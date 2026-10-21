# -*- coding: utf-8 -*-
"""
data_admin.py — 資料管理頁 (由 app.py 載入)

讓不會用指令列的人也能維護資料：
  1. 資料庫位置：指定要讀哪一份 expenses.duckdb (例如公司共用資料夾)，存進
     dashboard_config.json，下次開啟沿用。
  2. 更新資料：輸入轉檔結果所在的資料夾，按按鈕就執行 process_data.run()
     (增量更新，只處理新增/修改的 Excel)。
  3. 匯率：一鍵抓 ECB 月平均匯率，或直接在表格裡修改/新增 (改過的列標成
     「手動」，之後自動更新不會覆蓋)。
  4. 讀取失敗的檔案清單。
"""

from pathlib import Path

import duckdb
import pandas as pd
import streamlit as st

import fx
import process_data
from settings import CONFIG_PATH, get_db_path, set_db_path

st.title("資料管理")
db_path = get_db_path()


def connect_rw() -> duckdb.DuckDBPyConnection:
    """寫入用連線；資料庫被別人鎖住時顯示提示並停止。"""
    try:
        db_path.parent.mkdir(parents=True, exist_ok=True)
        return duckdb.connect(str(db_path))
    except duckdb.Error as e:
        st.error(f"資料庫目前無法寫入 (可能有其他人正在更新)：{e}")
        st.stop()


# ============================================================
# 1. 資料庫位置
# ============================================================

st.subheader("1. 資料庫位置")
if db_path.is_file():
    con = duckdb.connect(str(db_path), read_only=True)
    try:
        rows, files, last = con.execute(
            "SELECT (SELECT count(*) FROM expense_items), (SELECT count(*) FROM etl_files), "
            "(SELECT max(processed_at) FROM etl_files)").fetchone()
    finally:
        con.close()
    st.success(f"目前使用：`{db_path}`　｜　{db_path.stat().st_size / 1e6:,.1f} MB、"
               f"明細 {rows:,} 筆、來源 Excel {files:,} 份、最後匯入 {last:%Y-%m-%d %H:%M}"
               if last else f"目前使用：`{db_path}` (還沒有資料)")
else:
    st.warning(f"目前指定的資料庫不存在：`{db_path}`。可以改指向既有的資料庫，"
               "或直接在下方「更新資料」建立一份新的。")

with st.form("db_path_form"):
    new_path = st.text_input("資料庫檔案路徑", str(db_path),
                             help="例如公司共用資料夾裡的 \\\\server\\share\\費用分析\\expenses.duckdb")
    if st.form_submit_button("儲存位置"):
        p = Path(new_path.strip().strip('"'))
        if p.suffix.lower() != ".duckdb":
            st.error("請輸入 .duckdb 檔案的完整路徑")
        else:
            set_db_path(p)
            st.cache_data.clear()
            st.rerun()
st.caption(f"設定存在 `{CONFIG_PATH}`。")

# ============================================================
# 2. 更新資料
# ============================================================

st.subheader("2. 從轉檔結果更新資料")
st.caption("掃描資料夾 (含所有子資料夾) 裡「XXX-辨識等級(高)」的轉檔結果 Excel。"
           "只會處理新增或修改過的檔案；更新期間其他人的儀表板可能暫時讀不到資料。")
with st.form("etl_form"):
    roots_text = st.text_area("轉檔結果所在資料夾 (一行一個)",
                              placeholder=r"D:\發票\印尼MAERSK 2025-2026")
    c1, c2, c3 = st.columns(3)
    level = c1.selectbox("辨識等級", ["高", "中", "低"], index=0)
    supplier_folder = c2.text_input("只讀這個供應商的資料夾 (選填)", placeholder="例如 MAERSK")
    full = c3.checkbox("完整重建 (清空後重新匯入)", value=False,
                       help="修改費用名稱對照規則後需要完整重建")
    submitted = st.form_submit_button("開始更新", type="primary")

if submitted:
    roots = [Path(line.strip().strip('"')) for line in roots_text.splitlines() if line.strip()]
    missing = [str(r) for r in roots if not r.is_dir()]
    if not roots:
        st.error("請輸入至少一個資料夾")
    elif missing:
        st.error("找不到資料夾：" + "、".join(missing))
    else:
        bar = st.progress(0.0, text="掃描資料夾中...")
        log_lines = []
        log_box = st.empty()

        def echo(msg: str) -> None:
            log_lines.append(msg)
            log_box.code("\n".join(log_lines[-15:]), language=None)

        def progress(done: int, total: int) -> None:
            if done % 50 == 0 or done == total:
                bar.progress(done / total, text=f"讀取 Excel {done:,}/{total:,}")

        try:
            result = process_data.run(
                [r.resolve() for r in roots], db_path, level, supplier_folder.strip() or None,
                full=full, use_threads=True, echo=echo, progress=progress)
        except duckdb.Error as e:
            st.error(f"資料庫目前無法寫入 (可能有其他人正在使用或更新)：{e}")
        else:
            bar.progress(1.0, text="完成")
            st.cache_data.clear()
            st.success(f"完成：新增/修改 {result['processed']:,} 份、寫入 {result['rows']:,} 筆明細、"
                       f"讀取失敗 {result['failed']:,} 份，耗時 {result['seconds']:,.0f} 秒。")

# ============================================================
# 3. 匯率
# ============================================================

st.subheader("3. 匯率 (1 USD = ? 原幣)")
if not db_path.is_file():
    st.info("建立資料庫後才能維護匯率。")
    st.stop()

col_a, col_b = st.columns([1, 3])
if col_a.button("🌐 從 ECB 更新匯率", help="向歐洲央行 (Frankfurter API) 抓每日匯率算月平均；只送出幣別與日期"):
    con = connect_rw()
    try:
        process_data.init_db(con)
        msg = fx.update_fx_rates(con)
        st.cache_data.clear()
        col_b.success(msg)
    except Exception as e:  # noqa: BLE001  網路失敗等
        col_b.error(f"更新失敗：{e}")
    finally:
        con.close()

con = duckdb.connect(str(db_path), read_only=True)
try:
    has_fx = con.execute("SELECT count(*) FROM information_schema.tables "
                         "WHERE table_name = 'fx_rates'").fetchone()[0]
    rates = con.execute("SELECT currency, year, month, per_usd, source FROM fx_rates "
                        "ORDER BY currency, year, month").df() if has_fx else \
        pd.DataFrame(columns=["currency", "year", "month", "per_usd", "source"])
finally:
    con.close()

edited = st.data_editor(
    rates, num_rows="dynamic", hide_index=True, width="stretch", key="fx_editor",
    column_config={
        "currency": st.column_config.TextColumn("幣別", required=True, max_chars=3),
        "year": st.column_config.NumberColumn("年", required=True, min_value=2000, max_value=2100, format="%d"),
        "month": st.column_config.NumberColumn("月", required=True, min_value=1, max_value=12, format="%d"),
        "per_usd": st.column_config.NumberColumn("1 USD = ?", required=True, min_value=0.0001, format="%,.4f"),
        "source": st.column_config.TextColumn("來源", disabled=True),
    })
if st.button("💾 儲存匯率修改"):
    key = ["currency", "year", "month"]
    new = edited.dropna(subset=key + ["per_usd"]).copy()
    new["currency"] = new["currency"].str.upper().str.strip()
    merged = new.merge(rates, on=key, how="left", suffixes=("", "_old"))
    # 新增或匯率有改的列標成「手動」，自動更新不會覆蓋
    changed = merged["per_usd_old"].isna() | (merged["per_usd"] != merged["per_usd_old"])
    new.loc[changed.values, "source"] = fx.SOURCE_MANUAL
    con = connect_rw()
    try:
        fx.ensure_table(con)
        con.register("edited_rates", new[key + ["per_usd", "source"]])
        con.execute("BEGIN")
        con.execute("DELETE FROM fx_rates")
        con.execute("INSERT INTO fx_rates (currency, year, month, per_usd, source) "
                    "SELECT currency, year, month, per_usd, source FROM edited_rates")
        con.execute("COMMIT")
    finally:
        con.close()
    st.cache_data.clear()
    st.success(f"已儲存 {len(new):,} 筆匯率 (其中 {int(changed.sum()):,} 筆標記為手動)")
    st.rerun()

# ============================================================
# 4. 讀取失敗的檔案
# ============================================================

st.subheader("4. 讀取失敗的檔案")
con = duckdb.connect(str(db_path), read_only=True)
try:
    errors = con.execute("SELECT source_file AS 檔案, error AS 原因, processed_at AS 時間 "
                         "FROM etl_files WHERE error IS NOT NULL ORDER BY 3 DESC").df()
finally:
    con.close()
if errors.empty:
    st.caption("沒有讀取失敗的檔案。")
else:
    st.dataframe(errors, hide_index=True, width="stretch")
