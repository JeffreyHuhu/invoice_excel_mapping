# -*- coding: utf-8 -*-
"""
app.py
======
多供應商帳單自動化系統 (Streamlit 網頁)

流程：辨識哪間供應商帳單 → 選取對應的帳單辨識系統 → 擷取資料 → 跟正確答案
       Excel 逐欄比對算出正確率 (重複執行不同擷取設定，直到 100% 或試完為止)

用法：
    pip install streamlit pdfplumber openpyxl pandas
    streamlit run app.py

擴充：要支援第 21 家供應商，去 suppliers.py 加一組 detect_xxx()/parse_xxx()
並呼叫 register_supplier() 即可，這支 app.py 完全不用改。
"""

import datetime
import tempfile
import os

import streamlit as st

# ---------------------------------------------------------------------------
# 防禦性匯入：Streamlit Cloud 在「未捕捉的例外」發生時，會把真正的錯誤訊息
# 隱藏成一句通用的 "original error message is redacted"，讓人完全看不出
# 到底是哪個套件沒裝好。這裡改成自己 try/except 去匯入，一旦失敗就用
# st.error() 印出明確、不會被隱藏的訊息，並提示排解步驟。
# 最常見原因：requirements.txt 是後來才加上/修改的，但 Streamlit Cloud
# 沒有重新完整安裝套件 —— 到 App 右下角「Manage app」選單點
# 「Reboot app」(或「Clear cache」再重新部署) 就能強制重建環境解決。
# ---------------------------------------------------------------------------
try:
    import pandas as pd

    import suppliers  # noqa: F401  (只是為了觸發所有供應商的 register_supplier())
    from extract_utils import (
        FIELD_CODES,
        DISPLAY_HEADERS,
        NA,
        find_best_extraction,
        read_reference_excel,
        build_excel,
        normalize_value,
        detect_supplier,
        extract_text_from_pdf,
        supplier_label,
        list_registered_suppliers,
    )
except ImportError as e:
    st.set_page_config(page_title="多供應商帳單自動化系統 - 啟動失敗", layout="wide")
    st.error(
        f"❌ 系統啟動失敗，缺少必要的 Python 套件：`{e}`\n\n"
        "這通常代表 Streamlit Cloud 沒有正確安裝 `requirements.txt` 裡列出的套件"
        "(常發生在 requirements.txt 是後來才新增/修改，但環境沒有重新完整建置時)。"
        "請依序嘗試：\n\n"
        "1. 確認 repo 根目錄的 `requirements.txt` 存在，且內容包含 "
        "`streamlit`、`pdfplumber`、`openpyxl`、`pandas`\n"
        "2. 到這個 App 頁面右下角點 **Manage app** → 選單裡點 **Reboot app**，"
        "強制重新安裝所有套件\n"
        "3. 若還是失敗，同樣在 **Manage app** 裡可以看到完整安裝 log，"
        "確認實際是哪個套件安裝失敗"
    )
    st.stop()

st.set_page_config(page_title="多供應商帳單自動化系統", layout="wide")
st.title("📄➡️📊 多供應商帳單自動化系統")
st.caption(
    "流程：① 自動辨識帳單屬於哪家供應商 → ② 套用該供應商專屬的擷取規則 → "
    "③ 跟正確答案 Excel 逐欄比對算出正確率 (重複嘗試多組設定，直到 100% 或試完為止)。"
)

registered = list_registered_suppliers()
with st.expander(f"🏷️ 目前系統已支援 {len(registered)} 家供應商的辨識規則"):
    for s in registered:
        st.write(f"- **{s['label']}** (代碼: `{s['key']}`)")
    st.caption("要新增供應商，請在 suppliers.py 增加一組辨識/擷取規則並註冊，不用改這個網頁程式。")

col_pdf, col_ref = st.columns(2)
with col_pdf:
    uploaded_pdfs = st.file_uploader(
        "① 上傳 PDF 帳單 (可一次選取多個檔案、可混合不同供應商)",
        type=["pdf"], accept_multiple_files=True, key="pdf_upload",
    )
with col_ref:
    reference_excel = st.file_uploader(
        "② (選填) 上傳正確答案 Excel，用來自動比對並算出正確率",
        type=["xlsx"], accept_multiple_files=False, key="reference_upload",
    )
    st.caption("沒有上傳的話，會改用『資料完整度』(欄位是否成功抓到值，而非 N/A) 作為代理指標。")

if not uploaded_pdfs:
    st.info("請先上傳至少一份 PDF 帳單。")
    st.stop()

reference_df = None
if reference_excel is not None:
    try:
        reference_df = read_reference_excel(reference_excel.getvalue())
        if reference_df.empty:
            st.warning("⚠️ 正確答案 Excel 讀不到任何資料列，將改用資料完整度分數。")
            reference_df = None
    except Exception as e:  # noqa: BLE001
        st.warning(f"⚠️ 正確答案 Excel 讀取失敗 ({e})，將改用資料完整度分數。")

has_reference = reference_df is not None
score_label = "正確率" if has_reference else "資料完整度"

all_export_rows = []
score_rows = []
compare_rows_all = []
attempts_log = []
unknown_files = []

with st.spinner("解析中 (每份 PDF 先辨識供應商，再重複嘗試多組擷取設定)..."):
    for f in uploaded_pdfs:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(f.getvalue())
            tmp_path = tmp.name
        try:
            # 第①步：辨識供應商 (用預設參數快速判斷一次)
            quick_text = extract_text_from_pdf(tmp_path)
            supplier_key = detect_supplier(quick_text)

            if supplier_key is None:
                unknown_files.append(f.name)

            # 若有正確答案 Excel，先用預設設定拿 Invoice No，藉此在正確答案
            # 裡找出這份 PDF 對應的那一列 (可能有好幾家供應商的答案混在同一份Excel)
            reference_rows = None
            if has_reference and "invoice_no" in reference_df.columns:
                from extract_utils import parse_invoice_pdf
                _, quick_rows, _ = parse_invoice_pdf(tmp_path, supplier_key=supplier_key)
                inv_no = normalize_value(quick_rows[0].get("invoice_no"))
                mask = reference_df["invoice_no"].map(normalize_value) == inv_no
                candidate = reference_df[mask]
                if not candidate.empty:
                    reference_rows = candidate

            # 第②③步：套用對應供應商規則擷取，並重複嘗試設定直到最高正確率
            detected_key, rows, score, cfg, results, attempts = find_best_extraction(
                tmp_path, supplier_key=supplier_key, reference_rows=reference_rows
            )

            label = supplier_label(detected_key)
            for row in rows:
                row["__supplier_label"] = label
            all_export_rows.extend(rows)

            score_rows.append({
                "來源檔案": f.name,
                "辨識供應商": label,
                "Invoice No": rows[0].get("invoice_no", NA),
                score_label: f"{score * 100:.1f}%",
                "採用設定": cfg or "預設",
                "比對基準": "正確答案 Excel" if reference_rows is not None else "資料完整度自我檢查",
            })
            attempts_log.append({"檔案": f.name, "供應商": label, "嘗試紀錄": attempts})

            if results:
                for r in results:
                    compare_rows_all.append({"來源檔案": f.name, "供應商": label, **r})

        except Exception as e:  # noqa: BLE001
            st.error(f"❌ {f.name}：解析失敗 ({e})")
        finally:
            os.remove(tmp_path)

if unknown_files:
    st.warning(
        "⚠️ 以下檔案無法辨識供應商 (未命中任何已註冊的規則)，欄位皆以 N/A 填入。"
        "請到 suppliers.py 新增這家供應商的辨識/擷取規則：\n\n"
        + "\n".join(f"- {name}" for name in unknown_files)
    )

if not all_export_rows:
    st.stop()

# ---------------------------------------------------------------------------
# 轉檔結果預覽
# ---------------------------------------------------------------------------

preview_records = []
for row in all_export_rows:
    rec = {"供應商": row.get("__supplier_label", NA)}
    for c in FIELD_CODES:
        rec[DISPLAY_HEADERS[c].split("\n")[0]] = row.get(c, NA)
    preview_records.append(rec)
preview_df = pd.DataFrame(preview_records)

st.subheader("✏️ 轉檔結果預覽（可直接在表格中修正錯誤欄位；抓不到的欄位顯示 N/A）")
edited_df = st.data_editor(preview_df, num_rows="dynamic", use_container_width=True)

display_to_code = {DISPLAY_HEADERS[c].split("\n")[0]: c for c in FIELD_CODES}
export_rows = []
for _, r in edited_df.iterrows():
    row = {display_to_code[col]: r[col] for col in display_to_code}
    row["__supplier_label"] = r["供應商"]
    export_rows.append(row)

excel_bytes = build_excel(export_rows)
st.download_button(
    "⬇️ 下載 Excel (彙整所有供應商的轉檔結果)",
    data=excel_bytes,
    file_name=f"帳單轉檔結果_{datetime.date.today().isoformat()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# ---------------------------------------------------------------------------
# 自動顯示正確率 (%)
# ---------------------------------------------------------------------------

st.divider()
st.subheader(f"📊 轉檔{score_label}")

if not has_reference:
    st.info(
        "目前沒有上傳正確答案 Excel，以下是『資料完整度』：欄位有成功抓到值 "
        "(不是 N/A) 的比例，不是跟人工核對過的正確率；上傳正確答案 Excel "
        "後即可看到真正的逐欄比對正確率。"
    )

score_df = pd.DataFrame(score_rows)
st.dataframe(score_df, use_container_width=True)

avg_score = score_df[score_label].str.rstrip("%").astype(float).mean()
if has_reference and avg_score >= 100.0:
    st.success(f"🎉 整體{score_label}達到 100%！")
else:
    st.metric(f"整體{score_label} (所有 PDF 平均)", f"{avg_score:.1f}%")

with st.expander("🔍 查看每份 PDF 重複執行各組擷取設定時的分數"):
    for log in attempts_log:
        st.write(f"**{log['檔案']}** (供應商: {log['供應商']})")
        st.dataframe(pd.DataFrame(log["嘗試紀錄"]), use_container_width=True)

if has_reference and compare_rows_all:
    st.subheader("📋 逐欄比對明細（🟢相符 / 🔴不相符）")
    cmp_df = pd.DataFrame(compare_rows_all)

    def _highlight(row):
        color = "background-color:#C6EFCE" if row["結果"].startswith("✅") else "background-color:#FFC7CE"
        return [color] * len(row)

    st.dataframe(cmp_df.style.apply(_highlight, axis=1), use_container_width=True)
