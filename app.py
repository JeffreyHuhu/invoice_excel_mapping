# -*- coding: utf-8 -*-
"""
app.py
======
供應商帳單自動化系統 (全新重建版，鎖定 KUEHNE NAGEL 格式)

流程：
  1. 上傳 PDF 帳單 (Faktur Pajak + KN Sales Invoice 混合檔)
  2. (選填) 上傳正確答案 Excel (跟 KUEHNE_NAGEL_資料模板.xlsx 同樣欄位順序)
  3. 系統對每份 PDF 用多組擷取參數「重複執行」，每次都跟正確答案逐欄比對，
     一旦正確率達到 100% 就提早停止，否則回傳嘗試過最高的正確率
  4. 抓不到的欄位一律顯示 'N/A'
  5. 正確率(%) 直接顯示在網頁上，並用紅綠燈標示每個欄位是否相符
"""

import datetime
import tempfile
import os

import streamlit as st
import pandas as pd

from extract_utils import (
    FIELD_CODES,
    DISPLAY_HEADERS,
    NA,
    find_best_extraction,
    read_reference_excel,
    build_excel,
    normalize_value,
)

st.set_page_config(page_title="供應商帳單自動化系統 (KUEHNE NAGEL)", layout="wide")
st.title("📄➡️📊 供應商帳單自動化系統")
st.caption(
    "上傳 PDF 帳單後，系統會用多組擷取設定『重複執行』並自動跟正確答案 Excel 比對，"
    "挑出正確率最高的結果 (可達 100% 就提早停止)，抓不到的欄位一律顯示 N/A。"
)

col_pdf, col_ref = st.columns(2)
with col_pdf:
    uploaded_pdfs = st.file_uploader(
        "① 上傳 PDF 帳單 (可一次選取多個檔案)",
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

all_rows = []
score_rows = []
compare_rows = []
attempts_log = []

with st.spinner("解析中 (每份 PDF 會重複嘗試多組設定，直到 100% 正確或試完所有設定)..."):
    for f in uploaded_pdfs:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(f.getvalue())
            tmp_path = tmp.name
        try:
            reference = None
            if has_reference:
                # 先用預設設定快速解析拿 Invoice No，藉此在正確答案 Excel
                # 裡找出這份 PDF 對應的那一列
                quick_row = find_best_extraction(tmp_path, reference=None, configs=[{}])[0]
                inv_no = normalize_value(quick_row.get("invoice_no"))
                mask = reference_df["invoice_no"].map(normalize_value) == inv_no
                candidate = reference_df[mask]
                if not candidate.empty:
                    reference = candidate.iloc[0].to_dict()

            best_row, best_score, best_cfg, best_results, attempts = find_best_extraction(
                tmp_path, reference=reference
            )

            row_with_file = {"來源檔案": f.name, **{DISPLAY_HEADERS[c].split("\n")[0]: best_row[c] for c in FIELD_CODES}}
            all_rows.append((f.name, best_row))

            score_rows.append({
                "來源檔案": f.name,
                "Invoice No": best_row.get("invoice_no", NA),
                score_label: f"{best_score * 100:.1f}%",
                "採用設定": best_cfg or "預設",
                "比對基準": "正確答案 Excel" if reference is not None else "資料完整度自我檢查",
            })
            attempts_log.append({"檔案": f.name, "嘗試紀錄": attempts})

            if best_results is not None:
                for r in best_results:
                    compare_rows.append({"來源檔案": f.name, **r})

        except Exception as e:  # noqa: BLE001
            st.error(f"❌ {f.name}：解析失敗 ({e})")
        finally:
            os.remove(tmp_path)

if not all_rows:
    st.stop()

# ---------------------------------------------------------------------------
# 轉檔結果預覽
# ---------------------------------------------------------------------------

preview_records = []
for fname, row in all_rows:
    rec = {"來源檔案": fname}
    for c in FIELD_CODES:
        rec[DISPLAY_HEADERS[c].split("\n")[0]] = row.get(c, NA)
    preview_records.append(rec)
preview_df = pd.DataFrame(preview_records)

st.subheader("✏️ 轉檔結果預覽（可直接在表格中修正錯誤欄位；抓不到的欄位顯示 N/A）")
edited_df = st.data_editor(preview_df, num_rows="dynamic", use_container_width=True)

# 把編輯後的表格轉回內部欄位代碼，準備輸出 Excel
display_to_code = {DISPLAY_HEADERS[c].split("\n")[0]: c for c in FIELD_CODES}
export_rows = []
for _, r in edited_df.iterrows():
    export_rows.append({display_to_code[col]: r[col] for col in display_to_code})

excel_bytes = build_excel(export_rows)
st.download_button(
    "⬇️ 下載 Excel (符合 KUEHNE NAGEL 範本格式)",
    data=excel_bytes,
    file_name=f"KUEHNE_NAGEL_轉檔結果_{datetime.date.today().isoformat()}.xlsx",
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
        st.write(f"**{log['檔案']}**")
        st.dataframe(pd.DataFrame(log["嘗試紀錄"]), use_container_width=True)

if has_reference and compare_rows:
    st.subheader("📋 逐欄比對明細（🟢相符 / 🔴不相符）")
    cmp_df = pd.DataFrame(compare_rows)

    def _highlight(row):
        color = "background-color:#C6EFCE" if row["結果"].startswith("✅") else "background-color:#FFC7CE"
        return [color] * len(row)

    st.dataframe(cmp_df.style.apply(_highlight, axis=1), use_container_width=True)
