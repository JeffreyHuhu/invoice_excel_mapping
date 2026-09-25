# -*- coding: utf-8 -*-
"""
app.py
======
Step 1：供應商 PDF 帳單 -> Excel 轉檔 網頁小程式 (Streamlit)

執行方式：
    pip install streamlit pdfplumber openpyxl pandas
    streamlit run app.py

功能：
  1. 使用者可一次上傳多張 PDF 帳單 (可編輯 PDF，非掃描檔)
  2. 程式自動抽取抬頭欄位 + 費用明細，展開成表格預覽
  3. 使用者可以在網頁上直接「手動微調」抓錯的欄位 (data_editor)
  4. 一鍵下載成符合「資料模板-DWIHARTA.xlsx」格式的 Excel
     (保留範本原有的兩列表頭：抓取欄位 / 顯示欄位)
"""

import io
import datetime
import tempfile
import os

import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from extract_utils import parse_invoice_pdf, invoice_to_rows, TEMPLATE_COLUMNS

# 範本第一列(內部用) / 第二列(顯示用) 的中英對照文字
DISPLAY_HEADERS = {
    "Invoice Date": "INVOICE DATE\n發票日期",
    "Invoice No": "INVOICE NO\n發票號碼",
    "Payment can be Transfer to": "SUPPLIER\n供應商名稱",
    "TO": "CONSIGNEE\n集團公司名稱",
    "Description\nVAT": "Description\n費用名稱",
    "Amount (IDR": "AMOUNT\n金額",
    "Order No.": "Order No.",
    "Arrive Date": "Arrive Date\n到達日",
    "on board date": "on board date\n開船日",
    "B/L NO.": "B/L NO.\n提單號碼",
    "MBL NO.": "MBL NO.\n主提單號碼",
    "Port of Loading": "Port of Loading\n啟運港",
    "Port of discharge": "Port of discharge\n目的港",
    "Volume": "Volume\n材積",
    "Vessel": "Vessel\n船名",
    "Container no.": "Container no.\n貨櫃號碼",
}


# ---------------------------------------------------------------------------
# 核心函式
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def _parse_uploaded_pdf(file_bytes: bytes, file_name: str):
    """把上傳的 PDF bytes 存成暫存檔後丟給 extract_utils 解析。
    用 st.cache_data 快取，避免同一份檔案重複解析。
    """
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        parsed = parse_invoice_pdf(tmp_path)
        rows = invoice_to_rows(parsed)
        for row in rows:
            row["__來源檔案"] = file_name
        return rows, len(parsed.line_items)
    finally:
        os.remove(tmp_path)


def build_excel(df: pd.DataFrame) -> bytes:
    """把整理好的 DataFrame 輸出成符合範本格式的 Excel (bytes)。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "DWIHARTA"

    bold = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="DDEBF7")
    wrap_center = Alignment(wrap_text=True, vertical="center", horizontal="center")

    # 第一列：內部欄位代碼 (對應原範本 A1='抓取欄位')
    ws.cell(row=1, column=1, value="抓取欄位")
    # 第二列：顯示給人看的中英文欄名 (對應原範本 A2='顯示欄位')
    ws.cell(row=2, column=1, value="顯示欄位")

    for idx, col_name in enumerate(TEMPLATE_COLUMNS, start=2):  # B欄開始
        ws.cell(row=1, column=idx, value=col_name)
        cell2 = ws.cell(row=2, column=idx, value=DISPLAY_HEADERS.get(col_name, col_name))
        cell2.font = bold
        cell2.fill = header_fill
        cell2.alignment = wrap_center
        ws.column_dimensions[get_column_letter(idx)].width = 20

    # 資料從第3列開始
    for r, (_, row) in enumerate(df.iterrows(), start=3):
        for idx, col_name in enumerate(TEMPLATE_COLUMNS, start=2):
            value = row.get(col_name, "")
            if isinstance(value, (datetime.date, datetime.datetime)):
                cell = ws.cell(row=r, column=idx, value=value)
                cell.number_format = "yyyy-mm-dd"
            else:
                ws.cell(row=r, column=idx, value=value)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Streamlit 介面
# ---------------------------------------------------------------------------

st.set_page_config(page_title="供應商帳單 PDF → Excel 轉檔工具", layout="wide")
st.title("📄➡️📊 供應商帳單 PDF → Excel 自動轉檔工具")
st.caption("上傳可編輯的 PDF 帳單，自動抽取欄位並整理成 Excel，取代人工逐筆輸入。")

uploaded_files = st.file_uploader(
    "上傳 PDF 帳單 (可一次選取多個檔案)",
    type=["pdf"],
    accept_multiple_files=True,
)

if uploaded_files:
    all_rows = []
    parse_errors = []

    with st.spinner("解析中..."):
        for f in uploaded_files:
            try:
                rows, n_items = _parse_uploaded_pdf(f.getvalue(), f.name)
                if n_items == 0:
                    parse_errors.append(f"⚠️ {f.name}：找不到任何費用明細，請確認 PDF 是否為可編輯文字檔。")
                all_rows.extend(rows)
            except Exception as e:  # noqa: BLE001
                parse_errors.append(f"❌ {f.name}：解析失敗 ({e})")

    if parse_errors:
        for msg in parse_errors:
            st.warning(msg)

    if all_rows:
        df = pd.DataFrame(all_rows)
        # 欄位排序：來源檔案放最前面方便核對，其餘照範本順序
        display_cols = ["__來源檔案"] + TEMPLATE_COLUMNS
        df = df[display_cols].rename(columns={"__來源檔案": "來源檔案"})

        st.subheader("✏️ 轉檔結果預覽（可直接在表格中修正錯誤欄位）")
        edited_df = st.data_editor(
            df,
            num_rows="dynamic",
            use_container_width=True,
        )

        # 還原欄位名稱以便輸出
        export_df = edited_df.rename(columns={"來源檔案": "__來源檔案"})
        export_df = export_df[TEMPLATE_COLUMNS]

        excel_bytes = build_excel(export_df)

        st.download_button(
            label="⬇️ 下載 Excel (符合 DWIHARTA 範本格式)",
            data=excel_bytes,
            file_name=f"DWIHARTA_轉檔結果_{datetime.date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.info(f"共處理 {len(uploaded_files)} 份 PDF，展開為 {len(export_df)} 筆費用明細列。"
                "下載後的 Excel 可直接交給 Step2 查核比對程式 (verify.py) 進行正確率檢查。")
else:
    st.info("請先上傳至少一份 PDF 帳單。")
