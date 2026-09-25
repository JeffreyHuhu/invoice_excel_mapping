# -*- coding: utf-8 -*-
"""
app.py
======
供應商 PDF 帳單自動化系統 (整合 Step 1 轉檔與 Step 2 查核比對)

功能：
  1. 頁籤 1：PDF 帳單轉 Excel 網頁小程式 (可互動編輯、下載 Excel)
  2. 頁籤 2：自動化查核比對小程式 (同時上傳 PDF 與填好的 Excel，計算正確率並下載紅綠燈查核報告)
"""

import io
import datetime
import tempfile
import os

import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from extract_utils import parse_invoice_pdf, invoice_to_rows, TEMPLATE_COLUMNS

# ---------------------------------------------------------------------------
# 樣式與對照設定
# ---------------------------------------------------------------------------

GREEN = PatternFill("solid", fgColor="C6EFCE")
RED = PatternFill("solid", fgColor="FFC7CE")

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

HEADER_FIELD_MAP = {
    "Invoice Date": "invoice_date",
    "Invoice No": "invoice_no",
    "Payment can be Transfer to": "supplier",
    "TO": "consignee",
    "Order No.": "order_no",
    "Arrive Date": "arrive_date",
    "on board date": "onboard_date",
    "B/L NO.": "bl_no",
    "MBL NO.": "mbl_no",
    "Port of Loading": "port_of_loading",
    "Port of discharge": "port_of_discharge",
    "Volume": "volume",
    "Vessel": "vessel",
    "Container no.": "container_no",
}


# ---------------------------------------------------------------------------
# 共用輔助函式
# ---------------------------------------------------------------------------

def normalize(value) -> str:
    """把值標準化成字串，方便比對"""
    if value is None:
        return ""
    if isinstance(value, (datetime.date, datetime.datetime)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).strip().upper()
    text = text.replace(",", "")
    text = text.rstrip(",")
    return text


def build_excel(df: pd.DataFrame) -> bytes:
    """把整理好的 DataFrame 輸出成符合範本格式的 Excel (bytes)。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "DWIHARTA"

    bold = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="DDEBF7")
    wrap_center = Alignment(wrap_text=True, vertical="center", horizontal="center")

    ws.cell(row=1, column=1, value="抓取欄位")
    ws.cell(row=2, column=1, value="顯示欄位")

    for idx, col_name in enumerate(TEMPLATE_COLUMNS, start=2):
        ws.cell(row=1, column=idx, value=col_name)
        cell2 = ws.cell(row=2, column=idx, value=DISPLAY_HEADERS.get(col_name, col_name))
        cell2.font = bold
        cell2.fill = header_fill
        cell2.alignment = wrap_center
        ws.column_dimensions[get_column_letter(idx)].width = 20

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


def load_excel_rows_from_bytes(excel_bytes: bytes) -> pd.DataFrame:
    """從上傳的 Excel bytes 讀取資料列"""
    wb = load_workbook(io.BytesIO(excel_bytes), data_only=True)
    ws = wb.active
    records = []
    for row in ws.iter_rows(min_row=3, values_only=False):
        values = {}
        empty = True
        for idx, col_name in enumerate(TEMPLATE_COLUMNS, start=2):
            cell = row[idx - 1]
            values[col_name] = cell.value
            if cell.value not in (None, ""):
                empty = False
        if not empty:
            records.append(values)
    return pd.DataFrame(records)


def compare_invoice(excel_group: pd.DataFrame, parsed_pdf) -> list:
    results = []
    first_row = excel_group.iloc[0]

    for col_name, attr_name in HEADER_FIELD_MAP.items():
        excel_val = first_row.get(col_name)
        pdf_val = getattr(parsed_pdf.header, attr_name)
        if attr_name in ("invoice_date", "arrive_date", "onboard_date") and pdf_val is None:
            pdf_val = getattr(parsed_pdf.header, attr_name + "_raw")

        match = normalize(excel_val) == normalize(pdf_val)
        results.append({
            "欄位": col_name,
            "Excel值": excel_val,
            "PDF重新解析值": pdf_val,
            "結果": "相符" if match else "不相符",
        })

    excel_items = set()
    for _, r in excel_group.iterrows():
        desc = normalize(r.get("Description\nVAT"))
        amt = normalize(r.get("Amount (IDR"))
        if desc or amt:
            excel_items.add((desc, amt))

    pdf_items = set(
        (normalize(li.description), normalize(li.amount_idr))
        for li in parsed_pdf.line_items
    )

    matched_items = excel_items & pdf_items
    only_in_excel = excel_items - pdf_items
    only_in_pdf = pdf_items - excel_items

    for desc, amt in sorted(matched_items):
        results.append({
            "欄位": "費用明細",
            "Excel值": f"{desc} / {amt}",
            "PDF重新解析值": f"{desc} / {amt}",
            "結果": "相符",
        })
    for desc, amt in sorted(only_in_excel):
        results.append({
            "欄位": "費用明細",
            "Excel值": f"{desc} / {amt}",
            "PDF重新解析值": "(PDF未找到對應項目)",
            "結果": "不相符",
        })
    for desc, amt in sorted(only_in_pdf):
        results.append({
            "欄位": "費用明細",
            "Excel值": "(Excel缺漏)",
            "PDF重新解析值": f"{desc} / {amt}",
            "結果": "不相符",
        })

    return results


def generate_report_bytes(result_df: pd.DataFrame, summary_df: pd.DataFrame) -> bytes:
    wb = Workbook()

    # Sheet1: 摘要
    ws1 = wb.active
    ws1.title = "正確率摘要"
    ws1.append(["欄位", "比對筆數", "相符筆數", "正確率"])
    for cell in ws1[1]:
        cell.font = Font(bold=True)
    for idx, row in summary_df.iterrows():
        ws1.append([idx, int(row["比對筆數"]), int(row["相符筆數"]), row["正確率"]])
    for col_letter, width in zip("ABCD", (22, 12, 12, 12)):
        ws1.column_dimensions[col_letter].width = width

    # Sheet2: 明細
    ws2 = wb.create_sheet("逐筆比對明細")
    cols = ["Invoice No", "欄位", "Excel值", "PDF重新解析值", "結果"]
    ws2.append(cols)
    for cell in ws2[1]:
        cell.font = Font(bold=True)
    for _, r in result_df.iterrows():
        ws2.append([r[c] for c in cols])
        fill = GREEN if r["結果"] == "相符" else RED
        ws2.cell(row=ws2.max_row, column=5).fill = fill
    widths = (18, 16, 30, 30, 10)
    for col_letter, width in zip("ABCDE", widths):
        ws2.column_dimensions[col_letter].width = width

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Streamlit 網頁主介面
# ---------------------------------------------------------------------------

st.set_page_config(page_title="供應商帳單自動化與查核系統", layout="wide")
st.title("📄➡️📊 供應商帳單自動化系統 (轉檔與查核)")

tab1, tab2 = st.tabs(["📄 Step 1：PDF 轉 Excel 轉檔工具", "🔍 Step 2：帳單查核比對工具"])

# --- 頁籤 1：PDF 轉 Excel ---
with tab1:
    st.header("Step 1：PDF 帳單轉 Excel 網頁小程式")
    st.caption("上傳可編輯的 PDF 帳單，自動抽取欄位並整理成 Excel，取代人工逐筆輸入。")

    uploaded_files = st.file_uploader(
        "上傳 PDF 帳單 (可一次選取多個檔案)",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_upload_tab1"
    )

    if uploaded_files:
        all_rows = []
        parse_errors = []

        with st.spinner("解析中..."):
            for f in uploaded_files:
                try:
                    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                        tmp.write(f.getvalue())
                        tmp_path = tmp.name
                    try:
                        parsed = parse_invoice_pdf(tmp_path)
                        rows = invoice_to_rows(parsed)
                        for row in rows:
                            row["__來源檔案"] = f.name
                        if not parsed.line_items:
                            parse_errors.append(f"⚠️ {f.name}：找不到任何費用明細，請確認 PDF 是否為可編輯文字檔。")
                        all_rows.extend(rows)
                    finally:
                        os.remove(tmp_path)
                except Exception as e:
                    parse_errors.append(f"❌ {f.name}：解析失敗 ({e})")

        if parse_errors:
            for msg in parse_errors:
                st.warning(msg)

        if all_rows:
            df = pd.DataFrame(all_rows)
            display_cols = ["__來源檔案"] + TEMPLATE_COLUMNS
            df = df[display_cols].rename(columns={"__來源檔案": "來源檔案"})

            st.subheader("✏️ 轉檔結果預覽（可直接在表格中修正錯誤欄位）")
            edited_df = st.data_editor(
                df,
                num_rows="dynamic",
                use_container_width=True,
            )

            export_df = edited_df.rename(columns={"來源檔案": "__來源檔案"})
            export_df = export_df[TEMPLATE_COLUMNS]

            excel_bytes = build_excel(export_df)

            st.download_button(
                label="⬇️ 下載 Excel (符合 DWIHARTA 範本格式)",
                data=excel_bytes,
                file_name=f"DWIHARTA_轉檔結果_{datetime.date.today().isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.info("請先上傳至少一份 PDF 帳單。")

# --- 頁籤 2：查核比對工具 ---
with tab2:
    st.header("Step 2：帳單查核比對工具 (計算正確率)")
    st.caption("同時上傳原始 PDF 帳單與填好的 Excel 檔案，系統將自動比對並產出紅綠燈查核報告。")

    col_a, col_b = st.columns(2)
    with col_a:
        verify_pdfs = st.file_uploader(
            "1. 上傳原始 PDF 帳單 (可多選)",
            type=["pdf"],
            accept_multiple_files=True,
            key="pdf_upload_tab2"
        )
    with col_b:
        verify_excel = st.file_uploader(
            "2. 上傳填好的 Excel 檔案 (單選)",
            type=["xlsx"],
            accept_multiple_files=False,
            key="excel_upload_tab2"
        )

    if verify_pdfs and verify_excel:
        if st.button("🚀 開始執行查核比對", type="primary"):
            with st.spinner("比對中，請稍候..."):
                try:
                    # 讀取並重新解析所有上傳的 PDF
                    ground_truth = {}
                    for f in verify_pdfs:
                        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                            tmp.write(f.getvalue())
                            tmp_path = tmp.name
                        try:
                            parsed = parse_invoice_pdf(tmp_path)
                            invoice_no = normalize(parsed.header.invoice_no)
                            if invoice_no:
                                ground_truth[invoice_no] = parsed
                        finally:
                            os.remove(tmp_path)

                    # 讀取上傳的 Excel
                    excel_df = load_excel_rows_from_bytes(verify_excel.getvalue())

                    if excel_df.empty:
                        st.error("❌ 上傳的 Excel 中沒有讀到任何資料列，請確認格式是否正確。")
                    else:
                        all_results = []
                        for invoice_no, group in excel_df.groupby(excel_df["Invoice No"].map(normalize)):
                            parsed_pdf = ground_truth.get(invoice_no)
                            if parsed_pdf is None:
                                all_results.append({
                                    "Invoice No": invoice_no,
                                    "欄位": "(整張帳單)",
                                    "Excel值": invoice_no,
                                    "PDF重新解析值": "(找不到對應的PDF檔案)",
                                    "結果": "不相符",
                                })
                                continue
                            for rec in compare_invoice(group, parsed_pdf):
                                all_results.append({"Invoice No": invoice_no, **rec})

                        result_df = pd.DataFrame(all_results)

                        # 計算摘要
                        summary = (
                            result_df.groupby("欄位")["結果"]
                            .apply(lambda s: pd.Series({
                                "比對筆數": len(s),
                                "相符筆數": (s == "相符").sum(),
                                "正確率": f"{(s == '相符').mean() * 100:.1f}%",
                            }))
                            .unstack()
                        )
                        summary = summary[["比對筆數", "相符筆數", "正確率"]]
                        total = len(result_df)
                        matched = (result_df["結果"] == "相符").sum()
                        overall = pd.DataFrame(
                            [{"比對筆數": total, "相符筆數": matched,
                              "正確率": f"{matched / total * 100:.1f}%" if total else "N/A"}],
                            index=["整體 (Overall)"],
                        )
                        summary_df = pd.concat([summary, overall])

                        st.success("✅ 查核比對完成！")
                        
                        st.subheader("📊 正確率摘要統計")
                        st.dataframe(summary_df, use_container_width=True)

                        report_bytes = generate_report_bytes(result_df, summary_df)
                        st.download_button(
                            label="⬇️ 下載詳細紅綠燈查核報告 (Excel)",
                            data=report_bytes,
                            file_name=f"查核報告_{datetime.date.today().isoformat()}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                except Exception as e:
                    st.error(f"❌ 執行比對時發生錯誤：{e}")
    else:
        st.info("請同時上傳「原始 PDF 帳單」以及「要檢查的 Excel 檔案」以開始比對。")
