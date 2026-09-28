# -*- coding: utf-8 -*-
"""
verify_app.py
=============
供應商帳單人工檢核系統 (人工查核 + 回饋收集工具)

這是跟「供應商帳單辨識對比系統」(app.py) 分開的獨立 Streamlit App，
專門給人工「逐欄」核對系統擷取出來的資料是否正確，用於：

  1. 上傳一份供應商帳單 PDF (不需要正確答案 Excel)。
  2. 系統自動辨識供應商、擷取欄位 (跟 app.py 用同一套 suppliers.py /
     extract_utils.py 核心引擎，共用同一套辨識/擷取規則，兩邊不會有
     擷取邏輯不一致的問題)。
  3. 人工逐欄核對：同一張發票的抬頭欄位 (發票日期、發票號碼、供應商名稱...
     等，整張發票只會有一個值) 只顯示一次；費用名稱/金額這種同一張發票
     可能有多筆、逐筆不同的欄位，才依品項逐筆列出。每個欄位預設
     「沒有動作 = 視為忽略 (不計入回饋)」，只有按下「❌ 標記為錯誤」才會
     展開輸入框，讓使用者填入正確答案。
  4. 核對完成後，把「所有被標記為錯誤的欄位 + 使用者填的正確答案」匯出
     成一份 Excel (回饋記錄)，可以把這份 Excel 拿給開發端 (或直接請
     Claude 執行 add-invoice-supplier 這套流程) 依照回饋記錄修正
     suppliers.py 裡對應供應商的 parse() 規則，改完之後回到這個網頁、
     重新上傳同一份 PDF 再核對一次，確認錯誤欄位已經修正。

跟 app.py 的差異：
  - app.py 是「有正確答案 Excel 時的自動化比對」，回答「正確率幾%」。
  - verify_app.py 是「沒有正確答案 Excel 時的人工逐欄查核」，回答「這一
    欄到底對不對、不對的話正確答案是什麼」，產出的 Excel 就是下一次可以
    拿來當作「正確答案」或「除錯依據」的原始素材。
  - 兩支 App 各自獨立部署 (Streamlit Cloud 各自建一個 App，Main file
    path 分別指向 app.py / verify_app.py)，但共用同一個 GitHub repo、
    同一套 suppliers.py / extract_utils.py，新增或修正供應商規則時，
    兩邊會同時生效，不用分別維護兩份擷取邏輯。
"""

import hashlib
import io
import tempfile

import streamlit as st

# ---------------------------------------------------------------------------
# 匯入失敗時，印出清楚的錯誤訊息 (原因/排解方式跟 app.py 完全一樣：通常是
# Streamlit Cloud 沒有重新安裝 requirements.txt，到 Manage app 點 Reboot
# app 即可)，不要讓 Streamlit 預設的模糊訊息蓋掉真正的錯誤內容。
# ---------------------------------------------------------------------------
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    import suppliers  # noqa: F401 (觸發所有供應商的 register_supplier())
    from extract_utils import (
        NA,
        FIELD_CODES,
        DISPLAY_HEADERS,
        find_best_extraction,
        detect_supplier,
        extract_text_from_pdf,
        supplier_label,
        list_registered_suppliers,
    )
except ImportError as e:
    st.set_page_config(page_title="供應商帳單人工檢核系統 - 啟動失敗", layout="wide")
    st.error(
        f"❌ 系統啟動失敗，缺少必要的 Python 套件：`{e}`\n\n"
        "這通常代表 Streamlit Cloud 沒有正確安裝 `requirements.txt` 裡列出的套件"
        "(常發生在 requirements.txt 是後來才新增/修改，但環境沒有重新完整建置時)。"
        "請到這個 App 頁面右下角點 **Manage app** → 選單裡點 **Reboot app**，"
        "強制重新安裝所有套件。"
    )
    st.stop()

st.set_page_config(page_title="供應商帳單人工檢核系統", layout="wide")

# ---------------------------------------------------------------------------
# 顯示欄位設定：只顯示使用者指定的欄位與順序 (不顯示 Order No.)。
# 抬頭欄位 (整張發票只會有一個值，例如發票日期、供應商名稱) 跟品項欄位
# (費用名稱／金額，同一張發票可能有多筆費用、逐筆不同) 分開處理：抬頭
# 欄位同一張發票只列一次，品項欄位才依每一筆費用逐筆列出。
# ---------------------------------------------------------------------------
HEADER_DISPLAY_FIELDS = [
    ("invoice_date", "INVOICE DATE\n發票日期"),
    ("invoice_no", "發票號碼"),
    ("supplier", "供應商名稱"),
    ("consignee", "集團公司名稱"),
    ("order_no", "Order No."),
    ("arrive_date", "到達日"),
    ("onboard_date", "開船日"),
    ("bl_no", "提單號碼"),
    ("mbl_no", "主提單號碼"),
    ("port_of_loading", "啟運港"),
    ("port_of_discharge", "目的港"),
    ("volume", "材積"),
    ("vessel", "船名"),
    ("container_no", "貨櫃號碼"),
]
ITEM_DISPLAY_FIELDS = [
    ("description", "費用名稱"),
    ("amount", "金額"),
]

# ---------------------------------------------------------------------------
# 全站樣式：字體/按鈕顏色沿用「供應商帳單辨識對比系統」(app.py) 同一套
# 視覺規範 (淺藍色按鈕、16px 字體)，讓兩支 App 看起來是同一套系統的兩個
# 頁面，不會有風格不一致的違和感；標題列的 flex 排版手法也沿用 app.py，
# 把「重新查核」按鈕推到標題右側 (右上方)。
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    div.stButton > button,
    div.stDownloadButton > button,
    div.stFormSubmitButton > button,
    div[data-testid="stButton"] > button,
    div[data-testid="stDownloadButton"] > button,
    div[data-testid="stFormSubmitButton"] > button {
        font-size: 16px !important;
        font-weight: 800 !important;
        padding: 1.4em 2em !important;
        height: auto !important;
        border-radius: 14px !important;
        background-color: #90CAF9 !important;
        border-color: #90CAF9 !important;
        color: #0D3B66 !important;
    }
    div.stButton > button:hover,
    div.stDownloadButton > button:hover,
    div.stFormSubmitButton > button:hover,
    div[data-testid="stButton"] > button:hover,
    div[data-testid="stDownloadButton"] > button:hover,
    div[data-testid="stFormSubmitButton"] > button:hover {
        background-color: #64B5F6 !important;
        border-color: #64B5F6 !important;
        color: #0D3B66 !important;
    }
    div[data-testid="stHeading"] h3,
    div[data-testid="stMarkdownContainer"] h3,
    section.main h3 {
        font-size: 24px !important;
    }
    /* 標題列改成「依內容自動縮寬」的 flex 排版，並把「重新查核」按鈕推到
       最右邊 (space-between 撐開)，讓按鈕落在畫面的右上角，做法跟
       app.py 的標題列一致。 */
    div:has(> #title-row-marker) + div[data-testid="stHorizontalBlock"] {
        display: flex !important;
        flex-wrap: nowrap !important;
        align-items: center !important;
        justify-content: space-between !important;
    }
    div:has(> #title-row-marker) + div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
        flex: 0 0 auto !important;
        width: auto !important;
        min-width: 0 !important;
    }
    /* 整體版面再簡潔一點：縮小 Streamlit 預設區塊之間的垂直間距、頁面
       上下留白、分隔線的留白，畫面才不會顯得鬆散。 */
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 2rem !important;
    }
    div[data-testid="stVerticalBlock"] {
        gap: 0.4rem !important;
    }
    hr {
        margin: 0.6rem 0 !important;
    }
    /* 上傳 PDF 的拖曳框放大兩倍：Streamlit 的 file_uploader 沒有內建大小
       參數，直接放大它的內部元素 (虛線框內距/最小高度、說明文字、圖示、
       Browse files 按鈕、已上傳檔案清單)，做法跟供應商帳單自動化辨識
       系統 (app.py) 當初「放大兩倍」的版本一致。 */
    [data-testid="stFileUploaderDropzone"] {
        padding: 2.4em 2em !important;
        min-height: 220px !important;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] span,
    [data-testid="stFileUploaderDropzoneInstructions"] small {
        font-size: 2em !important;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] small {
        font-size: 1.5em !important;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] svg {
        width: 2.4em !important;
        height: 2.4em !important;
    }
    [data-testid="stFileUploaderDropzone"] button {
        font-size: 1.4em !important;
        padding: 0.8em 1.4em !important;
    }
    [data-testid="stFileUploaderFile"] {
        font-size: 1.6em !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

if "uploader_version" not in st.session_state:
    st.session_state["uploader_version"] = 0

st.markdown('<span id="title-row-marker"></span>', unsafe_allow_html=True)
title_col, reset_col = st.columns([3, 1])
with title_col:
    st.markdown(
        '<h1 style="font-size:40px; margin:0;">📝 供應商帳單人工檢核系統</h1>',
        unsafe_allow_html=True,
    )
with reset_col:
    st.write("")  # 讓按鈕跟標題文字的垂直位置對齊
    reset_clicked = st.button(
        "🔄 重新查核（換一份帳單）",
        help="清空目前的擷取結果、標記狀態，並清除已上傳的 PDF，方便重新上傳下一份帳單。",
    )
if reset_clicked:
    st.session_state.pop("verify_result", None)
    st.session_state.pop("verify_signature", None)
    for k in list(st.session_state.keys()):
        if k.startswith("err_") or k.startswith("fix_") or k.startswith("confirmed_"):
            del st.session_state[k]
    st.session_state["uploader_version"] += 1
    st.rerun()

st.markdown(
    """
    <div style="font-size:16px; color:#000000; font-weight:500; line-height:1.5; margin:8px 0 12px;">
        <div>① 上傳一份供應商帳單 PDF，系統自動辨識供應商並擷取欄位</div>
        <div>② 逐欄核對擷取結果：沒問題不用動作，有錯誤才按「❌ 標記為錯誤」並填入正確答案</div>
        <div>③ 核對完成後下載「回饋記錄 Excel」，可據此修正供應商帳單辨識對比系統的擷取規則</div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.expander("ℹ️ 目前已支援的供應商清單"):
    for s in list_registered_suppliers():
        st.write(f"- **{s['label']}** (代碼: `{s['key']}`)")
    st.caption(
        "沒有列在上面的供應商，系統會顯示「未知供應商 (無法辨識)」，"
        "所有欄位都會是 N/A，這種情況也可以照樣核對、標記正確答案，"
        "回饋記錄可以做為之後新增這家供應商規則的素材。"
    )

# ---------------------------------------------------------------------------
# ① 上傳 PDF
# ---------------------------------------------------------------------------
st.subheader("① 上傳供應商帳單 PDF")
uploaded_pdf = st.file_uploader(
    "選擇一份 PDF 帳單",
    type=["pdf"],
    accept_multiple_files=False,
    key=f"pdf_upload_{st.session_state['uploader_version']}",
)

if not uploaded_pdf:
    st.info("請先上傳一份 PDF 帳單。")
    st.stop()

file_signature = hashlib.md5(uploaded_pdf.getvalue()).hexdigest()

run_clicked = st.button("▶️ 開始擷取", type="primary", use_container_width=True)

# 一定要按下「▶️ 開始擷取」才會真的執行擷取；換了新檔案不會自動觸發，
# 避免使用者還沒按按鈕、畫面就自己跑起來，甚至誤用上一份檔案的殘留結果。
if run_clicked:
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(uploaded_pdf.getvalue())
        tmp_path = tmp.name

    with st.spinner("解析中 (自動辨識供應商、嘗試不同擷取參數挑出資料最完整的一次)..."):
        quick_text = extract_text_from_pdf(tmp_path)
        supplier_key = detect_supplier(quick_text)
        result = find_best_extraction(tmp_path, supplier_key=supplier_key, reference_rows=None)

    st.session_state["verify_result"] = result
    st.session_state["verify_signature"] = file_signature
    # 換了新檔案重新擷取，把之前殘留的標記狀態清掉，避免誤把上一份帳單
    # 的錯誤標記/確認狀態誤植到這一份帳單上。
    for k in list(st.session_state.keys()):
        if k.startswith("err_") or k.startswith("fix_") or k.startswith("confirmed_"):
            del st.session_state[k]

# 目前上傳的檔案還沒按過「開始擷取」(或換了檔案但沿用的是上一份的結果)，
# 就不要顯示任何擷取結果，避免看到不是這份檔案的資料。
if st.session_state.get("verify_signature") != file_signature:
    st.info("請按上方「▶️ 開始擷取」，開始擷取這份帳單的資料。")
    st.stop()

result = st.session_state.get("verify_result")
if result is None:
    st.stop()

rows = result["rows"]
detected_key = result["supplier_key"]

st.success(
    f"辨識供應商：**{supplier_label(detected_key)}** "
    f"（代碼：`{detected_key or '無法辨識'}`）"
)

# ---------------------------------------------------------------------------
# 依 invoice_no 分組：一般供應商只有一組 (一張帳單一個 invoice_no)，
# multi_page 供應商 (一份 PDF 可能塞好幾張獨立發票) 會自動依 invoice_no
# 分成好幾組。同一張發票內，抬頭欄位 (發票日期、供應商名稱...) 只會有一
# 個值，只有費用名稱／金額會逐筆不同，所以抬頭欄位只顯示、核對一次，
# 底下才依品項逐筆列出費用名稱／金額。
# ---------------------------------------------------------------------------
seen_invoice: "dict[str, list[int]]" = {}
invoice_order: "list[str]" = []
for idx, row in enumerate(rows):
    inv_no = row.get("invoice_no") or NA
    if inv_no not in seen_invoice:
        seen_invoice[inv_no] = []
        invoice_order.append(inv_no)
    seen_invoice[inv_no].append(idx)
invoice_groups = [(inv_no, seen_invoice[inv_no]) for inv_no in invoice_order]

# ---------------------------------------------------------------------------
# ② 逐欄人工核對
# ---------------------------------------------------------------------------
st.subheader("② 逐欄人工核對（沒有標記 = 視為忽略，不會出現在回饋記錄）")

marked_count = 0
total_fields = 0


def _field_row(label: str, value, err_key: str, fix_key: str) -> None:
    global marked_count, total_fields
    total_fields += 1
    col_label, col_value, col_flag, col_fix = st.columns([2, 3, 2, 3])
    col_label.markdown(f"<span style='font-size:16px; font-weight:700;'>{label}</span>", unsafe_allow_html=True)
    # 用固定 16px 的 span 顯示擷取值，不要用 st.write() 直接印數字/文字，
    # 避免金額 (數字) 跟其他欄位 (文字) 顯示出來的字體大小不一致。
    col_value.markdown(f"<span style='font-size:16px;'>{value}</span>", unsafe_allow_html=True)
    marked = col_flag.checkbox("❌ 標記為錯誤", key=err_key)
    if marked:
        marked_count += 1
        col_fix.text_input(
            "✏️ 正確答案（留空代表正確答案應為 N/A）",
            key=fix_key,
            label_visibility="visible",
        )
    else:
        col_fix.write("")


for inv_no, idxs in invoice_groups:
    st.markdown(f"#### 📄 發票號碼：{inv_no}")
    header_row = rows[idxs[0]]

    for code, label in HEADER_DISPLAY_FIELDS:
        _field_row(
            label,
            header_row.get(code, NA),
            f"err_{file_signature}_{inv_no}_{code}",
            f"fix_{file_signature}_{inv_no}_{code}",
        )

    # 直接列出這張發票「所有」費用項目跟金額，並且直接在這裡放「標記為
    # 錯誤」的核對欄位，不用另外收合/展開才看得到，一份 PDF 帳單可能有
    # 好幾張發票、每張發票又有好幾筆費用，這樣使用者一眼就能看到全部
    # 品項並直接核對，不用多一層點擊。
    st.markdown(f"**費用項目與金額一覽（共 {len(idxs)} 筆）**")
    for pos, row_idx in enumerate(idxs, start=1):
        row = rows[row_idx]
        for code, label in ITEM_DISPLAY_FIELDS:
            _field_row(
                label,
                row.get(code, NA),
                f"err_{file_signature}_{row_idx}_{code}",
                f"fix_{file_signature}_{row_idx}_{code}",
            )
    st.divider()

st.info(f"目前已標記 **{marked_count}** / {total_fields} 個欄位為錯誤，其餘視為忽略。")

# 全部欄位都正確 (沒有任何標記) 的話，要求使用者按下「確認無誤」才能繼續
# 產生執行結果，避免使用者根本還沒看完就直接下載。標記過至少一個錯誤，
# 則視為已經完成核對，不需要再另外確認。
confirm_key = f"confirmed_{file_signature}"
if marked_count == 0:
    if st.button("✅ 確認無誤（此帳單所有欄位都正確）"):
        st.session_state[confirm_key] = True
        st.rerun()

ready_for_export = marked_count > 0 or st.session_state.get(confirm_key, False)

# ---------------------------------------------------------------------------
# ③ 下載執行結果 / 回饋記錄
# ---------------------------------------------------------------------------
st.subheader("③ 下載執行結果")

if not ready_for_export:
    st.caption("請先核對以上欄位：有錯誤請標記並填入正確答案；全部正確請按上方「✅ 確認無誤」。")


def _resolve_final_value(code: str, row_idx: int, inv_no: str, row: dict, marks: dict):
    """算出這個欄位「最終」該輸出的值：如果被標記為錯誤，改用人工填的正確
    答案；沒有標記就沿用系統擷取值。品項欄位 (description/amount) 用
    row_idx 當 key、抬頭欄位用 inv_no 當 key，跟畫面上核對時用的 key
    規則完全一致。
    """
    item_err_key = f"err_{file_signature}_{row_idx}_{code}"
    if marks.get(item_err_key):
        return marks.get(f"fix_{file_signature}_{row_idx}_{code}", "") or "N/A"
    header_err_key = f"err_{file_signature}_{inv_no}_{code}"
    if marks.get(header_err_key):
        return marks.get(f"fix_{file_signature}_{inv_no}_{code}", "") or "N/A"
    return row.get(code, NA)


def _build_result_excel(rows_: list, invoice_groups_: list, marks: dict, supplier_key_: str) -> bytes:
    """把人工核對後的「最終」資料 (擷取值，被標記錯誤的欄位改用人工輸入
    的正確答案) 輸出成跟使用者提供的 Excel 範本一樣的格式：表頭在第1列、
    從 A 欄開始依標準欄位順序排列，一張工作表對應這份帳單的供應商，
    可以直接拿來當作這家供應商之後的「正確答案 Excel」使用。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = (supplier_key_ or "未知供應商")[:31]

    bold = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="DDEBF7")
    wrap_center = Alignment(wrap_text=True, vertical="center", horizontal="center")

    for idx, code in enumerate(FIELD_CODES, start=1):
        cell = ws.cell(row=1, column=idx, value=DISPLAY_HEADERS[code])
        cell.font = bold
        cell.fill = header_fill
        cell.alignment = wrap_center
        ws.column_dimensions[get_column_letter(idx)].width = 20

    r = 2
    for inv_no, idxs in invoice_groups_:
        for row_idx in idxs:
            row_ = rows_[row_idx]
            for idx, code in enumerate(FIELD_CODES, start=1):
                ws.cell(row=r, column=idx, value=_resolve_final_value(code, row_idx, inv_no, row_, marks))
            r += 1

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


if ready_for_export:
    result_bytes = _build_result_excel(rows, invoice_groups, st.session_state, detected_key)
    st.download_button(
        "⬇️ 下載執行結果 (Excel)",
        data=result_bytes,
        file_name=f"執行結果_{uploaded_pdf.name.rsplit('.', 1)[0]}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption(
        "執行結果 Excel 是核對後的「最終」資料 (標記錯誤的欄位已換成您填的正確答案)，"
        "欄位順序跟範本一致，可直接拿來當作這家供應商之後的正確答案 Excel。"
    )

st.subheader("④ 下載回饋記錄")


def _build_feedback_excel(file_name: str, supplier_lbl: str, supplier_key_: str,
                           rows_: list, invoice_groups_: list, marks: dict) -> bytes:
    """把「被標記為錯誤的欄位 + 使用者填的正確答案」整理成一份 Excel，
    欄位：檔案名稱／供應商／發票號碼／品項序號／欄位代碼／欄位名稱／
    系統擷取值／人工輸入的正確答案，方便日後依此修正 suppliers.py。
    抬頭欄位 (整張發票共用) 的「品項序號」欄填「抬頭欄位」，跟逐筆的費用
    明細區分開來。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "回饋記錄"

    headers = ["檔案名稱", "供應商", "供應商代碼", "發票號碼", "品項序號",
               "欄位代碼", "欄位名稱", "系統擷取值", "人工輸入正確答案"]
    bold = Font(bold=True)
    fill = PatternFill("solid", fgColor="DDEBF7")
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = bold
        cell.fill = fill
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[get_column_letter(c)].width = 20

    r = 2
    for inv_no, idxs in invoice_groups_:
        header_row = rows_[idxs[0]]
        for code, label in HEADER_DISPLAY_FIELDS:
            err_key = f"err_{file_signature}_{inv_no}_{code}"
            fix_key = f"fix_{file_signature}_{inv_no}_{code}"
            if not marks.get(err_key):
                continue
            correct_value = marks.get(fix_key, "") or "N/A"
            ws.cell(row=r, column=1, value=file_name)
            ws.cell(row=r, column=2, value=supplier_lbl)
            ws.cell(row=r, column=3, value=supplier_key_ or "")
            ws.cell(row=r, column=4, value=inv_no)
            ws.cell(row=r, column=5, value="抬頭欄位")
            ws.cell(row=r, column=6, value=code)
            ws.cell(row=r, column=7, value=label.split("\n")[-1])
            ws.cell(row=r, column=8, value=header_row.get(code, NA))
            ws.cell(row=r, column=9, value=correct_value)
            r += 1

        for pos, row_idx in enumerate(idxs, start=1):
            row_ = rows_[row_idx]
            for code, label in ITEM_DISPLAY_FIELDS:
                err_key = f"err_{file_signature}_{row_idx}_{code}"
                fix_key = f"fix_{file_signature}_{row_idx}_{code}"
                if not marks.get(err_key):
                    continue
                correct_value = marks.get(fix_key, "") or "N/A"
                ws.cell(row=r, column=1, value=file_name)
                ws.cell(row=r, column=2, value=supplier_lbl)
                ws.cell(row=r, column=3, value=supplier_key_ or "")
                ws.cell(row=r, column=4, value=inv_no)
                ws.cell(row=r, column=5, value=pos)
                ws.cell(row=r, column=6, value=code)
                ws.cell(row=r, column=7, value=label)
                ws.cell(row=r, column=8, value=row_.get(code, NA))
                ws.cell(row=r, column=9, value=correct_value)
                r += 1

    if r == 2:
        ws.cell(row=2, column=1, value="（目前沒有任何欄位被標記為錯誤）")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


if marked_count == 0:
    st.caption("目前沒有任何欄位被標記為錯誤，標記完成後即可在這裡下載回饋記錄 Excel。")
else:
    feedback_bytes = _build_feedback_excel(
        uploaded_pdf.name, supplier_label(detected_key), detected_key,
        rows, invoice_groups, st.session_state,
    )
    st.download_button(
        "⬇️ 下載回饋記錄 (Excel)",
        data=feedback_bytes,
        file_name=f"回饋記錄_{uploaded_pdf.name.rsplit('.', 1)[0]}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption(
        "把這份 Excel 交給負責維護「供應商帳單辨識對比系統」的開發端 "
        "(或直接請 Claude 依照 add-invoice-supplier 的流程) 修正對應供應商的擷取規則，"
        "改完後回到這個網頁重新上傳同一份 PDF，確認錯誤欄位已修正。"
    )
