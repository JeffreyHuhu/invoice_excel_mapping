# -*- coding: utf-8 -*-
"""
extract_utils.py
=================
供應商 (DWIHARTA) 海運帳單 PDF -> 結構化資料 的共用解析模組。

這支模組被兩支程式共用：
  1. app.py      (Step1：PDF -> Excel 轉檔小工具，Streamlit 網頁程式)
  2. verify.py   (Step2：查核比對小工具，檢查 Excel 是否忠實反映 PDF 內容)

設計理念：
  - 帳單為「可編輯的 PDF」(非掃描圖檔)，因此用 pdfplumber 直接抽取文字，
    不需要 OCR。若日後供應商改用掃描檔，可在 extract_text_from_pdf()
    內加入 OCR fallback (例如 pytesseract)。
  - 抬頭欄位 (Invoice No、B/L No...) 用正規表示式從整段文字中截取。
  - 費用明細 (Description / Amount) 用逐行比對抽取，一張帳單可能有多筆
    費用，最終會展開成多列 (每筆費用一列)，並把抬頭欄位重複填在每一列，
    這樣才符合 Excel 範本「一列 = 一筆費用」的結構。
"""

import re
import datetime
from dataclasses import dataclass, field
from typing import List, Dict, Optional

import pdfplumber


# ---------------------------------------------------------------------------
# 1. 基本工具函式
# ---------------------------------------------------------------------------

# 印尼文月份對照表，用來把 "04 Agustus 2026" 這種日期字串轉成 Python date
INDO_MONTHS = {
    "januari": 1, "februari": 2, "maret": 3, "april": 4,
    "mei": 5, "juni": 6, "juli": 7, "agustus": 8,
    "september": 9, "oktober": 10, "november": 11, "desember": 12,
}


def parse_indo_date(text: Optional[str]) -> Optional[datetime.date]:
    """把 '04 Agustus 2026' 這類印尼文日期字串轉成 datetime.date。
    若解析失敗，回傳 None (呼叫端可以選擇保留原始字串)。
    """
    if not text:
        return None
    text = text.strip()
    m = re.match(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", text)
    if not m:
        return None
    day, month_name, year = m.groups()
    month = INDO_MONTHS.get(month_name.lower())
    if not month:
        return None
    try:
        return datetime.date(int(year), month, int(day))
    except ValueError:
        return None


def clean_amount(text: str) -> Optional[float]:
    """把 '5,820,000' 這種千分位字串轉成 float。抓不到時回傳 None。"""
    if text is None:
        return None
    text = text.replace(",", "").replace(".", "").strip()
    if not text.isdigit():
        return None
    return float(text)


def extract_text_from_pdf(pdf_path: str) -> str:
    """讀取 PDF 全部頁面的文字並合併成一個字串。

    帳單是「可編輯 PDF」，pdfplumber 可以直接抓到文字層。
    若遇到沒有文字層的掃描檔，這裡會回傳空字串，
    使用端應提示使用者改用 OCR 工具或人工輸入。
    """
    full_text_parts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            full_text_parts.append(page_text)
    return "\n".join(full_text_parts)


# ---------------------------------------------------------------------------
# 2. 資料結構
# ---------------------------------------------------------------------------

@dataclass
class LineItem:
    """一筆帳單費用明細"""
    description: str
    amount_idr: Optional[float]
    vat: Optional[float] = None


@dataclass
class InvoiceHeader:
    """帳單抬頭欄位，對應 Excel 範本 B~E、H~Q 欄"""
    invoice_date: Optional[datetime.date] = None
    invoice_date_raw: str = ""
    invoice_no: str = ""
    supplier: str = ""          # D欄 Payment can be Transfer to
    consignee: str = ""         # E欄 TO
    order_no: str = ""
    arrive_date: Optional[datetime.date] = None
    arrive_date_raw: str = ""
    onboard_date: Optional[datetime.date] = None
    onboard_date_raw: str = ""
    bl_no: str = ""
    mbl_no: str = ""
    port_of_loading: str = ""
    port_of_discharge: str = ""
    volume: str = ""
    vessel: str = ""
    container_no: str = ""
    source_file: str = ""


@dataclass
class ParsedInvoice:
    header: InvoiceHeader
    line_items: List[LineItem] = field(default_factory=list)
    raw_text: str = ""


# ---------------------------------------------------------------------------
# 3. 正規表示式規則 (抬頭欄位)
# ---------------------------------------------------------------------------
#
# 重要背景：這份帳單是「左右兩欄並排」的版面 (例如同一列左邊是
# "Order No. : IMPJK26070038"，右邊緊接著是 "Port Of Loading : SHANGHAI, CHINA")。
# pdfplumber 抽取文字時，同一個 y 座標高度的內容會被接成同一行字串，
# 所以如果只用「抓到冒號後面所有文字」的寫法，會把右欄欄位的文字也一併抓進來。
#
# 解法：每個欄位的值只抓到「下一個已知欄位名稱出現之前」或「連續兩個以上空白
# (欄與欄之間的排版留白)」或「換行」為止，三個條件任一成立就停止擷取。

# 所有帳單上可能出現的欄位/標題關鍵字，用來當作「這裡是下一欄，該停止擷取」的判斷依據
_LABEL_WORDS = [
    r"Invoice\s*No\.?", r"Invoice\s*Date", r"Shipment\s*Type", r"Page",
    r"Order\s*No\.?", r"Arrive\s*Date", r"On\s*[Bb]oard\s*Date",
    r"B/L\s*No\.?", r"MBL\s*No\.?", r"Shipper", r"Consignee",
    r"Port\s*Of\s*Loading", r"Port\s*Of\s*Discharge", r"Reference",
    r"Volume", r"Vessel", r"Container\s*No\.?", r"Payment\s*Term",
    r"Payment\s*can\s*be\s*Transfer\s*to", r"Bank\s*Information",
    r"I\s*N\s*V\s*O\s*I\s*C\s*E", r"Sub\s*Total", r"GRAND\s*TOTAL",
    r"V\s*A\s*T", r"E\.?\s*&\s*O\.?\s*E", r"Halaman", r"\bTO\b",
]
_NEXT_LABEL_ALT = "(?:" + "|".join(_LABEL_WORDS) + ")"

# 停止條件：遇到「連續2個以上空白」或「換行」或「字串結尾」或「下一個已知欄位名稱」
_STOP_LOOKAHEAD = rf"(?=\s{{2,}}|\n|$|\s*{_NEXT_LABEL_ALT}\b)"

# 每一個 key 對應 InvoiceHeader 的欄位名稱，value 是抓「冒號後面那段文字，
# 直到遇到下一欄位為止」的 pattern
HEADER_PATTERNS: Dict[str, str] = {
    "invoice_no":         rf"Invoice\s*No\.?\s*:\s*(.+?){_STOP_LOOKAHEAD}",
    "invoice_date_raw":   rf"Invoice\s*Date\s*:\s*(.+?){_STOP_LOOKAHEAD}",
    "order_no":           rf"Order\s*No\.?\s*:\s*(.+?){_STOP_LOOKAHEAD}",
    "arrive_date_raw":    rf"Arrive\s*Date\s*:\s*(.+?){_STOP_LOOKAHEAD}",
    "onboard_date_raw":   rf"On\s*[Bb]oard\s*Date\s*:\s*(.+?){_STOP_LOOKAHEAD}",
    "bl_no":              rf"\bB/L\s*No\.?\s*:\s*(.+?){_STOP_LOOKAHEAD}",
    "mbl_no":              rf"\bMBL\s*No\.?\s*:\s*(.+?){_STOP_LOOKAHEAD}",
    "port_of_loading":    rf"Port\s*Of\s*Loading\s*:\s*(.+?){_STOP_LOOKAHEAD}",
    "port_of_discharge":  rf"Port\s*Of\s*Discharge\s*:\s*(.+?){_STOP_LOOKAHEAD}",
    "volume":             rf"Volume\s*:\s*(.+?){_STOP_LOOKAHEAD}",
    "vessel":              rf"Vessel\s*:\s*(.+?){_STOP_LOOKAHEAD}",
    "container_no":       rf"Container\s*No\.?\s*:\s*(.+?){_STOP_LOOKAHEAD}",
    "consignee":           rf"\bTO\s*:\s*(.+?){_STOP_LOOKAHEAD}",
}


def clean_volume(value: str) -> str:
    """Volume 欄位常見格式是 'FCL - 1x40HC' 或 'LCL - 5 CBM'，
    其中 'FCL -' / 'LCL -' 只是裝櫃方式的前綴，實際要的是後面的櫃型/數量，
    所以這裡把該前綴去掉，只保留 '1x40HC' 這種實際資訊。
    """
    if not value:
        return value
    return re.sub(r"^(FCL|LCL)\s*-\s*", "", value.strip(), flags=re.IGNORECASE).strip()

# 費用明細行：例如
#   "JASA PPJK/FORWARDER 1 IDR 300,000 300,000 33,000"
#   "TRUCKING CHARGES 40 1 IDR 5,820,000 5,820,000"
# 格式：描述(可含數字/符號) 數量 幣別 單價 金額 [VAT]
LINE_ITEM_PATTERN = re.compile(
    r"^(?P<desc>.+?)\s+(?P<qty>\d+)\s+(?P<curr>IDR|USD)\s+"
    r"(?P<rate>[\d,]+)\s+(?P<amount>[\d,]+)"
    r"(?:\s+(?P<vat>[\d,]+))?\s*$"
)

# 用來判斷哪些行不是費用明細 (表頭行、備註行等)，避免誤判
LINE_ITEM_EXCLUDE_KEYWORDS = (
    "Description", "Sub Total", "GRAND TOTAL", "V A T", "Payment Term",
)


# ---------------------------------------------------------------------------
# 4. 主要解析函式
# ---------------------------------------------------------------------------

def _extract_header(text: str, source_file: str) -> InvoiceHeader:
    header = InvoiceHeader(source_file=source_file)

    for field_name, pattern in HEADER_PATTERNS.items():
        m = re.search(pattern, text, flags=re.MULTILINE)
        if m:
            value = m.group(1).strip().rstrip(",").strip()
            setattr(header, field_name, value)

    # Volume 欄位去掉 'FCL -' / 'LCL -' 前綴，只保留櫃型/數量本身
    header.volume = clean_volume(header.volume)

    # 供應商名稱：取自 "Payment can be Transfer to :" 之後的第一個非空白行
    m = re.search(r"Payment can be Transfer to\s*:?\s*\n\s*(.+)", text)
    if m:
        supplier_line = m.group(1).strip()
        # 同樣防止右欄或下一個標籤被誤黏進來
        stop_m = re.search(_STOP_LOOKAHEAD, supplier_line)
        if stop_m:
            supplier_line = supplier_line[:stop_m.start()]
        header.supplier = supplier_line.strip()

    # 日期字串轉成日期物件（若解析失敗仍保留原始字串）
    header.invoice_date = parse_indo_date(header.invoice_date_raw)
    header.arrive_date = parse_indo_date(header.arrive_date_raw)
    header.onboard_date = parse_indo_date(header.onboard_date_raw)

    return header


def _extract_line_items(text: str) -> List[LineItem]:
    items: List[LineItem] = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if any(kw in line for kw in LINE_ITEM_EXCLUDE_KEYWORDS):
            continue
        m = LINE_ITEM_PATTERN.match(line)
        if not m:
            continue
        desc = m.group("desc").strip()
        amount = clean_amount(m.group("amount"))
        vat = clean_amount(m.group("vat")) if m.group("vat") else None
        items.append(LineItem(description=desc, amount_idr=amount, vat=vat))
    return items


def parse_invoice_pdf(pdf_path: str) -> ParsedInvoice:
    """對外主要入口：讀取一份 PDF 帳單，回傳結構化的 ParsedInvoice。"""
    text = extract_text_from_pdf(pdf_path)
    header = _extract_header(text, source_file=pdf_path)
    line_items = _extract_line_items(text)
    return ParsedInvoice(header=header, line_items=line_items, raw_text=text)


# ---------------------------------------------------------------------------
# 5. 轉成 Excel 範本列 (dict list)，欄位對應 資料模板-DWIHARTA.xlsx 的 B~Q 欄
# ---------------------------------------------------------------------------

# 欄位代碼對照 Excel 範本第一列「抓取欄位」名稱，方便 app.py / verify.py 共用
TEMPLATE_COLUMNS = [
    "Invoice Date", "Invoice No", "Payment can be Transfer to", "TO",
    "Description\nVAT", "Amount (IDR", "Order No.", "Arrive Date",
    "on board date", "B/L NO.", "MBL NO.", "Port of Loading",
    "Port of discharge", "Volume", "Vessel", "Container no.",
]


def invoice_to_rows(parsed: ParsedInvoice) -> List[Dict]:
    """把一份解析完的帳單展開成多列 (一筆費用一列)，
    每一列是一個 dict，key 對應 TEMPLATE_COLUMNS。
    若帳單完全沒抓到費用明細，仍會輸出一列（費用欄留空），避免整張帳單被漏掉。
    """
    h = parsed.header
    common = {
        "Invoice Date": h.invoice_date or h.invoice_date_raw,
        "Invoice No": h.invoice_no,
        "Payment can be Transfer to": h.supplier,
        "TO": h.consignee,
        "Order No.": h.order_no,
        "Arrive Date": h.arrive_date or h.arrive_date_raw,
        "on board date": h.onboard_date or h.onboard_date_raw,
        "B/L NO.": h.bl_no,
        "MBL NO.": h.mbl_no,
        "Port of Loading": h.port_of_loading,
        "Port of discharge": h.port_of_discharge,
        "Volume": h.volume,
        "Vessel": h.vessel,
        "Container no.": h.container_no,
    }

    rows = []
    items = parsed.line_items or [LineItem(description="", amount_idr=None)]
    for item in items:
        row = dict(common)
        row["Description\nVAT"] = item.description
        row["Amount (IDR"] = item.amount_idr
        rows.append(row)
    return rows
