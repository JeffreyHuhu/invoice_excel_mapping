# -*- coding: utf-8 -*-
"""
suppliers.py
============
各供應商專屬的「辨識規則」與「擷取規則」，全部在這支檔案裡註冊到
extract_utils.SUPPLIER_REGISTRY。

新增第 21 家供應商的作法：
  1. 在下面仿照 DWIHARTA / KUEHNE_NAGEL 的寫法，寫一個
       def detect_xxx(text: str) -> bool: ...
     和一個
       def parse_xxx(text: str) -> dict: ...
     parse_xxx() 要回傳 {"header": {...14個抬頭欄位...},
                          "items": [{"description":.., "amount":..}, ...]}
  2. 檔案最後呼叫一次
       register_supplier("XXX_KEY", "顯示名稱", detect_xxx, parse_xxx)
  3. 不用改 extract_utils.py 或 app.py，系統會自動把新供應商加入辨識清單。

detect() 的判斷順序很重要：越有可能誤判的規則(太寬鬆的關鍵字)要放在
register_supplier() 呼叫順序的後面，因為 detect_supplier() 是依序詢問、
第一個回傳 True 的就採用。
"""

import re
import datetime
from typing import Dict, List, Optional

from extract_utils import register_supplier

# 印尼文月份對照表 (DWIHARTA 帳單用印尼文寫月份)
INDO_MONTHS = {
    "januari": 1, "februari": 2, "maret": 3, "april": 4,
    "mei": 5, "juni": 6, "juli": 7, "agustus": 8,
    "september": 9, "oktober": 10, "november": 11, "desember": 12,
}


def _search(pattern: str, text: str) -> Optional[str]:
    m = re.search(pattern, text, flags=re.MULTILINE)
    if not m:
        return None
    return m.group(1).strip().rstrip(",").strip()


# ===========================================================================
# 供應商 1：DWIHARTA (海運帳單，Invoice + 費用明細表格)
# ===========================================================================
#
# 版面特徵：左右兩欄並排 (例如同一列左邊 "Order No. : IMPJK26070038"，
# 右邊緊接著 "Port Of Loading : SHANGHAI, CHINA")。pdfplumber 抽取文字時，
# 同一 y 座標高度的內容會被接成同一行字串，所以每個欄位都用「遇到下一個
# 已知欄位關鍵字、或連續兩個以上空白、或換行就停止擷取」的方式抓值，
# 避免把右欄內容一起抓進來。

def detect_dwiharta(text: str) -> bool:
    return "DWIHARTA" in text.upper()


_DWIHARTA_LABEL_WORDS = [
    r"Invoice\s*No\.?", r"Invoice\s*Date", r"Shipment\s*Type", r"Page",
    r"Order\s*No\.?", r"Arrive\s*Date", r"On\s*[Bb]oard\s*Date",
    r"B/L\s*No\.?", r"MBL\s*No\.?", r"Shipper", r"Consignee",
    r"Port\s*Of\s*Loading", r"Port\s*Of\s*Discharge", r"Reference",
    r"Volume", r"Vessel", r"Container\s*No\.?", r"Payment\s*Term",
    r"Payment\s*can\s*be\s*Transfer\s*to", r"Bank\s*Information",
    r"I\s*N\s*V\s*O\s*I\s*C\s*E", r"Sub\s*Total", r"GRAND\s*TOTAL",
    r"V\s*A\s*T", r"E\.?\s*&\s*O\.?\s*E", r"Halaman", r"\bTO\b",
]
_DWIHARTA_NEXT = "(?:" + "|".join(_DWIHARTA_LABEL_WORDS) + ")"
_DWIHARTA_STOP = rf"(?=\s{{2,}}|\n|$|\s*{_DWIHARTA_NEXT}\b)"

_DWIHARTA_HEADER_PATTERNS: Dict[str, str] = {
    "invoice_no":         rf"Invoice\s*No\.?\s*:\s*(.+?){_DWIHARTA_STOP}",
    "invoice_date":       rf"Invoice\s*Date\s*:\s*(.+?){_DWIHARTA_STOP}",
    "order_no":           rf"Order\s*No\.?\s*:\s*(.+?){_DWIHARTA_STOP}",
    "arrive_date":        rf"Arrive\s*Date\s*:\s*(.+?){_DWIHARTA_STOP}",
    "onboard_date":       rf"On\s*[Bb]oard\s*Date\s*:\s*(.+?){_DWIHARTA_STOP}",
    "bl_no":              rf"\bB/L\s*No\.?\s*:\s*(.+?){_DWIHARTA_STOP}",
    "mbl_no":             rf"\bMBL\s*No\.?\s*:\s*(.+?){_DWIHARTA_STOP}",
    "port_of_loading":    rf"Port\s*Of\s*Loading\s*:\s*(.+?){_DWIHARTA_STOP}",
    "port_of_discharge":  rf"Port\s*Of\s*Discharge\s*:\s*(.+?){_DWIHARTA_STOP}",
    "volume":             rf"Volume\s*:\s*(.+?){_DWIHARTA_STOP}",
    "vessel":             rf"Vessel\s*:\s*(.+?){_DWIHARTA_STOP}",
    "container_no":       rf"Container\s*No\.?\s*:\s*(.+?){_DWIHARTA_STOP}",
    "consignee":          rf"\bTO\s*:\s*(.+?){_DWIHARTA_STOP}",
}

_DWIHARTA_LINE_ITEM_PATTERN = re.compile(
    r"^(?P<desc>.+?)\s+(?P<qty>\d+)\s+(?P<curr>IDR|USD)\s+"
    r"(?P<rate>[\d,]+)\s+(?P<amount>[\d,]+)"
    r"(?:\s+(?P<vat>[\d,]+))?\s*$"
)
_DWIHARTA_EXCLUDE_KEYWORDS = ("Description", "Sub Total", "GRAND TOTAL", "V A T", "Payment Term")


def _dwiharta_clean_volume(value: Optional[str]) -> Optional[str]:
    """Volume 欄位常見寫法 'FCL - 1x40HC'，'FCL -' 只是裝櫃方式前綴，
    實際要的是後面的櫃型/數量，這裡去掉前綴只留 '1x40HC'。
    """
    if not value:
        return value
    return re.sub(r"^(FCL|LCL)\s*-\s*", "", value.strip(), flags=re.IGNORECASE).strip()


def _dwiharta_parse_indo_date(text: Optional[str]) -> Optional[str]:
    """把 '04 Agustus 2026' 轉成 'DD.MM.YYYY'，跟其他供應商日期格式統一，
    方便比對時不因為日期寫法不同而誤判不相符。抓不到就回傳原始字串。
    """
    if not text:
        return text
    m = re.match(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", text.strip())
    if not m:
        return text
    day, month_name, year = m.groups()
    month = INDO_MONTHS.get(month_name.lower())
    if not month:
        return text
    return f"{int(day):02d}.{month:02d}.{year}"


def parse_dwiharta(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {}
    for code, pattern in _DWIHARTA_HEADER_PATTERNS.items():
        header[code] = _search(pattern, text)

    header["volume"] = _dwiharta_clean_volume(header.get("volume"))
    for date_code in ("invoice_date", "arrive_date", "onboard_date"):
        header[date_code] = _dwiharta_parse_indo_date(header.get(date_code))

    # 供應商名稱：取自 "Payment can be Transfer to :" 之後的第一個非空白行
    m = re.search(r"Payment can be Transfer to\s*:?\s*\n\s*(.+)", text)
    if m:
        supplier_line = m.group(1).strip()
        stop_m = re.search(_DWIHARTA_STOP, supplier_line)
        if stop_m:
            supplier_line = supplier_line[:stop_m.start()]
        header["supplier"] = supplier_line.strip()

    items: List[Dict] = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line or any(kw in line for kw in _DWIHARTA_EXCLUDE_KEYWORDS):
            continue
        m = _DWIHARTA_LINE_ITEM_PATTERN.match(line)
        if not m:
            continue
        amount_str = m.group("amount").replace(",", "")
        items.append({
            "description": m.group("desc").strip(),
            "amount": int(amount_str) if amount_str.isdigit() else None,
        })

    return {"header": header, "items": items}


# ===========================================================================
# 供應商 2：KUEHNE NAGEL (Faktur Pajak 稅務發票 + KN Sales Invoice 混合檔)
# ===========================================================================
#
# 這家供應商的 PDF 跟 DWIHARTA 完全不同版型：
#   - 第1頁是印尼稅務局格式的「Faktur Pajak」(稅務發票)，供應商名稱、
#     費用名稱(Description)、金額(Amount) 是從這一頁抓的。
#   - 第2頁起是 Kuehne Nagel 自己的「KN Sales Invoice」，抬頭欄位
#     (Invoice No/Date、日期、港口、船名、追蹤號碼...) 是從這幾頁抓的，
#     版面同樣左右兩欄並排，沿用跟 DWIHARTA 一樣的「遇到下一個已知欄位
#     就停止」策略，只是欄位關鍵字換成 Kuehne Nagel 帳單上的英文標籤。
#   - 這種帳單沒有海運提單，改用「KN TRACKING NO.」頂替 B/L NO. 欄位；
#     MBL NO. 與 Container no. 本來就沒有對應資訊，缺漏後續會自動補 N/A。

def detect_kuehne_nagel(text: str) -> bool:
    upper = text.upper()
    return "KUEHNE" in upper and "NAGEL" in upper


_KN_LABEL_WORDS = [
    r"INVOICE TO", r"SALES INVOICE", r"KN TRACKING NO\.?", r"KN ACCOUNTING NO\.?",
    r"ACCOUNT NO\.?", r"INVOICE NO\.?\s*/\s*DATE", r"DUE DATE", r"CONSIGNEE", r"SHIPPER",
    r"NOTIFY", r"FIRST VESSEL NA", r"VESSEL NAME", r"FIRST VOYAGE NU", r"VOYAGE",
    r"PL\.\s*OF RECEIPT", r"MOVEMENT TYPE", r"P\.\s*OF LOADING", r"ETD/ATD",
    r"P\.\s*OF DISCHARGE", r"ETA/ATA", r"PL\.\s*OF DELIVERY", r"DANGEROUS GOODS",
    r"TERMS OF TRADE", r"INSURAN\.\s*STATUS", r"MARKS\s*&\s*NOS", r"CHARGE",
    r"GTN SHIPPING ORDER NUMBER", r"COMMERCIAL INVOICE NUMBER", r"ESI TOKEN",
]
_KN_NEXT = "(?:" + "|".join(_KN_LABEL_WORDS) + ")"
_KN_STOP = rf"(?=\s{{2,}}|\n|$|\s*{_KN_NEXT}\b)"


def _kn_clean_amount(text: Optional[str]) -> Optional[int]:
    """印尼式數字 '1.521.098,00' (句點=千分位、逗號=小數) -> 1521098。"""
    if not text:
        return None
    text = text.strip().replace(".", "")
    text = text.split(",")[0]
    return int(text) if text.isdigit() else None


def _kn_clean_port(value: Optional[str]) -> Optional[str]:
    """'LOS ANGELES, CA' -> 'LOS ANGELES'：只取城市名，逗號後的州/國碼截掉。"""
    if not value:
        return value
    return value.split(",")[0].strip()


def _kn_clean_tracking_no(value: Optional[str]) -> Optional[str]:
    """'1076 510 528' -> '1076510528'：把數字群組間的空白去掉。"""
    if not value:
        return value
    return re.sub(r"\s+", "", value.strip())


def parse_kuehne_nagel(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {}

    header["supplier"] = _search(r"Pengusaha Kena Pajak:\s*\n\s*Nama\s*:\s*(.+)", text)
    description = _search(r"\n([A-Za-z][^\n]+)\nRp\s+[\d.,]+\s*x", text)
    amount_raw = _search(
        r"Harga Jual\s*/\s*Penggantian\s*/\s*Uang Muka\s*/\s*Termin\s+([\d.,]+)", text
    )

    header["consignee"] = _search(rf"INVOICE TO\s*\n\s*(.+?){_KN_STOP}", text)

    m = re.search(r"INVOICE NO\.\s*/\s*DATE\s+(\S+)\s+(\d{1,2}\.\d{1,2}\.\d{4})", text)
    if m:
        header["invoice_no"], header["invoice_date"] = m.group(1), m.group(2)

    header["bl_no"] = _kn_clean_tracking_no(_search(rf"KN TRACKING NO\.\s*(.+?){_KN_STOP}", text))
    header["order_no"] = _search(rf"GTN SHIPPING ORDER NUMBER\s*\n\s*(.+?){_KN_STOP}", text)
    header["vessel"] = _search(rf"VESSEL NAME\s*:\s*(.+?){_KN_STOP}", text)
    header["port_of_loading"] = _kn_clean_port(_search(rf"P\.\s*OF LOADING\s*:\s*(.+?){_KN_STOP}", text))
    header["port_of_discharge"] = _kn_clean_port(_search(rf"P\.\s*OF DISCHARGE\s*:\s*(.+?){_KN_STOP}", text))
    header["onboard_date"] = _search(rf"ETD/ATD\s*:\s*(.+?){_KN_STOP}", text)
    header["arrive_date"] = _search(rf"ETA/ATA\s*:\s*(.+?){_KN_STOP}", text)

    m2 = re.search(r"AS PER ATTACHED\s+([\d,\.]+)\s+([\d,\.]+)", text)
    header["volume"] = m2.group(2) if m2 else None

    # 這種帳單本來就沒有海運提單/貨櫃資訊，留空給 expand_to_rows() 補 N/A
    header["mbl_no"] = None
    header["container_no"] = None

    items = [{"description": description, "amount": _kn_clean_amount(amount_raw)}]

    return {"header": header, "items": items}


# ===========================================================================
# 註冊所有供應商
# ===========================================================================
# detect_kuehne_nagel() 的關鍵字 "KUEHNE"+"NAGEL" 很具體不會誤判，
# detect_dwiharta() 同理只認 "DWIHARTA" 字樣，兩者互不衝突，順序不影響結果。
# 未來新增供應商時，關鍵字盡量挑「該公司獨有、不會出現在其他供應商帳單」
# 的字串 (公司全名、統編、帳單編號前綴等)。

register_supplier("DWIHARTA", "PT. DWIHARTA LOGISTINDO", detect_dwiharta, parse_dwiharta)
register_supplier("KUEHNE_NAGEL", "KUEHNE NAGEL INDONESIA", detect_kuehne_nagel, parse_kuehne_nagel)
