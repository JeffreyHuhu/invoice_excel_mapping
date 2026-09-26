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


# 費用明細行 (在 KN Sales Invoice 的 "CHARGE" 區塊裡)，例如：
#   "OVERTIME CHARGES EXPORT 85.00 USD 1,521,098 7"
# 格式：費用名稱 ... 外幣金額 USD 台幣/印尼幣金額 VAT代碼
# 這是「使用者真正要的 Description」，跟 Faktur Pajak (稅務發票) 上寫的
# 那種概括性文字 (例如 "Domestic transport & ancillary services") 是兩回事：
# 稅務發票的文字只是報稅用的統稱，帳單上實際列出的每筆費用名稱才是查核
# 要比對的欄位。這裡用「結尾是 USD 金額+金額+VAT代碼」的格式辨識這種行，
# 一張帳單若有多筆費用，會抓到多筆、各自展開成一列 (呼應 DWIHARTA
# 「一筆費用一列」的設計)。
_KN_CHARGE_LINE_PATTERN = re.compile(
    r"^(?P<desc>[A-Z][A-Za-z0-9 /\-\.,()&]+?)\s+[\d,]+\.\d{2}\s*USD\s+"
    r"(?P<amount>[\d,]+)\s+\d+\s*$",
    re.MULTILINE,
)


def _kn_clean_amount_comma_thousands(text: Optional[str]) -> Optional[int]:
    """英式數字 '1,521,098' (逗號=千分位，沒有小數) -> 1521098。

    注意：這跟 _kn_clean_amount() 是不同的格式！_kn_clean_amount() 是給
    Faktur Pajak (稅務發票) 上的印尼式數字用的 ('1.521.098,00'，句點千分位
    、逗號小數)；Sales Invoice 費用明細行上的金額則是英式數字 (逗號千分位)。
    兩種格式的千分位/小數符號剛好相反，混用會把 '1,521,098' 誤判成 '1'
    (把第一個逗號當成小數點切斷)，所以分開兩個函式，不要共用。
    """
    if not text:
        return None
    text = text.strip().replace(",", "")
    return int(text) if text.isdigit() else None


def _kn_extract_charge_items(text: str) -> List[Dict]:
    """從 Sales Invoice 的費用明細行抓出每一筆 {description, amount}。
    抓不到任何一行時回傳空 list，呼叫端會用 Faktur Pajak 的概括描述當備援，
    確保至少有一筆資料而不是整張都是 N/A。
    """
    items = []
    for m in _KN_CHARGE_LINE_PATTERN.finditer(text):
        amount = _kn_clean_amount_comma_thousands(m.group("amount"))
        items.append({"description": m.group("desc").strip(), "amount": amount})
    return items


def parse_kuehne_nagel(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {}

    header["supplier"] = _search(r"Pengusaha Kena Pajak:\s*\n\s*Nama\s*:\s*(.+)", text)

    header["consignee"] = _search(rf"INVOICE TO\s*\n\s*(.+?){_KN_STOP}", text)

    m = re.search(r"INVOICE NO\.\s*/\s*DATE\s+(\S+)\s+(\d{1,2}\.\d{1,2}\.\d{4})", text)
    if m:
        header["invoice_no"], header["invoice_date"] = m.group(1), m.group(2)

    header["bl_no"] = _kn_clean_tracking_no(_search(rf"KN TRACKING NO\.\s*(.+?){_KN_STOP}", text))
    # 注意：Order No. 欄位對應的是「COMMERCIAL INVOICE NUMBER」(例如
    # RYH0810074280)，不是緊接在旁邊的「GTN SHIPPING ORDER NUMBER」
    # (例如 VB26072003093319) —— 這兩個編號在版面上前後相鄰、很容易搞混，
    # 但使用者核對過的正確答案是以 COMMERCIAL INVOICE NUMBER 為準。
    header["order_no"] = _search(rf"COMMERCIAL INVOICE NUMBER\s*\n\s*(.+?){_KN_STOP}", text)
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

    items = _kn_extract_charge_items(text)
    if not items:
        # 備援：真的抓不到任何費用明細行時，退回用稅務發票 (Faktur Pajak)
        # 的概括描述+總金額，至少不要整張全空。
        description = _search(r"\n([A-Za-z][^\n]+)\nRp\s+[\d.,]+\s*x", text)
        amount_raw = _search(
            r"Harga Jual\s*/\s*Penggantian\s*/\s*Uang Muka\s*/\s*Termin\s+([\d.,]+)", text
        )
        items = [{"description": description, "amount": _kn_clean_amount(amount_raw)}]

    return {"header": header, "items": items}


# ===========================================================================
# 供應商 3：INDO PROSTIME EXPRESS (空運帳單 AIRFREIGHT INVOICE)
# ===========================================================================
#
# 版面特徵：跟 DWIHARTA/KUEHNE NAGEL 一樣是左右並排欄位，但這份帳單並排欄位
# 之間只用「一個空白」分隔 (不是兩個以上)，例如：
#   "MAWB : 160-1655 0811 FREIGHT : COLLECT"
#   "HAWB : PT20975640 QUANTITY : 1.00 PKGS"
# 所以「遇到連續兩個以上空白就停止」這招在這裡沒用，改成完全依賴「遇到下
# 一個已知欄位關鍵字就停止」(不要求前面有幾個空白)。另外「Number :」那一列
# 因為版面關係跟收件人地址欄混在一起 (例如
# "HARDASES ABADI INDONESIA,.PT Number : 101995")，所以 invoice_no/
# invoice_date 直接用「只抓數字/日期格式」的簡單寫法，不用停止清單。
#
# 這家帳單沒有「開船日」欄位，但 FLIGHT 欄位尾碼其實藏著日期 (例如
# 'CX 777/01092026' 代表 01.09.2026)，所以 onboard_date 是從已經抓到的
# FLIGHT(vessel) 值反推出來的，不是直接抓某個「欄位:」的值。

def detect_indoprostime(text: str) -> bool:
    return "PROSTIME" in text.upper()


_INDO_LABEL_WORDS = [
    r"A/C\s*No\.?", r"A/C\s*name", r"Swift\s*code", r"Bank",
    r"Number", r"Date", r"Order\s*No\.?", r"Currency\s*/\s*Rate",
    r"MAWB", r"HAWB", r"QUANTITY", r"SHIPPER", r"GROSS\s*/\s*KG",
    r"CNEE", r"CHARGE\s*/\s*KG", r"ORIGIN", r"FLIGHT", r"FREIGHT",
    r"PKGS", r"D\s*E\s*S\s*C\s*R\s*I\s*P\s*T\s*I\s*O\s*N", r"AMOUNT",
    r"Total", r"Cheque",
]
_INDO_NEXT = "(?:" + "|".join(_INDO_LABEL_WORDS) + ")"
# 注意：這裡不像 DWIHARTA/KN 要求 "\s{2,}"，因為並排欄位只隔一個空白；
# 只要「後面接著任何空白 + 下一個已知欄位關鍵字」就停止擷取。
_INDO_STOP = rf"(?=\s+{_INDO_NEXT}\b|\n|$)"


def _indo_add_space_after_pt(value: Optional[str]) -> Optional[str]:
    """'PT.HARDASES ABADI INDONESIA' -> 'PT. HARDASES ABADI INDONESIA'：
    PDF 版面把 'PT.' 跟公司名稱黏在一起、中間沒有空白，這裡補回一個空白，
    跟正確答案 Excel 裡的寫法對齊 (正確答案 'PT.' 後面有空白)。
    """
    if not value:
        return value
    return re.sub(r"^(PT\.)(?=\S)", r"\1 ", value.strip())


def _indo_parse_flight_date(vessel_value: Optional[str]) -> Optional[str]:
    """從 FLIGHT 欄位尾碼反推日期，例如 'CX 777/01092026' -> '01/09/2026'
    (尾碼 8 碼數字是 DDMMYYYY)。抓不到就回傳 None，讓 expand_to_rows()
    補 N/A。
    """
    if not vessel_value:
        return None
    m = re.search(r"(\d{2})(\d{2})(\d{4})\s*$", vessel_value.strip())
    if not m:
        return None
    day, month, year = m.groups()
    return f"{day}/{month}/{year}"


def _indo_clean_amount(text: Optional[str]) -> Optional[int]:
    """英式數字 '416,729' (逗號千分位) -> 416729。"""
    if not text:
        return None
    text = text.strip().replace(",", "")
    return int(text) if text.isdigit() else None


# 費用明細行，例如：
#   "FREIGHT CHARGE (USD 23.54) (REIMBURSEMENT) 416,729"
# 格式：全大寫的費用名稱 + 一個或兩個備註用括號 + 最後的金額。要求描述文字
# 全大寫 (不能有小寫字母)，是為了避免誤抓到版面上其他也帶括號跟數字的雜訊
# 行 (例如帳單最上方的 "Phone (021) 6268280")。
_INDO_CHARGE_LINE_PATTERN = re.compile(
    r"^(?P<desc>[A-Z][A-Z /]+?)\s*\(.+?\)(?:\s*\(.+?\))?\s+(?P<amount>[\d,]+)\s*$",
    re.MULTILINE,
)
# 備援：少數帳單可能沒有括號備註，只有「全大寫費用名稱 + 金額」，這裡再抓
# 一次，但要排除 "Total ..." 這種小計行 (不是實際費用明細)。
_INDO_SIMPLE_CHARGE_LINE_PATTERN = re.compile(
    r"^(?!Total\b)(?P<desc>[A-Z][A-Z /]+?)\s+(?P<amount>[\d,]+)\s*$",
    re.MULTILINE,
)


def _indo_extract_charge_items(text: str) -> List[Dict]:
    items = []
    for m in _INDO_CHARGE_LINE_PATTERN.finditer(text):
        items.append({
            "description": m.group("desc").strip(),
            "amount": _indo_clean_amount(m.group("amount")),
        })
    if not items:
        for m in _INDO_SIMPLE_CHARGE_LINE_PATTERN.finditer(text):
            items.append({
                "description": m.group("desc").strip(),
                "amount": _indo_clean_amount(m.group("amount")),
            })
    return items


def parse_indoprostime(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {}

    header["invoice_no"] = _search(r"Number\s*:\s*(\d+)", text)
    header["invoice_date"] = _search(r"Date\s*:\s*(\d{1,2}/\d{1,2}/\d{4})", text)
    header["supplier"] = _search(rf"A/C\s*name\s*:\s*(.+?){_INDO_STOP}", text)
    header["consignee"] = _indo_add_space_after_pt(
        _search(rf"\bCNEE\s*:\s*(.+?){_INDO_STOP}", text)
    )
    header["order_no"] = _search(rf"Order\s*No\.?\s*:\s*(.+?){_INDO_STOP}", text)
    header["bl_no"] = _search(rf"\bHAWB\s*:\s*(.+?){_INDO_STOP}", text)
    header["mbl_no"] = _search(rf"\bMAWB\s*:\s*(.+?){_INDO_STOP}", text)
    header["port_of_loading"] = _search(rf"\bORIGIN\s*:\s*(.+?){_INDO_STOP}", text)
    header["vessel"] = _search(rf"\bFLIGHT\s*:\s*(.+?){_INDO_STOP}", text)
    header["volume"] = _search(rf"GROSS\s*/\s*KG\s*:\s*(.+?){_INDO_STOP}", text)
    header["onboard_date"] = _indo_parse_flight_date(header.get("vessel"))

    # 這種空運帳單沒有到達日/目的港/貨櫃資訊，留空給 expand_to_rows() 補 N/A
    header["arrive_date"] = None
    header["port_of_discharge"] = None
    header["container_no"] = None

    items = _indo_extract_charge_items(text)
    if not items:
        items = [{"description": None, "amount": None}]

    return {"header": header, "items": items}


# ===========================================================================
# 供應商 4：MAERSK (PT Maersk Logistics Indonesia，海運帳單 + Faktur Pajak)
# ===========================================================================
#
# 版面特徵：
#   - 第1、2頁是 Maersk 自己的 Invoice，抬頭欄位 (POL/POD/ETD/ETA/Sold to...)
#     一樣是左右並排、只隔一個空白 (跟 INDOPROSTIME 同款排版問題)，用同一套
#     「遇到下一個已知欄位就停止」策略處理。
#   - 費用明細表格在 Invoice 頁面裡會因為欄寬不夠而「自動換行」，例如
#     "Origin Terminal Handling Charge" 的 "Charge" 兩個字被擠到下一行，
#     單獨用 Invoice 頁面的文字很難可靠重組回完整描述。所幸第3頁的
#     Faktur Pajak (稅務發票) 用不同版面列出同樣 4 筆費用，這裡描述文字
#     完整沒有被換行，所以費用明細改成從 Faktur Pajak 頁面擷取 (效果比
#     KUEHNE NAGEL 反過來：那家帳單是 Faktur Pajak 描述太籠統、要改抓
#     Sales Invoice 明細；這家則是 Invoice 頁面描述被換行拆散、要改抓
#     Faktur Pajak 明細)。
#   - 沒有海運提單，改用 "FCR" (Forwarder's Cargo Receipt) 編號頂替
#     B/L NO. 欄位；MBL NO.、Vessel、Container no. 這張帳單本來就沒有，
#     缺漏後續會自動補 N/A。

def detect_maersk(text: str) -> bool:
    return "MAERSK" in text.upper()


_MAERSK_LABEL_WORDS = [
    r"Place\s*of\s*Receipt", r"Place\s*of\s*Delivery", r"SHPR", r"CNEE",
    r"POL", r"POD", r"ETD", r"ETA", r"Sold\s*to", r"Total\s*Packages",
    r"Weight", r"Volume", r"Units",
]
_MAERSK_NEXT = "(?:" + "|".join(_MAERSK_LABEL_WORDS) + ")"
# 跟 INDOPROSTIME 一樣：並排欄位只隔一個空白，靠「下一個已知欄位關鍵字」
# 停止，不能靠連續空白判斷。
_MAERSK_STOP = rf"(?=\s+{_MAERSK_NEXT}\b|\n|$)"

_EN_MONTHS_ABBR = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


def _maersk_parse_date(text: Optional[str]) -> Optional[str]:
    """把 'ETD:/ETA:' 這種英文縮寫月份日期 '29-Aug-2026' 轉成 'DD.MM.YYYY'。

    這裡一定要轉成跟其他供應商一樣的 'DD.MM.YYYY' 數字格式，因為使用者的
    正確答案 Excel 這兩欄是 Excel 內建的日期格式(儲存格型態是日期,不是
    純文字)，比對時 normalize_value() 會把日期物件格式化成 'DD.MM.YYYY'，
    如果這裡保留原始的 '29-Aug-2026' 文字就永遠對不起來。
    """
    if not text:
        return text
    m = re.match(r"(\d{1,2})-([A-Za-z]{3})-(\d{4})", text.strip())
    if not m:
        return text
    day, month_abbr, year = m.groups()
    month = _EN_MONTHS_ABBR.get(month_abbr.lower())
    if not month:
        return text
    return f"{int(day):02d}.{month:02d}.{year}"


def _maersk_clean_amount_id(text: Optional[str]) -> Optional[int]:
    """印尼式數字 '623.699,00' (句點=千分位、逗號=小數) -> 623699。
    跟 KUEHNE NAGEL 的 _kn_clean_amount() 是同一種格式，各供應商各自維護
    一份，避免互相牽動。
    """
    if not text:
        return None
    text = text.strip().replace(".", "")
    text = text.split(",")[0]
    return int(text) if text.isdigit() else None


# Faktur Pajak (稅務發票) 頁面的費用明細行，例如：
#   "2 Origin Terminal Handling Charge 149.659,00"
# 格式：項次 + 費用名稱 + 印尼式金額。用這一頁的明細而不是 Invoice 頁面，
# 是因為 Invoice 頁面的表格欄寬不夠，長一點的費用名稱 (例如這筆) 會被自動
# 換行拆成兩行，很難可靠重組；Faktur Pajak 頁面版面不同，同一筆描述都在
# 同一行，擷取起來穩定得多。
_MAERSK_TAX_ITEM_PATTERN = re.compile(
    r"^\d+\s+(?P<desc>[A-Za-z][A-Za-z0-9 /\-\.]*?)\s+(?P<amount>[\d.]+,\d{2})\s*$",
    re.MULTILINE,
)


def _maersk_extract_items(text: str) -> List[Dict]:
    items = []
    for m in _MAERSK_TAX_ITEM_PATTERN.finditer(text):
        items.append({
            "description": m.group("desc").strip(),
            "amount": _maersk_clean_amount_id(m.group("amount")),
        })
    return items


def parse_maersk(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {}

    header["invoice_no"] = _search(r"\bINVOICE\s+(\d+)\b", text)
    header["invoice_date"] = _search(r"Invoice\s*Date\s*:\s*(\S+)", text)
    header["supplier"] = _search(r"Issued by:\s*\n(.+)", text)
    header["consignee"] = _search(rf"Sold\s*to\s*:\s*(.+?){_MAERSK_STOP}", text)
    header["order_no"] = _search(r"Commercial Invoice No\s*:\s*(\S+)", text)
    header["bl_no"] = _search(r"\bFCR\s+(\S+)", text)
    header["port_of_loading"] = _search(rf"\bPOL\s*:\s*(.+?){_MAERSK_STOP}", text)
    header["port_of_discharge"] = _search(rf"\bPOD\s*:\s*(.+?){_MAERSK_STOP}", text)
    header["onboard_date"] = _maersk_parse_date(
        _search(rf"\bETD\s*:\s*(.+?){_MAERSK_STOP}", text)
    )
    header["arrive_date"] = _maersk_parse_date(
        _search(rf"\bETA\s*:\s*(.+?){_MAERSK_STOP}", text)
    )
    header["volume"] = _search(r"Volume\s*:\s*([\d.]+)\s*m3", text)

    # 這張帳單沒有主提單/船名/貨櫃資訊，留空給 expand_to_rows() 補 N/A
    header["mbl_no"] = None
    header["vessel"] = None
    header["container_no"] = None

    items = _maersk_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None}]

    return {"header": header, "items": items}


# ===========================================================================
# 供應商 5：HYPER MEGA SHIPPING (⚠️ 特殊：一份 PDF 裡有好幾張各自獨立的發票)
# ===========================================================================
#
# 這家供應商跟前四家有一個根本性的不同：其他供應商不管 PDF 有幾頁，整份都
# 是「同一張發票」(欄位分散在不同頁面，但發票號碼只有一個)；HYPER MEGA
# 則是把好幾張完全獨立的發票 (各自的發票號碼、日期、費用明細都不一樣)
# 直接合併成一份 PDF、一頁就是一張發票。如果照舊把所有頁面文字合併成一個
# 字串再抓欄位，正規表示式只會抓到「第一個」出現的發票號碼/日期，後面幾張
# 發票的資料會直接遺失。
#
# 解法：註冊時傳入 multi_page=True，detect_hyper_mega()/parse_hyper_mega()
# 改成「逐頁」被呼叫 (extract_utils.py 的 _parse_multi_page_pdf() 會對每一
# 頁分別呼叫 detect()，承認的頁面才呼叫 parse() 並各自展開成列)，之後比對
# 正確率時也會用 invoice_no 把列分組、各自跟正確答案表格裡對應的發票比對
# (compare_multi_invoice_rows())，而不是只比對第一張發票。
#
# 版面特徵 (跟 INDOPROSTIME/MAERSK 一樣是單一空白分隔的並排欄位)：
#   - "Shipper" 那一列版面重疊、文字層是亂序的 (例如
#     'Shipper   : THE LOOK (MACAO...OFFSHOVReEss) eClO LTD : VANCOUVER 047S')，
#     但很穩定地固定用「這一列最後一個冒號後面的內容」代表 Vessel 欄位值，
#     不需要真的解出 Shipper 公司名稱本身。
#   - "Consignee" 那一列在公司名稱後面也有一段類似的亂碼，用「第一個左括號
#     之前」的文字取代 stop-lookahead 就好。
#   - 少數頁面 MBL No. 後面會多一個「-」尾巴 (排版問題)，擷取後要去掉。

def detect_hyper_mega(text: str) -> bool:
    return "HYPER MEGA" in text.upper()


_HYPER_LABEL_WORDS = [
    r"I\s*N\s*V\s*O\s*I\s*C\s*E", r"NPWP", r"Invoice\s*No\.?", r"Invoice\s*Date",
    r"Shipment\s*Type", r"Page", r"Order\s*No\.?", r"Port\s*of\s*Origin",
    r"Arrive\s*Date", r"Port\s*of\s*Discharge", r"B/L\s*No\.?", r"Reference",
    r"MBL\s*No\.?", r"CBM\s*/\s*Package", r"Shipper", r"Consignee",
    r"Description", r"Payment\s*Term", r"Sub\s*Total", r"DPP", r"V\s*A\s*T",
    r"GRAND\s*TOTAL",
]
_HYPER_NEXT = "(?:" + "|".join(_HYPER_LABEL_WORDS) + ")"
_HYPER_STOP = rf"(?=\s+{_HYPER_NEXT}\b|\n|$)"

_EN_MONTHS_FULL = {
    "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
    "july": 7, "august": 8, "september": 9, "october": 10, "november": 11,
    "december": 12,
}


def _hyper_parse_date(text: Optional[str]) -> Optional[str]:
    """'06 October 2025' -> '06.10.2025'，跟 MAERSK 一樣要轉成
    'DD.MM.YYYY'，因為正確答案 Excel 的 Arrive Date 欄位是 Excel 日期
    格式，比對時 normalize_value() 會把它格式化成這種點分隔格式。
    """
    if not text:
        return text
    m = re.match(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", text.strip())
    if not m:
        return text
    day, month_name, year = m.groups()
    month = _EN_MONTHS_FULL.get(month_name.lower())
    if not month:
        return text
    return f"{int(day):02d}.{month:02d}.{year}"


def _hyper_clean_amount_id(text: Optional[str]) -> Optional[int]:
    """印尼式數字 '16.990,00' (句點=千分位、逗號=小數) -> 16990。"""
    if not text:
        return None
    text = text.strip().replace(".", "")
    text = text.split(",")[0]
    return int(text) if text.isdigit() else None


def _hyper_clean_trailing_dash(value: Optional[str]) -> Optional[str]:
    """少數頁面 MBL No. 後面多一個「-」尾巴 (排版問題)，例如
    'OOLU2311247390-'，去掉結尾的連字號/多餘符號。
    """
    if not value:
        return value
    return value.strip().rstrip("-").strip()


# 費用名稱裡「- 9 okt」這種印尼文月份縮寫的日期註記，使用者的正確答案會把
# 它當成備註拿掉 (例如 'ADMINISTRATION FEE - 9 okt' -> 'ADMINISTRATION
# FEE')，但像 'STORAGE CHARGE - Masa 1 : 3 Days' 這種「- Masa ...」是用來
# 區分兩筆不同 STORAGE CHARGE 的關鍵字，正確答案反而完整保留，不能一併
# 清掉。所以只清「- 數字 + 印尼文月份縮寫」這種明確是日期的樣式。
_HYPER_DESC_DATE_SUFFIX = re.compile(
    r"\s*-\s*\d{1,2}\s*(?:jan|feb|mar|apr|mei|jun|jul|agu|sep|okt|nov|des)\w*\s*$",
    re.IGNORECASE,
)


def _hyper_clean_description(desc: Optional[str]) -> Optional[str]:
    if not desc:
        return desc
    return _HYPER_DESC_DATE_SUFFIX.sub("", desc).strip()


# 費用明細行，例如 "OCEAN FREIGHT CHARGES 1,000 IDR 16.990,00 16.990,00 1,1 %"
# 或 "MONITORING CHARGE 2,000 IDR 300.000,00 600.000,00 11 %"：
# 費用名稱 + 數量 + 幣別 IDR + 單價 + 金額 + VAT百分比。要的是「金額」那個
# 數字 (VAT% 前面那個)，不是單價 (前一個數字)，因為數量不是 1 時兩者不同
# (例如 MONITORING CHARGE 數量 2、單價 300,000、金額才是 600,000)。
_HYPER_ITEM_PATTERN = re.compile(
    r"^(?P<desc>[A-Z][A-Za-z0-9 \-:.]+?)\s+[\d,]+\s+IDR\s+[\d.,]+\s+"
    r"(?P<amount>[\d.,]+)\s+[\d,]+\s*%\s*$",
    re.MULTILINE,
)


def _hyper_extract_items(text: str) -> List[Dict]:
    items = []
    for m in _HYPER_ITEM_PATTERN.finditer(text):
        items.append({
            "description": _hyper_clean_description(m.group("desc").strip()),
            "amount": _hyper_clean_amount_id(m.group("amount")),
        })
    return items


def parse_hyper_mega(text: str) -> Dict:
    """注意：這是 multi_page 供應商，text 是『一頁』的文字 (不是整份 PDF
    合併後的文字)，呼叫端 (_parse_multi_page_pdf()) 只會對通過
    detect_hyper_mega() 的頁面呼叫這個函式。
    """
    header: Dict[str, Optional[str]] = {}

    header["invoice_no"] = _search(rf"Invoice\s*No\.?\s*:\s*(.+?){_HYPER_STOP}", text)
    header["invoice_date"] = _hyper_parse_date(
        _search(rf"Invoice\s*Date\s*:\s*(.+?){_HYPER_STOP}", text)
    )
    header["supplier"] = _search(r"(?i)^\s*(PT\.\s*Hyper\s*Mega\s*Shipping)", text)
    header["consignee"] = _search(r"Consignee\s*:\s*([^(\n]+)", text)
    header["order_no"] = _search(rf"Order\s*No\.?\s*:\s*(.+?){_HYPER_STOP}", text)
    header["port_of_loading"] = _search(rf"Port\s*of\s*Origin\s*:\s*(.+?){_HYPER_STOP}", text)
    header["arrive_date"] = _hyper_parse_date(
        _search(rf"Arrive\s*Date\s*:\s*(.+?){_HYPER_STOP}", text)
    )
    header["port_of_discharge"] = _search(rf"Port\s*of\s*Discharge\s*:\s*(.+?){_HYPER_STOP}", text)
    header["bl_no"] = _search(rf"B/L\s*No\.?\s*:\s*(.+?){_HYPER_STOP}", text)
    header["mbl_no"] = _hyper_clean_trailing_dash(
        _search(rf"MBL\s*No\.?\s*:\s*(.+?){_HYPER_STOP}", text)
    )
    # Vessel 欄位版面跟 Shipper 重疊、文字層亂序，但「這一列最後一個冒號
    # 後面的內容」穩定地就是 Vessel 值 (參見本區塊開頭的說明)。
    header["vessel"] = _search(r"Shipper\s*:.*:\s*(.+?)\s*$", text)
    m = re.search(r"CBM\s*/\s*Package\s*:\s*(?:LCL|FCL)\s*/\s*([\d,]+)\s*/", text)
    header["volume"] = m.group(1).replace(",", ".") if m else None

    # 這種帳單沒有獨立的「開船日」欄位、也沒有貨櫃資訊，缺漏後續會自動
    # 補 N/A。
    header["onboard_date"] = None
    header["container_no"] = None

    items = _hyper_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None}]

    return {"header": header, "items": items}


# ===========================================================================
# 註冊所有供應商
# ===========================================================================
# detect_kuehne_nagel() 的關鍵字 "KUEHNE"+"NAGEL" 很具體不會誤判，
# detect_dwiharta() 同理只認 "DWIHARTA" 字樣，detect_indoprostime() 只認
# "PROSTIME" 字樣，detect_maersk() 只認 "MAERSK" 字樣，detect_hyper_mega()
# 只認 "HYPER MEGA" 字樣，五者互不衝突，順序不影響結果。未來新增供應商時，
# 關鍵字盡量挑「該公司獨有、不會出現在其他供應商帳單」的字串 (公司全名、
# 統編、帳單編號前綴等)；如果新供應商也是「一份 PDF 有好幾張獨立發票」，
# 記得跟 HYPER MEGA 一樣在 register_supplier() 傳入 multi_page=True。

register_supplier("DWIHARTA", "PT. DWIHARTA LOGISTINDO", detect_dwiharta, parse_dwiharta)
register_supplier("KUEHNE_NAGEL", "KUEHNE NAGEL INDONESIA", detect_kuehne_nagel, parse_kuehne_nagel)
register_supplier("INDOPROSTIME", "PT. INDO PROSTIME EXPRESS", detect_indoprostime, parse_indoprostime)
register_supplier("MAERSK", "PT MAERSK LOGISTICS INDONESIA", detect_maersk, parse_maersk)
register_supplier(
    "HYPER_MEGA", "PT. HYPER MEGA SHIPPING", detect_hyper_mega, parse_hyper_mega,
    multi_page=True,
)
