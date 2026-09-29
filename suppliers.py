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

# ===========================================================================
# 供應商 6a：YJE (ShenZhen) International Logistics (⚠️ multi_page)
# ===========================================================================
#
# 這家帳單版面很單純：每一頁就是一張完整獨立的發票 (跟 HYPER MEGA 一樣，
# 一份 PDF 裡可能塞好幾張不同發票號碼/日期/費用明細的獨立帳單，一頁一張)，
# 所以一樣用 multi_page=True 逐頁辨識/擷取。
#
# 版面特徵：
#   - 標題公司名稱是 "YJE (ShenZhen) International Logistics Co., Ltd."，
#     但正確答案 Excel 的 SUPPLIER 欄位填的是帳單下方 "AGENT:" 那一行
#     ("PT TATA HARMONI SARANATAMA")，不是抬頭公司名稱。
#   - PDF 裡的冒號是全形 "："，不是半形 ":"，抓值時兩種都要接受。
#   - "TO：" 收件人欄位偶爾會多一個句點 ("PT. POU YUEN INDONESIA")，
#     偶爾沒有 ("PT POU YUEN INDONESIA")，正確答案統一是不帶句點的寫法，
#     擷取後要把 "PT." 開頭正規化成 "PT "。
#   - 金額用歐式寫法 (逗號當小數點)，例如 "US$1,72" 代表 1.72、
#     "US$115,00" 代表 115.00，轉換時把逗號換成小數點再轉成數字。
#   - 沒有 Order No./Arrive Date/on board date/MBL NO./Port of
#     Loading/Port of discharge/Volume/Vessel/Container no.，全部留空
#     讓 expand_to_rows() 補 N/A。

def detect_yje(text: str) -> bool:
    upper = text.upper()
    return "YJE" in upper and "INTERNATIONAL LOGISTICS" in upper


def _yje_clean_consignee(value: Optional[str]) -> Optional[str]:
    """'PT. POU YUEN INDONESIA' -> 'PT POU YUEN INDONESIA'：正確答案統一
    不帶句點的寫法，PDF 上偶爾會多一個句點，擷取後去掉。
    """
    if not value:
        return value
    return re.sub(r"^(PT)\.\s*", r"\1 ", value.strip())


def _yje_clean_amount(text: Optional[str]) -> Optional[float]:
    """歐式金額 'US$1,72' (逗號=小數點) -> 1.72。"""
    if not text:
        return None
    text = text.strip().replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


_YJE_ITEM_PATTERN = re.compile(
    r"^(?P<desc>[A-Za-z][A-Za-z0-9 +\-]*?)\s+\d+\s+SET\s+US\$(?P<amount>[\d,\.]+)\s*$",
    re.MULTILINE,
)


def _yje_extract_items(text: str) -> List[Dict]:
    items = []
    for m in _YJE_ITEM_PATTERN.finditer(text):
        items.append({
            "description": m.group("desc").strip(),
            "amount": _yje_clean_amount(m.group("amount")),
        })
    return items


def parse_yje(text: str) -> Dict:
    """注意：這是 multi_page 供應商，text 是『一頁』的文字。"""
    header: Dict[str, Optional[str]] = {}

    header["invoice_no"] = _search(r"Inv\.?\s*No\s*[:：]\s*(\S+)", text)
    header["invoice_date"] = _search(r"Inv\s*Date\s*[:：]\s*(\S+)", text)
    header["supplier"] = _search(r"AGENT\s*[:：]\s*(.+)$", text)
    header["consignee"] = _yje_clean_consignee(_search(r"TO\s*[:：]\s*(.+)$", text))
    header["bl_no"] = _search(r"HAWB\s*[:：]\s*(\S+)", text)

    # 這種帳單沒有 Order No./Arrive Date/開船日/MBL NO./港口/材積/船名/
    # 貨櫃資訊，留空給 expand_to_rows() 補 N/A
    for code in ("order_no", "arrive_date", "onboard_date", "mbl_no",
                 "port_of_loading", "port_of_discharge", "volume",
                 "vessel", "container_no"):
        header[code] = None

    items = _yje_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None}]

    return {"header": header, "items": items}


# ===========================================================================
# 供應商 6b：PT TATA HARMONI SARANATAMA (⚠️ multi_page，YJE 帳單的第二種格式)
# ===========================================================================
#
# 這是同一個供應商關係下的「第二種帳單格式」：YJE 的帳單上會寫
# "AGENT: PT TATA HARMONI SARANATAMA"，而這裡是 TATA 自己開立的印尼文
# 稅務/報關費用發票，版面、欄位、金額格式都跟 YJE 那份完全不同，所以另外
# 寫一組 detect_tata()/parse_tata()，用不同的 key 註冊 (辨識關鍵字互不
# 衝突：YJE 帳單裡沒有 "TAMAN DUTAMAS" 字樣，TATA 帳單裡也沒有
# "INTERNATIONAL LOGISTICS" 字樣)。
#
# 版面特徵 (跟 INDOPROSTIME/MAERSK 一樣是單一空白分隔的並排欄位)：
#   - 一份 PDF 一樣可能塞好幾張獨立發票、一頁一張，所以也是 multi_page=True。
#   - "NO. DESCRIPTION AMOUNT (Rp)" 表格裡每筆費用前面有項次編號，
#     "HANDLING FEE" 那一筆後面偶爾會多一個貨運追蹤號碼尾巴 (例如
#     "HANDLING FEE YJE123965498")，正確答案只要 "HANDLING FEE" 本身，
#     要把追蹤號碼去掉；判斷方式是看緊接在費用名稱後面的那個字有沒有
#     包含數字 (追蹤號碼一定有數字，費用名稱單字不會)。
#   - 正確答案除了表格裡列出的費用明細，還多一筆 "DPP"（未稅金額，
#     表格下方單獨列出的小計），要額外抓出來當成一筆 Description/Amount。
#   - 金額是英式千分位、無小數 ('159,840' -> 159840)。
#   - Arrive Date/on board date 兩欄的值相同，都是來自 "Date :" 那個
#     出貨日期欄位 (跟到達日/開船日不是嚴格對應，但正確答案兩欄填的是
#     同一個值)。
#   - SUPPLIER 固定是發票最下方 "Nama :" 那一行 ("PT. TATA HARMONI
#     SARANATAMA")。

def detect_tata(text: str) -> bool:
    return "TAMAN DUTAMAS" in text.upper()


_TATA_LABEL_WORDS = [
    r"MAWB", r"HAWB", r"Telp", r"Attn", r"Flight", r"Date", r"QTY",
    r"Shipper", r"Address", r"Consignee", r"ORIGIN",
]
_TATA_NEXT = "(?:" + "|".join(_TATA_LABEL_WORDS) + ")"
_TATA_STOP = rf"(?=\s+{_TATA_NEXT}\b|\n|$)"


def _tata_clean_amount(text: Optional[str]) -> Optional[int]:
    """英式數字 '4,585,728' (逗號千分位，無小數) -> 4585728。"""
    if not text:
        return None
    text = text.strip().replace(",", "")
    return int(text) if text.isdigit() else None


def _tata_parse_date(text: Optional[str]) -> Optional[str]:
    """'2026-08-24' -> '24.08.2026'，跟其他供應商一樣統一成 'DD.MM.YYYY'
    (正確答案 Excel 這幾欄是 Excel 日期格式)。
    """
    if not text:
        return text
    m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", text.strip())
    if not m:
        return text
    year, month, day = m.groups()
    return f"{int(day):02d}.{int(month):02d}.{year}"


# 費用明細行，例如 "1 GUDANG FEE Rp 159,840" 或
# "4 HANDLING FEE YJE123965498 Rp 4,585,728"：項次 + 費用名稱 + (偶爾多一個
# 含數字的追蹤號碼尾巴) + Rp + 金額。用「後面緊接的字是否含數字」判斷那個字
# 是追蹤號碼還是費用名稱的一部分，避免把 "LOCAL EXPRESS CHARGE" 這種全大寫
# 費用名稱誤判成「名稱+尾碼」。
_TATA_ITEM_PATTERN = re.compile(
    r"^\d+\s+(?P<desc>[A-Z][A-Z ]*?)(?:\s+(?=[A-Z0-9]*\d)[A-Z0-9]+)?\s+Rp\s+(?P<amount>[\d,]+)\s*$",
    re.MULTILINE,
)
_TATA_DPP_PATTERN = re.compile(r"^DPP\s+Rp\s+(?P<amount>[\d,]+)\s*$", re.MULTILINE)


def _tata_extract_items(text: str) -> List[Dict]:
    items = []
    for m in _TATA_ITEM_PATTERN.finditer(text):
        items.append({
            "description": m.group("desc").strip(),
            "amount": _tata_clean_amount(m.group("amount")),
        })
    dpp_m = _TATA_DPP_PATTERN.search(text)
    if dpp_m:
        items.append({"description": "DPP", "amount": _tata_clean_amount(dpp_m.group("amount"))})
    return items


def parse_tata(text: str) -> Dict:
    """注意：這是 multi_page 供應商，text 是『一頁』的文字。"""
    header: Dict[str, Optional[str]] = {}

    header["invoice_no"] = _search(r"(?m)^NO\s*:\s*(\S+)", text)
    header["invoice_date"] = _tata_parse_date(
        _search(r"INV\s*Date\s*:\s*(\d{4}-\d{1,2}-\d{1,2})", text)
    )
    header["supplier"] = _search(r"Nama\s*:\s*(PT\..+)$", text)
    header["consignee"] = _search(rf"\bTO\s*:\s*(.+?){_TATA_STOP}", text)
    header["mbl_no"] = _search(rf"\bMAWB\s*:\s*(.+?){_TATA_STOP}", text)
    header["bl_no"] = _search(rf"\bHAWB\s*:\s*(.+?){_TATA_STOP}", text)
    header["vessel"] = _search(rf"\bFlight\s*:\s*(.+?){_TATA_STOP}", text)
    header["volume"] = _search(rf"\bQTY\s*:\s*(.+?){_TATA_STOP}", text)
    header["port_of_loading"] = _search(rf"\bORIGIN\s*:\s*(.+?){_TATA_STOP}", text)

    # 注意：不能用 r"\bDate\s*:\s*..." 這種寬鬆寫法，因為同一頁最上方
    # "INV Date :2026-08-24 15:55:13" 也符合這個樣式 (字面上 "Date :" 前面
    # 剛好也有一個單字邊界)，會被誤抓成「出貨日期」；這裡用負向後顧排除
    # 緊接在 "INV " 後面的那個 "Date :"，只抓 Consignee 那一行真正的出貨
    # 日期 "Date :2026-08-16"。
    ship_date = _tata_parse_date(
        _search(r"(?<!INV )Date\s*:\s*(\d{4}-\d{1,2}-\d{1,2})", text)
    )
    header["arrive_date"] = ship_date
    header["onboard_date"] = ship_date

    # 這種帳單沒有 Order No./目的港/貨櫃資訊，留空給 expand_to_rows() 補 N/A
    header["order_no"] = None
    header["port_of_discharge"] = None
    header["container_no"] = None

    items = _tata_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None}]

    return {"header": header, "items": items}


# ---------------------------------------------------------------------------
# YJE 與 TATA 合併成同一個「供應商群組」註冊 (重要修正)
# ---------------------------------------------------------------------------
#
# 問題根因：使用者反映「有時候 TATA 的帳單會跟 YJE 合成同一份 PDF 一起請
# 款」，但之前是把 YJE、TATA 各自獨立註冊成兩個供應商 (key="YJE"、
# key="TATA")。多供應商辨識/擷取流程是「先用整份 PDF 的文字判斷屬於哪一個
# 供應商 key (detect_supplier())，再固定用那一個 key 對應的
# detect()/parse() 逐頁處理 (_parse_multi_page_pdf())」。當同一份 PDF 裡
# 同時混著 YJE 格式頁面跟 TATA 格式頁面時，整份 PDF 的文字裡「YJE」跟
# 「TATA」的關鍵字都會出現，detect_supplier() 依序詢問每個已註冊供應商，
# 一旦先問到的 "YJE" 承認了 (回傳 True)，就會把整份 PDF 都當成 "YJE" 這個
# key 處理；之後逐頁擷取時，只會用 YJE 自己的 detect_yje()/parse_yje() 去
# 跑每一頁，TATA 格式的頁面因為通不過 detect_yje() 而被直接跳過、不會被
# 擷取，也就不會出現在「逐欄比對明細」裡——這就是使用者回報「TATA 沒有顯示
# 出來」的真正原因，不是 TATA 的擷取規則本身壞掉。
#
# 修正方式：不要分成兩個供應商 key，改成註冊「同一個」multi_page 供應商，
# 但 detect()/parse() 內部『每一頁』都同時检查兩種格式——是 YJE 格式就用
# parse_yje()，是 TATA 格式就用 parse_tata()，兩種格式都不是才判定這頁不
# 屬於這個供應商。這樣不管一份 PDF 裡是純 YJE、純 TATA，還是兩種混在一起，
# 每一頁都會被個別正確辨識並擷取，不會因為「整份 PDF 先綁定同一個 key」而
# 漏掉另一種格式的頁面。

def detect_yje_or_tata(text: str) -> bool:
    return detect_yje(text) or detect_tata(text)


def parse_yje_or_tata(text: str) -> Dict:
    """注意：這是 multi_page 供應商，text 是『一頁』的文字。每一頁各自
    判斷是 YJE 格式還是 TATA 格式，呼叫對應的 parse 函式；因為只有先通過
    detect_yje_or_tata() 的頁面才會呼叫到這裡，所以這兩個條件已經涵蓋了
    所有會進來的頁面。
    """
    if detect_yje(text):
        return parse_yje(text)
    return parse_tata(text)


# ===========================================================================
# 供應商 7：PT. TRANS DAYA PRIMA (⚠️ multi_page：一份 PDF 裡有好幾張獨立發票)
# ===========================================================================
#
# 版面特徵：跟 YJE/TATA/HYPER MEGA 一樣，一份 PDF 裡可能塞好幾張各自獨立的
# "Sales Invoice"，一頁一張 (各自不同的發票號碼/日期/費用明細)，所以一樣要
# multi_page=True 逐頁辨識/擷取。
#
#   - 抬頭三欄位 (Invoice Date、Invoice Number、收件人 Consignee) 版面固定
#     在同一個位置：
#       "Invoice Date Invoice Number
#        26 Jan 2026 0183-TDP-GSI2-I-2026
#        PT. GLOSTAR INDONESIA
#        PT. GLOSTAR INDONESIA - JL. Raya Sukabumi ..."
#     日期/發票號碼在同一行，收件人在緊接著的下一行 (下下一行才是重複一次
#     的收件人+地址，不用管)。
#   - ⚠️ 費用明細『不一定』只有一筆：有些帳單一張只有一筆費用，有些一張有
#     兩、三筆。而且**正確答案的 Order No. 欄位是「每一筆費用明細各自的
#     DO No.」，不是整張帳單共用同一個值**——同一張帳單裡不同費用列的
#     Order No. 可能不一樣。所以這裡改成逐行掃描，每一筆費用明細各自帶著
#     自己的 order_no 一起回傳 (參見 extract_utils.expand_to_rows()：item
#     dict 裡如果帶了 HEADER_FIELD_CODES 裡的欄位，會覆蓋掉整張帳單共用的
#     抬頭值，只套用在這一列)。
#   - 費用明細行有兩種版面：
#     (a) 完整版 (有 DO No./DO Date/部門代碼)，例如：
#         "GSI-2 FCL Trucking 40" GTO20260120-077 20 Jan 2026 B 9308 JIN 1 4.920.000 4.920.000"
#         格式：費用名稱 + DO No. + DO Date + 部門代碼(單一大寫字母+數字+
#         字母) + 數量 + 單價 + 總價。DO No. 若因為欄寬不夠被自動換行拆成
#         兩截 (例如 "00055/TDP/SJ-" 留在這一行、"01/2026" 被擠到下一行單
#         獨一行)，這裡會偵測「這一截結尾是 '-' 或 '/' 這種明顯還沒結束的
#         符號，且下一行是一個獨立的短字串」，把兩截拼回完整的 DO No.。
#     (b) 簡化版 (沒有 DO No./DO Date/部門代碼，例如另外加收的手續費)：
#         "Handling GSI-2 Operational 2 300.000 600.000"
#         格式：費用名稱 + 數量 + 單價 + 總價，這種費用沒有對應的 DO No.，
#         Order No. 留 None 讓 expand_to_rows() 補 N/A。
#     逐行掃描時两種版面互斥判斷 (完整版比對不到才試簡化版)，避免像
#     "Biaya LOLO 1.431.900" 這種只有一個數字的稅金/雜費小計行被誤判成
#     費用明細 (簡化版需要「數量+單價+總價」三個數字都存在才會比對到)。
#   - 費用名稱 (Description) 偶爾帶有貨櫃呎吋的引號 (例如
#     'LCL Lokal Serang - GSI-2 40"' 代表 40 呎櫃)，正確答案裡這個引號被
#     拿掉了 (只留 'LCL Lokal Serang - GSI-2 40')，擷取後要清掉引號字元。
#   - 金額是印尼式千分位 (句點分隔、無小數)，例如 '2.975.000' -> 2975000。
#   - 目的港：帳單右側偶爾會有 "P.O.D : USA" 這一行，正確答案的 Port of
#     discharge 欄位就是抓這個值；沒有這一行的帳單維持 N/A。
#   - 沒有到達日/開船日/提單/主提單/啟運港/材積/船名/貨櫃資訊 (帳單上雖然
#     另外印了 Vessel/Flight/Cont No 等欄位，但正確答案不會把這些對應到
#     我們的欄位，一律留 N/A，避免抓多)，缺漏後續會自動補 N/A。

def detect_trans(text: str) -> bool:
    return "TRANS DAYA PRIMA" in text.upper()


# 完整版費用明細行 (費用名稱 + DO No. + DO Date + 部門代碼 + 數量 + 單價 +
# 總價)，見上方說明。用 fullmatch 逐行比對 (不是在整段文字裡 search)，
# 避免跟簡化版費用行搶著比對同一行文字。
_TRANS_FULL_ITEM_PATTERN = re.compile(
    r"^(?P<desc>.+?)\s+(?P<do_no>\S+)\s+(?P<do_date>\d{1,2}\s+[A-Za-z]{3}\s+\d{4})\s+"
    r"(?P<dept>[A-Z]\s+\d+\s+[A-Z]+)\s+(?P<qty>\d+)\s+(?P<unit_price>[\d.,]+)\s+"
    r"(?P<total>[\d.,]+)$"
)
# 簡化版費用明細行 (沒有 DO No./DO Date/部門代碼)：費用名稱 + 數量 + 單價 +
# 總價，三個數字缺一不可，避免誤吃到只有一個數字的稅金/雜費小計行。
# 注意：這種行偶爾在「費用名稱」跟「數量」之間會多夾一個部門代碼的雜字
# (例如 'Handling GSI-2 Operational 2 300.000 600.000' 裡的 'Operational'，
# 版面上其實落在 Department 欄位的位置，只是這一行沒有印出完整代碼)，這裡
# 用一個可有可無的 dept 群組吃掉它，不要讓它被誤併進 desc 裡。
_TRANS_SIMPLE_ITEM_PATTERN = re.compile(
    r"^(?=[A-Za-z])(?P<desc>.+?)\s+(?:(?P<dept>[A-Za-z]+)\s+)?"
    r"(?P<qty>\d+)\s+(?P<unit_price>[\d.,]+)\s+(?P<total>[\d.,]+)$"
)


def _trans_parse_date(text: Optional[str]) -> Optional[str]:
    """'26 Jan 2026' -> '26.01.2026'，跟其他供應商一樣統一成 'DD.MM.YYYY'
    (正確答案 Excel 這欄是 Excel 日期格式)。沿用 MAERSK 已定義的
    _EN_MONTHS_ABBR 英文月份縮寫對照表，不用重複定義一份。
    """
    if not text:
        return text
    m = re.match(r"(\d{1,2})\s+([A-Za-z]{3})\s+(\d{4})", text.strip())
    if not m:
        return text
    day, month_abbr, year = m.groups()
    month = _EN_MONTHS_ABBR.get(month_abbr.lower())
    if not month:
        return text
    return f"{int(day):02d}.{month:02d}.{year}"


def _trans_clean_amount(text: Optional[str]) -> Optional[int]:
    """印尼式數字 '2.975.000' (句點千分位，無小數) -> 2975000。"""
    if not text:
        return None
    text = text.strip().replace(".", "").replace(",", "")
    return int(text) if text.isdigit() else None


def _trans_clean_desc(text: Optional[str]) -> Optional[str]:
    """去掉貨櫃呎吋標示裡的引號字元 (例如 '... 40"' -> '... 40')，正確答案
    是不帶引號的寫法。
    """
    if not text:
        return text
    return text.replace('"', "").strip()


def _trans_extract_items(text: str) -> List[Dict]:
    """逐行掃描費用明細表格區塊：完整版比對不到才試簡化版；完整版的
    DO No. 若被換行拆成兩截，往下併一行。回傳的每一筆 item 都各自帶著
    自己的 order_no (完整版有、簡化版沒有→None)，讓 expand_to_rows()
    對這一列做逐筆覆蓋，而不是整張帳單套用同一個 Order No.。
    """
    lines = text.split("\n")
    items: List[Dict] = []
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        m = _TRANS_FULL_ITEM_PATTERN.match(line.strip())
        if m:
            do_no = m.group("do_no")
            if i + 1 < len(lines):
                nxt = lines[i + 1].strip()
                looks_like_suffix = (
                    do_no.endswith(("-", "/"))
                    and nxt
                    and re.fullmatch(r"\S+", nxt)
                    and not _TRANS_FULL_ITEM_PATTERN.match(nxt)
                    and not _TRANS_SIMPLE_ITEM_PATTERN.match(nxt)
                )
                if looks_like_suffix:
                    do_no = do_no + nxt
                    i += 1  # 續行已經併進上一筆，跳過不要再單獨處理一次
            # 注意：Amount 欄位對應的是「Unit Price」(單價) 那一欄，不是
            # 「Total Price」(單價 x 數量)——之前送測的樣本數量都剛好是 1，
            # 單價跟總價數字一樣看不出差別；這次數量有一筆是 2 (見下面簡化
            # 版費用行)，正確答案填的是單價 (300.000) 不是總價 (600.000)，
            # 才發現真正要抓的是 Unit Price 這一欄。
            items.append({
                "description": _trans_clean_desc(m.group("desc")),
                "amount": _trans_clean_amount(m.group("unit_price")),
                "order_no": do_no.strip(),
            })
            i += 1
            continue

        m2 = _TRANS_SIMPLE_ITEM_PATTERN.match(line.strip())
        if m2:
            items.append({
                "description": _trans_clean_desc(m2.group("desc")),
                "amount": _trans_clean_amount(m2.group("unit_price")),
                "order_no": None,
            })
        i += 1

    return items


def parse_trans(text: str) -> Dict:
    """注意：這是 multi_page 供應商，text 是『一頁』的文字。"""
    header: Dict[str, Optional[str]] = {}

    m = re.search(
        r"Invoice\s*Number\s*\n\s*(\d{1,2}\s+[A-Za-z]{3}\s+\d{4})\s+(\S+)\s*\n(.+?)\n",
        text,
    )
    if m:
        header["invoice_date"] = _trans_parse_date(m.group(1))
        header["invoice_no"] = m.group(2).strip()
        header["consignee"] = m.group(3).strip()

    header["supplier"] = _search(r"^(PT\.\s*TRANS\s*DAYA\s*PRIMA)", text)
    header["port_of_discharge"] = _search(r"P\.O\.D\s*:\s*(.+?)\s*$", text)

    # 這種帳單沒有到達日/開船日/提單/主提單/啟運港/材積/船名/貨櫃資訊，留空
    # 給 expand_to_rows() 補 N/A。Order No. 現在是逐筆費用各自帶自己的值
    # (見 _trans_extract_items())，這裡的抬頭預設值留 None，沒有費用明細
    # 帶值時才會用到 (整張帳單抓不到任何費用行的防禦性情況)。
    for code in ("arrive_date", "onboard_date", "bl_no", "mbl_no",
                 "port_of_loading", "volume", "vessel", "container_no",
                 "order_no"):
        header[code] = None

    items = _trans_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None}]

    return {"header": header, "items": items}


# ===========================================================================
# 供應商 8：PT. FEMARIA BUANA CARGO (單頁、單張發票，費用明細表格)
# ===========================================================================
#
# 版面特徵：抬頭欄位分散在多行、常常跟右邊另一個欄位擠在同一行 (例如
# "To : Pouchen Indonesia,PT Invoice IDR No. : 53299/PCI/INV/2026")，用
# 「遇到下一個已知欄位關鍵字就停止擷取」的方式抓值。
#
# 費用明細行格式："編號.(代碼)-費用名稱 中間的參考碼 金額"，例如：
#   "1.(602)-LIFTOFF HJS/K20260136472IHM 699.300,00"
#   "5.HANDLING CHARGES PPh 23 300.000,00"     (中間參考碼可能是兩個字 "PPh 23")
#   "4.BIAYA SEGEL CONTAINER - 25.000,00"       (中間參考碼可能只是一個 "-")
# 金額都是最後一欄 (Amount IDR)，因為 Amout USD / Kurs 兩欄是空的，
# pdfplumber 抽取出來的文字只剩一個數字。
#
# 已知問題：原始 PDF 裡有些費用名稱底層文字本身就沒有空格 (不是
# pdfplumber 抽取參數的問題，用任何 x_tolerance 測試結果都一樣)，例如
# "THC-TERMINALHANDLINGCHARGE" 應該是 "THC-TERMINAL HANDLING CHARGE"。
# 這種已知的沾黏文字用 _FEMARIA_DESC_FIXUPS 對照表修正，之後如果遇到其他
# 沾黏的費用名稱，比照這個表補上新的一筆即可，不用改其他程式碼。
_FEMARIA_DESC_FIXUPS = {
    "THC-TERMINALHANDLINGCHARGE": "THC-TERMINAL HANDLING CHARGE",
}

_FEMARIA_ITEM_PATTERN = re.compile(
    r"^\d+\.\s*(?P<desc>.+?)\s+(?P<ref>-|PPh\s*23|\S+)\s+(?P<amount>[\d.,]+)$"
)


def detect_femaria(text: str) -> bool:
    return "FEMARIA BUANA CARGO" in text.upper()


def _femaria_parse_date(text: Optional[str]) -> Optional[str]:
    """'2026-01-09' (YYYY-MM-DD) -> '09.01.2026'，統一成跟其他供應商一樣
    的 'DD.MM.YYYY'。
    """
    if not text:
        return text
    m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", text.strip())
    if not m:
        return text
    year, month, day = m.groups()
    return f"{int(day):02d}.{int(month):02d}.{year}"


def _femaria_clean_amount(text: Optional[str]) -> Optional[int]:
    """印尼式數字 '2.276.380,00' (句點千分位、逗號小數) -> 2276380。"""
    if not text:
        return None
    text = text.strip().split(",")[0].replace(".", "")
    return int(text) if text.lstrip("-").isdigit() else None


def _femaria_clean_desc(text: Optional[str]) -> Optional[str]:
    """去掉開頭 '(代碼)-' 的費用代碼前綴，並修正已知的沾黏文字。"""
    if not text:
        return text
    text = re.sub(r"^\(\d+\)-", "", text.strip()).strip()
    return _FEMARIA_DESC_FIXUPS.get(text, text)


def _femaria_add_space_after_comma(value: Optional[str]) -> Optional[str]:
    """'Pouchen Indonesia,PT' -> 'Pouchen Indonesia, PT'。"""
    if not value:
        return value
    return re.sub(r",(\S)", r", \1", value.strip())


def _femaria_extract_items(text: str) -> List[Dict]:
    items: List[Dict] = []
    for line in text.split("\n"):
        m = _FEMARIA_ITEM_PATTERN.match(line.strip())
        if not m:
            continue
        items.append({
            "description": _femaria_clean_desc(m.group("desc")),
            "amount": _femaria_clean_amount(m.group("amount")),
        })
    return items


def parse_femaria(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {}

    header["supplier"] = _search(r"^(PT\.\s*FEMARIA\s*BUANA\s*CARGO)", text)
    header["consignee"] = _femaria_add_space_after_comma(
        _search(r"To\s*:\s*(.+?)\s+Invoice\s+IDR\s+No", text)
    )
    header["invoice_no"] = _search(r"Invoice\s+IDR\s+No\.\s*:\s*(\S+)", text)
    header["invoice_date"] = _femaria_parse_date(
        _search(r"Jakarta\s*,\s*(\d{4}-\d{1,2}-\d{1,2})", text)
    )
    header["volume"] = _search(r"Detail pty.*?:\s*\((.+?)\)", text)
    header["vessel"] = _search(r"Vessel\s*/\s*Voyage\s*:\s*(.+?)\s+ATA", text)
    header["bl_no"] = _search(r"BL/AWB\s+No\.\s*:\s*(\S+)", text)
    header["mbl_no"] = header["bl_no"]

    # 這種帳單沒有到達日/開船日/啟運港/目的港/貨櫃號碼/Order No.，留空給
    # expand_to_rows() 補 N/A (SPJK/SPJM Date、ATA、Uitslag Date 印出來的都是
    # 假的佔位日期 '0000-00-00 00:00:00'，不是真的日期，所以不擷取)。
    for code in ("order_no", "arrive_date", "onboard_date",
                 "port_of_loading", "port_of_discharge", "container_no"):
        header[code] = None

    items = _femaria_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None}]

    return {"header": header, "items": items}


# ===========================================================================
# 供應商 9：PT. EXPRESS MAXIMUM (單頁、單張發票，費用明細表格)
# ===========================================================================
#
# 版面特徵：費用明細每一列在 PDF 裡實際橫跨兩行 (SHIPPER 欄位名稱、ATTN
# 欄位人名太長會換行接到下一行，例如 "BIG" 換行接 "CORPORATION")，但這兩個
# 欄位都不是正確答案要比對的欄位，所以直接用 fullmatch 只認第一行 (有完整
# 9 個欄位、以兩個逗號分隔數字結尾的那一行)，換行的接續行 (只有零星文字、
# 對不上這個樣式) 直接跳過即可，不需要特別處理續行合併。
#
# 費用明細行的 9 個欄位依序是：
#   DATE  SHIPPER  AWB_NO  DESCRIPTION  ATTN  C/T  W/T  CHARGE(USD)
#   CHARGE(IDR)  TAX(IDR)
# 正確答案的 Amount 欄位對應 CHARGE(IDR) (倒數第二個數字)，不是 CHARGE(USD)
# 也不是最後的 TAX(IDR)。到達日/開船日/提單號碼/主提單號碼都是這一列自己
# 的 DATE/AWB_NO，不是整張帳單共用一個值，所以每筆費用都各自帶自己的
# arrive_date/onboard_date/bl_no/mbl_no，讓 expand_to_rows() 做逐筆覆蓋
# (即使這次的樣本兩筆費用剛好同一個 AWB/日期，未來遇到同一張發票裡有多個
# 不同貨運批次時也不會出錯)。
_EXPRESS_ITEM_PATTERN = re.compile(
    r"^\d{1,2}/[A-Za-z]{3}/\d{2}\s+\S+\s+(?P<awb>\S+)\s+(?P<desc>[A-Z]+)\s+\S+\s+"
    r"[\d.]+\s+[\d.]+\s+[\d.,]+\s+(?P<amount>[\d,]+)\s+[\d,]+$"
)
_EXPRESS_DATE_PATTERN = re.compile(r"^(\d{1,2})/([A-Za-z]{3})/(\d{2})")


def detect_express(text: str) -> bool:
    return "EXPRESS MAXIMUM" in text.upper()


def _express_add_space_after_pt(value: Optional[str]) -> Optional[str]:
    """'PT.EXPRESS MAXIMUM' -> 'PT. EXPRESS MAXIMUM'。"""
    if not value:
        return value
    return re.sub(r"^(PT\.)(\S)", r"\1 \2", value.strip())


def _express_parse_long_date(text: Optional[str]) -> Optional[str]:
    """'1 September 2026' -> '01.09.2026' (跟 HYPER_MEGA 共用的英文月份
    全名對照表 _EN_MONTHS_FULL)。
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


def _express_parse_item_date(text: Optional[str]) -> Optional[str]:
    """'26/Aug/26' (DD/英文月份縮寫/YY) -> '26.08.2026'，年份只有兩位數，
    這張帳單樣本都是 2000 年後的日期，補成 20YY。
    """
    if not text:
        return text
    m = _EXPRESS_DATE_PATTERN.match(text.strip())
    if not m:
        return text
    day, month_abbr, year_2digit = m.groups()
    month = _EN_MONTHS_ABBR.get(month_abbr.lower())
    if not month:
        return text
    return f"{int(day):02d}.{month:02d}.20{year_2digit}"


def _express_clean_amount(text: Optional[str]) -> Optional[int]:
    """英式千分位逗號 '3,737,160' -> 3737160。"""
    if not text:
        return None
    text = text.strip().replace(",", "")
    return int(text) if text.isdigit() else None


def _express_extract_items(text: str) -> List[Dict]:
    """逐行掃描費用明細表格區塊，只認完整的 9 欄費用行 (換行的續行文字對
    不上樣式，直接被跳過)。每筆費用各自帶自己的 arrive_date/onboard_date/
    bl_no/mbl_no (都來自同一行的 DATE/AWB_NO)，交給 expand_to_rows() 做
    逐筆覆蓋。
    """
    items: List[Dict] = []
    for line in text.split("\n"):
        stripped = line.strip()
        m = _EXPRESS_ITEM_PATTERN.match(stripped)
        if not m:
            continue
        date_m = _EXPRESS_DATE_PATTERN.match(stripped)
        item_date = _express_parse_item_date(date_m.group(0)) if date_m else None
        awb = m.group("awb").strip()
        items.append({
            "description": m.group("desc").strip(),
            "amount": _express_clean_amount(m.group("amount")),
            "arrive_date": item_date,
            "onboard_date": item_date,
            "bl_no": awb,
            "mbl_no": awb,
        })
    return items


def parse_express(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {}

    header["supplier"] = _express_add_space_after_pt(
        _search(r"^(PT\.\s*EXPRESS\s*MAXIMUM)", text)
    )
    header["consignee"] = _search(r"COMPANY\s*:\s*(.+?)\s*$", text)
    header["invoice_no"] = _search(r"INVOICE\s+NO\s*:\s*(\S+)", text)
    header["invoice_date"] = _express_parse_long_date(
        _search(r"Tangerang\s*,\s*(\d{1,2}\s+[A-Za-z]+\s+\d{4})", text)
    )

    # 這種帳單沒有 Order No./啟運港/目的港/材積/船名/貨櫃號碼，留空給
    # expand_to_rows() 補 N/A。到達日/開船日/提單號碼/主提單號碼改成逐筆
    # 費用各自帶自己的值 (見下面 _express_extract_items())。
    for code in ("order_no", "port_of_loading", "port_of_discharge",
                 "volume", "vessel", "container_no",
                 "arrive_date", "onboard_date", "bl_no", "mbl_no"):
        header[code] = None

    items = _express_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None}]

    return {"header": header, "items": items}


register_supplier("DWIHARTA", "PT. DWIHARTA LOGISTINDO", detect_dwiharta, parse_dwiharta)
register_supplier("KUEHNE_NAGEL", "KUEHNE NAGEL INDONESIA", detect_kuehne_nagel, parse_kuehne_nagel)
register_supplier("INDOPROSTIME", "PT. INDO PROSTIME EXPRESS", detect_indoprostime, parse_indoprostime)
register_supplier("MAERSK", "PT MAERSK LOGISTICS INDONESIA", detect_maersk, parse_maersk)
register_supplier(
    "HYPER_MEGA", "PT. HYPER MEGA SHIPPING", detect_hyper_mega, parse_hyper_mega,
    multi_page=True,
)
register_supplier(
    "YJE_TATA", "YJE (ShenZhen) International Logistics / PT TATA HARMONI SARANATAMA",
    detect_yje_or_tata, parse_yje_or_tata, multi_page=True,
)
register_supplier(
    "TRANS_DAYA_PRIMA", "PT. TRANS DAYA PRIMA", detect_trans, parse_trans,
    multi_page=True,
)
register_supplier("FEMARIA", "PT. FEMARIA BUANA CARGO", detect_femaria, parse_femaria)
# ===========================================================================
# 供應商 10：PT. PAN EKSPRES INTERNATIONAL (單頁、單張發票，費用明細表格)
# ===========================================================================
#
# 版面特徵：左右兩欄並排 (跟 DWIHARTA 類似)，例如同一列左邊是收件公司/
# 抬頭欄位，右邊緊接著 WO.No/Inv.No/Inv.Date 等欄位，用「遇到下一個已知
# 欄位關鍵字就停止擷取」的方式抓值。CONSIGNEE 比較特殊：「TO :」跟它的值
# 不在同一行 (那一行右邊剛好被 "WO.No : ..." 佔走)，值是接在下一行、
# 銜接到 "Inv.No" 關鍵字前面。
#
# 到達日/開船日分別對應 ETA (估計到達時間) / ETD (估計出發時間)，日期格式
# 是 'DD-Mon-YYYY' (三字英文月份縮寫)，沿用 MAERSK 已定義的
# _EN_MONTHS_ABBR 對照表。
#
# 費用明細行格式："編號 費用名稱 數量 單位 幣別 單價 金額"，例如：
#   "1 HANDLING SERVICE 1 Dokumen IDR 450,000.00 450,000.00"
# 表頭本身就明寫最後一欄是 "Amount"，所以直接取最後一個數字。
def detect_pan_ekspres(text: str) -> bool:
    return "PAN EKSPRES INTERNATIONAL" in text.upper()


_PAN_ITEM_PATTERN = re.compile(
    r"^\d+\s+(?P<desc>.+?)\s+[\d.]+\s+\S+\s+[A-Z]{3}\s+[\d,]+\.\d{2}\s+"
    r"(?P<amount>[\d,]+\.\d{2})$"
)


def _pan_parse_date(text: Optional[str]) -> Optional[str]:
    """'09-Sep-2026' -> '09.09.2026' (跟 MAERSK/TRANS_DAYA_PRIMA 共用的
    英文月份縮寫對照表 _EN_MONTHS_ABBR)。
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


def _pan_clean_amount(text: Optional[str]) -> Optional[int]:
    """'450,000.00' (逗號千分位、句點小數) -> 450000。"""
    if not text:
        return None
    text = text.strip().replace(",", "")
    try:
        return int(round(float(text)))
    except ValueError:
        return None


def _pan_extract_items(text: str) -> List[Dict]:
    items: List[Dict] = []
    for line in text.split("\n"):
        m = _PAN_ITEM_PATTERN.match(line.strip())
        if not m:
            continue
        items.append({
            "description": m.group("desc").strip(),
            "amount": _pan_clean_amount(m.group("amount")),
        })
    return items


def parse_pan_ekspres(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {}

    header["supplier"] = _search(r"^(PT\.\s*PAN\s*EKSPRES\s*INTERNATIONAL)", text)
    header["consignee"] = _search(r"TO\s*:.*?\n\s*(.+?)\s+Inv\.No", text)
    header["invoice_no"] = _search(r"Inv\.No\s*:\s*(\S+)", text)
    header["invoice_date"] = _pan_parse_date(
        _search(r"Inv\.Date\s*:\s*(\d{1,2}-[A-Za-z]{3}-\d{4})", text)
    )
    header["bl_no"] = _search(r"BL\.No\s*:\s*(\S+)", text)
    header["mbl_no"] = header["bl_no"]
    header["vessel"] = _search(r"Vessel\s*:\s*(.+?)\s*$", text)
    header["port_of_loading"] = _search(r"POL\s*:\s*(.+?)\s+No AJU", text)
    header["port_of_discharge"] = _search(r"POD\s*:\s*(.+?)\s+Party", text)
    header["volume"] = _search(r"Party\s*:\s*(.+?)\s*$", text)
    header["arrive_date"] = _pan_parse_date(
        _search(r"ETA\s*:\s*(\d{1,2}-[A-Za-z]{3}-\d{4})", text)
    )
    header["onboard_date"] = _pan_parse_date(
        _search(r"ETD\s*:\s*(\d{1,2}-[A-Za-z]{3}-\d{4})", text)
    )

    # 這種帳單沒有 Order No./貨櫃號碼 (Container No 這欄印出來但值是空的)，
    # 留空給 expand_to_rows() 補 N/A。
    header["order_no"] = None
    header["container_no"] = None

    items = _pan_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None}]

    return {"header": header, "items": items}


register_supplier("EXPRESS_MAXIMUM", "PT. EXPRESS MAXIMUM", detect_express, parse_express)
register_supplier(
    "PAN_EKSPRES", "PT. PAN EKSPRES INTERNATIONAL",
    detect_pan_ekspres, parse_pan_ekspres,
)


# ===========================================================================
# 供應商 11：PT. PANCARAN SRIKANDI LOGISTIK (單頁、單張發票，費用明細表格)
# ===========================================================================
#
# 版面特徵：
#   - Invoice No 的標籤跟值分兩行印出 ("Invoice No:" 換行才是
#     "INV.2609.0014")。
#   - REFERENCE NO. (對應 Order No.) 跟 Container No. 這兩個欄位的值都很
#     長，會換行接到下一行，而且換行處常常跟右邊另一欄的文字擠在同一行
#     (例如 REFERENCE NO. 那一行右邊接著 Qty/Commodity 說明文字)。正確
#     答案 Excel 裡這兩欄本身也是保留原本 PDF 斷行位置的多行文字 (用
#     '\n' 分段，比對時 normalize_value() 會把換行當空白處理，所以只要
#     斷行的文字內容跟順序一樣，斷在哪一行不影響比對結果)。
#   - REFERENCE NO. 只取「值」的第一個字詞 (中間用空白隔開的字才是接在
#     它後面的 Qty/Commodity 說明，不算 Order No. 的一部分)，續行同理只
#     取續行的第一個字詞。
#   - Container No. 沒有额外文字混在同一行，續行只要是「全部由大寫字母/
#     數字/逗號組成」的整行文字，就當作是還在延續同一個貨櫃號碼清單，
#     遇到不是這種格式的行 (例如換頁後單獨一個逗號、或後面的 Amount
#     表格) 就停止。
def detect_pancaran(text: str) -> bool:
    return "PANCARAN SRIKANDI LOGISTIK" in text.upper()


_PANCARAN_ITEM_PATTERN = re.compile(
    r"^\d+\s+(?P<desc>.+?)\s+\d+\s+[\d,]+\s+(?P<amount>[\d,]+)$"
)


def _pancaran_clean_amount(text: Optional[str]) -> Optional[int]:
    """'75,660,000' (逗號千分位、無小數) -> 75660000。"""
    if not text:
        return None
    text = text.strip().replace(",", "")
    return int(text) if text.isdigit() else None


def _pancaran_extract_order_no(text: str) -> Optional[str]:
    lines = text.split("\n")
    for i, line in enumerate(lines):
        m = re.search(r"REFERENCE\s+NO\.\s*:\s*(\S+)", line)
        if not m:
            continue
        parts = [m.group(1)]
        if i + 1 < len(lines):
            nxt = lines[i + 1].strip()
            if nxt.startswith("/"):
                m2 = re.match(r"(\S+)", nxt)
                if m2:
                    parts.append(m2.group(1))
        return "\n".join(parts)
    return None


def _pancaran_extract_container_no(text: str) -> Optional[str]:
    lines = text.split("\n")
    for i, line in enumerate(lines):
        m = re.search(r"JOB\s+NO\.\s*:\s*\S+\s+(.+)$", line)
        if not m:
            continue
        parts = [m.group(1).strip()]
        j = i + 1
        while j < len(lines):
            nxt = lines[j].strip()
            if len(nxt) > 1 and re.fullmatch(r"[A-Z0-9,]+", nxt):
                parts.append(nxt)
                j += 1
            else:
                break
        return "\n".join(parts)
    return None


def _pancaran_extract_items(text: str) -> List[Dict]:
    items: List[Dict] = []
    for line in text.split("\n"):
        m = _PANCARAN_ITEM_PATTERN.match(line.strip())
        if not m:
            continue
        items.append({
            "description": m.group("desc").strip(),
            "amount": _pancaran_clean_amount(m.group("amount")),
        })
    return items


def parse_pancaran(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {}

    header["supplier"] = _search(r"A\.N\.\s*(PT\..+?)\s*$", text)
    header["consignee"] = _search(r"CONSIGNEE\s*:\s*(.+?)\s*$", text)
    header["invoice_no"] = _search(r"Invoice\s+No\s*:\s*\n\s*(\S+)", text)
    header["invoice_date"] = _express_parse_long_date(
        _search(r"DATE\s*:\s*(\d{1,2}\s+[A-Za-z]+\s+\d{4})", text)
    )
    header["order_no"] = _pancaran_extract_order_no(text)
    header["container_no"] = _pancaran_extract_container_no(text)

    # 這種帳單沒有到達日/開船日/提單號碼/主提單號碼/啟運港/目的港/材積/
    # 船名，留空給 expand_to_rows() 補 N/A。
    for code in ("arrive_date", "onboard_date", "bl_no", "mbl_no",
                 "port_of_loading", "port_of_discharge", "volume", "vessel"):
        header[code] = None

    items = _pancaran_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None}]

    return {"header": header, "items": items}


register_supplier(
    "PANCARAN", "PT. PANCARAN SRIKANDI LOGISTIK",
    detect_pancaran, parse_pancaran,
)


# ===========================================================================
# 供應商 12：PT BIROTIKA SEMESTA / DHL EXPRESS (掃描檔 PDF，需要 OCR)
# ===========================================================================
#
# 這是本系統第一家「掃描檔」供應商：PDF 本身沒有文字層 (extract_text()
# 一律回傳空字串)，全部文字都要靠 extract_utils.py 的 OCR 備援機制
# (_ocr_page_text()) 辨識出來，這支 parse 函式只要跟其他供應商一樣處理
# 「已經是文字」的內容就好，不用自己碰 OCR。
#
# 這份帳單一張發票橫跨兩頁掃描圖檔：第 1 頁是「服務類型/金額」總覽表
# (品項/金額都在這一頁，OCR 品質也最好)；第 2 頁是「Air Waybill/Shippers
# Reference/Shipment Origin Date」明細表，且第 2 頁整頁被掃描成轉了 90 度
# (OCR 備援機制裡的方向偵測會自動轉正，這裡不用處理)，轉正後 OCR 出來的
# 文字順序仍然會把同一列的欄位「擠成一行」(不像原始表格那樣分欄)，但因為
# 一張發票只有一列資料，用「這一行前三個數字依序是 Air Waybill/Shippers
# Reference/Shipment Origin Date」的位置規則就能可靠取出。
#
# 已知的 OCR 誤判：Invoice Number 裡的數字 '0' 常被辨識成英文字母 'O'
# (例如 'JKTIR00818492' 被讀成 'JKTIROO818492' 或 'JKTIRO0818492')，這個
# 發票編號的格式固定是「英文字母開頭 + 純數字」，前面的英文字母部分剛好
# 不含字母 O，所以直接把整個擷取到的編號裡的 'O' 全部換成 '0' 是安全的
# 修正方式；之後如果遇到其他 OCR 常誤判的字元，比照這個做法在擷取後加一
# 個修正函式即可。
def detect_dhl(text: str) -> bool:
    upper = text.upper()
    return "BIROTIKA SEMESTA" in upper or "DHL EXPRESS" in upper


_DHL_ITEM_PATTERN = re.compile(r"^([A-Z][A-Z /]+[A-Z])\s+([\d,]+)$")
_DHL_AWB_ROW_PATTERN = re.compile(
    r"(\d{6,12})\s+(\d{6,12})\s+(\d{1,2}-\d{1,2}-\d{4})\s+JKT"
)


def _dhl_fix_ocr_zero(value: Optional[str]) -> Optional[str]:
    """OCR 常把發票編號裡的數字 '0' 認成英文字母 'O'，這裡統一換回來。"""
    if not value:
        return value
    return value.replace("O", "0")


def _dhl_parse_date(text: Optional[str]) -> Optional[str]:
    """'31-12-2024' (DD-MM-YYYY) -> '31.12.2024'。"""
    if not text:
        return text
    m = re.match(r"(\d{1,2})-(\d{1,2})-(\d{4})", text.strip())
    if not m:
        return text
    day, month, year = m.groups()
    return f"{int(day):02d}.{int(month):02d}.{year}"


def _dhl_clean_amount(text: Optional[str]) -> Optional[int]:
    """'662,295' (逗號千分位、無小數) -> 662295。"""
    if not text:
        return None
    text = text.strip().replace(",", "")
    return int(text) if text.isdigit() else None


_DHL_SUMMARY_LINE_PATTERN = re.compile(
    r"^([A-Z][A-Z /&]+[A-Z])\s+\d+\s+[\d.]+\s+\d+\s+([\d,]+)\s+[\d,]+\s+[\d,]+\s+[\d,]+$"
)


def _dhl_extract_items(text: str) -> List[Dict]:
    """擷取每一筆費用項目：

    1. 第 1 頁最上面「Type of Service」彙總表的服務類型列 (例如
       "EXPRESS WORLDWIDE NONDOC 6 23.00 8 6,292,646 3,083,135 112,509
       9,488,290")，欄位依序是服務名稱/件數/總重量/項目數/Standard
       Shipping Charge/Extra Charges/VAT/Total(含稅)，我們要的金額是
       「Standard Shipping Charge」欄，不是最後含稅總額；如果這欄是 0
       (代表這張帳單沒有基本運費，只有額外費用，例如關稅類帳單) 就不算
       一個項目。
    2. 「Analysis of Extra Charges」區塊逐行擷取費用名稱/金額，這個區塊
       OCR 品質最好、沒有跟其他欄位擠在同一行，比第 2 頁明細表可靠。只
       在「Analysis of Extra Charges」到「Total Extra Charges」這個範圍
       內找，避免誤吃到後面付款資訊裡剛好也符合「大寫字+空白+數字」格式
       的行 (例如匯款帳號 "ACCOUNT NO 956907730")。
    """
    items: List[Dict] = []
    lines = text.split("\n")

    for line in lines:
        m = _DHL_SUMMARY_LINE_PATTERN.match(line.strip())
        if not m:
            continue
        amount = _dhl_clean_amount(m.group(2))
        if amount:  # 0 或抓不到就不算一筆項目
            items.append({"description": m.group(1).strip(), "amount": amount})
        break  # 一張帳單只有一種服務類型的彙總列

    in_extra_section = False
    for line in lines:
        stripped = line.strip()
        if stripped.upper().startswith("ANALYSIS OF EXTRA CHARGES"):
            in_extra_section = True
            continue
        if not in_extra_section:
            continue
        if stripped.upper().startswith("TOTAL EXTRA CHARGES"):
            break
        m = _DHL_ITEM_PATTERN.match(stripped)
        if not m:
            continue
        desc = m.group(1).strip()
        if desc in ("TOTAL EXTRA CHARGES", "TOTAL DISCOUNTS", "TOTAL VAT"):
            continue
        items.append({
            "description": desc,
            "amount": _dhl_clean_amount(m.group(2)),
        })
    return items


def parse_dhl(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {}

    # 收件公司名稱位置有兩種版面：
    # (a) OCR 掃描版：信封抬頭最上面一行就是公司名稱，跟其他欄位標籤各佔
    #     一整行，直接抓整份文字的第一行即可。
    # (b) 有內嵌文字層的版面：公司名稱跟 "Invoice Number:" 標籤同一行
    #     (例如 "POU YUEN INDONESIA PT Invoice Number: BDOIR00043036")，
    #     這時第一行其實是 "DHL Express" 這種文件抬頭，不是公司名稱，要
    #     改抓 "Invoice Number:" 前面那一段文字。
    first_line = text.split("\n", 1)[0].strip()
    m = re.search(r"^([^\n]+?)[ \t]+Invoice[ \t]+Number[ \t]*:", text, re.MULTILINE)
    header["consignee"] = m.group(1).strip() if m else (first_line or None)
    header["supplier"] = _search(r"(PT\s+BIROTIKA\s+SEMESTA)\s*/\s*DHL\s+EXPRESS", text)
    if header["supplier"]:
        header["supplier"] = f"{header['supplier']} / DHL EXPRESS"
    # OCR 把這四個標籤跟四個值分別掃描成「標籤欄一整排、值欄一整排」
    # (先四行標籤：Invoice Number:/Account Number:/Tax ID:/Invoice Date:，
    # 空一行，再四行對應的值)，不是每個標籤緊接著自己的值，所以不能直接
    # 用「Invoice Number: 後面第一個字」抓 (那樣會抓到下一個標籤
    # "Account")，要整塊比對位置對應。
    m = re.search(
        r"Invoice\s+Number\s*:\s*\n"
        r"Account\s+Number\s*:\s*\n"
        r"Tax\s+ID\s*:\s*\n"
        r"Invoice\s+Date\s*:\s*\n\s*\n"
        r"(\S+)\n(\S+)\n(\S+)\n(\d{1,2}-\d{1,2}-\d{4})",
        text,
    )
    if m:
        header["invoice_no"] = _dhl_fix_ocr_zero(m.group(1))
        header["invoice_date"] = _dhl_parse_date(m.group(4))
    else:
        # 備援：萬一標籤/值的行數對不上 (OCR 結果不穩定)，改用發票號碼
        # 固定格式 (JKTxxx開頭) 直接在全文找，日期則從 "Invoice Date:"
        # 後面找最近的一個日期。
        header["invoice_no"] = _dhl_fix_ocr_zero(
            _search(r"\b([A-Z]{2,6}[O0-9]{6,})\b", text)
        )
        header["invoice_date"] = _dhl_parse_date(
            _search(r"Invoice\s+Date\s*:\s*\n?\s*(\d{1,2}-\d{1,2}-\d{4})", text)
        )

    m = _DHL_AWB_ROW_PATTERN.search(text)
    if m:
        header["bl_no"] = m.group(1)
        header["order_no"] = m.group(2)
        header["onboard_date"] = _dhl_parse_date(m.group(3))
    else:
        header["bl_no"] = None
        header["order_no"] = None
        header["onboard_date"] = None

    # 這種帳單沒有到達日/主提單號碼/啟運港/目的港/材積/船名/貨櫃號碼，
    # 留空給 expand_to_rows() 補 N/A。
    for code in ("arrive_date", "mbl_no", "port_of_loading",
                 "port_of_discharge", "volume", "vessel", "container_no"):
        header[code] = None

    items = _dhl_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None}]

    return {"header": header, "items": items}


register_supplier(
    "DHL", "PT BIROTIKA SEMESTA / DHL EXPRESS",
    detect_dhl, parse_dhl,
)


# ===========================================================================
# 供應商 13：PT DSV Transport Indonesia (multi_page，一份 PDF 裡有好幾張
# 各自獨立的發票，但『一張發票』實際上橫跨兩個實體頁面)
# ===========================================================================
#
# 這份 PDF 版面比較特別：一張發票印出來是 2 頁 ("Page 1 of 2"/"Page 2 of
# 2")，一份 PDF 又可能塞好幾張這樣的發票 (這次樣本是 3 張發票、共 6 頁)。
# 好在每張發票需要的欄位 (抬頭 + 費用明細金額) 全部都印在「Page 1 of 2」
# 那一頁上，「Page 2 of 2」只有 SUBTOTAL/PPN/TOTAL 這種彙總數字跟匯款
# 資訊，我們的 16 個標準欄位都用不到，所以不用真的把兩頁文字合併，讓
# detect_dsv() 只認「Page 1 of 2」那一頁 (靠這頁才有的 "CHARGES IN IDR"
# 表頭關鍵字判斷)，multi_page 架構逐頁辨識時「Page 2 of 2」會被跳過，剛好
# 就是我們要的效果 (不會多出一筆重複/空白的列)。
#
# 費用明細每一項的說明文字會換行 (例如 "THC (Terminal Handling Charge) -
# Greater of (Min\nRate IDR 73427,00, 8,471 Cubic Meter(s) @ IDR\n
# 73427,00/M3)")，但只有「第一行」的結尾會是這一項的 PPN 金額 + 正式金額
# 兩個數字，續行都是說明文字沒有這種「結尾兩個數字」的樣式，所以逐行比對
# 就能只抓到每一項的第一行，不用特別處理續行合併。正確答案的 Amount 對應
# 的是 "CHARGES IN IDR" 這一欄 (行尾最後一個數字)，不是前面的
# "PPN IN IDR" 那一欄。
def detect_dsv(text: str) -> bool:
    return "CHARGES IN IDR" in text.upper()


_DSV_ITEM_PATTERN = re.compile(
    r"^(?P<desc>.+?)\s+-\s+.*?\s+(?P<ppn>[\d.]+)\s+(?P<amount>[\d.]+)$",
    re.MULTILINE,
)
_DSV_VESSEL_BL_PATTERN = re.compile(
    r"^(?P<vessel>.+?)\s*/\s*(?P<voyage>\S+)\s*/\s*\S+\s*/\s*\S+\s+"
    r"(?P<ocean_bl>\S+)\s+(?P<house_bl>\S+)$",
    re.MULTILINE,
)


def _dsv_parse_date(text: Optional[str]) -> Optional[str]:
    """'03-Sep-26' (DD-英文月份縮寫-YY，兩位數年份) -> '03.09.2026'，沿用
    既有的 _EN_MONTHS_ABBR 對照表，年份補上世紀 (這批帳單都是 20XX 年)。
    """
    if not text:
        return text
    m = re.match(r"(\d{1,2})-([A-Za-z]{3})-(\d{2})", text.strip())
    if not m:
        return text
    day, month_abbr, year_2digit = m.groups()
    month = _EN_MONTHS_ABBR.get(month_abbr.lower())
    if not month:
        return text
    return f"{int(day):02d}.{month:02d}.20{year_2digit}"


def _dsv_clean_amount(text: Optional[str]) -> Optional[int]:
    """'3.140.183' (句點千分位、無小數) -> 3140183。"""
    if not text:
        return None
    text = text.strip().replace(".", "")
    return int(text) if text.isdigit() else None


def _dsv_extract_items(text: str) -> List[Dict]:
    items: List[Dict] = []
    for m in _DSV_ITEM_PATTERN.finditer(text):
        items.append({
            "description": m.group("desc").strip(),
            "amount": _dsv_clean_amount(m.group("amount")),
        })
    return items


def parse_dsv(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {}

    header["supplier"] = "PT DSV Transport Indonesia"
    header["invoice_no"] = _search(r"INVOICE\s+(\S+)\s+ORIGINAL", text)

    m = re.search(
        r"^(.+?)\s+INVOICE DATE\s+(\d{1,2}-[A-Za-z]{3}-\d{2})", text, re.MULTILINE
    )
    if m:
        header["consignee"] = m.group(1).strip()
        header["invoice_date"] = _dsv_parse_date(m.group(2))
    else:
        header["consignee"] = None
        header["invoice_date"] = None

    header["order_no"] = _search(r"OWNER'S REFERENCE\s*\n\s*(\S+)", text)

    m = re.search(
        r"=\s*([^,]+),.*?(\d{1,2}-[A-Za-z]{3}-\d{2})\s+\S+\s*=\s*([^,]+),.*?"
        r"(\d{1,2}-[A-Za-z]{3}-\d{2})",
        text,
    )
    if m:
        header["port_of_loading"] = m.group(1).strip()
        header["onboard_date"] = _dsv_parse_date(m.group(2))
        header["port_of_discharge"] = m.group(3).strip()
        header["arrive_date"] = _dsv_parse_date(m.group(4))
    else:
        header["port_of_loading"] = None
        header["onboard_date"] = None
        header["port_of_discharge"] = None
        header["arrive_date"] = None

    m = _DSV_VESSEL_BL_PATTERN.search(text)
    if m:
        header["vessel"] = f"{m.group('vessel').strip()} / {m.group('voyage').strip()}"
        header["bl_no"] = m.group("house_bl").strip()
        header["mbl_no"] = m.group("ocean_bl").strip()
    else:
        header["vessel"] = None
        header["bl_no"] = None
        header["mbl_no"] = None

    m = re.search(r"^(\S+)\s*\(", text, re.MULTILINE)
    header["container_no"] = m.group(1) if m else None

    m = re.search(r"[\d.]+\s*KG\s+([\d.]+)\s*M3", text)
    header["volume"] = m.group(1) if m else None

    items = _dsv_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None}]

    return {"header": header, "items": items}


register_supplier(
    "DSV", "PT DSV Transport Indonesia",
    detect_dsv, parse_dsv, multi_page=True,
)
