# -*- coding: utf-8 -*-
"""
expense_normalize.py
====================
費用分析系統的資料清洗規則 (純函式，不碰檔案/資料庫，方便單獨測試)：

  - parse_date()       ：各種日期寫法 (31.Aug.2026、2026-08-31、31/08/2026、
                         印尼文月份、Excel 日期序號...) 轉成 datetime.date
  - parse_amount()     ：金額/稅額 (含 Rp、千分位逗號或點、括號負數) 轉成 float
  - parse_float()      ：一般小數 (點一律當小數點)
  - parse_volume()     ：材積原文 -> 櫃型櫃數 / TEU / CBM / KG (FCL > CBM > KG)
  - normalize_expense():費用名稱正規化，回傳 (標準費用項目, 費用大類)
  - normalize_port()   ：港口名稱統一大小寫與空白
  - detect_currency()  ：金額幣別 (帳單上的幣別 > 供應商 > 國家資料夾
                         folder_currency()；都判斷不出來回傳 None，詢問使用者)
  - clean_text()       ：一般文字欄位去空白，N/A / - / 空字串一律變成 None

要新增「同性質費用合併」的對照，改 EXPENSE_ALIASES 即可；要調整費用大類，
改 EXPENSE_CATEGORIES (由上往下比對，先符合的先用)。
"""

import re
import unicodedata
from datetime import date, datetime, timedelta
from typing import Dict, Optional, Tuple

# 視為「沒有值」的佔位字
_NA_VALUES = {"", "N/A", "NA", "N.A.", "-", "--", "NONE", "NULL", "NAN", "NIL"}


def clean_text(value) -> Optional[str]:
    """一般文字欄位：去頭尾空白、合併連續空白；佔位字一律回傳 None。
    數字型態的發票號碼 (例如 5617840405) 轉成不帶 .0 的字串。"""
    if value is None:
        return None
    if isinstance(value, float):
        if value != value:  # NaN
            return None
        if value.is_integer():
            value = int(value)
    text = re.sub(r"\s+", " ", str(value)).strip()
    if text.upper() in _NA_VALUES:
        return None
    return text


# ============================================================
# 日期
# ============================================================

# 印尼文/英文全名月份 -> 英文三碼縮寫，讓 strptime 的 %b 認得
_MONTH_WORDS = {
    "JANUARI": "Jan", "JANUARY": "Jan", "FEBRUARI": "Feb", "FEBRUARY": "Feb",
    "MARET": "Mar", "MARCH": "Mar", "APRIL": "Apr", "MEI": "May", "JUNI": "Jun",
    "JUNE": "Jun", "JULI": "Jul", "JULY": "Jul", "AGUSTUS": "Aug", "AUGUST": "Aug",
    "AGU": "Aug", "AGT": "Aug", "AGS": "Aug", "SEPTEMBER": "Sep", "SEPT": "Sep",
    "OKTOBER": "Oct", "OCTOBER": "Oct", "OKT": "Oct", "NOVEMBER": "Nov",
    "NOPEMBER": "Nov", "NOP": "Nov", "DESEMBER": "Dec", "DECEMBER": "Dec", "DES": "Dec",
}
_MONTH_RE = re.compile(r"[A-Za-z]+")

# 由上往下嘗試；日在前的寫法優先 (印尼帳單習慣 日/月/年)
_DATE_FORMATS = (
    "%d.%b.%Y", "%d-%b-%Y", "%d %b %Y", "%d/%b/%Y", "%d-%b-%y", "%d %b %y",
    "%b %d %Y", "%b %d, %Y", "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d",
    "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%m-%y", "%Y%m%d",
)

# Excel 日期序號的合理範圍 (約 1982 ~ 2064 年)，避免把金額誤當日期
_EXCEL_SERIAL_MIN, _EXCEL_SERIAL_MAX = 30000, 60000
_EXCEL_EPOCH = datetime(1899, 12, 30)


def _month_to_abbr(match: re.Match) -> str:
    word = match.group(0)
    abbr = _MONTH_WORDS.get(word.upper())
    if abbr:
        return abbr
    # 英文三碼縮寫統一成首字大寫 (AUG -> Aug)
    return word[:1].upper() + word[1:].lower() if len(word) == 3 else word


def parse_date(value) -> Optional[date]:
    """把各種日期寫法轉成 date；轉不出來回傳 None。"""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        if value != value:  # NaN
            return None
        if _EXCEL_SERIAL_MIN <= value <= _EXCEL_SERIAL_MAX:
            return (_EXCEL_EPOCH + timedelta(days=float(value))).date()
        return None

    text = clean_text(value)
    if not text:
        return None
    # 去掉時間部分 ("2026-08-31 00:00:00"、"2026-08-31T10:00")
    text = re.split(r"[ T]\d{1,2}:\d{2}", text)[0].strip()
    text = _MONTH_RE.sub(_month_to_abbr, text)
    text = re.sub(r"\s*([./-])\s*", r"\1", text)  # "31 . Aug . 2026" -> "31.Aug.2026"
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


# ============================================================
# 數字
# ============================================================

_NUM_KEEP_RE = re.compile(r"[^0-9,.\-]")


def parse_amount(value) -> Optional[float]:
    """金額/稅額轉成 float；轉不出來回傳 None。

    千分位判斷 (印尼帳單常見 1.234.567 或 1,234,567，也可能 1.234.567,89)：
      - 同時有 . 和 , ：最後出現的那個是小數點，另一個是千分位
      - 只有一種符號且出現多次：千分位
      - 只有一種符號出現一次：後面剛好 3 位數字時當千分位 (IDR 金額沒有
        三位小數)，否則當小數點
    """
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return None if value != value else float(value)

    text = str(value).strip()
    if not text or text.upper() in _NA_VALUES:
        return None
    negative = text.startswith("(") and text.endswith(")")
    text = _NUM_KEEP_RE.sub("", text)
    if text.endswith("-"):  # 會計格式的負號放在後面 "1,234-"
        negative, text = True, text[:-1]
    if text.startswith("-"):
        negative, text = True, text[1:]
    if not text or not re.search(r"\d", text):
        return None

    has_dot, has_comma = "." in text, "," in text
    if has_dot and has_comma:
        decimal = "." if text.rfind(".") > text.rfind(",") else ","
        thousands = "," if decimal == "." else "."
        text = text.replace(thousands, "").replace(decimal, ".")
    elif has_dot or has_comma:
        sep = "." if has_dot else ","
        parts = text.split(sep)
        if len(parts) > 2 or len(parts[-1]) == 3:
            text = text.replace(sep, "")
        else:
            text = text.replace(sep, ".")
    try:
        number = float(text)
    except ValueError:
        return None
    return -number if negative else number


def parse_float(value) -> Optional[float]:
    """材積等一般數字：點一律當小數點，逗號當千分位。"""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return None if value != value else float(value)
    text = clean_text(value)
    if not text:
        return None
    match = re.search(r"-?\d[\d,]*(?:\.\d+)?", text)
    if not match:
        return None
    try:
        return float(match.group(0).replace(",", ""))
    except ValueError:
        return None


# 材積 (Volume) 拆成數字 (使用者 2026-10-27)。轉檔結果的寫法 (所有供應商)：
#   FCL「數量X櫃型」(1X40HC, 多種用 ", " 隔開) > LCL「CBM M3」(7.010 M3) > 只有重量「KG」(2082 KG)
#   舊格式只有數字沒有單位 (例如 MAERSK 47.545)：當成 CBM。
# TEU：20 呎 = 1，40/45 呎 = 2。
CONTAINER_TYPES = ("20GP", "40GP", "40HC", "45HC")
_FCL_RE = re.compile(r"(\d+)\s*[X×*]\s*(20|40|45)\s*['’]?\s*(?:FT\b|FEET\b)?\s*([A-Z]{2})?",
                     re.IGNORECASE)
_CBM_RE = re.compile(r"(\d[\d,]*(?:\.\d+)?)\s*(?:M3|M³|CBM)\b", re.IGNORECASE)
_KG_RE = re.compile(r"(\d[\d,]*(?:\.\d+)?)\s*(?:KGS?|KILOS?)\b", re.IGNORECASE)
_BARE_NUMBER_RE = re.compile(r"\d[\d,]*(?:\.\d+)?")
# 運送型態字樣 (使用者 2026-10-27)：以整張帳單 (費用名稱 + 材積) 判斷，兩種都有以 FCL 為準
#   有 FCL 或櫃型 (40HC、20' GP、1X40) = 整櫃；有 CFS、CBM (或 LCL) = 散貨
FCL_WORD_RE = re.compile(r"\bFCL\b|\b(?:20|40|45)\s*(?:'|FT|FEET)?\s*"
                         r"(?:GP|DC|DV|HC|HQ|RF|RH|OT|FR|ST|SD)\b|\b\d+\s*X\s*(?:20|40|45)\b",
                         re.IGNORECASE)
LCL_WORD_RE = re.compile(r"\bCFS\b|\bCBM\b|\bLCL\b", re.IGNORECASE)


def _container_type(size: str, code: Optional[str]) -> str:
    """呎數 + 櫃型代碼 -> CONTAINER_TYPES (高櫃 HC/HQ 算 40HC，其餘 40 呎算 40GP)。"""
    if size == "20":
        return "20GP"
    if size == "45":
        return "45HC"
    return "40HC" if (code or "").upper() in ("HC", "HQ", "RH") else "40GP"


def parse_volume(value) -> Dict[str, object]:
    """材積原文 -> {"mode": FCL/LCL/空運快遞/None, "20GP".."45HC": 櫃數, "teu", "cbm", "kg"}。
    依 FCL > CBM > KG 的順序只取一種；看不懂回傳全空。"""
    out: Dict[str, object] = {"mode": None, "teu": None, "cbm": None, "kg": None}
    out.update({t: None for t in CONTAINER_TYPES})
    if value is None or isinstance(value, bool):
        return out
    if isinstance(value, (int, float)):
        if value == value and value > 0:
            out.update(mode="LCL", cbm=float(value))
        return out
    text = clean_text(value)
    if not text:
        return out
    fcl = _FCL_RE.findall(text)
    if fcl:
        for qty, size, code in fcl:
            t = _container_type(size, code)
            out[t] = (out[t] or 0) + int(qty)
        out["teu"] = float(sum((out[t] or 0) * (1 if t == "20GP" else 2) for t in CONTAINER_TYPES))
        out["mode"] = "FCL"
        return out
    for regex, key, mode in ((_CBM_RE, "cbm", "LCL"), (_KG_RE, "kg", "空運快遞")):
        m = regex.search(text)
        if m:
            out.update({key: float(m.group(1).replace(",", "")), "mode": mode})
            return out
    if _BARE_NUMBER_RE.fullmatch(text):
        number = float(text.replace(",", ""))
        if number > 0:
            out.update(mode="LCL", cbm=number)
    return out


# ============================================================
# 費用名稱
# ============================================================

# 縮寫/複數/錯字展開 (在去標點、轉大寫之後套用，由上往下)
_EXPANSIONS = (
    (r"\bO\s?THC\b", "ORIGIN TERMINAL HANDLING CHARGE"),
    (r"\bD\s?THC\b", "DESTINATION TERMINAL HANDLING CHARGE"),
    (r"\bTHC\b", "TERMINAL HANDLING CHARGE"),
    (r"\bTERMINALHANDLINGCHARGE\b", "TERMINAL HANDLING CHARGE"),
    (r"\bCHA?RG(?:E|ES)?\b|\bCHGS?\b|\bCHRGS?\b", "CHARGE"),
    (r"\bFEES\b", "FEE"),
    (r"\bDOCS?\b|\bDOCUMENTS?\b|\bDOKUMEN\b", "DOCUMENTATION"),
    (r"\bB\s?L\b", "BL"),
    (r"\bCOSTOMS\b|\bCUSTOM\b", "CUSTOMS"),
    (r"\bWEIGHT\b", "WEIGHING"),
    (r"\bADM\b|\bADMIN\b", "ADMINISTRATION"),
    (r"\bTRANPORTATION\b", "TRANSPORTATION"),  # CONG THANH 對帳單標題的拼字
    (r"\bLIFTOFF\b", "LIFT OFF"),
)
_EXPANSION_RES = tuple((re.compile(p), r) for p, r in _EXPANSIONS)

# 費用名稱裡要拿掉的雜訊：括號內容、追蹤號碼這類「字母+長串數字」、
# 純數字/金額、幣別
_NOISE_RES = (
    re.compile(r"\(.*?\)"),
    re.compile(r"\b[A-Z]{0,4}\d{6,}\b"),
    re.compile(r"\b(?:USD|IDR|RP|SGD|EUR)\b"),
    # 備註用的日期，例如 'ADMINISTRATION FEE - 9 okt'
    re.compile(r"\b\d{1,2}\s*(?:JAN|FEB|MAR|APR|MAY|MEI|JUN|JUL|AUG|AGU|AGT|SEP|"
               r"OCT|OKT|NOV|NOP|DEC|DES)[A-Z]*\b(?:\s*\d{2,4}\b)?"),
    # 附加費的費率，例如越南燃油附加費 'PHỤ PHÍ XĂNG DẦU UP 12%' / 'UP 15%'
    re.compile(r"\bUP\s*\d+(?:[.,]\d+)?\s*%?"),
)

# 正規化後的 key -> 統一顯示名稱 (同性質費用合併在這裡加)
EXPENSE_ALIASES = {
    "CFS RECEIVING CHARGE": "CFS Receiving Charge",
    "CFS CHARGE": "CFS Receiving Charge",
    "CFS RECEIVING FEE": "CFS Receiving Charge",
    "ORIGIN TERMINAL HANDLING CHARGE": "Origin Terminal Handling Charge",
    "TERMINAL HANDLING CHARGE ORIGIN": "Origin Terminal Handling Charge",
    "DESTINATION TERMINAL HANDLING CHARGE": "Destination Terminal Handling Charge",
    "TERMINAL HANDLING CHARGE DESTINATION": "Destination Terminal Handling Charge",
    "FCR ISSUANCE FEE": "FCR Issuance Fee",
    "FCR FEE": "FCR Issuance Fee",
    "FCR ISSUANCE CHARGE": "FCR Issuance Fee",
    "VGM WEIGHING FEE": "VGM Weighing Fee",
    "VGM WEIGHING CHARGE": "VGM Weighing Fee",
    "VGM FILING FEE": "VGM Filing Fee",
    "VGM FILING CHARGE": "VGM Filing Fee",
    "VGM FEE": "VGM Fee",
    "ORIGIN DOCUMENTATION FEE": "Origin Documentation Fee",
    "ORIGIN DOCUMENTATION CHARGE": "Origin Documentation Fee",
    "DOCUMENTATION FEE": "Documentation Fee",
    "DOCUMENTATION CHARGE": "Documentation Fee",
    # 越南文費用名稱 (去掉聲調後比對)
    "PHI VAN CHUYEN": "Transportation Fee (Phí vận chuyển)",
    "PHI DICH VU": "Service Fee (Phí dịch vụ)",
    "PHU PHI XANG DAU": "Fuel Surcharge (Phụ phí xăng dầu)",
    "PHAT SINH": "Incurred (Phát sinh)",
    "NANG HA": "Lift On/Off (Nâng hạ)",
    "LUU BAI": "Storage (Lưu bãi)",
    "PT HA BAI SOM": "Early Drop Fee (Phí hạ bãi sớm)",
    "CSHT": "Port Infrastructure Fee (CSHT)",
    "D O H L TERMINAL HANDLING CHARGE CIC FEE": "D/O, H/L, THC, CIC Fee",
    "FREIGHT SEA PER P L": "Freight Sea (per P/L)",
    "FREIGHT AIR PER P L": "Freight Air (per P/L)",
}

# 費用大類：(大類名稱, 關鍵字 regex)，由上往下比對，先符合的先用
EXPENSE_CATEGORIES = (
    ("碼頭處理費 (THC)", r"TERMINAL HANDLING|\bCSHT\b|\bD O H L\b"),
    ("VGM 費用", r"\bVGM\b"),
    ("CFS / 倉儲", r"\bCFS\b|WAREHOUSE|GUDANG|STORAGE|STUFFING|UNSTUFFING|LUU BAI"),
    ("文件 / 提單費", r"DOCUMENTATION|\bFCR\b|\bCBL\b|\bBL\b|BILL OF LADING|\bAWB\b|\bEDI\b|"
                     r"ISSUANCE|\bSEAL\b|SHIPPING INSTRUCTION"),
    ("運費", r"FREIGHT|\bBAF\b|\bCAF\b|BUNKER|SURCHARGE"),
    ("報關 / 檢驗", r"CUSTOMS|CLEARANCE|DECLARATION|\bPEB\b|\bPIB\b|INSPECTION|QUARANTINE|"
                   r"FUMIGATION"),
    ("拖車 / 陸運 / 滯櫃", r"TRUCKING|HAULAGE|DELIVERY|TRANSPORT|PICK ?UP|LIFT ON|LIFT OFF|"
                         r"DETENTION|DEMURRAGE|VAN CHUYEN|XANG DAU|NANG HA|HA BAI"),
    ("手續 / 管理 / 行政", r"HANDLING|AGENCY|ADMINISTRATION|SERVICE|COMMISSION|MANAGEMENT|"
                         r"BOOKING|COMPLIAN|LATE PAYMENT|DICH VU"),
)
_CATEGORY_RES = tuple((name, re.compile(p)) for name, p in EXPENSE_CATEGORIES)
OTHER_CATEGORY = "其他"
UNKNOWN_EXPENSE = "(未填費用名稱)"

# 顯示名稱用 Title Case 時要保持全大寫的縮寫
_ACRONYMS = {"CFS", "VGM", "FCR", "CBL", "THC", "BL", "AWB", "EDI", "PEB", "PIB", "LCL",
             "FCL", "DO", "PPN", "VAT", "BAF", "CAF", "ISPS", "EMKL", "AE", "LSS",
             "TEU", "FEU", "DG", "USD", "IDR", "HBL", "MBL", "CY"}


def expense_key(description) -> Optional[str]:
    """費用名稱 -> 比對用 key：轉大寫、去雜訊與標點、展開縮寫、去掉重複片語
    (例如 'THC-TERMINAL HANDLING CHARGE' 展開後會變成兩次 TERMINAL
    HANDLING CHARGE)。"""
    text = clean_text(description)
    if not text:
        return None
    # 越南文等帶聲調的字母先去掉聲調 (PHỤ PHÍ XĂNG DẦU -> PHU PHI XANG DAU)，
    # 不然下面去標點時會把整個字母刪掉
    text = text.replace("Đ", "D").replace("đ", "d")
    text = "".join(ch for ch in unicodedata.normalize("NFKD", text)
                   if not unicodedata.combining(ch))
    text = text.upper().replace("&", " AND ")
    for noise in _NOISE_RES:
        text = noise.sub(" ", text)
    text = re.sub(r"[^A-Z0-9/ ]+", " ", text)
    text = text.replace("/", " ")
    text = re.sub(r"\s+", " ", text).strip()
    for pattern, repl in _EXPANSION_RES:
        text = pattern.sub(repl, text)
    text = re.sub(r"\s+", " ", text).strip()
    # 連續重複的片語只留一次 (最長 6 個字)
    for n in range(6, 0, -1):
        text = re.sub(rf"\b((?:\w+ ){{{n - 1}}}\w+)(?: \1\b)+", r"\1", text)
    return text or None


def _title(key: str) -> str:
    return " ".join(w if w in _ACRONYMS else w.capitalize() for w in key.split())


def normalize_expense(description) -> Tuple[str, str]:
    """回傳 (標準費用項目, 費用大類)。"""
    key = expense_key(description)
    if not key:
        return UNKNOWN_EXPENSE, OTHER_CATEGORY
    item = EXPENSE_ALIASES.get(key) or _title(key)
    for name, pattern in _CATEGORY_RES:
        if pattern.search(key):
            return item, name
    return item, OTHER_CATEGORY


# ============================================================
# 幣別
# ============================================================

# 帳單、供應商、國家資料夾都判斷不出幣別時不猜，幣別留空 (None)，
# 由 invoice-expense-db 列出清單詢問使用者，再補進下面的規則。
# (帳單上的美金只出現在費用名稱裡，例如 'FREIGHT CHARGE (USD 23.54)'，
#  金額欄已經是當地幣別，所以不從費用名稱判斷幣別)
DEFAULT_CURRENCY: Optional[str] = None
# 個別供應商的金額幣別跟預設不同時加在這裡 (key：SUPPLIER 欄位轉大寫)
SUPPLIER_CURRENCY: dict = {
    # 越南供應商 (Excel 對帳表，金額為越南盾)
    "CÔNG TY TNHH GIAO NHẬN THƯƠNG MẠI RỒNG Á CHÂU": "VND",
    "CÔNG TY TNHH DỊCH VỤ GIAO NHẬN VẬN TẢI VÀ THƯƠNG MẠI CÔNG THÀNH": "VND",
}

# INVOICE NO 不是供應商自己的帳單號碼 (例如寶成的發票號)、帳單也沒有發票日期的
# 供應商 (key：SUPPLIER 欄位轉大寫；使用者 2026-10-28 指定)：
#   - 分析日期：發票日期空白時改用 on board date (轉檔結果仍是 N/A)。
#   - 重複判斷：同一個發票號碼會出現在不同週的帳單、費用項目不同，所以不能
#     用 (供應商, 發票號碼) 判斷；改成整列 (發票號碼+費用名稱+金額+提單號碼) 相同才算。
SUPPLIER_NO_OWN_INVOICE = {
    "CÔNG TY TNHH GIAO NHẬN THƯƠNG MẠI RỒNG Á CHÂU",
}


def is_no_own_invoice(supplier) -> bool:
    """用開頭比對：轉檔結果的名稱後面可能有英文備註 "(… English …)"。"""
    sup = clean_text(supplier)
    return bool(sup) and sup.upper().startswith(tuple(SUPPLIER_NO_OWN_INVOICE))


# 幣別判斷順序 (使用者 2026-10-25 指定)：
#   1. 帳單上的幣別 (金額文字裡的幣別代碼/符號，_CURRENCY_RES)
#   2. 供應商 (SUPPLIER_CURRENCY 指定名稱 -> 公司名稱的國別特徵)
#   3. 國家資料夾 (供應商轉檔結果\<國家>\...，COUNTRY_FOLDER_CURRENCY)
#   4. 都判斷不出來：留空 (None)，詢問使用者
# 目前帳單只有中國、東南亞、印度、台灣；新增其他國家時再補規則。
COUNTRY_CURRENCY = {
    "CN": "CNY", "TW": "TWD", "IN": "INR", "ID": "IDR", "VN": "VND", "TH": "THB",
    "MY": "MYR", "SG": "SGD", "PH": "PHP", "KH": "KHR", "LA": "LAK", "MM": "MMK",
    "BN": "BND", "TL": "USD",  # 東帝汶通用美元
}
# 公司名稱的國別特徵 (依序比對，先符合的為準)
_COUNTRY_NAME_RES = (
    # 越南：CÔNG TY / CTY (公司)、TNHH (有限責任)、越南文字母
    ("VN", re.compile(r"\bC[ÔO]NG\s*TY\b|\bCTY\b|TNHH|[ĂÂĐÊÔƠƯăâđêôơư]|VIET\s*NAM",
                      re.IGNORECASE)),
    # 泰國：泰文字母
    ("TH", re.compile(r"[฀-๿]")),
    # 印尼：PT / CV 開頭、名稱含 INDONESIA
    ("ID", re.compile(r"^(PT|CV)\b\.?|INDONESIA", re.IGNORECASE)),
    # 馬來西亞：SDN. BHD.
    ("MY", re.compile(r"\bSDN\.?\s*BHD\b", re.IGNORECASE)),
    # 新加坡：PTE. LTD.
    ("SG", re.compile(r"\bPTE\.?\s*LTD\b", re.IGNORECASE)),
    # 印度：PVT. LTD. / PRIVATE LIMITED
    ("IN", re.compile(r"\bPVT\.?\s*LTD\b|\bPRIVATE\s+LIMITED\b", re.IGNORECASE)),
)
# 國家資料夾名稱 -> 幣別 (路徑中任一層資料夾名稱完全相同才算)
COUNTRY_FOLDER_CURRENCY = {
    "中國": "CNY", "中国": "CNY", "大陸": "CNY", "台灣": "TWD", "臺灣": "TWD",
    "印度": "INR", "印尼": "IDR", "越南": "VND", "泰國": "THB", "馬來西亞": "MYR",
    "新加坡": "SGD", "菲律賓": "PHP", "柬埔寨": "KHR", "寮國": "LAK", "緬甸": "MMK",
    "汶萊": "BND", "東帝汶": "USD",
}


def folder_currency(path) -> Optional[str]:
    """檔案路徑 -> 國家資料夾的幣別 (離檔案最近的國家資料夾為準)；沒有回傳 None。"""
    if not path:
        return None
    for part in reversed(re.split(r"[\\/]", str(path))):
        if part.strip() in COUNTRY_FOLDER_CURRENCY:
            return COUNTRY_FOLDER_CURRENCY[part.strip()]
    return None


def supplier_country(supplier) -> Optional[str]:
    """供應商名稱 -> 國別代碼 (COUNTRY_CURRENCY 的 key)；看不出來回傳 None。"""
    text = clean_text(supplier)
    if not text:
        return None
    for country, pattern in _COUNTRY_NAME_RES:
        if pattern.search(text):
            return country
    return None


# 帳單上的幣別代碼/符號 (依序比對；USD 要在 SGD/BND 前面，US$ 才不會被當成 S$；
# TWD 要在 CNY 前面，「新台幣…元」才不會被當成人民幣)
_CURRENCY_RES = (
    ("USD", re.compile(r"\bUSD\b|US\$|^\s*\$")),
    ("IDR", re.compile(r"\bIDR\b|\bRP\.?", re.IGNORECASE)),
    ("VND", re.compile(r"\bVND\b|\bVNĐ\b|₫|\bĐỒNG\b", re.IGNORECASE)),
    ("SGD", re.compile(r"\bSGD\b|S\$")),
    ("BND", re.compile(r"\bBND\b|B\$")),
    ("MYR", re.compile(r"\bMYR\b|\bRM\s*\d")),
    ("THB", re.compile(r"\bTHB\b|฿|\bBAHT\b", re.IGNORECASE)),
    ("PHP", re.compile(r"\bPHP\b|₱")),
    ("KHR", re.compile(r"\bKHR\b|៛|\bRIEL\b", re.IGNORECASE)),
    ("LAK", re.compile(r"\bLAK\b|₭|\bKIP\b", re.IGNORECASE)),
    ("MMK", re.compile(r"\bMMK\b|\bKYAT\b", re.IGNORECASE)),
    ("INR", re.compile(r"\bINR\b|₹|\bRS\.?\s*\d", re.IGNORECASE)),
    ("EUR", re.compile(r"\bEUR\b|€")),
    ("TWD", re.compile(r"\bTWD\b|\bNTD?\$?|新[台臺]幣")),
    ("CNY", re.compile(r"\bCNY\b|\bRMB\b|人民[币幣]|[¥￥]")),
)


def detect_currency(amount_raw, supplier,
                    default: Optional[str] = DEFAULT_CURRENCY) -> Optional[str]:
    """金額幣別：帳單上的幣別 (金額文字，例如 'USD 1,234'、'Rp 50.000') 優先，
    其次供應商 (SUPPLIER_CURRENCY -> 名稱國別 COUNTRY_CURRENCY)，最後 default
    (呼叫端傳國家資料夾的幣別，見 folder_currency())；都沒有回傳 None。"""
    if isinstance(amount_raw, str):
        for code, pattern in _CURRENCY_RES:
            if pattern.search(amount_raw):
                return code
    sup = clean_text(supplier)
    if sup:
        # 轉檔結果會在非英文名稱後面加英文括號備註，比對時去掉
        sup = re.sub(r"\s*\([^()]*\)\s*$", "", sup)
    if sup and sup.upper() in SUPPLIER_CURRENCY:
        return SUPPLIER_CURRENCY[sup.upper()]
    country = supplier_country(sup)
    if country in COUNTRY_CURRENCY:
        return COUNTRY_CURRENCY[country]
    return default


# ============================================================
# 港口
# ============================================================

def normalize_port(value) -> Optional[str]:
    """港口名稱：統一 Title Case (BYHALIA -> Byhalia、LOS ANGELES ->
    Los Angeles)；只有一個三碼以下的代碼 (例如 CGK) 時保持大寫。"""
    text = clean_text(value)
    if not text:
        return None
    # 擷取時偶爾連後面的船名航次一起抓到 ("Jakarta Vessel/voyage:XIN YAN TAI/ 272N")
    text = re.split(r"\s+VESSEL\b", text, maxsplit=1, flags=re.IGNORECASE)[0].strip()
    if len(text) <= 3 and text.isalpha():
        return text.upper()
    return " ".join(w[:1].upper() + w[1:].lower() for w in text.split(" "))
