# -*- coding: utf-8 -*-
"""
expense_normalize.py
====================
費用分析系統的資料清洗規則 (純函式，不碰檔案/資料庫，方便單獨測試)：

  - parse_date()       ：各種日期寫法 (31.Aug.2026、2026-08-31、31/08/2026、
                         印尼文月份、Excel 日期序號...) 轉成 datetime.date
  - parse_amount()     ：金額/稅額 (含 Rp、千分位逗號或點、括號負數) 轉成 float
  - parse_float()      ：材積等一般小數 (點一律當小數點)
  - normalize_expense():費用名稱正規化，回傳 (標準費用項目, 費用大類)
  - normalize_port()   ：港口名稱統一大小寫與空白
  - clean_text()       ：一般文字欄位去空白，N/A / - / 空字串一律變成 None

要新增「同性質費用合併」的對照，改 EXPENSE_ALIASES 即可；要調整費用大類，
改 EXPENSE_CATEGORIES (由上往下比對，先符合的先用)。
"""

import re
from datetime import date, datetime, timedelta
from typing import Optional, Tuple

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
}

# 費用大類：(大類名稱, 關鍵字 regex)，由上往下比對，先符合的先用
EXPENSE_CATEGORIES = (
    ("碼頭處理費 (THC)", r"TERMINAL HANDLING"),
    ("VGM 費用", r"\bVGM\b"),
    ("CFS / 倉儲", r"\bCFS\b|WAREHOUSE|GUDANG|STORAGE|STUFFING|UNSTUFFING"),
    ("文件 / 提單費", r"DOCUMENTATION|\bFCR\b|\bCBL\b|\bBL\b|BILL OF LADING|\bAWB\b|\bEDI\b|"
                     r"ISSUANCE|\bSEAL\b|SHIPPING INSTRUCTION"),
    ("運費", r"FREIGHT|\bBAF\b|\bCAF\b|BUNKER|SURCHARGE"),
    ("報關 / 檢驗", r"CUSTOMS|CLEARANCE|DECLARATION|\bPEB\b|\bPIB\b|INSPECTION|QUARANTINE|"
                   r"FUMIGATION"),
    ("拖車 / 陸運 / 滯櫃", r"TRUCKING|HAULAGE|DELIVERY|TRANSPORT|PICK ?UP|LIFT ON|LIFT OFF|"
                         r"DETENTION|DEMURRAGE"),
    ("手續 / 管理 / 行政", r"HANDLING|AGENCY|ADMINISTRATION|SERVICE|COMMISSION|MANAGEMENT|"
                         r"BOOKING|COMPLIAN|LATE PAYMENT"),
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
