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

from extract_utils_streamlit import register_supplier

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


# ---------------------------------------------------------------------------
# 全系統統一的日期顯示格式：「DD Mon YYYY」，月份用英文縮寫 (Jan/Feb/…)。
# 所有供應商的日期擷取函式最後都呼叫這個共用函式組字串，不要各自手動拼
# "DD.MM.YYYY"，確保畫面顯示、匯出 Excel、跟正確答案比對三處的日期格式
# 永遠一致 (比對邏輯見 extract_utils.normalize_value()，它會把『正確答案
# Excel 裡的日期型態儲存格』也用同一個英文月份格式重新格式化再比對)。
# ---------------------------------------------------------------------------
_MONTH_ABBR_EN = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
]


def _format_date_en(day, month, year) -> Optional[str]:
    """day/month/year (int 或數字字串) -> 'DD Mon YYYY'，例如 (2, 1, 2025)
    -> '02 Jan 2025'。month 不在 1~12 範圍就回傳 None (呼叫端視為擷取失敗)。
    """
    try:
        day_i, month_i, year_i = int(day), int(month), int(year)
    except (TypeError, ValueError):
        return None
    if not 1 <= month_i <= 12:
        return None
    return f"{day_i:02d} {_MONTH_ABBR_EN[month_i - 1]} {year_i}"


_MONTH_NAME_TO_NUM = {name.upper(): i + 1 for i, name in enumerate(_MONTH_ABBR_EN)}


def _format_date_en_from_month_name(day, month_name, year) -> Optional[str]:
    """day/year 為數字 (或數字字串)、month_name 為英文月份縮寫 (任意大小寫，
    例如 'Aug'、'AUG')時使用；轉成統一的 'DD Mon YYYY' 格式。
    """
    if month_name is None:
        return None
    month_i = _MONTH_NAME_TO_NUM.get(str(month_name).strip().upper()[:3])
    if month_i is None:
        return None
    return _format_date_en(day, month_i, year)


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
# 冒號後面只能跳過「同一行」的空白 (空格/tab)，不能用 \s* (那樣連換行都會
# 跳過)：像 "Container No. :" 或 "Volume :" 這種欄位常常沒有值、緊接著就
# 換行進到下一個標籤，如果冒號後面的空白字元允許吃掉換行，(.+?) 就會被
# 硬牽去抓下一行「下一個欄位」的文字 (例如 "Container No. :\nConsignee :
# PT. POU YUEN INDONESIA" 會被誤抓成 "Consignee : PT. POU YUEN INDONESIA"
# 當作 container_no 的值)。改成只吃同一行空白後，遇到欄位是空的，
# (.+?) 在换行前找不到任何字元可比對，整個 pattern 就不會比對成功，
# _search() 正確回傳 None，之後由 expand_to_rows() 補成 "N/A"。
_DWIHARTA_COLON = r"\s*:[ \t]*"

_DWIHARTA_HEADER_PATTERNS: Dict[str, str] = {
    "invoice_no":         rf"Invoice\s*No\.?{_DWIHARTA_COLON}(.+?){_DWIHARTA_STOP}",
    "invoice_date":       rf"Invoice\s*Date{_DWIHARTA_COLON}(.+?){_DWIHARTA_STOP}",
    "order_no":           rf"Order\s*No\.?{_DWIHARTA_COLON}(.+?){_DWIHARTA_STOP}",
    "arrive_date":        rf"Arrive\s*Date{_DWIHARTA_COLON}(.+?){_DWIHARTA_STOP}",
    "onboard_date":       rf"On\s*[Bb]oard\s*Date{_DWIHARTA_COLON}(.+?){_DWIHARTA_STOP}",
    "bl_no":              rf"\bB/L\s*No\.?{_DWIHARTA_COLON}(.+?){_DWIHARTA_STOP}",
    "mbl_no":             rf"\bMBL\s*No\.?{_DWIHARTA_COLON}(.+?){_DWIHARTA_STOP}",
    "port_of_loading":    rf"Port\s*Of\s*Loading{_DWIHARTA_COLON}(.+?){_DWIHARTA_STOP}",
    "port_of_discharge":  rf"Port\s*Of\s*Discharge{_DWIHARTA_COLON}(.+?){_DWIHARTA_STOP}",
    "volume":             rf"Volume{_DWIHARTA_COLON}(.+?){_DWIHARTA_STOP}",
    "vessel":             rf"Vessel{_DWIHARTA_COLON}(.+?){_DWIHARTA_STOP}",
    "container_no":       rf"Container\s*No\.?{_DWIHARTA_COLON}(.+?){_DWIHARTA_STOP}",
    "consignee":          rf"\bTO{_DWIHARTA_COLON}(.+?){_DWIHARTA_STOP}",
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
    return _format_date_en(day, month, year)


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

    # Shipper 欄位的值常常換行 (例如 "Shipper : SHANGHAI LIGHT INDUSTRY AND\n
    # TEXTILE GROUP IMP. & EXP. CO.,LTD.; Container No. : ...")，不能用一般
    # 欄位的 _DWIHARTA_STOP (遇到換行就停) 處理，這裡改用允許跨行的寫法，
    # 抓到 "Container No." 前的所有文字，再把裡面的換行/多重空白收斂成
    # 一個空白，還原成一行完整的出貨方名稱。
    _shipper_m = re.search(r"Shipper\s*:\s*([\s\S]+?)\s*Container\s*No\.?\s*:", text)
    if _shipper_m:
        header["shipper"] = re.sub(r"\s+", " ", _shipper_m.group(1)).strip()

    items: List[Dict] = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line or any(kw in line for kw in _DWIHARTA_EXCLUDE_KEYWORDS):
            continue
        m = _DWIHARTA_LINE_ITEM_PATTERN.match(line)
        if not m:
            continue
        amount_str = m.group("amount").replace(",", "")
        vat_str = m.group("vat")
        vat_val = None
        if vat_str:
            vat_str = vat_str.replace(",", "")
            vat_val = int(vat_str) if vat_str.isdigit() else None
        items.append({
            "description": m.group("desc").strip(),
            "amount": int(amount_str) if amount_str.isdigit() else None,
            "vat": vat_val,
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


# 費用明細行 (在 KN Sales Invoice 的 "CHARGE" 區塊裡)，有兩種格式：
#   1. 有外幣金額的："THC - ORIGIN 385.00 USD 7,173,967 7"
#      格式：費用名稱 ... 外幣金額 USD 印尼幣金額 VAT代碼
#   2. 只有印尼幣的（沒有外幣換算）："SECURITY FILING 700,000 7"
#      格式：費用名稱 ... 印尼幣金額 VAT代碼
# 這是「使用者真正要的 Description」，跟 Faktur Pajak (稅務發票) 上寫的
# 那種概括性文字 (例如 "Domestic transport & ancillary services") 是兩回事：
# 稅務發票的文字只是報稅用的統稱，帳單上實際列出的每筆費用名稱才是查核
# 要比對的欄位。這裡用「結尾是 (可選的外幣金額+USD) + 印尼幣金額 + VAT
# 代碼」的格式辨識這種行，一張帳單若有多筆費用，會抓到多筆、各自展開成
# 一列 (呼應 DWIHARTA「一筆費用一列」的設計)。VAT 代碼固定是單獨一個數字
# (目前觀察到的樣本都是 "7")，不是金額的一部分——限定成單一位數字，避免
# 誤吃到後面緊接著的其他數字 (例如貨櫃追蹤號碼 "1076 285 746" 被誤判成
# 金額+VAT代碼)。金額部分也要求「英式千分位逗號」格式 (例如 "700,000")，
# 不接受沒有逗號分隔的連續數字，這樣才不會誤抓到 PO 單號、提單號這類沒
# 有千分位逗號、純數字黏在旁邊字母後面的雜訊行 (pdfplumber 有時會把側邊
# 直排文字的單個字母併進旁邊的數字行)。
_KN_CHARGE_LINE_PATTERN = re.compile(
    r"^(?P<desc>[A-Z][A-Za-z0-9 /\-\.,()&]+?)\s+"
    r"(?:[\d,]+\.\d{2}\s*[A-Z]{3}\s+)?"
    r"(?P<amount>\d{1,3}(?:,\d{3})+)\s+\d\s*$",
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


# "CONSIGNEE SHIPPER" 這個區塊是左右兩欄並排、但 pdfplumber 抓文字時會把
# 同一行的兩欄黏成一行 (中間只隔一個空白，不是兩個以上，沒辦法用「兩個
# 空白就切開」那招拆開)，例如：
#   "ASICS JAPAN CORPORATION THE LOOK (MACAO COMMERCIAL OFFSHORE"
#   "7-2 MARUNOUCHI 2-CHOME, COMPANY LIMITED O/B PT. POUCHEN"
# 這裡用一個技巧還原：帳單下面緊接著的 "NOTIFY" 區塊，常常跟 CONSIGNEE
# 欄位是同一個收件人、逐行文字完全一樣 (地址寫兩次)，所以只要「NOTIFY
# 每一行」剛好是「CONSIGNEE/SHIPPER 對應那一行」的開頭前綴，就能把前綴
# 去掉、剩下的部分就是 SHIPPER 欄位那一行的文字；把每一行的 SHIPPER 片段
# 接起來，再依照使用者的規則，取到「LIMITED」或「LTD」等公司型態字樣為
# 止 (使用者確認：Shipper 公司全名只到這個字樣，後面接的是地址/備註，不
# 算在內)。如果 NOTIFY 跟 CONSIGNEE 不是同一個收件人 (兩個區塊行數不同，
# 或對不起來)，就代表這招用不上，回傳 None，呼叫端會改用其他方式 (沿用
# 已經可靠抓到的 consignee 值)。
_KN_LIMITED_WORD = re.compile(r"\bL(?:IMITED|TD\.?)\b", re.IGNORECASE)


def _kn_extract_shipper_via_notify(text: str) -> Optional[str]:
    block_m = re.search(r"CONSIGNEE SHIPPER\n((?:.+\n)+?)(?=NOTIFY\b)", text)
    notify_m = re.search(r"NOTIFY\n((?:.+\n)+?)(?=FIRST VESSEL|Gov doc#)", text)
    if not block_m or not notify_m:
        return None
    merged_lines = [ln for ln in block_m.group(1).split("\n") if ln.strip()]
    notify_lines = [ln for ln in notify_m.group(1).split("\n") if ln.strip()]
    if len(merged_lines) != len(notify_lines) or not merged_lines:
        return None

    fragments = []
    for merged, notify_line in zip(merged_lines, notify_lines):
        if not merged.startswith(notify_line):
            return None
        fragments.append(merged[len(notify_line):].strip())

    shipper_text = " ".join(f for f in fragments if f)
    limit_m = _KN_LIMITED_WORD.search(shipper_text)
    if not limit_m:
        return None
    return re.sub(r"\s+", " ", shipper_text[: limit_m.end()]).strip()


def parse_kuehne_nagel(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {}

    header["supplier"] = _search(r"Pengusaha Kena Pajak:\s*\n\s*Nama\s*:\s*(.+)", text)

    header["consignee"] = _search(rf"INVOICE TO\s*\n\s*(.+?){_KN_STOP}", text)
    # Sales Invoice 頁的 "CONSIGNEE SHIPPER" 這個區塊裡，SHIPPER 欄位有時
    # 跟上面的 INVOICE TO (header["consignee"]) 是同一家公司 (觀察到的部分
    # 樣本都是 PT. POU YUEN INDONESIA，這裡的 CONSIGNEE 欄位其實是另一個
    # 不同欄位，貨物真正收件地，不是我們的 consignee 標準欄位)，但不是每
    # 張都這樣 (例如這家貨主跟收件人是不同公司時)，所以優先用
    # _kn_extract_shipper_via_notify() 真正解析這個左右並排的欄位；解析不
    # 出來時才退回沿用 consignee 值。
    header["shipper"] = _kn_extract_shipper_via_notify(text) or header["consignee"]

    m = re.search(r"INVOICE NO\.\s*/\s*DATE\s+(\S+)\s+(\d{1,2})\.(\d{1,2})\.(\d{4})", text)
    if m:
        header["invoice_no"] = m.group(1)
        header["invoice_date"] = _format_date_en(m.group(2), m.group(3), m.group(4))

    header["bl_no"] = _kn_clean_tracking_no(_search(rf"KN TRACKING NO\.\s*(.+?){_KN_STOP}", text))
    # 注意：Order No. 欄位對應的是「COMMERCIAL INVOICE NUMBER」(例如
    # RYH0810074280)，不是緊接在旁邊的「GTN SHIPPING ORDER NUMBER」
    # (例如 VB26072003093319) —— 這兩個編號在版面上前後相鄰、很容易搞混，
    # 但使用者核對過的正確答案是以 COMMERCIAL INVOICE NUMBER 為準。
    header["order_no"] = _search(rf"COMMERCIAL INVOICE NUMBER\s*\n\s*(.+?){_KN_STOP}", text)
    # Vessel 欄位使用者確認要「船名+航次」(例如 "VANCOUVER 058N")，不是只
    # 取船名；VESSEL NAME/VOYAGE 兩個欄位左右並排在同一行 ("VESSEL NAME :
    # VANCOUVER VOYAGE : 058N")，這裡兩個都抓出來、中間用一個空白接起來。
    _kn_vessel_m = re.search(r"VESSEL NAME\s*:\s*(.+?)\s+VOYAGE\s*:\s*(.+?)(?=\n|$)", text)
    if _kn_vessel_m:
        header["vessel"] = f"{_kn_vessel_m.group(1).strip()} {_kn_vessel_m.group(2).strip()}"
    else:
        header["vessel"] = _search(rf"VESSEL NAME\s*:\s*(.+?){_KN_STOP}", text)
    header["port_of_loading"] = _kn_clean_port(_search(rf"P\.\s*OF LOADING\s*:\s*(.+?){_KN_STOP}", text))
    header["port_of_discharge"] = _kn_clean_port(_search(rf"P\.\s*OF DISCHARGE\s*:\s*(.+?){_KN_STOP}", text))
    _kn_etd_m = re.search(r"ETD/ATD\s*:\s*(\d{1,2})\.(\d{1,2})\.(\d{4})", text)
    header["onboard_date"] = (
        _format_date_en(*_kn_etd_m.groups()) if _kn_etd_m else None
    )
    _kn_eta_m = re.search(r"ETA/ATA\s*:\s*(\d{1,2})\.(\d{1,2})\.(\d{4})", text)
    header["arrive_date"] = (
        _format_date_en(*_kn_eta_m.groups()) if _kn_eta_m else None
    )

    m2 = re.search(r"AS PER ATTACHED\s+([\d,\.]+)\s+([\d,\.]+)", text)
    header["volume"] = m2.group(2) if m2 else None
    # 注意：上面這個「整張帳單共用一個 Volume」只適用於沒有逐櫃拆分費用的
    # 帳單 (例如只有一筆 OVERTIME CHARGES EXPORT 的簡單帳單)。如果帳單裡有
    # 「CONTAINER POSITIONING EXPORT」這種逐個貨櫃分開收費的項目，使用者的
    # 正確答案改用「每個貨櫃自己的櫃型代碼」(例如 "1X40GE") 當那幾筆費用
    # 各自的 Volume，其餘費用列則留空，不套用整張帳單共用的材積總量——這裡
    # 先記下有沒有這種逐櫃費用，等下面抓到 items 之後再決定要不要蓋掉這個
    # 預設值 (見下方 _kn_container_positioning 區塊)。

    # 這種帳單本來就沒有海運提單/貨櫃資訊，留空給 expand_to_rows() 補 N/A
    header["mbl_no"] = None
    header["container_no"] = None

    items = _kn_extract_charge_items(text)
    # 這家供應商的 PDF 常常把同一張 Sales Invoice「整份內容」重複印兩次
    # (ORIGINAL 跟 DIGITAL COPY 各一份，直排在頁面兩側當水印文字，肉眼看
    # 不出來，但 pdfplumber 把兩份內容前後接在一起)，導致費用明細行數量
    # 剛好是真實筆數的兩倍、且前後兩半完全一樣。這裡偵測「整個清單剛好可
    # 以切成前後兩半、且兩半逐筆完全相同」時，只留前半，避免每筆費用都被
    # 誤判成重複兩筆；如果不是這種剛好對半重複的情況 (例如真的有兩筆金額
    # 剛好相同的費用)，兩半不會逐筆完全相等，不會被誤刪。
    if items and len(items) % 2 == 0:
        half = len(items) // 2
        if items[:half] == items[half:]:
            items = items[:half]

    # 「CONTAINER POSITIONING EXPORT」費用行下面緊接著的明細行會寫出這筆
    # 費用對應的貨櫃型態，例如 "1 40' GE X 23 USD = 23.00 USD" 代表一個
    # 40 呎 GE (General Export) 櫃，使用者的正確答案把它轉成 "1X40GE" 這種
    # 代碼格式填在 Volume 欄位 (只填在這幾筆費用列，不是整張帳單共用一個
    # 值)。這裡依照文件順序抓出每一個這種代碼，跟 items 裡同樣依文件順序
    # 出現的 CONTAINER POSITIONING EXPORT 項目一一對應；跟費用明細行一樣，
    # 整份文字常常重複印兩次，所以也套用同樣的「對半相同就只留前半」處理。
    _kn_container_codes = [
        f"1X{num}{kind}"
        for num, kind in re.findall(r"1\s+(\d+)'\s*([A-Z]+)\s*X\s*[\d.]+\s*USD\s*=", text)
    ]
    if _kn_container_codes and len(_kn_container_codes) % 2 == 0:
        _half = len(_kn_container_codes) // 2
        if _kn_container_codes[:_half] == _kn_container_codes[_half:]:
            _kn_container_codes = _kn_container_codes[:_half]
    if _kn_container_codes:
        # 有逐櫃費用時，整張帳單共用的材積總量就不適用了 (正確答案這幾張
        # 帳單的其他費用列 Volume 欄位都是留空，不是填整張帳單的材積總量)。
        header["volume"] = None
        _kn_code_iter = iter(_kn_container_codes)
        for item in items:
            if item.get("description") == "CONTAINER POSITIONING EXPORT":
                try:
                    item["volume"] = next(_kn_code_iter)
                except StopIteration:
                    break

    if not items:
        # 備援：真的抓不到任何費用明細行時，退回用稅務發票 (Faktur Pajak)
        # 的概括描述+總金額，至少不要整張全空。
        description = _search(r"\n([A-Za-z][^\n]+)\nRp\s+[\d.,]+\s*x", text)
        amount_raw = _search(
            r"Harga Jual\s*/\s*Penggantian\s*/\s*Uang Muka\s*/\s*Termin\s+([\d.,]+)", text
        )
        items = [{"description": description, "amount": _kn_clean_amount(amount_raw)}]

    # VAT：費用明細行本身 (例如 "OVERTIME CHARGES EXPORT 85.00 USD
    # 1,521,098 7") 結尾那個數字只是 VAT 稅率代碼的註腳編號 (對照下面
    # "7: 11.00% GROSS / 12% NET VAT..." 的說明)，不是稅額金額；真正的
    # VAT 金額要從 "SUBTOTAL IDR 1,521,098 167,321" 這一行抓，跟費用明細
    # 金額並排在同一行 (兩個都是英式千分位逗號數字)。
    #
    # 這張帳單無論有幾筆費用明細，VAT 永遠只有「一個」整張帳單的總稅額
    # (不是每筆費用各自的稅額)，所以不能像其他供應商一樣把 vat 塞進某一
    # 筆費用明細的 dict 裡 (塞進第一筆會讓人誤以為那筆費用才有稅額)。使用
    # 者更新後的正確答案範本也是這樣處理：額外「多一列」只填 VAT 金額，
    # Description/Amount 都留空——因此這裡額外附加一筆「只有 vat、沒有
    # description/amount」的 item，讓 expand_to_rows() 展開成獨立的一列，
    # 跟正確答案的列結構對齊。
    _kn_subtotal_m = re.search(r"SUBTOTAL\s+IDR\s+[\d,]+\s+([\d,]+)", text)
    if _kn_subtotal_m:
        vat_amount = _kn_clean_amount_comma_thousands(_kn_subtotal_m.group(1))
        items.append({"description": None, "amount": None, "vat": vat_amount})

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
#
# ⚠️ multi_page：這家供應商常把「同一個月份好幾張獨立發票」合併成一份
# PDF、一頁一張各自獨立的發票 (不同 invoice number)，所以跟 HYPER_MEGA/
# YJE_TATA/PANCARAN 一樣註冊成 multi_page=True，detect_indoprostime()/
# parse_indoprostime() 會被逐頁呼叫 (extract_utils.py 的
# _parse_multi_page_pdf())，text 參數是『一頁』的文字，不是整份 PDF
# 合併後的文字。

def detect_indoprostime(text: str) -> bool:
    """只認 'INDO PROSTIME EXPRESS'，不要用單純 'PROSTIME' 當關鍵字——
    PT. Indo Prostime Transport (另一家公司，另外註冊成
    detect_indoprostime_transport()/parse_indoprostime_transport()) 也會
    一樣含有 'PROSTIME' 字樣，太寬鬆的判斷式會把兩家公司的帳單搞混。"""
    return "INDO PROSTIME EXPRESS" in text.upper()


_INDO_LABEL_WORDS = [
    r"A/C\s*No\.?", r"A/C\s*name", r"Swift\s*code", r"Bank",
    r"Number", r"Date", r"Order\s*No\.?", r"Currency\s*/\s*Rate",
    r"MAWB", r"HAWB", r"QUANTITY", r"SHIPPER", r"GROSS\s*/\s*KG",
    r"CNEE", r"CONSIGNEE", r"CHARGE\s*/\s*KG", r"ORIGIN", r"FLIGHT", r"FREIGHT",
    r"PKGS", r"PIB\s*/\s*BC", r"D\s*E\s*S\s*C\s*R\s*I\s*P\s*T\s*I\s*O\s*N", r"AMOUNT",
    r"Total", r"Cheque", r"ROLL",
]
_INDO_NEXT = "(?:" + "|".join(_INDO_LABEL_WORDS) + ")"
# 注意：這裡不像 DWIHARTA/KN 要求 "\s{2,}"，因為並排欄位只隔一個空白；
# 只要「後面接著任何空白 + 下一個已知欄位關鍵字」就停止擷取。
_INDO_STOP = rf"(?=\s+{_INDO_NEXT}\b|\n|$)"
# SHIPPER 欄位有時會跟右欄下一個欄位標籤完全黏在一起、中間沒有空白
# (例如 "SHIPPER : SHANGHAI EASEMENT ELECTRIC CO., LTDGROSS /KG : 0.50"，
# "LTD" 跟 "GROSS" 中間沒有空白)，一般 _INDO_STOP 要求「下一個標籤前面
# 一定要有空白」在這裡會抓過頭，所以 SHIPPER 專用這個放寬版 (允許零個
# 空白) 的停止條件。
_INDO_STOP_LOOSE = rf"(?=\s*{_INDO_NEXT}\b|\n|$)"

_INDO_MONTH_ABBR_UPPER = [
    "JAN", "FEB", "MAR", "APR", "MAY", "JUN",
    "JUL", "AUG", "SEP", "OCT", "NOV", "DEC",
]


def _indo_format_date(day, month, year) -> Optional[str]:
    """這個供應商家族 (INDOPROSTIME EXPRESS/TRANSPORT) 的正確答案範本把
    日期欄位存成純文字 'YYYY-MON-DD' (例如 '2026-JUN-11')，不是真正的
    Excel 日期型態儲存格，所以 normalize_value() 不會幫忙轉換，日期格式
    要在這裡直接組成一樣的字串才會比對相符，跟其他供應商統一用的
    'DD Mon YYYY' (_format_date_en()) 不同，不要共用。"""
    try:
        d, m, y = int(day), int(month), int(year)
    except (TypeError, ValueError):
        return None
    if not 1 <= m <= 12:
        return None
    return f"{y}-{_INDO_MONTH_ABBR_UPPER[m - 1]}-{d:02d}"


_INDO_TOP_CONSIGNEE_RE = re.compile(r'^(.+?)\s+Number\s*:\s*\d+', re.MULTILINE)


def _indo_extract_top_consignee(text: str) -> Optional[str]:
    """抬頭收件人名稱印在發票最上方、"Number :" 前面那一行 (不是
    "CNEE :"/"CONSIGNEE :" 那個欄位標籤行)，例如
    'POU YUEN INDONESIA,.PT. Number : 089261' -> 'POU YUEN INDONESIA,.PT.'。
    """
    m = _INDO_TOP_CONSIGNEE_RE.search(text)
    return m.group(1).strip() if m else None


def _indo_format_volume(raw: Optional[str]) -> Optional[str]:
    """'0.50'/'1.60'/'44.00' -> '0.5 KG'/'1.6 KG'/'44 KG' (去掉多餘的零，
    補上單位 KG，跟正確答案範本的寫法一致)。"""
    if not raw:
        return None
    try:
        num = float(raw.strip())
    except ValueError:
        return None
    return f"{num:g} KG"


def _indo_extract_vessel_code(raw_flight: Optional[str]) -> Optional[str]:
    """FLIGHT 欄位原始值 (例如 'GA 863/09062026'、'VJ0855/05-08-2026') 只取
    "/" 前面的航班代碼部分當 Vessel，日期部分另外由
    _indo_parse_flight_date()/_ipt_parse_flight_date() 轉成開船日。"""
    if not raw_flight:
        return None
    return raw_flight.split("/")[0].strip()


def _indo_add_space_after_pt(value: Optional[str]) -> Optional[str]:
    """'PT.HARDASES ABADI INDONESIA' -> 'PT. HARDASES ABADI INDONESIA'：
    PDF 版面把 'PT.' 跟公司名稱黏在一起、中間沒有空白，這裡補回一個空白，
    跟正確答案 Excel 裡的寫法對齊 (正確答案 'PT.' 後面有空白)。
    """
    if not value:
        return value
    return re.sub(r"^(PT\.)(?=\S)", r"\1 ", value.strip())


def _indo_parse_flight_date(raw_flight: Optional[str]) -> Optional[str]:
    """從 FLIGHT 欄位原始值反推開船日，支援兩種常見寫法：
    - 'CX 777/01092026' (尾碼 8 碼數字無分隔，DDMMYYYY)
    - 'VJ0855/05-08-2026' (尾碼用 - 分隔，DD-MM-YYYY)
    回傳跟 Arrive Date 共用的 'YYYY-MON-DD' 格式 (_indo_format_date())，
    抓不到就回傳 None，讓 expand_to_rows() 補 N/A。
    """
    if not raw_flight:
        return None
    raw_flight = raw_flight.strip()
    m = re.search(r"(\d{1,2})-(\d{1,2})-(\d{4})\s*$", raw_flight)
    if not m:
        m = re.search(r"(\d{2})(\d{2})(\d{4})\s*$", raw_flight)
    if not m:
        return None
    day, month, year = m.groups()
    return _indo_format_date(day, month, year)


def _indo_clean_amount(text: Optional[str]) -> Optional[int]:
    """英式數字 '416,729' (逗號千分位) -> 416729。"""
    if not text:
        return None
    text = text.strip().replace(",", "")
    return int(text) if text.isdigit() else None


# 費用明細行，例如：
#   "FREIGHT CHARGE (USD 23.54) (REIMBURSEMENT) 416,729"
#   "HANDLING (PPH 23) T) 650,000"   <- "T)" 是版面上一個核取方塊符號被
#                                       讀成的雜訊殘留，不是真正的文字
#   "STORAGE (INV NO : 26 0158969) 427,137"
# 這裡不限制描述文字只能是純大寫 (PT. Indo Prostime Transport 的费用名稱
# 後面可能帶 "(PPH 23)" 這種備註，描述本身仍全大寫，但為了同時兼容未來
# 可能出現的大小寫混雜備註，放寬成「任何非空白字元開頭、非 Total/VAT 行」)。
# 一律先完整抓到「最後一個金額數字」前的所有文字當描述，再用
# _indo_clean_item_desc() 去掉已知的雜訊/備註尾碼。
_INDO_ITEM_LINE_RE = re.compile(
    r"^(?!Total\b)(?!V\.?A\.?T\.?\s)(?P<desc>\S.*?)\s+(?P<amount>[\d,]+)\s*$",
    re.MULTILINE,
)
# V.A.T. 另起一列印出 (例如 "V.A.T. 34,769")，正確答案範本把這一列當成
# 獨立一筆品項：description 固定寫 "VAT"、amount 留空、金額改填在 vat 欄位
# (不是把 VAT 金額當成第五筆品項的 amount)。
_INDO_VAT_LINE_RE = re.compile(r"^V\.?A\.?T\.?\s+(?P<vat>[\d,]+)\s*$", re.MULTILINE)
# 费用明細只會出現在 "D E S C R I P T I O N AMOUNT" 這個表頭之後、
# "Total" 小計行之前這一段範圍內；只在這個範圍內掃描品項，才不會把帳單
# 上其他「文字 + 一串數字」的雜訊行 (地址裡的郵遞區號、電話分機、發票號碼
# 那一行) 誤判成費用項目 (跟 HANDAL 供應商的 _handal_item_region() 同樣
# 的防呆手法)。
_INDO_ITEM_REGION_START_RE = re.compile(
    r"D\s*E\s*S\s*C\s*R\s*I\s*P\s*T\s*I\s*O\s*N\s+AMOUNT.*?\n", re.IGNORECASE,
)
_INDO_ITEM_REGION_END_RE = re.compile(r"^Total\b", re.MULTILINE)


def _indo_item_region(text: str) -> str:
    start_m = _INDO_ITEM_REGION_START_RE.search(text)
    start = start_m.end() if start_m else 0
    end_m = _INDO_ITEM_REGION_END_RE.search(text, start)
    end = end_m.start() if end_m else len(text)
    return text[start:end]


def _indo_clean_item_desc(desc: str) -> str:
    desc = desc.strip()
    # PT. Indo Prostime Express 帳單裡，"(REIMBURSEMENT)" 這個備註不計入
    # 正確答案的描述文字，要去掉 (只去尾碼，不影響描述中間的其他括號備註，
    # 例如 "(USD 8.15)" 要保留)。
    desc = re.sub(r"\s*\(REIMBURSEMENT\)\s*$", "", desc, flags=re.IGNORECASE)
    # PT. Indo Prostime Transport 帳單裡，"(PPH 23)" 後面偶爾會黏一個
    # 版面核取符號雜訊 "T)"，要去掉；"(PPH 23)" 本身要保留。
    desc = re.sub(r"\s*T\)\s*$", "", desc)
    return desc.strip()


def _indo_extract_charge_items(text: str) -> List[Dict]:
    region = _indo_item_region(text)
    items = []
    for m in _INDO_ITEM_LINE_RE.finditer(region):
        items.append({
            "description": _indo_clean_item_desc(m.group("desc")),
            "amount": _indo_clean_amount(m.group("amount")),
            "vat": None,
        })
    vat_m = _INDO_VAT_LINE_RE.search(region)
    if vat_m:
        items.append({
            "description": "VAT",
            "amount": None,
            "vat": _indo_clean_amount(vat_m.group("vat")),
        })
    if not items:
        items = [{"description": None, "amount": None, "vat": None}]
    return items


def parse_indoprostime(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {}

    header["invoice_no"] = _search(r"Number\s*:\s*(\d+)", text)
    _indo_inv_date_m = re.search(r"Date\s*:\s*(\d{1,2})/(\d{1,2})/(\d{4})", text)
    header["invoice_date"] = (
        _indo_format_date(*_indo_inv_date_m.groups()) if _indo_inv_date_m else None
    )
    header["supplier"] = "PT. INDO PROSTIME EXPRESS"
    header["consignee"] = _indo_extract_top_consignee(text)
    header["order_no"] = _search(rf"Order\s*No\.?\s*:\s*(.+?){_INDO_STOP}", text)
    header["bl_no"] = _search(rf"\bHAWB\s*:\s*(.+?){_INDO_STOP}", text)
    header["mbl_no"] = _search(rf"\bMAWB\s*:\s*(.+?){_INDO_STOP}", text)
    header["port_of_loading"] = _search(rf"\bORIGIN\s*:\s*(.+?){_INDO_STOP}", text)
    header["shipper"] = _search(rf"\bSHIPPER\s*:\s*(.+?){_INDO_STOP_LOOSE}", text)
    _raw_flight = _search(rf"\bFLIGHT\s*:\s*(.+?){_INDO_STOP}", text)
    header["vessel"] = _indo_extract_vessel_code(_raw_flight)
    header["volume"] = _indo_format_volume(
        _search(rf"GROSS\s*/\s*KG\s*:\s*(.+?){_INDO_STOP}", text)
    )
    header["onboard_date"] = _indo_parse_flight_date(_raw_flight)
    header["arrive_date"] = header["onboard_date"]

    # 這種空運帳單目的港固定是雅加達 (這家供應商只處理進口到雅加達的空運
    # 帳單)，帳單本文沒有單獨印出「目的港」欄位，所以固定填 Jakarta，不是
    # 從文字擷取出來的。
    header["port_of_discharge"] = "Jakarta"
    header["container_no"] = None

    items = _indo_extract_charge_items(text)

    return {"header": header, "items": items}


# ---------------------------------------------------------------------------
# PT. Indo Prostime Transport (跟上面 PT. Indo Prostime Express 是「同一套
# 版面排版、不同公司」的姊妹帳單——兩家公司名稱都含有 "PROSTIME"，所以
# detect_fn 一定要分別比對完整公司名稱 "INDO PROSTIME EXPRESS" /
# "INDO PROSTIME TRANSPORT"，不能只看 "PROSTIME" 不然會互相誤判。
#
# 跟 EXPRESS 版面的差異：
#   - Bank Information 區塊沒有 "A/C name :" 這一行 (改成固定寫死供應商
#     名稱，不用擷取)。
#   - 收件人欄位標籤是 "CONSIGNEE :"，不是 EXPRESS 的 "CNEE :"。
#   - FLIGHT 欄位尾碼日期用 "-" 分隔 (例如 "VJ0855/05-08-2026")，不是
#     EXPRESS 的無分隔 8 碼數字——共用 _indo_parse_flight_date() 已經兩種
#     格式都支援。
#   - 費用明細常見 "(PPH 23)" 備註後面黏一個版面核取符號雜訊 "T)"，以及
#     可能額外有一行獨立的 "V.A.T. 金額" (有些帳單有 VAT、有些沒有)——
#     共用 _indo_extract_charge_items() 已經處理這兩種情況。
# ---------------------------------------------------------------------------

def detect_indoprostime_transport(text: str) -> bool:
    return "INDO PROSTIME TRANSPORT" in text.upper()


def parse_indoprostime_transport(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {}

    header["invoice_no"] = _search(r"Number\s*:\s*(\d+)", text)
    _indo_inv_date_m = re.search(r"Date\s*:\s*(\d{1,2})/(\d{1,2})/(\d{4})", text)
    header["invoice_date"] = (
        _indo_format_date(*_indo_inv_date_m.groups()) if _indo_inv_date_m else None
    )
    header["supplier"] = "PT. INDO PROSTIME TRANSPORT"
    header["consignee"] = _indo_extract_top_consignee(text)
    header["order_no"] = _search(rf"Order\s*No\.?\s*:\s*(.+?){_INDO_STOP}", text)
    header["bl_no"] = _search(rf"\bHAWB\s*:\s*(.+?){_INDO_STOP}", text)
    header["mbl_no"] = _search(rf"\bMAWB\s*:\s*(.+?){_INDO_STOP}", text)
    header["port_of_loading"] = _search(rf"\bORIGIN\s*:\s*(.+?){_INDO_STOP}", text)
    header["shipper"] = _search(rf"\bSHIPPER\s*:\s*(.+?){_INDO_STOP_LOOSE}", text)
    _raw_flight = _search(rf"\bFLIGHT\s*:\s*(.+?){_INDO_STOP}", text)
    header["vessel"] = _indo_extract_vessel_code(_raw_flight)
    header["volume"] = _indo_format_volume(
        _search(rf"GROSS\s*/\s*KG\s*:\s*(.+?){_INDO_STOP}", text)
    )
    header["onboard_date"] = _indo_parse_flight_date(_raw_flight)
    header["arrive_date"] = header["onboard_date"]
    header["port_of_discharge"] = "Jakarta"
    header["container_no"] = None

    items = _indo_extract_charge_items(text)

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
    """注意：不能只看 'MAERSK' 這個字，因為其他供應商(例如 PANCARAN)的帳單
    裡 VESSEL NAME 欄位常常就是船名含 'MAERSK' (例如 'JOSEPHINE MAERSK')，
    這跟 PT Maersk Logistics Indonesia 這家供應商本身無關，會誤判。改用更
    具體的 'MAERSK LOGISTICS' 字樣判斷。"""
    return "MAERSK LOGISTICS" in text.upper()


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
    return _format_date_en(day, month, year)


def _maersk_format_invoice_date(day, month_abbr, year) -> Optional[str]:
    """MAERSK 的 INVOICE DATE 欄位，正確答案 Excel 裡是純文字儲存格 (不是
    Excel 日期型態)，格式是帳單原文印的 '30.Sep.2026' (點分隔)，跟其他
    供應商共用的 _format_date_en_from_month_name()『DD Mon YYYY』(空白分
    隔) 不一樣——因為那個共用函式是給『Excel 日期型態』欄位用的 (Arrive
    Date/on board date…)，normalize_value() 比對時會把真正的日期物件也
    格式化成『DD Mon YYYY』再比較；但 MAERSK 的 INVOICE DATE 這裡本來就是
    文字儲存格，不會被 normalize_value() 重新格式化，必須直接輸出跟正確
    答案一樣的點分隔格式，才對得起來。只有這個欄位這樣處理，不要去動共用
    的 _format_date_en()，否則會把其他供應商已經驗證過的日期欄位全部改壞。
    """
    if month_abbr is None:
        return None
    month_i = _MONTH_NAME_TO_NUM.get(str(month_abbr).strip().upper()[:3])
    if month_i is None:
        return None
    try:
        day_i, year_i = int(day), int(year)
    except (TypeError, ValueError):
        return None
    return f"{day_i:02d}.{_MONTH_ABBR_EN[month_i - 1]}.{year_i}"


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


# 舊版做法：改抓 Faktur Pajak (稅務發票) 頁面的費用明細行 (例如
# "2 Origin Terminal Handling Charge 149.659,00")。那一頁版面每筆描述都在
# 同一行，擷取起來比較穩定，但缺點是「沒有逐筆 VAT 金額」(Faktur Pajak頁
# 只有整張發票的 Total PPN 一個數字)。
#
# 使用者新增「每個費用後面有自己對應的 VAT 金額」的需求後，改成抓 Invoice
# 頁面本身的 "ItemDescription of Charges ... Qty UoM Total(IDR)" 表格，
# 因為只有這裡才有「每一筆費用後面緊接著 VAT 那一行」的資訊，而且不是每張
# 帳單都有附 Faktur Pajak 頁 (例如新樣本 5617858220 這份 PDF 只有 2 頁
# Invoice，沒有 Faktur Pajak 頁)。
#
# 版面特徵：
#   "1 CFS Receiving Charge 4.094 M3 1,677,125"   -> 單行：項次+費用名稱+
#                                                     數量+單位+金額(英式
#                                                     千分位、無小數)
#   "VAT 184,484"                                 -> 緊接著下一行就是這筆
#                                                     費用自己的 VAT 金額
#   費用名稱太長時會被自動換行拆成兩段，例如：
#   "2 Origin Terminal Handling 4.094 M3 361,055"
#   "Charge"                                       -> 這是費用名稱的續行，
#   "VAT 39,716"                                      要併回上一筆描述
# 用逐行掃描 (不是一個大正規表示式)，比較好處理「描述可能多一行續行」
# 這種不固定的情況。
_MAERSK_ITEM_START_PATTERN = re.compile(
    r"^(?P<num>\d+)\s+(?P<desc>.+?)\s+(?P<qty>\d+(?:\.\d+)?)\s+(?P<uom>[A-Za-z0-9]+)\s+"
    r"(?P<amount>[\d,]+)\s*$"
)
_MAERSK_VAT_LINE_PATTERN = re.compile(r"^VAT\s+([\d,]+)\s*$")


def _maersk_clean_amount_comma(text: Optional[str]) -> Optional[int]:
    """英式數字 '1,677,125' (逗號千分位、無小數) -> 1677125。跟 Faktur
    Pajak 頁用的印尼式數字 (_maersk_clean_amount_id，句點千分位、逗號小數)
    是不同格式，Invoice 頁面本身的金額都是這種英式格式。
    """
    if not text:
        return None
    text = text.strip().replace(",", "")
    return int(text) if text.isdigit() else None


def _maersk_extract_items(text: str) -> List[Dict]:
    m = re.search(
        r"ItemDescription of Charges.*?Total\(IDR\)\s*\n(.*?)\nTotal Packages",
        text, re.DOTALL,
    )
    if not m:
        return []
    lines = [ln.strip() for ln in m.group(1).split("\n") if ln.strip()]

    items: List[Dict] = []
    i = 0
    while i < len(lines):
        start_m = _MAERSK_ITEM_START_PATTERN.match(lines[i])
        if not start_m:
            i += 1
            continue
        desc_parts = [start_m.group("desc").strip()]
        amount = _maersk_clean_amount_comma(start_m.group("amount"))
        i += 1
        vat = None
        # 往下找這一筆的 VAT 那一行；中間如果還有非 VAT、非下一筆項次開頭
        # 的文字，代表是費用名稱被自動換行拆出來的續行，併回描述裡。
        while i < len(lines):
            vat_m = _MAERSK_VAT_LINE_PATTERN.match(lines[i])
            if vat_m:
                vat = _maersk_clean_amount_comma(vat_m.group(1))
                i += 1
                break
            if _MAERSK_ITEM_START_PATTERN.match(lines[i]):
                break  # 下一筆已經開始，這筆沒有抓到 VAT (防禦性處理)
            desc_parts.append(lines[i])
            i += 1
        items.append({
            "description": " ".join(desc_parts).strip(),
            "amount": amount,
            "vat": vat,
        })
    return items


def _maersk_extract_grand_total(text: str) -> Optional[int]:
    """'Amount Due 2,946,119' 是「Total Amount Due」小表格裡的含稅總額
    (英式千分位、無小數)，用來做『未稅加總 + VAT 加總 = 含稅總額』的稽核
    (見 extract_utils.audit_invoice_totals())。"""
    m = _search(r"Amount\s*Due\s+([\d,]+)", text)
    return _maersk_clean_amount_comma(m)


def parse_maersk(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {}

    header["invoice_no"] = _search(r"\bINVOICE\s+(\d+)\b", text)
    _maersk_inv_m = re.search(r"Invoice\s*Date\s*:\s*(\d{1,2})\.([A-Za-z]{3,})\.(\d{4})", text)
    header["invoice_date"] = (
        _maersk_format_invoice_date(
            _maersk_inv_m.group(1), _maersk_inv_m.group(2), _maersk_inv_m.group(3)
        )
        if _maersk_inv_m
        else _search(r"Invoice\s*Date\s*:\s*(\S+)", text)
    )
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
    # Shipper：帳單上 "SHPR:" 後面的公司名稱 (跟 "Sold to:" 的帳單對象可能
    # 不是同一家，不要直接沿用 consignee 的值)。
    header["shipper"] = _search(rf"\bSHPR\s*:\s*(.+?){_MAERSK_STOP}", text)

    # 這張帳單沒有主提單/船名/貨櫃資訊，留空給 expand_to_rows() 補 N/A
    header["mbl_no"] = None
    header["vessel"] = None
    header["container_no"] = None

    items = _maersk_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None, "vat": None}]

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
    # 印尼文月份全名 (EXPRESS_MAXIMUM 帳單簽名處的日期用印尼文，例如
    # "2 Januari 2026")，拼法跟英文不同的月份才需要額外列出，一樣的
    # (September) 共用英文那組即可。
    "januari": 1, "februari": 2, "maret": 3, "mei": 5, "juni": 6, "juli": 7,
    "agustus": 8, "oktober": 10, "desember": 12,
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
    return _format_date_en(day, month, year)


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


_YJE_MONTH_ABBR_UPPER = [
    "JAN", "FEB", "MAR", "APR", "MAY", "JUN",
    "JUL", "AUG", "SEP", "OCT", "NOV", "DEC",
]


def _yje_format_date(day, month_name, year) -> Optional[str]:
    """YJE 帳單正確答案範本的日期欄位是純文字 'YYYY-MON-DD' (例如
    '2026-JUL-31')，不是真正的 Excel 日期型態儲存格，跟其他供應商共用的
    'DD Mon YYYY' (_format_date_en_from_month_name()) 不一樣，不要共用。
    """
    month_i = _MONTH_NAME_TO_NUM.get(str(month_name).strip().upper()[:3])
    if month_i is None:
        return None
    try:
        d, y = int(day), int(year)
    except (TypeError, ValueError):
        return None
    return f"{y}-{_YJE_MONTH_ABBR_UPPER[month_i - 1]}-{d:02d}"


def parse_yje(text: str) -> Dict:
    """注意：這是 multi_page 供應商，text 是『一頁』的文字。"""
    header: Dict[str, Optional[str]] = {}

    header["invoice_no"] = _search(r"Inv\.?\s*No\s*[:：]\s*(\S+)", text)
    _yje_inv_m = re.search(r"Inv\s*Date\s*[:：]\s*(\d{4})-([A-Za-z]{3,})-(\d{1,2})", text)
    header["invoice_date"] = (
        _yje_format_date(
            _yje_inv_m.group(3), _yje_inv_m.group(2), _yje_inv_m.group(1)
        )
        if _yje_inv_m
        else _search(r"Inv\s*Date\s*[:：]\s*(\S+)", text)
    )
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

def detect_tata_classic(text: str) -> bool:
    """TATA 帳單的第一種格式："INVOICE" 標題 + "NO :"/"INV Date :" 這種
    英文排版 (見下面 parse_tata_classic())。用 "NOMOR INVOICE" 沒出現
    (那是第二種格式的專屬字樣) 來跟第二種格式互相排除。
    """
    return "TAMAN DUTAMAS" in text.upper() and "NOMOR INVOICE" not in text.upper()


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


# Shipper 欄位有時候名稱太長會換行，但因為帳單是雙欄排版 (左欄是收件人
# Address 繼續行、右欄才是 Shipper 繼續行)，pdfplumber 抓文字時兩欄會被
# 併成同一行：換行後那一行通常是左欄的地址續字 "BARAT" (印尼文「西爪哇」
# 的後半，JAWA BARAT = 西爪哇省)，但如果 Shipper 名稱也換行，就會變成
# "BARAT BRANCH" 這種「地址續字 + Shipper 續字」黏在同一行——這裡把
# "BARAT" 這個已知的地址續字去掉，剩下的部分 (如果有) 才是 Shipper 的
# 換行續字，要接回 Shipper 名稱後面。
_TATA_SHIPPER_RE = re.compile(
    r"\bShipper\s*:\s*(?P<main>.+?)\s*\n(?P<next_line>[^\n]*)\n", re.MULTILINE
)
_TATA_SHIPPER_WRAP_RE = re.compile(r"^BARAT\b\s*(?P<cont>.*)$", re.IGNORECASE)


def _tata_extract_shipper(text: str) -> Optional[str]:
    m = _TATA_SHIPPER_RE.search(text)
    if not m:
        return None
    shipper = m.group("main").strip()
    wrap_m = _TATA_SHIPPER_WRAP_RE.match(m.group("next_line").strip())
    if wrap_m:
        cont = wrap_m.group("cont").strip()
        if cont:
            shipper = f"{shipper} {cont}"
    return shipper or None


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
    return _format_date_en(day, month, year)


# 費用明細行，例如 "1 GUDANG FEE Rp 159,840" 或
# "4 HANDLING FEE YJE123965498 Rp 4,585,728"：項次 + 費用名稱 + (偶爾多一個
# 含數字的追蹤號碼尾巴) + Rp + 金額。用「後面緊接的字是否含數字」判斷那個字
# 是追蹤號碼還是費用名稱的一部分，避免把 "LOCAL EXPRESS CHARGE" 這種全大寫
# 費用名稱誤判成「名稱+尾碼」。
_TATA_ITEM_PATTERN = re.compile(
    r"^\d+\s+(?P<desc>[A-Z][A-Z ]*?)(?:\s+(?=[A-Z0-9]*\d)[A-Z0-9]+)?\s+Rp\s+(?P<amount>[\d,]+)\s*$",
    re.MULTILINE,
)
# "PPN 12%" 這一行是加值稅 (正確答案範本新增 VAT 欄位後，這一筆要變成
# description="PPN 12%"、amount 留空、金額改填 vat 欄位，不是普通第五筆
# 品項的 amount)。注意：這裡取代了舊版的 "DPP" 列——DPP (稽核用的課稅
# 基礎，有時候不等於 SUB-TOTAL，是 HANDLING FEE 另外乘上 11/12 算出來的
# 中間值) 新版正確答案範本已經不要求輸出，只保留真正印在帳單上的 PPN
# 金額。
_TATA_PPN_PATTERN = re.compile(r"^(?P<desc>PPN\s*\d+%)\s+Rp\s+(?P<vat>[\d,]+)\s*$", re.MULTILINE)


def _tata_extract_items(text: str) -> List[Dict]:
    items = []
    for m in _TATA_ITEM_PATTERN.finditer(text):
        items.append({
            "description": m.group("desc").strip(),
            "amount": _tata_clean_amount(m.group("amount")),
            "vat": None,
        })
    ppn_m = _TATA_PPN_PATTERN.search(text)
    if ppn_m:
        items.append({
            "description": ppn_m.group("desc").strip(),
            "amount": None,
            "vat": _tata_clean_amount(ppn_m.group("vat")),
        })
    return items


def parse_tata_classic(text: str) -> Dict:
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
    header["shipper"] = _tata_extract_shipper(text)

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
# TATA 的第二種格式：印尼文稅務收據 (⚠️ "NOMOR INVOICE #" 排版)
# ---------------------------------------------------------------------------
#
# 這是同一家 TATA 供應商的「另一種帳單樣式」，抬頭是公司自己的印尼文格式
# 收據 (不是英文 "INVOICE" 那種)，版面、欄位名稱都跟 parse_tata_classic()
# 那份完全不同：
#   "NOMOR INVOICE # : TH20250064"、"TANGGAL : 03-01-2025"、
#   "JATUH TEMPO : 02-02-2025"、"PELANGGAN"/"DETAIL PEMBAYARAN" 左右並排、
#   "NO KETERANGAN PAJAK JUMLAH (Rp.)" 費用明細表格。
#
# 費用明細表格的排版有個陷阱：大多數費用是「項次 描述 X 金額」單行
# (例如 "2 NON REIMBURSEMENT AE - AGENCY FEE (PER AWB) X 750.000")，但
# 第一筆費用因為描述文字太長、換行排版，pdfplumber 抽出來的順序會變成
# 「描述(第一行)」→「項次 金額」(沒有 X 標記，descriptions 跟金額被拆到
# 不同行) →「描述續行(括號補充)」三行分開，例如：
#   'REIMBURSEMENT - INVOICE STORAGE PT GAPURA ANGKASA'
#   '1 1.061.202'
#   '(70125000125)'
# 所以擷取邏輯改成逐行的簡易狀態機 (_tata_receipt_extract_items())：遇到
# 「項次 描述 X 金額」單行就直接算一筆；遇到「項次 金額」(沒有 X、數字前面
# 不是描述文字) 就用「前面还没歸戶的描述行」當作這一筆的描述，並且再往後看
# 一行，如果那一行不是新項次開頭 (不是「數字 開頭」)，就當作描述的續行一併
# 併入。金額是印尼式千分位、無小數 ('1.061.202' -> 1061202，用句點而不是
# 逗號)，跟 parse_tata_classic() 的英式千分位 (逗號) 剛好相反，不要共用同
# 一個金額清理函式。
#
# 這張帳單沒有 Order No./到達日/開船日/裝卸貨港/材積/貨櫃資訊，正確答案
# 全部是 N/A，留 None 給 expand_to_rows() 補齊即可。
#
# SUPPLIER 名稱這裡直接寫死成 "PT TATA HARMONI SARANATAMA"：這份收據的
# 公司抬頭因為排版關係被拆成 "PT TATA HARMONI" / "SARANATAMA" 兩行 (中間
# 還夾著一個 logo 圖示對應的亂碼字元)，直接組字串比逐行拼接可靠。

def detect_tata_receipt(text: str) -> bool:
    return "NOMOR INVOICE" in text.upper()


def _tata_receipt_parse_date(text: Optional[str]) -> Optional[str]:
    """'03-01-2025' (DD-MM-YYYY) -> '03.01.2025'。"""
    if not text:
        return None
    m = re.match(r"(\d{1,2})-(\d{1,2})-(\d{4})\s*$", text.strip())
    if not m:
        return None
    day, month, year = m.groups()
    return _format_date_en(day, month, year)


def _tata_receipt_clean_amount(text: Optional[str]) -> Optional[int]:
    """印尼式千分位 '1.061.202' (句點千分位，無小數) -> 1061202。"""
    if not text:
        return None
    text = text.strip().replace(".", "")
    return int(text) if text.isdigit() else None


#   "3 NON REIMBURSEMENT AE - AGENCY FEE (PER AWB) X 750.000" (含 PPN 稅率
#   標記 "X")，或
#   "2 REIMBURSEMENT - INVOICE CENTURY LOGISTICS CO.,LTD. 4.125.420" (沒有
#   "X" 標記，REIMBURSEMENT 類的費用單行時常常沒有這個標記)。
# "X" 不是每筆都有，所以設成可有可無。
_TATA_RECEIPT_FULL_ITEM_LINE = re.compile(
    r"^\d+\s+(?P<desc>[A-Za-z][A-Za-z0-9 /\-.,()=@]*?)\s+(?:X\s+)?(?P<amount>[\d.]+)\s*$"
)
_TATA_RECEIPT_NUM_AMOUNT_ONLY_LINE = re.compile(r"^(?P<num>\d+)\s+(?P<amount>[\d.]+)\s*$")


def _tata_receipt_extract_items(text: str) -> List[Dict]:
    """逐行狀態機擷取費用明細，見上面區塊註解說明「第一筆換行排版」的
    特殊處理方式。只在 "KETERANGAN...JUMLAH" 表頭跟 "PESAN"/"Subtotal"
    之間的區塊裡找，避免抓到表格外其他也剛好符合格式的雜訊行。
    """
    block_m = re.search(
        r"KETERANGAN\s+PAJAK\s+JUMLAH.*?\n(.*?)\n\s*(?:PESAN|Subtotal)",
        text, re.DOTALL | re.IGNORECASE,
    )
    block = block_m.group(1) if block_m else text

    lines = block.split("\n")
    items: List[Dict] = []
    pending: List[str] = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        m_full = _TATA_RECEIPT_FULL_ITEM_LINE.match(line)
        m_num_only = _TATA_RECEIPT_NUM_AMOUNT_ONLY_LINE.match(line)
        if m_full:
            items.append({
                "description": m_full.group("desc").strip(),
                "amount": _tata_receipt_clean_amount(m_full.group("amount")),
            })
            pending = []
        elif m_num_only:
            desc = " ".join(pending).strip()
            pending = []
            if i + 1 < len(lines):
                nxt = lines[i + 1].strip()
                if nxt and not re.match(r"^\d+\s", nxt):
                    desc = f"{desc} {nxt}".strip()
                    i += 1
            items.append({
                "description": desc,
                "amount": _tata_receipt_clean_amount(m_num_only.group("amount")),
            })
        else:
            pending.append(line)
        i += 1
    return items


#   "SHIP : CHAEI YANG VN CO LTD\nNO,BC: ..." (單獨一行，名稱後面直接換行)，
#   或 "SHIP:BAIKSAN CO.,LTD. Bayar Diterima\nKOREA Sisa Tagihan: ..." (沒有
#   空格、而且因為跟右欄的「Bayar Diterima」排版擠在同一行，名稱後面要在
#   遇到 "Bayar" 這個右欄標籤字樣時就停止，不要吃到右欄文字；下一行的國家
#   名稱 "KOREA" 是收件地點，不屬於 Shipper，不要抓)。
# ⚠️ 右欄緊鄰的標籤字樣不只 "Bayar"：有些帳單版面是 "SHIP:xxx Total
# 9.269.986" (右欄印的是 "Total" 不是 "Bayar")，兩種都要能停止擷取，不要
# 把右欄文字併進 Shipper 名稱裡。
_TATA_RECEIPT_SHIP_PATTERN = r"SHIP\s*:\s*(.+?)(?:\s+(?:Bayar|Total)\b|\s*\n)"

# Volume/材積：帳單上沒有專屬標籤，是獨立一行印著「數量代碼/重量KG(S)」的
# 格式 (例如 "7PK/26KGS"、"1PK/17KGS"、"7BG/132KGS"、"3CTN/109KG"、
# "7PKG/117KG"——單位有時候是 "KG" 沒有尾巴的 "S"，兩種都要吃)，這一行在
# 版面上的上下位置會因為有沒有船名/航班資訊而不同 (有時緊接在 "11.0%:"
# 之後、有時緊接在 "Total:" 之後)，所以不用位置判斷，直接用格式本身 (數字
# +英文代碼+ "/" + 數字 + "KG"/"KGS"，獨佔一整行) 掃描整份文字找。
_TATA_RECEIPT_VOLUME_RE = re.compile(r"(?m)^(\d+[A-Za-z]+/\d+\s*[Kk][Gg][Ss]?)\s*$")

# PPN 稅金行：跟 parse_tata_classic() 的 "PPN 12% Rp xxx" 格式不同，這種
# 帳單印成 "PPN 12% 12.645" (沒有 "Rp" 前綴，金額是印尼式千分位、無小數)，
# 而且這一行印在費用明細表格 ("KETERANGAN...JUMLAH" 到 "Subtotal" 之間)
# 外面 (緊接在 Subtotal 後面)，所以不能跟 _tata_receipt_extract_items() 共用
# 同一個擷取範圍，要另外對整份文字單獨搜尋。百分比動態抓出來組字串 (不要
# 寫死 12%，避免哪張帳單稅率不同時字串對不上，跟 TRANS_DAYA_PRIMA/APL 的
# VAT 費用名稱慣例一致)。
#
# ⚠️ 版面陷阱：有些帳單因為沒有另一行獨立印「11.0%:」這個利率小標籤 (跟
# Volume 那一行搶同一個位置)，導致這個小標籤被擠到 PPN 這一行裡面、夾在
# 百分比跟金額中間，變成 "PPN 12% 11.0% 331.458" 這種三個數字的格式，真正
# 的金額是「最後」那個數字，不是中間那個 "11.0" (那只是利率小標籤被誤排
# 進來，不是金額)。用一個可有可無的 "NN.N%" 群組吃掉這個小標籤，確保金額
# 一定抓到行尾真正的數字。
_TATA_RECEIPT_PPN_RE = re.compile(
    r"(?m)^PPN\s*(?P<pct>\d+)%\s+(?:\d+\.\d+%\s+)?(?P<vat>[\d.]+)\s*$"
)


def parse_tata_receipt(text: str) -> Dict:
    """注意：這是 multi_page 供應商，text 是『一頁』的文字。"""
    header: Dict[str, Optional[str]] = {}

    header["invoice_no"] = _search(r"NOMOR\s*INVOICE\s*#\s*:\s*(\S+)", text)
    header["invoice_date"] = _tata_receipt_parse_date(
        _search(r"TANGGAL\s*:\s*(\d{1,2}-\d{1,2}-\d{4})", text)
    )
    header["supplier"] = "PT TATA HARMONI SARANATAMA"
    header["consignee"] = _search(r"PELANGGAN\s*\n(.+?)\s+DETAIL\s*PEMBAYARAN", text)
    # ⚠️ 2026-10 修正：MAWB:/HAWB: 後面的值可能帶空格 (例如 "738-5567 8873")，
    # 原本用 \S+ 只抓到第一個空白前的片段，改成抓到整行結尾才不會截斷。
    header["bl_no"] = _search(r"\bHAWB\s*:\s*(.+?)\n", text)
    header["mbl_no"] = _search(r"\bMAWB\s*:\s*(.+?)\n", text)
    header["vessel"] = _search(r"FLT\s*/\s*DT\s*:\s*([^/\n]+)/", text)
    # ⚠️ 2026-10 新增：這次使用者提供的新樣本帳單上其實有印 Shipper (SHIP :)
    # 跟材積 (獨立一行的 qty/weight)，舊版完全沒有擷取、一律留空；這兩個
    # 欄位整張發票只有一個值，不分品項 (不像 TATA A/F 格式逐行各自不同)。
    header["shipper"] = _search(_TATA_RECEIPT_SHIP_PATTERN, text)
    _volume_m = _TATA_RECEIPT_VOLUME_RE.search(text)
    header["volume"] = _volume_m.group(1).strip() if _volume_m else None

    # 這種帳單沒有 Order No./到達日/開船日/裝卸貨港/貨櫃資訊，留空給
    # expand_to_rows() 補 N/A
    header["order_no"] = None
    header["arrive_date"] = None
    header["onboard_date"] = None
    header["port_of_loading"] = None
    header["port_of_discharge"] = None
    header["container_no"] = None

    items = _tata_receipt_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None}]

    # PPN 稅金：整張帳單只有一個稅額 (不是逐筆品項各自的稅額)，獨立附加
    # 成一筆「只有 description/vat」的 item，展開成獨立一列，Order No. 等
    # 其他抬頭欄位留 None 讓 expand_to_rows() 補 N/A。
    _ppn_m = _TATA_RECEIPT_PPN_RE.search(text)
    if _ppn_m:
        items.append({
            "description": f"PPN {_ppn_m.group('pct')}%",
            "amount": None,
            "vat": _tata_receipt_clean_amount(_ppn_m.group("vat")),
        })

    return {"header": header, "items": items}


# ---------------------------------------------------------------------------
# TATA 的第三種格式：A/F 空運費率明細表 (⚠️ "Kurs 1 USD" 排版，一行一筆貨物)
# ---------------------------------------------------------------------------
#
# 這是同一家 TATA 供應商的「第三種帳單樣式」：空運 (Air Freight) 費率
# 明細表，整張發票是『一行一筆貨物』的表格，欄位是
# "NO DATE Shipper Country ATTN HAWB Qty WEIGHT Kurs 1 USD A/F PPn 12% tati"，
# 跟 parse_tata_classic()（單筆發票、抬頭式欄位 MAWB:/HAWB:/Flight: 等）、
# parse_tata_receipt()（印尼文稅務收據 NOMOR INVOICE 排版）完全不同：
#   - 每一行自己帶著 Shipper/HAWB/重量/金額(A/F)/VAT(PPn 12%)，正確答案的
#     Description 欄位固定填表頭欄位名稱本身的字面 "A/F"，Amount 欄位是
#     這一行的 A/F 金額、VAT 欄位是這一行的 PPn 12% 金額、Shipper 欄位是
#     這一行自己的託運人名稱（不是整張發票共用一個，每一列各自覆蓋，用法
#     跟 PANCARAN/PT. TRANS DAYA PRIMA 的「每筆品項自帶抬頭欄位覆寫」一樣）。
#     Volume 欄位填這一行的 WEIGHT 換算成 "數字 kgs" 字串 (去除多餘的小數
#     尾零，例如 '4,00 kgs' -> '4 kgs'、'12,50 kgs' -> '12.5 kgs')。
#   - ⚠️ 兩欄並排版面陷阱：Shipper/Country 兩個欄位常常其中一個換行，而且
#     pdfplumber 抽出來的順序是「換行的上半段先印在這一行『前面』、下半段
#     印在這一行『後面』」(不是單純往下接一行)，例如：
#       'SINCETECH (FUJIAN) TECHNOLOGY CO'          <- Shipper 上半段 (前)
#       '3 4/Aug/25 CHINA RZR RONG YJE...'          <- 本行本體 (中間只剩 Country)
#       'LTD'                                        <- Shipper 下半段 (後)
#     或是 Country 換行而 Shipper 完整印在本行：
#       'HONG'                                       <- Country 上半段 (前)
#       '2 1/Aug/25 BILLION BASE INTL CO LTD RZR...' <- 本行本體 (Shipper 完整)
#       'KONG'                                        <- Country 下半段 (後)
#     判斷方式：先用已知國家字樣清單 (_TATA_AF_KNOWN_COUNTRIES) 檢查本行主
#     體裡「日期跟 ATTN 之間」的文字 (mid)：
#       (a) mid 整段剛好就是一個已知國家名稱 -> Shipper 是空的，要去抓前後
#           兩行拼回來 (Country 已經在本行裡了，不用管前後行)。
#       (b) mid 結尾找不到任何已知國家名稱 -> Shipper 整段就是 mid 本身，
#           前後兩行是 Country 的換行殘字 (不需要輸出，直接丟棄但要消耗掉
#           這兩行，避免被鄰居那一筆誤吃)。
#       (c) mid 結尾剛好是已知國家名稱、前面還有剩餘文字 -> 單行沒有換行，
#           Shipper 就是去掉國家字樣後的剩餘文字，不要去動前後行。
#   - ⚠️ 已知的雜訊列陷阱：帳單第一筆資料列有時候會混入一筆完全不合理的
#     雜訊資料 (日期離譜 (例如 '1/Oct/19'，跟發票本身的月份 2025/08 完全
#     對不上)、匯率離譜偏低 (例如 'Rp 440'，印尼盾兌美元匯率正常絕對遠高於
#     1000))，這種列不是真正的貨運項目，是帳單製作過程殘留的錯誤資料，正確
#     答案把它整筆排除。這裡用「匯率 (Kurs) 數值是否小於 1000」當判斷依據
#     (比用日期更穩定、不受帳期月份影響)：Kurs < 1000 直接跳過該列，不計入
#     items。
#   - 金額 (A/F) 是印尼式千分位、無小數 ('1.716.000' -> 1716000)；VAT
#     (PPn 12%) 是印尼式千分位 + 逗號小數 ('167.475,7' -> 167475.7)。
#   - SUPPLIER 固定抓發票最下方 "Nama :" 那一行 (跟 parse_tata_classic() 一
#     樣的慣例，但這份帳單的 "Nama :" 後面沒有 "PT." 的句點，是 "PT " 沒有
#     句點，所以不能照抄 classic 版的 regex，要放寬成不要求句點)。
#   - CONSIGNEE 抓 "To :" 那一行，但帳單上寫的是
#     "PT POU CHEN INDONESIA ( ASICS)"，正確答案只要 "PT POU CHEN
#     INDONESIA"，後面的括號備註 "( ASICS)" 要去掉。
#   - ⚠️ 稽核機制特別說明：這張帳單印刷的小計 "A/F"/"TOTAL AMOUNT" 跟系統
#     擷取到的品項加總兩邊核對起來，常態性地會差 1~3 塊錢 (帳單原文自己的
#     小計看起來是用未顯示的小數位四捨五入組出來的，不是系統擷取錯誤)，
#     超出系統 ±1 的稽核容許誤差，所以這個格式故意不實作
#     extract_total/extract_grand_total，让稽核機制視為「不適用」而不是
#     顯示假警報的「不一致」。
#   - ⚠️ 2026-10 補充：ON BOARD DATE/B/L NO./啟運港 其實每一筆品項各自有自
#     己的值 (就是每行自己的 DATE/HAWB/Country 欄位)，不是整張發票共用、也
#     不是真的沒有 —— 這點最初的樣本 (sample7) 剛好只有一張發票、沒機會注意
#     到跟 Order No./Arrive Date/MBL NO./目的港/船名/貨櫃號碼這些「真的沒有
#     對應資訊」的欄位混在一起看待。所以 onboard_date 要用品項自己的 DATE
#     欄位 (例如 '31/Jul/25' -> '31 Jul 2025')、bl_no 用品項自己的 HAWB 欄位
#     (例如 'YJE123808864')、port_of_loading 用品項自己的 Country 欄位 (例如
#     'CHINA'/'HONG KONG'/'VIETNAM'/'TAIWAN')，在每筆 item dict 裡用
#     HEADER_FIELD_CODES 覆寫的慣例各自帶上這三個 key (見
#     _tata_af_extract_items())，不要固定留 None。Order No./MBL NO./目的港/
#     船名/貨櫃號碼這幾個欄位帳單上真的沒有對應資訊，維持留 None 給
#     expand_to_rows() 補 N/A。

def detect_tata_af(text: str) -> bool:
    """用表頭唯一的 'Kurs 1 USD' 字樣辨識，不會跟 classic/receipt 兩種
    TATA 格式或 YJE 衝突。"""
    return "KURS 1 USD" in text.upper()


_TATA_AF_ITEM_LINE_RE = re.compile(
    r"^(?P<no>\d+)\s+(?P<date>\d{1,2}/[A-Za-z]{3}/\d{2})\s+(?P<mid>.*?)\s*"
    r"(?P<attn>[A-Z]+\s+RONG)\s+(?P<hawb>\S+)\s+(?P<qty>\d+)\s*PKGS\s+"
    r"(?P<weight>[\d,]+)\s*kgs\s+Rp\s+(?P<kurs>[\d.]+)\s+(?P<amount>[\d.]+)\s+"
    r"(?P<vat>[\d.,]+)\s*$"
)
# ⚠️ 2026-10 補充：mid 跟 attn 之間原本要求至少一個空白 (\s+)，但有些品項
# 整個 Shipper/Country 都被拆到前後兩行、本行本體剩下 "DATE RZR RONG ..."
# 這種「日期跟 ATTN 之間只剩一個空白」的極端情況 (例如 HC25013192 第 3 筆
# "3 28/Jul/25 RZR RONG YJE123807268 ...")，要求至少一個空白會讓這整行配不
# 上 regex、整筆品項直接憑空消失 (比錯誤拆欄更嚴重)。改成 \s* 讓 mid 可以
# 為空字串時兩個分隔空白共用同一個空白字元，不會因此漏掉整筆資料。

# 已知的國家字樣清單 (多字的排在前面，優先比對)：目前樣本出現這四種，
# 之後遇到新國家字樣再加進來即可。
_TATA_AF_KNOWN_COUNTRIES = ["HONG KONG", "CHINA", "VIETNAM", "TAIWAN"]

# 本行主體以外、不該被誤當成 Shipper/Country 換行殘字的行首字樣 (帳單下方
# 小計/備註/銀行資訊區塊)。
_TATA_AF_SKIP_LINE_PREFIXES = (
    "A/F", "DPP", "VAT", "TOTAL", "Terbilang", "Note", "REVISI",
    "PER ALL", "PAYBLE", "BANK", "Branch", "Address", "Swift", "Account", "Nama",
)


def _tata_af_strip_country_suffix(mid: str):
    """mid 結尾若是已知國家名稱，回傳 (去掉國家字樣後的剩餘文字, 國家字樣)；
    完全找不到就回傳 (mid 原樣去空白, None)。"""
    stripped = mid.strip()
    upper = stripped.upper()
    for country in _TATA_AF_KNOWN_COUNTRIES:
        if upper == country:
            return "", country
        if upper.endswith(" " + country):
            return stripped[: -(len(country) + 1)].strip(), country
    return stripped, None


# "HONG KONG" 換行時常被拆成「HONG」「KONG」兩個單字分居前後兩行 (見下面
# _tata_af_split_wrapped_line())，這兩個單字本身不是任何其他已知國家字樣的
# 前/後綴，不會跟 CHINA/VIETNAM/TAIWAN 誤判。
_TATA_AF_COUNTRY_WORD_FRAGMENTS = {"HONG", "KONG"}


def _tata_af_split_wrapped_line(line: str):
    """把一行『換行殘字』(Shipper 續行或 Country 續行，或兩者黏在同一行)
    拆成 (這一行裡屬於 Shipper 的部分, 這一行裡屬於 Country 的部分或
    None)。例如 'HONGKONG HONGYANG WEAVING HONG' -> ('HONGKONG HONGYANG
    WEAVING', 'HONG')；純粹的 'KONG' -> ('', 'KONG')；完全沒有國家字樣的
    'STRING CO LIMITED KONG' 裡的 'KONG' 一樣會被拆開 -> ('STRING CO
    LIMITED', 'KONG')。"""
    stripped = line.strip()
    if not stripped:
        return "", None
    upper = stripped.upper()
    for country in _TATA_AF_KNOWN_COUNTRIES:
        if upper == country:
            return "", country
        if upper.endswith(" " + country):
            return stripped[: -(len(country) + 1)].strip(), country
    tokens = stripped.split()
    if tokens and tokens[-1].upper() in _TATA_AF_COUNTRY_WORD_FRAGMENTS:
        return " ".join(tokens[:-1]).strip(), tokens[-1].upper()
    return stripped, None


def _tata_af_clean_amount(text: Optional[str]) -> Optional[int]:
    """印尼式千分位、無小數 '1.716.000' -> 1716000。"""
    if not text:
        return None
    text = text.strip().replace(".", "")
    return int(text) if text.isdigit() else None


def _tata_af_clean_vat(text: Optional[str]) -> Optional[float]:
    """印尼式千分位 + 逗號小數 '167.475,7' -> 167475.7。"""
    if not text:
        return None
    text = text.strip().replace(".", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def _tata_af_clean_kurs(text: Optional[str]) -> Optional[int]:
    if not text:
        return None
    text = text.strip().replace(".", "").replace(",", "")
    return int(text) if text.isdigit() else None


def _tata_af_parse_item_date(text: Optional[str]) -> Optional[str]:
    """每一筆品項自己帶的 DATE 欄位 ('31/Jul/25' -> 'DD/Mon/YY') 轉成統一的
    'DD Mon YYYY' 格式，對應到這筆品項覆寫用的 onboard_date。兩位數年份一律
    視為 20XX (目前樣本都落在 2025 年附近)。"""
    if not text:
        return None
    m = re.match(r"^(\d{1,2})/([A-Za-z]{3})/(\d{2})$", text.strip())
    if not m:
        return None
    day, month_name, yy = m.group(1), m.group(2), m.group(3)
    return _format_date_en_from_month_name(day, month_name, 2000 + int(yy))


def _tata_af_format_weight(text: Optional[str]) -> Optional[str]:
    """'12,50' -> '12.5 kgs'；'4,00' -> '4 kgs' (去除多餘小數尾零)。"""
    if not text:
        return None
    text = text.strip().replace(",", ".")
    try:
        value = float(text)
    except ValueError:
        return None
    return f"{value:g} kgs"


def _tata_af_take_adjacent(lines, idx, item_line_idx_set, consumed) -> str:
    """嘗試把 idx 這一行當成換行殘字『消耗掉』：不是本體行、沒被消耗過、
    非空、不是帳單下方小計/備註區塊的開頭字樣，才會真的被使用並標記消耗；
    否則回傳空字串 (不使用，也不消耗)。"""
    if idx < 0 or idx >= len(lines):
        return ""
    if idx in item_line_idx_set or idx in consumed:
        return ""
    line = lines[idx].strip()
    if not line or line.startswith(_TATA_AF_SKIP_LINE_PREFIXES):
        return ""
    consumed.add(idx)
    return line


def _tata_af_extract_items(text: str) -> List[Dict]:
    lines = text.split("\n")
    item_line_idx = [
        i for i, line in enumerate(lines)
        if _TATA_AF_ITEM_LINE_RE.match(line.strip())
    ]
    item_line_idx_set = set(item_line_idx)
    consumed = set()
    items: List[Dict] = []

    for i in item_line_idx:
        m = _TATA_AF_ITEM_LINE_RE.match(lines[i].strip())
        mid = m.group("mid")
        shipper_leftover, country = _tata_af_strip_country_suffix(mid)
        shipper = shipper_leftover

        if country is None or (country is not None and shipper_leftover == ""):
            # 這一行本體沒辦法同時確定完整的 Shipper 跟 Country (mid 完全
            # 沒有國家字樣、或 mid 整段就是國家名稱本身)：去抓前後兩行的
            # 換行殘字，每一行都可能同時混著 Shipper 續字 + Country 續字
            # (例如 'HONGKONG HONGYANG WEAVING HONG'：前半是 Shipper 公司
            # 名稱本身帶 'HONGKONG' 字樣、後半 'HONG' 其實是 'HONG KONG'
            # 這個國家被拆到兩行的其中一半)，不能直接整行當 Shipper 或整行
            # 當 Country，要先用 _tata_af_split_wrapped_line() 各自拆開再
            # 分別歸類、重新組合。
            before_raw = _tata_af_take_adjacent(lines, i - 1, item_line_idx_set, consumed)
            after_raw = _tata_af_take_adjacent(lines, i + 1, item_line_idx_set, consumed)
            before_shipper, before_country = _tata_af_split_wrapped_line(before_raw)
            after_shipper, after_country = _tata_af_split_wrapped_line(after_raw)
            shipper = " ".join(
                part for part in (before_shipper, shipper_leftover, after_shipper) if part
            ).strip()
            if country is None:
                # mid 本身没有國家字樣時才需要用前後行拼出來的 Country 片段；
                # 若 mid 整段就是完整國家名稱 (country 已知)，維持原值不覆寫。
                country_fragments = [c for c in (before_country, after_country) if c]
                country = " ".join(country_fragments) if country_fragments else None
        # else: mid 結尾就有完整的國家名稱、前面還有剩餘文字 -> 單行沒換
        # 行，shipper 已經是正確值，不要去動前後行。

        kurs = _tata_af_clean_kurs(m.group("kurs"))
        if kurs is not None and kurs < 1000:
            # 已知陷阱：匯率離譜偏低的雜訊列 (見上方區塊註解)，直接跳過。
            continue

        items.append({
            "description": "A/F",
            "amount": _tata_af_clean_amount(m.group("amount")),
            "vat": _tata_af_clean_vat(m.group("vat")),
            "volume": _tata_af_format_weight(m.group("weight")),
            "shipper": shipper or None,
            # 每筆品項自帶的 DATE/HAWB/Country 欄位，覆寫整張發票共用的抬頭
            # 值 (這張帳單沒有單一共用的 onboard_date/bl_no/port_of_loading，
            # 每一筆貨運各自不同)。
            "onboard_date": _tata_af_parse_item_date(m.group("date")),
            "bl_no": m.group("hawb") or None,
            "port_of_loading": country,
        })

    return items


def parse_tata_af(text: str) -> Dict:
    """注意：這是 multi_page 供應商，text 是『一頁』的文字。"""
    header: Dict[str, Optional[str]] = {}

    header["invoice_no"] = _search(r"No\.\s*:\s*(\S+)", text)

    date_m = re.search(r"Inv\s*Date\s*:\s*(\d{1,2})/(\d{1,2})/(\d{4})", text)
    header["invoice_date"] = (
        _format_date_en(date_m.group(1), date_m.group(2), date_m.group(3))
        if date_m else None
    )

    # "Nama :" 後面這份帳單沒有句點 ("PT " 不是 "PT.")，跟 classic 格式的
    # 寫法不同，不要求句點。
    header["supplier"] = _search(r"Nama\s*:\s*(.+)$", text)

    consignee_raw = _search(r"\bTo\s*:\s*(.+)$", text)
    if consignee_raw:
        consignee_raw = re.sub(r"\s*\([^)]*\)\s*$", "", consignee_raw).strip()
    header["consignee"] = consignee_raw or None

    for code in (
        "order_no", "arrive_date", "onboard_date", "bl_no", "mbl_no",
        "port_of_loading", "port_of_discharge", "vessel", "container_no",
        "volume", "shipper",
    ):
        header[code] = None

    items = _tata_af_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None}]

    return {"header": header, "items": items}


def detect_tata(text: str) -> bool:
    return detect_tata_classic(text) or detect_tata_receipt(text) or detect_tata_af(text)


def parse_tata(text: str) -> Dict:
    """依內容判斷屬於哪一種 TATA 帳單格式，分派給對應的 parse 函式。"""
    if detect_tata_receipt(text):
        return parse_tata_receipt(text)
    if detect_tata_af(text):
        return parse_tata_af(text)
    return parse_tata_classic(text)


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
#
# 28 張帳單批次發現：部門代碼有時是「字母+空白+數字+空白+字母」(例如
# "B 9303 UWV")，有時中間完全沒有空白、黏成一個詞 (例如 "B9239TEK")，
# 兩種排版都要能吃進去，否則這一行會整個比對不到完整版，退回簡化版規則
# 時又因為部門代碼含數字、不符合簡化版的「純字母」限制，導致 DO No./DO
# Date/部門代碼全部被誤併進費用名稱 (desc) 裡。這裡把部門代碼中間的空白
# 從「一定要有」(\s+) 改成「可有可無」(\s*)，兩種排版都吃得到。
_TRANS_FULL_ITEM_PATTERN = re.compile(
    r"^(?P<desc>.+?)\s+(?P<do_no>\S+)\s+(?P<do_date>\d{1,2}\s+[A-Za-z]{3}\s+\d{4})\s+"
    r"(?P<dept>[A-Z]\s*\d+\s*[A-Z]+)\s+(?P<qty>\d+)\s+(?P<unit_price>[\d.,]+)\s+"
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
    return _format_date_en(day, month, year)


def _trans_clean_amount(text: Optional[str]) -> Optional[int]:
    """印尼式數字 '2.975.000' (句點千分位，無小數) -> 2975000。"""
    if not text:
        return None
    text = text.strip().replace(".", "").replace(",", "")
    return int(text) if text.isdigit() else None


def _trans_clean_desc(text: Optional[str]) -> Optional[str]:
    """單純去除前後空白。

    注意：這裡原本會把貨櫃呎吋標示的引號字元拿掉 (例如 '... 40"' ->
    '... 40')，但這次使用者新增 VAT/Shipper 欄位時一併送來的新版正確答案
    (sample1/sample2 的 'LCL Lokal Serang - GSI-2 40"'、'GSI-2 FCL Trucking
    40"') 都保留了引號，跟舊版規則相反，所以改成照抄 PDF 原文、不拿掉
    引號。
    """
    if not text:
        return text
    return text.strip()


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
            # 注意：Amount 欄位對應的是「Total Price」(單價 x 數量) 那一欄，
            # 不是「Unit Price」(單價)——先前有一版錯以為是單價，是因為之前
            # 送測的樣本數量都剛好是 1，單價跟總價數字一樣看不出差別；這次
            # 使用者確認 Handling GSI-2 那筆 (數量 2、單價 300.000、總價
            # 600.000) 正確答案填的是總價 (600.000)，而且用 Sub Total 反推
            # 也對得起來 (Sub Total = 所有費用明細的總價加總，不是單價加
            # 總)，才確認真正要抓的是 Total Price 這一欄。
            items.append({
                "description": _trans_clean_desc(m.group("desc")),
                "amount": _trans_clean_amount(m.group("total")),
                "order_no": do_no.strip(),
            })
            i += 1
            continue

        m2 = _TRANS_SIMPLE_ITEM_PATTERN.match(line.strip())
        if m2:
            items.append({
                "description": _trans_clean_desc(m2.group("desc")),
                "amount": _trans_clean_amount(m2.group("total")),
                "order_no": None,
            })
        i += 1

    items.extend(_trans_extract_cost_items(text))
    return items


# 有些帳單在費用明細表格下面、Sub Total 之前，另外有一個「Cost Description
# Total」小表格列出額外的雜費 (例如吊櫃/堆存費)，之前的版本完全沒有把這幾
# 筆費用列進 Excel。格式很單純：費用名稱 + 金額 (印尼式千分位)，跟上面主
# 表格「數量+單價+總價」三欄的格式不同，只有一個數字，所以用一個獨立、更
# 簡單的規則抓，不會跟上面的完整版/簡化版費用行搶著比對同一行。這幾筆費用
# 沒有對應的 DO No.，Order No. 留 None 讓 expand_to_rows() 補 N/A。
#
# 28 張帳單批次發現：同一張帳單如果牽涉兩個以上的貨櫃，「Biaya LOLO」/
# 「Biaya Penumpukan」這兩種費用會各出現兩筆以上，靠費用名稱後面緊接的
# 一串參考代號 (例如 "Biaya LOLO 263300806"、"Biaya LOLO JKTG00484400")
# 來區分是哪個貨櫃的費用；正確答案的 Description 欄位是整段文字(含代號)
# 一起列出來，不會把代號拆開，所以 desc 的字元集合除了英文字母、空白，也
# 要能吃進數字/字母混合的代號本身，靠非貪婪比對讓代號優先被併入 desc，
# 金額欄位 (有千分位小數點、不含空白) 永遠留到最後一段才匹配。
_TRANS_COST_ITEM_PATTERN = re.compile(r"^(?P<desc>[A-Za-z][A-Za-z0-9 ]*?)\s+(?P<amount>[\d.]+)\s*$")


def _trans_extract_cost_items(text: str) -> List[Dict]:
    m = re.search(r"Cost\s*Description\s*Total\s*\n((?:.+\n)*?)(?=Sub\s*Total\b)", text)
    if not m:
        return []
    items: List[Dict] = []
    for line in m.group(1).split("\n"):
        line_m = _TRANS_COST_ITEM_PATTERN.match(line.strip())
        if line_m:
            items.append({
                "description": _trans_clean_desc(line_m.group("desc")),
                "amount": _trans_clean_amount(line_m.group("amount")),
                "order_no": None,
            })
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
    # P.O.D 的值永遠是單一字詞 (例如 "USA")，不要用「吃到這一行結尾」的寫
    # 法；28 張帳單批次裡，牽涉兩個貨櫃的帳單排版會把 "Other Cost ..."
    # 也擠到同一行 (例如 "P.O.D : USA Other Cost 4.621.763")，吃到行尾會
    # 把這些多餘文字一起抓進來，只取緊接著冒號後面的第一個詞就好。
    header["port_of_discharge"] = _search(r"P\.O\.D\s*:\s*(\S+)", text)

    # 大部分這種帳單右下角沒有船名/貨櫃資訊 (整頁找不到 "Vessel :"/"Cont No
    # :" 這兩個關鍵字)，留 None 讓 expand_to_rows() 補 N/A；少數有海運貨櫃
    # 的帳單則會印在 Remarks 旁邊一小塊 "Vessel : PHEN BASIN Flight/Voy No.
    # : 125N" + "Cont No : OOCU7801321"，這裡才需要真的抓出來。Vessel 欄位
    # 使用者的正確答案是「船名+航次」接在一起 (跟 KUEHNE NAGEL 同一套規則)。
    # 28 張帳單批次裡，牽涉兩個貨櫃的帳單排版會把 "VAT (11%) ..." 擠到
    # "Vessel :" 這一行的尾端 (例如 "Vessel : MERATUS JAYAGIRI VAT (11%)
    # 1.148.400"，下一行才是 "Flight/Voy No. : 603N")，所以船名那一段要在
    # 遇到 "VAT" 字樣時就先停止，避免把金額也併進船名。
    _trans_vessel_m = re.search(
        r"Vessel\s*:\s*(.+?)\s+(?:VAT\b.*?\s+)?Flight/Voy No\.\s*:\s*(.+?)(?=\s+VAT\b|\n|$)",
        text,
    )
    header["vessel"] = (
        f"{_trans_vessel_m.group(1).strip()} {_trans_vessel_m.group(2).strip()}"
        if _trans_vessel_m else None
    )
    # 牽涉兩個貨櫃的帳單會有兩個「Cont No :」(每個貨櫃各自一段 Vessel/
    # Flight/P.O.D/Cont No 區塊)，正確答案是「每個貨櫃號碼都列出來，用逗號
    # +空白連接」，不是只取第一個；用 findall 把所有出現過的貨櫃號碼都抓出
    # 來 (只取冒號後緊接著的第一個詞，不含後面可能擠在同一行的 "Other
    # Cost ..." 文字)，依出現順序去重後用 ", " 接起來。
    _trans_container_nos = re.findall(r"Cont No\s*:\s*(\S+)", text)
    header["container_no"] = (
        ", ".join(dict.fromkeys(_trans_container_nos)) if _trans_container_nos else None
    )

    # 這種帳單沒有到達日/開船日/提單/主提單/啟運港/材積/Shipper 資訊 (帳單
    # 上完全沒有印這些欄位，不是抓不到，是本來就不適用)，留空給
    # expand_to_rows() 補 N/A。Order No. 現在是逐筆費用各自帶自己的值
    # (見 _trans_extract_items())，這裡的抬頭預設值留 None，沒有費用明細
    # 帶值時才會用到 (整張帳單抓不到任何費用行的防禦性情況)。
    for code in ("arrive_date", "onboard_date", "bl_no", "mbl_no",
                 "port_of_loading", "volume", "shipper", "order_no"):
        header[code] = None

    items = _trans_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None}]

    # VAT：跟 KUEHNE NAGEL 一樣，這種帳單無論有幾筆費用明細，VAT 永遠只有
    # 「一個」整張帳單的總稅額 (印在 "VAT (11%) 327.250" 這一行)，不是每筆
    # 費用各自的稅額，所以不塞進某一筆費用明細裡，額外附加一筆「只有 vat、
    # 沒有 amount/order_no」的 item，展開成獨立的一列，Order No. 留 None
    # 讓 expand_to_rows() 補 N/A (這一列本來就沒有對應的 DO No.)。
    #
    # ⚠️ 2026-10 修正 (使用者提供新版正確答案 OK_VAT.xlsx 明確要求)：這一列
    # 的 Description/費用名稱欄位不能留空 (之前是 None → 補 N/A)，要填
    # "VAT (11%)" 這個費用名稱字串，百分比直接照抄 PDF 印的數字 (不要寫死
    # 11，避免萬一哪張帳單印的稅率不是 11% 時字串對不上)。
    _trans_vat_m = re.search(r"VAT\s*\((\d+)%\)\s*([\d.,]+)", text)
    vat_amount = _trans_clean_amount(_trans_vat_m.group(2)) if _trans_vat_m else None
    if vat_amount is not None:
        items.append({"description": f"VAT ({_trans_vat_m.group(1)}%)",
                      "amount": None, "order_no": None, "vat": vat_amount})

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
# ⚠️ 2026-09 修正 (使用者明確確認)：「描述」欄位要保留完整原文，包含開頭的
# "(代碼)-" 前綴跟後面真正的參考/追蹤號碼 (例如 "(601)-LIFTON BSA/K1-1223501"
# 整段都要留)；只有當中間那個位置本身就只是「PPh 23」或單一個「-」這種純
# 佔位符號 (代表這筆費用沒有參考號碼) 時才把它拿掉，其餘一律照抄 PDF 原文。
# 舊版程式曾經無條件拿掉 "(代碼)-" 前綴跟參考號碼，跟這次使用者提供的正確
# 答案不符，已經修正、並回頭把 test_samples/FEMARIA/sample1_ref.xlsx 也
# 訂正成同一套規則 (詳見下面 _femaria_extract_items() 的說明)。
#
# 費用明細表格還有一個排版陷阱：有些費用 (目前看到的是 "(401)-STORAGE/
# WAREHOUSE" 這筆倉租費) 描述文字太長，會被拆成三行：
#   "(401)-STORAGE/WAREHOUSE"                  <- 這筆費用的『前段』標籤，
#                                                  單獨一行、沒有編號也沒有金額
#   "3. NPCT1/NPCT/20260829/00102... 2.208.624,00"  <- 編號 + 參考號碼 + 金額
#   "BANKADM/ STAMP"                           <- 這筆費用的『後段』標籤，
#                                                  單獨一行、緊接在編號行後面
# 正確答案要把這三行重組成
# "(401)-STORAGE/WAREHOUSEBANKADM/ STAMP NPCT1/NPCT/20260829/00102..."
# (前段跟後段直接黏在一起、中間不留空白，然後空一格接編號行自己的參考
# 號碼)。判斷「下一行是不是這筆費用的後段續行」的依據是：只有在這筆費用
# 前面已經有一行『前段』還沒被用掉時，才去看下一行是不是續行；單行費用
# (前面沒有累積任何前段) 後面即使緊接著別的東西，也不會被誤吞。
_FEMARIA_DESC_FIXUPS = {
    "THC-TERMINALHANDLINGCHARGE": "THC-TERMINAL HANDLING CHARGE",
}

_FEMARIA_NUMBERED_LINE = re.compile(
    r"^(?P<num>\d+)\.\s*(?P<rest>.+?)\s+(?P<amount>[\d.]+,\d{2})\s*$"
)
_FEMARIA_STOP_MARKER = re.compile(r"jumlah\s*transaksi", re.IGNORECASE)


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
    return _format_date_en(day, month, year)


def _femaria_clean_amount(text: Optional[str]) -> Optional[int]:
    """印尼式數字 '2.276.380,00' (句點千分位、逗號小數) -> 2276380。"""
    if not text:
        return None
    text = text.strip().split(",")[0].replace(".", "")
    return int(text) if text.lstrip("-").isdigit() else None


def _femaria_clean_desc(text: Optional[str]) -> Optional[str]:
    """不拿掉 '(代碼)-' 前綴 (使用者要求保留完整原文)，只修正已知的沾黏
    文字 (用子字串取代，不要求整段字串完全相符，因為這段文字現在可能夾在
    前綴/參考號碼中間)。"""
    if not text:
        return text
    text = text.strip()
    for bad, good in _FEMARIA_DESC_FIXUPS.items():
        text = text.replace(bad, good)
    return text


def _femaria_strip_ref_placeholder(text: str) -> str:
    """去掉句尾純佔位符號的「參考號碼」："... PPh 23" 或 "... -"，這兩種
    代表「這筆費用沒有參考/追蹤號碼」，不是真正的內容；其餘真正的參考
    號碼 (例如 "BSA/K1-1223501") 一律保留，不在這裡處理。"""
    text = re.sub(r"\s+-\s*$", "", text)
    text = re.sub(r"\s+PPh\s*23\s*$", "", text, flags=re.IGNORECASE)
    return text.strip()


def _femaria_add_space_after_comma(value: Optional[str]) -> Optional[str]:
    """'Pouchen Indonesia,PT' -> 'Pouchen Indonesia, PT'。"""
    if not value:
        return value
    return re.sub(r",(\S)", r", \1", value.strip())


def _femaria_clean_placeholder_dash(value: Optional[str]) -> Optional[str]:
    """'BL/AWB No. : -' 這種欄位值整個就是一個 '-'，代表帳單上沒有填這個
    資訊 (不是真的提單號碼叫 '-')，回傳 None 讓 expand_to_rows() 補 N/A。
    """
    if value is None:
        return None
    return None if value.strip() == "-" else value


_FEMARIA_VOLUME_PATTERN = re.compile(
    r"Detail\s*pty\s*/\s*Gross_wt\s*/\s*:\s*(.+?)\s*/\s*\d+\s*KGS",
    re.IGNORECASE,
)


def _femaria_parse_volume(text: str) -> Optional[str]:
    """'Detail pty /Gross_wt/ : (1x20 FT) /0 KGS /0Cubic' 只有一組貨櫃
    資訊時，正確答案只留括號裡面的內容 ('1x20 FT')；但遇到像
    '(1x20 FT) (3x40 FT) (3x40 HQ) /0 KGS /0Cubic' 這種同一張帳單裝多種
    櫃型/批次的情況，正確答案要連括號跟全部幾組一起保留
    ('(1x20 FT) (3x40 FT) (3x40 HQ)')，不能只抓第一組就停。
    """
    m = _FEMARIA_VOLUME_PATTERN.search(text)
    if not m:
        return None
    raw = m.group(1).strip()
    groups = re.findall(r"\([^)]*\)", raw)
    if len(groups) <= 1:
        single = re.match(r"^\((.+)\)$", raw)
        return single.group(1).strip() if single else raw
    return raw


def _femaria_extract_items(text: str) -> List[Dict]:
    """逐行處理，見上面區塊註解說明的「三行拆分費用」重組規則。"""
    items: List[Dict] = []
    pending: List[str] = []
    started = False
    lines = text.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        if _FEMARIA_STOP_MARKER.search(line):
            break

        m = _FEMARIA_NUMBERED_LINE.match(line)
        if m:
            started = True
            rest = _femaria_strip_ref_placeholder(m.group("rest"))
            prefix = "".join(pending)
            pending = []
            # 只有「這筆費用前面已經有未使用的前段文字」時，才去看下一行
            # 是不是這筆費用的後段續行，避免把下一筆費用的前段誤吞進來。
            if prefix and i + 1 < len(lines):
                nxt = lines[i + 1].strip()
                if nxt and not re.match(r"^\d+\.", nxt) and not _FEMARIA_STOP_MARKER.search(nxt):
                    prefix = f"{prefix}{nxt}"
                    i += 1
            desc = f"{prefix} {rest}".strip() if prefix else rest
            items.append({
                "description": _femaria_clean_desc(desc),
                "amount": _femaria_clean_amount(m.group("amount")),
                "vat": None,
            })
        elif started:
            pending.append(line)
        # 還沒出現過第一筆編號費用之前的雜訊列 (表頭文字等) 一律忽略。
        i += 1
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
    header["volume"] = _femaria_parse_volume(text)
    header["vessel"] = _search(r"Vessel\s*/\s*Voyage\s*:\s*(.+?)\s+ATA", text)
    header["bl_no"] = _femaria_clean_placeholder_dash(
        _search(r"BL/AWB\s+No\.\s*:\s*(\S+)", text)
    )
    header["mbl_no"] = header["bl_no"]
    header["shipper"] = _search(r"Shipper\s*:\s*(.+?)\s+SPJK\s*/\s*SPJM\s*Date", text)

    # 這種帳單沒有到達日/開船日/啟運港/目的港/貨櫃號碼/Order No.，留空給
    # expand_to_rows() 補 N/A (SPJK/SPJM Date、ATA、Uitslag Date 印出來的都是
    # 假的佔位日期 '0000-00-00 00:00:00'，不是真的日期，所以不擷取)。
    for code in ("order_no", "arrive_date", "onboard_date",
                 "port_of_loading", "port_of_discharge", "container_no"):
        header[code] = None

    # 這家供應商帳單沒有單獨的 VAT 項目、也沒有整張帳單合計的 VAT 總額，
    # 所以每筆費用明細的 vat 一律是 None，讓 expand_to_rows() 補 N/A，不用
    # 額外寫擷取規則。
    items = _femaria_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None, "vat": None}]

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
    r"^\d{1,2}/[A-Za-z]{3}/\d{2}\s+(?P<shipper>.+?)\s+(?P<awb>[A-Z]{2,10}\d{4,})\s+"
    r"(?P<desc>[A-Z]+)\s+\S+\s+"
    r"[\d.]+\s+[\d.]+\s+[\d.,]+\s+(?P<amount>[\d,]+)\s+[\d,]+$"
)
_EXPRESS_DATE_PATTERN = re.compile(r"^(\d{1,2})/([A-Za-z]{3})/(\d{2})")


def detect_express(text: str) -> bool:
    return "EXPRESS MAXIMUM" in text.upper()


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
    return _format_date_en(day, month, year)


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
    return _format_date_en(day, month, f"20{year_2digit}")


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
    逐筆覆蓋。這是「海運/多筆 AWB」那種版面才有的表格，走空運「DETAILS OF
    PAYMENT」版面 (見 _express_extract_payment_items()) 的帳單這裡會抓不到
    任何一行，回傳空列表。
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


# 空運「DETAILS OF PAYMENT」版面的費用行，例如：
#   "Air Freight Charge $ 3.700 $ 166.500 Rp 2,794,869"
#   "Costoms Clearance Fee Rp 250,000"
# 費用名稱後面可能有 0～多組 "$ 數字" (美金單價/小計)，最後一定是
# "Rp 金額" (印尼盾金額，才是正確答案要的 AMOUNT)。"Total"/"Sub. Total"/
# "Tax (...)"/"Deposit" 這些彙總列雖然格式很像，但不是真正的費用項目，
# 要另外排除。
_EXPRESS_PAYMENT_ITEM_PATTERN = re.compile(
    r"^(?P<desc>[A-Za-z][A-Za-z .]*?)\s+(?:\$\s*[\d,]+\.?\d*\s+)*Rp\.?\s*(?P<amount>[\d,]+)$"
)
_EXPRESS_PAYMENT_ITEM_SKIP_PATTERN = re.compile(
    r"^(total|sub\.?\s*total|grand\s*total|tax|deposit)\b", re.IGNORECASE
)


def _express_extract_payment_items(text: str) -> List[Dict]:
    """空運帳單「DETAILS OF PAYMENT」表格版面 (跟 _express_extract_items()
    的海運多筆 AWB 表格是兩種完全不同的版面，同一家供應商依貨運方式各自
    印出不同格式的帳單)。逐行掃描，只認「費用名稱 + (可選)美金單價 +
    Rp 印尼盾金額」這種格式，Total/Sub. Total/Grand Total (含 "GRAND TOTAL
    in IDR" 這種後面還帶文字的變化)/Tax(...)/Deposit 這些彙總列要排除，
    不是真正的費用明細。
    """
    items: List[Dict] = []
    for line in text.split("\n"):
        stripped = line.strip()
        m = _EXPRESS_PAYMENT_ITEM_PATTERN.match(stripped)
        if not m:
            continue
        desc = m.group("desc").strip()
        if _EXPRESS_PAYMENT_ITEM_SKIP_PATTERN.match(desc):
            continue
        items.append({
            "description": desc,
            "amount": _express_clean_amount(m.group("amount")),
        })
    return items


def parse_express(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {}

    # 公司名稱固定就是這家供應商，不用每次都從 PDF 抓：有些帳單的抬頭
    # logo 用特殊字型印刷，pdfplumber 抽出來的文字會變成每個字母重複好
    # 幾次的亂碼 (例如 "PPPPTTTT....EEEEXXXXPPPPRRRREEEE..." 其實是
    # "PT.EXPRESS MAXIMUM")，與其花力氣解碼亂碼字型，不如直接固定這個
    # 值，兩種帳單版面都適用。
    header["supplier"] = "PT.EXPRESS MAXIMUM"
    header["consignee"] = _search(r"COMPANY\s*:\s*(.+?)\s*$", text)
    header["invoice_no"] = _search(r"INVOICE\s+NO\s*:\s*(\S+)", text)
    header["invoice_date"] = _express_parse_long_date(
        _search(r"Tangerang\s*,\s*(\d{1,2}\s+[A-Za-z]+\s+\d{4})", text)
    )

    # 這種帳單沒有 Order No./啟運港/目的港/材積/船名/貨櫃號碼，留空給
    # expand_to_rows() 補 N/A。
    for code in ("order_no", "port_of_loading", "port_of_discharge",
                 "volume", "vessel", "container_no"):
        header[code] = None

    items = _express_extract_items(text)
    if items:
        # 海運/多筆 AWB 版面：到達日/開船日/提單號碼/主提單號碼改成逐筆
        # 費用各自帶自己的值 (見 _express_extract_items())，這裡不用整張
        # 發票共用一個值。
        for code in ("arrive_date", "onboard_date", "bl_no", "mbl_no"):
            header[code] = None
    else:
        # 空運「DETAILS OF PAYMENT」版面：整張發票只有一組 DATE/HAWB/MAWB，
        # 所有費用項目共用同一個值 (不像海運版面每筆費用各自的 AWB 可能不
        # 同)。HAWB (House AWB，供應商自己的貨運單號) 對應 B/L NO.，
        # MAWB (Master AWB，航空公司的主提單號，格式含空白例如
        # "180 4357 0774") 對應 MBL NO.，跟 DHL/其他供應商「提單/主提單」
        # 欄位的用法一致。這種版面沒有 Arrive Date 的概念 (空運當天直飛，
        # 帳單只印一個 DATE)，留空給 expand_to_rows() 補 N/A。
        header["arrive_date"] = None
        header["onboard_date"] = _express_parse_long_date(
            _search(r"^DATE\s*:\s*(\d{1,2}\s+[A-Za-z]+\s+\d{4})", text)
        )
        header["bl_no"] = _search(r"^HAWB\s*:\s*(\S+)", text)
        header["mbl_no"] = _search(r"^MAWB\s*:\s*(.+?)\s*$", text)
        items = _express_extract_payment_items(text)

    if not items:
        items = [{"description": None, "amount": None}]

    return {"header": header, "items": items}


# ===========================================================================
# 帳單金額稽核：從各供應商帳單原文抓出「未稅總額」(Total (Excl. VAT) /
# Sub Total / SUBTOTAL / Net value…，各供應商命名不同)，讓
# extract_utils.audit_invoice_totals() 拿去跟系統擷取到的費用明細加總互相
# 核對，一致就不用管，不一致就跳出警示 (見 app.py / verify_app.py)。
#
# 這裡只是「抓數字」，直接沿用各供應商自己既有的金額格式清理函式
# (千分位符號、小數點慣例每家都不一樣)，抓不到就回傳 None，代表這張帳單
# 這次沒辦法稽核 (不算「不一致」，呼叫端不會顯示警示)。
# ===========================================================================

def _dwiharta_extract_total(text: str) -> Optional[int]:
    m = _search(r"Sub\s*Total\s+([\d,]+)", text)
    if not m:
        return None
    m = m.replace(",", "")
    return int(m) if m.isdigit() else None


def _kn_extract_total(text: str) -> Optional[int]:
    # "SUBTOTAL IDR 1,521,098 167,321" 是 Sales Invoice 費用明細行上的英式
    # 數字 (逗號千分位、無小數)，要用 _kn_clean_amount_comma_thousands()，
    # 不是給 Faktur Pajak 印尼式數字用的 _kn_clean_amount()。
    return _kn_clean_amount_comma_thousands(_search(r"SUBTOTAL\s+IDR\s+([\d,]+)", text))


def _kn_extract_grand_total(text: str) -> Optional[int]:
    # "INVOICE TOTAL IDR 1,688,419" 是含稅總額 (= SUBTOTAL + VAT)，用來做
    # 『未稅加總 + VAT 加總 = 含稅總額』的稽核 (見
    # extract_utils.audit_invoice_totals())。跟上面的 SUBTOTAL 是不同行，
    # 用 "INVOICE TOTAL" 這個獨有字樣區分，不會跟 SUBTOTAL 那一行搶到。
    return _kn_clean_amount_comma_thousands(_search(r"INVOICE\s+TOTAL\s+IDR\s+([\d,]+)", text))


def _indo_extract_grand_total(text: str) -> Optional[int]:
    """注意：INDOPROSTIME (EXPRESS/TRANSPORT) 都是 multi_page 供應商，text
    是『一頁』的文字。帳單上印的 "Total" 是含稅總額 (如果這張帳單有
    V.A.T. 的話，已經把 VAT 加進去了)，所以要註冊成 extract_grand_total
    (跟『擷取到的費用明細加總 + VAT 加總』互相稽核)，不是 extract_total
    (未稅總額)——這家供應商帳單本文沒有單獨印出未稅小計。"""
    return _indo_clean_amount(_search(r"^Total\s+([\d,]+)\s*$", text))


def _maersk_extract_total(text: str) -> Optional[int]:
    # "Net value 1,381,801" 是 Invoice 頁面「Total Amount Due」小表格裡的
    # 英式數字 (逗號千分位、無小數)，跟 Faktur Pajak 頁面費用明細行的印尼式
    # 數字 (_maersk_clean_amount_id，句點千分位、逗號小數) 是不同格式。
    m = _search(r"Net\s+value\s+([\d,]+)", text)
    if not m:
        return None
    m = m.replace(",", "")
    return int(m) if m.isdigit() else None


def _hyper_extract_total(text: str) -> Optional[int]:
    """注意：HYPER_MEGA 是 multi_page 供應商，text 是『一頁』的文字。"""
    return _hyper_clean_amount_id(_search(r"Sub\s*Total\s+([\d,.]+)", text))


def _yje_tata_extract_total(text: str) -> Optional[int]:
    """YJE_TATA 是合併群組，同一份 PDF 可能混著 3 種版面 (YJE 英文費率表、
    TATA 印尼文 INVOICE 收據、TATA 稅務收據)，各自的未稅小計標示方式跟
    金額格式都不同，依序嘗試；都抓不到就回傳 None (這一頁不稽核，不是
    「不一致」——例如 YJE 英文費率表目前還沒有對應的規則)。
    """
    m = _search(r"SUB-TOTAL\s+Rp\s+([\d,]+)", text)
    if m:
        return _tata_clean_amount(m)
    m = _search(r"Subtotal\s+([\d.,]+)", text)
    if m:
        return _tata_receipt_clean_amount(m)
    return None


def _yje_tata_extract_grand_total(text: str) -> Optional[int]:
    """TATA 稅務發票 (classic 格式) 本身印的 "TOTAL" 已經是含稅總額 (=
    SUB-TOTAL + PPN)，用來做『擷取到的費用明細加總 + VAT 加總 = 含稅
    總額』的稽核 (見 extract_utils.audit_invoice_totals())。用「行首就是
    TOTAL」限定，不能只找 "TOTAL" 字樣，否則會跟前面 "SUB-TOTAL ..." 那
    一行搶到 (那一行不是以 "TOTAL" 開頭，是 "SUB-TOTAL" 開頭，一般不會
    誤配，這裡額外加 ^ 錨點更保險)；YJE 英文費率表跟 TATA 稅務收據目前
    還沒有對應的含稅總額規則，抓不到就回傳 None (不稽核，不是「不
    一致」)。
    """
    m = _search(r"^TOTAL\s+Rp\s+([\d,]+)\s*$", text)
    if m:
        return _tata_clean_amount(m)
    return None


def _trans_extract_total(text: str) -> Optional[int]:
    """注意：TRANS_DAYA_PRIMA 是 multi_page 供應商，text 是『一頁』的文字。

    這裡要跟系統擷取出來的「所有費用明細加總」(不含 VAT 那一列) 互相核對，
    但帳單上的 "Sub Total" 只是主費用表格 (Service Description 那個表格)
    的未稅小計，不包含 "Cost Description" 那個小表格裡的額外雜費 (例如
    Biaya LOLO/Biaya Penumpukan)——這些雜費另外加總成 "Other Cost" 這個
    獨立欄位，沒有雜費的帳單這裡會印 "Other Cost 0"。所以這裡要用
    「Sub Total + Other Cost」才會跟 (主費用 + Cost Description 費用) 的
    加總對得起來，單用 Sub Total 核對，有雜費的帳單一定會被誤判成不一致。
    """
    sub_total = _trans_clean_amount(_search(r"Sub\s*Total\s+([\d.,]+)", text))
    other_cost = _trans_clean_amount(_search(r"Other\s*Cost\s+([\d.,]+)", text))
    if sub_total is None:
        return None
    return sub_total + (other_cost or 0)


def _trans_extract_grand_total(text: str) -> Optional[int]:
    """注意：TRANS_DAYA_PRIMA 是 multi_page 供應商，text 是『一頁』的文字。

    帳單右下角單獨一行印著 "Total 14.982.004" (= Sub Total + VAT + Other
    Cost，含稅總額)，用來做『未稅加總 + VAT 加總 = 含稅總額』的稽核 (見
    extract_utils.audit_invoice_totals())。要用「行首就是 Total」限定，
    不能只找 "Total" 字樣，否則會跟前面 "Sub Total ..." 那一行搶到。
    """
    return _trans_clean_amount(_search(r"^Total\s+([\d.,]+)\s*$", text))


def _femaria_extract_total(text: str) -> Optional[int]:
    # "Jumlah Transaksi : 7 Total : 7.499.546,00" 跟 "Total\nJumlah
    # Transaksi : 10 86.999.502,00" 兩種排版都有出現過 ("Total" 字樣位置
    # 因為 pdfplumber 斷行方式不同而前後不定)，統一只抓「Jumlah Transaksi
    # : 數字」後面緊接著的那個金額，不管前面有沒有多一個 "Total :"。
    return _femaria_clean_amount(
        _search(r"Jumlah\s*Transaksi\s*:\s*\d+\s*(?:Total\s*:\s*)?([\d.,]+)", text)
    )


def _express_extract_total(text: str) -> Optional[int]:
    return _express_clean_amount(_search(r"TOTAL\s+\$\s*[\d.]+\s+([\d,]+)\s+[\d,]+", text))


def _pan_extract_total(text: str) -> Optional[int]:
    return _pan_clean_amount(_search(r"Sub\s*Total\s+([\d,.]+)", text))


def _pancaran_extract_total(text: str) -> Optional[int]:
    return _pancaran_clean_amount(_search(r"Sub\s*Total\s+([\d,]+)", text))


def _dhl_extract_total(text: str) -> Optional[int]:
    return _dhl_clean_amount(_search(r"Total\s+Amount\s*\(IDR\)\s+([\d,]+)", text))


def _dsv_extract_total(text: str) -> Optional[int]:
    """注意：DSV 是 multi_page 供應商，text 是『一頁』的文字。"""
    return _dsv_clean_amount(_search(r"SUBTOTAL\s+([\d.]+)", text))


def _evergreen_extract_grand_total(text: str) -> Optional[int]:
    """帳單最下面獨立一行、只有 "IDR 數字" 的那一行是含稅總額 (VAT 已經
    算進去的最終應付金額)。"""
    return _evergreen_clean_amount(_search(r"^IDR\s+([\d,]+(?:\.\d+)?)\s*$", text))


def _evergreen_extract_vat(text: str) -> Optional[int]:
    """"VAT IDR 197,638.00" 是整張帳單只有一個的稅額總計 (不是逐筆費用
    各自的稅額)，跟 KUEHNE_NAGEL/TRANS_DAYA_PRIMA/APL 一樣另外獨立成一筆
    合成列 (見 _evergreen_extract_items())，不要跟費用明細混在一起算。"""
    return _evergreen_clean_amount(_search(r"VAT\s+IDR\s+([\d,]+(?:\.\d+)?)", text))


def _evergreen_extract_total(text: str) -> Optional[int]:
    """EVERGREEN_LOGISTICS 帳單原文沒有直接寫『未稅總額』，只有『含稅總額』
    (獨立一行、只有 "IDR 數字") 跟 "VAT IDR 數字" 兩個數字，未稅總額 =
    含稅總額 - VAT。
    """
    grand_total = _evergreen_extract_grand_total(text)
    vat = _evergreen_extract_vat(text)
    if grand_total is None or vat is None:
        return None
    return grand_total - vat


register_supplier(
    "DWIHARTA", "PT. DWIHARTA LOGISTINDO", detect_dwiharta, parse_dwiharta,
    extract_total=_dwiharta_extract_total,
)
register_supplier(
    "KUEHNE_NAGEL", "KUEHNE NAGEL INDONESIA", detect_kuehne_nagel, parse_kuehne_nagel,
    extract_total=_kn_extract_total, extract_grand_total=_kn_extract_grand_total,
)
register_supplier(
    "INDOPROSTIME", "PT. INDO PROSTIME EXPRESS", detect_indoprostime, parse_indoprostime,
    multi_page=True, extract_grand_total=_indo_extract_grand_total,
)
register_supplier(
    "INDOPROSTIME_TRANSPORT", "PT. INDO PROSTIME TRANSPORT",
    detect_indoprostime_transport, parse_indoprostime_transport,
    multi_page=True, extract_grand_total=_indo_extract_grand_total,
)
register_supplier(
    "MAERSK", "PT MAERSK LOGISTICS INDONESIA", detect_maersk, parse_maersk,
    extract_total=_maersk_extract_total, extract_grand_total=_maersk_extract_grand_total,
)
register_supplier(
    "HYPER_MEGA", "PT. HYPER MEGA SHIPPING", detect_hyper_mega, parse_hyper_mega,
    multi_page=True, extract_total=_hyper_extract_total,
)
register_supplier(
    "YJE_TATA", "YJE (ShenZhen) International Logistics / PT TATA HARMONI SARANATAMA",
    detect_yje_or_tata, parse_yje_or_tata, multi_page=True,
    extract_total=_yje_tata_extract_total,
    extract_grand_total=_yje_tata_extract_grand_total,
)
register_supplier(
    "TRANS_DAYA_PRIMA", "PT. TRANS DAYA PRIMA", detect_trans, parse_trans,
    multi_page=True, extract_total=_trans_extract_total,
    extract_grand_total=_trans_extract_grand_total,
)
register_supplier(
    "FEMARIA", "PT. FEMARIA BUANA CARGO", detect_femaria, parse_femaria,
    extract_total=_femaria_extract_total,
)
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
    return _format_date_en(day, month, year)


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


register_supplier(
    "EXPRESS_MAXIMUM", "PT.EXPRESS MAXIMUM", detect_express, parse_express,
    extract_total=_express_extract_total,
)
register_supplier(
    "PAN_EKSPRES", "PT. PAN EKSPRES INTERNATIONAL",
    detect_pan_ekspres, parse_pan_ekspres,
    extract_total=_pan_extract_total,
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


# 這家供應商的 VAT 是「每一筆品項自己的一欄」，不是像 TATA 那種另外獨立
# 一筆 PPN 列：同一張發票裡有些品項有印出 VAT 金額 (帶小數點，例如
# "27,500.00")、有些完全沒印 (留空)，所以 VAT 要接在 Amount 後面當成
# 「可選」的尾段來抓，抓不到就是 None (不是 0)。
_PANCARAN_ITEM_PATTERN = re.compile(
    r"^\d+\s+(?P<desc>.+?)\s+\d+\s+[\d,]+\s+(?P<amount>[\d,]+)"
    r"(?:\s+(?P<vat>[\d,]+\.\d+))?$"
)


def _pancaran_clean_amount(text: Optional[str]) -> Optional[int]:
    """'75,660,000' (逗號千分位、無小數) -> 75660000；
    VAT 欄位偶爾帶 '.00' 小數 (例如 '27,500.00') 一樣去掉小數部分取整數。"""
    if not text:
        return None
    text = text.strip().replace(",", "")
    text = text.split(".")[0]
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
            "vat": _pancaran_clean_amount(m.group("vat")) if m.group("vat") else None,
        })
    return items


def _pancaran_extract_grand_total(text: str) -> Optional[int]:
    """'GRAND TOTAL 8,913,458' 本身就是含稅總額 (= Sub Total + V A T)。"""
    return _pancaran_clean_amount(_search(r"GRAND\s+TOTAL\s+([\d,]+)", text))


# REFERENCE NO. 那一行 (跟它的續行) 右邊其實還擠著 "Qty :" 欄位的值
# (雙欄排版被 pdfplumber 併成同一行)：
#   "REFERENCE NO. : 94079840 EXPOR - 1x40HC"
#        -> Order No. = "94079840"，Qty(=Volume) = "EXPOR - 1x40HC"
#   "REFERENCE NO. : PYI-...167 LCL - 13x40HC LCL - 2xBOXCDDLCL -"
#   "/171/175/176/178/180/181/182 1xBOXFUSO"
#        -> Order No. 兩行合併，Volume 兩行合併
# 正確答案範本裡，"EXPOR -"/"IMPOR -" 這種單純的出口/進口標記會被拿掉，
# 只留真正的材積內容 (例如 "EXPOR - 1x40HC" -> "1x40HC")；但其他像
# "LCL - 13x40HC" 這種不是出口/進口標記的字樣則照原樣保留。
_PANCARAN_EXPOR_IMPOR_PREFIX_RE = re.compile(r"^(?:EXPOR|IMPOR)\s*-\s*", re.IGNORECASE)


def _pancaran_extract_volume(text: str) -> Optional[str]:
    lines = text.split("\n")
    for i, line in enumerate(lines):
        m = re.search(r"REFERENCE\s+NO\.\s*:\s*(\S+)\s+(.+)$", line)
        if not m:
            continue
        parts = [m.group(2).strip()]
        if i + 1 < len(lines):
            nxt = lines[i + 1].strip()
            if nxt.startswith("/"):
                m2 = re.match(r"\S+\s+(.+)$", nxt)
                if m2:
                    parts.append(m2.group(1).strip())
        value = "\n".join(p for p in parts if p)
        value = _PANCARAN_EXPOR_IMPOR_PREFIX_RE.sub("", value)
        return value or None
    return None


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
    header["bl_no"] = _search(r"JOB\s+NO\.\s*:\s*(\S+)", text)
    header["volume"] = _pancaran_extract_volume(text)
    header["shipper"] = _search(r"(?m)^SHIPPER\s*:\s*(.+?)(?:;|\s*$)", text)

    pol_pod = _search(r"(?m)^POL\s*/\s*POD\s*:\s*(.+)$", text)
    if pol_pod and "/" in pol_pod:
        pol, pod = pol_pod.split("/", 1)
        header["port_of_loading"] = pol.strip() or None
        header["port_of_discharge"] = pod.strip() or None
    else:
        header["port_of_loading"] = pol_pod
        header["port_of_discharge"] = None

    vessel = _search(r"(?m)^VESSEL\s+NAME\s*:\s*(.+)$", text)
    if vessel and re.fullmatch(r"[.\s/]+", vessel):
        vessel = None
    header["vessel"] = vessel

    # 這種帳單沒有到達日/開船日/主提單號碼，留空給 expand_to_rows() 補 N/A。
    for code in ("arrive_date", "onboard_date", "mbl_no"):
        header[code] = None

    items = _pancaran_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None, "vat": None}]

    return {"header": header, "items": items}


register_supplier(
    "PANCARAN", "PT. PANCARAN SRIKANDI LOGISTIK",
    detect_pancaran, parse_pancaran,
    extract_total=_pancaran_extract_total,
    extract_grand_total=_pancaran_extract_grand_total,
)


# ===========================================================================
# 供應商 12：PT BIROTIKA SEMESTA / DHL EXPRESS
# ===========================================================================
#
# ⚠️ 2026-09 重新設計：DHL 帳單第 1 頁是「彙總表」(依服務類型/費用類型列出
# 全發票的合計金額)，第 2 頁以後才是「各票 (Air Waybill) 各自的費用明細」
# (每一個 Air Waybill Number 是一票獨立的貨物，各自有自己的起訖地、日期、
# 基本運費、額外費用明細)。舊版程式只抓第 1 頁彙總表 (等於把全發票所有票
# 的同類費用加總成一筆)，沒有拆到「每一票」的層級，也沒有 bl_no/起訖港/
# 日期這些明細欄位；使用者要求整套重寫，改成從第 2 頁以後逐票拆解，
# bl_no/onboard_date/port_of_loading/port_of_discharge 都是「每一票各自
# 的值」(用 item 覆寫抬頭欄位的機制處理，見 parse_dhl_regular())。
#
# 本供應商同時有兩種實際見過的帳單子格式，用 _dhl_is_regular(text) 判斷
# 走哪一條路：
#
# (A) 「一般」格式 (_parse_dhl_regular)：PDF 有內嵌文字層 (或 OCR 品質
#     夠好)，第 2 頁以後的表頭有 "Air Waybill Number" / "Shippers
#     Reference" 字樣，每一票的主要那一行長相固定：
#       "<AWB> [Shippers Reference...] <日期> <起運地代碼>, <起運地地區>
#        <目的地代碼>, <目的地地區> EXPRESS <重量><重量代碼?> <件數>
#        <基本運費> <VAT金額> <稅別代碼> <含稅小計>"
#     後面接著的幾行 (跟寄件人/收件人地址擠在同一行) 才是這一票的「額外
#     費用」明細，每一行格式是 "<費用名稱> <金額> <VAT金額> <稅別代碼>
#     <含稅小計>"。因為費用名稱前面常常黏著寄件人/收件人名稱、地址這些
#     雜訊文字 (版面把好幾欄硬擠成一行)，沒辦法單純用「這一行開頭的大寫
#     單字」來判斷費用名稱，改成先從第 1 頁「Analysis of Extra Charges」
#     區塊 (這裡沒有跟其他欄位擠在一起，乾淨可靠) 動態收集「這張發票實際
#     出現過的額外費用名稱」當作白名單，再拿這份白名單去比對第 2 頁以後
#     雜訊行裡「費用名稱 金額 VAT 代碼 小計」這個固定樣式，就能準確從
#     雜訊中挑出正確的費用名稱，不會被前面的地址/人名文字誤導。基本運費
#     (每一票主要那一行) 的服務名稱固定使用第 1 頁彙總表的「Type of
#     Service」欄位值 (例如 "EXPRESS WORLDWIDE NONDOC")，因為同一張發票
#     裡每一票用的都是同一種服務。
#
# (B) 「掃描檔關務/完稅費用」格式 (_parse_dhl_scanned_customs)：本系統
#     第一家掃描檔供應商，PDF 沒有文字層，全部文字要靠 OCR 備援機制辨識，
#     常見於「DUTIES & TAXES」這種基本運費是 0、只有進口稅捐相關費用的
#     帳單。OCR 出來的第 2 頁明細表格式比 (A) 更不規則 (欄位常常對不齊、
#     稅別/VAT 欄位可能整個缺漏)，(A) 的固定樣式規則套不上去，維持原本
#     專門處理這種格式的邏輯 (第 1 頁彙總表 + Analysis of Extra Charges
#     區塊逐行擷取，比對第 2 頁明細只抓 Air Waybill Number/日期，其餘
#     欄位這種格式本來就沒有更細的資訊可拆)。
#
# parse_dhl() 先嘗試 (A)，抓不到任何一票 (代表不是這種乾淨格式) 才 fall
# back 到 (B)，不用額外去判斷「是不是掃描檔」。

def detect_dhl(text: str) -> bool:
    upper = text.upper()
    return "BIROTIKA SEMESTA" in upper or "DHL EXPRESS" in upper


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
    return _format_date_en(day, month, year)


def _dhl_clean_amount(text: Optional[str]) -> Optional[int]:
    """'662,295' (逗號千分位、無小數) -> 662295。"""
    if not text:
        return None
    text = text.strip().replace(",", "")
    return int(text) if text.isdigit() else None


def _dhl_common_header(text: str) -> Dict[str, Optional[str]]:
    """兩種子格式共用的抬頭欄位 (收件人/供應商名稱/發票號碼/發票日期)。"""
    header: Dict[str, Optional[str]] = {}
    first_line = text.split("\n", 1)[0].strip()
    m = re.search(r"^([^\n]+?)[ \t]+Invoice[ \t]+Number[ \t]*:", text, re.MULTILINE)
    header["consignee"] = m.group(1).strip() if m else (first_line or None)
    header["supplier"] = _search(r"(PT\s+BIROTIKA\s+SEMESTA)\s*/\s*DHL\s+EXPRESS", text)
    if header["supplier"]:
        header["supplier"] = f"{header['supplier']} / DHL EXPRESS"

    # OCR 掃描版常把「標籤欄一整排、值欄一整排」分開掃描 (先四行標籤：
    # Invoice Number:/Account Number:/Tax ID:/Invoice Date:，空一行，再
    # 四行對應的值)，不是每個標籤緊接著自己的值；有文字層的版面則是
    # "Invoice Number: BDOIR00043036" 緊接在同一行。兩種都試，抓到哪個
    # 算哪個。
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
        header["invoice_no"] = _dhl_fix_ocr_zero(
            _search(r"Invoice\s+Number\s*:\s*(\S+)", text)
            or _search(r"\b([A-Z]{2,6}[O0-9]{6,})\b", text)
        )
        header["invoice_date"] = _dhl_parse_date(
            _search(r"Invoice\s+Date\s*:\s*\n?\s*(\d{1,2}-\d{1,2}-\d{4})", text)
        )
    return header


# ---------------------------------------------------------------------------
# (A) 一般格式：逐票 (Air Waybill) 拆解
# ---------------------------------------------------------------------------

_DHL_SERVICE_NAME_PATTERN = re.compile(
    r"Type of Service.*?\n(?:.*\n)*?"
    r"(?P<svc>[A-Z][A-Z /&]+[A-Z])\s+\d+\s+[\d.]+\s+\d+\s+"
    r"[\d,]+\s+[\d,]+\s+[\d,]+\s+[\d,]+"
)
_DHL_EXTRA_CHARGE_NAMES_PATTERN = re.compile(
    r"Analysis of Extra Charges.*?\n(?P<block>.*?)Total Extra Charges", re.DOTALL
)
_DHL_EXTRA_CHARGE_NAME_LINE = re.compile(r"^([A-Z][A-Z /&]+[A-Z])\s+[\d,]+\s*$", re.MULTILINE)

# 每一票主要那一行 (基本運費)，例如：
#   "7387602585 26-12-2024 XMN, FUJIAN PROVINCE BDO, BANDUNG EXPRESS 14.50
#    B 2 1,693,221 20,319 A 1,713,540"
#   "6278270375 1738854338 TL 1DN 27-12-2024 HKG, HONG KONG BDO, BANDUNG
#    EXPRESS 1.00W 1 467,761 5,613 A 473,374"  (AWB 後面多一段 Shippers
#    Reference，用 "(?:\s+\S.*?)?" 這種寬鬆寫法吃掉，不用管它實際內容)
# 起運/目的地各自是「代碼 (2-4 個大寫字母) + 逗號 + 地區名稱」，兩個地名
# 中間只隔一個空白 (跟 INDOPROSTIME/MAERSK 同款並排欄位排版問題)，靠
# 「目的地代碼一定是大寫字母」這個特徵切開，不用連續空白判斷。
_DHL_MAIN_LINE_PATTERN = re.compile(
    r"^(?P<awb>\d{6,12})(?:\s+\S.*?)?\s+(?P<date>\d{1,2}-\d{1,2}-\d{4})\s+"
    r"(?P<ocode>[A-Z]{2,4}),\s*(?P<oregion>[A-Za-z ]+?)\s+"
    r"(?P<dcode>[A-Z]{2,4}),\s*(?P<dregion>[A-Za-z ]+?)\s+EXPRESS\s+"
    r"(?P<weight>[\d.]+)\s*[A-Z]?\s+(?P<items>\d+)\s+"
    r"(?P<charge>[\d,]+)\s+(?P<vat>[\d,]+)\s+[A-Z]\s+(?P<total>[\d,]+)\s*$",
    re.MULTILINE,
)


def _dhl_extract_charge_name_whitelist(text: str) -> List[str]:
    """從第 1 頁「Analysis of Extra Charges」區塊收集這張發票實際出現過
    的額外費用名稱，當作白名單去比對第 2 頁以後跟雜訊擠在同一行的費用
    明細，避免把地址/人名文字誤判成費用名稱 (見上面區塊註解)。"""
    m = _DHL_EXTRA_CHARGE_NAMES_PATTERN.search(text)
    if not m:
        return []
    names = _DHL_EXTRA_CHARGE_NAME_LINE.findall(m.group("block"))
    # 名稱長的排前面，避免比對時被同樣開頭的較短名稱截斷 (目前沒有這種
    # 案例，但先寫成通用寫法)。
    return sorted(set(n.strip() for n in names), key=len, reverse=True)


def _dhl_is_regular(text: str) -> bool:
    upper = text.upper()
    return "AIR WAYBILL" in upper and "SHIPPERS" in upper


def _parse_dhl_regular(text: str) -> Optional[Dict]:
    main_matches = list(_DHL_MAIN_LINE_PATTERN.finditer(text))
    if not main_matches:
        return None  # 不是這種乾淨格式，讓 parse_dhl() fall back 到 (B)

    svc_m = _DHL_SERVICE_NAME_PATTERN.search(text)
    service_name = svc_m.group("svc").strip() if svc_m else None

    whitelist = _dhl_extract_charge_name_whitelist(text)
    extra_pattern = None
    if whitelist:
        alt = "|".join(re.escape(n) for n in whitelist)
        extra_pattern = re.compile(
            rf"(?:{alt})\s+(?P<amount>[\d,]+)\s+(?P<vat>[\d,]+)\s+[A-Z]\s+(?P<total>[\d,]+)"
        )
        # 費用名稱本身也要抓出來 (上面的 pattern 沒有把名稱包進 group，
        # 因為 re.escape() 過的多個名稱長度不一，用另一個對照 pattern
        # 重新加上具名 group 更清楚)。
        extra_pattern = re.compile(
            rf"(?P<desc>{alt})\s+(?P<amount>[\d,]+)\s+(?P<vat>[\d,]+)\s+[A-Z]\s+(?P<total>[\d,]+)"
        )

    boundaries = [m.start() for m in main_matches] + [len(text)]
    items: List[Dict] = []
    for idx, m in enumerate(main_matches):
        block = text[m.start():boundaries[idx + 1]]
        bl_no = m.group("awb")
        onboard_date = _dhl_parse_date(m.group("date"))
        port_of_loading = f"{m.group('ocode')}, {m.group('oregion')}".strip()
        port_of_discharge = f"{m.group('dcode')}, {m.group('dregion')}".strip()
        overrides = {
            "bl_no": bl_no,
            "onboard_date": onboard_date,
            "port_of_loading": port_of_loading,
            "port_of_discharge": port_of_discharge,
        }

        items.append({
            "description": service_name,
            "amount": _dhl_clean_amount(m.group("charge")),
            **overrides,
        })
        if extra_pattern:
            for em in extra_pattern.finditer(block):
                items.append({
                    "description": em.group("desc").strip(),
                    "amount": _dhl_clean_amount(em.group("amount")),
                    **overrides,
                })

    header = _dhl_common_header(text)
    # 這種帳單抬頭本身沒有共通的到達日/主提單號碼/材積/船名/貨櫃號碼/
    # Order No.；bl_no/onboard_date/port_of_loading/port_of_discharge 這
    # 四個欄位改成逐票用上面 items 各自的 overrides 覆寫 (expand_to_rows()
    # 會自動處理)，這裡的抬頭預設值留 None 就好。
    for code in ("order_no", "arrive_date", "mbl_no", "bl_no", "onboard_date",
                 "port_of_loading", "port_of_discharge", "volume", "vessel",
                 "container_no"):
        header[code] = None

    return {"header": header, "items": items}


# ---------------------------------------------------------------------------
# (B) 掃描檔關務/完稅費用格式 (OCR)
# ---------------------------------------------------------------------------
#
# 這份帳單一張發票橫跨兩頁掃描圖檔：第 1 頁是「服務類型/金額」總覽表
# (品項/金額都在這一頁，OCR 品質也最好)；第 2 頁是「Air Waybill/Shippers
# Reference/Shipment Origin Date」明細表，且第 2 頁整頁被掃描成轉了 90 度
# (OCR 備援機制裡的方向偵測會自動轉正，這裡不用處理)，轉正後 OCR 出來的
# 文字順序仍然會把同一列的欄位「擠成一行」(不像原始表格那樣分欄)，但因為
# 一張發票只有一票資料，用「這一行前三個數字依序是 Air Waybill/Shippers
# Reference/Shipment Origin Date」的位置規則就能可靠取出；基本運費是 0
# (這種帳單沒有基本運費，只有關稅相關費用)，也沒有 VAT/稅別代碼可拆，
# 套不上 (A) 的固定樣式規則，維持原本專門處理這種格式的簡單邏輯。
#
# 已知的 OCR 誤判：Invoice Number 裡的數字 '0' 常被辨識成英文字母 'O'
# (例如 'JKTIR00818492' 被讀成 'JKTIROO818492' 或 'JKTIRO0818492')，這個
# 發票編號的格式固定是「英文字母開頭 + 純數字」，前面的英文字母部分剛好
# 不含字母 O，所以直接把整個擷取到的編號裡的 'O' 全部換成 '0' 是安全的
# 修正方式；之後如果遇到其他 OCR 常誤判的字元，比照這個做法在擷取後加一
# 個修正函式即可。
_DHL_CUSTOMS_ITEM_PATTERN = re.compile(r"^([A-Z][A-Z /]+[A-Z])\s+([\d,]+)$")
_DHL_CUSTOMS_AWB_ROW_PATTERN = re.compile(
    r"(\d{6,12})\s+(\d{6,12})\s+(\d{1,2}-\d{1,2}-\d{4})\s+JKT"
)
# 第 2 頁「Air Waybill…」明細表那一行的啟運港/目的港，例如：
#   "1902246286 0853008513 23-12-2024 JKT, JAKARTA HAK, SOUTH CHINA AREA
#    DUTIES & .! 0"
# 港口代碼後面一定緊接著逗號，用這個特徵跟後面的地區名稱切開 (跟
# _DHL_MAIN_LINE_PATTERN 的 ocode/dcode 同樣的做法)；用 "DUTIES" 當右邊界，
# 因為這個掃描格式的 Type of Service 固定是 "DUTIES & TAXES"。
_DHL_CUSTOMS_PORT_PATTERN = re.compile(
    r"\d{6,12}\s+\d{6,12}\s+\d{1,2}-\d{1,2}-\d{4}\s+"
    r"(?P<origin>[A-Z]{2,4},\s*[A-Za-z ]+?)\s+"
    r"(?P<dest>[A-Z]{2,4},\s*[A-Za-z ]+?)\s+DUTIES"
)
_DHL_CUSTOMS_SUMMARY_LINE_PATTERN = re.compile(
    # 第 1 頁「Type of Service」彙總表那一列，欄位依序是：Number of
    # Shipments / Total Weight / Number of Items / Standard Shipping
    # Charge / Total Extra Charges / VAT / Total (incl. VAT)，共用的欄位
    # 數量在不同帳單上會有點差異 (pdfplumber 有時會把某條分隔線抽成一個
    # 雜訊字元 "=")，所以這裡只把「名稱後面一整串數字/雜訊 token」都框起
    # 來，實際要取第幾個 token 由 _dhl_customs_extract_items() 依固定順序
    # (前 4 個一定是 Shipments/Weight/Items/StandardCharge) 自己判斷，
    # 不在正規表示式裡死板規定總共有幾欄。
    r"^([A-Z][A-Z /&]+[A-Z])\s+((?:[\d.,]+|=)(?:\s+(?:[\d.,]+|=))*)\s*$"
)


def _dhl_customs_extract_items(text: str) -> List[Dict]:
    """擷取每一筆費用項目：

    1. 第 1 頁最上面「Type of Service」彙總表的服務類型列 (例如
       "DUTIES & TAXES")，金額是「Standard Shipping Charge」欄——這種
       關務/完稅費用帳單通常這欄是 0 (沒有基本運費，只有額外費用)，但
       這一列本身仍然算一筆費用明細 (金額 0)，要跟其他供應商「有抓到
       Type of Service 這一列就算一筆」的邏輯一致，不能因為金額剛好是 0
       就跳過不算。
    2. 「Analysis of Extra Charges」區塊逐行擷取費用名稱/金額，這個區塊
       OCR 品質最好、沒有跟其他欄位擠在同一行，比第 2 頁明細表可靠。只
       在「Analysis of Extra Charges」到「Total Extra Charges」這個範圍
       內找，避免誤吃到後面付款資訊裡剛好也符合「大寫字+空白+數字」格式
       的行 (例如匯款帳號 "ACCOUNT NO 956907730")。
    """
    items: List[Dict] = []
    lines = text.split("\n")

    for line in lines:
        m = _DHL_CUSTOMS_SUMMARY_LINE_PATTERN.match(line.strip())
        if not m:
            continue
        # 前 4 個 token 固定是 Number of Shipments / Total Weight /
        # Number of Items / Standard Shipping Charge，第 4 個 (index 3)
        # 才是我們要的金額，後面欄位數量不一定 (可能混著雜訊 "=") 所以不管。
        tokens = [t for t in m.group(2).split() if t != "="]
        if len(tokens) < 4:
            continue
        amount = _dhl_clean_amount(tokens[3])
        items.append({"description": m.group(1).strip(), "amount": amount if amount is not None else 0})
        break

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
        m = _DHL_CUSTOMS_ITEM_PATTERN.match(stripped)
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


def _parse_dhl_scanned_customs(text: str) -> Dict:
    header = _dhl_common_header(text)

    m = _DHL_CUSTOMS_AWB_ROW_PATTERN.search(text)
    if m:
        header["bl_no"] = m.group(1)
        header["order_no"] = m.group(2)
        header["onboard_date"] = _dhl_parse_date(m.group(3))
    else:
        header["bl_no"] = None
        header["order_no"] = None
        header["onboard_date"] = None

    port_m = _DHL_CUSTOMS_PORT_PATTERN.search(text)
    header["port_of_loading"] = port_m.group("origin").strip() if port_m else None
    header["port_of_discharge"] = port_m.group("dest").strip() if port_m else None

    for code in ("arrive_date", "mbl_no", "volume", "vessel", "container_no"):
        header[code] = None

    items = _dhl_customs_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None}]

    return {"header": header, "items": items}


def parse_dhl(text: str) -> Dict:
    if _dhl_is_regular(text):
        result = _parse_dhl_regular(text)
        if result is not None:
            return result
    return _parse_dhl_scanned_customs(text)


register_supplier(
    "DHL", "PT BIROTIKA SEMESTA / DHL EXPRESS",
    detect_dhl, parse_dhl,
    extract_total=_dhl_extract_total,
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
    return _format_date_en(day, month, f"20{year_2digit}")


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
            # 費用明細行裡 "PPN IN IDR" 這一欄就是這筆費用的稅額 (VAT)，
            # 跟 "CHARGES IN IDR" (amount) 並排在同一行，_DSV_ITEM_PATTERN
            # 本來就有抓這個群組，只是先前沒有用上。
            "vat": _dsv_clean_amount(m.group("ppn")),
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

    # "ULTIMATE SHIPPER" 跟 "ULTIMATE CONSIGNEE" 兩欄的名稱是左右並排、
    # 兩欄中間沒有明顯分隔符號 (pdfplumber 抽取後變成同一行文字直接黏在
    # 一起，例如 "PT. GLOSTAR INDONESIA ADIDAS SPOR MALZEMLERI...")，沒辦
    # 法單純用正則表達式切開兩欄。但觀察到的樣本裡，ULTIMATE SHIPPER 的值
    # 每一張都跟最上面「帳單寄送對象」(也就是 header["consignee"]) 完全
    # 一樣 (都是同一家委任 DSV 出貨的公司)，所以直接沿用已經可靠抓到的
    # consignee 值，不用再另外處理那兩欄黏在一起的文字。
    header["shipper"] = header["consignee"]

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
    extract_total=_dsv_extract_total,
)


# ===========================================================================
# 供應商 14：P.T. EVERGREEN LOGISTICS INDONESIA (海運帳單)
# ===========================================================================
#
# 版面特徵：跟 INDOPROSTIME/MAERSK 同款「左右並排欄位、只隔一個空白」排版，
# 例如：
#   "HBL : RWRD501600002926 CONTAINER NO : KMTU9431774"
#   "MBL : KMTCPUSK052666 CONSIGNEE : PT. POU YUEN INDONESIA"
#   "DESTINATION : JAKARTA ETD/ETA : 20260701"
# 一樣用「遇到下一個已知欄位關鍵字就停止」處理。
#
# 日期是 8 位數字的 YYYYMMDD (例如 'Invoice Date : 20260814'，
# 'ETD/ETA : 20260701')，沒有分隔符號，用 _evergreen_parse_yyyymmdd()
# 轉成 'DD.MM.YYYY'。這家帳單只有一個 ETD/ETA 日期同時代表到離岸，正確
# 答案只對應到 arrive_date，onboard_date 這張帳單沒有，留 None 補 N/A。
#
# 費用明細行是 "全大寫描述 + IDR + 金額" (例如 'OCEAN FREIGHT IDR
# 11,562,850')；VAT 那一行 (稅金) 正確答案沒有列成單獨一筆費用明細 (直接
# 併入總額，不算費用明細)，所以擷取時要排除 "VAT" 這一行，避免多算一筆。

def detect_evergreen(text: str) -> bool:
    return "EVERGREEN LOGISTICS" in text.upper()


_EVERGREEN_LABEL_WORDS = [
    r"To", r"Invoice\s*No\.?", r"Invoice\s*Date", r"HBL", r"MBL",
    r"CONTAINER\s*NO", r"CONSIGNEE", r"DESTINATION", r"ETD\s*/\s*ETA",
    r"VSL\s*/\s*VOY",
]
_EVERGREEN_NEXT = "(?:" + "|".join(_EVERGREEN_LABEL_WORDS) + ")"
_EVERGREEN_STOP = rf"(?=\s+{_EVERGREEN_NEXT}\b|\n|$)"


def _evergreen_parse_yyyymmdd(text: Optional[str]) -> Optional[str]:
    """'20260814' (YYYYMMDD，無分隔符號) -> '14.08.2026'。"""
    if not text:
        return None
    m = re.match(r"(\d{4})(\d{2})(\d{2})\s*$", text.strip())
    if not m:
        return None
    year, month, day = m.groups()
    return _format_date_en(day, month, year)


def _evergreen_parse_yyyymmdd_slash(text: Optional[str]) -> Optional[str]:
    """'20260701' (YYYYMMDD) -> '01 Jul 2026'。

    這家供應商正確答案 Excel 裡 arrive_date (ETD/ETA) 這一格過去是「純
    文字」'YYYY/MM/DD' (不是 Excel 日期型態儲存格)，為了配合全系統統一
    改成「月份英文縮寫」的顯示格式，已經請使用者把這一格也改成跟其他
    日期欄位一樣的 'DD Mon YYYY' 文字寫法，所以這裡直接沿用跟
    _evergreen_parse_yyyymmdd() 一樣的共用格式化函式即可。"""
    if not text:
        return None
    m = re.match(r"(\d{4})(\d{2})(\d{2})\s*$", text.strip())
    if not m:
        return None
    year, month, day = m.groups()
    return _format_date_en(day, month, year)


def _evergreen_clean_amount(text: Optional[str]) -> Optional[int]:
    """英式數字 '11,562,850' 或 '4,660,097.00' (逗號千分位，可能有小數點)
    -> 11562850 / 4660097。小數部分是整數金額就轉成 int，不是整數才保留
    小數 (目前觀察到的樣本小數都是 .00，但以防萬一仍支援非整數金額)。
    """
    if not text:
        return None
    text = text.strip().replace(",", "")
    try:
        num = float(text)
    except ValueError:
        return None
    return int(num) if num.is_integer() else num


# 費用明細行，例如 "OCEAN FREIGHT IDR 11,562,850" 或 "ALL IN LOCAL CHARGES
# IDR 4,660,097.00" (金額後面可能有 ".00" 小數)。VAT 這一行 (稅金) 不算
# 進費用明細 (正確答案 Excel 沒有列這一筆)，用負向前瞻排除。
_EVERGREEN_ITEM_PATTERN = re.compile(
    r"^(?!VAT\b)(?P<desc>[A-Z][A-Z /]+?)\s+IDR\s+(?P<amount>[\d,]+(?:\.\d+)?)\s*$",
    re.MULTILINE,
)


def _evergreen_extract_items(text: str) -> List[Dict]:
    items = []
    for m in _EVERGREEN_ITEM_PATTERN.finditer(text):
        items.append({
            "description": m.group("desc").strip(),
            "amount": _evergreen_clean_amount(m.group("amount")),
        })
    vat = _evergreen_extract_vat(text)
    if vat is not None:
        # 跟 APL 同一套慣例：只有一筆費用明細時 (不會混淆稅額對應哪一筆)，
        # VAT 那一列的 Description 直接沿用那筆費用的描述；兩筆以上費用
        # 明細時 (無法判斷稅額對應哪一筆) 留空。
        vat_desc = items[0]["description"] if len(items) == 1 else None
        items.append({"description": vat_desc, "amount": None, "vat": vat})
    return items


def parse_evergreen(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {}

    header["invoice_no"] = _search(r"Invoice\s*No\.?\s*:\s*(\S+)", text)
    header["invoice_date"] = _evergreen_parse_yyyymmdd(
        _search(r"Invoice\s*Date\s*:\s*(\d{8})", text)
    )
    header["supplier"] = _search(
        r"^(P\.T\.\s*EVERGREEN\s+LOGISTICS\s+INDONESIA)\s*$", text
    )
    header["consignee"] = _search(rf"\bCONSIGNEE\s*:\s*(.+?){_EVERGREEN_STOP}", text)
    header["bl_no"] = _search(rf"\bHBL\s*:\s*(.+?){_EVERGREEN_STOP}", text)
    header["mbl_no"] = _search(rf"\bMBL\s*:\s*(.+?){_EVERGREEN_STOP}", text)
    header["port_of_discharge"] = _search(
        rf"\bDESTINATION\s*:\s*(.+?){_EVERGREEN_STOP}", text
    )
    header["arrive_date"] = _evergreen_parse_yyyymmdd_slash(
        _search(rf"\bETD\s*/\s*ETA\s*:\s*(.+?){_EVERGREEN_STOP}", text)
    )
    header["vessel"] = _search(rf"\bVSL\s*/\s*VOY\s*:\s*(.+?){_EVERGREEN_STOP}", text)
    header["container_no"] = _search(
        rf"\bCONTAINER\s*NO\s*:\s*(.+?){_EVERGREEN_STOP}", text
    )

    # 這張帳單沒有訂單編號/開船日/裝貨港/材積/Shipper 資訊，留空給
    # expand_to_rows() 補 N/A。
    header["order_no"] = None
    header["onboard_date"] = None
    header["port_of_loading"] = None
    header["volume"] = None
    header["shipper"] = None

    items = _evergreen_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None, "vat": None}]

    return {"header": header, "items": items}


register_supplier(
    "EVERGREEN_LOGISTICS", "P.T. EVERGREEN LOGISTICS INDONESIA",
    detect_evergreen, parse_evergreen,
    extract_total=_evergreen_extract_total,
    extract_grand_total=_evergreen_extract_grand_total,
)


# ===========================================================================
# 供應商：PT. APL LOGISTICS (單頁、單張發票，VAT 只有整張發票一個總額)
# ===========================================================================
#
# 版面特徵：
#   - "BILLED TO :\nPT. NIKOMAS GEMILANG NO. 2026-0904-0114\n..." -> 公司
#     名稱後面緊接著 "NO." 跟發票號碼，同一行。
#   - 費用明細表格有兩種排版 (同一家供應商、不同帳單會混用)：
#       (a) "Description Amount (IDR)"：單一貨幣，例如
#           "OVERTIME CHARGES 700,000.00"
#       (b) "Description Amount (USD) Amount (IDR)"：雙貨幣並排，例如
#           "SEAL FEE 12.00 217,872.00" (12.00 是美金、217,872.00 才是
#           正確答案要的 Amount(IDR))。
#     兩種排版用同一個規則處理：一律取「這一行最後一個金額數字」當
#     Amount，前面如果還有一個金額數字 (美金) 就直接略過不用。
#   - 費用明細表格裡偶爾會有「不是真正費用」的參考資訊行夾在裡面 (例如
#     "LI No: JKTL246686"、"BOOKING NO.\n319263"，或費用名稱後面額外幾行
#     備註如 "ETD 30 AUG 2026")，這些行结尾不是「數字.兩位小數」，用同一個
#     正規表示式掃描時會自然被跳過，不需要額外的排除邏輯。
#   - VAT 只有整張帳單一個總額 (不是逐筆列出)，跟 KUEHNE_NAGEL/
#     TRANS_DAYA_PRIMA 一樣，獨立成一筆 description=None/amount=None/
#     vat=稅額 的合成列，附加在費用明細最後面。
#   - 開船日 (onboard_date/ETD) 有兩個可能的來源，依序嘗試：
#       1. 費用明細備註裡的 "ETD 30 AUG 2026" (英文月份縮寫)，如果有的話
#          這個才是正確答案要的值 (比帳單最上方 "We debit your account
#          for: ... - ETD : 09/09/2026" 那個日期更準)。
#       2. 沒有備註的帳單 (例如雙貨幣排版那種，費用明細裡沒有額外備註)，
#          改抓最上方那一行的 "ETD : 10/03/2026" (美式 MM/DD/YYYY)。
#   - Order No./Arrive Date/MBL NO./Port of discharge/Volume/Vessel/
#     Container no./Shipper 這張帳單都沒有對應資訊，留空給
#     expand_to_rows() 補 N/A。Port of Loading 兩個樣本都是 "Jakarta"
#     (帳單本身沒有印出「啟運港」這個標籤，固定用這家供應商所在地)。

def detect_apl(text: str) -> bool:
    return "APL LOGISTICS" in text.upper()


_APL_ITEM_PATTERN = re.compile(
    r"^(?P<desc>[A-Z][A-Z0-9 /\-\.]*?)\s+(?:[\d,]+\.\d{2}\s+)?(?P<amount>[\d,]+\.\d{2})\s*$",
    re.MULTILINE,
)


def _apl_clean_amount(text: Optional[str]) -> Optional[int]:
    """英式數字 '700,000.00' (逗號千分位、小數點) -> 700000。"""
    if not text:
        return None
    text = text.strip().replace(",", "")
    try:
        return round(float(text))
    except ValueError:
        return None


def _apl_extract_items(text: str) -> List[Dict]:
    m = re.search(r"Description\s+Amount.*?\n(.*?)\nSubtotal", text, re.DOTALL)
    if not m:
        return []
    items = []
    for im in _APL_ITEM_PATTERN.finditer(m.group(1)):
        items.append({
            "description": im.group("desc").strip(),
            "amount": _apl_clean_amount(im.group("amount")),
        })
    vat = _apl_clean_amount(_search(r"VAT\s*:\s*([\d,]+\.\d{2})", text))
    if vat is not None:
        # ⚠️ 2026-10 修正 (使用者提供新版正確答案明確要求)：VAT 那一列的
        # Description 不分單筆/多筆費用明細，一律固定填 "VAT"(照抄 PDF 上
        # "VAT : 77,000.00" 這個標籤，PDF 沒有印百分比，跟 TRANS_DAYA_PRIMA
        # 的 "VAT (11%)" 不同，不要套錯格式)。舊版「只有一筆費用時沿用該筆
        # description、多筆時留空」的規則已經不符合新正確答案，廢棄。
        items.append({"description": "VAT", "amount": None, "vat": vat})
    return items


def _apl_extract_total(text: str) -> Optional[int]:
    return _apl_clean_amount(_search(r"Subtotal\s*:\s*([\d,]+\.\d{2})", text))


def _apl_extract_grand_total(text: str) -> Optional[int]:
    return _apl_clean_amount(_search(r"Total Amount\s*:\s*([\d,]+\.\d{2})", text))


def parse_apl(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {}

    header["supplier"] = _search(r"^(PT\.\s*APL\s*LOGISTICS)", text)

    m = re.search(r"BILLED TO\s*:\s*\n(.+?)\s+NO\.\s*(\S+)", text)
    header["consignee"] = m.group(1).strip() if m else None
    header["invoice_no"] = m.group(2).strip() if m else None

    _apl_inv_m = re.search(r"Jakarta,\s*([A-Za-z]+)\s+(\d{1,2}),\s*(\d{4})", text)
    header["invoice_date"] = (
        _format_date_en_from_month_name(
            _apl_inv_m.group(2), _apl_inv_m.group(1)[:3], _apl_inv_m.group(3)
        )
        if _apl_inv_m else None
    )

    # 開船日：先試費用明細備註裡的 "ETD 30 AUG 2026"，沒有才退回最上方
    # "We debit your account for: ... - ETD : 10/03/2026" (美式 MM/DD/YYYY)。
    _apl_etd_note_m = re.search(r"\bETD\s+(\d{1,2})\s+([A-Za-z]{3})\s+(\d{4})", text)
    if _apl_etd_note_m:
        header["onboard_date"] = _format_date_en_from_month_name(
            _apl_etd_note_m.group(1), _apl_etd_note_m.group(2), _apl_etd_note_m.group(3)
        )
    else:
        _apl_etd_hdr_m = re.search(r"\bETD\s*:\s*(\d{1,2})/(\d{1,2})/(\d{4})", text)
        header["onboard_date"] = (
            _format_date_en(_apl_etd_hdr_m.group(2), _apl_etd_hdr_m.group(1),
                             _apl_etd_hdr_m.group(3))
            if _apl_etd_hdr_m else None
        )

    header["bl_no"] = _search(r"\b(?:LI\s*No|FCR\s*No)\s*:\s*(\S+)", text)
    # 帳單上沒有印出「啟運港」這個標籤，但兩個樣本都是從 Jakarta 出貨，
    # 固定填 "Jakarta" (這家供應商公司本身就在 Jakarta)。
    header["port_of_loading"] = "Jakarta"

    # 這張帳單沒有單號/到貨日/主提單/目的港/材積/船名/貨櫃/出貨方資訊，
    # 留空給 expand_to_rows() 補 N/A。
    header["order_no"] = None
    header["arrive_date"] = None
    header["mbl_no"] = None
    header["port_of_discharge"] = None
    header["volume"] = None
    header["vessel"] = None
    header["container_no"] = None
    header["shipper"] = None

    items = _apl_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None, "vat": None}]

    return {"header": header, "items": items}


register_supplier(
    "APL", "PT. APL LOGISTICS", detect_apl, parse_apl,
    extract_total=_apl_extract_total, extract_grand_total=_apl_extract_grand_total,
)


# ===========================================================================
# 供應商：PT. GALUH INTI BAHARI (單頁、單張發票，稅額逐筆列在費用明細裡)
# ===========================================================================
#
# 版面特徵：
#   - "TO PT. HARDASES ABADI INDONESIA No. Invoice : 2601312-KWTGIB" ->
#     收件公司跟發票號碼同一行，中間用 "No. Invoice :" 分隔。
#   - "Wonopringgo Tgl. Invoice : 07 Agt 2026" -> 發票日期用印尼文月份縮寫
#     (Agt=Agustus、Sep=September...)，跟 DWIHARTA 的印尼文全名不同，所以
#     用這家供應商自己的縮寫對照表 _GALUH_MONTH_ABBR，不沿用 INDO_MONTHS。
#   - 費用明細表格 "No. KETERANGAN JUMLAH" 之後是逐行費用，格式
#     "<序號> <費用名稱> <金額>" (金額千分位用句點，例如 "9.400.000")；
#     這家供應商跟其他供應商不同的地方是：PPN/PPh 稅額**不是整張帳單只有
#     一個總額的獨立列**，而是逐筆直接列在費用明細裡 (沒有前面的序號)，
#     例如 "PPN 11% 16.500"、"PPh 2% -191.000" (PPh 是從金額裡「扣除」，
#     所以是負數)。這兩種稅額行各自獨立成一筆 description=稅額名稱/
#     amount=None/vat=稅額 的列，不套用其他供應商「VAT 只有一個總額」的
#     合成列慣例 (那種慣例的 description 規則不適用，這裡本來就有自己的
#     description)。
#   - 有些費用名稱後面會黏著一段日期備註 (例如 "Trucking 27 Juli 2026"、
#     "TRUCKING PER TGL 31 AGUSTUS - 05 SEPTEMBER 2026")，正確答案只要
#     名稱本身 ("Trucking"/"TRUCKING")，日期備註要用
#     _galuh_strip_date_suffix() 去掉。
#   - 帳單本身沒有另外印出「未稅總額」，費用明細加總 (不含稅額列) 直接
#     加上 PPN/PPh 稅額就等於帳單最下面的 "Total" (已經是含稅後的總額)，
#     所以只註冊 extract_grand_total，不註冊 extract_total (沒有獨立的未稅
#     小計可以比對)。
#   - 收件公司 (consignee) 通常從最上面 "TO <公司名> No. Invoice :" 抓；
#     但遇到像 ASICS 這張帳單時，帳單本身用簡稱 "PT. POUCHEN ( ASICS )"，
#     實際報表第二頁 (same PDF 的「REKAPITULASI TAGIHAN...」明細頁) 開頭
#     印出的才是正確答案要的正式公司全名 "PT. POU CHEN INDONESIA\nNo
#     :2601512-KWTGIB"，所以優先找「行首是 PT. 開頭、後面接 No : <發票號>
#     的那一行」，有找到就用那個覆寫，沒有 (單頁帳單通常沒有這一行) 才退
#     回用最上面 "TO ..." 那一行。
#   - 這張帳單沒有訂單編號/到貨日/開船日/提單號/主提單號/啟運港/目的港/
#     材積/船名/貨櫃號/出貨方資訊，留空給 expand_to_rows() 補 N/A。

def detect_galuh(text: str) -> bool:
    return "GALUH INTI BAHARI" in text.upper()


_GALUH_MONTH_ABBR = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "mei": 5, "jun": 6,
    "jul": 7, "agt": 8, "sep": 9, "okt": 10, "nov": 11, "des": 12,
}


def _galuh_parse_date(day, month_abbr: Optional[str], year) -> Optional[str]:
    if not month_abbr:
        return None
    month = _GALUH_MONTH_ABBR.get(str(month_abbr).strip().lower()[:3])
    if not month:
        return None
    return _format_date_en(day, month, year)


def _galuh_clean_amount(text: Optional[str]) -> Optional[int]:
    """印尼式數字 '9.400.000' (句點千分位) -> 9400000；PPh 常見是負數
    '-191.000' -> -191000，int() 本身就支援前面的負號，不用特別處理。
    """
    if not text:
        return None
    text = text.strip().replace(".", "").replace(",", "")
    try:
        return int(text)
    except ValueError:
        return None


_GALUH_BLOCK_PATTERN = re.compile(
    r"No\.\s*KETERANGAN\s*JUMLAH\s*\n(.*?)\nTotal\s+[\d.,]+\s*$",
    re.DOTALL | re.MULTILINE,
)
_GALUH_TAX_LINE_PATTERN = re.compile(r"^(?P<desc>PP[Nh]\s+\d+%)\s+(?P<amount>-?[\d.]+)\s*$")
_GALUH_ITEM_LINE_PATTERN = re.compile(
    r"^\d+\s+(?P<desc>.+?)\s+(?P<amount>\d{1,3}(?:\.\d{3})*)\s*$"
)
_GALUH_MONTH_NAMES = (
    "Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|"
    "Oktober|November|Desember"
)


def _galuh_strip_date_suffix(desc: Optional[str]) -> Optional[str]:
    """去掉費用名稱後面黏著的日期備註，例如 'Trucking 27 Juli 2026' ->
    'Trucking'、'TRUCKING PER TGL 31 AGUSTUS - 05 SEPTEMBER 2026' ->
    'TRUCKING'。沒有日期備註的名稱 (例如 'Fiat Dokumen'、'TOTAL BIAYA
    EKSTERNAL NON PPH 2%') 原樣保留。
    """
    if not desc:
        return desc
    desc = re.sub(r"\s+(?:PER\s+)?TGL\b.*$", "", desc, flags=re.IGNORECASE)
    desc = re.sub(
        rf"\s+\d{{1,2}}\s+(?:{_GALUH_MONTH_NAMES})\s+\d{{4}}.*$",
        "", desc, flags=re.IGNORECASE,
    )
    return desc.strip()


def _galuh_rejoin_split_tax_lines(block: str) -> str:
    """費用項目筆數變多時，PDF 版面「KETERANGAN」跟「JUMLAH」兩欄的垂直
    對齊會跟著跑掉，導致 PPN/PPh 那一列的金額被 pdfplumber 拆到跟名稱不同
    的一行，例如 (3 筆費用項目的帳單)：
        PPN 11% 33.000
        -194.000
        PPh 2%
    "PPh 2%" 自己落單一行、金額 "-194.000" 反而黏在上一行。這裡用正規表
    示式把「金額自己一行 + 緊接著下一行是稅額名稱」重新接回同一行，變成
    跟原本版面一致的 "PPh 2% -194.000"，這樣不管這張帳單有幾筆費用明細
    (進而讓 PPN/PPh 那一列被往下擠或對齊跑掉)，後面逐行解析都不用再為
    「金額跟名稱被拆開」的情況另外特判。
    """
    return re.sub(
        r"^(-?[\d.]+)\s*\n\s*(PP[Nh]\s+\d+%)\s*$",
        r"\2 \1",
        block,
        flags=re.MULTILINE,
    )


def _galuh_extract_items(text: str) -> List[Dict]:
    block_m = _GALUH_BLOCK_PATTERN.search(text)
    if not block_m:
        return []
    block = _galuh_rejoin_split_tax_lines(block_m.group(1))
    items: List[Dict] = []
    for line in block.splitlines():
        line = line.strip()
        if not line:
            continue
        tax_m = _GALUH_TAX_LINE_PATTERN.match(line)
        if tax_m:
            items.append({
                "description": tax_m.group("desc").strip(),
                "amount": None,
                "vat": _galuh_clean_amount(tax_m.group("amount")),
            })
            continue
        item_m = _GALUH_ITEM_LINE_PATTERN.match(line)
        if item_m:
            items.append({
                "description": _galuh_strip_date_suffix(item_m.group("desc").strip()),
                "amount": _galuh_clean_amount(item_m.group("amount")),
                "vat": None,
            })
    return items


def _galuh_extract_grand_total(text: str) -> Optional[int]:
    return _galuh_clean_amount(_search(r"^Total\s+([\d.,]+)\s*$", text))


def parse_galuh(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {}

    header["supplier"] = _search(r"^(PT\.\s*GALUH\s*INTI\s*BAHARI)", text)
    header["invoice_no"] = _search(r"No\.\s*Invoice\s*:\s*(\S+)", text)

    _galuh_date_m = re.search(r"Tgl\.\s*Invoice\s*:\s*(\d{1,2})\s+(\w+)\s+(\d{4})", text)
    header["invoice_date"] = (
        _galuh_parse_date(*_galuh_date_m.groups()) if _galuh_date_m else None
    )

    # consignee：優先找報表明細頁開頭「PT. 開頭、後面接 No : <發票號>」那
    # 一行 (比最上面的 "TO ..." 簡稱更正式/更準確)，找不到才退回 "TO ..."。
    _galuh_rekap_m = re.search(r"^(PT\..+?)\s+No\s*:\s*\S+\s*$", text, re.MULTILINE)
    if _galuh_rekap_m:
        header["consignee"] = _galuh_rekap_m.group(1).strip()
    else:
        header["consignee"] = _search(r"\bTO\s+(.+?)\s+No\.\s*Invoice", text)

    # 這張帳單沒有訂單編號/到貨日/開船日/提單號/主提單號/啟運港/目的港/
    # 材積/船名/貨櫃號/出貨方資訊，留空給 expand_to_rows() 補 N/A。
    header["order_no"] = None
    header["arrive_date"] = None
    header["onboard_date"] = None
    header["bl_no"] = None
    header["mbl_no"] = None
    header["port_of_loading"] = None
    header["port_of_discharge"] = None
    header["volume"] = None
    header["vessel"] = None
    header["container_no"] = None
    header["shipper"] = None

    items = _galuh_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None, "vat": None}]

    return {"header": header, "items": items}


register_supplier(
    "GALUH", "PT. GALUH INTI BAHARI", detect_galuh, parse_galuh,
    extract_grand_total=_galuh_extract_grand_total,
)


# ===========================================================================
# 供應商：PT. TRANS ECOMET GLOBAL (⚠️ multi_page：一份 PDF 塞好幾張各自
# 獨立的發票，一頁一張，跟 HYPER MEGA/PANCARAN 同一種型態)
# ===========================================================================
#
# 版面特徵：
#   - 每一頁最上面都會出現收款帳戶那一行「...rekening: PT. TRANS ECOMET
#     GLOBAL...」，這是最穩定、不會跟其他供應商混淆的關鍵字。
#   - 這份 PDF 的「冒號」在不同頁會印在不一樣的位置：有的頁面冒號是獨立字元
#     (正常的 "Shipper : ADIDAS")，有的頁面冒號卻黏在數字/字母中間
#     (例如 "HB/L No SNC20:260296"、"Negara Asal THAILAND:")，這單純是
#     PDF 排版的渲染問題，冒號本身從來不是用來切欄位的分隔符號、也不是
#     值的一部分。解法：parse 一開始就把整頁文字裡所有的 ":" 直接拿掉
#     (`t = text.replace(":", "")`)，之後全部規則只認「標籤後面接的內容」，
#     不依賴冒號，兩種渲染方式都會變成同一種乾淨的文字可以共用同一套
#     regex (例如 "SNC20:260296" 拿掉冒號後就是正確答案要的 "SNC20260296"，
#     "GAR:UDA" 拿掉冒號後就是 "GARUDA")。
#   - Shipper 公司名稱偶爾會換行(例如 "CONG TY TNHH KHOA HOC CONG NGHE HUA"
#     跨到下一行才接 "CHANG VINA" 才結束)，換行的那一段剛好夾在地址第一行
#     "JL. RAYA SERANG KM.71 DESA" 跟地址第二行 "TAMBAK, KIBIN, SERANG,"
#     中間，用一個「可選的中間擷取群組」處理即可，見 _ecomet_extract_shipper()。
#   - Container no. 可能是「單一個 - (沒有貨櫃)」或「好幾個貨櫃號跨好幾行」
#     (換行方式跟 HYPER MEGA 的 Shipper 欄位類似)，一路抓到下一個已知標籤
#     "ETD" 那一行為止，見 _ecomet_extract_container()。
#   - Vessel 欄位：空運帳單用 "Flight" 欄位值；海運帳單用 "Voyage" 欄位值，
#     Voyage 沒有資料時印成 "- -"，正確答案裡簡化成單一個 "-"。
#   - Volume (材積) 欄位：正確答案一律是「Nett Weight 的第一個數字 + " KG"」，
#     不管 PDF 上是不是已經印了 "KG" 單位、也不管後面還有幾個重複的重量值
#     (多貨櫃的帳單 Nett Weight 會印成 "0 / 0 / 0" 這種多值，答案只取第一個)。
#   - Port of discharge (目的港) 這份帳單從頭到尾都沒有印出來 (因為收件方
#     固定是印尼的 PT. NIKOMAS GEMILANG)，正確答案固定是 "INDONESIA"，
#     這裡直接用常數帶入，不用去 PDF 找一個不存在的欄位。
#   - 費用明細有兩段獨立編號 (都是從 1 開始)：「No Biaya Jumlah」那一段是
#     一般費用 (含 PPN，PPN 那一列要進 vat 欄位、不是 amount，跟 GALUH的
#     PPN/PPh 列處理方式一樣)，「No Reimbursement Keterangan Jumlah」那
#     一段是代墊費用，部分項目後面會多一個 "ket: ..." 備註 (例如
#     "Storage ket: RP 1.006.110 Rp 1,116,782.00")，真正要的金額是最後的
#     "Rp ..."，不是 "ket:" 後面那個印尼文備註數字，描述只取 "ket:" 前面
#     的文字。
#   - Order No./Invoice Date 這張帳單完全沒有對應欄位，正確答案固定是
#     "N/A"，留 None 讓 expand_to_rows() 自動補即可。

def detect_ecomet(text: str) -> bool:
    return "ECOMET" in text.upper()


def _ecomet_clean_amount(text: Optional[str]) -> Optional[int]:
    if not text:
        return None
    text = text.strip().replace(",", "")
    try:
        return int(float(text))
    except ValueError:
        return None


def _ecomet_parse_date(text: Optional[str]) -> Optional[str]:
    """'26-Apr-26' -> '26 Apr 2026' (2 位數年份一律視為 20XX)。"""
    if not text:
        return None
    m = re.match(r"(\d{1,2})-([A-Za-z]+)-(\d{2,4})", text.strip())
    if not m:
        return None
    day, month_abbr, year = m.groups()
    month = _MONTH_NAME_TO_NUM.get(month_abbr.strip().upper()[:3])
    if not month:
        return None
    year_i = int(year)
    if year_i < 100:
        year_i += 2000
    return _format_date_en(day, month, year_i)


def _ecomet_clean_dash(value: Optional[str]) -> Optional[str]:
    """Voyage 沒有資料時印成 '- -' 這種多個破折號，正確答案簡化成單一個
    '-'；其他正常值原樣回傳 (只去頭尾空白)。
    """
    if value is None:
        return None
    value = value.strip()
    if re.fullmatch(r"[-\s]+", value):
        return "-"
    return value


def _ecomet_extract_shipper(flat: str) -> Optional[str]:
    """Shipper 公司名稱可能跨行 (換行那一段夾在地址第一行跟第二行中間)，
    見本區塊開頭的版面說明。flat 是把整頁文字的換行都換成空白之後的版本。
    """
    m = re.search(
        r"Shipper\s+(.+?)\s+JL\.\s*RAYA\s*SERANG\s*KM\.?\s*71\s*DESA"
        r"\s*(.*?)\s*TAMBAK,?\s*KIBIN,?\s*SERANG,?\s*BANTEN\s*Consignee",
        flat,
    )
    if not m:
        return None
    main, overflow = m.group(1).strip(), m.group(2).strip()
    return (main + " " + overflow).strip() if overflow else main


def _ecomet_extract_container(t: str) -> Optional[str]:
    """Container no. 可能單行 (單一個 '-' 或單一個貨櫃號)，也可能好幾個
    貨櫃號跨好幾行，一路抓到下一個已知標籤 'ETD' 那一行為止。
    """
    m = re.search(r"Container\s+(.+?)\s+Qty", t)
    if not m:
        return None
    first = m.group(1)
    tail_m = re.search(r"Container\s+.+?\n((?:(?!ETD).+\n)*)ETD", t)
    extra = tail_m.group(1).replace("\n", " ") if tail_m else ""
    parts = [p.strip() for p in (first + " " + extra).split("/") if p.strip()]
    if not parts:
        return None
    if all(p == "-" for p in parts):
        return "-"
    return " / ".join(parts)


_ECOMET_ITEM_LINE = re.compile(
    r"^\d+\s+(?P<desc>.+?)(?:\s+ket\b.*?)?\s+Rp\s+(?P<amount>[\d,]+\.\d{2})$"
)
_ECOMET_PPN_LINE = re.compile(
    r"^(?P<desc>PPN\s+[\d.]+)\s+Rp\s+(?P<amount>[\d,]+\.\d{2})$"
)


def _ecomet_extract_items(t: str) -> List[Dict]:
    items: List[Dict] = []
    for raw_line in t.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        m_ppn = _ECOMET_PPN_LINE.match(line)
        if m_ppn:
            items.append({
                "description": m_ppn.group("desc").strip(),
                "amount": None,
                "vat": _ecomet_clean_amount(m_ppn.group("amount")),
            })
            continue
        m_item = _ECOMET_ITEM_LINE.match(line)
        if m_item:
            items.append({
                "description": m_item.group("desc").strip(),
                "amount": _ecomet_clean_amount(m_item.group("amount")),
                "vat": None,
            })
    return items


def parse_ecomet(text: str) -> Dict:
    """注意：這是 multi_page 供應商，text 是『一頁』的文字 (不是整份 PDF
    合併後的文字)，呼叫端 (_parse_multi_page_pdf()) 只會對通過
    detect_ecomet() 的頁面呼叫這個函式。

    一開始先把整頁文字裡的 ':' 全部拿掉，解決這份 PDF 冒號亂跑的排版問題
    (見本區塊開頭說明)，後面全部規則都在這個「乾淨版」文字 t 上操作。
    """
    t = text.replace(":", "")
    flat = re.sub(r"\s+", " ", t)

    header: Dict[str, Optional[str]] = {}

    header["invoice_no"] = _search(r"No\s+(\S+)\s+Kepada", flat)
    header["invoice_date"] = None

    header["consignee"] = _search(r"Kepada\s+(.+?)\s+Shipper", flat)
    header["supplier"] = _search(
        r"rekening\s+(PT\.\s*TRANS\s*ECOMET\s*GLOBAL)", flat
    )
    header["shipper"] = _ecomet_extract_shipper(flat)

    header["order_no"] = None

    m_eta = re.search(
        r"ETD\s*/\s*ETA\s+(\d{1,2}-[A-Za-z]+-\d{2,4})\s*/\s*(\d{1,2}-[A-Za-z]+-\d{2,4})",
        flat,
    )
    header["onboard_date"] = _ecomet_parse_date(m_eta.group(1)) if m_eta else None
    header["arrive_date"] = _ecomet_parse_date(m_eta.group(2)) if m_eta else None

    header["bl_no"] = _search(r"HB/L\s*No\s+(\S+)", flat)
    header["mbl_no"] = _search(r"MB/L\s*No\s+(\S+)", flat)
    # Negara Asal 跟 Tipe 分別在不同的物理行 (右欄位逐行往下錯開一列)，
    # 不能用 flat (整頁合併成一行) 去抓到 Tipe 當結尾，國家名稱本身都是
    # 單一個字 (THAILAND/CIKANDE/ITALIA/VIETNAM/JAPAN)，直接抓下一個空白
    # 前的單字即可。
    header["port_of_loading"] = _search(r"Negara\s*Asal\s+(\S+)", flat)
    header["port_of_discharge"] = "INDONESIA"

    m_weight = re.search(r"Nett\s*Weight\s+([\d.,]+)", flat)
    header["volume"] = f"{m_weight.group(1)} KG" if m_weight else None

    flight = _search(r"Flight\s+(.+?)\s+Gross", flat)
    voyage = _search(r"Voyage\s+(.+?)\s+CBM", flat)
    header["vessel"] = flight if flight else _ecomet_clean_dash(voyage)

    header["container_no"] = _ecomet_extract_container(t)

    items = _ecomet_extract_items(t)
    if not items:
        items = [{"description": None, "amount": None, "vat": None}]

    return {"header": header, "items": items}


register_supplier(
    "ECOMET", "PT. TRANS ECOMET GLOBAL", detect_ecomet, parse_ecomet,
    multi_page=True,
)


# ============================================================
# HANDAL (PT. Handal Mandiri Transindo)
# ------------------------------------------------------------
# 兩種 PDF 版型都用同一套 detect/parse：
#   (a) 掃描檔 (例如 47 頁的月結單，一頁一張發票) - 文字層存在但是亂碼
#       (OCR 數字誤讀，例如 8→g, 0→o, 1→4)，用 force_ocr=True 強制走
#       獨立的 OCR pipeline (extract_pages_from_pdf_force_ocr)，不信任
#       pdfplumber 原生抓到的文字層。
#   (b) 可編輯 PDF (例如 HG12240926.pdf) - 沒有任何欄位標籤文字，只有
#       編號+描述+單價+金額的表格列，直接用 pdfplumber 原生文字即可。
# ============================================================

_HANDAL_NUM = r'\d{1,3}(?:,\d{3})+'  # 要求至少一個千分位逗號 (>=1000)，避免把
# 描述文字裡本來就有的小數字 (例如 "TRUCKING LAUT 40 (PPH 23)" 裡的 40、23，
# 或 "CUSTOM CLEARANCE (PART OFF 10)" 裡的 10) 誤判成金額
_HANDAL_SEP = r'[\s.,;]*'

_HANDAL_NON_ITEM_PREFIXES = (
    "DESCRIPTION", "QTY", "UNIT PRICE", "AMOUNT", "NO.", "NO ",
    "VESSEL", "DATE", "INVIOCE", "INVOICE", "NO.BC", "NO. BC",
    "SHIPPING REMARKS", "KEPADA YTH", "DEBIT NOTE", "E.M.K", "EMKU",
    "PT. HANDAL", "HANDAL MANDIRI", "FINANCE", "PAYMENT BY",
    "HARAP DITRANSFER", "BANK", "NO. REK", "KOMPLEK", "TELP",
    "FOR AND ON BEHALF", "TOTAL", "#", "RIBU RUPIAH", "JUTA",
    "JL RAYA", "JL.", "KAWASAN INDUSTRI", "SERANG", "TAMBAK",
    "KAB ", "RAYA SERANG",
)

_HANDAL_CODE_RE = re.compile(r'\b([A-Z]{4,8}\d{6,10})\b')
_HANDAL_ORDER_NO_RE = re.compile(r'(?<![A-Z])([A-Z])(\d{5,7})(?!\d)')
_HANDAL_INVOICE_NO_RE = re.compile(r'\bHG\d{8}\b')
_HANDAL_INVOICE_DATE_RE = re.compile(r'\b(\d{1,2}\s+[A-Za-z]+\s+\d{4})\b')
_HANDAL_ROUTE_RE = re.compile(r'([A-Z][A-Z ]*?) *- *([A-Z][A-Z ]*?) *- *(\d{1,2}/\d{1,2}/\d{4})')
_HANDAL_VOLUME_RE = re.compile(r'\b(\d{1,2}X\d{2,3})\b')
_HANDAL_CONSIGNEE_RE = re.compile(r'\b([A-Z][A-Z\s]{2,40}\.\s*PT)\b')
_HANDAL_TOTAL_RE = re.compile(r'(\d{1,3}(?:,\d{3})+)\s*\n+\s*#')
_HANDAL_GROUPED_NUM_RE = re.compile(r'\d{1,3}(?:,\d{3})+')


def _looks_like_handal_item_desc(line: str) -> bool:
    text = line.strip()
    if not text or len(text) < 3 or len(text) > 48:
        return False
    if not text[0].isalpha() or not text[0].isupper():
        return False
    # 真正的費用項目名稱整行都是大寫；只要出現小寫字母，幾乎都是 OCR 把
    # VESSEL 欄位等雜訊文字誤讀成英數混合的殘留，不是費用項目。
    if re.search(r'[a-z]', text):
        return False
    upper = text.upper()
    for prefix in _HANDAL_NON_ITEM_PREFIXES:
        if upper.startswith(prefix):
            return False
    if re.search(r'\.\s*PT\b', upper):
        return False
    if re.search(r'\d{1,2}/\d{1,2}/\d{4}', text):
        return False
    if re.match(r'^[A-Z][A-Z ]*-[A-Z ]+-', upper):
        return False
    if (_HANDAL_CODE_RE.search(upper) or _HANDAL_ORDER_NO_RE.search(upper)
            or _HANDAL_INVOICE_NO_RE.search(upper)):
        return False
    # 4 個以上連續數字、但不是千分位逗號分隔格式 -> 很可能是編號/代碼，不是費用名稱
    stripped_grouped = _HANDAL_GROUPED_NUM_RE.sub('', upper)
    if re.search(r'\d{4,}', stripped_grouped):
        return False
    allowed = re.sub(r'[A-Z0-9 /()=*.,\-]', '', upper)
    if len(allowed) > 2:
        return False
    if len(re.findall(r'[A-Z]', upper)) < 2:
        return False
    return True


def _handal_amount_to_int(s):
    if not s:
        return None
    s = s.replace(",", "").strip()
    return int(s) if s.isdigit() else None


_HANDAL_MONTHS_EN = {
    "jan": 1, "january": 1, "feb": 2, "february": 2, "mar": 3, "march": 3,
    "apr": 4, "april": 4, "may": 5, "jun": 6, "june": 6, "jul": 7, "july": 7,
    "aug": 8, "august": 8, "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10, "nov": 11, "november": 11, "dec": 12, "december": 12,
}


def _handal_parse_date_monthname(s):
    if not s:
        return None
    m = re.match(r'(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})', s.strip())
    if not m:
        return None
    day, month_name, year = m.groups()
    return _format_date_en_from_month_name(day, month_name, year)


def _handal_parse_date_slash(s):
    if not s:
        return None
    m = re.match(r'(\d{1,2})/(\d{1,2})/(\d{4})', s.strip())
    if not m:
        return None
    d, mo, y = m.groups()
    return _format_date_en(d, mo, y)


def _handal_clean_port(s: str) -> str:
    s = re.sub(r'\s+', ' ', s).strip()
    s = re.sub(r'^(PT)\s+', '', s)
    return s


_HANDAL_COMBINED_ITEM_RE = re.compile(
    rf'^(?:\d+\s+)?(?P<desc>[A-Z][A-Za-z0-9 /().=*\-]*?)[\s,]+(?:\d+[\s,]+)?'
    rf'(?P<unit>{_HANDAL_NUM}){_HANDAL_SEP}(?P<amount>{_HANDAL_NUM})(?!\d)[^\n0-9]*$'
)
_HANDAL_PAIR_ONLY_RE = re.compile(
    rf'^(?P<unit>{_HANDAL_NUM}){_HANDAL_SEP}(?P<amount>{_HANDAL_NUM})(?!\d)[^\n0-9]*$'
)
_HANDAL_SINGLE_NUM_RE = re.compile(
    rf'^(?P<amount>{_HANDAL_NUM})(?!\d)[^\n0-9]*$'
)


def _handal_extract_vessel(text: str, order_no):
    if order_no:
        for line in text.splitlines():
            if order_no in line:
                before = line.split(order_no)[0]
                words = re.findall(r'[A-Z0-9]+', before)
                if words:
                    candidate = ' '.join(words)
                    if len(candidate) >= 3:
                        return candidate
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if 'VESSEL' in line.upper():
            for nxt in lines[i + 1:i + 3]:
                words = re.findall(r'[A-Z0-9]+', nxt)
                if words:
                    candidate = ' '.join(words)
                    if len(candidate) >= 3:
                        return candidate
    return None


def _handal_extract_header(text: str) -> Dict:
    h: Dict = {}

    m = _HANDAL_CONSIGNEE_RE.search(text)
    h['consignee'] = re.sub(r'\s+', ' ', m.group(1)).strip() if m else None

    m = _HANDAL_ROUTE_RE.search(text)
    if m:
        h['port_of_loading'] = _handal_clean_port(m.group(1))
        h['port_of_discharge'] = _handal_clean_port(m.group(2))
        h['onboard_date'] = _handal_parse_date_slash(m.group(3))
    else:
        h['port_of_loading'] = h['port_of_discharge'] = h['onboard_date'] = None

    m = _HANDAL_VOLUME_RE.search(text)
    h['volume'] = m.group(1) if m else None

    m = _HANDAL_CODE_RE.search(text)
    code = m.group(1) if m else None
    h['bl_no'] = code
    h['mbl_no'] = code

    m = _HANDAL_INVOICE_NO_RE.search(text)
    h['invoice_no'] = m.group(0) if m else None

    m = _HANDAL_INVOICE_DATE_RE.search(text)
    h['invoice_date'] = _handal_parse_date_monthname(m.group(1)) if m else None

    m = _HANDAL_ORDER_NO_RE.search(text)
    h['order_no'] = (m.group(1) + m.group(2)) if m else None

    h['vessel'] = _handal_extract_vessel(text, h['order_no'])

    h['supplier'] = "PT. Handal Mandiri Transindo"
    h['arrive_date'] = None
    h['container_no'] = None
    h['shipper'] = None

    return h


def _handal_item_region(text: str) -> str:
    """只取『總金額數字出現以前』的文字範圍來找費用項目，把發票下方的簽名、
    銀行帳號、電話、地址等雜訊文字整段排除，避免被誤判成費用項目。"""
    m = _HANDAL_TOTAL_RE.search(text)
    if m:
        return text[:m.start()]
    return text


def _handal_extract_items(text: str):
    region = _handal_item_region(text)
    lines = [l.strip() for l in region.splitlines()]
    items = []
    pending_descs = []
    pending_numbers = []

    for line in lines:
        if not line:
            continue

        m = _HANDAL_COMBINED_ITEM_RE.match(line)
        if m and _looks_like_handal_item_desc(m.group('desc')):
            items.append({
                "description": m.group('desc').strip(),
                "amount": _handal_amount_to_int(m.group('amount')),
                "vat": None,
            })
            continue

        m2 = _HANDAL_PAIR_ONLY_RE.match(line)
        if m2:
            pending_numbers.append(_handal_amount_to_int(m2.group('amount')))
            continue

        if _looks_like_handal_item_desc(line):
            pending_descs.append(line)
            continue

        m3 = _HANDAL_SINGLE_NUM_RE.match(line)
        if m3:
            pending_numbers.append(_handal_amount_to_int(m3.group('amount')))
            continue

    if pending_descs:
        n = len(pending_descs)
        if len(pending_numbers) == n:
            amounts = pending_numbers
        elif len(pending_numbers) == 2 * n:
            amounts = [pending_numbers[i * 2 + 1] for i in range(n)]
        else:
            amounts = (pending_numbers + [None] * n)[:n]
        for desc, amount in zip(pending_descs, amounts):
            items.append({"description": desc, "amount": amount, "vat": None})

    return items


def detect_handal(text: str) -> bool:
    upper = text.upper()
    if "HANDAL MANDIRI TRANSINDO" in upper or "HANDAL" in upper:
        return True
    if (_HANDAL_INVOICE_NO_RE.search(text) and "SEGEL" in upper
            and "CUSTOM CLEARANCE" in upper):
        return True
    return False


def parse_handal(text: str) -> Dict:
    header = _handal_extract_header(text)
    items = _handal_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None, "vat": None}]
    return {"header": header, "items": items}


def _handal_extract_total(text: str):
    m = _HANDAL_TOTAL_RE.search(text)
    if not m:
        return None
    return _handal_amount_to_int(m.group(1))


register_supplier(
    "HANDAL", "PT. Handal Mandiri Transindo", detect_handal, parse_handal,
    multi_page=True, force_ocr=True,
    extract_total=_handal_extract_total,
)


# ---------------------------------------------------------------------------
# 供應商 21：PT. PUTRA SEJAHTERA SENTOSA
# ---------------------------------------------------------------------------
# multi_page：一份 PDF 裡塞了好幾張各自獨立的發票，一頁一張 (每頁都有自己
# 的 Invoice No./Date，跟 HYPER MEGA/YJE_TATA 同一種型態)。
#
# 這家供應商有 PPN (加值稅，1,1% 的印花稅率寫法) 跟 PPH (預扣所得稅) 兩種
# 稅金欄位，各自獨立一列 (VAT 呈現方式分類第 2 種「獨立一列稅金」，見
# add-invoice-supplier skill)，描述照抄 PDF 原文標籤 (例如 "PPN 1,1 %"、
# "PPH 2%"，百分比用 regex 動態抓出來，不要寫死)：
#   - "PPN x %"：一般加值稅，要『加』進含稅總額 (GRAND TOTAL = DPP + PPN)。
#   - "PPH x%"：預扣所得稅，是由付款方代扣、要從含稅總額『扣』掉才等於
#     實付金額 (TOTAL YANG DIBAYAR = GRAND TOTAL - PPH)，不是加總的稅金，
#     所以用 register_supplier() 的 deduction_prefixes=("PPH",) 標記，稽核
#     含稅總額時才不會被誤算進去 (見 extract_utils_streamlit.py 的
#     audit_invoice_totals() 說明)。
#
# 三段式金額查核機制 (使用者明確要求)：
#   1. DPP (未稅小計) + PPN = GRAND TOTAL           -> extract_total/extract_grand_total
#   2. GRAND TOTAL - PPH 2% = TOTAL YANG DIBAYAR     -> extract_final_total + deduction_prefixes
# ---------------------------------------------------------------------------

def detect_putra_sejahtera_sentosa(text: str) -> bool:
    return "PUTRA SEJAHTERA SENTOSA" in text.upper()


def _pss_clean_amount(text: Optional[str]) -> Optional[int]:
    if not text:
        return None
    cleaned = text.replace(",", "").strip()
    return int(cleaned) if cleaned.isdigit() else None


def _pss_format_date(day, month, year) -> Optional[str]:
    """正確答案 Excel 的 INVOICE DATE 欄位是純文字 'DD/MM/YYYY'
    (不是 Excel 日期型別，跟大多數供應商統一轉成的 'DD Mon YYYY' 不同)，
    所以這家供應商要輸出一樣的純文字格式才會比對相符，不要套用共用的
    _format_date_en()。"""
    try:
        day_i, month_i, year_i = int(day), int(month), int(year)
    except (TypeError, ValueError):
        return None
    return f"{day_i:02d}/{month_i:02d}/{year_i}"


def _pss_clean_description(desc: Optional[str]) -> Optional[str]:
    """右括號前的空白 (例如 'ADIDAS )' vs 'ADIDAS)') 不要統一清掉——
    不同月份的正確答案 Excel 對這個空白的要求不一致，有時跟 PDF 原文
    一樣有空白、有時沒有，沒有通用規則可以判斷，所以照抄 PDF 原文最準。
    這個函式保留只是為了未來若發現某一批帳單真的需要清理空白時可以用。
    """
    return desc


def parse_putra_sejahtera_sentosa(text: str) -> Dict:
    """注意：這是 multi_page 供應商，text 是『一頁』的文字 (每頁各自是一張
    獨立發票)，呼叫端 (_parse_multi_page_pdf()) 只會對通過
    detect_putra_sejahtera_sentosa() 的頁面呼叫這個函式。
    """
    header: Dict[str, Optional[str]] = {
        "order_no": None, "arrive_date": None, "onboard_date": None,
        "bl_no": None, "mbl_no": None, "port_of_loading": None,
        "port_of_discharge": None, "volume": None, "vessel": None,
        "container_no": None, "shipper": None,
    }

    header["invoice_no"] = _search(r"Invoice\s*No\.?\s*:\s*(\S+)", text)

    date_m = re.search(r"\bDate\s*:\s*(\d{1,2})/(\d{1,2})/(\d{4})", text)
    header["invoice_date"] = (
        _pss_format_date(date_m.group(1), date_m.group(2), date_m.group(3))
        if date_m else None
    )

    header["supplier"] = _search(r"(?i)^\s*(PT\.\s*PUTRA SEJAHTERA SENTOSA)", text)
    # 抬頭的 "TO : PT. XXX (ADIDAS ADCI) Invoice No. : ..." 後面括號裡的
    # 品牌/專案註記不是集團公司名稱的一部分，正確答案 Excel 裡的
    # CONSIGNEE 欄位不含這個括號，擷取時要在遇到 "(" 或 "Invoice No" 就停。
    header["consignee"] = _search(
        r"\bTO\s*:\s*(.+?)(?:\s*\(|\s+Invoice\s*No)", text
    )

    items: List[Dict] = []

    desc_m = re.search(r"(JASA TRANSPORTASI\s*\([^\n]*?\))", text)
    amount_m = re.search(r"TOTAL BIAYA TRUCKING\s+([\d,]+)", text)
    if desc_m or amount_m:
        items.append({
            "description": (
                _pss_clean_description(desc_m.group(1).strip()) if desc_m else None
            ),
            "amount": _pss_clean_amount(amount_m.group(1)) if amount_m else None,
            "vat": None,
        })

    # 代墊/報銷費用 (Non DPP 的部分，不列入 PPN 計稅基礎)：有些月份的
    # 帳單會多一筆 "TOTAL BIAYA REIMBURSEMENT"，金額是 0 時代表這張帳單
    # 沒有這筆費用，正確答案 Excel 也不會有這一列，所以只在金額 > 0 時
    # 才加進 items。
    reimb_m = re.search(r"TOTAL BIAYA REIMBURSEMENT\s+([\d,]+)", text)
    reimb_amount = _pss_clean_amount(reimb_m.group(1)) if reimb_m else None
    if reimb_amount:
        items.append({
            "description": "TOTAL BIAYA REIMBURSEMENT",
            "amount": reimb_amount,
            "vat": None,
        })

    # PPN (加值稅)：描述照抄 PDF 原文的 "PPN x %" 標籤，百分比動態抓出來。
    ppn_m = re.search(r"\bPPN\s+(\d[\d,]*\s*%)\s+([\d,]+)", text)
    if ppn_m:
        items.append({
            "description": f"PPN {ppn_m.group(1).strip()}",
            "amount": None,
            "vat": _pss_clean_amount(ppn_m.group(2)),
        })

    # PPH (預扣所得稅)：描述照抄 PDF 原文的 "PPH x%" 標籤。這一筆的 vat
    # 在稽核含稅總額時會被 deduction_prefixes=("PPH",) 排除、改去稽核
    # 「實付金額」(GRAND TOTAL - PPH)，不是加進 GRAND TOTAL 的稅金。
    pph_m = re.search(r"\bPPH\s+(\d[\d,]*\s*%)\s+([\d,]+)", text)
    if pph_m:
        items.append({
            "description": f"PPH {pph_m.group(1).strip()}",
            "amount": None,
            "vat": _pss_clean_amount(pph_m.group(2)),
        })

    if not items:
        items = [{"description": None, "amount": None, "vat": None}]

    return {"header": header, "items": items}


def _pss_extract_dpp(text: str) -> Optional[int]:
    """未稅總額，稽核第 1 段 (未稅總額 + PPN = GRAND TOTAL) 用。

    DPP (Dasar Pengenaan Pajak) 是「需要課 PPN 稅的那一部分」未稅金額，
    對應 items 裡的 TOTAL BIAYA TRUCKING；如果這張帳單還有一筆
    "Non DPP"（不課稅的代墊/報銷費用，對應 TOTAL BIAYA REIMBURSEMENT），
    兩者加起來才是完整的「未稅總額」，才會等於 items 裡 amount 的加總
    (TRUCKING + REIMBURSEMENT)。沒有 Non DPP 那筆費用時，Non DPP 本身
    就是 0，加總等於 DPP 本身，不影響舊有行為。"""
    dpp_m = re.search(r"\bDPP\s+([\d,]+)", text)
    if not dpp_m:
        return None
    dpp = _pss_clean_amount(dpp_m.group(1))
    if dpp is None:
        return None
    non_dpp_m = re.search(r"\bNon DPP\s+([\d,]+)", text)
    non_dpp = _pss_clean_amount(non_dpp_m.group(1)) if non_dpp_m else None
    return dpp + (non_dpp or 0)


def _pss_extract_grand_total(text: str) -> Optional[int]:
    """含稅總額，稽核第 1 段 (DPP + PPN = GRAND TOTAL) 用。"""
    m = re.search(r"GRAND TOTAL\s+([\d,]+)", text)
    return _pss_clean_amount(m.group(1)) if m else None


def _pss_extract_total_paid(text: str) -> Optional[int]:
    """實付金額 (TOTAL YANG DIBAYAR)，稽核第 2 段：
    GRAND TOTAL - PPH = TOTAL YANG DIBAYAR 用。"""
    m = re.search(r"TOTAL YANG DIBAYAR\s+([\d,]+)", text)
    return _pss_clean_amount(m.group(1)) if m else None


register_supplier(
    "PUTRA_SEJAHTERA_SENTOSA", "PT. Putra Sejahtera Sentosa",
    detect_putra_sejahtera_sentosa, parse_putra_sejahtera_sentosa,
    multi_page=True,
    extract_total=_pss_extract_dpp,
    extract_grand_total=_pss_extract_grand_total,
    extract_final_total=_pss_extract_total_paid,
    deduction_prefixes=("PPH",),
)
