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

import calendar
import collections
import difflib
import os
import re
import datetime
from typing import Dict, List, Optional, Tuple

from extract_utils_streamlit import limited_pages, register_supplier, set_supplier_number

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
# 材積 (Volume) 統一寫法 (使用者 2026-10-27 規定，所有供應商適用)：
#   1. 整櫃 (FCL) 有貨櫃數：「數量X櫃型」，多種櫃型用 ", "，例如 "1X40HC"、
#      "1X20GP, 1X40HC"。
#   2. 散貨 (LCL) 沒有貨櫃數：CBM 數字 + " M3"，例如 "7.010 M3"。
#   3. 只有重量 (空運/快遞，沒有 CBM)：重量 + " KG"，例如 "2082 KG"。
# 各供應商只負責抓出貨櫃/CBM/重量，格式一律交給 _fmt_volume()。
# ---------------------------------------------------------------------------

def _fmt_volume(containers=None, cbm=None, kg=None) -> Optional[str]:
    """containers：[(數量, 櫃型)]，櫃型去掉撇號、空白、FT 並轉大寫 ("40'HC" ->
    "40HC"、"20 FT" -> "20")，同櫃型數量合併。cbm / kg：帳單印的數字 (字串或
    數字，呼叫端先把千分位/小數符號整理好)，0 視為沒有。"""
    counts: Dict[str, int] = {}
    for qty, kind in containers or ():
        key = re.sub(r"['\s]|FT\b", "", str(kind).upper())
        try:
            counts[key] = counts.get(key, 0) + int(float(qty))
        except (TypeError, ValueError):
            continue
    parts = [f"{n}X{key}" for key, n in counts.items() if n]
    if parts:
        return ", ".join(parts)
    for value, unit in ((cbm, "M3"), (kg, "KG")):
        if value is None or str(value).strip() in ("", "N/A"):
            continue
        text = str(value).strip()
        try:
            if float(text) == 0:
                continue
        except ValueError:
            continue
        return f"{text} {unit}"
    return None


_VOLUME_BOX_RE = re.compile(r"(\d+)\s*[xX]\s*(\d{2}\s*'?\s*(?:[A-Za-z]{1,3}\b)?)")


def _volume_from_text(raw: Optional[str]) -> Optional[str]:
    """把帳單上的材積文字轉成統一寫法：'FCL - 1x40HC' / '1X20,6X40' /
    '(1x20 FT) (3x40 HQ)' -> 櫃量；'2.5 CBM' / '27.450 M3' -> M3；'614KG' -> KG。
    認不出來的照原文。"""
    if not raw:
        return raw
    boxes = _VOLUME_BOX_RE.findall(raw)
    if boxes:
        return _fmt_volume(containers=boxes)
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*(?:CBM|M3)\b", raw, re.IGNORECASE)
    if m:
        return _fmt_volume(cbm=m.group(1).replace(",", "."))
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*KGS?\b", raw, re.IGNORECASE)
    if m:
        return _fmt_volume(kg=m.group(1).replace(",", "."))
    return raw


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
    # 國內報關 (Shipment Type : Customs) 帳單的欄位
    r"Domestic\s*Origin", r"Domestic\s*Dest", r"Doc\s*No\.?", r"Order\s*Date", r"Customer",
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
    # 幣別有時沒印 (例如 INJK-2504-0454 最後兩筆)，設為選填
    r"^(?P<desc>.+?)\s+(?P<qty>\d+)\s+(?:(?P<curr>IDR|USD)\s+)?"
    r"(?P<rate>[\d,]+)\s+(?P<amount>[\d,]+)"
    r"(?:\s+(?P<vat>[\d,]+))?\s*$"
)
_DWIHARTA_EXCLUDE_KEYWORDS = ("Description", "Sub Total", "GRAND TOTAL", "V A T", "Payment Term",
                              "A/C", "Halaman")  # 銀行帳號列、分頁字樣不是費用


def _dwiharta_clean_volume(value: Optional[str]) -> Optional[str]:
    """Volume 欄位常見寫法 'FCL - 1x40HC'，'FCL -' 只是裝櫃方式前綴，
    實際要的是後面的櫃型/數量，這裡去掉前綴只留 '1x40HC'。
    """
    if not value:
        return value
    # 材積統一寫法 (見 _fmt_volume)："1x40HC" -> "1X40HC"
    return _volume_from_text(
        re.sub(r"^(FCL|LCL)\s*-\s*", "", value.strip(), flags=re.IGNORECASE).strip())


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
        value = _search(pattern, text)
        # 欄位空白時會抓到同一行的下一個欄位 (例如 "Arrive Date : Port Of
        # Discharge : JAKARTA")，值以已知標籤開頭就當作沒有值
        if value and re.match(rf"{_DWIHARTA_NEXT}\s*:", value):
            value = None
        header[code] = value

    header["volume"] = _dwiharta_clean_volume(header.get("volume"))
    for date_code in ("invoice_date", "arrive_date", "onboard_date"):
        header[date_code] = _dwiharta_parse_indo_date(header.get(date_code))

    # 供應商名稱：取自 "Payment can be Transfer to :" 之後的第一個非空白行；
    # 跨頁時中間會夾分頁字樣 "Halaman 1 dari 2" (例如 INJK-2508-0710)，跳過
    m = re.search(r"Payment can be Transfer to\s*:?\s*\n(?:\s*Halaman\b.*\n)*\s*(.+)", text)
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
        shipper = _shipper_m.group(1)
        # 國內報關帳單：Shipper 下一行是 "Doc No. : Consignee : ..." 等其他
        # 欄位 (Container No. 在更下面)，遇到這些標籤就停
        shipper = re.split(r"\n\s*(?:Doc\s*No\.?|Consignee|Domestic\s*(?:Origin|Dest))\s*:",
                           shipper)[0]
        header["shipper"] = re.sub(r"\s+", " ", shipper).strip()

    # 國內報關 (Shipment Type : Customs) 帳單 (使用者提示)：啟運港 = Domestic
    # Origin、目的港 = Domestic Dest；沒有 B/L，B/L NO. 與 MBL NO. 用 Order No.
    if not header.get("port_of_loading"):
        header["port_of_loading"] = _search(
            rf"Domestic\s*Origin{_DWIHARTA_COLON}(.+?){_DWIHARTA_STOP}", text)
    if not header.get("port_of_discharge"):
        header["port_of_discharge"] = _search(
            rf"Domestic\s*Dest{_DWIHARTA_COLON}(.+?){_DWIHARTA_STOP}", text)
    if not header.get("bl_no") and not header.get("mbl_no") and header.get("order_no"):
        header["bl_no"] = header["mbl_no"] = header["order_no"]

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
#     MBL NO. 本來就沒有對應資訊，缺漏後續會自動補 N/A。
#   - 帳單辨識以 KN 帳單 (右上/右下角有 Kuehne Nagel 字樣) 為準，稅務發票
#     不一定每份都有；幣別依各國帳單而定，不寫死 IDR。

def detect_kuehne_nagel(text: str) -> bool:
    upper = text.upper()
    # 越南 KN (VN10，越南文電子發票/附件清單) 另有規則 KN_VN
    if "KUEHNE + NAGEL" in upper and ("HÓA ĐƠN" in upper or "HOÁ ĐƠN" in upper or "VIỆT NAM" in upper):
        return False
    return "KUEHNE" in upper and "NAGEL" in upper


def kn_invoice_page(page: str) -> bool:
    """KN 帳單頁 = Sales Invoice (有 "SALES INVOICE" 與 "INVOICE NO.")。同一份
    PDF 常夾帶空運提單 (Air Waybill)、承運條款、代墊廠商發票等附件，上面也
    印了 Kuehne Nagel，但不是帳單 (使用者規定只擷取帳單頁)。"""
    upper = page.upper()
    return "SALES INVOICE" in upper and "INVOICE NO." in upper


# 空運 Sales Invoice (使用者 2026-10 樣本 JKT1402378、JKT1406381)：
#   "AWB NO(S) : 1066557183 / 910-18077076 FLIGHT NO(S) : WY/850; WY/117"
#   "ORIGIN AIRPORT : JAKARTA DEST. AIRPORT : NEW YORK, NY"
#   "FLIGHT DATE : 06.07.2025 ARRIVAL DATE : 07.07.2025"
#   貨物行 "236 ADIZERO ... 1,823.00 27.136 4523.00" (毛重、材積 CBM、計費重)
#   費用 "OUTLAYS GARUDA 14,503,404 Z" (代墊費，名稱只寫 OUTLAYS，使用者規定)


def _kn_parse_air(text: str, header: Dict) -> List[Dict]:
    """空運帳單：覆寫運輸相關抬頭欄位，回傳費用明細。"""
    charge_line = re.compile(
        rf"^(?P<desc>[A-Z][^\n]*?)\s+(?P<amount>{_KN_AMOUNT})\s+[A-Z0-9]\s*$", re.MULTILINE)
    awb = re.search(r"AWB NO\(S\)\s*:\s*(\S+)\s*/\s*(\S+)", text)
    if awb:
        header["bl_no"], header["mbl_no"] = awb.group(1), awb.group(2)
    header["vessel"] = _search(r"FLIGHT NO\(S\)\s*:\s*(.+?)\s*$", text)
    header["port_of_loading"] = _search(r"ORIGIN AIRPORT\s*:\s*(.+?)\s+DEST\.\s*AIRPORT", text)
    header["port_of_discharge"] = _search(r"DEST\.\s*AIRPORT\s*:\s*(.+?)\s*$", text)
    for code, label in (("onboard_date", "FLIGHT DATE"), ("arrive_date", "ARRIVAL DATE")):
        m = re.search(rf"{label}\s*:\s*(\d{{1,2}})\.(\d{{1,2}})\.(\d{{4}})", text)
        header[code] = _format_date_en(*m.groups()) if m else None
    cargo = re.search(r"VOLUME / CBM[^\n]*\n[^\n]*?([\d,]+\.\d+)\s+([\d.]+)\s+([\d.]+)\s*$",
                      text, re.MULTILINE)
    # 空運有印 CBM，材積寫 CBM (統一寫法見 _fmt_volume)
    header["volume"] = _fmt_volume(cbm=cargo.group(2), kg=cargo.group(1).replace(",", "")) if cargo else None
    header["container_no"] = None
    items = []
    for m in charge_line.finditer(text):
        desc = m.group("desc").strip()
        if desc.upper().startswith(("VAT TOTAL", "SUBTOTAL", "TOTAL")):
            continue
        if desc.upper().startswith("OUTLAYS"):
            desc = "OUTLAYS"   # 代墊廠商名稱不寫 (使用者規定)
        items.append({"description": desc,
                      "amount": _kn_clean_amount_comma_thousands(m.group("amount"))})
    return items


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

# 另一種版面 (2026-10-30 使用者抽查 JKT1365247)：兩個標籤在一行、值在下一行，
#   "P. OF LOADING ETD/ATD"
#   ": JAKARTA : 07.01.2025"
# 先合併成一般版面 "P. OF LOADING : JAKARTA ETD/ATD : 07.01.2025"，後面的規則不用改
_KN_LABEL_PAIRS = [
    (r"FIRST VESSEL(?: NA)?", r"FIRST VOYAGE(?: NU)?"), (r"VESSEL NAME", r"VOYAGE"),
    (r"PL\. OF RECEIPT", r"MOVEMENT TYPE"), (r"P\. OF LOADING", r"ETD/ATD"),
    (r"P\. OF DISCHARGE", r"ETA/ATA"), (r"PL\. OF DELIVERY", r"DANGEROUS GOODS"),
    (r"TERMS OF TRADE", r"INSURAN\. STATUS"),
]
_KN_SPLIT_LABEL_RE = re.compile(
    r"(?m)^(?P<a>" + "|".join(a for a, _ in _KN_LABEL_PAIRS) + r")\s+"
    r"(?P<b>" + "|".join(b for _, b in _KN_LABEL_PAIRS) + r")[ \t]*\n"
    r":[ \t]*(?P<v1>[^\n:]*?)[ \t]*:[ \t]*(?P<v2>[^\n]*)$")


def _kn_join_split_labels(text: str) -> str:
    return _KN_SPLIT_LABEL_RE.sub(
        lambda m: f"{m['a']} : {m['v1']}  {m['b']} : {m['v2']}", text)


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
#
# 幣別不寫死 IDR：KN 各國帳單的本地幣別不同 (印尼 IDR、其他國家可能是
# USD/EUR...)，本地幣金額也可能帶兩位小數 (例如 "1,234.56")，所以金額
# 接受「英式千分位逗號 (可帶兩位小數)」或「三位數以內 + 兩位小數」兩種，
# 仍然不接受沒有逗號/小數點的純數字 (理由同上，避免誤抓單號雜訊)。
_KN_AMOUNT = r"\d{1,3}(?:,\d{3})+(?:\.\d{2})?|\d{1,3}\.\d{2}"
_KN_CHARGE_LINE_PATTERN = re.compile(
    r"^(?P<desc>[A-Z][A-Za-z0-9 /\-\.,()&]+?)\s+"
    r"(?:[\d,]+\.\d{2}\s*[A-Z]{3}\s+)?"
    # 行尾是 VAT 代碼：數字，或 "Z" (N/A OUT OF SCOPE OF VAT，使用者 2026-10 樣本 JKT1394609)，
    # 或 "E" (0 % ZERO RATED EXPORT OF SERVICE，2026-10-30 JKT1418536 DEMURRAGE AND DETENTION)、
    # "O" (JKT1425521)。只收已知代碼 (貨物說明 "PRS 1,072 C" 不是費用)，而且要在同一行
    # (不可跨行：「GW 51.63」下一行的「G」不是代碼)
    rf"(?P<amount>{_KN_AMOUNT})[ \t]+(?:\d|[ZEO])[ \t]*$",
    re.MULTILINE,
)
# 不含稅的帳單 VAT 合計是 "0"，合計行剛好也像費用行，要排除
_KN_TOTAL_LINE_DESC = re.compile(r"^(?:SUBTOTAL|INVOICE\s+(?:VAT\s+)?TOTAL)\b")


def _kn_clean_amount_comma_thousands(text: Optional[str]):
    """英式數字 '1,521,098' / '1,234.56' (逗號=千分位、句點=小數) -> 數字。
    小數部分是 0 時回傳 int (例如 1521098)，否則回傳 float。

    注意：這跟 _kn_clean_amount() 是不同的格式！_kn_clean_amount() 是給
    Faktur Pajak (稅務發票) 上的印尼式數字用的 ('1.521.098,00'，句點千分位
    、逗號小數)；Sales Invoice 費用明細行上的金額則是英式數字 (逗號千分位)。
    兩種格式的千分位/小數符號剛好相反，混用會把 '1,521,098' 誤判成 '1'
    (把第一個逗號當成小數點切斷)，所以分開兩個函式，不要共用。
    """
    if not text:
        return None
    text = text.strip().replace(",", "")
    if not re.fullmatch(r"\d+(?:\.\d+)?", text):
        return None
    value = float(text)
    return int(value) if value.is_integer() else value


# SUBTOTAL / 含稅總額行，幣別代碼用任意三個大寫字母 (不寫死 IDR)：
#   "SUBTOTAL IDR 12,244,622 1,346,908"   (未稅合計 + VAT 合計)
#   "INVOICE TOTAL IDR 13,591,530"         (含稅總額；有些國家寫 AMOUNT DUE)
# 幣別前面可能多一個字母 ("SUBTOTAL N IDR 19,099,286 2,100,921"，2026-10-30 JKT1370073)
_KN_SUBTOTAL_PATTERN = r"SUBTOTAL\s+(?:[A-Z]\s+)?[A-Z]{3}\s+([\d,.]+)\s+([\d,.]+)"
_KN_GRAND_TOTAL_PATTERN = r"(?:INVOICE\s+TOTAL|AMOUNT\s+DUE|TOTAL\s+DUE)\s+[A-Z]{3}\s+([\d,.]+)"


# SUPPLIER：以 KN 帳單「右下角」的公司名稱為準 (稅務發票不是每份 PDF 都有)。
# 頁尾第一行長這樣，最右邊 (行尾) 那個就是右下角欄位：
#   "PT. Kuehne Nagel Indonesia Tel:+62 ... The Energy Building, 5&6 fl. SCBD PT. Kuehne Nagel Indonesia"
# 輸出統一成大寫、去掉開頭的 "PT." 公司型態字樣 -> "KUEHNE NAGEL INDONESIA"。
_KN_FOOTER_NAME = re.compile(
    r"(?:\bPT\.?\s+)?\b(Kuehne\s*\+?\s*Nagel\b[A-Za-z .&()]*?)\s*$",
    re.IGNORECASE | re.MULTILINE,
)


def _kn_extract_supplier(text: str) -> Optional[str]:
    # 優先找頁尾那行 (同一行有 Tel:/Fax:)，找不到才退回任何一行行尾的 KN 公司名稱
    candidates = [
        m for m in _KN_FOOTER_NAME.finditer(text)
        if re.search(r"\b(?:Tel|Fax)\s*:", text[text.rfind("\n", 0, m.start()) + 1:m.end()])
    ] or list(_KN_FOOTER_NAME.finditer(text))
    if not candidates:
        return None
    return re.sub(r"\s+", " ", candidates[0].group(1)).strip(" .,").upper()


# 貨櫃明細 (MARKS & NOS 區塊)：「SEAL NO.」那一行的上一行就是櫃號那一行，
# 同一行接著寫數量與櫃型，例如：
#   "FSCU5030780 1 40' GE CONTAINER SAID TO CONTAIN: 4222.68 49.423"
#   "SEAL NO. 412 CARTON(S)"
# 一個櫃號對應一個櫃型，櫃號填 Container no.，櫃型彙總成 Volume。
_KN_CONTAINER_LINE = re.compile(
    r"^([A-Z]{4}\d{7})\s+(\d+)\s+(\d{2})'\s*([A-Z]{2})\b[^\n]*\n\s*SEAL NO\.",
    re.MULTILINE,
)


def _kn_extract_containers(text: str) -> List[tuple]:
    """回傳 [(櫃號, 數量, 呎數, 櫃型), ...]，依文件順序、去除重複
    (整份內容常被重複印兩次)。"""
    seen = {}
    for m in _KN_CONTAINER_LINE.finditer(text):
        seen.setdefault(m.group(1), (m.group(1), int(m.group(2)), m.group(3), m.group(4)))
    return list(seen.values())


def _kn_container_volume(containers: List[tuple]) -> Optional[str]:
    """整櫃 (FCL) 的櫃量：同樣櫃型合併數量，櫃型只取第一個字母，
    例如 40' GE + 40' HC + 20' GE -> "1X40G+1X40H+1X20G"。"""
    # 材積統一寫法 (見 _fmt_volume)：多種櫃型用 ", " 隔開
    return _fmt_volume(containers=[(qty, f"{size}{kind[0]}") for _, qty, size, kind in containers])


def _kn_extract_po_numbers(text: str) -> Optional[str]:
    """DESCRIPTION OF GOODS 裡的 "PO NO." 清單：PO NO. 下面逐行都是 PO 號碼，
    遇到第一行不是純號碼的 (例如 "INV NO."、"INV."、"HS CODE") 就停止。
    這段常夾著直排水印的單一字母 (C/L/A/T...)，比對前先把它們去掉。"""
    # 也有寫成 "PO NO#" 的 (JKT1365247、JKT1370073)
    m = re.search(r"PO NO[.#][^\n]*\n(.*)", text, re.DOTALL)
    if not m:
        return None
    numbers: List[str] = []
    for line in m.group(1).split("\n"):
        line = re.sub(r"\b[A-Z]\b", "", line).strip()
        if not line:
            continue
        if not re.fullmatch(r"\d{6,}(?:\s+\d{6,})*", line):
            break
        numbers.extend(line.split())
    return ",".join(dict.fromkeys(numbers)) or None


def _kn_extract_charge_items(text: str) -> List[Dict]:
    """從 Sales Invoice 的費用明細行抓出每一筆 {description, amount}。
    抓不到任何一行時回傳空 list，呼叫端會用 Faktur Pajak 的概括描述當備援，
    確保至少有一筆資料而不是整張都是 N/A。
    """
    items = []
    for m in _KN_CHARGE_LINE_PATTERN.finditer(text):
        if _KN_TOTAL_LINE_DESC.match(m.group("desc")):
            continue
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
    text = _kn_join_split_labels(text)
    header: Dict[str, Optional[str]] = {}

    # SUPPLIER 以 KN 帳單右下角為準；只有右下角抓不到時才退回稅務發票的賣方名稱
    header["supplier"] = (
        _kn_extract_supplier(text)
        or _search(r"Pengusaha Kena Pajak:\s*\n\s*Nama\s*:\s*(.+)", text)
    )

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
    # 沒有 COMMERCIAL INVOICE NUMBER 的帳單，改用 DESCRIPTION OF GOODS 裡的
    # "PO NO." 清單 (多個 PO 用 "," 串接，不加空白)。
    header["order_no"] = (
        _search(rf"COMMERCIAL INVOICE NUMBER\s*\n\s*(.+?){_KN_STOP}", text)
        or _kn_extract_po_numbers(text)
    )
    # 使用者 2026-10-30：寫「AS PER ATTACHED」但 PDF 裡沒有附件號碼 (CAK1107369 等) -> N/A
    if header["order_no"] and not re.search(r"\d", header["order_no"]):
        header["order_no"] = _kn_extract_po_numbers(text)
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

    # Volume / Container no.：
    #   - 整櫃 (FCL)：MARKS & NOS 區塊每個「SEAL NO.」上一行是櫃號+櫃型，
    #     櫃號串成 Container no.，櫃型彙總成櫃量 (例如 "1X40G+1X40H+1X20G")
    #     填 Volume，整張帳單每一列都一樣。
    #   - 散貨 (LCL)：沒有櫃號，Volume 填 "AS PER ATTACHED 重量 材積" 的材積數字。
    containers = _kn_extract_containers(text)
    if containers:
        header["container_no"] = ", ".join(c[0] for c in containers)
        header["volume"] = _kn_container_volume(containers)
    else:
        header["container_no"] = None
        m2 = re.search(r"AS PER ATTACHED\s+([\d,\.]+)\s+([\d,\.]+)", text)
        # 散貨：CBM (統一寫法見 _fmt_volume)
        header["volume"] = _fmt_volume(cbm=m2.group(2).replace(",", "")) if m2 else None

    # 這種帳單沒有主提單，留空給 expand_to_rows() 補 N/A
    header["mbl_no"] = None

    if "AWB NO(S)" in text.upper():
        items = _kn_parse_air(text, header)
    else:
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
    # Description 統一填 "VAT"、Amount 留空——因此這裡額外附加一筆這樣的
    # item，讓 expand_to_rows() 展開成獨立的一列，跟正確答案的列結構對齊。
    _kn_subtotal_m = re.search(_KN_SUBTOTAL_PATTERN, text)
    if _kn_subtotal_m:
        vat_amount = _kn_clean_amount_comma_thousands(_kn_subtotal_m.group(2))
        items.append({"description": "VAT", "amount": None, "vat": vat_amount})

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
    # 帳單打錯的年份 (例如 FLIGHT "CX3245/19020025" 的 0025) 不輸出錯誤日期
    if not 1 <= m <= 12 or y < 2000:
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
    code = raw_flight.split("/")[0].strip()
    # FLIGHT 欄只印日期、沒有航班代碼 ("FLIGHT : 19-11-2025"，INV. POU CHEN DESEMBER 2025)：
    # 沒有船名，日期仍由 _indo_parse_flight_date() 取開船日
    if re.fullmatch(r"\d{1,2}-\d{1,2}-\d{4}|\d{8}", code):
        return None
    return code


def _indo_add_space_after_pt(value: Optional[str]) -> Optional[str]:
    """'PT.HARDASES ABADI INDONESIA' -> 'PT. HARDASES ABADI INDONESIA'：
    PDF 版面把 'PT.' 跟公司名稱黏在一起、中間沒有空白，這裡補回一個空白，
    跟正確答案 Excel 裡的寫法對齊 (正確答案 'PT.' 後面有空白)。
    """
    if not value:
        return value
    return re.sub(r"^(PT\.)(?=\S)", r"\1 ", value.strip())


def _indo_parse_flight_date(raw_flight: Optional[str]) -> Optional[str]:
    """從 FLIGHT 欄位原始值反推開船日，支援兩種常見寫法 (帳單日期印錯時照抄，例如
    "CI761/287052025" -> 2025-MAY-87，已登記 validation_streamlit.KNOWN_SOURCE_ISSUES，不要改這裡)：
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
    """英式數字 '416,729' (逗號千分位) -> 416729。美金計價的帳單金額帶小數
    ('8,132,460.00'，[NIKOMAS]_IPE MARET 第 102 頁) -> 8132460。"""
    if not text:
        return None
    text = text.strip().replace(",", "")
    m = re.fullmatch(r"(\d+)(?:\.(\d{1,2}))?", text)
    if not m:
        return None
    if m.group(2) and int(m.group(2)):
        return float(text)
    return int(m.group(1))


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
    r"^(?!Total\b)(?!V\.?A\.?T\.?\s)(?P<desc>\S.*?)\s+(?P<amount>[\d,]+(?:\.\d{2})?)\s*$",
    re.MULTILINE,
)
# V.A.T. 另起一列印出 (例如 "V.A.T. 34,769")，正確答案範本把這一列當成
# 獨立一筆品項：description 固定寫 "VAT"、amount 留空、金額改填在 vat 欄位
# (不是把 VAT 金額當成第五筆品項的 amount)。
_INDO_VAT_LINE_RE = re.compile(r"^V\.?A\.?T\.?\s+(?P<vat>[\d,]+(?:\.\d{2})?)\s*$", re.MULTILINE)
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
# 供應商 4：MAERSK (PT Maersk Logistics Indonesia，海運帳單)
# ===========================================================================
#
# 欄位對應 (依使用者在正確答案 Excel 上方提示列的說明)：
#   INVOICE DATE   "Invoice Date:" (保留帳單原文 '30.Sep.2026' 點分隔格式)
#   INVOICE NO     "Original" 下方 "INVOICE 5617858220" 的號碼
#   SUPPLIER       "Issued by:" 下一行；沒有的話取帳單下方公司名稱
#   CONSIGNEE      "Sold to :"
#   Description/AMOUNT/VAT  "Description of Charges" 表格，每筆費用下一行
#                  緊接著自己的 "VAT xxx"
#   Order No.      "Commercial Invoice No:" (可能有多個號碼、以逗號分隔且跨行)
#   Arrive Date / on board date   "ETA:" / "ETD:"
#   B/L NO.        "FCR" 後面的號碼
#   Port of Loading / discharge   "POL:" / "POD:"
#   Volume         該 FCR 費用的 M3 數量；抓不到才用 "Volume :" 總材積
#   Shipper        "SHPR:"
#   MBL NO.、Vessel、Container no. 帳單沒有，留給 expand_to_rows() 補 N/A。
#
# 一張發票可能包含好幾票貨 (FCR)：每個 "FCR xxx" 開頭的區塊各有自己的
# Sold to/POL/POD/ETD/ETA/SHPR 與費用表格，區塊以費用表格最後一行的區塊
# 小計 (單獨一個數字) 結束。各區塊的運輸欄位寫進該區塊每筆 item，覆蓋整張
# 發票共用的抬頭值 (expand_to_rows() 支援 item 帶抬頭欄位)。
#
# 版面特徵：POL/POD/ETD/ETA/SHPR... 左右並排、只隔一個空白 (跟 INDOPROSTIME
# 同款排版)，用「遇到下一個已知欄位關鍵字就停止」的策略處理。
#
# 金額查核：每筆費用都有自己的 VAT；"Total Amount Due" 小表格的
#   Net value  = 費用金額加總          (extract_total)
#   Amount Due = 費用金額 + VAT 加總   (extract_grand_total)
# 兩項都相符時，VAT 加總也必然等於小表格的 VAT 那一列。
#
# 已知原帳單問題 (使用者核對原帳單確認，不是擷取錯誤，不要改規則去「修」)：
#   - ETA 本身印的就是「--」：Arrive Date 照原樣擷取成 "--"。
#   - ETA 早於 ETD。
# 這兩項登記在 validation_streamlit.KNOWN_SOURCE_ISSUES，驗證時只當附註
# (INFO，不反紅、不影響信心等級)，並註記「屬原帳單問題」。

def detect_maersk(text: str) -> bool:
    """注意：不能只看 'MAERSK' 這個字，因為其他供應商(例如 PANCARAN)的帳單
    裡 VESSEL NAME 欄位常常就是船名含 'MAERSK' (例如 'JOSEPHINE MAERSK')，
    這跟 PT Maersk Logistics Indonesia 這家供應商本身無關，會誤判。改用更
    具體的 'MAERSK LOGISTICS' 字樣判斷。

    另外排除印尼電子稅務發票 (e-Faktur / Faktur Pajak，檔名常見
    'OutputTaxInvoice-...')：賣方同樣是 MAERSK LOGISTICS INDONESIA，但版面
    跟 MAERSK 帳單完全不同，硬套 parse_maersk 只會抓出錯誤資料。暫時歸到
    未知供應商，之後再另外建立電子稅務發票的辨識規則。
    注意：正常 MAERSK 帳單後面也可能附一頁 Faktur Pajak，所以只排除
    「一開頭就是 Faktur Pajak」的單獨電子稅務發票。"""
    upper = text.upper()
    # 只看抬頭：掃描檔 OCR 常把 "Kode dan Nomor Seri" 讀錯，不能靠它判斷
    if upper.lstrip().startswith("FAKTUR PAJAK"):
        return False
    # 越南 MAERSK (VN07，越南文電子發票) 另有規則
    if "MAERSK LOGISTICS & SERVICES VI" in upper or "HÓA ĐƠN GIÁ TRỊ GIA TĂNG" in upper:
        return False
    # DSV 帳單 (頁首 "INVOICE ID610184985 ORIGINAL"、費用表 "CHARGES IN IDR") 內文可能提到
    # MAERSK LOGISTICS，不是 MAERSK 的帳單 (2026-10-30 DSV "INVOICE - ID610184985 …")
    if "CHARGES IN IDR" in upper and re.search(r"^INVOICE\s+ID\d+\s+(?:ORIGINAL|COPY)\b", upper, re.MULTILINE):
        return False
    return "MAERSK LOGISTICS" in upper


_MAERSK_LABEL_WORDS = [
    r"Place\s*of\s*Receipt", r"Place\s*of\s*Delivery", r"SHPR", r"CNEE",
    r"POL", r"POD", r"ETD", r"ETA", r"Sold\s*to", r"Total\s*Packages",
    r"Weight", r"Volume", r"Units",
]
_MAERSK_NEXT = "(?:" + "|".join(_MAERSK_LABEL_WORDS) + ")"
# 跟 INDOPROSTIME 一樣：並排欄位只隔一個空白，靠「下一個已知欄位關鍵字」
# 停止，不能靠連續空白判斷。
_MAERSK_STOP = rf"(?=\s+{_MAERSK_NEXT}\b|\n|$)"

# 英文月份縮寫對照表 (後面其他供應商也共用這份)；也收印尼文縮寫
# (Mei/Agt/Okt/Des，例如 TRANS DAYA PRIMA "21 Agt 2025")
_EN_MONTHS_ABBR = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
    "mei": 5, "agt": 8, "agu": 8, "okt": 10, "des": 12,
}


def _maersk_parse_date(text: Optional[str]) -> Optional[str]:
    """ETD/ETA 的 '29-Aug-2026' -> 'DD Mon YYYY' ('29 Aug 2026')。
    認不出格式就原樣回傳。"""
    if not text:
        return text
    m = re.match(r"(\d{1,2})-([A-Za-z]{3})-(\d{4})", text.strip())
    if not m:
        return text
    return _format_date_en_from_month_name(*m.groups()) or text


def _maersk_extract_invoice_date(text: str) -> Optional[str]:
    """INVOICE DATE 保留帳單原文的點分隔格式 '30.Sep.2026' (正確答案 Excel
    是這種文字儲存格；比對與輸出 Excel 時 _coerce_to_date() 都認得)。"""
    m = re.search(r"Invoice\s*Date\s*:\s*(\d{1,2})\.([A-Za-z]{3,})\.(\d{4})", text)
    if not m:
        return _search(r"Invoice\s*Date\s*:\s*(\S+)", text)
    month_i = _MONTH_NAME_TO_NUM.get(m.group(2).upper()[:3])
    if month_i is None:
        return None
    return f"{int(m.group(1)):02d}.{_MONTH_ABBR_EN[month_i - 1]}.{m.group(3)}"


def _maersk_clean_amount(text: Optional[str]) -> Optional[int]:
    """英式數字 '1,677,125' (逗號千分位、無小數) -> 1677125。"""
    if not text:
        return None
    text = text.strip().replace(",", "")
    return int(text) if text.isdigit() else None


# 費用明細表格 "ItemDescription of Charges ... Qty UoM Total(IDR)"：
#   "1 CFS Receiving Charge 4.094 M3 1,677,125"  -> 項次+名稱+數量+單位+金額
#   "VAT 184,484"                                -> 這筆費用自己的 VAT
#   名稱太長會被換行，例如：
#   "2 Origin Terminal Handling 4.094 M3 361,055"
#   "Charge"                                     -> 續行，併回上一筆名稱
#   "VAT 39,716"
#   未達最低計費量時，金額不在項次那一行，而是另起一行：
#   "9 CFS Receiving Charge 0.850 M3"
#   "VAT 25,830"
#   "Min. Quantity Price 2,348.19 234,819"       -> 最後一個數字才是金額
# 用逐行掃描處理這些不固定的情況。
# 掃描檔 OCR 常把項次讀成 "3." (多一個句點)，所以句點設為選填
_MAERSK_ITEM_START_PATTERN = re.compile(
    r"^(?P<num>\d+)\.?\s+(?P<desc>.+?)\s+(?P<qty>\d+(?:\.\d+)?)\s+"
    r"(?P<uom>[A-Za-z][A-Za-z0-9]*)(?:\s+(?P<amount>[\d,]+))?\s*$"
)
_MAERSK_VAT_LINE_PATTERN = re.compile(r"^VAT\s+([\d,]+)\s*$")
_MAERSK_MIN_PRICE_PATTERN = re.compile(r"^Min\.?\s*Quantity\s*Price\b.*?([\d,]+)\s*$")
# 費用表格最後的區塊小計：整行只有一個數字
_MAERSK_TABLE_END_PATTERN = re.compile(r"^[\d,]+$")


def _maersk_extract_items(block: str) -> Tuple[List[Dict], Optional[str]]:
    """擷取一個 FCR 區塊的費用明細，回傳 (items, 該區塊 M3 數量)。"""
    # 文字層是 "ItemDescription"，掃描檔 OCR 是 "Item Description"
    m = re.search(r"Item\s*Description\s+of\s+Charges[^\n]*\n(.*)", block, re.DOTALL)
    if not m:
        return [], None
    lines = []
    for ln in m.group(1).split("\n"):
        ln = ln.strip()
        if _MAERSK_TABLE_END_PATTERN.match(ln):
            break
        if ln:
            lines.append(ln)

    items: List[Dict] = []
    volume = None
    i = 0
    while i < len(lines):
        start_m = _MAERSK_ITEM_START_PATTERN.match(lines[i])
        if not start_m:
            i += 1
            continue
        if volume is None and start_m.group("uom").upper() == "M3":
            volume = start_m.group("qty")
        desc_parts = [start_m.group("desc").strip()]
        amount = _maersk_clean_amount(start_m.group("amount"))
        vat = None
        i += 1
        # 到下一筆項次開頭之前：VAT 行、最低計費金額行，其餘是名稱續行
        while i < len(lines) and not _MAERSK_ITEM_START_PATTERN.match(lines[i]):
            vat_m = _MAERSK_VAT_LINE_PATTERN.match(lines[i])
            min_m = _MAERSK_MIN_PRICE_PATTERN.match(lines[i])
            if vat_m:
                vat = _maersk_clean_amount(vat_m.group(1))
            elif min_m and amount is None:
                amount = _maersk_clean_amount(min_m.group(1))
            elif vat is None:
                desc_parts.append(lines[i])
            i += 1
        items.append({
            "description": " ".join(desc_parts).strip(),
            "amount": amount,
            "vat": vat,
        })
    return items, volume


def _maersk_split_fcr_blocks(text: str) -> List[str]:
    """依 "FCR xxx" 行切出每一票貨的區塊 (第一個 FCR 之前是發票抬頭)；
    找不到 FCR 行時整份當成一個區塊。"""
    starts = [m.start() for m in re.finditer(r"^FCR\s+\S+\s*$", text, re.MULTILINE)]
    if not starts:
        return [text]
    return [text[s:e] for s, e in zip(starts, starts[1:] + [len(text)])]


def _maersk_order_no(text: str) -> Optional[str]:
    """"Commercial Invoice No:" 後面的號碼；多個號碼時會以逗號分隔並跨行
    (甚至從號碼中間斷行)，所以把到 "Total Amount Due" 之前的文字去掉空白
    直接接起來。"""
    m = re.search(r"Commercial\s*Invoice\s*No\s*:\s*(.*?)\n\s*Total\s*Amount\s*Due",
                  text, re.DOTALL)
    if not m:
        return _search(r"Commercial\s*Invoice\s*No\s*:\s*(\S+)", text)
    return re.sub(r"\s+", "", m.group(1)) or None


def _maersk_extract_total(text: str) -> Optional[int]:
    """"Total Amount Due" 小表格的 'Net value' (未稅總額)。"""
    return _maersk_clean_amount(_search(r"Net\s+value\s+([\d,]+)", text))


def _maersk_extract_grand_total(text: str) -> Optional[int]:
    """"Total Amount Due" 小表格的 'Amount Due' (含稅總額)。"""
    return _maersk_clean_amount(_search(r"Amount\s*Due\s+([\d,]+)", text))


def parse_maersk(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {}

    header["invoice_no"] = _search(r"\bINVOICE\s+(\d+)\b", text)
    header["invoice_date"] = _maersk_extract_invoice_date(text)
    header["order_no"] = _maersk_order_no(text)
    # SUPPLIER：第 2 頁 "Issued by:" 下一行；只有第 1 頁時退而取費用表格
    # 下方 ("Total Packages" 那行之後) 的公司名稱
    header["supplier"] = (
        _search(r"Issued by:\s*\n(.+)", text)
        or _search(r"Total\s*Packages.*\n(.+)", text)
    )
    header["consignee"] = _search(rf"Sold\s*to\s*:\s*(.+?){_MAERSK_STOP}", text)
    total_volume = _search(r"Volume\s*:\s*([\d.,]+)\s*m3", text)

    items: List[Dict] = []
    for n, block in enumerate(_maersk_split_fcr_blocks(text)):
        block_items, block_volume = _maersk_extract_items(block)
        shipment = {
            "bl_no": _search(r"^FCR\s+(\S+)", block),
            "port_of_loading": _search(rf"\bPOL\s*:\s*(.+?){_MAERSK_STOP}", block),
            "port_of_discharge": _search(rf"\bPOD\s*:\s*(.+?){_MAERSK_STOP}", block),
            "onboard_date": _maersk_parse_date(
                _search(rf"\bETD\s*:\s*(.+?){_MAERSK_STOP}", block)
            ),
            "arrive_date": _maersk_parse_date(
                _search(rf"\bETA\s*:\s*(.+?){_MAERSK_STOP}", block)
            ),
            # 材積統一寫法 (見 _fmt_volume)：散貨 CBM
            "volume": _fmt_volume(cbm=(block_volume or total_volume or "").replace(",", "") or None),
            # Shipper：帳單上 "SHPR:" 後面的公司名稱 (跟 "Sold to:" 的帳單
            # 對象可能不是同一家，不要直接沿用 consignee 的值)。
            "shipper": _search(rf"\bSHPR\s*:\s*(.+?){_MAERSK_STOP}", block),
        }
        if n == 0:
            header.update(shipment)  # 抬頭沿用第一票貨的值
        items.extend({**item, **shipment} for item in block_items)

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
    upper = text.upper()
    # TATA 的代墊費用說明也會寫 "REIMBURSEMENT - PT HYPER MEGA SHIPPING (INJK…)"
    # (使用者 2026-10 抽樣 TH20251541)：抬頭是 TATA 就不是 HYPER MEGA
    if "TATA HARMONI" in upper[:300]:
        return False
    return "HYPER MEGA" in upper


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
    # 材積統一寫法 (見 _fmt_volume)："0,5100" -> "0.5100 M3"
    header["volume"] = _fmt_volume(cbm=m.group(1).replace(",", ".")) if m else None

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
    # 材積統一寫法 (見 _fmt_volume)："3 PKG / 32 KG" -> "32 KG"
    header["volume"] = _volume_from_text(_search(rf"\bQTY\s*:\s*(.+?){_TATA_STOP}", text))
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
# 名稱可能有單引號、分號等任何符號 (例如 "LIFT OFF 1x20'"、"INT'L"、
# "MITRA; (LIFT OFF")，所以名稱不限字元；先判斷下面的「項次 (X) 金額」行
_TATA_RECEIPT_FULL_ITEM_LINE = re.compile(
    r"^\d+\s+(?P<desc>[A-Za-z].*?)\s+(?:X\s+)?(?P<amount>[\d.]+)\s*$"
)
# 名稱換行時的金額行：「項次 金額」或「項次 X 金額」(例如 "4 X 600.000")
_TATA_RECEIPT_NUM_AMOUNT_ONLY_LINE = re.compile(
    r"^(?P<num>\d+)\s+(?:X\s+)?(?P<amount>[\d.]+)\s*$")


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
        m_num_only = _TATA_RECEIPT_NUM_AMOUNT_ONLY_LINE.match(line)
        m_full = None if m_num_only else _TATA_RECEIPT_FULL_ITEM_LINE.match(line)
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
    # 材積統一寫法 (見 _fmt_volume)："3CTN/109KG" -> "109 KG"
    header["volume"] = _volume_from_text(_volume_m.group(1).strip()) if _volume_m else None

    # 海運 (使用者提示，例如 SC2608-1058)：表格下方印 "HB/L:" "MB/L:"
    # "SHIPPER:" "ORIGIN:" "VESSEL: ... Total 43.332.814" (右欄 Total 擠在同一行)
    # 與每櫃一行的 "FCL 1x20': EGHU3354377"
    header["bl_no"] = header["bl_no"] or _search(r"\bHB/L\s*:\s*(.+?)\s*\n", text)
    header["mbl_no"] = header["mbl_no"] or _search(r"\bMB/L\s*:\s*(.+?)\s*\n", text)
    header["shipper"] = header["shipper"] or _search(
        r"\bSHIPPER\s*:\s*(.+?)(?:\s+(?:Bayar|Total)\b|\s*\n)", text)
    header["vessel"] = header["vessel"] or _search(
        r"\bVESSEL\s*:\s*(.+?)(?:\s+(?:Bayar|Total)\b|\s*\n)", text)
    header["port_of_loading"] = _search(
        r"\bORIGIN\s*:\s*(.+?)(?:\s+(?:Bayar|Total)\b|\s*\n)", text)
    containers = re.findall(r"(?m)^\s*(?:FCL|LCL)\s+[^:\n]*:\s*([A-Z]{4}\d{7})", text)
    header["container_no"] = ", ".join(containers) or None

    # 這種帳單沒有 Order No./到達日/開船日/目的港，留空給 expand_to_rows() 補 N/A
    header["order_no"] = None
    header["arrive_date"] = None
    header["onboard_date"] = None
    header["port_of_discharge"] = None

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
    return _fmt_volume(kg=f"{value:g}")   # 材積統一寫法 (見 _fmt_volume)


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
    # DHL Express 帳單 (抬頭 "DHL Express … REGULAR INVOICE") 的附件頁可能出現 TATA
    # 的關鍵字 (2026-10-27 抽樣 SRGR003222411 等 8 份被誤認)：看抬頭排除
    head = text[:300].upper()
    if head.lstrip().startswith("DHL EXPRESS") and ("INVOICE" in head or "DEBIT NOTE" in head):
        return False
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
# 費用名稱也會有 "&"、"/" (例如 "Biaya Dokumen & Other charges"、
# "Biaya Bongkar / Muat")，所以 desc 不限字元，只要求英文字母開頭。
_TRANS_COST_ITEM_PATTERN = re.compile(r"^(?P<desc>[A-Za-z].*?)\s+(?P<amount>[\d.]+)\s*$")


def _trans_extract_cost_items(text: str) -> List[Dict]:
    # 兩頁的帳單：Cost Description 在第 1 頁底部，Sub Total 在第 2 頁，
    # 所以也要在頁尾 "Halaman 1 dari 2" 或文字結尾停下 (2026-10-31
    # GSI-1 AGUSTUS 2025 等 74 張雜費全部漏抓，金額勾稽不符)。
    m = re.search(
        r"Cost\s*Description\s*Total\s*\n((?:.*\n)*?)(?=Sub\s*Total\b|Halaman\s+\d+\s+dari\b|\Z)",
        text + "\n",
    )
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
    # 材積統一寫法 (見 _fmt_volume)："(1x20 FT) (3x40 HQ)" -> "1X20, 3X40HQ"
    header["volume"] = _volume_from_text(_femaria_parse_volume(text))
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


def _dwiharta_amount(pattern: str, text: str) -> Optional[int]:
    m = _search(pattern, text)
    m = m.replace(",", "") if m else None
    return int(m) if m and m.isdigit() else None


def _dwiharta_extract_vat_total(text: str) -> Optional[int]:
    """帳單合計區的 "V A T 201,348" (每筆費用各自的 VAT 欄加總)。"""
    return _dwiharta_amount(r"(?m)^[ \t]*V\s?A\s?T[ \t]+([\d,]+)[ \t]*$", text)


def _dwiharta_extract_grand_total(text: str) -> Optional[int]:
    """"E.& O.E GRAND TOTAL 37,479,532" = Sub Total + V A T。"""
    return _dwiharta_amount(r"GRAND\s*TOTAL\s+([\d,]+)", text)


def _kn_extract_total(text: str) -> Optional[int]:
    # "SUBTOTAL IDR 1,521,098 167,321" 是 Sales Invoice 費用明細行上的英式
    # 數字 (逗號千分位、無小數)，要用 _kn_clean_amount_comma_thousands()，
    # 不是給 Faktur Pajak 印尼式數字用的 _kn_clean_amount()。
    return _kn_clean_amount_comma_thousands(_search(_KN_SUBTOTAL_PATTERN, text))


def _kn_extract_grand_total(text: str) -> Optional[int]:
    # "INVOICE TOTAL IDR 1,688,419" 是含稅總額 (= SUBTOTAL + VAT)，用來做
    # 『未稅加總 + VAT 加總 = 含稅總額』的稽核 (見
    # extract_utils.audit_invoice_totals())。跟上面的 SUBTOTAL 是不同行，
    # 用 "INVOICE TOTAL" 這個獨有字樣區分，不會跟 SUBTOTAL 那一行搶到。
    return _kn_clean_amount_comma_thousands(_search(_KN_GRAND_TOTAL_PATTERN, text))


def _indo_extract_grand_total(text: str) -> Optional[int]:
    """注意：INDOPROSTIME (EXPRESS/TRANSPORT) 都是 multi_page 供應商，text
    是『一頁』的文字。帳單上印的 "Total" 是含稅總額 (如果這張帳單有
    V.A.T. 的話，已經把 VAT 加進去了)，所以要註冊成 extract_grand_total
    (跟『擷取到的費用明細加總 + VAT 加總』互相稽核)，不是 extract_total
    (未稅總額)——這家供應商帳單本文沒有單獨印出未稅小計。"""
    return _indo_clean_amount(_search(r"^Total\s+([\d,]+(?:\.\d{2})?)\s*$", text))


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
    """注意：DSV 是 multi_page 供應商，text 是『一頁』的文字。
    頁首有 "INVOICE ID610295705 ORIGINAL/COPY" 時回傳 {發票號碼: 總額}，直接歸給這張。
    費用說明可能提到另一張發票 ("THE ALREADY BILL UNDER ID610295704")，核心用「頁面上
    出現的發票號碼」歸屬時會配錯張 (2026-10-30 GLOSTAR_12DECEMBER：705 的小計配給 704)。"""
    total = _dsv_clean_amount(_search(r"SUBTOTAL\s+([\d.]+)", text))
    inv = _search(r"^INVOICE\s+(ID\d+)\b", text) if total is not None else None
    return {inv: total} if inv else total


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


# --- DWIHARTA Excel 對帳彙總表 rekap (使用者 2026-10-28：要當帳單擷取) --------
# 一張工作表好幾個發票區塊 (2026-10-28 抽樣 "rekap 202602")：發票號碼 INJK-2602-0896
# 在表頭上方，每列一趟車，TRUCKING 到 JUMLAH 之前都是費用，"TOTAL" 列勾稽，
# 下方簽名日期 "Jakarta, 13 Feb 2026" 當發票日期。沒有稅額 (稅在 PDF 發票上)。
def _dwiharta_sheets(book) -> Tuple[List[Dict], List[Dict]]:
    docs, totals = [], []
    for _, rows in book.sheets:
        heads = [i for i, r in enumerate(rows)
                 if {"TRUCKING", "JUMLAH", "FEET"} <= {_vn_text(c).strip().upper() for c in r}]
        for n, head in enumerate(heads):
            stop = heads[n + 1] if n + 1 < len(heads) else len(rows)
            above = " ".join(_vn_text(c) for r in rows[max(0, head - 7):head] for c in r if c not in (None, ""))
            inv = re.findall(r"INJK-\d{4}-\d{4}", above)
            if not inv:
                continue
            labels = [_vn_text(c).strip().upper() for c in rows[head]]
            start, end = labels.index("TRUCKING"), labels.index("JUMLAH")
            cont_col = next((k for k, v in enumerate(labels) if "CONTAINER" in v), None)
            feet_col = labels.index("FEET")
            consignee = next((_vn_text(c).strip() for r in rows[max(0, head - 7):head] for c in r
                              if _vn_text(c).strip().upper().startswith("PT.")), None)

            def label(k):
                sub = _vn_text(_vn_cell(rows[head + 1], k)).strip()
                return re.sub(r"\s+", " ", f"{_vn_text(rows[head][k]).strip()} {sub}").strip()

            items, total, date = [], None, None
            for r in rows[head + 1:stop]:
                text = " ".join(_vn_text(c) for c in r if c not in (None, ""))
                if _vn_num(_vn_cell(r, 0)) is not None and total is None:
                    box = _vn_text(_vn_cell(r, cont_col)).strip() if cont_col is not None else ""
                    feet = _vn_text(_vn_cell(r, feet_col)).strip()
                    extra = {"container_no": box if re.search(r"\d", box) else None,
                             "volume": f"1X{feet}" if re.fullmatch(r"\d{2}'?", feet) else None}
                    for k in range(start, end):
                        amount = _vn_amount(_vn_cell(r, k))
                        if amount:
                            items.append({"description": label(k), "amount": amount, "vat": None, **extra})
                elif total is None and re.search(r"\bTOTAL\b", text.upper()):
                    total = _vn_num(_vn_cell(r, end))
                m = re.search(r"Jakarta\s*,\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", text, re.I)
                if m and date is None:
                    month = INDO_MONTHS.get(m.group(2).lower()) or next(
                        (i for i in range(1, 13) if datetime.date(2000, i, 1).strftime("%b").lower()
                         == m.group(2)[:3].lower()), None)
                    if month:
                        date = datetime.date(int(m.group(3)), int(month), int(m.group(1))).strftime("%d %b %Y")
            if not items:
                continue
            docs.append({"header": {"invoice_no": inv[-1], "supplier": "PT. DWIHARTA LOGISTINDO",
                                    "consignee": consignee, "invoice_date": date}, "items": items})
            totals.append({"invoice_no": inv[-1], "total": total, "grand_total": None, "final_total": None})
    return docs or [{"header": {}, "items": [{"description": None, "amount": None}]}], totals


register_supplier(
    "DWIHARTA", "PT. DWIHARTA LOGISTINDO", detect_dwiharta, parse_dwiharta,
    extract_total=_dwiharta_extract_total,
    extract_vat_total=_dwiharta_extract_vat_total,
    extract_grand_total=_dwiharta_extract_grand_total,
    sheet_parse=lambda book: _dwiharta_sheets(book)[0], sheet_totals=lambda book: _dwiharta_sheets(book)[1],
)
register_supplier(
    "KUEHNE_NAGEL", "KUEHNE NAGEL INDONESIA", detect_kuehne_nagel, parse_kuehne_nagel,
    extract_total=_kn_extract_total, extract_grand_total=_kn_extract_grand_total,
    invoice_page=kn_invoice_page,
    # FA Single Invoice 後面約 20 頁是掃描附件：不 OCR (一份 1~7 分鐘 → 1~2 秒)
    ocr_attachments=False,
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
    # 材積統一寫法 (見 _fmt_volume)：散貨用 "Measurement : 3" (CBM)，沒有才用
    # "Gross Weight : 1296" (KG)；"Party : 48 ROLLS" 是件數，不是材積
    header["volume"] = _fmt_volume(cbm=_search(r"Measurement\s*:\s*([\d.]+)", text),
                                   kg=_search(r"Gross Weight\s*:\s*([\d.]+)", text))
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
    # 陸運帳單 (INV.2510.0007~0012) 的 Container No. 欄印車牌 ("B 9197 PCC,")：照抄，不要改擷取規則。
    # 使用者 2026-10-30 確認是原帳單問題，登記在 validation_streamlit.KNOWN_SOURCE_ISSUES["PANCARAN"]
    lines = text.split("\n")
    for i, line in enumerate(lines):
        m = re.search(r"JOB\s+NO\.\s*:\s*\S+\s+(.+)$", line)
        if not m:
            continue
        parts = [m.group(1).strip()]
        prev = lines[i - 1].strip() if i else ""
        # 貨櫃清單在 JOB NO. 上一行，最後一個號碼被切開
        # ("…,DENU9999888,TCLU" + "JOB NO. : EXPJK25110020 6576305,"，使用者 2026-10 樣本 INV.2511.0045)；
        # 也會切在英文字母中間 ("…,MTK" + "U4230035," INV.2505.0004、"…,BM" + "OU4893979," INV.2511.0044)：
        # 兩段英文字母合起來 4 碼才接
        tail = re.search(r"(?:^|,)([A-Z]{1,4})$", prev)
        head = re.match(r"([A-Z]{0,3})\d{7}\b", parts[0])
        if (tail and head and re.fullmatch(r"[A-Z0-9,]+", prev)
                and len(tail.group(1)) + len(head.group(1)) == 4):
            parts = [(prev + parts[0]).rstrip(",")]
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
    # REFERENCE NO. 只寫裝載方式 (例如 "LCL"，沒有任何數字) 時不是單號，
    # Order No. 改用 JOB NO. (使用者提示，例如 INV.2509.0051)
    if header["order_no"] and not re.search(r"\d", header["order_no"]):
        header["order_no"] = header["bl_no"]
    # 材積統一寫法 (見 _fmt_volume，使用者 2026-10-27)：卡車車型也照「數量X車型」，
    # "LCL - 13x40HC LCL - 2xBOXCDDLCL - 1xBOXFUSO" -> "13X40HC, 2XBOXCDD, 1XBOXFUSO"
    raw_volume = _pancaran_extract_volume(text)
    boxes = re.findall(r"(\d+)\s*[xX]\s*([A-Za-z0-9]+?)(?=LCL|FCL|[^A-Za-z0-9]|$)", raw_volume or "")
    header["volume"] = _fmt_volume(containers=boxes) or raw_volume
    header["shipper"] = _search(r"(?m)^SHIPPER\s*:\s*(.+?)(?:;|\s*$)", text)

    pol_pod = _search(r"(?m)^POL\s*/\s*POD\s*:\s*(.+)$", text)
    if pol_pod and "/" in pol_pod:
        pol, pod = pol_pod.split("/", 1)
        header["port_of_loading"] = pol.strip() or None
        header["port_of_discharge"] = pod.strip() or None
    else:
        header["port_of_loading"] = pol_pod
        header["port_of_discharge"] = None

    # 沒有船名時帳單印 ". / ." 或 "- / -"：照抄原文 (使用者答案)，驗證層把
    # 「船名沒有文字」列為 PANCARAN 的例外
    header["vessel"] = _search(r"(?m)^VESSEL\s+NAME\s*:\s*(.+)$", text)

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

def _is_dhl_vnpt(text: str) -> bool:
    """越南 DHL-VNPT 快遞 (VN04) 的帳單：別家的 detect 都要排除 (附件裡常出現 APL、DSV
    等託運人名稱，2026-10-28 抽樣)。"""
    return re.search(r"DHL\s*-?\s*VNPT", text.upper()) is not None


def detect_dhl(text: str) -> bool:
    upper = text.upper()
    if _is_dhl_vnpt(text):
        return False
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
    # 地名可能有括號、句點、撇號、連字號 ("MAA, CHENNAI (MADRAS)"，2026-10-27 抽樣)
    r"(?P<ocode>[A-Z]{2,4}),\s*(?P<oregion>[A-Za-z][A-Za-z ().'\-]*?)\s+"
    r"(?P<dcode>[A-Z]{2,4}),\s*(?P<dregion>[A-Za-z][A-Za-z ().'\-]*?)\s+EXPRESS\s+"
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
            # 材積 = 每票的 Weight in Kg (使用者 2026-10-27 答案 "14.5 KG")
            "volume": _fmt_volume(kg=f"{float(m.group('weight')):g}"),
        }

        # 每筆費用自帶 VAT (使用者 2026-10-27 答案)
        items.append({
            "description": service_name,
            "amount": _dhl_clean_amount(m.group("charge")),
            "vat": _dhl_clean_amount(m.group("vat")),
            **overrides,
        })
        if extra_pattern:
            for em in extra_pattern.finditer(block):
                items.append({
                    "description": em.group("desc").strip(),
                    "amount": _dhl_clean_amount(em.group("amount")),
                    "vat": _dhl_clean_amount(em.group("vat")),
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

    # 彙總表可能有好幾種服務 (EXPRESS WORLDWIDE DOC / NONDOC / DOMESTIC，2026-10-27
    # 抽樣)：從表頭到 "Total" 列之間每一列都是一筆
    in_summary = False
    for line in lines:
        stripped_line = line.strip()
        if stripped_line.startswith("Type of Service"):
            in_summary = True
            continue
        # OCR 掃描檔的表頭常讀壞 (in_summary 沒設到)：抓到服務列之後遇到 Total 也停，
        # 不然後面 "NPWP 0013 1077 4306 2000" 會被當成一筆 (f008 抽樣)
        if (in_summary or items) and stripped_line.upper().startswith("TOTAL"):
            break
        m = _DHL_CUSTOMS_SUMMARY_LINE_PATTERN.match(stripped_line)
        if not m:
            continue
        # 前 4 個 token 固定是 Number of Shipments / Total Weight /
        # Number of Items / Standard Shipping Charge，第 4 個 (index 3)
        # 才是我們要的金額，後面欄位數量不一定 (可能混著雜訊 "=") 所以不管。
        tokens = [t for t in m.group(2).split() if t != "="]
        if len(tokens) < 4:
            continue
        amount = _dhl_clean_amount(tokens[3])
        # 金額 0 (DUTIES & TAXES) 的 VAT 是 0；其他服務的 VAT 由 _dhl_allocate_vat() 分配
        items.append({"description": m.group(1).strip(), "amount": amount if amount is not None else 0,
                      "vat": 0 if not amount else None})
    # 第 2 頁每筆額外費用 "DUTY TAX PAID 290,000 293,190" = 名稱、金額、含稅小計：
    # VAT = 含稅小計 - 金額 (使用者 2026-10-27 答案：DUTY TAX PAID VAT 3,190，其餘 0)
    line_totals = {}
    for line in lines:
        for mm in re.finditer(r"([A-Z][A-Z &]+[A-Z])\s+([\d,]+)\s+([\d,]+)\s*$", line.strip()):
            amt, tot = _dhl_clean_amount(mm.group(2)), _dhl_clean_amount(mm.group(3))
            if amt is not None and tot is not None and tot >= amt:
                line_totals[(mm.group(1).strip(), amt)] = tot - amt

    in_extra_section = False
    # OCR 讀壞 "Analysis of Extra Charges" 標題時，從彙總表的 Total 列下一行開始
    has_title = any(ln.strip().upper().startswith("ANALYSIS OF EXTRA CHARGES") for ln in lines)
    for line in lines:
        stripped = line.strip()
        if stripped.upper().startswith("ANALYSIS OF EXTRA CHARGES") or (
                not has_title and not in_extra_section and items
                and re.match(r"Total\s+\d", stripped)):
            in_extra_section = True
            continue
        if not in_extra_section:
            continue
        if stripped.upper().startswith("TOTAL EXTRA CHARGES"):
            break
        # OCR 雜訊："FUELSURCHARGE —— 1,756,394 ," -> "FUEL SURCHARGE 1,756,394"
        stripped = re.sub(r"[\s—–_|.,-]+$", "", stripped)
        stripped = re.sub(r"\s+[—–_|-]+\s+", " ", stripped)
        stripped = stripped.replace("FUELSURCHARGE", "FUEL SURCHARGE")
        m = _DHL_CUSTOMS_ITEM_PATTERN.match(stripped)
        if not m:
            continue
        desc = m.group(1).strip()
        if desc in ("TOTAL EXTRA CHARGES", "TOTAL DISCOUNTS", "TOTAL VAT"):
            continue
        amount = _dhl_clean_amount(m.group(2))
        vat = next((v for (name, amt), v in line_totals.items()
                    if amt == amount and name.endswith(desc)), None)
        items.append({"description": desc, "amount": amount, "vat": vat})
    return items


_DHL_VAT_CODE_RE = re.compile(
    r"^[A-Z][a-z]?\s+[A-Za-z][A-Za-z ]*?\s+(?P<rate>\d+(?:\.\d+)?)%\s+(?P<base>[\d,]+)\s+(?P<vat>[\d,]+)\s*$",
    re.MULTILINE)


def _dhl_allocate_vat(text: str, items: List[Dict]) -> None:
    """只有第 1 頁彙總表 (沒有逐票明細) 時，費用沒有自己的 VAT：用「Analysis of
    VAT」每個稅別的稅基 (例如 "A Taxable 1.10% 580,000 6,380") 找出加總剛好等於
    稅基的費用，稅額分給這些費用 (多筆時依金額比例，尾差放最後一筆)；其餘費用
    VAT = 0 (使用者 2026-10-27 答案)。找不到對應的費用時，稅額單獨一筆。"""
    from itertools import combinations
    pending = [it for it in items if it.get("vat") is None and it.get("amount")]
    if not pending:
        for it in items:
            if it.get("vat") is None:
                it["vat"] = 0
        return
    codes = [(m.group("base"), m.group("vat")) for m in _DHL_VAT_CODE_RE.finditer(text)]
    if not codes:
        # 稅別那行讀不到稅額 (OCR)：用第 1 頁 "Total Amount (IDR) 未稅 VAT 含稅"
        m = re.search(r"Total\s+Amount\s*\(IDR\)\s+([\d,]+)\s+([\d,]+)\s+[\d,]+", text)
        codes = [m.groups()] if m else []
    for base_text, vat_text in codes:
        base, vat = _dhl_clean_amount(base_text), _dhl_clean_amount(vat_text)
        if not base or not vat:
            continue
        group = tuple(pending) if sum(it["amount"] for it in pending) == base else None
        for size in range(1, min(len(pending), 6) + 1) if not group else ():
            group = next((g for g in combinations(pending, size)
                          if sum(it["amount"] for it in g) == base), None)
            if group:
                break
        if not group:
            items.append({"description": "VAT", "amount": None, "vat": vat})
            continue
        left = vat
        for it in group[:-1]:
            it["vat"] = round(vat * it["amount"] / base)
            left -= it["vat"]
        group[-1]["vat"] = left
        pending = [it for it in pending if it not in group]
    for it in items:
        if it.get("vat") is None and it.get("amount") is not None:
            it["vat"] = 0


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
    _dhl_allocate_vat(text, items)
    if not items:
        items = [{"description": None, "amount": None}]

    return {"header": header, "items": items}


# (C) INBOUND DUTY INVOICE (JKTZR…，2026-10-27 抽樣)：一票，表頭 "Waybill Origin Shipper
#     Destination…"，第一行 "5087356190 SGN, HO CHI MINH JKT, JAKARTA"；費用行
#     "07-05-2026 IMPORT EXPORT DUTIES 370,461 z 370,461" (名稱、金額、稅別、[VAT]、含稅)，
#     左邊可能夾著 "Description of Goods: SLAB TEST: 09 SET " 這段貨物說明。
_DHL_DUTY_ITEM_RE = re.compile(
    r"(?P<desc>[A-Z][A-Z &/]*[A-Z])\s+(?P<amount>[\d,]+)\s+(?P<code>[A-Za-z]{1,2})\s+"
    r"(?:(?P<vat>[\d,]+)\s+)?(?P<total>[\d,]+)\s*$")


def _dhl_is_duty(text: str) -> bool:
    # OCR 可能讀掉標題 (f002 抽樣)：也看費用表頭 "Charge Description"
    return ("DUTY INVOICE" in text[:400].upper() or "Charge Description" in text) \
        and "Type of Service" not in text


def _parse_dhl_duty(text: str) -> Dict:
    header = _dhl_common_header(text)
    m = re.search(r"^(\d{8,12})\s+([A-Z]{3},\s*[A-Z][A-Z .()'-]*?)\s+([A-Z]{3},\s*[A-Z][A-Z .()'-]*?)\s*$",
                  text, re.MULTILINE)
    if m:
        header.update(bl_no=m.group(1), port_of_loading=m.group(2), port_of_discharge=m.group(3))
    items: List[Dict] = []
    for line in text.splitlines():
        line = line.strip()
        if line.upper().startswith(("ANALYSIS OF VAT", "TOTAL")):
            if items:
                break
            continue
        date = re.match(r"(\d{1,2}-\d{1,2}-\d{4})\s+", line)
        if date:
            header.setdefault("onboard_date", _dhl_parse_date(date.group(1)))
            line = line[date.end():]
        # 左欄的貨物說明 ("Description of Goods: SLAB TEST: 09 SET ") 去掉
        # (OCR 會把 Description 讀成 "Detcription"，所以只認 "of Goods:")
        line = re.sub(r"^.*?of Goods:.*?\d+\s*(?:SETS?|PCS?|CTNS?|PKGS?|UNITS?|EA)\s+(?:0\s+)?",
                      "", line, flags=re.I)
        line = re.sub(r"^\s*0?\s?ISBURSEMENT\b", "DISBURSEMENT", line)   # OCR 掉了 D
        im = _DHL_DUTY_ITEM_RE.search(line)
        if not im:
            continue
        amount = _dhl_clean_amount(im.group("amount"))
        total = _dhl_clean_amount(im.group("total"))
        vat = _dhl_clean_amount(im.group("vat")) if im.group("vat") else (
            total - amount if amount is not None and total is not None else None)
        items.append({"description": im.group("desc").strip(), "amount": amount, "vat": vat})
    for code in ("order_no", "arrive_date", "mbl_no", "volume", "vessel", "container_no"):
        header.setdefault(code, None)
    return {"header": header, "items": items or [{"description": None, "amount": None}]}


def parse_dhl(text: str) -> Dict:
    if _dhl_is_duty(text):
        result = _parse_dhl_duty(text)
    elif _dhl_is_regular(text) and (regular := _parse_dhl_regular(text)) is not None:
        result = regular
    else:
        result = _parse_dhl_scanned_customs(text)
    # OCR 把抬頭讀成 "PT GIROTIKA SEMESTA" 時供應商會抓不到：供應商固定
    result["header"]["supplier"] = result["header"].get("supplier") or "PT BIROTIKA SEMESTA / DHL EXPRESS"
    return result


def _dhl_extract_grand_total(text: str) -> Optional[int]:
    """含稅總額 = 每一票最後單獨一行的總費用 (例如 "2,917,670" = 該票費用 + VAT)
    加總 (使用者 2026-10-27 提示：每一票的項目總額 + VAT = 每一票的總費用)；
    找不到就用第 1 頁 Total Amount (IDR) 的含稅總額。"""
    # 只看逐票明細 (第一個 "Air Waybill" 表頭到最後一個 "Service Sub Total")，
    # 不吃到後面附件 (報關單、商業發票) 裡單獨一行的數字
    start = text.find("Air Waybill")
    end = text.rfind("Service Sub Total")
    detail = text[start:end] if 0 <= start < end else ""
    per_shipment = re.findall(r"^\s*(\d{1,3}(?:,\d{3})+)\s*$", detail, re.MULTILINE)
    if per_shipment:
        return sum(_dhl_clean_amount(v) for v in per_shipment)
    m = re.search(r"Total\s+Amount\s*\(IDR\)\s+[\d,]+\s+[\d,]+\s+([\d,]+)", text)
    return _dhl_clean_amount(m.group(1)) if m else None


register_supplier(
    "DHL", "PT BIROTIKA SEMESTA / DHL EXPRESS",
    detect_dhl, parse_dhl,
    extract_total=_dhl_extract_total, extract_grand_total=_dhl_extract_grand_total,
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
    """每張發票除了正本 ("INVOICE ID610286388 ORIGINAL") 還附一頁內容相同的
    副本 ("… COPY")。只有副本的頁不算 (逐頁擷取時會略過)，不然發票號碼抓
    不到、費用也會重複算一次。"""
    if "CHARGES IN IDR" not in text.upper():
        return False
    copy_only = (re.search(r"^INVOICE\s+\S+\s+COPY\b", text, re.MULTILINE)
                 and not re.search(r"^INVOICE\s+\S+\s+ORIGINAL\b", text, re.MULTILINE))
    return not copy_only


_DSV_ITEM_PATTERN = re.compile(
    # 舊版 PPN 欄前面有稅率 "11%=333.313" (2026-10-30 DSV 2024-12 "Multiple Invoices … Total 2_6")
    r"^(?P<desc>.+?)\s+-\s+.*?\s+(?:\d+%=)?(?P<ppn>[\d.]+)\s+(?P<amount>[\d.]+)$",
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


# 有些費用名稱後面沒有 " - 說明" ("Detention 598.541 5.441.280"，使用者 2026-10
# 樣本 ID610196477)；只在 CHARGES 表格裡用，避免抓到其他兩個數字結尾的行
_DSV_ITEM_PLAIN_PATTERN = re.compile(
    # 免稅費用的 PPN 欄寫 "Exempt" ("Reimbursement Storage JAS Exempt 265.243")
    r"^(?P<desc>[A-Za-z].*?)\s+(?:\d+%=)?(?P<ppn>\d{1,3}(?:\.\d{3})*|Exempt)\s+(?P<amount>\d{1,3}(?:\.\d{3})*)$",
    re.MULTILINE,
)


def _dsv_extract_items(text: str) -> List[Dict]:
    items: List[Dict] = []
    matches = list(_DSV_ITEM_PATTERN.finditer(text))
    # 續頁沒有 "TOTAL CHARGES"，到 "Continued Over" 為止 (抽樣 ID610286388)
    sec = re.search(r"CHARGES IN IDR\s*\n(.*?)\n\s*(?:TOTAL CHARGES|Continued Over)", text, re.DOTALL)
    if not matches:
        matches = list(_DSV_ITEM_PLAIN_PATTERN.finditer(sec.group(1))) if sec else []
    elif sec:
        # 同一張也可能混著沒有 " - 說明" 的費用 ("B/L (Bill of Lading) 16.500 150.000"，
        # 2026-10-28 抽樣 ID610199356)：表格裡其他行也用簡單格式補上，依原文順序
        dash_lines = {m.group(0) for m in matches}
        plain = [m for m in _DSV_ITEM_PLAIN_PATTERN.finditer(sec.group(1)) if m.group(0) not in dash_lines]
        if plain:
            order = sec.group(1)
            matches = sorted([m for m in matches if m.group(0) in order] + plain,
                             key=lambda m: order.find(m.group(0))) + [m for m in matches if m.group(0) not in order]
    for m in matches:
        items.append({
            # 名稱後面的費率說明不要 ("Overtime charge IDR 175000,00/hours/truck")
            "description": re.sub(r"\s+IDR\s+[\d.,]+/.*$", "", m.group("desc").strip()),
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

    # 欄位空白時下一行就是 "GOODS DESCRIPTION" 標題，不是單號 (2026-10-30 ID610184985)；
    # 撇號也可能是 "OWNER’'S"
    # 整行照抄 (使用者 2026-10-30 決定，例如 "FIAT, OCTOBER, PWJ")
    header["order_no"] = _search(r"OWNER\S{0,3}S REFERENCE\s*\n\s*(?!GOODS DESCRIPTION)(.+)", text)

    # 港口照抄整段 "IDJKT = Jakarta, Java, Indonesia" (使用者 2026-10-25 規定)。
    # ETA 可能沒有印 ("… 28-Feb-25 ZAJNB = Johannesburg, South Africa"，樣本 ID610160512)
    m = re.search(
        r"^([A-Z]{5}\s*=\s*[^\n]+?)\s+(\d{1,2}-[A-Za-z]{3}-\d{2})\s+([A-Z]{5}\s*=\s*[^\n]+?)"
        r"(?:\s+(\d{1,2}-[A-Za-z]{3}-\d{2}))?\s*$",
        text, re.MULTILINE,
    )
    if m:
        # 名稱太長會換行，行尾留下逗號 ("MYPKG = Port Klang (Pelabuhan Klang),")
        header["port_of_loading"] = m.group(1).strip(" ,")
        header["onboard_date"] = _dsv_parse_date(m.group(2))
        header["port_of_discharge"] = m.group(3).strip(" ,")
        header["arrive_date"] = _dsv_parse_date(m.group(4))
    else:
        header["port_of_loading"] = None
        header["onboard_date"] = None
        header["port_of_discharge"] = None
        header["arrive_date"] = None

    # 船名/提單只在 "… HOUSE BILL OF LADING" 表頭到 "ORIGIN" 之間找，避免抓到頁尾網址。
    # 併櫃 (LCL) 只有一個號碼 "/ EDC1118207"：那是 House B/L，船名、MBL 填 N/A
    # (使用者 2026-10 樣本 ID610160512)
    header["vessel"] = header["bl_no"] = header["mbl_no"] = None
    sec = re.search(r"HOUSE BILL OF LADING\s*\n(.*?)\n\s*ORIGIN\b", text, re.DOTALL)
    if sec:
        m = _DSV_VESSEL_BL_PATTERN.search(sec.group(1))
        if m:
            header["vessel"] = f"{m.group('vessel').strip()} / {m.group('voyage').strip()}"
            header["bl_no"] = m.group("house_bl").strip()
            header["mbl_no"] = m.group("ocean_bl").strip()
        else:
            m = re.search(r"^\s*/\s*([A-Z0-9]+)\s*$", sec.group(1), re.MULTILINE)
            header["bl_no"] = m.group(1) if m else None

    # 貨櫃號碼只在 "CONTAINERS" 區塊找；沒有這個區塊 (併櫃) 時填 N/A，
    # 不要抓到費用名稱 "THC (Terminal Handling Charge)"
    # 只收標準貨櫃號碼 (4 個字母 + 7 個數字)；空運的貨盤 "176-22905820 (LD-1…)" 不是貨櫃，填 N/A
    sec = re.search(r"^CONTAINERS\s*\n(.*?)(?:\n\s*(?:Total CO|CO₂|CHARGES)\b|\Z)",
                    text, re.DOTALL | re.MULTILINE)
    m = re.search(r"^([A-Z]{4}\d{7})\s*\(", sec.group(1), re.MULTILINE) if sec else None
    header["container_no"] = m.group(1) if m else None

    m = re.search(r"[\d.]+\s*KG\s+([\d.]+)\s*M3", text)
    # VOLUME 欄被其他文字蓋住讀不到時，用費用說明的 "1,001 Cubic Meter(s)"
    m_cbm = None if m else re.search(r"(\d+),(\d{3})\s+Cubic Meter", text)
    # 材積統一寫法 (見 _fmt_volume)：整櫃數 "PCIU9246720 (40HC - …) / FCL" ->
    # "1X40HC"；散貨寫 CBM
    fcl = re.findall(r"^\S+\s*\((\w+)\s*-[^)]*\)\s*/\s*FCL", sec.group(1), re.MULTILINE) if sec else []
    header["volume"] = _fmt_volume(
        containers=[(1, kind) for kind in fcl],
        cbm=(m.group(1) if m else f"{m_cbm.group(1)}.{m_cbm.group(2)}" if m_cbm else None))

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
    upper = text.upper()
    # 越南 APL (VN11，越南文電子發票) 另有規則 APL_VN
    if "APL LOGISTICS VIỆT NAM" in upper or "HÓA ĐƠN GIÁ TRỊ GIA TĂNG" in upper or _is_dhl_vnpt(text):
        return False
    # DSV 越南帳單的出貨方可能是 "APL Logistics Vietnam Company Limited"
    # (2026-10-28 DSV AIR & SEA_JUL.2026_BY SHIPMENT.pdf 第 169 頁)，不是 APL 開的帳單
    if "DSV AIR & SEA" in upper:
        return False
    return "APL LOGISTICS" in upper


def apl_invoice_page(page: str) -> bool:
    """APL 帳單頁 = 有 "BILLED TO" 的頁。同一份 PDF 常夾帶 Faktur Pajak、
    寶成的領料單 (Items release form) 等附件 (使用者 2026-10 樣本 2025-1217-0135)。"""
    upper = page.upper()
    return "BILLED TO" in upper and not upper.lstrip().startswith("FAKTUR PAJAK")


# 名稱可能有 "%" (例如 "OVERNIGHT 100%")
_APL_ITEM_PATTERN = re.compile(
    # 名稱也可能有撇號 ("THC - 96'"，2026-10-28 抽樣 INV. 2025-0221-0084)
    r"^(?P<desc>[A-Z][A-Z0-9 /\-\.%']*?)\s+(?:[\d,]+\.\d{2}\s+)?(?P<amount>[\d,]+\.\d{2})\s*$",
    re.MULTILINE,
)


def _apl_clean_amount(text: Optional[str]):
    """英式數字 '700,000.00' -> 700000；有小數時保留 ('180,995.52' ->
    180995.52)：雙貨幣排版的 IDR 金額常有小數，逐筆四捨五入會讓加總跟
    Subtotal 差幾元 (例如 2025 年 FCR 131065 那張差 2 元)。"""
    if not text:
        return None
    text = text.strip().replace(",", "")
    try:
        value = round(float(text), 2)
    except ValueError:
        return None
    return int(value) if value.is_integer() else value


def _apl_extract_items(text: str) -> List[Dict]:
    # 表頭可能是粗體疊印 "DDeessccrriippttiioonn" (2026-10-28 抽樣 INV. 2025-0307-0099)
    m = re.search(r"(?:Description|DDeessccrriippttiioonn)\s+Amount.*?\n(.*?)\nSubtotal", text, re.DOTALL)
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

    # 掃描檔 OCR 會把 logo 讀成同一行 ("APL PT. APL LOGISTICS")，不限行首
    header["supplier"] = _search(r"(PT\.\s*APL\s*LOGISTICS)", text)

    # 掃描檔 OCR 會在 "BILLED TO:" 後面多讀出雜訊 ("BILLED TO: “et :")
    m = re.search(r"BILLED TO\s*:[^\n]*\n\s*(.+?)\s+NO\.\s*(\S+)", text)
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
    invoice_page=apl_invoice_page,
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
# 掃描檔 OCR 會把千分位句點讀成逗號 ("-2.528,800")，金額接受兩種符號
# 稅額名稱大小寫不一定 ("PPN 11%"、"PPn 11%"、"PPh 2%")
_GALUH_TAX_LINE_PATTERN = re.compile(r"^(?P<desc>PP[NnHh]\s+\d+%)\s+(?P<amount>-?[\d.,]+)\s*$")
_GALUH_ITEM_LINE_PATTERN = re.compile(
    r"^\d+\s+(?P<desc>.+?)\s+(?P<amount>\d{1,3}(?:\.\d{3})*)\s*$"
)
# 掃描版：從 "Tgl. Invoice" 那行之後到 "Terbilang" (金額大寫) 為止。表頭
# "KETERANGAN JUMLAH" 有時 OCR 讀不到，Total 那行常是亂碼，都不當邊界
_GALUH_SCAN_BLOCK_PATTERN = re.compile(
    r"Tgl\.\s*Invoice[^\n]*\n(.*?)\n[^\n]*Terbilang", re.DOTALL | re.IGNORECASE)
_GALUH_SCAN_ITEM_LINE_PATTERN = re.compile(
    r"^(?P<desc>[A-Za-z].+?)\s+(?P<amount>\d{1,3}(?:\.\d{3})+)\s*$")
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
        # 也去掉日期區間 ("Trucking 06 - 14 April 2026")
        rf"\s+\d{{1,2}}(?:\s*-\s*\d{{1,2}})?\s+(?:{_GALUH_MONTH_NAMES})\s+\d{{4}}.*$",
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
        r"^(-?[\d.]+)\s*\n\s*(PP[NnHh]\s+\d+%)\s*$",
        r"\2 \1",
        block,
        flags=re.MULTILINE,
    )


def _galuh_extract_items(text: str) -> List[Dict]:
    block_m = _GALUH_BLOCK_PATTERN.search(text)
    item_pattern = _GALUH_ITEM_LINE_PATTERN
    if not block_m:
        # 掃描檔 (使用者 2026-10 樣本 0500084/KWTGIB/2025)
        block_m = _GALUH_SCAN_BLOCK_PATTERN.search(text)
        item_pattern = _GALUH_SCAN_ITEM_LINE_PATTERN
    if not block_m:
        return []
    block = _galuh_rejoin_split_tax_lines(block_m.group(1))
    items: List[Dict] = []
    for line in block.splitlines():
        line = line.strip()
        if not line:
            continue
        if item_pattern is _GALUH_SCAN_ITEM_LINE_PATTERN:
            # 掃描版 OCR 會帶出表格線與序號 ("1 |TRUCKING ..."、"| PPh 2% ...")
            line = re.sub(r"^(?:\d+\s*)?\|+\s*", "", line)
            # 序號後面沒有表格線 ("1 TRUCKING PER TGL …"，2026-10-28 抽樣 0500085)
            line = re.sub(r"^\d{1,2}\s+(?=[A-Za-z])", "", line)
            # 行尾表格線 OCR 也會讀成 "}" / "]" ("… 14.130.000 }"，2026-10-30 0500003)
            line = re.sub(r"\s*[|}\]]+\s*$", "", line)
        tax_m = _GALUH_TAX_LINE_PATTERN.match(line)
        if tax_m:
            items.append({
                "description": tax_m.group("desc").strip(),
                "amount": None,
                "vat": _galuh_clean_amount(tax_m.group("amount")),
            })
            continue
        item_m = item_pattern.match(line)
        # 掃描版的範圍到 Terbilang 為止，會包含 "Total 184.730.049" 那行
        if item_m and re.fullmatch(r"total", item_m.group("desc").strip(), re.IGNORECASE):
            continue
        if item_m:
            items.append({
                "description": _galuh_strip_date_suffix(item_m.group("desc").strip()),
                "amount": _galuh_clean_amount(item_m.group("amount")),
                "vat": None,
            })
    return items


def _galuh_extract_grand_total(text: str) -> Optional[int]:
    total = _galuh_clean_amount(_search(r"^Total\s+([\d.,]+)\s*$", text))
    if total is None:
        # 掃描檔 OCR 常漏掉 Total 那行，用金額大寫 "Terbilang: ... rupiah" 換算
        words = _search(r"Terbilang\s*:\s*\|?\s*([A-Za-z ]+?)\s+rupiah", text)
        total = _indo_words_to_int(words) if words else None
    return total


def parse_galuh(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {}

    # 統一輸出正式名稱：掃描檔 OCR 常少空格、前面黏 logo 文字
    # ("PT. GALUH INTIBAHARI"、"“GALUH PT. GALUH INTI BAHARI")
    header["supplier"] = ("PT. GALUH INTI BAHARI"
                          if re.search(r"PT[.,]\s*GALUH\s*INTI\s*BAHARI", text) else None)
    # 掃描檔 OCR 常把標點讀錯 ("No, Invoice ; 0500084/KWTGIB/2025")
    header["invoice_no"] = _search(r"No[.,]\s*Invoice\s*[:;]\s*(\S+)", text)

    # 掃描檔 OCR 常把日、月黏在一起 ("Tgl. Invoice ; 03Jan 2025")
    _galuh_date_m = re.search(r"Tgl\.\s*Invoice\s*[:;]\s*(\d{1,2})\s*([A-Za-z]+)\s+(\d{4})", text)
    header["invoice_date"] = (
        _galuh_parse_date(*_galuh_date_m.groups()) if _galuh_date_m else None
    )
    if not header["invoice_date"]:
        # 掃描檔 OCR 有時把日期放到標籤上一行 ("… RW 000 - 03 Feb 2025" /
        # "… Tgl. Invoice :")：在 No. Invoice 到 Tgl. Invoice 之間找
        sec = re.search(r"No[.,]\s*Invoice(.*?)Tgl\.\s*Invoice", text, re.DOTALL)
        for d in re.finditer(r"\b(\d{1,2})\s*([A-Za-z]{3,})\s+(\d{4})\b", sec.group(1) if sec else ""):
            header["invoice_date"] = _galuh_parse_date(*d.groups())
            if header["invoice_date"]:
                break

    # consignee：優先找報表明細頁開頭「PT. 開頭、後面接 No : <發票號>」那
    # 一行 (比最上面的 "TO ..." 簡稱更正式/更準確)，找不到才退回 "TO ..."。
    _galuh_rekap_m = re.search(r"^(PT\..+?)\s+No\s*:\s*\S+\s*$", text, re.MULTILINE)
    if _galuh_rekap_m:
        header["consignee"] = _galuh_rekap_m.group(1).strip()
    else:
        header["consignee"] = (
            # 掃描檔 OCR 會讀成 "TO :" 自己一行、公司名稱在下一行 (2026-10-30 0500086)
            _search(r"\bTO\s*:?\s+(.+?)\s+No[.,]\s*Invoice", text)
            # 掃描檔 OCR 把 "No. Invoice" 排到 "TO …" 上一行，"TO …" 自己一行
            or _search(r"(?m)^\s*TO\s+(PT\b.+?)\s*$", text))

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
    if items and not any(re.match(r"PP[NnHh]", it.get("description") or "") for it in items):
        # 掃描檔 OCR 把名稱欄與金額欄分開讀："PPh 2%" 單獨一行、"-4.618.400" 單獨一行
        # (2026-10-28 抽樣 "Nike - Per 20 januari - 01 Februari 25 - 0500150")
        label = _search(r"^\s*(PP[Hh]\s+\d+%)\s*$", text)
        neg = _search(r"^\s*(-\d{1,3}(?:[.,]\d{3})+)\s*$", text)
        if label and neg:
            items.append({"description": label, "amount": None, "vat": _galuh_clean_amount(neg)})
    if not items:
        items = [{"description": None, "amount": None, "vat": None}]

    return {"header": header, "items": items}


register_supplier(
    "GALUH", "PT. GALUH INTI BAHARI", detect_galuh, parse_galuh,
    extract_grand_total=_galuh_extract_grand_total,
    # 掃描帳單的金額欄跟名稱分得很開，預設 OCR 模式會漏掉整欄金額
    ocr_psm=6,
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
    # 帳單整頁沒有發票日期：使用者規定用 ETD (開船日) 當發票日期
    header["invoice_date"] = header["onboard_date"]

    header["bl_no"] = _search(r"HB/L\s*No\s+(\S+)", flat)
    # 冒號後面多一個空白時，拿掉冒號會變成兩段 ("BSWCG: K260717" -> BSWCGK260717，
    # 使用者 2026-10-25 確認)
    m_split = re.search(r"HB/L\s*No\s+([A-Z0-9]+):\s+([A-Z0-9]*\d[A-Z0-9]*)\b", text)
    if m_split:
        header["bl_no"] = m_split.group(1) + m_split.group(2)
    header["mbl_no"] = _search(r"MB/L\s*No\s+(\S+)", flat)
    # Negara Asal 跟 Tipe 分別在不同的物理行 (右欄位逐行往下錯開一列)，
    # 不能用 flat (整頁合併成一行) 去抓到 Tipe 當結尾，國家名稱本身都是
    # 單一個字 (THAILAND/CIKANDE/ITALIA/VIETNAM/JAPAN)，直接抓下一個空白
    # 前的單字即可。
    # 標籤也可能被值切開："Negara KOREA" / "Asal :"、"Nett 3412" / "Weight :"
    # (使用者 2026-10 樣本 C25-03-0200)
    header["port_of_loading"] = (_search(r"Negara\s*Asal\s+(\S+)", flat)
                                 or _search(r"Negara\s+(\S+)\s+Asal\b", flat))
    header["port_of_discharge"] = "INDONESIA"

    m_weight = (re.search(r"Nett\s*Weight\s+([\d.,]+)", flat)
                or re.search(r"Nett\s+([\d.,]+)\s+Weight\b", flat))
    # 材積統一寫法 (見 _fmt_volume，使用者 2026-10-27)：Tipe 有櫃數寫櫃量
    # ("FCL // 1X40'HC COLLECT" -> "1X40HC"、"LCL - 2X40' + 1X20'" -> "2X40, 1X20")；
    # 沒有就寫 CBM ("CBM 25.21"，空運是 "0 / 0" 視為沒有)；再沒有寫 Nett Weight KG
    tipe = _search(r"Tipe\s+([^\n]*)", text.replace(":", "")) or ""
    boxes = re.findall(r"(\d+)X(\d{2}'?(?:HC)?)", tipe)
    header["volume"] = _fmt_volume(
        containers=boxes,
        cbm=_search(r"\bCBM\s+([\d.,]+)", flat),
        kg=m_weight.group(1) if m_weight else None)

    flight = _search(r"Flight\s+(.+?)\s+Gross", flat)
    voyage = _search(r"Voyage\s+(.+?)\s+CBM", flat)
    header["vessel"] = flight if flight else _ecomet_clean_dash(voyage)

    header["container_no"] = _ecomet_extract_container(t)

    items = _ecomet_extract_items(t)
    if not items:
        items = [{"description": None, "amount": None, "vat": None}]

    return {"header": header, "items": items}


def _ecomet_extract_grand_total(text: str) -> Optional[int]:
    """金額稽核：費用 + 代墊 (Reimbursement) + PPN = Grand Total (使用者 2026-10-25
    要求)。沒有代墊表時帳單只印 "Total Rp …"。"""
    t = text.replace(":", "")
    m = (re.search(r"Grand\s*Total\s+Rp\.?\s*([\d,.]+)", t, re.IGNORECASE)
         or re.search(r"^\s*Total\s+Rp\.?\s*([\d,.]+)", t, re.MULTILINE | re.IGNORECASE))
    return _ecomet_clean_amount(m.group(1)) if m else None


# --- ECOMET Excel 對帳彙總表 REKAP EXPORT (使用者 2026-10-28：要當帳單擷取) ----
# 每張工作表一個品牌 (ASICS/NIKE/ADIDAS)，抬頭 "INV : D26-03-0224"，每列一個貨櫃：
# TRUCKING 到 JUMLAH 之間的欄都是費用 (OT/ON、REPO、EARLY STACKING、LOLO / BANDARA…)，
# 下一列是各欄合計 (JUMLAH = 全部費用，用來勾稽)。沒有稅額與含稅合計。
# 表格下方留有舊範本 (2019 年資料、"TO : Bpk UUT")，只讀第一段。
# 同一發票號碼另有 "xxx REVISE" 工作表時，以 REVISE 為準 (2026-10-28 抽樣
# "REKAP EXPORT NIKOMAS PERIODE 23 FEBRUARI - 28 FEBRUARI 2026"：ASICS REVISE / ASICS)。
_ECOMET_SUPPLIER = "PT. TRANS ECOMET GLOBAL"


def _ecomet_rekap_sheet(name: str, rows: List[list]) -> Optional[Tuple[Dict, Dict]]:
    head = next((i for i, r in enumerate(rows[:20])
                 if {"TRUCKING", "JUMLAH", "CONTAINER"} <= {_vn_text(c).strip().upper() for c in r}), None)
    if head is None:
        return None
    labels = [_vn_text(c).strip().upper() for c in rows[head]]
    start, end = labels.index("TRUCKING"), labels.index("JUMLAH")
    col = {k: labels.index(k) for k in ("CONTAINER", "FEET", "JOB NO :") if k in labels}
    top = " ".join(_vn_text(c) for r in rows[:head] for c in r if c not in (None, ""))
    inv = re.search(r"INV\s*:\s*([A-Z]\d{2}-\d{2}-\d+)", top)
    if not inv:
        return None
    consignee = next((_vn_text(c).strip() for r in rows[1:head] for c in r
                      if _vn_text(c).strip().upper().startswith("PT.") and "ECOMET" not in _vn_text(c).upper()), None)

    def label(k):
        sub = _vn_text(_vn_cell(rows[head + 1], k)).strip() if head + 1 < len(rows) else ""
        return re.sub(r"\s+", " ", f"{_vn_text(rows[head][k]).strip()} {sub}").strip()

    items: List[Dict] = []
    total = None
    dates = []
    for r in rows[head + 1:]:
        if _vn_num(_vn_cell(r, 0)) is not None:          # 明細列 (NO 欄是序號)
            feet = _vn_text(_vn_cell(r, col.get("FEET"))).strip() if "FEET" in col else ""
            extra = {"container_no": _vn_text(_vn_cell(r, col["CONTAINER"])).strip() or None
                     if "CONTAINER" in col else None,
                     "volume": f"1X{feet}" if re.fullmatch(r"\d{2}'?", feet) else None,
                     "order_no": _vn_text(_vn_cell(r, col["JOB NO :"])).strip() or None
                     if "JOB NO :" in col else None}
            for k in range(start, end):
                amount = _vn_amount(_vn_cell(r, k))
                if amount:
                    items.append({"description": label(k), "amount": amount, "vat": None, **extra})
            d = _vn_cell(r, 2)
            if isinstance(d, (datetime.date, datetime.datetime)):
                dates.append(d)
        elif not _vn_text(_vn_cell(r, 0)).strip() and _vn_num(_vn_cell(r, start)) is not None:
            total = _vn_num(_vn_cell(r, end))                 # 各欄合計列：第一段到這裡結束
            break
    if not items:
        return None
    # 表上沒有發票日期：暫用最後一個裝櫃日 (TANGGAL STUFFING)，待使用者確認
    date = max(dates).strftime("%d %b %Y") if dates else None
    header = {"invoice_no": inv.group(1), "supplier": _ECOMET_SUPPLIER, "consignee": consignee,
              "invoice_date": date}
    return ({"header": header, "items": items},
            {"invoice_no": inv.group(1), "total": total, "grand_total": None, "final_total": None})


def _ecomet_sheets(book) -> Tuple[List[Dict], List[Dict]]:
    found: Dict[str, Tuple[bool, Tuple[Dict, Dict]]] = {}
    for name, rows in book.sheets:
        got = _ecomet_rekap_sheet(name, rows)
        if not got:
            continue
        inv, revise = got[1]["invoice_no"], "REVISE" in name.upper()
        if inv not in found or (revise and not found[inv][0]):
            found[inv] = (revise, got)
    docs = [g[0] for _, g in found.values()]
    return docs or [{"header": {}, "items": [{"description": None, "amount": None}]}], \
        [g[1] for _, g in found.values()]


register_supplier(
    "ECOMET", "PT. TRANS ECOMET GLOBAL", detect_ecomet, parse_ecomet,
    multi_page=True, extract_grand_total=_ecomet_extract_grand_total,
    sheet_parse=lambda book: _ecomet_sheets(book)[0], sheet_totals=lambda book: _ecomet_sheets(book)[1],
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
# 掃描檔 OCR 會把數量 "3" 讀成 "-" ("SEGEL 25,000 - 75,000")
_HANDAL_SEP = r'[\s.,;+\-]*'

_HANDAL_NON_ITEM_PREFIXES = (
    "DESCRIPTION", "QTY", "UNIT PRICE", "AMOUNT", "NO.", "NO ",
    "VESSEL", "DATE", "INVIOCE", "INVOICE", "NO.BC", "NO. BC",
    "SHIPPING REMARKS", "KEPADA YTH", "DEBIT NOTE", "E.M.K", "EMKU",
    "PT. HANDAL", "HANDAL MANDIRI", "FINANCE", "PAYMENT BY",
    "HARAP DITRANSFER", "BANK", "NO. REK", "KOMPLEK", "TELP",
    "FOR AND ON BEHALF", "TOTAL", "#", "RIBU RUPIAH", "JUTA",
    "JL RAYA", "JL.", "KAWASAN INDUSTRI", "SERANG", "TAMBAK",
    "KAB ", "RAYA SERANG",
    # 收件地址 (例如 GLOSTAR 的 "BOJONG CIKEMBAR SUKABUMI JAWA BARAT," 與
    # 單獨一行的 "INDONESIA")，不是費用項目
    "JL ", "BOJONG", "INDONESIA",
)

_HANDAL_CODE_RE = re.compile(r'\b([A-Z]{4,8}\d{6,10})\b')
_HANDAL_ORDER_NO_RE = re.compile(r'(?<![A-Z])([A-Z])(\d{5,7})(?!\d)')
# 發票號碼後面可能多一個字母，例如 "HG01240125S"
_HANDAL_INVOICE_NO_RE = re.compile(r'\bHG\d{8}[A-Z]?\b')
_HANDAL_INVOICE_DATE_RE = re.compile(r'\b(\d{1,2}\s+[A-Za-z]+\s+\d{4})\b')
_HANDAL_ROUTE_RE = re.compile(r'([A-Z][A-Z ]*?) *- *([A-Z][A-Z ]*?) *- *(\d{1,2}/\d{1,2}/\d{4})')
# 多種櫃型用逗號連起來，例如 "1X20,6X40"
_HANDAL_VOLUME_RE = re.compile(r'\b(\d{1,2}X\d{2,3}(?:,\s*\d{1,2}X\d{2,3})*)\b')
# 沒有欄位標籤的版型 (使用者 2026-10 樣本 HG02360225)：B/L 是獨立一行
# "250162C (000177)"，MBL 是 B/L 上面某一行的單一代碼 ("EGLV149411233840"、
# "002FX00666")，發票日期+號碼的下一行是「船名 航次 Order No.」
_HANDAL_BL_LINE_RE = re.compile(r'^\s*(\d{6}[A-Z]\s*\(\d{6}\))\s*$', re.MULTILINE)
_HANDAL_CONSIGNEE_RE = re.compile(r'\b([A-Z][A-Z\s]{2,40}\.\s*PT)\b')
# 總額要自己一行：掃描檔 OCR 常漏掉總額那行，舊規則會把最後一筆費用
# ("COLLECTION FEE 701,000 701,000") 的金額當成總額 (使用者 2026-10 樣本 NIKO-MEI 2025)
_HANDAL_TOTAL_RE = re.compile(r'^\s*(\d{1,3}(?:,\d{3})+)\s*\n+\s*#', re.MULTILINE)
# 金額大寫 "# Tiga Puluh Dua Juta ... Rupiah #"
_HANDAL_WORDS_RE = re.compile(r'^\s*#\s*([A-Za-z ]+?)\s+Rupiah\s*#', re.MULTILINE | re.IGNORECASE)
_INDO_DIGITS = {"satu": 1, "dua": 2, "tiga": 3, "empat": 4, "lima": 5, "enam": 6,
                "tujuh": 7, "delapan": 8, "sembilan": 9}
_INDO_SCALES = {"ribu": 1_000, "juta": 1_000_000, "milyar": 10 ** 9, "miliar": 10 ** 9}


def _indo_words_to_int(words: str) -> Optional[int]:
    """印尼文金額大寫轉數字，例如 'Tiga Puluh Dua Juta Dua Ribu' -> 32002000。
    有不認得的字 (OCR 雜訊) 就回傳 None，不猜。"""
    total = group = num = 0
    for w in words.lower().split():
        if w in _INDO_DIGITS:
            num += _INDO_DIGITS[w]
        elif w == "sepuluh":
            num += 10
        elif w == "sebelas":
            num += 11
        elif w == "seratus":
            group += 100
        elif w == "seribu":
            total += 1000
        elif w == "belas":
            num += 10
        elif w == "puluh":
            group += num * 10
            num = 0
        elif w == "ratus":
            group += num * 100
            num = 0
        elif w in _INDO_SCALES:
            total += (group + num) * _INDO_SCALES[w]
            group = num = 0
        else:
            return None
    value = total + group + num
    return value or None
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
    # 船名行常有 OCR 底線 ("SINAR SANUR 136S_")；"INVIOCE" 標籤會被讀成 "NVI"、"VMIOCE"
    if '_' in text or difflib.SequenceMatcher(None, text, "INVOICE").ratio() >= 0.6:
        return False
    upper = text.upper()
    for prefix in _HANDAL_NON_ITEM_PREFIXES:
        if upper.startswith(prefix):
            return False
    if re.search(r'\.\s*PT\b', upper):
        return False
    # 掃描檔的雜訊行：船名那行的 "MESSRS" 標籤 (OCR 讀成 "MSS)"、"MESS 55]")、
    # 印花稅票 "SEPULUH RIBU RUPIAH" / "METERAI"
    if re.match(r'^(?:MESSRS|M[ES]{1,6}[\W\d])', upper) or re.search(r'RUPIA|SEPULUH|^METE', upper):
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
    # 名稱後面可能有 OCR 讀出的標點雜訊 ("SEGEL , ; 25,000 25,000")
    rf'^(?:\d+\s+)?(?P<desc>[A-Z][A-Za-z0-9 /().=*\-]*?)[\s,;:]+(?:\d+[\s,;:]+)?'
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
    # 散貨 (LCL) 沒有櫃型，材積欄寫重量 (例如 "614KG")
    m_kg = None if m else re.search(r'\b(\d[\d,.]*\s?KG)\b', text)
    # 材積統一寫法 (見 _fmt_volume)："1X20,6X40" -> "1X20, 6X40"、"614KG" -> "614 KG"
    h['volume'] = _volume_from_text(m.group(1) if m else (m_kg.group(1) if m_kg else None))

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

    bl_m = _HANDAL_BL_LINE_RE.search(text)
    if bl_m:
        # 沒有欄位標籤的版型：依位置抓
        h['bl_no'] = re.sub(r'\s+', ' ', bl_m.group(1))
        before = [ln.strip() for ln in text[:bl_m.start()].splitlines() if ln.strip()]
        codes = [ln for ln in before if re.fullmatch(r'[A-Z0-9]{8,20}', ln) and re.search(r'\d', ln)]
        h['mbl_no'] = codes[-1] if codes else None
        lines = text.splitlines()
        for i, ln in enumerate(lines):
            if _HANDAL_INVOICE_NO_RE.search(ln) and i + 1 < len(lines):
                vm = re.match(r'^\s*(.+?)\s+(\d{6})\s*$', lines[i + 1])
                if vm:
                    h['vessel'], h['order_no'] = vm.group(1).strip(), vm.group(2)
                break

    h['supplier'] = "PT. Handal Mandiri Transindo"
    h['arrive_date'] = None
    h['container_no'] = None
    h['shipper'] = None

    return h


def _handal_words_total(text: str) -> Optional[int]:
    m = _HANDAL_WORDS_RE.search(text)
    return _indo_words_to_int(m.group(1)) if m else None


def _handal_item_region(text: str) -> str:
    """只取『總金額數字出現以前』的文字範圍來找費用項目，把發票下方的簽名、
    銀行帳號、電話、地址等雜訊文字整段排除，避免被誤判成費用項目。
    總額數字跟金額大寫不一致時，那個數字其實是最後一筆費用 (欄位分開的
    掃描版)，改在金額大寫那行切。"""
    m = _HANDAL_TOTAL_RE.search(text)
    words = _handal_words_total(text)
    if m and (words is None or _handal_amount_to_int(m.group(1)) == words):
        return text[:m.start()]
    m = _HANDAL_WORDS_RE.search(text)
    return text[:m.start()] if m else text


def _handal_pair_blocks(nums: List) -> Optional[List]:
    """欄位分開的掃描版：單價欄、金額欄一段一段交錯 (4 個單價、4 個金額、
    3 個單價、3 個金額…)，每段單價跟金額一樣。找出一對對相同的段落，取金額段。"""
    out, i = [], 0
    while i < len(nums):
        for k in range(1, (len(nums) - i) // 2 + 1):
            if nums[i:i + k] == nums[i + k:i + 2 * k]:
                out += nums[i + k:i + 2 * k]
                i += 2 * k
                break
        else:
            return None
    return out


def _handal_pick_amounts(nums: List, n: int, target: Optional[int]) -> List:
    """名稱跟數字分開時，從數字清單挑出 n 筆金額；有總額時挑加總相符的排法。"""
    cands = []
    if len(nums) == n:
        cands.append(nums)
    elif len(nums) == 2 * n:
        cands += [nums[1::2], nums[n:]]
        blocks = _handal_pair_blocks(nums)
        if blocks and len(blocks) == n:
            cands.append(blocks)
    for c in cands:
        if target is not None and None not in c and sum(c) == target:
            return c
    return cands[0] if cands else (nums + [None] * n)[:n]


def _handal_extract_items(text: str):
    region = _handal_item_region(text)
    # 有 "DESCRIPTION" 表頭時從表頭開始 (上面的船名行不是費用)；
    # "For and on behalf" (簽名欄) 之後只收數字 (欄位分開的掃描版金額在後面)
    head = re.search(r'^\s*DESCRIPTION\b.*$', region, re.MULTILINE)
    if head:
        region = region[head.end():]
    lines = [l.strip() for l in region.splitlines()]
    items = []
    pending_descs = []
    pending_numbers = []
    in_sign = False

    for line in lines:
        # 掃描檔 OCR 會帶出表格線 ("SEGEL | 25,000 25,000") 與行尾的撇號
        # ("TRUCKING LAUT 40 (PPH 23) 3,420,000 '"，抽樣 PCI-FEB 2025)
        line = re.sub(r'\s*\|\s*', ' ', line).strip()
        line = re.sub(r"\s+[^\w\s()]{1,2}$", "", line)             # 行尾雜訊 "'"
        line = re.sub(r"(?<=\s)[^\w\s(),.\-/]{1,2}(?=\s)", "", line)  # 中間雜訊 "°"
        line = re.sub(r"\s{2,}", " ", line)
        if not line:
            continue
        if re.match(r'for and on', line, re.IGNORECASE):
            in_sign = True
        if in_sign and not re.match(r'^[\d,.\s]+$', line):
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

        # 名稱只接單價、金額欄在別處 ("THC 3,168,000")：只留名稱，金額從數字欄配
        mu = re.match(rf'^(?P<desc>.+?)\s+{_HANDAL_NUM}$', line)
        if mu and _looks_like_handal_item_desc(mu.group('desc')):
            pending_descs.append(mu.group('desc').strip())
            continue

        if _looks_like_handal_item_desc(line):
            pending_descs.append(line)
            continue

        m3 = _HANDAL_SINGLE_NUM_RE.match(line)
        if m3:
            pending_numbers.append(_handal_amount_to_int(m3.group('amount')))
            continue

    if pending_descs:
        total = _handal_extract_total(text)
        done = sum(it["amount"] or 0 for it in items)
        if pending_numbers and total is not None and pending_numbers[-1] == total:
            pending_numbers = pending_numbers[:-1]   # 總額那行 (金額欄最後)
        amounts = _handal_pick_amounts(pending_numbers, len(pending_descs),
                                       None if total is None else total - done)
        for desc, amount in zip(pending_descs, amounts):
            items.append({"description": desc, "amount": amount, "vat": None})

    return items


def detect_handal(text: str) -> bool:
    """有些 HANDAL 帳單沒有供應商名稱，用發票號碼格式 "HG" + 8 位數字辨認
    (使用者提示)；再要求有印尼文金額大寫結尾 "… Rupiah #"，避免其他文件剛好
    出現類似號碼。舊規則要 "CUSTOM CLEARANCE"，帳單寫 "CUSTOMS CLEARANCE"
    時會漏認。"""
    upper = text.upper()
    if "HANDAL MANDIRI TRANSINDO" in upper or "HANDAL" in upper:
        return True
    # Excel 對帳彙總表 (2026-10-28 抽樣)：沒有公司名稱，用 "INVOICE NO : HET…" 加表頭 NO POL
    if _HANDAL_XLS_INV_RE.search(upper) and "NO POL" in upper:
        return True
    return bool(_HANDAL_INVOICE_NO_RE.search(text) and re.search(r"RUPIAH\s*#", upper))


# --- HANDAL Excel 對帳彙總表 (使用者 2026-10-28：要當帳單擷取) -----------------
# 每張工作表一張發票 (HT = 拖車 HET…、HL = 吊櫃 HEL…)，抬頭 "INVOICE NO : HET21691225"
# 或 "NO INVOICE = HET20550425"；每列一個貨櫃，FEET 右邊到 JUMLAH 之前都是費用
# (TRUCKING、CUSTOM、LIFT ON / LIFT OFF、BONGKARAN BANDARA…)，下一列是各欄合計。
# 沒有稅額與發票日期。沒有發票號碼的工作表 ("ON ADIDAS" 加收紀錄) 不擷取。
_HANDAL_XLS_INV_RE = re.compile(r"(?:INVOICE\s*NO|NO\s*INVOICE)\s*[:=]\s*(HE[A-Z]\d{8})")


def _handal_rekap_sheet(rows: List[list]) -> Optional[Tuple[Dict, Dict]]:
    head = next((i for i, r in enumerate(rows[:15])
                 if {"NO POL", "FEET", "JUMLAH"} <= {_vn_text(c).strip().upper() for c in r}), None)
    if head is None:
        return None
    top = " ".join(_vn_text(c) for r in rows[:head] for c in r if c not in (None, "")).upper()
    inv = _HANDAL_XLS_INV_RE.search(top)
    if not inv:
        return None
    labels = [_vn_text(c).strip().upper() for c in rows[head]]
    feet_col, end = labels.index("FEET"), labels.index("JUMLAH")
    cont_col = next((k for k, v in enumerate(labels) if v.startswith("NO CONT")), None)
    consignee = next((_vn_text(c).strip() for r in rows[:head] for c in r
                      if re.match(r"PT\.?\s", _vn_text(c).strip().upper())), None)
    if not consignee:
        # 抬頭沒寫 PT. ("NIKOMAS GEMILANG"，2026-10-28 抽樣 "NIKE 205 ( 22 JUNI - 27 JUNI ) NN648")：
        # 用第一列的公司名稱 (不是 "…FACTORY" 引號列、TRUCKING / INVOICE 列)
        consignee = next((t for r in rows[:head] for t in (_vn_text(c).strip() for c in r)
                          if re.fullmatch(r"[A-Z][A-Z .&,]{5,}", t.upper()) and not t.startswith('"')
                          and not re.match(r"(TRUCKING|INVOICE|NO\s)", t.upper())), None)

    def label(k):
        sub = _vn_text(_vn_cell(rows[head + 1], k)).strip() if head + 1 < len(rows) else ""
        return re.sub(r"\s+", " ", f"{_vn_text(rows[head][k]).strip()} {sub}").strip()

    items, total, dates = [], None, []
    for r in rows[head + 1:]:
        if _vn_num(_vn_cell(r, 0)) is not None:          # 明細列 (NO 欄是序號)
            feet = _vn_text(_vn_cell(r, feet_col)).strip()
            # 貨櫃欄寫車型 ("WINGBOX"、"BOX"，沒有數字) 時不是貨櫃號碼
            box = _vn_text(_vn_cell(r, cont_col)).strip() if cont_col is not None else ""
            extra = {"container_no": box if re.search(r"\d", box) else None,
                     "volume": f"1X{feet}" if re.fullmatch(r"\d{2}'?", feet) else None}
            for k in range(feet_col + 1, end):
                amount = _vn_amount(_vn_cell(r, k))
                if amount:
                    items.append({"description": label(k), "amount": amount, "vat": None, **extra})
            d = _vn_cell(r, 1)
            if isinstance(d, (datetime.date, datetime.datetime)):
                dates.append(d)
        elif _vn_num(_vn_cell(r, end)) is not None:
            total = _vn_num(_vn_cell(r, end))                 # 各欄合計列
            break
    if not items:
        return None
    # 表上沒有發票日期：暫用最後一個裝貨日 (TGL MUAT)，待使用者確認
    header = {"invoice_no": inv.group(1), "supplier": "PT. HANDAL MANDIRI TRANSINDO", "consignee": consignee,
              "invoice_date": max(dates).strftime("%d %b %Y") if dates else None}
    return ({"header": header, "items": items},
            {"invoice_no": inv.group(1), "total": total, "grand_total": None, "final_total": None})


def _handal_sheets(book) -> Tuple[List[Dict], List[Dict]]:
    got = [g for g in (_handal_rekap_sheet(rows) for _, rows in book.sheets) if g]
    return ([g[0] for g in got] or [{"header": {}, "items": [{"description": None, "amount": None}]}],
            [g[1] for g in got])


def parse_handal(text: str) -> Dict:
    header = _handal_extract_header(text)
    items = _handal_extract_items(text)
    if not items:
        items = [{"description": None, "amount": None, "vat": None}]
    return {"header": header, "items": items}


def _handal_extract_total(text: str):
    # 金額大寫比數字可靠：OCR 漏掉總額那行時，金額前一行是最後一筆費用
    words = _handal_words_total(text)
    if words is not None:
        return words
    m = _HANDAL_TOTAL_RE.search(text)
    return _handal_amount_to_int(m.group(1)) if m else None


register_supplier(
    "HANDAL", "PT. Handal Mandiri Transindo", detect_handal, parse_handal,
    multi_page=True, force_ocr=True,
    extract_total=_handal_extract_total,
    sheet_parse=lambda book: _handal_sheets(book)[0], sheet_totals=lambda book: _handal_sheets(book)[1],
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
    if not header["invoice_date"]:
        # 抬頭的 Date 讀不到時，改抓右下角簽署日期 "JAKARTA, 11 AGUSTUS 2026"
        # (使用者提示)
        sign_m = re.search(r"JAKARTA,\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", text, re.IGNORECASE)
        month = INDO_MONTHS.get(sign_m.group(2).lower()) if sign_m else None
        if month:
            header["invoice_date"] = _pss_format_date(sign_m.group(1), month, sign_m.group(3))

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

    # 其他費用合計行，例如 "TOTAL BIAYA LAINNYA ` 1,785,000 1,785,000"、
    # "TOTAL BIAYA KETERLAMBATAN Rp. 5,000,000" (名稱後面可能夾 "`"、"Rp.")；
    # 金額 > 0 才列 (使用者 2026-10 樣本)
    for m in re.finditer(r"(?m)^(TOTAL BIAYA (?!TRUCKING|REIMBURSEMENT)[A-Z ]+?)\s+[`'\s]*"
                         r"(?:Rp\.?\s*)?([\d,]+)", text):
        amount = _pss_clean_amount(m.group(2))
        if amount:
            items.append({"description": m.group(1).strip(), "amount": amount, "vat": None})

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
    """含稅總額，稽核第 1 段 (DPP + PPN = GRAND TOTAL) 用。延遲費這類沒有
    GRAND TOTAL 的帳單，改用 "TOTAL Rp. 5,000,000"。"""
    m = (re.search(r"GRAND TOTAL\s+([\d,]+)", text)
         or re.search(r"(?<!BIAYA )\bTOTAL\s+Rp\.?\s*([\d,]+)", text))
    return _pss_clean_amount(m.group(1)) if m else None


def _pss_extract_total_paid(text: str) -> Optional[int]:
    """實付金額 (TOTAL YANG DIBAYAR)，稽核第 2 段：
    GRAND TOTAL - PPH = TOTAL YANG DIBAYAR 用。"""
    m = re.search(r"TOTAL YANG DIBAYAR\s+([\d,]+)", text)
    return _pss_clean_amount(m.group(1)) if m else None


# --- PUTRA Excel 帳單 (使用者 2026-10-28：REKAP TAGIHAN 要當帳單擷取) ----------
# 兩種版面，每張顯示中的工作表是一張發票：
#   1. REKAP TAGIHAN (REKAPITULASI BIAYA)：每列一趟車，費用欄用代號列 A.~G. 標示
#      (A~F 課 PPN、G = LOLO NON-DPP)，H. = JUMLAH。表尾 DPP / PPN / Non DPP /
#      GRAND TOTAL / PPH 23 / TOTAL YANG DITAGIHKAN。發票號碼在抬頭 "NO. INV : PSS-…"
#      (ADIDAS 2025-01 抬頭有兩個 NO. INV，上一張的沒刪，取最後一個)。
#   2. INVOICE (跟 PDF 發票同格式)：工作表轉成文字，沿用 PDF 的擷取與三段稽核。
# 同一張發票 PDF 也有時，資料庫依「供應商 + 發票號碼」只留一份。
_PSS_SUPPLIER = "PT. PUTRA SEJAHTERA SENTOSA"


def _pss_cell_text(value) -> str:
    """儲存格 -> PDF 文字寫法：整數金額加千分位 (153490000.0 -> 153,490,000)。"""
    if isinstance(value, float) and abs(value - round(value)) < 0.01:
        value = int(round(value))
    if isinstance(value, int) and not isinstance(value, bool):
        return f"{value:,}"
    if isinstance(value, (datetime.date, datetime.datetime)):
        return value.strftime("%d/%m/%Y")
    return re.sub(r"\s+", " ", str(value)).strip()


def _pss_sheet_as_text(rows: List[list]) -> str:
    return "\n".join("  ".join(t for t in (_pss_cell_text(c) for c in r if c not in (None, "")) if t)
                     for r in rows)


def _pss_rekap_sheet(rows: List[list]) -> Optional[Tuple[Dict, Dict]]:
    """REKAP TAGIHAN 一張工作表 -> (發票, 合計)；不是這個版面回傳 None。"""
    head = next((i for i, r in enumerate(rows[:15])
                 if {"TANGGAL", "TRUCKING"} <= {_vn_text(c).strip().upper() for c in r}), None)
    if head is None:
        return None
    letters = next((i for i in range(head + 1, min(head + 4, len(rows)))
                    if sum(bool(re.fullmatch(r"[A-H]\.?", _vn_text(c).strip())) for c in rows[i]) >= 3), None)
    if letters is None:
        return None
    code = {k: _vn_text(c).strip().rstrip(".") for k, c in enumerate(rows[letters])
            if re.fullmatch(r"[A-H]\.?", _vn_text(c).strip())}
    fee_cols = [k for k, v in code.items() if v != "H"]
    sum_col = next((k for k, v in code.items() if v == "H"), None)

    def label(k):
        parts = [_vn_text(_vn_cell(rows[i], k)).strip() for i in range(head, letters)]
        return re.sub(r"\s+", " ", " ".join(p for p in parts if p))

    labels = [_vn_text(c).strip().upper() for c in rows[head]]
    container_col = labels.index("CONTAINER") if "CONTAINER" in labels else None
    feet_col = labels.index("FEET") if "FEET" in labels else None
    top = " ".join(_vn_text(c) for r in rows[:head] for c in r if c not in (None, ""))
    inv = re.findall(r"NO\.\s*INV\s*:\s*(\S+)", top)
    consignee = next((re.sub(r"\s*-\s*FCL.*$", "", _vn_text(c).strip()) for r in rows[:head] for c in r
                      if _vn_text(c).strip().upper().startswith("PT.") and "PUTRA" not in _vn_text(c).upper()), None)
    header = {"invoice_no": inv[-1] if inv else None, "supplier": _PSS_SUPPLIER, "consignee": consignee,
              "invoice_date": None}
    items: List[Dict] = []
    totals = {"invoice_no": header["invoice_no"], "total": None, "grand_total": None, "final_total": None}
    for r in rows[letters + 1:]:
        first = _vn_text(_vn_cell(r, 0)).strip()
        nums = [x for x in (_vn_num(c) for c in r) if x is not None]
        if _vn_num(_vn_cell(r, 0)) is not None:          # 明細列 (NO 欄是序號)
            box = _vn_text(_vn_cell(r, container_col)).strip() if container_col is not None else ""
            feet = _vn_text(_vn_cell(r, feet_col)).strip() if feet_col is not None else ""
            extra = {"container_no": box if box and box.upper() != "BOX" else None,
                     "volume": f"1X{feet}" if re.fullmatch(r"\d{2}'?", feet) else None}
            for k in fee_cols:
                amount = _vn_amount(_vn_cell(r, k))
                if amount:
                    items.append({"description": label(k), "amount": amount, "vat": None, **extra})
        elif not first and sum_col is not None and totals["total"] is None and nums:
            totals["total"] = _vn_num(_vn_cell(r, sum_col))   # 各欄合計列：JUMLAH = 全部費用
        elif nums:
            text, value = first.upper(), nums[-1]
            m = re.search(r"PPN\s*(\d[\d,]*)\s*%", text)
            if m:
                items.append({"description": f"PPN {m.group(1)} %", "amount": None, "vat": round(value, 2)})
            m = re.search(r"PPH\s*(?:\d+\s*,\s*)?(\d[\d,]*)\s*%", text)
            if m:
                items.append({"description": f"PPH {m.group(1)}%", "amount": None, "vat": round(value, 2)})
            if "GRAND TOTAL" in text:
                totals["grand_total"] = value
            elif "TOTAL YANG DITAGIH" in text or "TOTAL YANG DIBAYAR" in text:
                totals["final_total"] = value
        sign = re.search(r"JAKARTA\s*,\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})",
                         " ".join(_vn_text(c) for c in r if c not in (None, "")), re.I)
        if sign and INDO_MONTHS.get(sign.group(2).lower()):
            header["invoice_date"] = _pss_format_date(sign.group(1), INDO_MONTHS[sign.group(2).lower()],
                                                      sign.group(3))
    return {"header": header, "items": items or [{"description": None, "amount": None, "vat": None}]}, totals


def _pss_sheets(book) -> Tuple[List[Dict], List[Dict]]:
    docs, totals = [], []
    for name, rows in book.sheets:
        rekap = _pss_rekap_sheet(rows)
        if rekap:
            docs.append(rekap[0])
            totals.append(rekap[1])
            continue
        text = _pss_sheet_as_text(rows)
        if not re.search(r"Invoice\s*No", text, re.I) or "TOTAL BIAYA" not in text.upper():
            continue
        doc = parse_putra_sejahtera_sentosa(text)
        doc["header"]["supplier"] = doc["header"].get("supplier") or _PSS_SUPPLIER
        docs.append(doc)
        totals.append({"invoice_no": doc["header"].get("invoice_no"), "total": _pss_extract_dpp(text),
                       "grand_total": _pss_extract_grand_total(text),
                       "final_total": _pss_extract_total_paid(text)})
    return docs or [{"header": {}, "items": [{"description": None, "amount": None}]}], totals


register_supplier(
    "PUTRA_SEJAHTERA_SENTOSA", "PT. Putra Sejahtera Sentosa",
    detect_putra_sejahtera_sentosa, parse_putra_sejahtera_sentosa,
    multi_page=True,
    extract_total=_pss_extract_dpp,
    extract_grand_total=_pss_extract_grand_total,
    extract_final_total=_pss_extract_total_paid,
    deduction_prefixes=("PPH",),
    sheet_parse=lambda book: _pss_sheets(book)[0], sheet_totals=lambda book: _pss_sheets(book)[1],
)


# ===========================================================================
# 越南 Excel 帳單共用工具 (VN03 RỒNG Á CHÂU、VN05 CONG THANH)
# ===========================================================================
# 越南供應商用 Excel 對帳表 (Bảng kê / Statement) 請款，不是 PDF。核心引擎
# 只交出「顯示中」的工作表 (read_spreadsheet())，這裡直接讀儲存格，不從文字
# 切欄位。金額欄位是數字型態，不用處理千分位。

def _vn_text(value) -> str:
    """儲存格 -> 文字：整數型的浮點數 (例如報關單號 308933256500.0) 去掉 .0，
    連續空白壓成一個。"""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return re.sub(r"\s+", " ", str(value)).strip()


def _vn_label(value) -> str:
    """表頭文字比對用：轉大寫、換行與連續空白壓成一個空白。"""
    return _vn_text(value).upper()


def _vn_num(value) -> Optional[float]:
    """數字儲存格 (或純數字文字) -> float，空白/非數字回傳 None。"""
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = _vn_text(value)
    # 逗號當小數點的文字："237,60" (2026-10-28 抽樣 VN14 TOTAL USD 欄)。只有一個逗號、
    # 後面剛好 2 位數才算小數；"5,000" 這類千分位 (3 位數) 不受影響
    if re.fullmatch(r"-?\d+,\d{2}", text.strip()):
        text = text.strip().replace(",", ".")
    # 越南寫法點當千分位的文字："116.470" = 116,470 (2026-10-28 抽樣 VN16 "BANG KE POU LI
    # -T4.2026" PHÍ CÂN 欄)。每段剛好 3 位數才算
    elif re.fullmatch(r"-?\d{1,3}(?:\.\d{3})+", text.strip()):
        text = text.strip().replace(".", "")
    text = text.replace(",", "")
    try:
        return float(text) if text else None
    except ValueError:
        return None


def _vn_amount(value) -> Optional[float]:
    """金額：分攤出來的金額會有很多位小數，四捨五入到小數 2 位；0 視為沒有。"""
    num = _vn_num(value)
    if not num:
        return None
    return round(num, 2)


def _vn_date(value) -> Optional[str]:
    """日期儲存格 (datetime)、'DD/MM/YYYY' 或 'YYYY-MM-DD 00:00:00.0' 文字
    -> 'DD Mon YYYY'。"""
    if isinstance(value, (datetime.datetime, datetime.date)):
        return _format_date_en(value.day, value.month, value.year)
    text = _vn_text(value)
    m = re.match(r"^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$", text)
    if m:
        return _format_date_en(m.group(1), m.group(2), m.group(3))
    # 文字格式的日期時間，例如 '2026-08-03 00:00:00.0'、'20251202'
    # 也接受 '2025/02/17' (年/月/日，VN05 2025 年對帳表)
    m = (re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T][\d:.]*)?$", text)
         or re.match(r"^(20\d{2})(\d{2})(\d{2})(?:\.0)?$", text))
    return _format_date_en(m.group(3), m.group(2), m.group(1)) if m else None


def _vn_cell(row: list, col: Optional[int]):
    return row[col] if col is not None and col < len(row) else None


_VN_CONTAINER_RE = re.compile(r"^[A-Z]{4}\d{7}$")


def _vn_container(value) -> Optional[str]:
    """「CONT / XE」欄常混著車牌 (51C.91515) 或簡寫櫃號 (CTTU642)，只收完整的
    貨櫃號碼 (4 碼英文 + 7 碼數字)。"""
    text = _vn_text(value).upper().replace(" ", "")
    return text if _VN_CONTAINER_RE.match(text) else None


# ===========================================================================
# 供應商 VN05：CONG THANH (CTY TNHHDV GN VT VA TM CONG THANH，廠商代號 100008048)
# ===========================================================================
# 每個月分好幾批 (DOT 1/DOT 2…) 請款，每批兩個檔案：
#   1. 「POUSUNG EX T09.2026 DOT 1.xls」：兩個工作表各是一張對帳單
#        VC T9-1.2026 = STATEMENT OF EXPORT TRANPORTATION CHARGE (運費)
#        DV T9-1.2026 = STATEMENT OF EXPORT SERVICE CHARGE (服務費)
#      抬頭：賣方/買方名稱、"Form : 26/08 - 15/09/2026 .No: 09VC-1/2026"；
#      明細每列一筆出貨 (SHIP DATE / DEC 報關單號 / SO INV / 櫃型件數 / CBM /
#      Total amount / Remark 送貨地點 / CONT XE)，併車的出貨單價 0 (費用算在
#      同車第一筆)，0 元列不轉；最後 Total、VAT: (8%)、Grand total；簽核欄
#      "CONG THANH 15/09/2026" 是對帳單日期。
#   2. 「…-CHI HO.xls(x)」：代墊款 INVOICE LIST (第三方發票，例如 "Nang ha"
#      吊櫃上下)，每筆代墊自帶 VAT 欄 (每品項自帶 VAT)。同一張第三方發票分攤
#      到好幾筆出貨時，後面的出貨只列 INV / 報關單號 (Tờ khai)、沒有金額。
#      欄位對應 (使用者答案 test_samples/CONG_THANH/sample1)：每張第三方發票
#      各一張發票，INVOICE NO = Số Hóa đơn、INVOICE DATE = Ngày、Order No. =
#      這張發票所有的 INV、B/L NO. 與 MBL NO. = 所有報關單號 (逗號連接)；
#      合計列 "TỔNG CỘNG"。越南文費用名稱的英文備註在寫入轉檔結果時才加
#      (extract_utils.english_note())。
#      勾稽 (使用者提示)：每張發票 AMOUNT + VAT = 該列 TOTAL；TỔNG CỘNG 是三欄
#      各自的合計。帳單 TỔNG CỘNG 的 VAT 常是「未稅合計 × 8%」，跟逐列四捨五入
#      的 VAT 加總會有幾元尾差，容許每列 0.5。
# 辨識：賣方稅籍編號 0301434547。

_CONG_THANH_TAX_ID = "0301434547"


def detect_cong_thanh(text: str) -> bool:
    upper = text.upper()
    return _CONG_THANH_TAX_ID in text and ("CÔNG THÀNH" in upper or "CONG THANH" in upper)


def _ct_header(rows: List[list]) -> Dict[str, Optional[str]]:
    """對帳單抬頭：賣方、買方、對帳單號、對帳單日期 (簽核欄日期，沒有就用期間迄日)。"""
    text = "\n".join(" | ".join(_vn_text(v) for v in r if _vn_text(v)) for r in rows)
    header: Dict[str, Optional[str]] = {code: None for code in (
        "invoice_date", "invoice_no", "supplier", "consignee")}
    header["supplier"] = _search(r"Tên Đơn vị bán hàng:\s*([^|\n]+)", text)
    header["consignee"] = _search(r"Tên Đơn vị mua hàng:\s*\|?\s*([^|\n]+)", text)
    header["invoice_no"] = _search(r"\bNo:\s*([\w./-]+)", text)
    signed = re.search(r"CONG THANH\s+(\d{1,2})/(\d{1,2})/(\d{4})", text)
    period = re.search(r"Form\s*:\s*[\d/]+\s*-\s*(\d{1,2})/(\d{1,2})/(\d{4})", text)
    date_m = signed or period
    if date_m:
        header["invoice_date"] = _format_date_en(*date_m.groups())
    return header


def _ct_columns(rows: List[list], start: int) -> Tuple[Optional[int], Dict[str, int]]:
    """找表頭列 (含 'SO INV' 或 'Số Hóa đơn')，回傳 (表頭列 index, {表頭文字: 欄號})。"""
    for i in range(start, min(start + 20, len(rows))):
        labels = {_vn_label(v): j for j, v in enumerate(rows[i]) if _vn_label(v)}
        if "SO INV" in labels or "SỐ HÓA ĐƠN" in labels:
            return i, labels
    return None, {}


def _ct_statement(rows: List[list]) -> Optional[Dict]:
    """運費 (VC) / 服務費 (DV) 對帳單。"""
    header = _ct_header(rows)  # 簽核欄 (對帳單日期) 在表格最下面
    title = next((_vn_text(r[0]) for r in rows[:15]
                  if r and _vn_label(r[0]).startswith("STATEMENT OF")), "")
    # "STATEMENT OF EXPORT  TRANPORTATION CHARGE  SEPTEMBER OF , 2026" -> 費用名稱
    charge = re.sub(r"^STATEMENT OF\s+|\s+(JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|"
                    r"SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\b.*$", "", title, flags=re.I).strip()
    head_i, cols = _ct_columns(rows, 0)
    if head_i is None or not charge:
        return None
    c_date, c_inv, c_cbm = cols.get("SHIP DATE"), cols.get("SO INV"), cols.get("CBM")
    c_amount, c_remark, c_dec = cols.get("TOTAL AMOUNT"), cols.get("REMARK"), cols.get("DEC")
    c_cont, c_plate = cols.get("CONT XE"), cols.get("SỐ XE")

    items = []
    total = vat = grand = None
    sum_col = None
    prev = {"date": None, "remark": None, "cont": None, "plate": None}
    for row in rows[head_i + 1:]:
        labels = [_vn_label(v) for v in row if isinstance(v, str) and _vn_label(v)]
        # 合計區 (TOTAL / Total / VAT: (8%) / Grand total)：金額不一定在明細的
        # 金額欄 (服務費表是在前一欄)，取該列最右邊的數字。同一列有兩欄數字時
        # (2026-10-29 "POUSUNG EX T10.2025 DOT1.xls" DV 表：L 欄 = 明細加總、O 欄多
        # 1,250,000)，使用者確認 L 欄才是請款金額：Total 取等於明細加總的那一欄，
        # VAT / Grand total 取同一欄
        cells = [(j, _vn_num(v)) for j, v in enumerate(row) if not isinstance(v, str)]
        cells = [(j, n) for j, n in cells if n is not None]
        summary = cells[-1][1] if cells else None
        if labels and ("TOTAL" in labels or any(lb.startswith("VAT") for lb in labels)) and len(cells) > 1:
            detail = sum(it["amount"] or 0 for it in items)
            if "TOTAL" in labels and not any(lb.startswith("VAT") for lb in labels):
                sum_col = next((j for j, n in cells if abs(n - detail) <= 1), sum_col)
            summary = next((n for j, n in cells if j == sum_col), summary)
        elif sum_col is not None and cells and "GRAND TOTAL" in labels:
            summary = next((n for j, n in cells if j == sum_col), summary)
        if "GRAND TOTAL" in labels:
            grand = summary
            break
        if any(lb.startswith("VAT") for lb in labels):
            vat = summary
            continue
        if "TOTAL" in labels:
            total = summary
            continue
        so_inv = _vn_text(_vn_cell(row, c_inv))
        dec = _vn_text(_vn_cell(row, c_dec))
        amount = _vn_num(_vn_cell(row, c_amount))
        # DEC / SO INV 寫的是文字 (沒有數字) = 額外費用項目，例如
        # "Phát sinh Giữ gác"：費用名稱用這段文字 (寫入時加英文翻譯)，
        # 單號欄照抄原文，驗證層標示人工查核 (使用者規定)
        extra = next((t for t in (dec, so_inv) if t and not re.search(r"\d", t)), None)
        if extra and amount:
            items.append({
                "description": extra, "amount": amount, "vat": None,
                "invoice_no": so_inv if re.search(r"\d", so_inv) else header["invoice_no"],
                "invoice_date": _vn_date(_vn_cell(row, c_date)) or header["invoice_date"],
                "order_no": dec or None, "bl_no": so_inv or None, "mbl_no": so_inv or None,
                "port_of_loading": "VIETNAM", "port_of_discharge": "VIETNAM",
            })
            continue
        if not so_inv:
            # 沒有 SO INV 但有金額：對帳單層級的調整 (例如貨損扣款
            # "10 kiện hàng bị ướt…" -1,666,304)，記在對帳單號底下。費用名稱用
            # DEC 欄的說明文字 (寫入時加英文翻譯)；DEC 只是單號時寫 Additional charge
            if amount:
                name = dec if " " in dec else "Additional charge"
                items.append({"description": name, "amount": amount, "vat": None,
                              "invoice_no": header["invoice_no"],
                              "invoice_date": header["invoice_date"]})
            continue
        # 併車的出貨 (同一天、同一個送貨地點) 後面幾列常不重複填車牌，沿用上一列
        # SHIP DATE 是合併儲存格時，後面幾列空白：沿用上一列的日期
        ship_date = _vn_date(_vn_cell(row, c_date)) or prev["date"]
        remark = _vn_text(_vn_cell(row, c_remark))
        same_group = (ship_date, remark) == (prev["date"], prev["remark"])
        cont = _vn_text(_vn_cell(row, c_cont)) or (prev["cont"] if same_group else None)
        plate = _vn_text(_vn_cell(row, c_plate)) or (prev["plate"] if same_group else None)
        prev = {"date": ship_date, "remark": remark, "cont": cont, "plate": plate}
        items.append({
            # 金額空白 = 原帳單漏填，視為 0 (使用者規定)
            "description": charge, "amount": amount if amount is not None else 0.0, "vat": None,
            "invoice_no": so_inv, "invoice_date": ship_date,
            "order_no": dec or None,
            "arrive_date": ship_date, "onboard_date": ship_date,
            "bl_no": so_inv, "mbl_no": so_inv,
            # 越南境內拖車，啟運港 / 目的港一律 VIETNAM (使用者答案)
            "port_of_loading": "VIETNAM", "port_of_discharge": "VIETNAM",
            # 材積統一寫法 (見 _fmt_volume)：CBM 欄 -> "0.26 M3"
            "volume": _fmt_volume(cbm=_vn_num(_vn_cell(row, c_cbm))),
            "vessel": plate or None, "container_no": cont or None,
        })
    if vat:
        items.append({"description": "VAT", "amount": None, "vat": round(vat, 2),
                      "invoice_no": header["invoice_no"], "invoice_date": header["invoice_date"]})
    return {"header": header, "items": items,
            "totals": {"total": total, "vat_total": vat, "grand_total": grand}}


def _ct_paid_on_behalf(rows: List[list]) -> Optional[Dict]:
    """代墊款 (CHI HO) INVOICE LIST。"""
    header = _ct_header(rows)  # 簽核欄 (清單日期) 在表格最下面
    head_i, cols = _ct_columns(rows, 0)
    if head_i is None or "SỐ HÓA ĐƠN" not in cols:
        return None
    c_content, c_net, c_vat = cols.get("NỘI DUNG"), cols.get("TIỀN TRƯỚC THUẾ"), cols.get("VAT")
    c_total, c_inv, c_cont = cols.get("TOTAL"), cols.get("INV"), cols.get("CONT")
    c_date, c_no, c_decl = cols.get("NGÀY"), cols.get("SỐ HÓA ĐƠN"), cols.get("TỜ KHAI")
    items: List[Dict] = []
    row_totals: Dict[str, float] = {}
    total = vat_total = bill_total = None
    # 日期欄只打日數 (例如 5)，Excel 顯示成 1900 年 1 月：用清單裡最常見的年月補上
    # (使用者 2026-10-28 確認 T05.2026 DOT1 的 "5" 是 2026 年 5 月)
    months = collections.Counter((v.year, v.month) for v in (_vn_cell(r, c_date) for r in rows[head_i + 1:])
                     if isinstance(v, (datetime.datetime, datetime.date)) and v.year > 1900)
    for row in rows[head_i + 1:]:
        cell = _vn_cell(row, c_date)
        if isinstance(cell, (datetime.datetime, datetime.date)) and cell.year == 1900 \
                and cell.month == 1 and months:
            year, month = months.most_common(1)[0][0]
            try:
                row = list(row)
                row[c_date] = datetime.date(year, month, cell.day)
            except ValueError:      # 該月沒有這一天：保留原值
                pass
        inv = _vn_text(_vn_cell(row, c_inv))
        decl = _vn_text(_vn_cell(row, c_decl))
        amount = _vn_amount(_vn_cell(row, c_net))
        # 合計列：通常 A 欄寫 TỔNG CỘNG；有的帳單漏寫 (例如 T01.2026 DOT1 第 119
        # 列)，只有金額、沒有日期/發票號碼/內容/INV/報關單號的列也是合計列
        no_text = not any(_vn_text(_vn_cell(row, c))
                          for c in (c_date, c_no, c_content, c_inv, c_decl))
        if _vn_label(_vn_cell(row, 0)).startswith("TỔNG") or (amount is not None and no_text):
            total, vat_total, bill_total = (_vn_num(_vn_cell(row, c))
                                            for c in (c_net, c_vat, c_total))
            break
        if amount is None:
            # 同一張第三方發票分攤到的其他出貨：只列 INV / 報關單號，併進上一張發票
            if items:
                for key, value in (("orders", inv), ("decls", decl)):
                    if value and value not in items[-1][key]:
                        items[-1][key].append(value)
            continue
        invoice_no = _vn_text(_vn_cell(row, c_no)) or header["invoice_no"]
        row_totals[invoice_no] = row_totals.get(invoice_no, 0.0) + (_vn_num(_vn_cell(row, c_total)) or 0)
        items.append({
            "invoice_no": invoice_no,
            "invoice_date": _vn_date(_vn_cell(row, c_date)) or header["invoice_date"],
            "description": _vn_text(_vn_cell(row, c_content)) or None,
            "amount": amount, "vat": _vn_amount(_vn_cell(row, c_vat)),
            "container_no": _vn_container(_vn_cell(row, c_cont)),
            "orders": [inv] if inv else [], "decls": [decl] if decl else [],
        })
    for item in items:
        # Order No. = 這張發票的所有 INV；B/L NO. 與 MBL NO. = 所有報關單號 (Tờ khai)
        item["order_no"] = ",".join(item.pop("orders")) or None
        item["bl_no"] = item["mbl_no"] = ",".join(item.pop("decls")) or None
    # 勾稽：只核對每一票 AMOUNT + VAT = 該列 TOTAL，不核對整份清單的 TỔNG CỘNG
    # (使用者 2026-10-28 規定；total / vat_total / bill_total 只用來找合計列)
    audit = [{"invoice_no": inv, "total": None, "grand_total": round(value, 2),
              "final_total": None} for inv, value in row_totals.items()]
    return {"header": header, "items": items, "audit": audit}


def _ct_parse_book(book) -> List[Dict]:
    """每個工作表各是一張對帳單 (header + items + 帳單本身的合計)。"""
    parsed = []
    for _name, rows in book.sheets:
        top = " ".join(_vn_label(v) for r in rows[:15] for v in r if _vn_label(v))
        result = _ct_paid_on_behalf(rows) if "INVOICE LIST" in top else _ct_statement(rows)
        if result and result["items"]:
            parsed.append(result)
    return parsed


def parse_cong_thanh_sheets(book) -> List[Dict]:
    return [{"header": p["header"], "items": p["items"]} for p in _ct_parse_book(book)]


def _cong_thanh_totals(book) -> List[Dict]:
    entries = []
    net = vat = gross = None
    for p in _ct_parse_book(book):
        if "audit" in p:            # 代墊款清單：整份 + 每張第三方發票
            entries.extend(p["audit"])
            continue
        # 運費 / 服務費對帳單：每列的 INVOICE NO 是各筆出貨的 SO INV，同一個檔案
        # 的 VC / DV 兩張對帳單共用出貨單號，所以整份檔案一起勾稽：所有費用 +
        # VAT = 各對帳單 Grand total 的加總 (使用者提示)
        t = p["totals"]
        if t.get("total") is not None:
            net = (net or 0) + t["total"]
        if t.get("vat_total") is not None:
            vat = (vat or 0) + t["vat_total"]
        if t.get("grand_total") is not None:
            gross = (gross or 0) + t["grand_total"]
    if net is not None or gross is not None:
        entries.append({"invoice_no": None, "total": net, "vat_total": vat,
                        "grand_total": gross, "final_total": None})
    return entries


register_supplier(
    "CONG_THANH", "CTY TNHHDV GN VT VA TM CONG THANH", detect_cong_thanh, None,
    sheet_parse=parse_cong_thanh_sheets, sheet_totals=_cong_thanh_totals,
)


# ===========================================================================
# 供應商 VN03：RỒNG Á CHÂU (CÔNG TY TNHH GIAO NHẬN THƯƠNG MẠI RỒNG Á CHÂU，
# 英文 Asia Dragon Cargo = 單號開頭的 "ADC"，廠商代號 100008978)
# ===========================================================================
# 每週一批 (PART 1~4) 拖車/報關服務費 Bảng kê，出口、進口分開檔案，一個檔案
# 可能有好幾個顯示中的工作表 (例如 LCL + FCL)，都要讀：
#   - LCL+LOCAL - IMPORT、FCL - IMPORT：左右兩個表，左邊是整櫃/整票原始費用，
#       右邊 (從第二個 "INVOICE No" 開始) 是依重量分攤到每張寶成發票的實際
#       請款金額，費用取右邊；CLEARING EXPENSE / TOTAL 是小計欄。
#       FCL 的 SERVICE ALL = SERVICE FEE + extra custom sheet，取 SERVICE ALL、
#       不取這兩個組成欄。TOTAL = 所有費用 + VAT PASS THROUGH。
#   - AIR - EXPORT：BẢNG KÊ PHÍ VẬN CHUYỂN HÀNG XUẤT AIR，費用分兩組
#       PHI VAN CHUYEN (DON GIA 運費 + PHỤ PHÍ XĂNG DẦU 燃油附加 + THANH TIEN
#       小計 + PHAT SINH 額外費用) 與 PHI DICH VU (DON GIA 單價 + THANH TIEN)。
#   - FCL, LCL, AIR IKEA - EXPORT：SERVICE CHARGE / TRUCKING FEE / 燃油附加 /
#       D/O,H/L,THC,CIC / LIFT ON / LIFTOFF / CSHT / FUMIGATION，TỔNG CỘNG 小計。
#   每個表最後一列是逐欄合計 (出口有 "TOTAL"/"TỔNG" 字樣，進口第一欄空白)。
# 檔案裡常留著隱藏的工作副本 ("(2)"、"CHECK XE"、"THEO BG"、"gốc"…)，核心
# 引擎只交出顯示中的工作表。
#
# 欄位對應 (依使用者提供的進口正確答案 test_samples/RONG_A_CHAU/sample1)：
#   INVOICE DATE   帳單沒有 -> N/A
#   INVOICE NO     每列的 INVOICE No (寶成發票號)；Order No. 留 N/A
#   SUPPLIER       越南 Top20 清單上的名稱；CONSIGNEE 依檔名 (POU SUNG…)
#   Description    費用欄表頭原文 (每筆出貨的每個非 0 費用欄一列)
#   VAT            LCL：獨立一列 "VAT"；FCL：獨立一列 "VAT PASS THROUGH"；
#                  進口每筆出貨都列，金額 0 時 VAT 欄為 N/A
#   Arrive Date    LCL = DATE C/D；FCL = DELIVERY DATE FCL
#   on board date  DATE C/D
#   B/L NO.        AWB,B/L No；MBL NO. = Declaration (報關單號)
#   Port of discharge  PORT
#   Volume         LCL = KGS + " KG" (例如 "877.78 KG")；FCL = 櫃型件數
#                  (20' 1 個 -> "1X20G")
#   Container no.  FCL 的 Cont no
# 出口表 (使用者還沒提供正確答案) 比照辦理：INVOICE NO = INVOICE/SO INVOICE、
# MBL NO. = SO TO KHAI、on board date = NGÀY ĐÓNG HÀNG、Port of Loading =
# PORT、Volume = 櫃型件數，沒有櫃就用 KGS。
# 金額照抄帳單 (分攤出來的金額有很多位小數，不四捨五入)。
# 金額勾稽：整份檔案逐欄「明細加總 = 合計列」(帳單只有整張表的合計列，沒有
# 每張發票的合計)，結果套用到檔案裡的每張發票。

_RAC_SUPPLIER = "CÔNG TY TNHH GIAO NHẬN THƯƠNG MẠI RỒNG Á CHÂU"
_VN_CONSIGNEES = {
    # 檔名關鍵字 -> 越南 Top 公司名稱 (Vietnam Supplier Top 20.xlsx「越南by公司」)
    "POU SUNG": "POU SUNG VIETNAM COMPANY LIMITED",
    "POUSUNG": "POU SUNG VIETNAM COMPANY LIMITED",
    "POU CHEN": "POU CHEN VIETNAM ENTERPRISE LTD.",
    "POUYUEN": "POUYUEN VIETNAM COMPANY LIMITED",
    "POU YUEN": "POUYUEN VIETNAM COMPANY LIMITED",
    "YUE DE": "YUE DE VIETNAM COMPANY LIMITED",
    "POU LI": "POU LI VIETNAM COMPANY LIMITED",
    "POU HUNG": "POU HUNG VIETNAM COMPANY LIMITED",
    "PRIME ASIA": "PRIME ASIA (VIETNAM) CO.,LTD.",
    # POU PHONG 不在 Top 20 清單 (2026-10-28 抽樣 "CÔNG NỢ POU PHONG T06.2026…"、資料夾
    # "POUPHONG\…")：全名使用者 2026-10-28 確認；給 POU PHONG 的發票不計入，
    # 由 validation 的 EXCLUDED_CONSIGNEES 列 ERROR
    "POU PHONG": "POU PHONG VIETNAM COMPANY LIMITED",
    "POUPHONG": "POU PHONG VIETNAM COMPANY LIMITED",
}

# 費用欄 (表頭轉大寫後完全相同)；PHI VAN CHUYEN/PHI DICH VU 兩組另外處理
_RAC_CHARGE_LABELS = {
    "SERVICE CHARGE", "SERVICE FEE", "TRUCKING FEE", "D/O, H/L, THC,CIC FEE", "LIFT ON",
    "LIFTOFF FEE", "LIFT ON/OFF FEE", "CSHT", "FUMIGATION FEE", "FREIGHT AIR PER P/L",
    "FREIGHT SEA PER P/L", "CFS", "CIC", "THC", "D/O", "HANDLING", "STORAGE", "INCURRED",
    "EXTRA CUSTOM SHEET", "DEMURRAGE", "PHAT SINH", "SERVICE ALL",
    # 2026-10-28 抽樣 POU PHONG 對帳表 ("CÔNG NỢ POU PHONG T03.2026 ĐỢT 1 -NHAP" 等) 的右表欄名
    "CFS,THC,CIC PER P/L", "D/O, H/L FEE PER P/L", "STORAGE PER P/L", "THC/CIC/ DO VS CONT",
    "LIFT ON/OFF", "CFS,THC, CIC FEE", "CFS,THC,CIC FEE", "D/O, H/L FEE",
}
# SERVICE ALL 存在時，它的組成欄不另外轉 (避免重複計算)
_RAC_SERVICE_ALL_PARTS = {"SERVICE FEE", "EXTRA CUSTOM SHEET"}
_RAC_VAT_LABELS = {"VAT PASS THROUGH"}
# 小計/合計欄，不轉成費用也不列入勾稽
_RAC_SUBTOTAL_LABELS = {"TỔNG CỘNG", "TOTAL", "CLEARING EXPENSE PER P/L", "CLEARING EXPENSE"}
# 描述欄位：表頭 -> 代號 (同名欄位以第一次出現的為準，左右表內容相同)
_RAC_META_LABELS = {
    "INVOICE": "invoice_no", "SO INVOICE": "invoice_no", "INVOICE NO": "invoice_no",
    "SỐ BILL": "bl_no", "AWB,B/L NO": "bl_no",
    "DECLARATION": "mbl_no", "SO TO KHAI": "mbl_no",
    "DATE C/D": "date_cd", "DELIVERY DATE FCL": "delivery_date",
    "NGÀY ĐÓNG HÀNG": "stuffing_date", "PORT": "port", "KGS": "kgs",
    "20'": "cont_20", "40'": "cont_40", "CONT NO": "container_no", "CBM": "cbm",
}


def detect_rong_a_chau(text: str) -> bool:
    upper = text.upper()
    return ("BẢNG KÊ PHÍ VẬN CHUYỂN" in upper or "ADC - JOB FILE" in upper
            or ("CLEARING EXPENSE" in upper and "VAT PASS THROUGH" in upper))


def _rac_consignee(path: str) -> Optional[str]:
    """帳單上沒有收貨方：依檔名判斷，檔名沒有再看所在資料夾 (例如
    '03-100008978 … RỒNG Á CHÂU-POUSUNG')。"""
    upper = os.path.basename(path).upper()
    if not any(key in upper for key in _VN_CONSIGNEES):
        upper = path.upper()
    return next((name for key, name in _VN_CONSIGNEES.items() if key in upper), None)


def _rac_sheet_kind(name: str) -> str:
    upper = _vt_norm(name)
    # 越南文 "NHAP" (nhập = 進口) 也算進口 (2026-10-28 抽樣 POU PHONG："NHAP FCL  DOT 1"、
    # "NHAP LCL+AIR+NKTC  DOT 1")；FCL 可能不在開頭，有 LCL 字樣時以 LCL 為準
    if "IMPORT" in upper or re.search(r"\bNHAP\b", upper):
        fcl = upper.strip().startswith("FCL") or (
            re.search(r"\bFCL\b", upper) and "LCL" not in upper)
        return "fcl_import" if fcl else "lcl_import"
    return "export"


def _rac_table(rows: List[list]):
    """找表頭並分類每一欄，回傳 dict 或 None：
      head (表頭列 index)、meta {代號: 欄號}、charges [(欄號, 費用名稱)]、
      vat_cols [欄號]、billing [列入勾稽的欄號]、sub_header (有第二列表頭)。"""
    head = next((i for i, r in enumerate(rows[:10])
                 if any(_vn_label(v) in ("INVOICE", "SO INVOICE", "INVOICE NO") for v in r)), None)
    if head is None:
        return None
    top = [_vn_label(v).replace('"', "'") for v in rows[head]]
    # 第二列表頭 (例如 DON GIA / 20' / 1 MT) 只有文字；有數字或日期就是資料列
    nxt = rows[head + 1] if head + 1 < len(rows) else []
    is_sub = any(isinstance(v, str) and v.strip() for v in nxt) and not any(
        isinstance(v, (int, float, datetime.date)) and not isinstance(v, bool) for v in nxt)
    sub = [_vn_label(v).replace('"', "'") for v in nxt] if is_sub else []
    width = max(len(top), len(sub))
    top += [""] * (width - len(top))
    sub += [""] * (width - len(sub))
    # 合併儲存格的群組名稱往右補，直到下一個群組或沒有子表頭的欄
    group, groups = "", []
    for j in range(width):
        if top[j]:
            group = top[j]
        elif not sub[j]:
            group = ""
        groups.append(group)
    labels = [sub[j] or top[j] for j in range(width)]

    # 左右兩個表 (進口)：實際請款的是第二個 "INVOICE NO" 開始的右表
    inv_cols = [j for j, lb in enumerate(top) if lb == "INVOICE NO"]
    # 2025-10 版左表有兩個 INVOICE NO 欄 (共三個)：右表從最後一個開始
    block_start = inv_cols[-1] if len(inv_cols) > 1 else 0
    has_service_all = "SERVICE ALL" in labels[block_start:]

    meta: Dict[str, int] = {}
    charges, vat_cols, components = [], [], set()
    for j, label in enumerate(labels):
        if not label:
            continue
        field = _RAC_META_LABELS.get(label)
        if field and field not in meta:
            meta[field] = j
        if j < block_start:
            continue
        # 費用名稱照抄表頭原文；PHI VAN CHUYEN 組的「DON GIA」就是運費本身，
        # THANH TIEN 是小計；PHI DICH VU 組取 THANH TIEN (DON GIA 是單價)
        original = _vn_text(rows[head + 1][j] if sub[j] else rows[head][j])
        charge = None
        if groups[j] == "PHI VAN CHUYEN":
            charge = {"DON GIA": "PHI VAN CHUYEN", "THANH TIEN": None}.get(label, original)
        elif groups[j] == "PHI DICH VU":
            charge = "PHI DICH VU" if label == "THANH TIEN" else None
        elif has_service_all and label in _RAC_SERVICE_ALL_PARTS:
            components.add(j)
        elif label in _RAC_CHARGE_LABELS or label.startswith("PHỤ PHÍ XĂNG DẦU"):
            charge = original
        if charge:
            charges.append((j, charge))
        elif label in _RAC_VAT_LABELS:
            vat_cols.append(j)
    if not charges:
        return None
    # 勾稽範圍：第一個費用欄之後、不是小計/組成/描述欄的所有數字欄 (出現
    # 新的費用欄卻沒被辨識時，明細加總會跟合計列對不上，不會悄悄漏掉)
    skip = components | set(meta.values()) | {
        j for j in range(width) if labels[j] in _RAC_SUBTOTAL_LABELS
        or (groups[j] == "PHI VAN CHUYEN" and sub[j] == "THANH TIEN")
        or (groups[j] == "PHI DICH VU" and sub[j] == "DON GIA")}
    billing = [j for j in range(charges[0][0], width) if j not in skip and j not in vat_cols]
    # 每列的總金額欄 (費用 + VAT PASS THROUGH)：進口 TOTAL、IKEA 出口 TỔNG CỘNG；
    # 空運出口沒有這一欄
    row_total = next((j for j in range(width - 1, block_start - 1, -1)
                      if labels[j] in ("TOTAL", "TỔNG CỘNG")), None)
    return {"head": head, "meta": meta, "charges": charges, "vat_cols": vat_cols,
            "billing": billing, "sub_header": is_sub, "row_total": row_total}


def _rac_volume(row: list, meta: Dict[str, int], kind: str) -> Optional[str]:
    """材積統一寫法 (見 _fmt_volume)：FCL 櫃量 (例如 '1X20G')；LCL 沒有櫃：
    CBM 欄 (例如 '2.5 M3')，沒有 CBM 才寫重量 (例如 '877.78 KG')。"""
    containers = []
    for field, size in (("cont_20", "20G"), ("cont_40", "40G")):
        n = _vn_num(_vn_cell(row, meta.get(field)))
        if n:
            containers.append((n, size))
    if containers:
        return _fmt_volume(containers=containers)
    if kind == "fcl_import":
        return None
    cbm = _vn_num(_vn_cell(row, meta.get("cbm")))
    kgs = _vn_num(_vn_cell(row, meta.get("kgs")))
    return _fmt_volume(cbm=_vn_text(cbm) if cbm else None, kg=_vn_text(kgs) if kgs else None)


def _rac_parse_sheet(name: str, rows: List[list]):
    """一個工作表 -> (items, 合計列的費用合計, 合計列的 VAT 合計,
    {發票號: 該發票各列 TOTAL 的加總，有任何一列沒填 TOTAL 時為 None})。"""
    table = _rac_table(rows)
    if table is None:
        return None
    kind = _rac_sheet_kind(name)
    meta = table["meta"]
    start = table["head"] + (2 if table["sub_header"] else 1)
    items: List[Dict] = []
    total = vat_total = 0.0
    invoice_totals: Dict[str, Optional[float]] = {}
    last_port = None
    for row in rows[start:]:
        invoice_no = _vn_text(_vn_cell(row, meta.get("invoice_no")))
        first = _vn_label(_vn_cell(row, 0))
        if not invoice_no or first.startswith(("TOTAL", "TỔNG")):
            # 合計列：明細之後第一列有金額、但沒有發票號的列
            nums = [_vn_num(_vn_cell(row, j)) for j in table["billing"] + table["vat_cols"]]
            if any(n for n in nums):
                total = sum(_vn_num(_vn_cell(row, j)) or 0 for j in table["billing"])
                vat_total = sum(_vn_num(_vn_cell(row, j)) or 0 for j in table["vat_cols"])
                break
            continue
        if table["row_total"] is not None:
            row_total = _vn_num(_vn_cell(row, table["row_total"]))
            has_charge = any(_vn_num(_vn_cell(row, j)) for j, _ in table["charges"])
            if row_total is None and has_charge:
                invoice_totals[invoice_no] = None  # 這張發票有列沒填 TOTAL，無法勾稽
            elif invoice_no not in invoice_totals or invoice_totals[invoice_no] is not None:
                invoice_totals[invoice_no] = (invoice_totals.get(invoice_no) or 0) + (row_total or 0)
        date_cd = _vn_date(_vn_cell(row, meta.get("date_cd")))
        shared = {
            "invoice_no": invoice_no,
            "bl_no": _vn_text(_vn_cell(row, meta.get("bl_no"))) or None,
            "mbl_no": _vn_text(_vn_cell(row, meta.get("mbl_no"))) or None,
            "container_no": _vn_container(_vn_cell(row, meta.get("container_no"))),
            "volume": _rac_volume(row, meta, kind),
        }
        # 同一批貨的後續列常不重複填港口：空白時沿用上一列 (日期/櫃號不沿用)
        port = _vn_text(_vn_cell(row, meta.get("port"))) or last_port
        last_port = port
        if kind == "export":
            shared.update(onboard_date=_vn_date(_vn_cell(row, meta.get("stuffing_date"))),
                          port_of_loading=port)
        else:
            arrive = date_cd if kind == "lcl_import" else \
                _vn_date(_vn_cell(row, meta.get("delivery_date")))
            shared.update(arrive_date=arrive, onboard_date=date_cd, port_of_discharge=port)
        for j, charge in table["charges"]:
            amount = _vn_num(_vn_cell(row, j))
            if amount:
                items.append({"description": charge, "amount": amount, "vat": None, **shared})
        vat = sum(_vn_num(_vn_cell(row, j)) or 0 for j in table["vat_cols"])
        if kind != "export":
            items.append({"description": "VAT" if kind == "lcl_import" else "VAT PASS THROUGH",
                          "amount": None, "vat": vat or None, **shared})
        elif vat:
            # 出口表也有 VAT PASS THROUGH 欄時 (2026-10-28 抽樣 POU PHONG "XKTC" 工作表)，
            # 有金額才列；原本沒有 VAT 欄的出口表不受影響
            items.append({"description": "VAT PASS THROUGH", "amount": None, "vat": vat, **shared})
    return items, total, vat_total, invoice_totals


def _rac_parse_book(book) -> Optional[Dict]:
    """所有顯示中工作表的明細合併成一組 (每列各自帶 INVOICE NO)。"""
    items: List[Dict] = []
    total = vat_total = 0.0
    invoice_totals: Dict[str, Optional[float]] = {}
    for name, rows in book.sheets:
        parsed = _rac_parse_sheet(name, rows)
        if parsed:
            items.extend(parsed[0])
            total += parsed[1]
            vat_total += parsed[2]
            for inv, value in parsed[3].items():
                known = invoice_totals.get(inv, 0.0)
                invoice_totals[inv] = None if value is None or known is None else known + value
    if not items:
        return None
    header = {"invoice_date": None, "supplier": _RAC_SUPPLIER,
              "consignee": _rac_consignee(book.path)}
    return {"header": header, "items": items, "invoice_totals": invoice_totals,
            "totals": {"total": total, "vat_total": vat_total, "grand_total": total + vat_total}}


def parse_rong_a_chau_sheets(book) -> List[Dict]:
    parsed = _rac_parse_book(book)
    return [{"header": parsed["header"], "items": parsed["items"]}] if parsed else []


def _rong_a_chau_totals(book) -> List[Dict]:
    """勾稽 (使用者規定的查驗方式)：同一張發票的所有費用 + VAT PASS THROUGH =
    該發票的 TOTAL (各列 TOTAL 加總，含稅總額 grand_total)。另外保留整份檔案
    逐欄「明細加總 = 表格合計列」(invoice_no=None：未稅 total 與含稅 final_total)，
    抓得到合計列公式範圍漏算這類問題。空運出口沒有 TOTAL 欄，只有整份檔案勾稽。"""
    parsed = _rac_parse_book(book)
    if not parsed:
        return []
    net, gross = parsed["totals"]["total"], parsed["totals"]["grand_total"]
    entries = [{"invoice_no": None, "total": net, "vat_total": parsed["totals"]["vat_total"],
                "grand_total": None, "final_total": gross}]
    entries += [{"invoice_no": inv, "total": None, "grand_total": value, "final_total": None}
                for inv, value in parsed["invoice_totals"].items() if value is not None]
    return entries


register_supplier(
    "RONG_A_CHAU", "CÔNG TY TNHH GIAO NHẬN THƯƠNG MẠI RỒNG Á CHÂU", detect_rong_a_chau, None,
    sheet_parse=parse_rong_a_chau_sheets, sheet_totals=_rong_a_chau_totals,
)


# ===========================================================================
# 供應商 ID08：OPL (ORIENTAL POWER LOGISTICS，廠商代號 100004377)
# ===========================================================================
# 印尼端由 PT. NCL INTER LOGISTIK INDONESIA 開立 "SALES REIMBURSEMENT"
# (CUSTOMER CODE : ORIPOLO02 = Oriental Power)。版面左右兩欄並排，每個欄位
# 抓到下一個已知標籤就停：
#   INV NO. / INV DATE (17/Sep/2026) / YOUR REF (= Order No.，使用者提示)
#   SHIPPER / CONSIGNEE / AWB/BL NO. / MAWB/OBL NO. / POL / POD / VESSEL
#   ATD (開船日) / ETA (到達日) / CONTAINER NO / 20'/40'/45' : 1 / 0 / 0 (材積 1X20G)
# 費用明細：DESCRIPTION QTY UNIT RATE CUR EX-RATE VAT AMOUNT，VAT 欄是稅率
# (0%)，每筆稅額 = 金額 × 稅率；合計區 SUB TOTAL / VAT / TOTAL : IDR 做勾稽。

def detect_opl(text: str) -> bool:
    """印尼當地帳單 (PT. NCL INTER LOGISTIK INDONESIA 開立，OPL 在印尼的代理)。
    同一家也會附提單 (B/L) 文件，上面也有 NCL 名稱但不是請款單，所以
    要同時有 "INV NO." 才算。OPL 美金 Debit Note 不在這裡，見 detect_opl_usd()。"""
    upper = text.upper()
    return "NCL INTER LOGISTIK" in upper and re.search(r"\bINV\s*NO\.", upper) is not None


def detect_opl_usd(text: str) -> bool:
    """OPL 美金 Debit Note (台灣曜捷運通開立，不是印尼當地帳單)，供應商 OPL_USD。"""
    upper = text.upper()
    # 掃描的空運 INVOICE ("OLCCGK251026")：預設 OCR 漏讀 FILE NO，會被誤認成 ORIENTAL_MULTIMODAL。
    # 有公司名稱 + INVOICE/DEBIT NOTE 標題 + HB/L 就算；擷取時 _opl_pages 會重新 OCR
    scanned_dn = re.search(r"ORIENTAL\s*POWER\s*LOGISTICS", upper) and re.search(r"[HAUB]B/L", upper) \
        and re.search(r"(?m)^\W*(?:OUTBOUND\s+)?(?:INVOICE|DEBIT\s+NOTE)\W*$", upper)
    return _is_opl_dn(text) or bool(scanned_dn)


# --- OPL 美金 Debit Note (使用者 2026-10-28：OPL 就是曜捷運通，"OUTBOUND DN" 就是
# OPL 美金 Debit Note) ---------------------------------------------------------------
# 台灣曜捷運通開的 "OUTBOUND INVOICE" / "OUTBOUND DEBIT NOTE"，金額是美金。抬頭的
# 公司名稱每個字重複 4 次 (粗體疊印)，不用。欄位：DATE (MM/DD/YYYY，有些沒有)、
# S/O#、FILE NO、HB/L、ON BOARD、POL、Dest、CNEE、VSL/VOY/LINE、(CBM)/(KGS)、CNTR NO。
# 明細：ITEM NAME QUANTITY UNIT PRICE AMT COLLECT AMT ("SEA FREIGHT 2.877 CBM USD 36.00
# 103.57 US$ 103.57"、數量可能寫 "MINI")，合計 "TOTAL: US$ 487.58" 或只有 "US$ 259.95"。
# 供應商名稱固定 _OPL_DN_SUPPLIER：資料庫用它判斷幣別是 USD (SUPPLIER_CURRENCY)。
_OPL_DN_SUPPLIER = "ORIENTAL POWER LOGISTICS CO., LTD"
_OPL_DN_ITEM_RE = re.compile(
    # 數量："2.877 CBM"、"1 SET"、"MINI"、整櫃 "1X20'GP"
    r"^(?P<desc>[A-Z][A-Z0-9/&().\- ]*?)\s+(?:[\d.,]+\s+[A-Z]+|MINI|\d+X\d{2}'?[A-Z]{0,2})\s+USD\s+[\d,.]+\s+"
    r"(?P<amount>[\d,.]+)\s+US\$\s*[\d,.]+\s*$")


def _is_opl_dn(text: str) -> bool:
    upper = text.upper()
    # 掃描檔 OCR 偶爾把 FILE NO 讀成 PILE NO。空運帳單 (AT- 開頭的 FILE NO，2026-10-29 OPL
    # 資料夾 "AASZX2509417 (PT.NIKOMAS SZX 2ND).pdf") 抬頭只寫 "INVOICE"，要有公司名稱才算
    # 標題讀成亂碼時 ("JKTHCSE2601010H")，有 S/O# 欄也算；HB/L 會讀成 AB/L ("SPOE2509124…")、
    # UB/L ("HKGJKT25090024")、BB/L ("HKGJKT25110676")
    # OCR 也會把 DEBIT 讀成 "DEB iT"、S/O# 讀成 "5/O#" ("OLJKT2509010")
    title = re.search(r"OUTBOUND\s+(?:INVOICE|DEB\s?IT\s+NOTE)", upper) or re.search(r"[S5]/O#", upper) or (
        re.search(r"(?m)^\s*INVOICE\s*$", upper) and "ORIENTAL POWER" in upper)
    return bool(title) and re.search(r"[FP]ILE NO", upper) is not None \
        and re.search(r"[HAUB]B/L", upper) is not None


def _opl_dn_total(text: str) -> Optional[float]:
    # 掃描檔 OCR 會多讀出 "=" / ":" ("US$ =: 1,320.00")，或在前面多讀出雜訊 ("- 7 US$ 639.20")。
    # 掃描檔有時把右邊的 PREPAID AMT 欄拆成一行一個 "US$ 435.00"，總額在最後 ("KHHJKT2511640")
    # 總額前面也會多一個 "."、"—-" 或雜字 ("a US$  . 357.75"、"HKGJKT25080662"；"US$ —- 176.00")
    found = re.findall(r"(?m)^[^A-Za-z\n]{0,8}?(?:[a-z]{1,2}\s+){0,2}(?:TOTAL\s*:\s*)?US\$\s*[=:.\s\-—~]*?\s*(\d[\d,.]*(?:\s\.\d{2})?)\s*$",
                       text)
    # 小數點前多一個空白 ("US$ 479 .36")
    return _opl_amount(found[-1].replace(" ", "")) if found else None


# --- OPL 美金 Debit Note 掃描檔 (2026-10-29 OPL 資料夾 "HKGJKT25110203(NIKOMAS HKG 62ND).pdf"
# 等約 800 份：第 1 頁 DN 掃描、第 2 頁提單) -----------------------------------------------
# 明細欄位：ITEM NAME QUANTITY UNIT PRICE AMT PAYER PREPAID AMT。預設 OCR 會被表頭下的虛線
# (掃描歪了約 0.5 度) 干擾，第一列 O/F 讀成亂碼；先轉正、擦掉橫線，再用 600 dpi + psm 4
# (400 dpi 會把 HB/L 的 HKGJKT 讀成 HRGJKT、FILE NO 讀成 PILE NO)。明細對不上總額時改用
# 500 dpi 黑白再讀一次。
_OPL_SCAN_DPI = 600
_OPL_SCAN_ITEM_RE = re.compile(
    # 費用名稱 + 數量 ("1X40'HQ"、"1x20'GP"、"3.570 CBM"、"1 SET"、"2 SHPT"、"MINI"；OCR 會把
    # "1 SET" 讀成 "| SET"/"1 SEI"、"2 SHPT" 讀成 "2. SHPT")，後面接 US$ 單價、AMT、PAYER、
    # PREPAID/COLLECT AMT
    # 名稱和數量之間會多讀出 "—-"、"~=—"、"_" ("O/F —-1X20'GP"、"SE25091000010")
    # 行首也會多讀出 "— "、"~ "、"“" ("OLJKT2511021" 的 "— DOC-POD")
    r"^[|.\s\-—–~“\"'_«‘;:\\]*(?P<desc>[A-Z0-9]\S*(?:\s+\S+)*?)[\s.,\-—–~=_'‘]+"
    # 也會把 "2X40'HQ" 讀成 "2440 'HQ"、"5BLS"/"4BLS"/"10BLS" 讀成 "SBLS"/"S5BLS"/"ABLS"/"LOBLS"、
    # "1 SET" 讀成 "{ SET"/"1s SET"/"lL SET"、MINI 讀成 "MINT"/"MINTI"/"MINI]"、"KGS USD" 黏成 "KGSUSD"
    # "7X40'HQ" 讀成 "TX40'HQ"/"7K40'HQ" ("XM25100326")
    r"(?P<qty>\d+\s*[Xx]\s*\d{2}\S*|\d+\s*4\s*\d{2}\s?'\S*|[\dTIl]?[XxK]\d{2}'\S*|[\d.,]+\s*KGS(?=USD)"
    r"|S?\d[\d.,]*[a-z]?\s*[A-Z]{2,5}|[|iIlt][lLI]?\s+[A-Z]{2,5}"
    # "1CNTR" 讀成 "ICNTR" ("VN25E1609"、"VN25E1682")；"1 SET" 讀成一串雜字 "=—S—s1sSSET"
    # ("JKTHCSE2509073J")：有數字、結尾是單位就算數量
    # "8 CBM" 讀成 "& CBM" ("SHKJKT25110571A")；品名本身也有 "&"，所以後面一定要接單位
    r"|\d?[{(\[\])]\s*[A-Z]{2,5}|&\s*(?:CBM|SETS?|SHPTS?|CNTRS?)(?=\s)|[LI1]?[SOA0](?:BLS?|SETS?|SHPTS?|CBM)|MIN[ITl1]{1,2}[\]|]?\d*"
    r"|[|iIl]CNTRS?|\S*\d\S*(?:SETS?|SHPTS?|CNTRS?)"
    # 數量沒有單位 ("CUSTOM CLEARANCE-WAN 1 US$ 130.00"、"PORT SECURITY-C 4.63 US$ 1.30"，
    # 文字版 "JKTHCSE2508031A"、"JKTHCSE2508057A-G")
    r"|\d[\d.,]*(?=\s+U[S$]))"
    r"(?:\s+|(?<=KGS)(?=USD))(?P<rest>.*U[S$]{1,2}.*)$")
_OPL_SCAN_MONEY_RE = re.compile(r"(?<![\d.])(\d{1,3}(?:,\s?\d{3})*\.\d{2})(?![\d])")


def _opl_scan_item(line: str) -> Optional[Dict]:
    """掃描 DN 的一列明細。金額：AMT 欄 (倒數第 2 個金額) 跟行尾 PREPAID AMT 一樣就用它；
    不一樣時 (OCR 把其中一欄讀錯，例如 PREPAID 800.00 讀成 300.00) 用 AMT 欄，交給總額稽核把關。"""
    m = _OPL_SCAN_ITEM_RE.match(line)
    if not m:
        return None
    # OCR 會把小數點讀成兩個 ("339..23"、"140..00")
    # 千分位逗號讀成分號 ("US$1;500.00")
    rest = re.sub(r"(\d)\s?\.\.\s?(\d)", r"\1.\2", m.group("rest"))
    rest = re.sub(r"(\d);(\d{3}\.)", r"\1,\2", rest)
    # 小數點漏讀成空白 ("USD 39 23 39 23 … US$ 339..23"、"326 00")
    rest = re.sub(r"(?<![\d.,])(\d{1,3}(?:,\d{3})*) (\d{2})(?![\d.,])", r"\1.\2", rest)
    nums = [float(n.replace(",", "").replace(" ", "")) for n in _OPL_SCAN_MONEY_RE.findall(rest)]
    if not nums:
        return None
    # 港幣等外幣費用 ("GATE CHARGE 2 SET HKD 570.00 1,140.00 PT.… US$ 147.49"，"HKGJKT25090022A")：
    # 單價、AMT 是外幣，只有行尾 PREPAID AMT 是美金，一定用它
    if re.match(r"\W*(?:HK[D)]|RMB|CNY|TWD|NTD?|JPY|EUR|SGD)\b", rest):
        return {"description": m.group("desc").strip(" .,"), "amount": nums[-1], "vat": None,
                "_alt": nums[-1]}
    amount = nums[-1] if len(nums) < 3 or nums[-1] == nums[-2] else nums[-2]
    # 另一欄讀到的不同金額 (_parse_opl_dn 用它讓明細對平總額)
    other = nums[-1] if amount != nums[-1] else (nums[-2] if len(nums) >= 2 else amount)
    return {"description": m.group("desc").strip(" .,"), "amount": amount, "vat": None, "_alt": other}


def _opl_deskew(gray):
    """投影法找傾斜角 (-1.5~1.5 度)：轉正後逐列平均亮度的變異數最大的角度。只用 PIL。"""
    from PIL import Image
    small = gray.resize((gray.width // 4, gray.height // 4))
    best, best_score = 0.0, -1.0
    for step in range(-15, 16):
        rows = list(small.rotate(step / 10, fillcolor=255).resize((1, small.height), Image.BOX).getdata())
        mean = sum(rows) / len(rows)
        score = sum((v - mean) ** 2 for v in rows)
        if score > best_score:
            best, best_score = step / 10, score
    return gray.rotate(best, fillcolor=255, resample=Image.BICUBIC) if best else gray


def _opl_erase_hlines(gray, ratio: float = 0.25):
    """表格橫線 (一列像素暗點超過 ratio) 塗白，上下各多擦 1 列。"""
    from PIL import Image, ImageDraw
    bw = gray.point(lambda v: 0 if v < 128 else 255)
    rows = list(bw.resize((1, bw.height), Image.BOX).getdata())
    draw = ImageDraw.Draw(gray)
    for y, v in enumerate(rows):
        if v < 255 * (1 - ratio):
            draw.rectangle([0, y - 1, gray.width, y + 1], fill=255)
    return gray


def _opl_scan_ocr(pdf_path: str, page_index: int, dpi: int = _OPL_SCAN_DPI, bw: bool = False,
                  psm: int = 4) -> str:
    import pdfplumber
    import pytesseract
    from PIL import ImageOps
    with pdfplumber.open(pdf_path) as pdf:
        image = pdf.pages[page_index].to_image(resolution=dpi).original
    gray = _opl_erase_hlines(_opl_deskew(ImageOps.autocontrast(image.convert("L"))))
    if bw:
        gray = gray.point(lambda v: 0 if v < 160 else 255)
    return pytesseract.image_to_string(gray, config=f"--psm {psm}") or ""


def _opl_scan_score(text: str) -> int:
    """OCR 結果好不好：2 = 明細對平總額且有 CONSIGNEE、發票日期、發票號碼；1 = 只對平總額；0 = 對不平。"""
    parsed = _parse_opl_dn(text)
    total = _opl_dn_total(text)
    if total is None or abs(sum(i["amount"] or 0 for i in parsed["items"]) - total) > 0.01:
        return 0
    header = parsed["header"]
    return 2 if header.get("consignee") and header.get("invoice_date") and header.get("invoice_no") else 1


def _opl_dn_scan_page(fast_text: str) -> bool:
    """快速 OCR 看起來是 DN 頁：有公司名稱或標題，加上 DN 的任一欄位。快速 OCR 常漏讀
    FILE NO ("OLCXML2508090C05"、"SE25092500012")，所以 S/O#、HB/L、ITEM NAME、US$ 也算。"""
    fast = fast_text.upper()
    return re.search(r"ORIENTAL\s*POWER|OUTBOUND|DEBIT\s*NOTE", fast) is not None \
        and re.search(r"[FP]ILE\s*NO|S/O#|HB/L|ITEM\s*NAME|US\$", fast) is not None


def _opl_pages(pdf_path: str) -> List[str]:
    """逐頁文字：只有「掃描的 DN 頁」自己 OCR (見上)，其他頁回傳空字串，退回預設擷取。"""
    import pdfplumber
    from extract_utils_streamlit import _ocr_page_fast, _stamp_only_scan
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_index, pg in enumerate(limited_pages(pdf)):
            # 整頁掃描圖上只疊了少量文字層 ("JKTHCSE2510070J" 的蓋章)：當掃描頁 (見 _stamp_only_scan)
            if any(c["text"].strip() for c in pg.chars) and not _stamp_only_scan(pg, pg.extract_text() or ""):
                # 文字版 DN 跨頁 ("SHKJKT25090149A" 第 2 頁只有明細尾段和總額)：前一頁是沒有總額
                # 的文字版 DN、這頁是沒有 FILE NO 的續頁時，接到前一頁；其他文字頁維持預設擷取
                text = pg.extract_text() or ""
                prev = (pdf.pages[page_index - 1].extract_text() or "") \
                    if page_index and not pages[-1] and "OUTBOUND" in text else ""
                if prev and _is_opl_dn(prev) and _opl_dn_total(prev) is None and not _is_opl_dn(text):
                    pages[-1] = prev + "\n" + text
                pages.append("")
                continue
            if not _opl_dn_scan_page(_ocr_page_fast(pdf_path, page_index, pg)[0]):
                pages.append("")
                continue
            text = _opl_scan_ocr(pdf_path, page_index)
            score = _opl_scan_score(text)
            others = []
            # 明細對不上總額、或缺 CONSIGNEE/發票日期時，依序改用 500 dpi 黑白、600 dpi psm 6
            # 重讀，分數較高才採用
            for retry in ({"dpi": 500, "bw": True}, {"psm": 6}):
                if score == 2:
                    break
                alt = _opl_scan_ocr(pdf_path, page_index, **retry)
                alt_score = _opl_scan_score(alt)
                if alt_score > score:
                    text, score, alt = alt, alt_score, text
                others.append(alt)
            pages.append(_opl_fill_header(text, others, pdf_path, page_index))
    # 明細太多印到第 2 頁 ("SHKJKT25110571A")：續頁只有標題和明細，沒有 FILE NO，總額在續頁。
    # 前一頁沒有總額時，把續頁接在後面，續頁本身改當附件
    for i in range(len(pages) - 1):
        if pages[i] and pages[i + 1] and _opl_dn_total(pages[i]) is None \
                and not _is_opl_dn(pages[i + 1]):
            pages[i] = pages[i] + "\n" + pages[i + 1]
            pages[i + 1] = ""
    return pages


def _opl_fill_header(text: str, others: List[str], pdf_path: str, page_index: int) -> str:
    """採用的 OCR 版本缺 CONSIGNEE/發票號碼/發票日期時，從其他版本 (同一頁) 補那一行；
    CONSIGNEE 還是沒有，再單獨讀 MESSR 那一行 (_opl_messr_band)。補的行加在最後。"""
    header = _parse_opl_dn(text)["header"]
    extra = []
    for key in ("consignee", "invoice_no", "invoice_date"):
        if header.get(key):
            continue
        for alt in others:
            line = _opl_header_lines(alt)[key]
            if line and _parse_opl_dn(line)["header"].get(key):
                extra.append(line)
                break
    if not header.get("consignee") and not any(_parse_opl_dn(ln)["header"].get("consignee") for ln in extra):
        messr = _opl_messr_band(pdf_path, page_index)
        if messr:
            extra.append(f"MESSR : {messr}")
    return "\n".join([text] + extra)
    return pages


def _opl_dn_date(value: Optional[str]) -> Optional[str]:
    """'09/18/2026' (月/日/年) 或 'SEP. 18,2026' -> '18 Sep 2026'。"""
    m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})$", value or "")
    if m:
        return datetime.date(int(m.group(3)), int(m.group(1)), int(m.group(2))).strftime("%d %b %Y")
    m = re.match(r"([A-Za-z]{3})\.?\s*(\d{1,2})\s*,\s*(\d{4})", value or "")
    return _format_date_en_from_month_name(m.group(2), m.group(1), m.group(3)) if m else value


# 使用者 2026-10-29：OPL 美金 DN 的 CONSIGNEE 統一成標準全名。掃描檔 OCR 會讀出各種寫法與雜字
# ("PT.POUYUEN INDONESTA"、"PT GLOSTAR INDONESIA a"、"PT. NIKOMAS GEMILANG ON BOARD: …")。
# 比對時只看大寫字母 (去掉空白、標點)；對不到的維持原值。
_OPL_STD_CONSIGNEES = [
    (r"NIK\w{0,2}AS|GEM[IT1L]LANG", "PT. NIKOMAS GEMILANG"),
    (r"GLOSTAR", "PT. GLOSTAR INDONESIA"),
    (r"POUY[UV][ER]?[NR]", "PT. POU YUEN INDONESIA"),
    (r"POUCHENI", "PT. POU CHEN INDONESIA"),
    (r"HARDASES|HARD\w{0,4}NDONE", "PT. HARDASES ABADI INDONESIA"),
    (r"KAYUEN", "PT. KA YUEN INDONESIA"),
    (r"CHANGYANG", "PT. CHANG YANG MATERIAL PLASTICS"),
]


def _opl_std_consignee(value: Optional[str]) -> Optional[str]:
    if not value:
        return value
    letters = re.sub(r"[^A-Z]", "", value.upper())
    for pattern, name in _OPL_STD_CONSIGNEES:
        if re.search(pattern, letters):
            return name
    return value


def _opl_cnee(text: str) -> Optional[str]:
    """CNEE 欄。文字版 PDF 有時 CNEE 跟收件人 email 疊印，抽出來變成一個字一個字交錯
    ("E P m T a . i l N : I K p O …"，"SHKJKT25090149A")：這種值不用，改用 MESSR。"""
    value = _search(r"CNEE\s*:\s*(.+?)(?:\s+CLOSING\s*:|$)", text)
    words = (value or "").split()
    if not words or "@" in value or sum(len(w) == 1 for w in words) > len(words) / 2:
        return None
    return value


def _opl_messr(text: str) -> Optional[str]:
    """MESSR (請款對象)。掃描檔同一行右邊的 ON BOARD / POL 會讀進來，後面也常接小寫雜字
    ("MESSR :PT. NIKOMAS GEMILANG ee neat aan")：公司名稱是大寫，遇到含小寫的字就停。"""
    value = _search(r"(?m)MESS[RK]\s*:\s*(.+?)(?:\s+(?:POL|ON\s*BOARD)\s*:|$)", text)
    if not value:
        return None
    # 集團公司直接用標準全名 (帳單也有大小寫混合的 "PT. Glostar Indonesia"，"HKGJKT25090023")
    std = _opl_std_consignee(value)
    if std != value:
        return std
    words = []
    for word in value.split():
        if re.search(r"[a-z]", word) or not re.search(r"[A-Z]{2}", word) and words[-1:] != ["PT."]:
            break
        words.append(word)
    # 只讀到 "PT." 不算 (讓 _opl_fill_header 從其他 OCR 版本或單行 OCR 補)
    name = " ".join(words).strip(" |,")
    return name if re.search(r"[A-Z]{3}", re.sub(r"^P[TY][.,]?", "", name)) else None


def _opl_header_lines(text: str) -> Dict[str, Optional[str]]:
    """表頭欄位的整行原文 (給 _opl_pages 從其他 OCR 版本補缺的行)。"""
    keys = {"consignee": r"(?:CNEE|MESS[RK])\s*:", "invoice_no": r"[FP]ILE NO\s*:",
            "invoice_date": r"^\s*DATE\s*:\s*\d|印表日期"}
    return {k: _search(rf"(?m)^(.*(?:{pat}).*)$", text) for k, pat in keys.items()}


def _opl_messr_band(pdf_path: str, page_index: int) -> Optional[str]:
    """MESSR 那行跟下一行地址貼得很近，psm 4 整行漏掉 ("HKGJKT25080687")：找到 ON BOARD 的
    位置，只切它下面那一條 (左邊 60%，避開右邊的 POL)，用 psm 7 單行 OCR。"""
    import pdfplumber
    import pytesseract
    from PIL import ImageOps
    with pdfplumber.open(pdf_path) as pdf:
        image = pdf.pages[page_index].to_image(resolution=_OPL_SCAN_DPI).original
    gray = _opl_deskew(ImageOps.autocontrast(image.convert("L")))
    top = gray.crop((0, 0, gray.width, gray.height // 2))
    data = pytesseract.image_to_data(top, config="--psm 4", output_type=pytesseract.Output.DICT)
    for i, word in enumerate(data["text"]):
        if word.upper().startswith("BOARD") and i and data["text"][i - 1].upper() == "ON":
            y, h = data["top"][i], data["height"][i]
            # 行距每份略有不同：依序試三種上下範圍，讀到 MESSR 就停
            for start, end in ((0.7, 2.5), (0.5, 2.2), (1.0, 2.9)):
                band = gray.crop((0, int(y + h * start), int(gray.width * 0.6), int(y + h * end)))
                messr = _opl_messr(pytesseract.image_to_string(band, config="--psm 7"))
                if messr:
                    # OCR 常把 INDONESIA 讀成 INDONESTA
                    return re.sub(r"\bINDONES[TI1l]A\b", "INDONESIA", messr).rstrip(".")
            return None
    return None


def _opl_swap_alt(items: List[Dict], total: float) -> None:
    import itertools
    gap = round(total - sum(it["amount"] for it in items), 2)
    diff = [it for it in items if round(it["_alt"] - it["amount"], 2)]
    if not gap or len(diff) > 12:
        return
    for size in range(1, len(diff) + 1):
        hits = [combo for combo in itertools.combinations(diff, size)
                if round(sum(it["_alt"] - it["amount"] for it in combo) - gap, 2) == 0]
        if hits:
            if len(hits) == 1:
                for it in hits[0]:
                    it["amount"] = it["_alt"]
            return


def _parse_opl_dn(text: str) -> Dict:
    onboard = _opl_dn_date(_search(r"ON BOARD\s*:\s*([A-Z]{3}\.?\s*\d{1,2}\s*,\s*\d{4})", text))
    kg = _search(r"\(KGS\)\s*:\s*([\d,.]+)", text)
    cntr = re.findall(r"CNTR NO\s*:\s*([A-Z]{4}\d{7})(?:/(\d{2}'?[A-Z]{0,2}))?", text)
    # 整櫃 (明細數量寫 "1X20'GP") 用櫃型；併櫃 (CFS/CBM 計價) 用 CBM
    fcl = re.search(r"\s\d+X\d{2}'", text) is not None
    header = {
        "supplier": _OPL_DN_SUPPLIER,
        # 使用者 2026-10-28：發票號碼 = FILE NO (例 CT-2609538)，發票日期 = 印表日期 (09/24/2026)。
        # 同一個 FILE NO 可能有好幾張 HB/L 各自請款 (抽樣 CT-2609648)：資料庫對這個供應商
        # 改用逐列判斷重複 (expense_normalize.SUPPLIER_NO_OWN_INVOICE)，不會少算
        # 掃描檔 OCR 會把 "IP-" 讀成 "[P-" ("TXGJKT2509567")
        "invoice_no": re.sub(r"^[\[|lI1]P-", "IP-", _search(r"[FP]ILE NO\s*:\s*(\S+)", text) or "") or None,
        # 沒有印表日期時用 DATE，再沒有用開船日
        "invoice_date": _opl_dn_date(_search(r"印表日期\s*:\s*(\d{1,2}/\d{1,2}/\d{4})", text))
        or _opl_dn_date(_search(r"(?m)^\s*DATE\s*:\s*(\d{1,2}/\d{1,2}/\d{4})", text)) or onboard,
        # 有些 DN 沒有 CNEE 欄 ("HKGJKT25080662")：改用 MESSR (請款對象)
        "consignee": _opl_std_consignee(_opl_cnee(text) or _opl_messr(text)),
        # 掃描檔 OCR 會在冒號前後多讀 "=" / ";" ("S/O# =: 25110203"、"HB/L =: HKGJKT…")
        # "S/O# :; 25110369" ("OESIN25100084")、"S/O# + 25100365" ("XM25100326")
        "order_no": _search(r"[S5]/O#\s*[=:;]*\s*:[\s;]*(\S+)", text) or _search(r"[S5]/O#\s*[=;+]+\s*(\d+)", text),
        "onboard_date": onboard,
        # 胡志明 → 雅加達的 HB/L 開頭是 JKTHCSE，OCR 會讀成 "JETHCSE"/"SKTHCSE" ("HARDASES跨月份更新帳單"、
        # "JKTHCSE2510070V")
        # 台中 → 雅加達的 TXGJKT 會讀成 "TAGJKT" ("TXGJKT2509567")
        # JKT 的 J 讀成 I ("SHKIKT26010049")
        "bl_no": re.sub(r"^([A-Z]{3})[I1l]KT(?=\d)", r"\1JKT", re.sub(r"^TAGJKT(?=\d)", "TXGJKT", re.sub(
            r"^(?:J[A-Z]|[S5]K)THCSE(?=\d)", "JKTHCSE",
            _search(r"[HAUB]B/L\s*[=s|]*\s*:\s*(\S+)", text) or ""))) or None,
        "port_of_loading": _search(r"POL\s*:\s*(.+?)\s*$", text),
        "port_of_discharge": _search(r"Dest\s*:\s*(.+?)\s*$", text),
        "vessel": _search(r"VSL/VOY/LINE\s*:\s*(.+?)(?:\s+\(Normal\))?(?:\s+Sales\s*:|$)", text),
        "container_no": ", ".join(c for c, _ in cntr) if cntr else None,
        "volume": _fmt_volume(containers=[(1, size) for _, size in cntr if size] if fcl else None,
                              cbm=_search(r"\(CBM\)\s*:\s*([\d.]+)", text),
                              kg=kg.replace(",", "") if kg else None),
    }
    items = [{"description": m.group("desc").strip(), "amount": _opl_amount(m.group("amount")), "vat": None}
             for m in (_OPL_DN_ITEM_RE.match(ln.strip()) for ln in text.splitlines()) if m]
    # 掃描檔 (見 _opl_pages)：只看 ITEM NAME 到 Remark 之間。表頭會讀成 "ITEM bh NAME"，
    # 或整行漏掉 ("OLJKT2509026")：這時從 P'KGS 那行之後開始。掃描檔也可能有一兩列剛好
    # 符合上面的文字版規則 ("KHHJKT2512649" 只抓到 DOCUMENT 一列)，取抓到較多列的結果
    block = _search(r"(?s)ITEM.{0,6}NAME[^\n]*\n(.*?)(?:\n\s*Remark|\Z)", text) \
        or _search(r"(?s)P'KGS[^\n]*\n(.*?)(?:\n\s*Remark|\Z)", text) or ""
    scanned = [it for it in (_opl_scan_item(ln.strip()) for ln in block.splitlines()) if it]
    total = _opl_dn_total(text)
    if scanned and total is not None:
        # AMT 欄和 PREPAID AMT 欄讀到不同金額 (例如 49.00 / 69.00) 時，找「換成另一欄就剛好
        # 對平總額」的列 (兩個值都是帳單上讀到的，不是推算)。可能好幾列都讀錯
        # ("HKGJKT25100107A")：試所有組合，最少列數的組合只有一個才換
        _opl_swap_alt(scanned, total)
        # OCR 把 PREPAID AMT 欄整欄拆開讀 ("PT.NIKOMAS GEMIL US$ 120.00" 一列一行)，明細行只剩單價
        # ("CFS CHARGE 3 CBM US$ 40.00"，2026-10-30 "TXGJKT2509567")：這欄的筆數跟明細一樣、
        # 加總剛好等於總額時，改用這欄金額
        if round(sum(it["amount"] for it in scanned) - total, 2):
            payer = [_opl_amount(a.replace(" ", "")) for a in re.findall(
                r"(?m)^\W*P[TY]\s*\.?\s*[A-Z][^\n]*?US\$\s*[=:.\s\-—~_]*(\d[\d,]*\.\d{2})\.?\s*$", text)]
            if len(payer) == len(scanned) and None not in payer and round(sum(payer) - total, 2) == 0:
                for it, amt in zip(scanned, payer):
                    it["amount"] = amt
    for it in scanned:
        it.pop("_alt")
    if len(scanned) > len(items):
        items = scanned
    return {"header": header, "items": items or [{"description": None, "amount": None, "vat": None}]}


_OPL_LABELS = [
    r"CUSTOMER\s*CODE", r"INV\s*NO\.", r"INV\s*DATE", r"JOB\s*NO\.", r"TERM", r"SALESMAN",
    r"REF\s*JOB\s*NO\.", r"ATTN", r"SHIPPER", r"VESSEL", r"CONSIGNEE", r"POL", r"POD",
    r"AWB/BL\s*NO\.", r"MAWB/OBL\s*NO\.", r"GROSS\s*WT", r"ATD", r"VOLUME", r"ETA",
    r"NO\.\s*OF\s*PKGS", r"CARRIER", r"YOUR\s*REF", r"CONTAINER\s*NO", r"INCOTERM",
]
_OPL_STOP = r"(?=\s+(?:" + "|".join(_OPL_LABELS) + r")\s*:|\s*\n|\s*$)"


def _opl_field(label: str, text: str) -> Optional[str]:
    return _search(rf"(?<![A-Z/]){label}\s*:[ \t]*(\S.*?){_OPL_STOP}", text)


def _opl_amount(value: Optional[str]) -> Optional[float]:
    if not value:
        return None
    try:
        return float(value.replace(",", ""))
    except ValueError:
        return None


_OPL_ITEM_RE = re.compile(
    r"^(?P<desc>[A-Z].*?)\s+(?P<qty>[\d.]+)\s+(?P<unit>\S+)\s+(?P<rate>[\d,.]+)\s+"
    r"(?P<cur>[A-Z]{3})\s+(?P<ex>[\d,.]+)\s+(?P<vat>[\d.]+)%\s+(?P<amount>[\d,.]+)\s*$")


def _opl_volume(text: str) -> Optional[str]:
    """材積統一寫法 (見 _fmt_volume)："20'/40'/45' : 1 / 0 / 0" -> "1X20G"；
    沒有櫃數用 "VOLUME : 27.450 M3"，再沒有用 "GROSS WT : 2,592.00 KGS"。"""
    m = re.search(r"20'/40'/45'\s*:\s*(\d+)\s*/\s*(\d+)\s*/\s*(\d+)", text)
    boxes = [(n, f"{size}G") for n, size in zip(m.groups(), ("20", "40", "45"))] if m else []
    kg = _search(r"GROSS WT\s*:\s*([\d,.]+)\s*KGS", text)
    return _fmt_volume(containers=boxes,
                       cbm=_search(r"VOLUME\s*:\s*([\d.]+)\s*M3", text),
                       kg=kg.replace(",", "") if kg else None)


def _opl_date(value: Optional[str]) -> Optional[str]:
    """INV DATE '17/Sep/2026' -> '17 Sep 2026'；ATD/ETA '04/09/2026' 原樣保留。"""
    m = re.match(r"(\d{1,2})/([A-Za-z]{3})/(\d{4})$", value or "")
    return _format_date_en_from_month_name(*m.groups()) if m else value


def _opl_split_dns(text: str) -> List[str]:
    """一份 PDF 有好幾張 DN (使用者 2026-10-30 改成逐張擷取："SHKJKT25080429-432,461,506"
    6 張共用一個 FILE NO、"SHKJKT25110300A"、"NIKOMAS 跨月份更新帳單")：在每個
    OUTBOUND INVOICE / DEBIT NOTE 標題切開。每段都要剛好一個 FILE NO，否則不切 (回傳 [text])。"""
    starts = [m.start() for m in re.finditer(r"(?mi)^[^\n]{0,6}OUTBOUND\s+(?:INVOICE|DEB\s?IT\s+NOTE)", text)]
    if len(starts) < 2:
        return [text]
    chunks = [text[a:b] for a, b in zip(starts, starts[1:] + [len(text)])]
    chunks[0] = text[:starts[1]]
    if any(len(re.findall(r"[FP]ILE NO", c)) != 1 for c in chunks):
        return [text]
    # 同一張 DN 掃描兩次 (FILE NO、HB/L、總額都相同) 只留一張 (比照 DHL，使用者 2026-10-28)
    seen, kept = set(), []
    for c in chunks:
        key = (_search(r"[FP]ILE NO\s*:\s*(\S+)", c), _search(r"[HAUB]B/L\s*[=s|]*\s*:\s*(\S+)", c),
               _opl_dn_total(c))
        if key not in seen:
            seen.add(key)
            kept.append(c)
    return kept


# 逐張 DN 擷取時，每列帶自己 DN 的欄位
_OPL_DN_ROW_FIELDS = ("invoice_no", "invoice_date", "consignee", "order_no", "onboard_date", "bl_no",
                      "port_of_loading", "port_of_discharge", "vessel", "container_no", "volume")


def parse_opl(text: str) -> Dict:
    if _is_opl_dn(text):
        chunks = _opl_split_dns(text)
        if len(chunks) == 1:
            return _parse_opl_dn(text)
        parsed = [_parse_opl_dn(c) for c in chunks]
        # 掃描頁沒讀到日期時，用同一個 FILE NO 其他張的日期 (同一次印表)
        dates = {p["header"]["invoice_no"]: p["header"]["invoice_date"]
                 for p in reversed(parsed) if p["header"]["invoice_date"]}
        items = []
        for p in parsed:
            h = p["header"]
            h["invoice_date"] = h["invoice_date"] or dates.get(h["invoice_no"])
            items += [dict(it, **{k: h.get(k) for k in _OPL_DN_ROW_FIELDS}) for it in p["items"]]
        return {"header": parsed[0]["header"], "items": items}
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    header = {
        "supplier": lines[0] if lines else None,
        "invoice_no": _opl_field(r"INV\s*NO\.", text),
        "invoice_date": _opl_date(_opl_field(r"INV\s*DATE", text)),
        # "SALES REIMBURSEMENT" 版：標籤和值分兩行 ("CONSIGNEE :\n: PT. POU CHEN INDONESIA POL …"，
        # "NCLR26080891")
        "consignee": _opl_field(r"CONSIGNEE", text)
        or _search(r"(?m)CONSIGNEE\s*:[ \t]*\n\s*:\s*(.+?)(?:\s+POL\b.*)?$", text),
        "order_no": _opl_field(r"YOUR\s*REF", text),
        "arrive_date": _opl_field(r"ETA", text),
        "onboard_date": _opl_field(r"ATD", text),
        "bl_no": _opl_field(r"AWB/BL\s*NO\.", text),
        "mbl_no": _opl_field(r"MAWB/OBL\s*NO\.", text),
        "port_of_loading": _opl_field(r"POL", text),
        "port_of_discharge": _opl_field(r"POD", text),
        "volume": _opl_volume(text),
        "vessel": _opl_field(r"VESSEL", text),
        "container_no": _opl_field(r"CONTAINER\s*NO", text),
        "shipper": _opl_field(r"SHIPPER", text),
    }
    items = []
    for idx, ln in enumerate(lines):
        m = _OPL_ITEM_RE.match(ln)
        if not m and 0 < idx < len(lines) - 1 and not re.search(r"\d", lines[idx - 1] + lines[idx + 1]):
            # 費用名稱太長，數字自成一行，名稱分在上下兩行 ("FUEL\n1.0000 SHPT 88.200 USD … 1,532,475.88
            # \nSURCHARGE_CRECIMIENTO"，"COAU7269061950 SR-JKTHCSE2605025A")
            m = _OPL_ITEM_RE.match(f"{lines[idx - 1]} {lines[idx + 1]} {ln}")
        if m:
            amount = _opl_amount(m.group("amount"))
            rate = float(m.group("vat"))
            items.append({"description": m.group("desc"), "amount": amount,
                          "vat": round(amount * rate / 100, 2) if amount is not None else None})
    vat_total = _opl_vat_total(text)
    if items and vat_total is not None:
        # 合計區的 VAT 另列一筆 (使用者答案)
        items.append({"description": "VAT", "amount": None, "vat": vat_total})
    return {"header": header, "items": items}


def _opl_extract_total(text: str) -> Optional[float]:
    if _is_opl_dn(text):
        # 好幾張 DN：各張總額加總 (有一張讀不到總額就不稽核)
        totals = [_opl_dn_total(c) for c in _opl_split_dns(text)]
        return round(sum(totals), 2) if None not in totals else None
    return _opl_amount(_search(r"SUB\s*TOTAL\s*:\s*([\d,.]+)", text))


def _opl_vat_total(text: str) -> Optional[float]:
    return _opl_amount(_search(r"(?m)^\s*VAT\s*:\s*([\d,.]+)", text))


def _opl_extract_grand_total(text: str) -> Optional[float]:
    return _opl_amount(_search(r"(?<!SUB )TOTAL\s*:\s*IDR\s*([\d,.]+)", text))


register_supplier(
    "OPL", "ORIENTAL POWER LOGISTICS", detect_opl, parse_opl,
    extract_total=_opl_extract_total,
    extract_vat_total=_opl_vat_total,
    extract_grand_total=_opl_extract_grand_total,
    # 美金 DN 掃描檔：DN 頁自己 OCR，後面的提單頁當附件 (NCL 帳單沒有 DN 頁，維持全部頁面)
    page_text=_opl_pages, invoice_page=_is_opl_dn,
)
# 美金 DN 拆出成獨立供應商 OPL_USD (2026-10-31，使用者要求)：解析邏輯共用上面的 parse_opl
# (它會走 _is_opl_dn 分支)，只有偵測與供應商 key 不同。OPL (NCL 印尼當地帳單) 不再吃 DN。
register_supplier(
    "OPL_USD", "OPL曜捷", detect_opl_usd, parse_opl,
    extract_total=_opl_extract_total,
    extract_vat_total=_opl_vat_total,
    extract_grand_total=_opl_extract_grand_total,
    page_text=_opl_pages, invoice_page=_is_opl_dn,
)


# ===========================================================================
# 供應商編號 (依進度管控的 Top 供應商清單)
# ===========================================================================
# 印尼：Desktop\印尼進度管控 的「Indonesia Supplier Top 15, 2025-2026.xlsx」
# 與「印尼進度管控表」占比工作表 (2025-2026 請款金額排名 + 廠商代號)。
#   ID01~ID15 = 印尼 TOP15；TOP15 以外沿用同一份排名 (例如 FEMARIA 第 18 名
#   = ID18)。ID08 ORIENTAL POWER LOGISTICS 的 key 是 OPL (2026-10 建立)。
# 越南：Desktop\越南進度管控 的「Vietnam Supplier Top 20.xlsx」(01~20，資料夾
#   名稱「編號-廠商代號 公司名」)，建好規則後一樣在這裡登記 VN01~VN20。
# 編號只是對照資料 (顯示在網頁清單、轉檔畫面與轉檔總表)，不影響辨識順序。
# ===========================================================================
# 供應商 VN08：CÔNG TY TNHH DSV AIR & SEA (越南 DSV，2026-10 新增)
# ===========================================================================
#
# 兩種版型 (使用者 2026-10 樣本 SEP.2026 SAMPLE 1 / SAMPLE 2)：
#   (a) PRO FORMA INVOICE：一份 PDF 一張發票 (可跨頁，"Continued Over…")，版面
#       跟印尼 DSV 相同。INVOICE NO = SHIPMENT 號碼；SUPPLIER = 頁尾 Address 的
#       "DSV Air & Sea Company Limited"；CONSIGNEE = 抬頭 "INVOICE DATE" 左邊的
#       客戶名稱；費用名稱取 " - " 說明之前 (名稱本身可能有 " - "，例如
#       "Phí lao vụ - hàng xuất - Greater of …")；VAT 是每筆 "8%=204.000" 的金額。
#       空運：FLIGHT 第一段是航班號 (Vessel)、MAWB = MBL、HAWB = B/L。
#       稽核：SUBTOTAL = 費用加總；TOTAL VND = 費用 + VAT。
#   (b) BẢNG KÊ (越南文對帳表)：一列一票，INVOICE NO = Số SWB (也是 MBL)，
#       B/L = FCR，每票拆 4 筆費用 (Phí phát hành vận đơn / Phí VGM / Phí CFS /
#       Phí THC) + 1 筆 VAT ("Thuế 8% VAT")。RATE 欄是匯率，不取。日期只有
#       "3-Sep"，年份取 SO# 最後 4 碼 ("SOA06092026")。
#       ❗ 文字層的數字常把第一位拆開 ("2 49,793" = 249,793、CBM "5 3.471" =
#       53.471)：1~2 位數、沒有逗號/小數點的字接到下一個字。
#       稽核 (每票)：4 筆費用 = Thành tiền trước thuế；費用 + VAT = Tổng thành tiền。

def detect_dsv_vn(text: str) -> bool:
    upper = text.upper()
    return "DSV AIR & SEA" in upper and not _is_dhl_vnpt(text) and (
        "VND" in upper or "VIỆT NAM" in upper or "VIET NAM" in upper)


# 對帳表的一票：SWB FCR SO# [PO#] D-Mon 數字…
_DSV_VN_ROW_RE = re.compile(
    r"^(?P<swb>S[A-Z]{3}\d{6,8})\s+(?P<fcr>[A-Z]{3}\d{6,8})\s+(?P<so>\S+)\s+"
    r"(?:(?P<po>\S+)\s+)?(?P<day>\d{1,2})-(?P<mon>[A-Za-z]{3})\s+(?P<rest>[\d.,\s]+)$",
    re.MULTILINE)
_DSV_VN_FEES = ("Phí phát hành vận đơn", "Phí VGM", "Phí CFS", "Phí THC")


def _dsv_vn_dot_amount(text: Optional[str]) -> Optional[int]:
    """越南式數字 '2.550.000' (句點千分位) -> 2550000。"""
    if not text:
        return None
    digits = text.replace(".", "")
    return int(digits) if digits.isdigit() else None


def _dsv_vn_merge_split_numbers(tokens: List[str]) -> List[str]:
    """把被拆開的第一位數字接回去："2 49,793" -> "249,793"。"""
    out, prefix = [], ""
    for i, tok in enumerate(tokens):
        if re.fullmatch(r"\d{1,2}", tok) and i + 1 < len(tokens):
            prefix += tok
            continue
        out.append(prefix + tok)
        prefix = ""
    return out


def _dsv_vn_table_rows(text: str) -> List[Dict]:
    """對帳表每一票：{swb, fcr, po, date, cbm, fees[4], subtotal, vat, total}。"""
    rows = []
    for m in _DSV_VN_ROW_RE.finditer(text):
        tokens = m.group("rest").split()
        if tokens and re.fullmatch(r"\d{1,2}", tokens[-1]):
            tokens = tokens[:-1]                     # 最後的 Brand 代碼
        nums = _dsv_vn_merge_split_numbers(tokens)
        if len(nums) != 9:
            continue
        cbm, _rate, *money = nums
        money = [int(v.replace(",", "")) if v.replace(",", "").isdigit() else None
                 for v in money]
        year = re.search(r"(\d{4})$", m.group("so"))
        rows.append({
            "swb": m.group("swb"), "fcr": m.group("fcr"), "po": m.group("po"),
            "date": _format_date_en_from_month_name(
                m.group("day"), m.group("mon"), year.group(1) if year else None),
            "cbm": cbm, "fees": money[:4], "subtotal": money[4],
            "vat": money[5], "total": money[6],
        })
    return rows


def _parse_dsv_vn_table(text: str) -> Dict:
    companies = re.findall(r"^(CÔNG TY TNHH .+?)\s*$", text, re.MULTILINE)
    header: Dict[str, Optional[str]] = {code: None for code in (
        "arrive_date", "onboard_date", "port_of_loading", "port_of_discharge",
        "vessel", "container_no", "shipper")}
    header["supplier"] = companies[0] if companies else None
    header["consignee"] = companies[1] if len(companies) > 1 else None
    rate = re.search(r"thuế\s+(\d+%)\s+VAT", text, re.IGNORECASE)
    vat_desc = f"Thuế {rate.group(1)} VAT" if rate else "Thuế VAT"
    items = []
    for r in _dsv_vn_table_rows(text):
        common = {"invoice_no": r["swb"], "invoice_date": r["date"], "order_no": r["po"],
                  "bl_no": r["fcr"], "mbl_no": r["swb"], "volume": f"{r['cbm']} M3"}
        for desc, amount in zip(_DSV_VN_FEES, r["fees"]):
            items.append({"description": desc, "amount": amount, "vat": None, **common})
        items.append({"description": vat_desc, "amount": None, "vat": r["vat"], **common})
    return {"header": header, "items": items or [{"description": None, "amount": None}]}


#   (c) VAT INVOICE (電子增值稅發票，使用者 2026-10-27 新增樣本 "08-sample DSV\1_001_C26TPS_*.pdf")：
#       一份一頁一張。INVOICE NO = Số (Invoice No.) (只用號碼，不加 Serial 1C26TPS)；
#       B/L = "Số hợp đồng / Vận đơn"；訂單號碼 = "Tên phương tiện" (FIA/VV01210087)。
#       費用行 "1 Phí chứng từ hàng nhập (Delivery order fee) 1 530.000 530.000 8% 42.400"，
#       每筆自帶稅額；名稱太長時折到上下兩行 ("5 1 1.918.571 …" 這行沒有名稱)。
#       稽核：Cộng doanh thu (Amount) = 費用加總；Tổng cộng tiền thanh toán = 費用 + 稅額。
_DSV_VN_EINV_ITEM_RE = re.compile(
    r"^\d+\s+(?:(?P<desc>.+?)\s+)?(?P<qty>\d[\d.,]*)\s+(?P<price>[\d.]+)\s+(?P<amount>[\d.]+)\s+"
    r"(?P<rate>\d+%|KCT)\s+(?P<tax>[\d.]+)$")


def _is_dsv_vn_einvoice(text: str) -> bool:
    return "(VAT INVOICE)" in text and "Cộng doanh thu" in text


def _parse_dsv_vn_einvoice(text: str) -> Dict:
    h: Dict[str, Optional[str]] = {"supplier": "CÔNG TY TNHH DSV AIR & SEA"}
    # 頁面擷取有時把號碼排到標籤的上一行 ("104817\nSố (Invoice No.) :")
    h["invoice_no"] = (_search(r"Số \(Invoice No\.\)\s*:\s*(\d+)", text)
                       or _search(r"^(\d+)\s*\nSố \(Invoice No\.\)", text))
    m = re.search(r"Ngày \(day\)\s*(\d{1,2})\s*tháng \(month\)\s*(\d{1,2})\s*năm \(year\)\s*(\d{4})", text)
    h["invoice_date"] = _format_date_en(*m.groups()) if m else None
    h["consignee"] = _search(r"Tên khách hàng \(Buyer\)\s*:\s*(.+)", text)
    h["order_no"] = _search(r"\(Mean of transportation\)\s*:\s*(\S+)", text)
    h["bl_no"] = _search(r"\(Contract / Bill of lading No\.\)\s*:\s*(\S+)", text)
    lines = [ln.strip() for ln in text.splitlines()]
    start = next((i for i, ln in enumerate(lines) if ln.startswith("(No.) (Description)")), 0)
    items = []
    for i in range(start + 1, len(lines)):
        if lines[i].startswith("Cộng doanh thu"):
            break
        m = _DSV_VN_EINV_ITEM_RE.match(lines[i])
        if not m:
            continue
        desc = m.group("desc")
        if not desc:
            # 名稱折行：上一行 + 下一行 (都不是費用行)
            before = lines[i - 1] if i - 1 > start and not _DSV_VN_EINV_ITEM_RE.match(lines[i - 1]) else ""
            after = (lines[i + 1] if i + 1 < len(lines) and not _DSV_VN_EINV_ITEM_RE.match(lines[i + 1])
                     and not lines[i + 1].startswith("Cộng doanh thu") else "")
            desc = f"{before} {after}".strip()
        items.append({"description": desc, "amount": _dsv_vn_dot_amount(m.group("amount")),
                      "vat": _dsv_vn_dot_amount(m.group("tax"))})
    return {"header": h, "items": items or [{"description": None, "amount": None, "vat": None}]}


#   (d) 每週合輯的對帳表 (2026-10-28 抽樣 "DSV AIR & SEA_JAN.2026_BY WEEKLY"，BẢNG KÊ ĐÍNH KÈM
#       HÓA ĐƠN)：費用欄比 (b) 多 (TIO fee、OT、BKG fee)，空白欄從文字分不出是哪一欄；SWB 也
#       可能是純數字 (空運 507586059429)。_cargowise_pages() 用 pdfplumber 表格附上
#       "##DSVHDR## 欄名 | …" 與 "##DSVROW## 值 | …"，這裡依欄名對應：RATE 與
#       Thành tiền trước thuế 之間是費用，Thuế 8% VAT 單獨一筆，每票 (SWB) 自己勾稽。
def _dsv_vn_grid_lines(page) -> str:
    out = ""
    for table in page.extract_tables() or []:
        if not table or not any("SOA" in (c or "") for c in table[0]):
            continue
        cell = lambda c: re.sub(r"\s+", " ", c or "").strip()  # noqa: E731
        out += "\n##DSVHDR## " + " | ".join(cell(c) for c in table[0])
        for row in table[1:]:
            out += "\n##DSVROW## " + " | ".join(cell(c) for c in row)
    return out


def _dsv_vn_grid_num(value: str) -> Optional[int]:
    digits = re.sub(r"[\s,]", "", value or "")
    return int(digits) if digits.isdigit() else None


def _dsv_vn_grid_rows(text: str) -> List[Dict]:
    rows, labels = [], []
    for line in text.splitlines():
        if line.startswith("##DSVHDR## "):
            labels = line[len("##DSVHDR## "):].split(" | ")
            continue
        if not line.startswith("##DSVROW## ") or not labels:
            continue
        cells = dict(zip(labels, line[len("##DSVROW## "):].split(" | ")))
        swb = cells.get("Số SWB", "")
        if not re.fullmatch(r"[A-Z0-9]{8,}", swb):
            continue                                       # TOTAL 列
        start = labels.index("RATE") + 1 if "RATE" in labels else None
        end = next((i for i, lb in enumerate(labels) if lb.startswith("Thành tiền")), None)
        vat_label = next((lb for lb in labels if lb.startswith("Thuế")), None)
        total_label = next((lb for lb in labels if lb.startswith("Tổng")), None)
        if start is None or end is None:
            continue
        soa = re.search(r"(\d{4})$", cells.get("SOA #", ""))
        day = re.fullmatch(r"(\d{1,2})-([A-Za-z]{3})", cells.get("DATE", ""))
        rows.append({
            "swb": swb, "fcr": cells.get("FCR") or None, "po": cells.get("PO#") or None,
            "date": _format_date_en_from_month_name(day.group(1), day.group(2), soa.group(1))
            if day and soa else None,
            "cbm": cells.get("CBM") or None,
            "fees": [(labels[i], _dsv_vn_grid_num(cells.get(labels[i]))) for i in range(start, end)],
            "subtotal": _dsv_vn_grid_num(cells.get(labels[end])),
            "vat_desc": vat_label, "vat": _dsv_vn_grid_num(cells.get(vat_label)) if vat_label else None,
            "total": _dsv_vn_grid_num(cells.get(total_label)) if total_label else None,
        })
    return rows


def _parse_dsv_vn_grid(text: str) -> Dict:
    doc = _parse_dsv_vn_table(text)          # 抬頭 (供應商、客戶) 沿用 (b)
    items = []
    for r in _dsv_vn_grid_rows(text):
        common = {"invoice_no": r["swb"], "invoice_date": r["date"], "order_no": r["po"],
                  "bl_no": r["fcr"], "mbl_no": r["swb"], "volume": f"{r['cbm']} M3" if r["cbm"] else None}
        items += [{"description": desc, "amount": amount, "vat": None, **common}
                  for desc, amount in r["fees"] if amount]
        if r["vat"] is not None:
            items.append({"description": r["vat_desc"], "amount": None, "vat": r["vat"], **common})
    doc["items"] = items or [{"description": None, "amount": None}]
    return doc


def parse_dsv_vn(text: str) -> Dict:
    if "##DSVROW## " in text:
        return _parse_dsv_vn_grid(text)
    if _is_dsv_vn_einvoice(text):
        return _parse_dsv_vn_einvoice(text)
    if _DSV_VN_ROW_RE.search(text):        # 只判斷版型，不先解析整張表
        return _parse_dsv_vn_table(text)
    return _cw_parse(text, _search(r"(DSV Air & Sea Company Limited)", text))


def _dsv_vn_extract_total(text: str):
    """未稅總額：對帳表是每票 Thành tiền trước thuế；PRO FORMA 是每張 SUBTOTAL。"""
    if "##DSVROW## " in text:
        return {r["swb"]: r["subtotal"] for r in _dsv_vn_grid_rows(text)}
    if _is_dsv_vn_einvoice(text):
        return _dsv_vn_dot_amount(_search(r"Cộng doanh thu trên hóa đơn \(Amount\)\s*:\s*([\d.]+)", text))
    rows = _dsv_vn_table_rows(text)
    return {r["swb"]: r["subtotal"] for r in rows} if rows else _cw_extract_total(text)


def _dsv_vn_extract_grand_total(text: str):
    if "##DSVROW## " in text:
        return {r["swb"]: r["total"] for r in _dsv_vn_grid_rows(text)}
    if _is_dsv_vn_einvoice(text):
        return _dsv_vn_dot_amount(_search(r"\(Total amount\)\s*:\s*([\d.]+)", text))
    rows = _dsv_vn_table_rows(text)
    return {r["swb"]: r["total"] for r in rows} if rows else _cw_extract_grand_total(text)


# ===========================================================================
# 越南 CargoWise 帳單共用 (VN08 DSV PRO FORMA、VN09 MOL、VN13 DHL、VN20 YUSEN)
# ===========================================================================
#
# 這幾家都用 CargoWise 系統開單，版面幾乎一樣 (PRO FORMA INVOICE / Debit Note)：
#   - 一份 PDF 可能有好幾張發票，每張從 "Page 1 of N" 開始 (可跨頁，
#     "Continued Over…")，用它切開；每筆費用帶上自己那張的抬頭欄位。
#   - INVOICE NO = SHIPMENT 號碼 (常跟標籤黏在一起 "SHIPMENTSSGN0315854")。
#   - 費用行有兩種：
#     (a) "B/L FEE (ORG) - 1 House Bill(s) @ VND 8%=88.000 1.100.000"：名稱取
#         " - 計費說明" 之前 (名稱本身可能有 " - "，例如 DHL "Org. Terminal
#         Handling - Carrier")，VAT 是 "=" 後的金額，句點千分位。
#     (b) MOL "Phí vận đơn 1 - Shipment 1,000,000 8 % VND 1,080,000"：最後的
#         AMOUNT 是含稅，金額 = Qty × RATE、VAT = 含稅 - 金額；非 VND 的費用
#         名稱加註幣別 "(USD)" (使用者 2026-10 正確答案)。
#   - 稽核：每張 SUBTOTAL / TOTAL VND；MOL 只有 GRAND TOTAL (每個幣別一行，
#     加總比對)。
# 版面要靠字元大小/位置才分得開，見 _cargowise_pages()。

def _cargowise_pages(pdf_path: str) -> List[str]:
    """逐頁文字，先清掉兩種雜訊再擷取：
    1. 大型 "DRAFT" 浮水印 (字高 > 30，例如 YUSEN 的 F/A/R/D 會混進提單號碼)。
    2. MOL 疊在越南文費用名稱上的費用代碼 (8 號字，名稱是 7 號字，例如 "BL" +
       "Phí vận đơn" 被讀成 "BPLhFí vận đơn")。
    另外在頁尾加一行 "##SHIPPER## <出貨方>"：SHIPPER / CONSIGNEE 兩欄並排，
    只看文字分不出界線，用 CONSIGNEE 標題的 x 座標切開。"""
    import pdfplumber
    from extract_utils_streamlit import _ocr_page_with_psm
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_index, pg in enumerate(limited_pages(pdf)):
            if not any(c["text"].strip() for c in pg.chars):
                # 掃描頁 (YUSEN 把整份帳單掃描附在後面)：400 dpi + psm 6 才讀得對金額
                # (使用者 2026-10-27 要求掃描頁也要到等級高)
                text = _ocr_page_with_psm(pdf_path, page_index, 6)
                ship = _cw_ocr_shipment(pg)
                if ship:
                    text = f"SHIPMENT {ship}\n{text}"
                pages.append(text + _CW_PAGE_MARK)
                continue
            # 粗體疊印三次的頁 ("BBB///LLL FFFEEEEEE")：同位置重複的字只留一個。
            # dedupe_chars() 很慢 (一頁約 0.8 秒)，先確認真的有疊印才做
            seq = sorted(pg.chars, key=lambda c: (round(c["top"]), c["x0"]))
            dup = sum(1 for a, b in zip(seq, seq[1:])
                      if a["text"] == b["text"] and abs(a["x0"] - b["x0"]) < 1
                      and abs(a["top"] - b["top"]) < 1)
            if dup * 10 > len(seq):
                pg = pg.dedupe_chars()
            chars = [c for c in pg.chars if c.get("size", 0) <= 30]
            small = [c for c in chars if c.get("size", 0) < 7.5 and c["x0"] < 60]
            drop = set()
            for c in chars:
                # 代碼欄寬到 x≈60 ("O-CWAF")
                if c.get("size", 0) >= 7.9 and c["x0"] < 62:
                    if any(abs(s["top"] - c["top"]) <= 2 for s in small):
                        drop.add((round(c["x0"], 2), round(c["top"], 2), c["text"]))
            # MOL：貨物說明 (8 號字) 太長，溢出到費用表 (7 號字) 跟費用行疊在一起
            # (2026-10-28 抽樣 "MOL LOGICTICS_APR.2026_rv" SSEM11Z01827 "Phí khai AicFeR No…")。
            # 費用表 (第一個 7 號字幣別 "VND" 那行) 到 "Sub Total" 之間只留小字
            tail = pg.search(r"Sub Total for")
            cur = [w["top"] for w in pg.extract_words(extra_attrs=["size"])
                   if w["text"] == "VND" and w["size"] < 7.5] if tail else []
            if tail and cur:
                lo, hi = min(cur) - 3, tail[0]["top"]
                drop |= {(round(c["x0"], 2), round(c["top"], 2), c["text"]) for c in chars
                         if c.get("size", 0) >= 7.9 and lo < c["top"] < hi - 1}
            # YUSEN：貨物說明 (Helvetica 8 號字) 太長，溢出到費用表 (LucidaConsole 9 號字)，
            # 跟費用行疊在一起 (YUSEN_DEC.2025 SSGN0300089 "CTEARRMGIONEASL R EHMANADINLEIDN OG")。
            # 費用表標題 "DESCRIPTION … CHARGES IN" 到 "Continued Over" / "TOTAL CHARGES" 之間，
            # 只留 LucidaConsole 的字
            drop |= _cw_overflow_chars(pg, chars)
            keep ={(round(c["x0"], 2), round(c["top"], 2), c["text"]) for c in chars} - drop

            def _keep(obj, keep=keep):
                if obj.get("object_type") != "char":
                    return True
                return (round(obj["x0"], 2), round(obj["top"], 2), obj["text"]) in keep

            clean = pg.filter(_keep)
            text = clean.extract_text(y_tolerance=1) or ""
            words = clean.extract_words()
            hdr = next((w for w in words if w["text"] == "SHIPPER"), None)
            cons = next((w for w in words if hdr and w["text"] == "CONSIGNEE"
                         and abs(w["top"] - hdr["top"]) < 2), None)
            below = [w for w in words if hdr and w["top"] > hdr["top"] + 4]
            if hdr and cons and below:
                first = min(w["top"] for w in below)
                line = sorted((w for w in below if abs(w["top"] - first) < 2
                               and w["x0"] < cons["x0"] - 1), key=lambda w: w["x0"])
                if line:
                    text += "\n##SHIPPER## " + " ".join(w["text"] for w in line)
            if "SOA #" in text and "SWB" in text:
                text += _dsv_vn_grid_lines(clean)
            if ("TOTAL CHARGES" in text and not _CW_CHARGE_RE.search(text)
                    and not re.search(r"SUBTOTAL\s*:?\s*\d", text)):
                # 費用表的字型沒有文字編碼 (畫面看得到，文字層沒有)：整頁改用 OCR
                # (2026-10-28 抽樣 DHL_MAY.2026 的 PRO FORMA，例 S2602151382)
                text = _ocr_page_with_psm(pdf_path, page_index, 6)
            pages.append(text + _CW_PAGE_MARK)
    return pages


def _cw_ocr_shipment(pg) -> Optional[str]:
    """掃描頁右上角表頭的 SHIPMENT 號碼，單獨放大重新 OCR。整頁 OCR 常讀錯
    或漏掉這一格 (2026-10-29 YUSEN_SEP.2026：SSGN0315898 讀成 15596/15808、
    SSGN0315454 整格漏掉)。讀不到回傳 None，照用整頁 OCR 的結果。"""
    try:
        import pytesseract
        image = pg.to_image(resolution=300).original.convert("L")
        w, h = image.size
        # 表頭表格在右上角 (x 55%~100%、y 12%~26%)；只截這一塊，psm 6
        head = image.crop((int(w * 0.55), int(h * 0.12), w, int(h * 0.26)))
        text = pytesseract.image_to_string(head, config="--psm 6")
    except Exception:
        return None
    m = re.search(r"SHIPMENT\W*([A-Z]{3,5}?)(O*)([0-9O]{6,8})\b", text)
    if not m:
        return None
    prefix, extra, num = m.groups()
    num = num.replace("O", "0")
    if len(num) == 6 and extra:          # "SSGNO315640"：O 就是開頭的 0
        num = "0" + num
    return prefix + num if re.fullmatch(r"\d{7}", num) else None


def _cw_overflow_chars(pg, chars) -> set:
    """費用表範圍內、不是 LucidaConsole 費用字型的字 (溢出的貨物說明)。
    頁面沒有 LucidaConsole 費用字時不刪 (其他 CargoWise 版面不受影響)。"""
    words = pg.extract_words(extra_attrs=["fontname"])
    hdr = next((w for w in words if w["text"] == "DESCRIPTION" and "Bold" in w["fontname"]
                and any(v["text"] == "CHARGES" and v["x0"] > 400 and abs(v["top"] - w["top"]) < 2
                        for v in words)), None)
    if not hdr:
        return set()
    ends = [w["top"] for w in words if w["top"] > hdr["top"]
            and w["text"] in ("Continued", "TOTAL") and w["x0"] < 300]
    lo, hi = hdr["bottom"], min(ends) if ends else pg.height
    if not any("Lucida" in c.get("fontname", "") and lo < c["top"] < hi for c in chars):
        return set()
    return {(round(c["x0"], 2), round(c["top"], 2), c["text"]) for c in chars
            if lo < c["top"] < hi - 1 and "Lucida" not in c.get("fontname", "")
            and "Bold" not in c.get("fontname", "") and c.get("size", 0) < 8.5}


_CW_DATE = r"\d{1,2}-(?:[A-Za-z]{3}|\d{2})-\d{2}"
_CW_CHARGE_RE = re.compile(
    # VAT 欄也可能是 "Zero Rated" (國際運費零稅率，MOL 2026-01)
    # 美金發票的金額用逗號當小數點 ("8%=2,36 29,50"，YUSEN 2026-03)
    # DHL 免稅行 "Freight USD 500,00 @ 26386,000000 0% 1.13.193.000"：金額前的 "1." 是右下
    # "VAT RATE REASON" 表的原因編號 (1. VAT EXEMPT INTERNATIONAL TRANSPORT)，不是金額
    # (2026-10-28 抽樣 DHL_MAY.2026 S2601691598)
    r"^(?P<desc>.+?)(?: - (?:Greater|Base|Min|\d)[^\n]*?)?\s+(?:(?P<rate>\d+)%=(?P<vat>[\d.]+(?:,\d+)?)|Zero Rated"
    # OCR 頁的原因編號可能讀成 "*" 或 "1." 後面多一個空白 ("0% * 46.695.655")
    r"|0%\s+(?:\d\.\s?|\*\s)(?=\d{1,3}(?:\.\d{3})*(?:,\d{1,2})?$))\s*"
    r"(?P<amount>\d{1,3}(?:\.\d{3})*(?:,\d{1,2})?)$", re.MULTILINE)


def _cw_vn_number(text: Optional[str]):
    """'1.234.567' (越南盾) -> 1234567；'29,50' / '1.234,50' (外幣，逗號是小數點) -> 29.5。"""
    if not text:
        return None
    if "," in text:
        return float(text.replace(".", "").replace(",", "."))
    return _dsv_vn_dot_amount(text)
# 數量可能用逗號當小數點 ("18,8 - dỡ"：名稱被數量切開，"-" 後面不是單位而是名稱的後半)
_CW_MOL_CHARGE_RE = re.compile(
    r"^(?P<desc>.+?)\s+(?P<qty>\d+(?:[.,]\d+)?)\s*-\s+(?P<unit>\S+)\s+(?P<rate>[\d,]+(?:\.\d+)?)\s+"
    r"(?P<pct>\d+)\s*%\s+(?P<cur>[A-Z]{3})\s+(?P<gross>[\d,]+(?:\.\d+)?)$", re.MULTILINE)
_CW_MOL_NOQTY_RE = re.compile(
    r"^(?P<desc>\D.+?)\s+(?P<rate>[\d,]+(?:\.\d+)?)\s+(?P<pct>\d+)\s*%\s+(?P<cur>[A-Z]{3})\s+"
    r"(?P<gross>[\d,]+(?:\.\d+)?)$")


def _cw_date(value: Optional[str]) -> Optional[str]:
    """'28-Sep-26' 或 '29-09-26' (DHL) -> '28 Sep 2026'。"""
    if not value:
        return None
    m = re.fullmatch(r"(\d{1,2})-(\d{2})-(\d{2})", value)
    if m:
        return _format_date_en(m.group(1), m.group(2), f"20{m.group(3)}")
    return _dsv_parse_date(value)


def _cw_comma_number(text: str):
    """英式數字 '1,000,000' / '12.75' -> 數字。"""
    value = float(text.replace(",", ""))
    return int(value) if value.is_integer() else value


# 掃描頁 OCR 會把標籤後的冒號讀成 ")" "|" "!" "/" ("SHIPMENT|/SSGN0316985")
_CW_INV_RE = r"SHIPMENT[\s)|:\]!/]*(?!DETAILS)([A-Z][A-Z0-9]{5,})"
_CW_PAGE_MARK = "\n##PAGE##"


def _cw_invoice_no(text: str) -> Optional[str]:
    """SHIPMENT 號碼；掃描頁 OCR 會在字母跟數字中間多讀一個 O ("SSGNO0315882")。"""
    inv = _search(_CW_INV_RE, text)
    if inv:
        fixed = re.sub(r"^([A-Z]+?)O+(?=\d{7,}$)", r"\1", inv)
        inv = fixed if fixed != inv and re.fullmatch(r"[A-Z]{3,5}\d{7,}", fixed) else inv
    return inv


def _cw_docs(text: str) -> List[str]:
    """一份 PDF 有好幾張發票時切開：每頁抬頭都有 SHIPMENT 號碼，連續幾頁號碼
    相同 (或續頁讀不到號碼) 就是同一張；"Page 1 of N" 一定是新的一張 (同一個
    SHIPMENT 可能開兩張，例如 YUSEN SSGN0316461 另有一張貨櫃延滯費)。沒有號碼
    也沒有頁碼的頁 (疊印的附件頁) 不算。沒有分頁記號 (非 _cargowise_pages
    擷取) 時整份當一張。先修 OCR 的數字空白 ("8%=64 . 000"、"1100000 ,00")。"""
    text = re.sub(r"(?<=\d) *([.,]) *(?=\d)", r"\1", text)
    pages = [p for p in text.split(_CW_PAGE_MARK) if p.strip()]
    if len(pages) <= 1:
        return [text.replace(_CW_PAGE_MARK, "")]
    docs: List[List[str]] = []
    current = None
    for page in pages:
        inv = _cw_invoice_no(page)
        first_page = re.search(r"\bPage 1 of \d+", page)
        if inv is None and not re.search(r"\bPage \d+ of \d+", page):
            continue
        # "Page 2 of N" 以後一定是上一張的續頁 (掃描頁 OCR 號碼可能讀錯，不能只靠號碼)
        later_page = re.search(r"\bPage (?:[2-9]|\d{2,}) of \d+", page)
        if docs and not first_page and (inv is None or inv == current or later_page):
            docs[-1].append(page)
        else:
            docs.append([page])
            current = inv
    # 同一張單據掃描兩次 (號碼與 TOTAL 都相同) 只留一張 (使用者 2026-10-28：DHL 合輯第 19、20 頁
    # S2601526262)。同號碼但總額不同的是另一張發票 (YUSEN SSGN0316461)，照常保留
    out, seen = [], set()
    for d in ("\n".join(d) for d in docs):
        total = re.findall(r"(?m)^[\s|\"'`]*TOTAL\b[^\n]*?(\d[\d.,]*\d)\s*$", d)
        key = (_cw_invoice_no(d), total[-1] if total else None)
        if key[0] and key[1] and key in seen:
            continue
        seen.add(key)
        out.append(d)
    return out or [text]


def _cw_after(text: str, header_re: str, line_re: str, look: int = 4):
    """在 header_re 那一行之後幾行內，找第一行符合 line_re 的。"""
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        if re.search(header_re, ln):
            for nxt in lines[i + 1:i + 1 + look]:
                m = re.match(line_re, nxt.strip())
                if m:
                    return m
            return None
    return None


def _cw_parse_doc(doc: str, supplier: Optional[str]) -> List[Dict]:
    h: Dict[str, Optional[str]] = {"supplier": supplier}
    h["invoice_no"] = _cw_invoice_no(doc)
    # 掃描頁 OCR 會在標籤後多讀一個符號 ("INVOICE DATE) 19-Apr-26"，2026-10-28 抽樣 DHL 合輯)
    h["invoice_date"] = _cw_date(_search(r"\bDATE\s*[\[|()/]?\s*(\d{1,2}-[A-Za-z]{3}-\d{2})", doc))
    m = re.search(r"^(?!SHIPMENT)(.+?)\s*(?:INVOICE DATE|SHIPMENT\s*[A-Z][A-Z0-9]{5,})",
                  doc, re.MULTILINE)
    h["consignee"] = m.group(1).strip() if m else None
    h["shipper"] = _search(r"^##SHIPPER## (.+)$", doc)
    h["order_no"] = _search(r"\bPON:\s*(\S+)", doc)
    if not h["order_no"]:
        m = _cw_after(doc, r"OWNER'S REFERENCE", r"(\S*\d\S*)", look=1)
        h["order_no"] = m.group(1) if m else None
    vol = _search(r"KG\s+([\d.,]+)\s*M3", doc)
    kg = _search(r"([\d.,]+)\s*KG\s+[\d.,]+\s*M3", doc)
    # 材積統一寫法 (見 _fmt_volume)：CONTAINERS 區塊的櫃型 ("FFAU3490244 - 40HC")
    # 算櫃量；MOL 沒有這個區塊，用費用單位 "4 - 40HC" 的最大數量；都沒有寫 CBM
    boxes = [(1, k) for k in re.findall(r"^\s*[A-Z]{4}\d{7}\s*-\s*(\w+)", doc, re.MULTILINE)]
    if not boxes:
        per_type: Dict[str, int] = {}
        for qty, kind in re.findall(r"\s(\d+)\s+-\s+(\d{2}[A-Z]{2})\s", doc):
            per_type[kind] = max(per_type.get(kind, 0), int(qty))
        boxes = [(n, k) for k, n in per_type.items()]
    h["volume"] = _fmt_volume(containers=boxes, cbm=vol.replace(",", ".") if vol else None,
                              kg=kg.replace(",", ".") if kg else None)
    h["vessel"] = h["bl_no"] = h["mbl_no"] = None
    m = _cw_after(doc, r"OCEAN BILL OF LADING.*HOUSE BILL OF LADING",
                  r"(?P<v>.+?) / (?P<voy>\S+) / \S+ (?P<ocean>\S+) (?P<house>\S+)$")
    if m:
        h["vessel"] = f"{m.group('v')} / {m.group('voy')}"
        h["mbl_no"], h["bl_no"] = m.group("ocean"), m.group("house")
    else:
        m = _cw_after(doc, r"FLIGHT / DATE & REFERENCE\s+MAWB\s+HAWB",
                      r"(?P<flight>\S+) / .*?(?P<mawb>\S+) (?P<hawb>\S+)$")
        if m:
            h["vessel"], h["mbl_no"], h["bl_no"] = m.group("flight", "mawb", "hawb")
    m = _cw_after(doc, r"^ORIGIN\b",
                  rf"(?P<pol>.+?)\s+(?P<etd>{_CW_DATE})\s+(?P<pod>.+?)(?:\s+(?P<eta>{_CW_DATE}))?$")
    h["port_of_loading"] = m.group("pol").strip(" ,") if m else None
    h["port_of_discharge"] = m.group("pod").strip(" ,") if m else None
    h["onboard_date"] = _cw_date(m.group("etd")) if m else None
    h["arrive_date"] = _cw_date(m.group("eta")) if m else None
    sec = re.search(r"^CONTAINERS\s*\n(.*?)(?:\n\s*(?:Total CO|CO₂|CHARGES)|\Z)",
                    doc, re.DOTALL | re.MULTILINE)
    conts = [ln.strip() for ln in sec.group(1).split("\n")
             if re.match(r"\s*[A-Z]{4}\d{7}\b", ln)] if sec else []
    h["container_no"] = ", ".join(conts) or None

    cur = _search(r"CHARGES IN ([A-Z]{3})", doc)
    suffix = f" ({cur})" if cur and cur != "VND" else ""      # 外幣發票的費用名稱加幣別
    items = [{"description": im.group("desc").strip() + suffix,
              "amount": _cw_vn_number(im.group("amount")),
              "vat": _cw_vn_number(im.group("vat"))}
             for im in _CW_CHARGE_RE.finditer(doc)]
    doc_lines = doc.splitlines()
    mol_matches = []
    for i, line in enumerate(doc_lines):
        im = _CW_MOL_CHARGE_RE.match(line)
        if im:
            mol_matches.append((im, im.group("qty").replace(",", ".")))
            continue
        # 數量印在上一行 ("14,48 - M3" / "Phí khai thác hàng lẻ (CFS) 282,000 8 % VND 4,410,029"，
        # 2026-10-28 抽樣；數量的逗號是小數點)
        im = _CW_MOL_NOQTY_RE.match(line)
        prev = re.match(r"^\s*(\d+(?:,\d+)?)\s+-\s+\S+\s*$", doc_lines[i - 1]) if i else None
        if im and prev:
            mol_matches.append((im, prev.group(1).replace(",", ".")))
    for im, qty in mol_matches:
        # RATE 有時是單價 (數量 × 單價)，有時已經是金額 ("3 - Shipment 3,000,000 8 % 3,240,000")：
        # 選「× (1+稅率) 等於含稅金額」的那一個，都不對就用含稅金額回推
        gross, pct = _cw_comma_number(im.group("gross")), int(im.group("pct"))
        rate = _cw_comma_number(im.group("rate"))
        net = next((c for c in (round(float(qty) * rate, 2), rate)
                    if abs(c * (1 + pct / 100) - gross) <= 2), round(gross / (1 + pct / 100)))
        vat = round(gross - net, 2)
        desc = im.group("desc").strip()
        unit = im.groupdict().get("unit") or ""
        if unit and (not unit.isascii() or "(" in unit or unit.islower()):
            desc = f"{desc} {unit}"          # "Phí xếp" + "dỡ"、"Phí khai thác hàng lẻ" + "(CFS)"
        if im.group("cur") != "VND":
            desc += f" ({im.group('cur')})"
        items.append({"description": desc,
                      "amount": int(net) if float(net).is_integer() else net,
                      "vat": int(vat) if float(vat).is_integer() else vat})
    if not items:
        items = [{"description": None, "amount": None, "vat": None}]
    return [{**h, **it} for it in items]


def _cw_parse(text: str, supplier: Optional[str]) -> Dict:
    items = [row for doc in _cw_docs(text) for row in _cw_parse_doc(doc, supplier)]
    return {"header": {}, "items": items}


def _cw_totals(text: str, patterns, conv=_cw_vn_number):
    """每張發票的總額 {SHIPMENT 號碼: 金額}。patterns 依序試，第一個抓得到的
    為準 (同一個 pattern 抓到好幾個就加總，例如 MOL 每個幣別一行 GRAND TOTAL)。"""
    out, missing = {}, set()
    for doc in _cw_docs(text):
        inv = _cw_invoice_no(doc)
        for pattern in patterns:
            vals = re.findall(pattern, doc, re.MULTILINE)
            if inv and vals:
                # 同一個 SHIPMENT 開兩張時，總額相加 (明細也會併在同一組勾稽)
                out[inv] = out.get(inv, 0) + sum(conv(v) for v in vals)
                break
        else:
            missing.add(inv)
    # 同號碼的另一張讀不到總額 (掃描頁 OCR)：只有一半的總額，不勾稽這個號碼
    return {k: v for k, v in out.items() if k not in missing}


def _cw_extract_total(text: str):
    # 掃描頁 OCR 可能漏掉 "SUBTOTAL" 標籤，金額仍接在 "discrepancies." 後面 (YUSEN_SEP.2026)
    return _cw_totals(text, [r"SUBTOTAL[ \t]+(\d[\d.]*(?:,\d+)?)",
                             r"discrepancies\.[ \t]+(\d{1,3}(?:\.\d{3})+(?:,\d+)?)[ \t]*$"])


def _cw_extract_grand_total(text: str):
    # YUSEN 的金額有時印在 "TOTAL VND" 上一行 (字的高度略高，被拆成兩行)
    # 外幣發票是 "TOTAL USD" (金額 "31,86" 逗號是小數點)
    return (_cw_totals(text, [r"^TOTAL[ \t]+(?:VND|USD)[ \t]+(\d[\d.]*(?:,\d+)?)",
                              r"^(\d[\d.]*(?:,\d+)?)[ \t]*\nTOTAL[ \t]+(?:VND|USD)\s*$",
                              # 掃描頁 OCR 漏掉 "TOTAL VND" 標籤時，用頁尾 "Invoiced VND 7.171.200"
                              r"Invoiced[ \t]+(?:VND|USD)[ \t]+(\d{1,3}(?:\.\d{3})+(?:,\d+)?)\b"])
            or _cw_totals(text, [r"GRAND TOTAL for [A-Z]{3}[ \t]+(\d[\d,]*(?:\.\d+)?)"],
                          conv=_cw_comma_number))


register_supplier(
    "DSV_VN", "CÔNG TY TNHH DSV AIR & SEA", detect_dsv_vn, parse_dsv_vn,
    extract_total=_dsv_vn_extract_total, extract_grand_total=_dsv_vn_extract_grand_total,
    page_text=_cargowise_pages,
    # PRO FORMA 合輯裡夾的 VNPT 電子 VAT 發票頁是附件，不擷取 (費用與 PRO FORMA 重複；
    # 使用者 2026-10-28，JUL.2026_BY SHIPMENT 第 176 頁)。整份只有電子發票時一頁都
    # 不符合，invoice_text() 會改用全部頁面，照常擷取
    invoice_page=lambda page: not _is_dsv_vn_einvoice(page),
)


# VN13 DHL：PRO FORMA INVOICE，日期 "29-09-26"、港口只有代碼、訂單號碼取 "PON:"
def detect_dhl_vn(text: str) -> bool:
    return "DHL GLOBAL FORWARDING (VIETNAM)" in text.upper()


# 到貨通知 (Arrival Notice，2026-10-28 抽樣) 附的「Air/Sea Freight Charge Sheet」：
#   "Freight HKD 940,00 3.408,120000 0,00 3.203.633,00" (名稱、幣別、金額、匯率、VAT、折合越南盾含稅)
#   "Dest. Document VND 500.000,00 40.000,00 540.000,00" (越南盾沒有匯率)；名稱可能折到下一行。
#   金額 = 越南盾含稅 − VAT；稽核 Total (VND)。發票號碼 = INVOICE NO. (沒有用 SHIPMENT 號)。
_DHL_VN_CHARGE_RE = re.compile(
    r"^(?P<desc>[A-Za-z].*?)\s+(?P<cur>[A-Z]{3})\s+(?P<amt>[\d.]+,\d{2})\s+(?:[\d.]+,\d{3,}\s+)?"
    r"(?P<vat>[\d.]+,\d{2})\s+(?P<vnd>[\d.]+,\d{2})$")


def _dhl_vn_num(text: str) -> float:
    """'3.203.633,00' (句點千分位、逗號小數) -> 3203633.0。"""
    return float(text.replace(".", "").replace(",", "."))


def _is_dhl_vn_charge_sheet(text: str) -> bool:
    return "Freight Charge Sheet" in text and "Total (VND)" in text


def _parse_dhl_vn_charge_sheet(text: str) -> Dict:
    h: Dict[str, Optional[str]] = {"supplier": "DHL GLOBAL FORWARDING (VIETNAM) CORPORATION"}
    h["invoice_no"] = _search(r"INVOICE NO\.:\s*(\S+)", text) or _search(r"SHIPMENT\s*(S\d{8,})", text)
    m = re.search(r"DATE\s*(\d{1,2})-([A-Za-z]{3})-(\d{2})\b", text)
    h["invoice_date"] = _format_date_en_from_month_name(m.group(1), m.group(2), "20" + m.group(3)) if m else None
    h["consignee"] = _search(r"WILL BE INVOICED TO:\s*(.+?)\s*(?:\(|$)", text)
    h["bl_no"] = _search(r"HAWB\s*([A-Z0-9]{5,})", text)
    h["order_no"] = _search(r"SHIPMENT\s*(S\d{8,})", text)
    lines = [ln.strip() for ln in text.splitlines()]
    start = next((i for i, ln in enumerate(lines) if ln.startswith("CHARGE DESCRIPTION")), len(lines))
    items = []
    for i in range(start + 1, len(lines)):
        if lines[i].startswith("Total (VND)"):
            break
        m = _DHL_VN_CHARGE_RE.match(lines[i])
        if not m:
            continue
        desc = m.group("desc").strip()
        nxt = lines[i + 1] if i + 1 < len(lines) else ""
        # 名稱折到下一行只會是單一個字 ("Dest. Document" / "Handover")；"Filing Fee - AFR"
        # 這種多個字的行是另一個沒有金額的項目，不接
        if re.fullmatch(r"[A-Za-z][\w\-]*", nxt):
            desc = f"{desc} {nxt}"
        vat, vnd = _dhl_vn_num(m.group("vat")), _dhl_vn_num(m.group("vnd"))
        items.append({"description": desc, "amount": round(vnd - vat, 2), "vat": vat or None})
    return {"header": h, "items": items or [{"description": None, "amount": None, "vat": None}]}


def parse_dhl_vn(text: str) -> Dict:
    if _is_dhl_vn_charge_sheet(text):
        return _parse_dhl_vn_charge_sheet(text)
    return _cw_parse(text, _search(r"(DHL Global Forwarding \(Vietnam\) Corporation)", text))


def _dhl_vn_grand(text: str):
    if _is_dhl_vn_charge_sheet(text):
        v = _search(r"Total \(VND\)\s*([\d.]+,\d{2})", text)
        return _dhl_vn_num(v) if v else None
    return _cw_extract_grand_total(text)


register_supplier(
    "DHL_VN", "DHL GLOBAL FORWARDING (VIETNAM) CORPORATION", detect_dhl_vn, parse_dhl_vn,
    extract_total=lambda t: None if _is_dhl_vn_charge_sheet(t) else _cw_extract_total(t),
    extract_grand_total=_dhl_vn_grand, page_text=_cargowise_pages,
)


# VN20 YUSEN：PRO FORMA TAX INVOICE，有 DRAFT 浮水印
def detect_yusen(text: str) -> bool:
    return "YUSEN LOGISTICS (VIETNAM)" in text.upper()


def parse_yusen(text: str) -> Dict:
    supplier = _search(r"(YUSEN LOGISTICS \(VIETNAM\) CO\., LTD\.[^\n]*?BRANCH)", text)
    if not supplier:
        # 2025 上半年舊版：名稱分兩行 "YUSEN LOGISTICS (VIETNAM) CO.LTD" / (Bank) / "HO CHI MINH BRANCH,…"。
        # 統一成新版寫法 (使用者 2026-10-29)
        m = re.search(r"YUSEN LOGISTICS \(VIETNAM\) CO\.\s*,?\s*LTD\.?\s*\n(?:Bank\s*\n)?\s*"
                      r"HO CHI MINH BRANCH", text)
        supplier = "YUSEN LOGISTICS (VIETNAM) CO., LTD. - HO CHI MINH BRANCH" if m else None
    return _cw_parse(text, supplier)


register_supplier(
    "YUSEN", "YUSEN LOGISTICS (VIETNAM) CO., LTD.", detect_yusen, parse_yusen,
    extract_total=_cw_extract_total, extract_grand_total=_cw_extract_grand_total,
    page_text=_cargowise_pages,
)


# VN09 MOL：Debit Note，公司名稱只在 logo 圖片裡 (文字層沒有)，固定填
# "MOL Logistics Vietnam Inc" (使用者 2026-10 正確答案)；House B/L 是 "MLG…"
def detect_mol(text: str) -> bool:
    return bool(re.match(r"\s*Debit Note - ", text)
                and (re.search(r"\bMLG[A-Z]{2}\d{6,8}\b", text) or "MOL LOGISTICS" in text.upper())) \
        or _is_molcs(text)


# --- MOL CONSOLIDATION SERVICE 越南增值稅發票 (2026-10-28 抽樣 "MOL LOGICTICS_SEP.2025" 等，63 頁) --
# 一頁一張 HÓA ĐƠN GIÁ TRỊ GIA TĂNG；這批 "Số (No)" 與日期都空白 (草稿)，發票號碼暫用備註裡的
# 訂艙號碼 (VN2090087)，有填號碼時用號碼 (待使用者確認)。明細 "1 PHÍ CHỨNG TỪ (FCR fee) Bill 1.00
# 1,000,000 1,000,000"，名稱太長時折到上下兩行；整張一個 VAT 單獨一列。
_MOLCS_ITEM_RE = re.compile(
    r"^\d+\s+(?:(?P<desc>.+?)\s+)?(?P<unit>[A-Za-z]{2,6})\s+(?P<qty>[\d.]+)\s+[\d,]+\s+(?P<amount>[\d,]+)$")


def _is_molcs(text: str) -> bool:
    upper = text.upper()
    return "MOL CONSOLIDATION SERVICE" in upper and "HÓA ĐƠN GIÁ TRỊ GIA TĂNG" in upper


def _molcs_pages(text: str) -> List[str]:
    return [p for p in text.split(_CW_PAGE_MARK) if _is_molcs(p) and "Cộng tiền hàng" in p]


def _molcs_num(value: Optional[str]) -> Optional[int]:
    digits = (value or "").replace(",", "")
    return int(digits) if digits.isdigit() else None


def _molcs_parse_page(page: str) -> List[Dict]:
    lines = [ln.strip() for ln in page.splitlines()]
    start = next((i for i, ln in enumerate(lines) if ln.startswith("A B C")), 0)
    stop = next((i for i, ln in enumerate(lines) if ln.startswith("Cộng tiền hàng")), len(lines))
    items, last = [], start
    for i in range(start + 1, stop):
        m = _MOLCS_ITEM_RE.match(lines[i])
        if not m:
            continue
        desc = m.group("desc")
        if not desc:
            before = lines[i - 1] if i - 1 > last else ""
            after = lines[i + 1] if i + 1 < stop and not _MOLCS_ITEM_RE.match(lines[i + 1]) else ""
            desc = f"{before} {after}".strip()
        items.append({"description": desc, "amount": _molcs_num(m.group("amount")), "vat": None,
                      "unit": m.group("unit"), "qty": m.group("qty")})
        last = i + (0 if m.group("desc") else 1)
    remark = " ".join(lines[last + 1:stop])
    no = _search(r"Số \(No\):\s*(\d+)", page)
    ref = _search(r"\b(VN\d{7})\b", remark)
    d = re.search(r"Ngày \(Date\)\s*(\d{1,2})\s*tháng \(month\)\s*(\d{1,2})\s*năm \(year\)\s*(\d{4})", page)
    # 海運 "(Tàu MAERSK NACKA,chuyến số 534S, HO CHI MINH-UBERHERRN)"、
    # 空運 "(Chuyến bay LH8021, HO CHI MINH- UBERHERRN)"
    ship = (re.search(r"\(Tàu\s+(.+?),\s*chuyến số\s+(\S+?),\s*(.+?)-\s*(.+?)\)", remark)
            or re.search(r"\(Chuyến\s+bay\s+(\S+?)()\s*,\s*(.+?)-\s*(.+?)\)", remark))
    cbm = next((it["qty"] for it in items if it["unit"].upper() == "CBM"), None)
    rate = _search(r"VAT rate\):\s*(\d+)\s*%", page)
    common = {
        "invoice_no": no or ref, "invoice_date": _format_date_en(*d.groups()) if d else None,
        "supplier": "CÔNG TY TNHH MOL CONSOLIDATION SERVICE (VIỆT NAM)",
        "consignee": _search(r"Tên đơn vị \(Company's name\):\s*(.+)", page),
        "bl_no": ref,
        "vessel": (f"{ship.group(1)} / {ship.group(2)}" if ship.group(2) else ship.group(1)) if ship else None,
        "port_of_loading": ship.group(3).strip() if ship else None,
        "port_of_discharge": ship.group(4).strip() if ship else None,
        "volume": f"{cbm} M3" if cbm else None,
    }
    rows = [{**common, "description": it["description"], "amount": it["amount"], "vat": None} for it in items]
    vat = _molcs_num(_search(r"VAT amount\):\s*([\d,]+)", page))
    if rows and vat is not None:
        rows.append({**common, "description": f"Thuế GTGT {rate}%" if rate else "Thuế GTGT", "amount": None,
                     "vat": vat})
    return rows


def parse_mol(text: str) -> Dict:
    if _is_molcs(text):
        rows = [r for p in _molcs_pages(text) for r in _molcs_parse_page(p)]
        return {"header": {}, "items": rows or [{"description": None, "amount": None}]}
    return _cw_parse(text, "MOL Logistics Vietnam Inc")


def _mol_totals(text: str, label: str):
    out = {}
    for p in _molcs_pages(text):
        rows = _molcs_parse_page(p)
        if rows:
            out[rows[0]["invoice_no"]] = _molcs_num(_search(label + r"\):\s*([\d,]+)", p))
    return out


register_supplier(
    "MOL", "MOL LOGISTICS VIETNAM INC", detect_mol, parse_mol,
    extract_total=lambda t: _mol_totals(t, r"Total before VAT") if _is_molcs(t) else None,
    extract_grand_total=lambda t: _mol_totals(t, r"Total amount") if _is_molcs(t) else _cw_extract_grand_total(t),
    page_text=_cargowise_pages,
)


# ===========================================================================
# 供應商 VN15：EXPEDITORS VIETNAM COMPANY LIMITED (一頁一張發票，multi_page)
# ===========================================================================
#
# 版面 (使用者 2026-10 樣本 SEP.2026 Van Thuan，58 頁)：
#   "CLIENT NO: G0201863 DATE 09/28/26" (美式 MM/DD/YY)、
#   "… TAX No: NUMBER E3Z3109683" (INVOICE NO)、下一行起是客戶名稱、
#   空運 "AWB/BL: 157-03039422 DESTINATION: JFK"、重量 "161 PCS 2082.000 K"；
#   海運 "CARRIER: PENANG BRIDGE"、"120N 48 PCS 5.120 CBM"。
#   費用 "THC 8 5,915,875" (名稱 + VAT 稅率 + 金額)；"BAS 50.00 / SET" 這類
#   計費說明行不是費用。VAT 是整張一列：Description "VAT 8%"、VAT 欄填 "VAT:"。
#   稽核：SUB TOTAL = 費用加總；INVOICE TOTAL = 費用 + VAT。

def detect_expeditors(text: str) -> bool:
    return "EXPEDITORS VIETNAM COMPANY LIMITED" in text.upper() and "TAX No:" in text


# 稅別欄可能是字母 "E" (免稅，空運費 "AIR FREIGHT E 712.50"，2026-10-28 抽樣)
_EXP_ITEM_RE = re.compile(r"^(?P<desc>.+?)\s+(?P<rate>\d{1,2}|E)\s+(?P<amount>\d[\d,]*(?:\.\d+)?)$")


def parse_expeditors(text: str) -> Dict:
    h: Dict[str, Optional[str]] = {code: None for code in (
        "order_no", "arrive_date", "onboard_date", "mbl_no", "port_of_loading",
        "container_no", "shipper")}
    h["supplier"] = _search(r"(Expeditors Vietnam Company Limited)", text)
    h["invoice_no"] = _search(r"TAX No:\s*NUMBER\s+(\S+)", text)
    m = re.search(r"CLIENT NO:\s*\S+\s+DATE\s+(\d{2})/(\d{2})/(\d{2})", text)
    h["invoice_date"] = _format_date_en(m.group(2), m.group(1), f"20{m.group(3)}") if m else None
    # 中間可能夾空行 (樣本 E3Z3101331)
    h["consignee"] = _search(r"CLIENT NO:[^\n]*\n\s*(.+?)\s+INVOICE\s*$", text)
    h["bl_no"] = _search(r"AWB/BL:\s*(\S+)", text)
    h["port_of_discharge"] = _search(r"DESTINATION:\s*(\S+)", text)
    h["vessel"] = _search(r"CARRIER:\s*(.+?)\s*$", text)
    # 材積統一寫法 (見 _fmt_volume)：海運 "5.120 CBM" -> "5.120 M3"；空運只有
    # 重量 "2082.000 K" -> "2082 KG"
    weight = _search(r"PCS\s+([\d.]+)\s+K\b", text)
    h["volume"] = _fmt_volume(cbm=_search(r"PCS\s+([\d.]+)\s+CBM", text),
                              kg=f"{float(weight):g}" if weight else None)
    block = re.search(r"Description\s+VAT\s+AMOUNT\s+[A-Z]{3}\s*\n(.*?)\nSUB TOTAL", text, re.DOTALL)
    items = []
    for line in (block.group(1).split("\n") if block else []):
        im = _EXP_ITEM_RE.match(line.strip())
        if im and not line.strip().startswith("BAS "):
            items.append({"description": im.group("desc"),
                          "amount": _cw_comma_number(im.group("amount")), "vat": None})
    # VAT 以下方稅率表為準 ("8 Standard VAT (8%) USD 87.50 7 94.50")：上方 "VAT:"
    # 有時字型壞掉讀成 "Z" (樣本 E3Z3101671)
    # 一張發票跨兩頁時稅率表在第二頁：只在有 SUB TOTAL 的那一頁取 VAT，先用 "VAT:" (2026-10-28 抽樣)
    vat = None
    if "SUB TOTAL" in text:
        vat = (_search(r"^VAT:\s*(\d[\d,.]*)", text)
               or _search(r"Standard VAT \(\d+%\)\s+[A-Z]{3}\s+[\d,.]+\s+([\d,.]+)\s+[\d,.]+", text))
    rate = _search(r"Standard VAT \((\d+%)\)", text)
    if vat:
        items.append({"description": f"VAT {rate}" if rate else "VAT",
                      "amount": None, "vat": _cw_comma_number(vat)})
    return {"header": h, "items": items or [{"description": None, "amount": None, "vat": None}]}


def _exp_extract_total(text: str):
    v = _search(r"SUB TOTAL:\s*([\d,.]+)", text)
    return _cw_comma_number(v) if v else None


def _exp_extract_grand_total(text: str):
    v = _search(r"INVOICE TOTAL:\s*([\d,.]+)", text)
    return _cw_comma_number(v) if v else None


register_supplier(
    "EXPEDITORS", "EXPEDITORS VIETNAM COMPANY LIMITED", detect_expeditors, parse_expeditors,
    multi_page=True, extract_total=_exp_extract_total, extract_grand_total=_exp_extract_grand_total,
)


# ===========================================================================
# 越南 Excel 對帳表共用擷取 (VN01、VN02、VN06 … 2026-10-27 新增)
# ===========================================================================
#
# 越南供應商多半用 Excel 對帳表 (Bảng kê / Statement) 請款，每家欄位不同，可是
# 骨架一樣：抬頭 (賣方、買方、日期、附發票號碼)、一到兩列表頭 (例如「LOẠI CONT」
# 下面再分 20'/40'/45')、一列一筆出貨、費用欄、VAT 欄、合計欄、最後合計列。
# 這裡寫一套共用程式，每家只給一份設定 (_VT_xxx)：
#   - 欄位用表頭文字找 (去聲調、轉大寫後「包含」比對)，不寫死欄位位置。
#   - 兩列表頭自動合併：上一列是合併儲存格時，往左找最近的標題，
#     例如 "LOẠI CONT" + "40'" -> "LOAI CONT 40'"。
#   - 慣例同 VN05 CONG THANH 的運費對帳表 (使用者答案)：每一列出貨各一張發票，
#     INVOICE NO = 設定的單號欄；每個費用欄是一筆費用 (名稱用表頭)，VAT 單獨一筆。
#     同一個單號出現在好幾列 (或好幾張工作表) 時，合併成同一張發票勾稽。
#   - 勾稽 (sheet_totals)：每張發票的費用加總 = 「未稅合計」欄；加上 VAT =
#     「含稅合計」欄 (有設定才做)。
#   - 材積照統一寫法 (_fmt_volume)：20'/40'/45' 櫃數 -> "1X40"；沒有櫃數用 CBM；
#     再沒有用重量 KG。
#
# 設定欄位 (都用表頭文字的片段，去聲調大寫，例如 "SO CONTAINER")：
#   header：辨識表頭列的文字；key：發票號碼欄 (可給多個，依序取第一個有值的)；
#   fees：[(表頭片段, 費用名稱或 None=用表頭)]；vat / vat_desc：VAT 欄與名稱；
#   total / grand：未稅合計、含稅合計欄；date：每列日期欄；fields：{標準欄位:
#   表頭片段}；boxes：{"20": 表頭, "40": 表頭, "45": 表頭}；cbm / kg：材積欄；
#   port：「起點 - 終點」合在一欄時用來拆成啟運港/目的港。


def _vt_norm(value) -> str:
    """表頭比對用：去越南文聲調、轉大寫、換行與連續空白壓成一個空白。"""
    import unicodedata
    text = _vn_text(value).replace("Đ", "D").replace("đ", "d")
    text = "".join(ch for ch in unicodedata.normalize("NFKD", text)
                   if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", text).strip().upper()


def _vt_header_row(rows: List[list], marker) -> Optional[int]:
    """表頭列：第一個包含 marker (可給多個) 的列。"""
    markers = marker if isinstance(marker, (list, tuple)) else (marker,)
    for i, row in enumerate(rows[:40]):
        if any(m in _vt_norm(c) for c in row for m in markers):
            return i
    return None


def _vt_labels(rows: List[list], h: int, two_rows: bool = True, above: bool = False) -> List[str]:
    """每欄的完整標題：表頭列 (合併儲存格往左補) + 下一列子標題 (two_rows)。
    above：有些欄的標題在表頭的上一列 (例如 VN14 "NGÀY HĐ / SỐ HĐ")，表頭列空白時用它。"""
    top = rows[h]
    sub = (rows[h + 1] if h + 1 < len(rows) else []) if two_rows else []
    up = rows[h - 1] if above and h > 0 else []
    width = max(len(top), len(sub), len(up))
    labels, last = [], ""
    for c in range(width):
        t = _vt_norm(top[c]) if c < len(top) else ""
        if not t and c < len(up):
            t = _vt_norm(up[c])
        s = _vt_norm(sub[c]) if c < len(sub) else ""
        # 子標題是數字或很短 (20'/40'/VND/USD) 時，上一列空白代表被合併
        if t:
            last = t
        elif s:
            t = last
        labels.append(f"{t} {s}".strip())
    return labels


def _vt_col(labels: List[str], *needles) -> Optional[int]:
    """第一個包含任一片段的欄位。片段開頭加 "=" 代表要完全相等；再加 "$" ("$=TOTAL")
    代表取最後一個符合的欄。"""
    for needle in needles:
        if not needle:
            continue
        order = list(enumerate(labels))
        if needle.startswith("$"):
            needle, order = needle[1:], order[::-1]
        for c, label in order:
            if (label == needle[1:]) if needle.startswith("=") else (needle in label):
                return c
    return None


def _vt_doc_date(rows: List[list]) -> Optional[str]:
    """抬頭日期："Ngày 28 tháng 09 năm 2026"、"Ngày 03/09/2026"、"DATE : 28/09/2026"。"""
    for row in rows[:20]:
        # 去聲調後比對 (Excel 的越南文有時是組合字元，直接比對 "Ngày" 會失敗)
        text = _vt_norm(" ".join(_vn_text(c) for c in row if c is not None))
        m = re.search(r"NGAY\s+(\d{1,2})\s+THANG\s+(\d{1,2})\s+NAM\s+(\d{4})", text)
        if m:
            return _format_date_en(*m.groups())
        m = re.search(r"(?:NGAY|DATE)\s*:?\s*(\d{1,2})/(\d{1,2})/(\d{4})", text)
        if m:
            return _format_date_en(*m.groups())
        # "Date: 20 - APR- 2025" (VN19 IMX 舊檔)
        m = re.search(r"DATE\s*:\s*(\d{1,2})\s*-\s*([A-Z]{3})\s*-\s*(\d{4})", text)
        if m:
            return _format_date_en_from_month_name(*m.groups())
        # 只有請款月份 "Month of debit: March, 2025" (VN01)：用該月最後一天 (使用者 2026-10-28 確認)
        # (VN11 APL 卡車表 "Period JUN 2026 (From…)" 也一樣)
        m = re.search(r"(?:MONTH OF DEBIT|PERIOD)\s*:?\s*([A-Z]{3})[A-Z]*[\s,.]*(\d{4})", text)
        months = ("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
        if m and m.group(1) in months:
            import calendar
            month, year = months.index(m.group(1)) + 1, int(m.group(2))
            return _format_date_en(calendar.monthrange(year, month)[1], month, year)
    return None


def _vt_consignee(rows: List[list]) -> Optional[str]:
    """買方：「Khách hàng : …」「Tên khách hàng : …」「Họ tên người mua hàng (Customer): …」
    「To : …」「CONSIGNEE : …」「Customer : …」。"""
    pat = re.compile(r"^(?:TEN\s+)?KHACH\s+HANG|NGUOI\s+MUA\s+HANG|^TO\b|^CONSIGNEE|^CUSTOMER")
    for row in rows[:20]:
        # 只取儲存格第一行 (下一行常是地址)
        cells = [_vn_text(str(c).splitlines()[0]) for c in row if c not in (None, "") and str(c).strip()]
        for i, cell in enumerate(cells):
            if pat.search(_vt_norm(cell)):
                value = cell.split(":", 1)[1] if ":" in cell else ""
                if not value.strip(" :：") and i + 1 < len(cells):
                    value = cells[i + 1]        # "To :" | "POU YUEN …" 值在下一格
                # 同一格常接著地址/稅號 ("… VIỆT NAM Địa chỉ: …")
                value = re.split(r"\s+(?:Đ[ịi]a\s+ch[ỉi]|Add(?:ress)?\b|MST\b|Tel\b)", value, flags=re.I)[0]
                value = value.strip(" :：")
                if value:
                    return value
    return None


def _vt_loose_date(text: str, year: Optional[str]) -> Optional[str]:
    """沒有年份或年份兩位的日期：'04/08'、'14/8'、'AUG.07'、'12/08/26'、'20+21/8/2026' (取第一天)。"""
    m = re.match(r"(\d{1,2})(?:\+\d{1,2})?/(\d{1,2})(?:/(\d{2}|\d{4}))?$", text.strip())
    if m:
        day, month, yy = m.groups()
        yy = ("20" + yy if yy and len(yy) == 2 else yy) or year
        return _format_date_en(day, month, yy) if yy else None
    m = re.match(r"([A-Za-z]{3})\.?\s*(\d{1,2})$", text.strip())
    months = ("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
    if m and m.group(1).upper() in months and year:
        return _format_date_en(m.group(2), str(months.index(m.group(1).upper()) + 1), year)
    return None


def _vt_parse_sheet(name: str, rows: List[list], cfg: Dict, supplier: str,
                    year_hint: Optional[str] = None) -> Tuple[List[Dict], Dict]:
    """一張對帳表 -> (費用列, {發票號碼: [未稅合計, 含稅合計]})。"""
    h = _vt_header_row(rows, cfg["header"])
    if h is None:
        return [], {}
    two_rows = cfg.get("two_rows", True)
    if two_rows and h + 1 < len(rows) and any(
            isinstance(c, (datetime.date, datetime.datetime)) or re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", _vn_text(c))
            for c in rows[h + 1]):
        two_rows = False        # 表頭下一列就是資料 (有日期)：只有一列表頭 (VN18 POU LI T04-26)
    labels = _vt_labels(rows, h, two_rows, cfg.get("above", False))
    if cfg.get("inherit_blank"):
        # 合併儲存格佔兩欄、第二欄沒有標題 (VN14 CFS 金額欄)：沿用左邊欄的標題
        for c in range(1, len(labels)):
            if not labels[c] and labels[c - 1] and any(_vn_num(_vn_cell(r, c)) for r in rows[h + 1:h + 40]):
                labels[c] = labels[c - 1]

    def find(v):
        return _vt_col(labels, *(v if isinstance(v, (list, tuple)) else (v,)))

    col = {k: find(v) for k, v in cfg.items()
           if k in ("total", "total_add", "grand", "vat", "date", "cbm", "kg", "port", "group_col", "base_key")}
    keys = [_vt_col(labels, k) for k in cfg["key"]]
    # 發票號碼由好幾欄用逗號連起來 (VN02 代墊吊櫃費：SỐ HĐ NÂNG, SỐ HĐ HẠ，使用者 2026-10-27)
    join_keys = [_vt_col(labels, k) for k in cfg.get("key_join", ())]
    fees = [(find(f[0]), f[1]) for f in cfg["fees"]]
    # 費用自己的發票號碼欄 (第三個元素)：費用欄右邊第一個符合的欄。例如 VN18 空運費
    # 與 Local Charge 各有一張紅色發票 (兩欄都叫 "Red invoice no.")
    fee_keys = {}
    for f in cfg["fees"]:
        c = find(f[0])
        if len(f) > 2 and c is not None:
            right = _vt_col(labels[c:], f[2])
            fee_keys[c] = c + right if right is not None else None
    def fee_name(c):
        """費用名稱照抄表頭原文：子標題是費用名稱就用子標題 ("Description" 底下的
        "Ocean Freight"、"CHI HỘ" 底下的 "PHÍ NÂNG (gồm VAT)")；子標題只是幣別或
        櫃型 ("USD"、"VND"、"20'") 時用上面的標題 (合併儲存格往左找)。"""
        sub = _vn_text(_vn_cell(rows[h + 1], c)) if two_rows and h + 1 < len(rows) else ""
        if sub and not re.fullmatch(r"(?:VND|USD|RMB|\d{2}'?|VAT.*|TOTAL)", sub, re.I):
            return sub
        for left in range(c, -1, -1):
            top = _vn_text(_vn_cell(rows[h], left))
            if top:
                return top
        return labels[c]

    # 同一欄只算一次 ("FUMIGATION USD" 與 "$FUMIGATION USD" 在只有一欄時指到同一欄)
    fees = [(c, desc or fee_name(c)) for i, (c, desc) in enumerate(fees)
            if c is not None and c not in {x for x, _ in fees[:i]}]
    if cfg.get("fees_between"):
        # 兩欄之間的每一欄都是明細費用 (VN06 雜費表：Ship Mode 和 UNIT PRICE/TOTAL 之間的
        # MISCELL、OT.、APP、TRAN、CUT…，每個月不一樣)
        a = find(cfg["fees_between"][0])
        b = next((c for c in range((a or 0) + 1, len(labels))
                  if any(labels[c] == n[1:] if n.startswith("=") else labels[c].startswith(n)
                         for n in cfg["fees_between"][1])), None) if a is not None else None
        for c in range(a + 1, b) if a is not None and b is not None else ():
            if labels[c] and c not in {x for x, _ in fees}:
                fees.append((c, fee_name(c)))
    remainder = None
    if cfg.get("remainder"):
        # ("小計欄", "費用名稱")：小計減掉上面列出的明細費用，見下面逐列處理
        sub_col = find(cfg["remainder"][0])
        if sub_col is not None:
            # 小計右邊沒有標題的那一欄是 2% 附加費 (IMP 表沒寫標題，EXP 表寫 "2%")
            nxt = sub_col + 1
            has_after = nxt < len(labels) and labels[nxt] in ("", "2%")
            remainder = (sub_col, cfg["remainder"][1], [c for c, _ in fees], has_after)
            if has_after and nxt not in {c for c, _ in fees}:
                fees.append((nxt, cfg.get("after_sub", "2%")))
    if cfg.get("fees_re"):
        # 費用欄自動抓：標題符合 fees_re 的每一欄，group(1) 當費用名稱 (VN01 TỔNG 每個月
        # 的費用欄都不一樣：D/O、TCS、PSS、A/F、O/F…)；fees_skip 排除合計欄
        for c, label in enumerate(labels):
            m_fee = re.match(cfg["fees_re"], label)
            # fees_pair：左邊欄符合時，這欄一定是費用 (不套 fees_skip)
            paired = cfg.get("fees_pair") and c > 0 and re.search(cfg["fees_pair"], labels[c - 1])
            if m_fee and not (cfg.get("fees_skip") and re.search(cfg["fees_skip"], label) and not paired) \
                    and c not in {x for x, _ in fees}:
                fees.append((c, m_fee.group(1).strip()))
    # 外幣費用名稱加註幣別 (同 VN09 MOL 的使用者答案 "… (USD)")
    # fee_cur：表頭沒寫幣別、但整張費用都是外幣 (VN14 CN 表，合計欄是 TOTAL (USD))
    # (標題寫明 VND 的欄不加，例如 VN14 "FUMIGATION (VND)")
    fees = [(c, f"{desc} ({cur})" if (cur := next((x for x in ("USD", "RMB") if re.search(
        rf"\b{x}\b", labels[c])), None if re.search(r"\bVND\b", labels[c]) else cfg.get("fee_cur")))
        and cur not in desc.upper() else desc) for c, desc in fees]
    fields = {code: find(lab) for code, lab in cfg.get("fields", {}).items()}
    boxes = {size: _vt_col(labels, lab) for size, lab in cfg.get("boxes", {}).items()}
    # 表內沒有客戶名稱時，看工作表名稱的寶成廠別 ("FCL POU CHEN" -> POU CHEN)
    pou = re.search(r"(?<![A-Z])POU\s*-?\s*(YUEN|HUNG|CHEN|SUNG|LI|SHENG|YI)(?![A-Z])", _vt_norm(name))
    common = {"supplier": supplier, "consignee": _vt_consignee(rows) or (f"POU {pou.group(1)}" if pou else None),
              "invoice_date": _vt_doc_date(rows), **cfg.get("const", {})}
    if cfg.get("date_sheet_mmdd") and not common["invoice_date"] and year_hint:
        # 工作表名稱是月日 ("0405")，年份取檔名 (VN06 TRUCK AIR+CFS IMPORT)
        m_md = re.search(r"(?<!\d)(\d{2})(\d{2})(?!\d)", name)
        if m_md:
            common["invoice_date"] = _format_date_en(m_md.group(2), m_md.group(1), year_hint)
    doc_key = None
    if cfg.get("doc_key"):
        # 整張對帳表一個號碼 (例如 VN19 "Mã bảng kê: PYV260901"、VN11 表尾
        # "Invoice No.: 28433")：當發票號碼
        all_text = " ".join(_vn_text(c) for r in rows for c in r if c not in (None, ""))
        m_doc = re.search(cfg["doc_key"], all_text)
        doc_key = m_doc.group(1) if m_doc else None
    if cfg.get("doc_date"):
        for r in rows:
            cells = [c for c in r if c not in (None, "")]
            for i, c in enumerate(cells[:-1]):
                if cfg["doc_date"] in _vt_norm(c) and _vn_date(cells[i + 1]):
                    common["invoice_date"] = _vn_date(cells[i + 1])
    items, totals = [], {}
    start = h + 2 if two_rows else h + 1
    end = len(rows)
    for i in range(start, len(rows)):
        first = _vt_norm(next((c for c in rows[i] if c not in (None, "")), ""))
        if re.match(r"^(TONG|TOTAL|TTL|CONG|VAT|XIN VUI|SO TK|THANH TIEN$)", first):
            end = i
            break
    # 整張對帳表只有一個 VAT (合計列下方 "VAT 8% | 1,638,528")：單獨一張，
    # 發票號碼用設定的對帳表號碼欄 (vat_doc) 或工作表名稱
    for row in rows[end:end + 6]:
        cells = [c for c in row if c not in (None, "")]
        # "THUẾ SUẤT GTGT (VAT RATE)" (舊字型讀成 "THUEÁ SUAÁT") 也是 VAT 列
        if cells and re.match(r"^(V\.?A\.?T\b|THUE)", _vt_norm(cells[0])) and cfg.get("vat_row"):
            nums = [_vn_amount(c) for c in cells[1:] if _vn_amount(c)]
            if nums:
                items.append({"description": _vn_text(cells[0]), "amount": None, "vat": nums[0],
                              **common, "invoice_no": doc_key or f"{name.strip()} VAT"})
            break
    if cfg.get("vat_total"):
        # 表尾的整張 VAT ("Total VAT amount | 2590843.48")，取該列最後一個數字
        for row in rows[end:]:
            if any(cfg["vat_total"] in _vt_norm(c) for c in row):
                nums = [_vn_amount(c) for c in row if _vn_amount(c)]
                if nums:
                    items.append({"description": cfg.get("vat_desc", "VAT"), "amount": None,
                                  "vat": nums[-1], **common, "invoice_no": doc_key or f"{name.strip()} VAT"})
                break
    last_key = last_date = None
    last_base, group_n, last_fields = "", 0, {}
    sheet_year = None
    if col.get("date") is not None:
        years = collections.Counter((_vn_date(_vn_cell(r, col["date"])) or "")[-4:] for r in rows[start:end])
        years.pop("", None)
        sheet_year = years.most_common(1)[0][0] if years else (common["invoice_date"] or "")[-4:] or None
    for row in rows[start:end]:
        key = next((_vn_text(_vn_cell(row, k)) for k in keys if _vn_text(_vn_cell(row, k))), "")
        # 一格裡換行的好幾個號碼 ("102001\n270960") 也拆開
        joined = list(dict.fromkeys(
            v for k in join_keys if k is not None
            for v in (_vn_cell(row, k).split() if isinstance(_vn_cell(row, k), str)
                      else [_vn_text(_vn_cell(row, k))]) if v))
        if key and joined:
            key = ", ".join(joined)
        if col.get("base_key") is not None:
            # 一票 (Booking) 好幾張發票 (VN01 FREIGHT STATEMENT)：發票號碼只寫在每張的第一列；
            # 合計欄 (group_col) 有值就是新的一張，沒有號碼時用「提單號」或「提單號-序號」
            base = _vn_text(_vn_cell(row, col["base_key"]))
            if col.get("group_col") is None:
                key = ""        # 舊版只有每票合計：發票號碼用提單號，整票一起勾稽
            if base:
                last_base, group_n, last_fields = base, 0, {}
            starts = col.get("group_col") is not None and _vn_num(_vn_cell(row, col["group_col"])) is not None
            if not key and (base or starts) and last_base:
                group_n += 1
                key = last_base if group_n == 1 else f"{last_base}-{group_n}"
        if not key:
            # 續行：單號空白但有費用 (VN14 同一票第二列只有 D/O 費)，算在上一張
            if cfg.get("carry") and last_key and any(_vn_amount(_vn_cell(row, c)) for c, _ in fees):
                key = last_key
            else:
                continue
        if cfg.get("einv_key"):
            # 電子發票只用號碼 ("1C26TYY-4484" -> "4484"，使用者 2026-10-27)
            # "1-C26TOD-28667" (VN14) 前面多一個 "1-"
            # "C26TOD/17995" 用斜線
            key = re.sub(r"^(?:\d-)?\d?[A-Z]\d{2}[A-Z]{2,4}\s*[-/]\s*", "", key)
        last_key = key
        # vat_doc_only：對帳表號只給整張的 VAT 列用，費用列用自己的號碼 (VN11 Inv#)
        key = key if cfg.get("vat_doc_only") else (doc_key or key)
        head = dict(common, invoice_no=key)
        for code, c in fields.items():
            if c is not None and _vn_text(_vn_cell(row, c)):
                value = _vn_cell(row, c)
                head[code] = (_vn_date(value) or _vn_text(value)) if code.endswith("_date") else _vn_text(value)
        for code in cfg.get("carry_fields", ()):
            # 續行沿用同一票上一列的欄位 (提單號、港口)
            if head.get(code):
                last_fields[code] = head[code]
            elif last_fields.get(code):
                head[code] = last_fields[code]
        if col.get("date") is not None:
            raw_date = _vn_cell(row, col["date"])
            # 日期後面可能接其他字 ("26/08/2026 - 146 KGS"，VN18)
            m_date = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", _vn_text(raw_date))
            row_date = _vn_date(raw_date) or (_format_date_en(*m_date.groups()) if m_date else None)
            # 沒有年份或年份兩位 ("04/08"、"AUG.07"、"12/08/26"、"20+21/8/2026")：
            # 年份用同一張表其他列的年份 (VN06 進口派車表，使用者 2026-10-27 確認是 2026)
            if not row_date and _vn_text(raw_date):
                row_date = _vt_loose_date(_vn_text(raw_date), sheet_year)
            # 同一趟車的後幾批不再寫日期 (VN06 進口派車表)：沿用上一列
            if not row_date and cfg.get("ffill_date"):
                row_date = last_date
            if row_date:
                head["invoice_date"] = last_date = row_date
        if col.get("port") is not None and " - " in _vn_text(_vn_cell(row, col["port"])):
            pol, pod = _vn_text(_vn_cell(row, col["port"])).split(" - ", 1)
            head.setdefault("port_of_loading", pol.strip())
            head.setdefault("port_of_discharge", pod.strip())
        def measure(name_):
            # Excel 浮點誤差 (6.912000000000001) 取到小數 3 位
            value = _vn_num(_vn_cell(row, col.get(name_))) if col.get(name_) is not None else None
            return f"{round(value, 3):g}" if value else None

        head["volume"] = _fmt_volume(
            containers=[(_vn_num(_vn_cell(row, c)), size) for size, c in boxes.items() if c is not None],
            cbm=measure("cbm"), kg=measure("kg"))
        # 費用名稱 "@欄位" 代表取該欄的文字 (例如 VN07 Remark 欄 "Late come charge")
        row_items = [{"description": (_vn_text(_vn_cell(row, find(desc[1:]))) or desc[1:])
                      if desc.startswith("@") else desc,
                      "amount": _vn_amount(_vn_cell(row, c)), "vat": None, **head}
                     for c, desc in fees if _vn_amount(_vn_cell(row, c))]
        for it in row_items:
            if isinstance(it["description"], str) and it["description"].startswith("-") and it["amount"]:
                # 減項 (費用名稱前加 "-")：金額取負號
                it["description"], it["amount"] = it["description"][1:], -abs(it["amount"])
        for it, (c, _) in zip(row_items, [f for f in fees if _vn_amount(_vn_cell(row, f[0]))]):
            own = _vn_text(_vn_cell(row, fee_keys.get(c))) if fee_keys.get(c) is not None else ""
            if own:
                # 電子發票只用號碼，去掉發票符號 ("1C26TSG-7734" -> "7734"，使用者 2026-10-27)
                # 符號也可能寫成 "1-C26TOD-8547" (2026-10-28 抽樣 VN14 "POU HUNG 03-2026 PHƯƠNG ĐÔNG export")
                it["invoice_no"] = re.sub(r"^\d-?[A-Z]\d{2}[A-Z]{2,4}-", "", own)
        if remainder is not None:
            # 小計減掉已列出的明細費用，剩下的就是主要費用 (VN06 SERVICE FEE：MISCELL 欄空白，
            # 金額只寫在第一個 TOTAL 欄)
            rest = (_vn_num(_vn_cell(row, remainder[0])) or 0) - sum(
                _vn_num(_vn_cell(row, c)) or 0 for c in remainder[2])
            if rest > 0.5:
                row_items.insert(0, {"description": remainder[1], "amount": rest, "vat": None, **head})
            # 沒有 2% 欄、只有「TOTAL (小計) | TOTAL (含 2%)」兩欄時，差額就是 2% 附加費
            last = col.get("total")
            if not remainder[3] and last not in (None, remainder[0]):
                extra = (_vn_num(_vn_cell(row, last)) or 0) - (_vn_num(_vn_cell(row, remainder[0])) or 0)
                if extra > 0.5:
                    row_items.append({"description": cfg.get("after_sub", "2%"), "amount": round(extra, 2),
                                      "vat": None, **head})
        if cfg.get("rest_fee") and col.get("total") is not None:
            # 合計欄減掉已列出的費用，剩下的補成一列 (VN06 舊版 EXP 表沒有 2% 欄，2026-10-28 抽樣
            # "PFP 20260223-20260228 Miscellaneous Fee"：PRICE 250,000、TOTAL 255,000)
            extra = (_vn_num(_vn_cell(row, col["total"])) or 0) - sum(it.get("amount") or 0 for it in row_items)
            if extra > 0.5:
                row_items.append({"description": cfg["rest_fee"], "amount": round(extra, 2), "vat": None, **head})
        vat =_vn_amount(_vn_cell(row, col.get("vat"))) if col.get("vat") is not None else None
        if vat:
            row_items.append({"description": cfg.get("vat_desc", "VAT"), "amount": None, "vat": vat, **head})
        if not row_items:
            continue
        items += row_items
        # 合計欄空白 (公式沒有存計算結果) 時當成沒有合計，不是 0
        t = totals.setdefault(key, [None, None])
        # total_add：另一個幣別的合計 (VN14 費用有 USD 也有 VND，使用者 2026-10-27 要求
        # 兩種都稽核)：USD 合計 + VND 合計 = 兩種費用的加總，任一幣別對不上都會不通過
        for i, name_ in ((0, "total"), (0, "total_add"), (1, "grand")):
            value = _vn_num(_vn_cell(row, col[name_])) if col.get(name_) is not None else None
            if value is not None:
                t[i] = (t[i] or 0) + value
    # 整張表一個 VAT 的列沒有自己的日期，抬頭也沒寫日期時：用這張表最晚的日期
    dated = [datetime.datetime.strptime(it["invoice_date"], "%d %b %Y") for it in items
             if it.get("invoice_date") and re.fullmatch(r"\d{2} \w{3} \d{4}", it["invoice_date"])]
    for it in items:
        if not it.get("invoice_date") and it.get("amount") is None and dated:
            it["invoice_date"] = max(dated).strftime("%d %b %Y")
    if cfg.get("sum_row"):
        # 沒有逐列合計的表：用表尾 TOTAL 列各費用欄的合計，跟整張表的費用加總勾稽
        # (invoice_no=None 代表整份檔案；VN05 POU CHEN 運費對帳表，使用者規定運費表核對整份合計)
        for row in rows[end:]:
            cells = [c for c in row if c not in (None, "")]
            if cells and re.match(cfg["sum_row"], _vt_norm(cells[0])):
                sheet_sum = sum(_vn_num(_vn_cell(row, c)) or 0 for c, _ in fees)
                prev = totals.get(None) or [0, None]
                totals[None] = [prev[0] + sheet_sum, None]
                break
    if cfg.get("sheet_total") and doc_key:
        # 整張對帳表一個號碼時，用表尾合計列勾稽：(未稅合計, 含稅合計) 的標題，
        # 例如 VN19 "TOTAL (VND)"、"SUBTOTAL (VND)"
        pair = [None, None]
        for row in rows[end:]:
            cells = [c for c in row if c not in (None, "")]
            for i, label in enumerate(cfg["sheet_total"]):
                if cells and label and _vt_norm(cells[0]).startswith(label):
                    nums = [_vn_num(c) for c in cells[1:] if _vn_num(c) is not None]
                    pair[i] = nums[-1] if nums else None
        totals[doc_key] = pair
    return items, totals


def _vt_requires_ok(rows: List[list], cfg: Dict) -> bool:
    """名稱符合的設定也要有 name_require 指定的欄 (同名工作表有新舊兩種版面時區分)。"""
    req = cfg.get("name_require")
    if not req:
        return True
    h = _vt_header_row(rows, cfg["header"])
    if h is None:
        return False
    labels = _vt_labels(rows, h, cfg.get("two_rows", True), cfg.get("above", False))
    return _vt_col(labels, *(req if isinstance(req, (list, tuple)) else (req,))) is not None


def _vt_sheets(book, cfgs: List[Tuple[str, Dict]]):
    """套好設定的工作表 (_vt_sheets_raw)，再去掉工作副本：一張表的費用 (發票號碼、費用名稱、
    金額、VAT) 全部包含在另一張表裡時只算一次 (2026-10-30 VN01：格式不同但內容一樣的
    "TONG"/"Sheet1"、"POU CHEN"/"POU CHEN (2)"，少幾欄的草稿 "9"，只有部分費用的 "TONG"
    vs "Bang ke"；以前兩張都算，金額重複)。"""
    got = list(_vt_sheets_raw(book, cfgs))
    if len(got) < 2:
        yield from got
        return
    sigs = [collections.Counter((it["invoice_no"], it["description"], it["amount"], it["vat"])
                                for it in _vt_parse_sheet(name, rows, cfg, "")[0])
            for name, rows, cfg in got]
    for i, item in enumerate(got):
        # 被另一張表完全包含 (內容一樣時保留前面那張)
        if sigs[i] and any(j != i and sigs[j] and not (sigs[i] - sigs[j])
                           and (sigs[j] != sigs[i] or j < i) for j in range(len(got))):
            continue
        yield item


def _vt_sheets_raw(book, cfgs: List[Tuple[str, Dict]]):
    """依工作表名稱 (正規表示式，比對去聲調大寫的名稱) 套設定；第一個符合的為準。
    整本都沒有名稱符合的工作表時 (2026-10-28 抽樣：工作表叫 "POU HUNG"、"TỔNG"、
    "XUAT"…)，改用表頭判斷：表頭列找得到、而且至少有一個費用欄的第一個設定。"""
    matched, rest = False, []
    # 內容完全相同的工作表只算一次 (2026-10-28 抽樣 VN02 "POU HUNG 09-2025"：
    # "HÀNG RỜI" 是 "NEO" 的副本，兩張都算會重複請款)
    seen = set()
    sheets = []
    for name, rows in book.sheets:
        sig = repr(rows)
        if sig not in seen:
            seen.add(sig)
            sheets.append((name, rows))
    for name, rows in sheets:
        for pattern, cfg in cfgs:
            if re.search(pattern, _vt_norm(name)) and _vt_requires_ok(rows, cfg):
                matched = True
                yield name, rows, cfg
                break
        else:
            rest.append((name, rows))
    # 有名稱符合的工作表時，其他工作表只套「有 require 必要欄位」的設定 (版面特徵明確，
    # 不會誤抓彙總表)；整本都沒有名稱符合時，所有設定都用表頭判斷
    for name, rows in rest:
        for _pattern, cfg in cfgs:
            if matched and not cfg.get("require"):
                continue
            if cfg.get("skip_sheet") and re.search(cfg["skip_sheet"], _vt_norm(name)):
                continue
            h = _vt_header_row(rows, cfg["header"])
            if h is None:
                continue
            labels = _vt_labels(rows, h, cfg.get("two_rows", True), cfg.get("above", False))
            req = cfg.get("require")
            if req and _vt_col(labels, *(req if isinstance(req, (list, tuple)) else (req,))) is None:
                continue        # 這個版面一定要有的欄 (區分很像的版面)
            if any(_vt_col(labels, r) is None for r in cfg.get("require_all", ())):
                continue        # 這幾欄都要有 (require 是任一欄)
            if any(_vt_col(labels, *(f[0] if isinstance(f[0], (list, tuple)) else (f[0],))) is not None
                   for f in cfg["fees"]) or (cfg.get("fees_re") and any(
                       re.match(cfg["fees_re"], lab) for lab in labels)) or (
                       cfg.get("fees_between") and _vt_col(labels, cfg["fees_between"][0]) is not None):
                yield name, rows, cfg
                break


def _vt_year(book) -> Optional[str]:
    """檔名裡的年份 ("04-2025 BINH TAN-…" -> "2025")，給只寫月日的工作表用。"""
    years = re.findall(r"(?<!\d)(20\d{2})(?!\d)", os.path.basename(getattr(book, "path", "") or ""))
    return years[-1] if years else None


_VT_PATH_CUSTOMERS = (
    # 路徑寫法 (去空白、去聲調大寫) -> 客戶簡稱，寫法同工作表名稱備援 "POU xxx"
    ("POUYUEN", "POU YUEN"), ("POUHUNG", "POU HUNG"), ("POUCHEN", "POU CHEN"),
    ("POUSUNG", "POU SUNG"), ("POUPHONG", "POU PHONG"), ("POULI", "POU LI"),
    ("YUEDE", "YUE DE"),
)


def _vt_path_customer(book) -> Optional[str]:
    """檔案內與工作表名稱都沒有客戶時，看檔名、再看所在資料夾 (2026-10-28 抽樣：
    VN06 "20260629-20260704(EXP-TOTAL).xls" 放在 "…\\POU HUNG\\2026\\08\\"、
    VN14 放在 "…\\YUEDE\\2026\\")。由檔名往上找，離檔案最近的為準。"""
    path = getattr(book, "path", "") or ""
    for part in reversed(re.split(r"[\\/]", path)):
        flat = re.sub(r"[^A-Z]", "", _vt_norm(part))
        for key, name in _VT_PATH_CUSTOMERS:
            if key in flat:
                return name
    return None


def _vn_full_consignee(name: Optional[str]) -> Optional[str]:
    """客戶簡稱 -> 全名 (_VN_CONSIGNEES)。去掉空白與符號後跟簡稱完全相同才換，其他照抄。"""
    flat = re.sub(r"[^A-Z]", "", _vt_norm(name or ""))
    return next((full for key, full in _VN_CONSIGNEES.items() if key.replace(" ", "") == flat), name)


def _vt_name_month(book) -> Optional[Tuple[int, int]]:
    """檔名裡的 (年, 月)："08.2026"、"T01-2026"、"T1.2026"、"AUGUST 2026" (2026-10-30 VN01
    "AUGUST 2026 IMPORT-YUEDE")。"""
    base = os.path.basename(getattr(book, "path", "") or "")
    m = re.search(r"(?<!\d)(\d{1,2})[.\-/](20\d{2})(?!\d)", base)
    if m and 1 <= int(m.group(1)) <= 12:
        return int(m.group(2)), int(m.group(1))
    months = ("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
    m = re.search(r"\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[A-Z]*[\s.,\-]*(20\d{2})(?!\d)", base.upper())
    if m:
        return int(m.group(2)), months.index(m.group(1)) + 1
    return None


def _vt_name_month_end(book) -> Optional[str]:
    """整本都沒有日期時，用檔名的月份，填該月最後一天 (使用者 2026-10-28 確認；VN06
    "PFP CFS -AIR TRUCK FEE 08.2026"、"PYV-CFS T01-2026-Air")。"""
    ym = _vt_name_month(book)
    if not ym:
        return None
    last = calendar.monthrange(*ym)[1]
    return datetime.date(ym[0], ym[1], last).strftime("%d %b %Y")


def _vt_make(supplier: str, cfgs: List[Tuple[str, Dict]]):
    """產生這家供應商的 (sheet_parse, sheet_totals)。"""
    def sheet_parse(book) -> List[Dict]:
        items = []
        for name, rows, cfg in _vt_sheets(book, cfgs):
            items += _vt_parse_sheet(name, rows, cfg, supplier, _vt_year(book))[0]
        # 同一張發票在某張表才有的欄位 (例如材積只在 LCC 表)，補到其他列
        known: Dict[str, Dict] = {}
        for it in items:
            for code, value in it.items():
                if value not in (None, "") and code not in ("description", "amount", "vat"):
                    known.setdefault(it["invoice_no"], {}).setdefault(code, value)
        items = [{**known.get(it["invoice_no"], {}), **{k: v for k, v in it.items() if v not in (None, "")},
                  "vat": it.get("vat"), "amount": it.get("amount")} for it in items]
        # 一本對帳表只有一個客戶：沒有客戶名稱的工作表 (VN03 TẠI CHỖ) 用其他表的
        customer = next((it["consignee"] for it in items if it.get("consignee")), None) \
            or _vt_path_customer(book)
        # 沒有日期的列 (整張表的 VAT 列、額外費用列)：用同一本對帳表最晚的日期 (VN16 Hang Xe Tai)
        dated = [datetime.datetime.strptime(it["invoice_date"], "%d %b %Y") for it in items
                 if isinstance(it.get("invoice_date"), str) and re.fullmatch(r"\d{2} \w{3} \d{4}", it["invoice_date"])]
        latest = max(dated).strftime("%d %b %Y") if dated else _vt_name_month_end(book)
        ym = _vt_name_month(book)
        for it in items:
            # 日期欄只打「日」(Excel 顯示成 1900 年，例 "28" -> 28 Jan 1900)：用檔名的年月補上
            # (2026-10-28 抽樣 VN16 "BANG KE POU LI -T1.2026 (1)" Cont Nhập 第 16 列)
            d = it.get("invoice_date")
            if ym and isinstance(d, str) and d.endswith(" 1900") and int(d[:2]) <= calendar.monthrange(*ym)[1]:
                it["invoice_date"] = datetime.date(ym[0], ym[1], int(d[:2])).strftime("%d %b %Y")
            # 客戶簡稱 ("POU HUNG") 統一成全名 (使用者 2026-10-28 確認)
            it["consignee"] = _vn_full_consignee(it.get("consignee") or customer)
            if latest and not it.get("invoice_date"):
                it["invoice_date"] = latest
        return [{"header": {}, "items": items or [{"description": None, "amount": None}]}]

    def sheet_totals(book) -> List[Dict]:
        merged: Dict[str, List[float]] = {}
        has_total = has_grand = False
        for name, rows, cfg in _vt_sheets(book, cfgs):
            has_total |= "total" in cfg or "sheet_total" in cfg
            has_grand |= "grand" in cfg or bool(cfg.get("sheet_total", (None, None))[1])
            for key, pair in _vt_parse_sheet(name, rows, cfg, supplier, _vt_year(book))[1].items():
                m = merged.setdefault(key, [None, None])
                for i, value in enumerate(pair):
                    if value is not None:
                        m[i] = (m[i] or 0) + value
        return [{"invoice_no": key, "total": t if has_total else None,
                 "grand_total": g if has_grand else None} for key, (t, g) in merged.items()]
    return sheet_parse, sheet_totals


def _vt_register(key: str, label: str, detect, supplier: str, cfgs) -> None:
    parse, totals = _vt_make(supplier, cfgs)
    register_supplier(key, label, detect, None, sheet_parse=parse, sheet_totals=totals)


# --- VN01 TRANS WAGON (CÔNG TY TNHH LIÊN VẬN QUỐC TẾ，MST 0302…) -------------
# 一個活頁簿三張正式對帳表 (各附一張越南增值稅發票)：S1-LCC 當地費用、S1-O.F
# 海運費、S1-EXW 工廠交貨費；另有「TỔNG」彙總表與「EXWORK」工作表 (重複，不取)。
# 發票號碼 = HAWB.No (同一票的三張表合併勾稽)。VAT 0%。
_VT_TRANS_WAGON = [
    (r"^S1-LCC", {"header": "HAWB.NO", "key": ["HAWB.NO"],
                  "fees": [("PHI CHUNG TU", None), ("PHI DAI LY", None), ("PHI PHU TROI", None),
                           ("PHI BOC XEP", None), ("PHI XEP DO", None), ("PHI VE SINH", None),
                           ("PHI QUAN LY", None), ("DRAYAGE", None), ("LOADING", None),
                           # "PHÍ KHÁC" 其他費用 (2026-10-30 "CONG NO T07-2026 (WAGON) POU CHEN" 11SEA2607097)
                           ("PHI KHAC", None)],
                  "vat": "VAT", "vat_desc": "VAT", "total": "TONG CONG (TOTAL) VND",
                  "grand": "=TONG CONG (TOTAL) TOTAL", "port": "CANG DI",
                  "boxes": {"20": "CONT 20", "40": "CONT 40"}, "cbm": "CBM",
                  "fields": {"bl_no": "HAWB.NO"}}),
    (r"^S1-O\.?F", {"header": "HAWB.NO", "key": ["HAWB.NO"],
                    "fees": [("CUOC BIEN", None)], "vat": "VAT", "vat_desc": "VAT",
                    "total": "CUOC BIEN", "grand": "=CUOC BIEN (OCEAN FREIGHT) TOTAL",
                    "port": "CANG DI", "fields": {"bl_no": "HAWB.NO"}}),
    (r"^S1-EXW", {"header": "HAWB.NO", "key": ["HAWB.NO"],
                  "fees": [("PHI EXWORK", None)], "vat": "VAT", "vat_desc": "VAT",
                  "total": "PHI EXWORK", "grand": "=PHI EXWORK TOTAL",
                  "port": "CANG DI", "fields": {"bl_no": "HAWB.NO"}}),
    # 2026-10-28 抽樣的另外兩種版面 (工作表名稱不固定，用表頭判斷)：
    # (a) FREIGHT STATEMENT：每列一項費用 (費用名稱在 Rate 欄：DO/THC/CFS/O/FRT)，
    #     同一張發票 (Invoice No) 好幾列，只有第一列寫發票號碼與合計
    (r"(?!)", {"header": "BOOKING NO", "key": ["=INVOICE NO", "BOOKING NO"], "carry": True,
               "fees": [("=VND", "@RATE")], "vat": "VAT (8%)", "vat_desc": "VAT 8%",
               "total": "TOTAL(VND) NOT VAT LOCAL", "grand": "=TOTAL(VND) VAT LOCAL CHARGE",
               "boxes": {"20": "MEASUREMENT 20'", "40": "MEASUREMENT 40'"}, "cbm": "MEASUREMENT CMB",
               "fields": {"bl_no": "BOOKING NO"}}),
    # (b) TỔNG：每票一列，每項費用 USD / VND 兩欄 (取 VND)，GRAND TOTAL = 當地費用含稅 + O/F + EXWORK
    (r"(?!)", {"header": "BILL NO", "key": ["BILL NO"], "fees": [],
               # "NOT VATLOCAL CHARGE" (少空格，2026-10-30 "CONG NO T02-2026 PYV IMP (AIR-BK)")
               "fees_re": r"TOTAL\(VND\)\s*(?:NOT\s*VAT\s*)?(.+?)(?:\s+DUONG B(?:AY|IEN))?$",
               # 小計欄：表頭有 "O/F+EXW" 也有 "VAT OF+EXW 0%" (2026-10-29 抽樣 POU LI T07/T08-2026)
               "fees_skip": r"LOCAL CHARGE|F\+EXW",
               # 左邊是 "LOCAL CHARGE (USD)" 的那欄是 Local Charge 費用 (USD 換算 VND)，不是小計
               # (2026-10-30 "Cong no T11-25 PYV IMP (SEA-BK)"：費用欄與小計欄標題一樣)
               "fees_pair": r"^LOCAL CHARGE \(USD\)",
               "vat": ("VAT (8%)", "=VAT"), "vat_desc": "VAT 8%", "grand": "GRAND TOTAL",
               "boxes": {"20": "MEASUREMENT 20", "40": "MEASUREMENT 40"}, "cbm": "MEASUREMENT CMB",
               "kg": "MEASUREMENT KGS",
               "fields": {"bl_no": "BILL NO", "arrive_date": "NGAY TAU DEN"}}),
    # (c) FREIGHT STATEMENT (Freight Detail 欄是費用名稱，AMOUNT /VND、VAT (8%)、每票第一列含稅合計)
    #     一票 (Booking) 可能開好幾張發票：右邊 TOTAL(VND) 三欄是每張發票的合計 (有的版本沒有，
    #     就用每票第一列的 Total Amount included VAT)
    (r"(?!)", {"header": "FREIGHT DETAIL", "key": ["=INVOICE NO"], "carry": True, "base_key": "BOOKING NO",
               "group_col": "TOTAL(VND) NOT VAT LOCAL", "carry_fields": ["bl_no", "port_of_loading"],
               "fees": [(("AMOUNT /VND", "AMOUNT VND"), "@FREIGHT DETAIL")],
               "vat": "VAT (8%)", "vat_desc": "VAT 8%", "total": "TOTAL(VND) NOT VAT LOCAL",
               "grand": ("=TOTAL(VND) VAT LOCAL CHARGE", "TOTAL AMOUNT INCLUDED VAT"),
               "boxes": {"20": "MEASUREMENT 20'", "40": "MEASUREMENT 40'"}, "cbm": "MEASUREMENT CMB",
               "fields": {"bl_no": "BOOKING NO"}}),
    # (e) FREIGHT STATEMENT (POU CHEN 2025-08~)：費用只有 USD 分類，越南盾只有 Local Charge 與
    #     OF+EXWORK 兩個合計；右邊每列一張發票 (INVOICE NO、DATE、未稅、VAT、含稅)
    (r"(?!)", {"header": "BOOKING NO", "key": ["=INVOICE NO"], "einv_key": True,
               "fees": [("TOTAL(VND) LOCAL CHARGE", "Local Charge"), ("TOTAL(VND) OF+EXWORK", "O/F + EXWORK")],
               "vat": "=VAT", "vat_desc": "VAT 8%", "total": "AMOUNT (NOT VAT)", "grand": "=TOTAL AMOUNT",
               "date": "=INVOICE DATE", "carry_fields": ["bl_no", "port_of_loading"],
               "boxes": {"20": "MEASUREMENT 20'", "40": "MEASUREMENT 40'"}, "cbm": "MEASUREMENT CMB",
               "fields": {"bl_no": "BOOKING NO", "port_of_loading": "NOTE PORT"}}),
    # (f) FREIGHT STATEMENT (POU CHEN 2025-06 進口，2026-10-29 抽樣 "CONG NO T06-25 POU CHEN-IMP")：
    #     每票兩列，各是一張發票 (右邊 INVOICE NO "1C25TYY - 5297"、DATE、未稅、VAT、含稅)；
    #     越南盾費用欄 FREIGHT SEA / AIR FREIGHT / THC,CFS FEE / D/O, HDLG FEE (Total local charges
    #     與 TOTAL 是小計)
    (r"(?!)", {"header": "BOOKING NO", "key": ["=INVOICE NO"], "einv_key": True,
               "require": "D/O, HDLG FEE",
               "fees": [("=FREIGHT SEA VND", "FREIGHT SEA"), ("=AIR FREIGHT VND", "AIR FREIGHT"),
                        ("=THC,CFS FEE VND", "THC,CFS FEE"), ("=D/O, HDLG FEE VND", "D/O, HDLG FEE")],
               "vat": "=VAT", "vat_desc": "VAT 8%", "total": "AMOUNT (NOT VAT)", "grand": "=TOTAL AMOUNT",
               "date": "=INVOICE DATE", "carry_fields": ["bl_no"],
               "boxes": {"20": "MEASUREMENT 20'", "40": "MEASUREMENT 40'"}, "cbm": "MEASUREMENT CMB",
               "fields": {"bl_no": "BOOKING NO"}}),
    # (d) 紅色發票清單：每列一張發票 (Red inv no，只用號碼)，金額已含稅 (沒有另列 VAT)
    (r"(?!)", {"header": "RED INV NO", "two_rows": False, "key": ["RED INV NO"], "einv_key": True,
               "fees": [("FREIGHT AIR", "Freight AIR"), ("FREIGHT SEA", "Freight SEA"),
                        ("MISCELLANEOUS", "Miscellaneous")],
               "total": "AMOUNT (VND)", "date": "=DATE", "fields": {"bl_no": "BILL NO"}}),
]


def detect_trans_wagon(text: str) -> bool:
    # 紅色發票清單 (Red inv no) 沒有公司名稱，用 TRANS WAGON 的發票符號 "1C26TYY" 辨識
    return "TRANS WAGON" in text.upper() or (
        "RED INV NO" in _vt_norm(text) and re.search(r"\b\dC\d{2}TYY\b", text) is not None)


# (e) PDF 電子增值稅發票 (HÓA ĐƠN GIÁ TRỊ GIA TĂNG，Serial 1C26TYY，2026-10-28 使用者
#     提供 1_C26TYY_11198.pdf)：費用行 "1 Phí xếp dỡ (THC) X40' 1 1.943.250 1.943.250 8%
#     155.460 2.098.710" (序號、名稱、單位、數量、單價、未稅、稅率、稅額、含稅)，每筆自帶
#     稅額；下一行起是 "POL - POD _ 船名/航次 _ 提單號碼"。數字是越式千分位 (點)。
_TW_PDF_ITEM_RE = re.compile(
    r"^\d+\s+(?P<desc>.+?)\s+(?P<unit>\S+)\s+(?P<qty>[\d.,]+)\s+(?P<price>[\d.,]+)\s+"
    r"(?P<amount>[\d.,]+)\s+(?P<rate>\d+%|KCT|KKKNT)\s+(?P<tax>[\d.,]+|-)\s+(?P<total>[\d.,]+)$",
    re.MULTILINE)
_TW_TAX_CODE = "0302101999"


def _tw_num(text: Optional[str]):
    """'1.943.250' -> 1943250 (越式千分位)。"""
    if not text or text == "-":
        return None
    try:
        value = float(text.replace(".", "").replace(",", "."))
    except ValueError:
        return None
    return int(value) if value.is_integer() else value


def detect_trans_wagon_pdf(text: str) -> bool:
    upper = text.upper()
    return _TW_TAX_CODE in text and "GIÁ TRỊ GIA TĂNG" in upper and "LIÊN VẬN QUỐC TẾ" in upper


def parse_trans_wagon_pdf(text: str) -> Dict:
    h: Dict[str, Optional[str]] = {"supplier": "CÔNG TY TNHH LIÊN VẬN QUỐC TẾ"}
    # 號碼印在「Số(No.):」標籤上一行；越南發票號碼只用號碼 (使用者 2026-10-27)
    h["invoice_no"] = (_search(r"^(\d{3,})\s*\n\s*S[ốo]\(No\.\)", text)
                       or _search(r"S[ốo]\(No\.\):\s*(\d+)", text))
    m = re.search(r"Ngày\(Date\)\s*(\d{1,2})\s*tháng\(month\)\s*(\d{1,2})\s*năm\(year\)\s*(\d{4})", text)
    h["invoice_date"] = _format_date_en(*m.groups()) if m else None
    h["consignee"] = _search(r"Mã của CQT[^\n]*\n\s*:\s*(.+)", text)
    items, boxes = [], []
    lines = text.splitlines()
    for i, line in enumerate(lines):
        m = _TW_PDF_ITEM_RE.match(line.strip())
        if not m:
            continue
        # 費用行下面到下一個費用行 / 合計之前的文字：航線 _ 船名 _ 提單號碼
        extra = []
        for nxt in lines[i + 1:]:
            if _TW_PDF_ITEM_RE.match(nxt.strip()) or "Tổng cộng" in nxt:
                break
            extra.append(nxt.strip())
        item = {"description": m.group("desc").strip(), "amount": _tw_num(m.group("amount")),
                "vat": _tw_num(m.group("tax"))}
        parts = [p.strip() for p in " ".join(extra).split(" _ ")]
        if len(parts) >= 3:
            route = re.split(r"\s+-\s+", parts[0], maxsplit=1)
            if len(route) == 2:
                item["port_of_loading"], item["port_of_discharge"] = route
            item["vessel"], item["bl_no"] = parts[1], parts[-1]
        unit = re.match(r"^X?(20|40|45)'?(HC|GP|DC)?$", m.group("unit").upper())
        if unit:
            boxes.append((_tw_num(m.group("qty")), unit.group(1) + (unit.group(2) or "")))
        items.append(item)
    for field in ("port_of_loading", "port_of_discharge", "vessel", "bl_no"):
        h[field] = next((it[field] for it in items if it.get(field)), None)
    h["volume"] = _fmt_volume(containers=boxes)
    return {"header": h, "items": items or [{"description": None, "amount": None, "vat": None}]}


def _tw_pdf_totals(text: str):
    m = re.search(r"Total payment\):\s*([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)", text)
    return tuple(_tw_num(g) for g in m.groups()) if m else (None, None, None)


_tw_sheet_parse, _tw_sheet_totals = _vt_make("CÔNG TY TNHH LIÊN VẬN QUỐC TẾ", _VT_TRANS_WAGON)
register_supplier(
    "TRANS_WAGON", "CÔNG TY TNHH LIÊN VẬN QUỐC TẾ (TRANS WAGON)",
    lambda text: detect_trans_wagon(text) or detect_trans_wagon_pdf(text), parse_trans_wagon_pdf,
    extract_total=lambda text: _tw_pdf_totals(text)[0],
    extract_grand_total=lambda text: _tw_pdf_totals(text)[2],
    sheet_parse=_tw_sheet_parse, sheet_totals=_tw_sheet_totals,
)


# --- VN02 VĨNH HƯNG LONG (拖車行) --------------------------------------------
# 兩個檔：運費對帳表 (CƯỚC VC，好幾張工作表，每列一個貨櫃) 與代墊吊櫃費 (CHI HỘ
# NÂNG HẠ，金額已含 VAT)。發票號碼 = 批號 SỐ LÔ (新批號優先，沒有用舊批號)；代墊吊櫃費
# 用第三方發票號碼 SỐ HĐ NÂNG, SỐ HĐ HẠ (逗號連起來，使用者 2026-10-27)，
# 同一批好幾個貨櫃合併勾稽；貨櫃號碼、櫃型 (20'/40'/45') 逐列。
_VT_VHL_COMMON = {
    "header": "SO CONTAINER", "key": ["SO LO MOI", "SO LO CU", "=SO LO", "SO CONTAINER"],
    "date": "NGAY HOAN THANH", "boxes": {"20": "LOAI CONT 20'", "40": "LOAI CONT 40'", "45": "LOAI CONT 45'"},
    "fields": {"container_no": "SO CONTAINER", "port_of_loading": ("CANG LAY", "=CANG"),
               "port_of_discharge": ("CANG HA", "CANG TRA")},
}
# 稽核只用含稅合計 (grand)：同一批號跨好幾張表 (運費、吊櫃、磅櫃、壓車)，各表只有含稅或
# 只有整列合計，未稅合計混在一起會對不上 (2026-10-28 抽樣)
_VT_VINH_HUNG_LONG = [
    # BK (IMP-EXP) 對帳表 (2026-10-28 抽樣，工作表 XUAT/NHẬP/NEO)：運費、吊櫃費、磅櫃費同一列，
    # TỔNG CỘNG = 整列合計；有 "TỔNG CỘNG HĐ NÂNG HẠ" 欄才是這個版面
    (r"(?!)", {**_VT_VHL_COMMON, "date": "=NGAY", "require": "TONG CONG HD NANG HA",
               # "CƯỚC" 單獨一欄 (2026-10-30 "JANUARY 2026- YUEDE" DU DUC -XUAT 表)
               "fees": [(x, None) for x in ("=CUOC VAN CHUYEN", "=CUOC VC", "=CUOC", "CUOC VC TRONG", "CUOC VC NGOAI",
                                            "HO TRO PHI CO SO HA TANG", "BD NANG", "=BDNH",
                                            "=PHI NANG", "=PHI HA", "PHI HA BAI SOM", "PHI HA CONT SOM",
                                            # "Phí hạ sớm" (2026-10-28 抽樣 "BK 012025 (IMP-EXP)")
                                            "=PHI HA SOM", "=VSC",
                                            "PHI D/O VSC", "=CAN XE", "PHI NEO XE",
                                            # "PHÍ CÂN" 算在 TỔNG CỘNG HĐ NÂNG HẠ 裡 (2026-10-28 抽樣
                                            # "POU HUNG 05-2025 Vinh Hung Long" XUAT NB 表)
                                            "=PHI CAN",
                                            # "PHÍ CÂN CONT" 在吊櫃費組裡、算進 TỔNG CỘNG HĐ NÂNG HẠ
                                            # (2026-10-30 "công nợ VHL T5-25" XUAT 表；f008 的 "CÂN CONT" 不取)
                                            "=PHI CAN CONT")],
               # "CÂN" 欄 (過磅費) 不算進 TỔNG CỘNG，另在磅櫃表請款，這裡不取 (f007 BK 072025)；
               # "CÂN CONT" 欄同樣不在 TỔNG 裡，使用者 2026-10-28 確認不擷取 (f008 công nợ VHL T6-25)
               "grand": ("=TONG CONG", "TONG CONG CUOC & PHI", "=TONG")}),
    # 壓車表 (NEO)：PHÍ NEO XE 是每日單價，TỔNG CỘNG = 單價 × SỐ NGÀY NEO 才是金額
    # (2026-10-28 抽樣 "POU HUNG 09-2025 Vinh Hung Long"：1,500,000 × 2 天 = 3,000,000)
    (r"(?!)", {**_VT_VHL_COMMON, "date": "NGAY VAO XUONG", "require": "PHI NEO XE",
               "fees": [("=TONG CONG", "PHÍ NEO XE")], "grand": "=TONG CONG"}),
    # 磅櫃 / 附加費表 (CÂN CONT、cân-phụ thu)
    (r"(?!)", {**_VT_VHL_COMMON, "date": ("=NGAY", "NGAY LAY"), "require": ("PHI CAN CONT", "CHI HO CAN"),
               "fees": [(x, None) for x in ("PT CAN CONT", "=PTCC", "PHI CAN CONT", "=PHU THU", "PHU THU CAN",
                                            "CHI HO CAN")],
               # 代墊過磅費已含 VAT：TỔNG 是含稅合計 (跟同一批的運費表合併勾稽)
               "grand": ("=TONG", "=TONG CONG")}),
    # 2026-10-28 抽樣補：散貨 (HÀNG LẺ：TIỀN CƯỚC VC、PHÍ BOT、代墊) 與超大櫃 (CONT QUÁ KHỔ：
    # VAT 欄叫 "VAT 8%"，同列另有代墊吊櫃費 PHÍ NÂNG/HẠ (GỒM VAT)、VSC)
    (r"CUOC VC|HANG LE", {**_VT_VHL_COMMON,
                  "fees": [("CUOC VC TRONG", None), ("CUOC VC NGOAI", None), ("=CUOC VAN CHUYEN", None),
                           ("=TIEN CUOC VC", None), ("SO TIEN PHU THU", None), ("PHI CO SO HA TANG", None),
                           ("PHU THU HA RONG", None),
                           ("BOI DUONG", None),
                           ("=PHI BOT", None), ("PHI NEO CONT (CO VAT)", None), ("=TIEN CHI HO", None),
                           # 散貨表的 "PHÍ NEO CONT" 算進 TỔNG TIỀN CƯỚC VC (2026-10-30 "JUNE 2026-YUEDE")
                           ("=PHI NEO CONT", None),
                           ("PHI NANG (GOM VAT)", None), ("PHI HA (GOM VAT)", None), ("CHI HO VSC", None)],
                  # VAT 欄也寫成 "VAT (8%)" (2026-10-30 "THÁNG 3- CƯỚC VC POU YUEN")
                  "vat": ("THUE VAT", "=VAT 8%", "=VAT (8%)"), "vat_desc": "Thuế VAT 8%",
                  # 沒有 VAT 欄的運費表，含稅合計就是 TỔNG TIỀN CƯỚC VẬN CHUYỂN
                  "grand": ("TONG CONG SO TIEN", "TONG TIEN CUOC VAN CHUYEN")}),
    (r"NEO CONT", {**_VT_VHL_COMMON, "date": None,
                   "fees": [("PHI NEO CONT", None), ("PHI NEO XE", None)], "vat": "THUE VAT",
                   "vat_desc": "Thuế VAT 8%",
                   "grand": ("TONG CONG SO TIEN", "=TONG CONG")}),
    (r"CHI HO", {**_VT_VHL_COMMON,
                 "fees": [("PHI NANG", None), ("PHI HA", None), ("VSC", None)],
                 "key_join": ["SO HD NANG", "SO HD HA"], "grand": "=TONG CONG"}),
]


def detect_vinh_hung_long(text: str) -> bool:
    return "VINH HUNG LONG" in _vt_norm(text)


# BK 對帳表的 "CÂN CONT" 放在吊櫃費組裡 (PHÍ NÂNG | PHÍ HẠ | CÂN CONT | SỐ HĐ)，合計欄叫 "TỔNG CỘNG"，
# 有算進合計，要擷取 (2026-10-28 抽樣 "BK 052025 (IMP-EXP)"：差 416,000 = CÂN CONT)。
# công nợ 版 (CÂN CONT 在吊櫃發票號碼後面、合計欄叫 "TỔNG") 照使用者決定不擷取
_VT_VINH_HUNG_LONG.insert(0, (r"(?!)", {
    **_VT_VINH_HUNG_LONG[0][1], "require_all": ("TONG CONG HD NANG HA", "=TONG CONG", "=CAN CONT"),
    "fees": _VT_VINH_HUNG_LONG[0][1]["fees"] + [("=CAN CONT", None)]}))
_vt_register("VINH_HUNG_LONG", "CÔNG TY TNHH VĨNH HƯNG LONG", detect_vinh_hung_long,
             "CÔNG TY TNHH VĨNH HƯNG LONG", _VT_VINH_HUNG_LONG)


# --- VN06 BÌNH TÂN (HỢP TÁC XÃ VẬN TẢI DỊCH VỤ DU LỊCH BÌNH TÂN，MST 0303140905) --
# 三種檔：雜費 Miscellaneous Fee (IMP/EXP，每列一批，VAT 在合計列下方，整張一個)、
# 出口運費 (NB/PHV)、進口派車表 (08-CFS/08-AIR，每列一批，逐列 VAT)。
# 發票號碼 = 批號 (SỐ LÔ / Lot No / Dec No)；車牌照 VN05 慣例填貨櫃號碼與船名。
_VT_BINH_TAN = [
    # 雜費表 (SERVICE FEE / IMP-TOTAL，2026-10-28 抽樣，工作表名稱不固定)：MISCELL | OT. | OTHER |
    # TOTAL (小計) | 2% | TOTAL。MISCELL 常空白，主要費用只寫在小計；合計取最後一個 TOTAL
    (r"(?!)", {"header": "PAY NO", "two_rows": False, "key": ["LOT NO", "DEC NO"], "require": "SHIP MODE",
               "fees": [], "fees_between": ("SHIP MODE", ("UNIT PRICE", "=TOTAL", "TOTAL")),
               "remainder": (("=TOTAL", "TOTAL 2%"), "Miscellaneous Fee"),
               "total": "$TOTAL", "vat_row": True, "date": "DEC DATE",
               "fields": {"order_no": "=INV"}, "cbm": "TOTAL CFS/AIR"}),
    # 出口運費 (EXPORT TRUCK，每個部門一張工作表)：每列一批 (SỐ LÔ HÀNG)，運費 + 過路費 (VÉ ĐƯỜNG)，
    # 表尾 "TOTAL :" 勾稽整張、"THUẾ SUẤT GTGT" 單獨一筆 VAT
    (r"(?!)", {"header": "SO LO HANG", "key": ["SO LO HANG"], "require": "VE DUONG",
               "fees": [("=SO TIEN", "Cước vận chuyển"), ("VE DUONG", "Vé đường")],
               "date": "=DATE", "sum_row": r"^TOTAL\s*:", "vat_row": True,
               "const": {"port_of_loading": "VIETNAM", "port_of_discharge": "VIETNAM"}}),
    # 同上的英文版 (2025-03)：KIND TRUCK | PRICE | AMOUNT | TICKET | F.TORY | EXPORT DECLARATION NO.
    (r"(?!)", {"header": "EXPORT DECLARATION NO", "key": ["EXPORT DECLARATION NO"], "require": "=TICKET",
               "fees": [("=AMOUNT", "Cước vận chuyển"), ("=TICKET", "Vé đường")],
               "date": "=DATE", "sum_row": r"^TOTAL\s*:", "vat_row": True,
               "const": {"port_of_loading": "VIETNAM", "port_of_discharge": "VIETNAM"}}),
    # 進口運費 PFQ (TRUCK AIR+CFS IMPORT，工作表名稱是日期 "0405")：Truck/Port/D/O,THC/Storage Fee、
    # VAT、Clearing Expense (未稅) 與 Total；每列的 Invoice No. 是批號
    (r"(?!)", {"header": "CLEARING", "key": ["INVOICE NO"], "require": "CLEARING EXPENSE",
               "fees": [("TRUCK", "Truck Fee"), ("PORT", "Port Fee"), ("D/O,THC", "D/O,THC Fee"),
                        ("STORAGE", "Storage Fee")],
               "vat": "=VAT", "vat_desc": "VAT 8%", "total": "CLEARING EXPENSE", "grand": "=TOTAL",
               "date_sheet_mmdd": True,
               "const": {"port_of_loading": "VIETNAM", "port_of_discharge": "VIETNAM"}}),
    # 舊版 (有 PRICE 欄)；沒有 PRICE 欄的 IMP/EXP 表改用上面的新版設定
    (r"^(IMP|EXP)$", {"header": "PAY NO", "two_rows": False, "key": ["LOT NO", "DEC NO"], "name_require": "=PRICE",
                      "fees": [("MISCELL", None), ("=OT.", None), ("=OTHER", None),
                               ("=PRICE", "Miscellaneous Fee"), ("=2%", "2%")], "rest_fee": "2%",
                      "total": "=TOTAL", "vat_row": True, "date": "DEC DATE",
                      "fields": {"order_no": "=INV"}, "cbm": "TOTAL CFS/AIR"}),
    (r"^(NB|PHV)$", {"header": "SO LO", "key": ["=SO LO"],
                     "fees": [("CUOC VC CHUA VAT", None), ("SO TIEN PHU THU", None)],
                     "total": "TONG TIEN THANH TOAN", "vat_row": True, "date": "=NGAY",
                     "cbm": "SO LUONG CBM",
                     "fields": {"vessel": "SO XE", "container_no": "SO XE", "order_no": "SO INV"}}),
    # --- 2026-10-28 抽樣補的運費表 (工作表名稱也有 CFS/AIR，用必要欄位區分) ---
    # PLV CFS / PLV AIR：AMOUNT | % TĂNG | SỐ TIỀN % TĂNG | AMOUNT SAU TĂNG | VAT | TOTAL
    (r"(?!)", {"header": "AMOUNT SAU TANG", "two_rows": False, "key": ["SO LO", "=LOT NO"],
               "require": "AMOUNT SAU TANG",
               "fees": [("=AMOUNT", "Cước vận chuyển"), ("SO TIEN % TANG", "Phụ phí xăng dầu")],
               "vat": "=VAT", "vat_desc": "VAT 8%", "total": "AMOUNT SAU TANG", "grand": "=TOTAL",
               "date": "NGAY NHAN HANG", "ffill_date": True, "cbm": "SO CBM", "kg": "SO KG",
               "const": {"port_of_loading": "VIETNAM", "port_of_discharge": "VIETNAM"},
               "fields": {"vessel": "=SO XE", "container_no": "=SO XE"}}),
    # CFS CHECK air：每趟車的運費 (第一個 Số tiền) 分攤到各批 (第二個 Số tiền)，加油價附加費
    (r"(?!)", {"header": "PHU PHI SO TIEN", "two_rows": False, "key": ["SO LO"], "require": "PHU PHI SO TIEN",
               "fees": [("$SO TIEN (CHUA VAT)", "Cước vận chuyển"), ("PHU PHI SO TIEN", "Phụ phí xăng dầu")],
               "vat": "THUE 8%", "vat_desc": "VAT 8%", "total": "TONG TIEN ( CHUA", "grand": "=TONG TIEN",
               "date": "NGAY NHAN HANG", "ffill_date": True, "kg": "SO KG",
               "const": {"port_of_loading": "VIETNAM", "port_of_discharge": "VIETNAM"},
               "fields": {"vessel": "=SO XE", "container_no": "=SO XE"}}),
    # BẢNG KÊ VẬN CHUYỂN (PYV-CFS BK)：Số tiền bao gồm phụ phí | THUẾ 8% | Tổng tiền，一趟車好幾批
    (r"(?!)", {"header": "SO TIEN BAO GOM", "two_rows": False, "key": ["=SO LO"], "require": "SO TIEN BAO GOM",
               "fees": [("SO TIEN BAO GOM", "Cước vận chuyển (gồm phụ phí)")],
               "vat": "THUE 8%", "vat_desc": "VAT 8%", "grand": "TONG TIEN",
               "date": "NGAY LAY", "ffill_date": True, "kg": "=KG",
               "const": {"port_of_loading": "VIETNAM", "port_of_discharge": "VIETNAM"},
               "fields": {"vessel": "=SO XE", "container_no": "=SO XE", "bl_no": "B/L FACT"}}),
    # 分攤版派車表 (2026-10-28 抽樣 "POU HUNG 04-2026 … BINH TAN" 工作表 CFS)：第一個 SỐ TIỀN (CHƯA VAT)
    # 是整趟運費，SỐ TIỀN CT 是分攤到這批的運費，加 % tăng giá dầu = 第二個 SỐ TIỀN (CHƯA VAT)
    (r"(?!)", {"header": "SO TIEN CT", "two_rows": False, "key": ["=SO LO"], "require": "SO TIEN CT",
               "fees": [("SO TIEN CT", "Cước vận chuyển"), ("% TANG GIA DAU", "Phụ phí xăng dầu")],
               "vat": "THUE 8%", "vat_desc": "VAT 8%", "total": "$SO TIEN (CHUA VAT)", "grand": "TONG TIEN",
               "date": "NGAY NHAN HANG", "ffill_date": True, "cbm": "SO CBM", "kg": "SO KG",
               "const": {"port_of_loading": "VIETNAM", "port_of_discharge": "VIETNAM"},
               "fields": {"vessel": "=SO XE", "container_no": "=SO XE"}}),
    # IMPORT TRUCK FEE：SỐ TIỀN (CHƯA VAT) | THUẾ 8% | TỔNG TIỀN，每列一批 (lot no.)
    (r"(?!)", {"header": "SO TIEN (CHUA VAT)", "two_rows": False, "key": ["=LOT NO.", "=SO LO"],
               "require": "SO TIEN (CHUA VAT)",
               "fees": [("SO TIEN (CHUA VAT)", "Cước vận chuyển")], "vat": "THUE 8%", "vat_desc": "VAT 8%",
               "total": "SO TIEN (CHUA VAT)", "grand": "TONG TIEN", "date": "NGAY NHAN HANG", "ffill_date": True,
               "cbm": "SO CBM", "kg": "SO KG",
               "const": {"port_of_loading": "VIETNAM", "port_of_discharge": "VIETNAM"},
               "fields": {"vessel": "=SO XE", "container_no": "=SO XE"}}),
    # PLV CHECK 派車表 (2026-10-28 抽樣 "PFP CFS -AIR TRUCK FEE 08.2026"，工作表 CFS/AIR PLV CHECK)：
    # AMOUNT | SỐ TIỀN % TĂNG | SỐ TIỀN SAU KHI TĂNG | VAT | TOTAL；一趟車好幾批，同車其他批金額 0
    (r"(?!)", {"header": "SO TIEN SAU KHI TANG", "two_rows": False, "key": ["SO LO"],
               "require": "SO TIEN SAU KHI TANG",
               "fees": [("=AMOUNT", "Cước vận chuyển"), ("SO TIEN % TANG", "Phụ phí xăng dầu")],
               "vat": "=VAT", "vat_desc": "VAT 8%", "total": "SO TIEN SAU KHI TANG", "grand": "=TOTAL",
               "date": "NGAY NHAN HANG", "ffill_date": True, "cbm": "SO CBM", "kg": "SO KG",
               "const": {"port_of_loading": "VIETNAM", "port_of_discharge": "VIETNAM"},
               "fields": {"vessel": "=SO XE", "container_no": "=SO XE"}}),
    # PYV 空運派車表 (2026-10-28 抽樣 "PYV-CFS T01-2026-Air")：表頭 "Số tiền ( chưa VAT)" 括號後多空白；
    # 兩欄都叫 Số lô，第二欄才是完整批號 (60102025025662 / PYV2025006854)
    (r"(?!)", {"header": "SO TIEN ( CHUA VAT)", "two_rows": False, "key": ["$=SO LO"],
               "require": "SO TIEN ( CHUA VAT)",
               "fees": [("SO TIEN ( CHUA VAT)", "Cước vận chuyển")], "vat": "TIEN THUE 8%", "vat_desc": "VAT 8%",
               "total": "SO TIEN ( CHUA VAT)", "grand": "=TONG TIEN", "date": "NGAY NHAN HANG",
               "ffill_date": True, "kg": "=SO KG",
               "const": {"port_of_discharge": "VIETNAM"},
               "fields": {"vessel": "=SO XE", "container_no": "=SO XE", "port_of_loading": "TEN CANG"}}),
    # 舊版 08-CFS / 08-AIR (有 ST/1 LÔ 欄)
    (r"CFS|AIR", {"header": "SO LO", "two_rows": False, "key": ["SO LO"], "name_require": "ST/1 LO",
                  "fees": [("ST/1 LO CHUA VAT", "Cước vận chuyển"), ("PHU PHI XANG DAU", None)],
                  "total": "TONG TIEN/1 LO", "vat": "VAT 8%", "vat_desc": "VAT 8%",
                  "grand": ("=TONG CONG", "=TOTAL"), "date": "NGAY NHAN HANG", "ffill_date": True,
                  "cbm": "CMB", "kg": "=KG",
                  "fields": {"vessel": "SO XE", "container_no": "SO XE", "port_of_loading": "=CANG"}}),
]


def detect_binh_tan(text: str) -> bool:
    # 不能只看 "BÌNH TÂN"：寶元的地址就在平新郡 (Quận Bình Tân)，別家帳單也會出現
    norm = _vt_norm(text)
    return ("0303140905" in norm or "HOP TAC XA VAN TAI DICH VU DU LICH" in norm
            or ("PFQ-" in norm and "MISCELLANEOUS FEE" in norm)
            or "DANH SACH XE VAN CHUYEN HANG" in norm
            # 2026-10-28 抽樣：沒有公司名稱的雜費表 (Pay No/Fact No/Lot No/Ship Mode 表頭)
            # 與簽名欄只寫 "HTX VT-DV-DL BÌNH TÂN" 的運費表
            or (all(k in norm for k in ("PAY NO", "FACT NO", "SHIP MODE"))
                and ("LOT NO" in norm or "DEC NO" in norm))
            or "HTX VT-DV-DL BINH TAN" in norm or "HTX VT DV DL BINH T" in norm   # 舊字型 "BÌNH TAÂN"
            or ("PTT NO" in norm and "ADVANCE NO" in norm and "CLEARING" in norm)
            # IMPORT TRUCK FEE：沒有公司名稱，用表頭組合辨識
            or (all(k in norm for k in ("NGAY NHAN HANG", "SO XE", "SO LO", "SO TIEN (CHUA VAT)", "LOT NO"))
                and "SONG TRIEU" not in norm))


_vt_register("BINH_TAN", "HỢP TÁC XÃ VẬN TẢI DỊCH VỤ DU LỊCH BÌNH TÂN", detect_binh_tan,
             "HỢP TÁC XÃ VẬN TẢI DỊCH VỤ DU LỊCH BÌNH TÂN", _VT_BINH_TAN)


# --- VN12 NGUYÊN HƯỞNG (空運請款) ---------------------------------------------
# 每項費用都有 RMB 與 VND 兩欄，只取 VND；發票號碼 = Waybill number。
_VT_NGUYEN_HUONG = [
    (r".", {"header": "WAYBILL", "key": ["WAYBILL"],
            "fees": [("AIR FREIGHT VND", "Air freight"), ("CUSTOMS FEES (BILL) VND", "Customs fees"),
                     ("BILL OF LADING FEE (BILL) VND", "Bill of lading fee"),
                     ("ADDITIONAL FEE (BILL) VND", "Additional fee"),
                     ("WAREHOUSING FEE (BILL) VND", "Warehousing fee"),
                     ("PAD EXPENSE VND", "Pad expense"), ("D/O (BILL) VND", "D/O")],
            "total": "TOTAL AMOUNT VND", "date": "PICKUP DATE", "kg": "WEIGHT",
            "fields": {"bl_no": "WAYBILL", "shipper": "=SHIPPER"}}),
]


def detect_nguyen_huong(text: str) -> bool:
    return "NGUYEN HUONG" in _vt_norm(text)


_vt_register("NGUYEN_HUONG", "CÔNG TY TNHH NGUYÊN HƯỞNG", detect_nguyen_huong,
             "CÔNG TY TNHH NGUYÊN HƯỞNG", _VT_NGUYEN_HUONG)


# --- VN14 ORIENTAL LOGISTICS MULTIMODAL (Công Ty TNHH Logistics Vận Tải Đa Phương Thức) --
# 三種表 (CN 進口、TRUCK、SEA 出口)。費用是 USD (少數 VND)，每列另有越南盾的
# 未稅金額、VAT、含稅金額與發票號碼 SỐ HĐ。費用照抄 USD 金額、名稱加 "(USD)"
# (同 MOL 慣例)，用每列的 USD 合計勾稽；VAT 是越南盾 (單獨一筆)。發票號碼 = 每列的
# 發票號碼 SỐ HĐ (沒有才用 HAWB/Bill No)，HAWB 填 B/L NO.。
_VT_OLM_FEES_CN = [(x, None) for x in (
    "OCEAN FREIGHT", "OTHER CHARGES", "CFS FEE", "=CIC", "D/O FEE", "=HANDLING FEE", "CLEANING FEE",
    "=THC", "=CMC", "=LSS", "=LCC", "=BAF", "AFR AMENDMENT", "=HANDLING",
    # 空運進口 (2026-10-28 抽樣 "POU HUNG 06-2026 PHƯƠNG ĐÔNG air import")
    "=AIR FREIGHT")]
_VT_ORIENTAL = [
    # BANGKE EXP (2026-10-28 抽樣 "BANGKE T2-2026 EXP"、"BANGKE T5/T6-2026 EXP")：費用欄同 IMP 版合併成
    # "VOL (CBM) THC"…，但 TOTAL VND = TOTAL USD × 匯率 (不是另外的越南盾費用)，只用 TOTAL USD 勾稽
    (r"(?!)", {"header": "VOL (CBM)", "require": "VOL (CBM) BORDER GATE FEE", "key": ["SO HD", "H-B/L"],
               "inherit_blank": True, "einv_key": True,
               "fees": [], "fees_re": r"^VOL \(CBM\) (?!CY QTY)(.+)$", "fees_skip": r"CHECK|\bVND\b",
               "fee_cur": "USD", "skip_sheet": r"^(LIST|BK)\b",
               "total": "TOTAL USD", "date": "NGAY HD", "cbm": "=VOL (CBM)",
               "fields": {"bl_no": "H-B/L", "mbl_no": "M-B/L", "order_no": ("LOT NO", "BOOKING NO"),
                          "port_of_loading": "=PORT"}}),
    # BANGKE IMP (2026-10-28 抽樣，工作表 "T6-2026")：三列表頭，費用欄合併後是 "VOL (CBM) D/O"…；
    # 費用 USD (標題寫 VND 的是越南盾)，用 TOTAL USD + TOTAL VND 勾稽；發票號碼 SỐ HĐ
    (r"(?!)", {"header": "VOL (CBM)", "require": "VOL (CBM) D/O", "key": ["SO HD", "H-B/L"], "inherit_blank": True,
               "einv_key": True,
               "fees": [], "fees_re": r"^VOL \(CBM\) (?!CY QTY)(.+)$", "fees_skip": r"CHECK", "fee_cur": "USD",
               # "list 1-8%"、"list 2-0%"、"BK 8%" 是主表依稅率拆開的副本，不取
               "skip_sheet": r"^(LIST|BK)\b",
               "total": "TOTAL USD", "total_add": "TOTAL VND", "date": "NGAY HD", "cbm": "=VOL (CBM)",
               "fields": {"bl_no": "H-B/L", "mbl_no": "M-B/L", "order_no": ("LOT NO", "BOOKING NO"),
                          "port_of_loading": "=PORT"}}),
    # 2024 PAYMENT FOR LOCAL CHARGE：Invoice No | B/L | 各項費用 (USD) | Total | VAT | Sum Total (USD/VND)
    (r"(?!)", {"header": "SUM TOTAL", "two_rows": False, "require": "SUM TOTAL", "key": ["INVOICE NO", "=B/L"],
               "fees": [], "fees_between": ("=B/L", ("=TOTAL",)), "fee_cur": "USD",
               "vat": "=VAT", "vat_desc": "VAT (USD)", "total": "=TOTAL", "grand": "SUM TOTAL (USD)",
               "date": "VAT INV DATE", "fields": {"bl_no": "=B/L"}}),
    # 每月發票清單 (ORIENTEL 2026-06：No | Invoice date | Invoice No | Bill No | Amount)
    (r"(?!)", {"header": "BILL NO", "two_rows": False, "require": "=AMOUNT", "key": ["INVOICE NO"],
               "einv_key": True, "fees": [("=AMOUNT", "Amount")], "date": "INVOICE DATE",
               "sum_row": r"^TOTAL", "fields": {"bl_no": "BILL NO"}}),
    (r"^CN", {"header": "BILL NO", "above": True, "key": ["SO HD", "BILL NO"], "fees": _VT_OLM_FEES_CN,
              "einv_key": True,
              "fee_cur": "USD", "ffill_date": True,   # 同一批第二張發票的 NGÀY HĐ 合併在上一列
              "total": "TOTAL (USD)", "vat": "=VAT", "vat_desc": "VAT", "date": "NGAY HD",
              "boxes": {"20": "CY QTY 20'", "40": "CY QTY 40'"}, "cbm": "CFS CBM",
              "fields": {"bl_no": "BILL NO", "order_no": "LOT NO", "port_of_loading": "ORIGIN PORT"}}),
    (r"TRUCK", {"header": "H-B/L", "key": ["SO HD", "H-B/L", "M-B/L"], "einv_key": True,
                "fees": [(x, None) for x in ("TRUCKING FEE", "BORDER GATE FEE", "CUSTOMS CLEARANCE FEE",
                                             "TOLL FEE", "LOLO FEE", "FUEL SURCHARGE", "FCR TRUCKING")],
                "total": "TOTAL USD", "total_add": "TOTAL VND", "vat": "=VAT", "vat_desc": "VAT",
                "date": "NGAY HD", "boxes": {"20": "CY QUANTITY 20", "40": "CY QUANTITY 40"}, "cbm": "CFS CBM",
                "fields": {"bl_no": "H-B/L", "mbl_no": "M-B/L", "order_no": "LOT NO",
                           "port_of_loading": "ORIGIN PORT"}}),
    (r"SEA", {"header": "H-B/L", "key": ["SO HD", "H-B/L", "M-B/L"], "einv_key": True,
              "fees": [(x, None) for x in (
                  "THC VND", "THC USD", "=EBS", "EFS USD", "CFS VND", "CFS USD", "=TRUCKING", "=LOCAL CHARGE",
                  "TCS CHARGE", "=AMS", "=BAF", "=EDI", "=SCANNING", "=RR", "=LSS", "OCEAN FREIGHT",
                  "AIR FREIGHT", "HANDLING FEE VND", "HANDLING FEE USD", "FUMIGATION VND",
                  "FUMIGATION USD", "SEAL FEE", "TELEX FEE", "BILL FEE",
                  # 2026-10-28 抽樣 "POUYUEN T6-2026 EXP"：另有一欄只寫 "SEAL" (USD)
                  "=SEAL",
                  # FUMIGATION 合併儲存格底下有兩欄 USD，第二欄也有金額且算進合計
                  # (2026-10-28 抽樣 "POUYUEN T4-2026 EXP-PHƯƠNG ĐÔNG"：每列 32.4)
                  "$FUMIGATION USD",
                  # 表頭拼錯 "Hanling fee" (USD) (2026-10-28 抽樣 "BANGKE T4-2026 EXP" SEA (2) 表)
                  "HANLING FEE",
                  # 2026-10-28 抽樣 "BANGKE T6-2026 EXP"：LOLO、DECLARATION，與第二個 LOCAL CHARGE 欄
                  "=LOLO", "=DECLARATION", "$LOCAL CHARGE",
                  # 沒寫幣別的 EFS 欄 (2026-10-28 抽樣 "POU HUNG 03-2026 PHƯƠNG ĐÔNG export" CN T 03 表)
                  "=EFS")],
              "total": "TOTAL USD", "total_add": "TOTAL VND", "vat": "=VAT", "vat_desc": "VAT",
              "date": "NGAY HD", "boxes": {"20": "CY QUANTITY 20", "40": "CY QUANTITY 40"}, "cbm": "CFS CBM",
              "fields": {"bl_no": "H-B/L", "mbl_no": "M-B/L", "order_no": "LOT NO",
                         "port_of_loading": "NOTE PORT"}}),
]


def detect_oriental_multimodal(text: str) -> bool:
    norm = _vt_norm(text)
    # 2026-10-28 抽樣：有的對帳表只寫越南文名稱 (… ĐA PHƯƠNG THỨC PHƯƠNG ĐÔNG) 或只有發票符號 C26TOD
    return ("ORIENTAL LOGISTICS MULTIMODA" in norm or "DA PHUONG THUC PHUONG DONG" in norm
            or "VAN TAI DA PHUONG THUC" in norm or re.search(r"\d-?C\d{2}TOD\b", norm) is not None)


# PDF (2026-10-28 抽樣)：
#  (a) DEBIT NOTE：費用行 "THC 9.000 / CBM x 5.400 8 USD 52.49" (單價/單位 × 數量、稅率、幣別、含稅金額)，
#      未稅 = 單價 × 數量，VAT = 差額；表尾 "AMOUNT: USD 238.14" 是含稅合計。
#  (b) 電子發票 (HÓA ĐƠN GIÁ TRỊ GIA TĂNG, Serial C26TOD)：費用行 "1 Cước vận tải quốc tế (O/F)
#      XMNSGN26050461 2 40'HQ 17.076.150 34.152.300"，表尾 Sub total / VAT / Total payment。
_OLM_DN_ITEM_RE = re.compile(
    r"^(?P<desc>.+?)\s+(?P<price>[\d.,]+)\s*/\s*(?P<unit>\S+)\s+x\s+(?P<qty>[\d.,]+)\s+(?P<rate>\d+)\s+"
    r"(?P<cur>[A-Z]{3})\s+(?P<amt>[\d.,]+)$", re.MULTILINE)
_OLM_EINV_ITEM_RE = re.compile(
    r"^\d+\s+(?P<desc>.+?)\s+(?P<bl>[A-Z]{4,}\d{6,}\S*)\s+(?P<qty>[\d.,]+)\s+(?P<unit>\S+)\s+"
    r"(?P<price>[\d.]+)\s+(?P<amt>[\d.]+)$", re.MULTILINE)


def _olm_num(text: str) -> float:
    return float(text.replace(",", ""))


def parse_oriental_pdf(text: str) -> Dict:
    h: Dict[str, Optional[str]] = {"supplier": "ORIENTAL LOGISTICS MULTIMODAL"}
    items: List[Dict] = []
    if "DEBIT NOTE" in text.upper() and _OLM_DN_ITEM_RE.search(text):
        h["invoice_no"] = _search(r"REF NO\.\s*:\s*(\S+)", text)
        h["invoice_date"] = _vn_pdf_date(_search(r"ISSUE DATE\s*:\s*(\S+)", text))
        h["consignee"] = _search(r"NAME & ADDRESS:[^\n]*\n(.+?)\s+ISSUE DATE", text)
        h["vessel"] = _search(r"VESSEL/VOYAGE\s*:\s*(.+?)\s+ONBOARD DATE", text)
        h["onboard_date"] = _vn_pdf_date(_search(r"ONBOARD DATE\s*:\s*(\S+)", text))
        m = re.search(r":\s*([A-Z][A-Z .,]+?)/\d{2}/\d{2}/\d{4}\s*:\s*(.+)$", text, re.MULTILINE)
        if m:
            h["port_of_loading"], h["port_of_discharge"] = m.group(1).strip(), m.group(2).strip()
        h["mbl_no"] = _search(r"MBL NO\s*:\s*(?!DESTINATION)(\S+)", text)   # MBL 空白時後面是 DESTINATION
        h["bl_no"] = _search(r"HBL INFO:\s*(\S+)", text)
        h["container_no"] = _search(r"CONTAINER:\s*([A-Z]{4}\d{7})", text)
        units = set()
        for m in _OLM_DN_ITEM_RE.finditer(text):
            net = round(_olm_num(m.group("price")) * _olm_num(m.group("qty")), 2)
            amt = _olm_num(m.group("amt"))
            cur = m.group("cur")
            units.add(m.group("unit").upper())
            items.append({"description": m.group("desc").strip() + ("" if cur == "VND" else f" ({cur})"),
                          "amount": net, "vat": round(amt - net, 2) or None})
        # 材積：按 CBM 計價的是散貨 (CBM)，按櫃計價的是整櫃 (櫃型)
        cbm = _search(r"HBL INFO:.*?([\d.]+)\s*CBM", text)
        boxes = re.findall(r"/(\d{2}'[A-Z]{2})\s", text) if "CBM" not in units else []
        h["volume"] = _fmt_volume(containers=[(1, b) for b in boxes], cbm=cbm)
    else:
        h["invoice_no"] = _search(r"S[ốo] \(Invoice No\.\):\s*(\d+)", text)
        h["invoice_date"] = _vn_pdf_date(_search(r"\(Date\):\s*(\S+)", text))
        h["consignee"] = _search(r"\(Company name\):\s*(.+)", text)
        m = re.search(r"HÀNH TRÌNH:\s*(.+?)\s+-\s+(.+?)\s*/\s*\n?\s*(.+)", text)
        if m:
            h["port_of_loading"], h["port_of_discharge"], h["vessel"] = (x.strip() for x in m.groups())
        boxes = []
        for m in _OLM_EINV_ITEM_RE.finditer(text):
            h.setdefault("bl_no", m.group("bl"))
            if re.match(r"\d{2}'", m.group("unit")):
                boxes.append((m.group("qty"), m.group("unit")))
            items.append({"description": m.group("desc").strip(),
                          "amount": _dsv_vn_dot_amount(m.group("amt")), "vat": None})
        vat = _dsv_vn_dot_amount(_search(r"\(VAT amount\):\s*([\d.]+)", text))
        rate = _search(r"\(Tax rate\):\s*(\d+%)", text)
        if vat:
            items.append({"description": f"VAT {rate}" if rate else "VAT", "amount": None, "vat": vat})
        # 同一櫃型在好幾行都列一次 (O/F、THC…)：取最大的數量
        best: Dict[str, float] = {}
        for qty, unit in boxes:
            best[unit] = max(best.get(unit, 0), float(qty.replace(",", ".")))
        h["volume"] = _fmt_volume(containers=list((q, u) for u, q in best.items()))
    return {"header": h, "items": items or [{"description": None, "amount": None, "vat": None}]}


def _olm_pdf_total(text: str):
    if "DEBIT NOTE" in text.upper():
        return None
    return _dsv_vn_dot_amount(_search(r"\(Sub total\):\s*([\d.]+)", text))


def _olm_pdf_grand(text: str):
    if "DEBIT NOTE" in text.upper():
        v = _search(r"^AMOUNT:\s*[A-Z]{3}\s*([\d.,]+)", text)
        return _olm_num(v) if v else None
    return _dsv_vn_dot_amount(_search(r"\(Total payment\):\s*([\d.]+)", text))


# 工作表叫 "CN T 03" 但版面是 SEA 對帳表 (表頭 H-B/L，沒有 BILL NO)：放在 CN 設定前面，
# 用必要欄位區分 (2026-10-28 抽樣 "POU HUNG 03-2026 PHƯƠNG ĐÔNG export")
_VT_ORIENTAL.insert(next(i for i, (p, _c) in enumerate(_VT_ORIENTAL) if p == r"^CN"),
                    (r"^CN", {**next(c for p, c in _VT_ORIENTAL if p == r"SEA"), "require": "H-B/L"}))
_olm_sheet_parse, _olm_sheet_totals = _vt_make("ORIENTAL LOGISTICS MULTIMODAL", _VT_ORIENTAL)
register_supplier(
    "ORIENTAL_MULTIMODAL", "CÔNG TY TNHH LOGISTICS VẬN TẢI ĐA PHƯƠNG THỨC (ORIENTAL)",
    detect_oriental_multimodal, parse_oriental_pdf,
    extract_total=_olm_pdf_total, extract_grand_total=_olm_pdf_grand,
    sheet_parse=_olm_sheet_parse, sheet_totals=_olm_sheet_totals,
)


# --- VN16 SONG TRIỀU (拖車行，Pou Li Tây Ninh) ---------------------------------
# 四張運費對帳表 (散貨車、轉倉貨櫃、出口櫃、進口櫃)。運費 = 未調漲運費 + 油價調漲
# 金額；另有裝卸、壓車、吊櫃上下、過磅等費用。發票號碼 = 批號 SỐ LÔ / SỐ INV。
_VT_ST_COMMON = {"date": "=NGAY", "port": None, "cbm": "SO LUONG CBM",
                 "boxes": {"20": "LOAI CONT 20'", "40": "LOAI CONT 40'", "45": "LOAI CONT 45'"},
                 "total": "=TONG", "vat_row": True,
                 "fields": {"vessel": "SO XE", "container_no": ("SO CONTAINER", "SO XE"),
                            "port_of_loading": ("DIEM DI", "DIEM DI"), "port_of_discharge": "DIEM DEN"}}
_VT_SONG_TRIEU = [
    (r"XE TAI", {**_VT_ST_COMMON, "header": "SO XE", "key": ["SO INV"],
                 "fees": [("PHI BOC", None), ("NEO XE", None), ("PHAT SINH", None),
                          ("CUOC VAN CHUYEN CHUA TANG", None), ("SO TIEN % TANG", None)]}),
    (r"CHUYEN KHO", {**_VT_ST_COMMON, "header": "SO CONTAINER", "key": ["SO INV", "SO CONTAINER"],
                     "fees": [("TROUBLE", None), ("=NEO", None), ("CUOC VAN CHUYEN CHUA TANG", None),
                              ("SO TIEN % TANG", None)]}),
    (r"CONT (XUAT|NHAP)", {**_VT_ST_COMMON, "header": "SO CONTAINER", "key": ["=SO LO", "SO CONTAINER"],
                           "fees": [("CUOC VC CHUA TANG", None), ("SO TIEN % TANG", None),
                                    ("BD NANG", None), (("PHI KHAC PHI", "PHI KHAC SO TIEN"), None),
                                    (("PHI NANG/HA NANG", "PHI NANG / HA NANG"), None),
                                    (("PHI NANG/HA HA", "PHI NANG / HA HA"), None), ("PHI CAN CAN", None),
                                    # 沖銷 (TỔNG 要減掉，2026-10-28 抽樣 T7.2026)；名稱前的 "-" 代表減項
                                    ("CAN TRU", "-Cấn trừ phí phát sinh")]}),
]


def detect_song_trieu(text: str) -> bool:
    norm = _vt_norm(text)
    return "SONG TRIEU" in norm or "033 1003828 194" in norm


_vt_register("SONG_TRIEU", "CÔNG TY TNHH SONG TRIỀU", detect_song_trieu,
             "CÔNG TY TNHH SONG TRIỀU", _VT_SONG_TRIEU)


# --- VN17 TIMES CARGO LOGISTIC (CÔNG TY TNHH DỊCH VỤ GIAO NHẬN HÀNG HÓA) ------
# Statement of account (SEA - IMPORT)：每列一票，有自己的 Invoice No.，金額是越南盾
# (已用 Rate Exchange 換算)；合計列 "TOTAL BALANCE"。發票號碼 = Invoice No.。
_VT_TIMES_CARGO = [
    # POU SUNG 的 STATEMENT OF ACCOUNT (2026-10-28 抽樣)：No | Service | Quantity | HBL | ETA | Invoice No. |
    # Invoice Date | Form | Serial | Description (各費用) | VAT | Total，每列一張電子發票
    (r"(?!)", {"header": "SERIAL", "require": "SERIAL", "key": ["INVOICE NO", "=HBL"],
               "fees": [], "fees_between": ("SERIAL", ("VAT",)), "vat": "VAT", "vat_desc": "VAT 8%",
               "grand": "$TOTAL", "date": "INVOICE DATE",
               "fields": {"bl_no": "=HBL", "arrive_date": "=ETA"}}),
    (r".", {"header": "HBL", "key": ["=INVOICE NO.", "=HBL"], "name_require": "RATE EXCHANGE",
            "fees": [(x, None) for x in (
                "OCEAN FREIGHT", "DESCRIPTION EXW", "THC FEE", "DO FEE", "HANDLING FEE", "CIC FEE",
                "LSS FEE", "CFS FEE", "LOADING FEE", "DESCRIPTION EBS", "DRAYAGE")],
            # 每個月的費用欄不一樣 (AIR FREIGHT、CLEANING FEE、OTHER CHARGE…)：INVOICE DATE 與 TOTAL
            # 之間的欄都是費用 (2026-10-28 抽樣)
            "fees_between": ("=INVOICE DATE", ("=TOTAL",)),
            "total": "=TOTAL", "date": "=INVOICE DATE", "cbm": "=QUANTITY",
            "fields": {"bl_no": "=HBL", "shipper": "=SHIPPER", "port_of_loading": "PLACE OF RECEIPT",
                       "port_of_discharge": "PLACE OF DELIVERY", "vessel": "VESSEL",
                       "arrive_date": "=ETA"}}),
]


def detect_times_cargo(text: str) -> bool:
    norm = _vt_norm(text)
    # 電子發票只有越南文名稱 (… THỜI GIAN) 與稅號 0309536491
    return ("TIMES CARGO LOGISTIC" in norm or "0309536491" in norm or "GIAO NHAN HANG HOA THOI GIAN" in norm
            # 沒有公司名稱的 STATEMENT OF ACCOUNT：發票符號 1C25TAA
            or ("STATEMENT OF ACC" in norm and re.search(r"\dC\d{2}TAA\b", norm) is not None))


# PDF (2026-10-28 抽樣)：
#  (a) 到貨通知 (THÔNG BÁO HÀNG ĐẾN / Arrival Notice)：費用行 "☑ CFS 2.300 CBM 35.00 80.50 (USD)"
#      (名稱、數量、單位、單價、金額、幣別)，"Total Charges 333.00 (USD)"，費用不含 VAT。
#  (b) 電子發票：費用行 "1 Cước vận chuyển quốc tế bằng đường biển Khối 1 2.366.100 2.366.100"
#      (名稱、單位、數量、單價、金額)，表尾 Cộng tiền hàng / Tiền thuế GTGT / Tổng cộng tiền thanh toán。
_TCL_AN_ITEM_RE = re.compile(
    r"^\W*(?P<desc>[A-Za-z].*?)\s+(?P<qty>[\d.,]+)\s+(?P<unit>\S+)\s+(?P<rate>[\d.,]+)\s+(?P<amt>[\d.,]+)\s+"
    r"\((?P<cur>[A-Z]{3})\)$", re.MULTILINE)
_TCL_EINV_ITEM_RE = re.compile(
    r"^\d+\s+(?P<desc>.+?)\s+(?P<unit>\S+)\s+(?P<qty>[\d.,]+)\s+(?P<price>[\d.]+)\s+(?P<amt>[\d.]+)$", re.MULTILINE)


def parse_times_cargo_pdf(text: str) -> Dict:
    h: Dict[str, Optional[str]] = {"supplier": "TIMES CARGO LOGISTIC CO., LTD"}
    items: List[Dict] = []
    if "Arrival Notice" in text:
        # 冒號有時印在上一行 (":" / "Số (No.) IM-LSEA507225")
        h["invoice_no"] = _search(r"Số \(No\.\)\s*:?\s*(\S+)", text)
        h["invoice_date"] = _vn_pdf_date(_search(r"\(Time\)\s*:?\s*\(?(\d{1,2}/\d{1,2}/\d{4})", text))
        h["consignee"] = _search(r"\(TO\):\s*(.+)", text)
        h["shipper"] = _search(r"\(Shipper\)\s*:\s*(.+)", text)
        m = re.search(r"\(Vessel\)\s*:\s*(.+?)\s+Chuyến \(Voy\.\)\s*:\s*(\S+)", text)
        h["vessel"] = f"{m.group(1)} / {m.group(2)}" if m else None
        m = re.search(r"\(From\)\s*:\s*(.+?)\s+Đến cảng \(Terminal\)\s*:\s*(.+)", text)
        if m:
            h["port_of_loading"], h["port_of_discharge"] = m.group(1).strip(), m.group(2).strip()
        m = re.search(r"\(Estimated Arrival\)\s*:\s*([A-Za-z]{3}) (\d{1,2}), (\d{4})", text)
        h["arrive_date"] = _format_date_en_from_month_name(m.group(2), m.group(1), m.group(3)) if m else None
        h["mbl_no"] = _search(r"\(M-B/L\)\s*:\s*(\S+)", text)
        h["bl_no"] = _search(r"\(H-B/L\)\s*:\s*(\S+)", text)
        h["container_no"] = _search(r"^([A-Z]{4}\d{7})/", text)
        # CFS/CFS 是散貨 (CBM)；整櫃寫 "(01)X40'HC"
        cbm = _search(r"([\d.]+)\s*CBM\s*$", text)
        box = re.search(r"\((\d+)\)\s*X\s*(\d{2}'?\w*)", text)
        lcl = "CFS/CFS" in text or "A PART OF" in text
        h["volume"] = _fmt_volume(containers=[] if lcl or not box else [(box.group(1), box.group(2))], cbm=cbm)
        for m in _TCL_AN_ITEM_RE.finditer(text):
            cur = m.group("cur")
            items.append({"description": m.group("desc").strip() + ("" if cur == "VND" else f" ({cur})"),
                          "amount": float(m.group("amt").replace(",", "")), "vat": None})
    else:
        h["invoice_no"] = (_search(r"^Số:\s*(\d+)", text) or "").lstrip("0") or None
        m = re.search(r"Ngày (\d{1,2}) tháng (\d{1,2}) năm (\d{4})", text)
        h["invoice_date"] = _format_date_en(*m.groups()) if m else None
        h["consignee"] = _search(r"^Tên đơn vị:\s*(.+)", text)
        h["bl_no"] = _search(r"B/L#:\s*(\S+)", text)
        m = re.search(r"B/L#:[^\n]*\n(.+?),?\s*-\s*(.+)\n(.+)", text)
        if m:
            h["port_of_loading"], h["port_of_discharge"], h["vessel"] = (x.strip() for x in m.groups())
        for m in _TCL_EINV_ITEM_RE.finditer(text):
            items.append({"description": m.group("desc").strip(),
                          "amount": _dsv_vn_dot_amount(m.group("amt")), "vat": None})
        vat = _dsv_vn_dot_amount(_search(r"Tiền thuế GTGT:\s*([\d.]+)", text))
        rate = _search(r"Thuế suất GTGT:\s*(\d+%)", text)
        if vat:
            items.append({"description": f"VAT {rate}" if rate else "VAT", "amount": None, "vat": vat})
    return {"header": h, "items": items or [{"description": None, "amount": None, "vat": None}]}


def _tcl_pdf_total(text: str):
    if "Arrival Notice" in text:
        v = _search(r"Total Charges\s+([\d.,]+)", text)
        return float(v.replace(",", "")) if v else None
    return _dsv_vn_dot_amount(_search(r"Cộng tiền hàng:\s*([\d.]+)", text))


def _tcl_pdf_grand(text: str):
    if "Arrival Notice" in text:
        return None
    return _dsv_vn_dot_amount(_search(r"Tổng cộng tiền thanh toán:\s*([\d.]+)", text))


_tcl_sheet_parse, _tcl_sheet_totals = _vt_make("TIMES CARGO LOGISTIC CO., LTD", _VT_TIMES_CARGO)
register_supplier(
    "TIMES_CARGO", "CÔNG TY TNHH DỊCH VỤ GIAO NHẬN HÀNG HÓA (TIMES CARGO)", detect_times_cargo,
    parse_times_cargo_pdf, extract_total=_tcl_pdf_total, extract_grand_total=_tcl_pdf_grand,
    sheet_parse=_tcl_sheet_parse, sheet_totals=_tcl_sheet_totals,
)


# --- VN18 DỊCH VỤ VẬN CHUYỂN SÀI GÒN (空運) -------------------------------------
# 每列一票：空運費 A/F (USD 與換算後 VND)、Local Charge、Break Bulk (都是 VND)，
# 各有自己的紅色發票號碼 (Red invoice no.)。使用者 2026-10-27：發票號碼用 Red invoice no.
# (每項費用取自己右邊那一欄；空白才用 HAWB)。日期欄是
# "26/08/2026 - 146 KGS" 這種文字。沒有逐列合計，不做逐張勾稽。
_VT_SAIGON = [
    # OBC 快遞的 BẢNG KÊ CÔNG NỢ (2026-10-28 抽樣，工作表 "POU SUNG-寶崧-NHAP")：每列一張提單 (Số Bill)，
    # THÀNH TIỀN VND 是金額 (已含其他費用)
    (r"(?!)", {"header": "SO BILL", "require": "THANH TIEN VND", "key": ["SO BILL"],
               "fees": [("THANH TIEN VND", "Cước vận chuyển (OBC)")], "date": "NGAY", "kg": "SO KG",
               "sum_row": r"^(TONG|TOTAL|THANH TIEN$)", "port": None,
               "fields": {"bl_no": "SO BILL", "port_of_loading": "CHUYEN DI"}}),
    (r".", {"header": "HAWB", "key": ["HAWB NO"], "name_require": "HAWB NO",
            "fees": [("AMOUNT (VND)", "Air freight (A/F)", "RED INVOICE NO"),
                     ("LOCAL CHARGE (VND)", "Local Charge", "RED INVOICE NO"),
                     ("BREAK BULK", "Break Bulk", "RED INVOICE NO")],
            "date": ("INV DATE", "DATE INV"), "kg": "CHARGEABLE WEIGHT",
            "fields": {"bl_no": "HAWB NO", "mbl_no": "MAWB NO"}}),
]


def detect_saigon_transport(text: str) -> bool:
    norm = _vt_norm(text)
    # 舊字型抬頭讀不出公司名稱：用 MST 0301440413 或表頭組合 (Hawb No + Red Invoice No + Local Charge)
    return ("DICH VU VAN CHUYEN SAI GON" in norm or "0301440413" in norm
            or all(k in norm for k in ("HAWB NO", "RED INVOICE", "LOCAL CHARGE (VND)")))


_vt_register("SAIGON_TRANSPORT", "CÔNG TY TNHH DỊCH VỤ VẬN CHUYỂN SÀI GÒN", detect_saigon_transport,
             "CÔNG TY TNHH DỊCH VỤ VẬN CHUYỂN SÀI GÒN", _VT_SAIGON)


# --- VN19 IMX EXPRESS (CÔNG TY TNHH DỊCH VỤ CHUYỂN PHÁT NHANH IMX，MST 0315868311) --
# 兩張對帳表：D/O FEE (USD 與 VND 兩欄，取 VND) 與 AIR FREIGHT (A/F、FUEL、P/U、HD 都是
# USD，名稱加 "(USD)"，用 TOTAL USD 勾稽)。使用者 2026-10-27 要求 USD 與 VND 都稽核：
# D/O FEE (VND) 用表尾 TOTAL (VND) / SUBTOTAL (VND) 勾稽，表尾 "V.A.T TAX (8%)" 單獨一筆。
# 發票號碼 = 對帳表號 (Mã bảng kê，每張表附一張發票)，AWB.No 填 B/L NO.。
_VT_IMX_COMMON = {"key": ["AWB.NO"], "date": "=DATE", "kg": "WEIGHT",
                  "doc_key": r"M[ãa]\s*b[ảa]ng\s*k[êe]\s*:\s*([A-Z0-9]+)",
                  "fields": {"bl_no": "AWB.NO", "shipper": "=SHIPPER", "vessel": "FLIGHT"}}
_VT_IMX = [
    (r"D\.?O FEE", {**_VT_IMX_COMMON, "header": "AWB.NO", "fees": [("D/O FEE VND", "D/O FEE")],
                    "sheet_total": ("TOTAL (VND)", "SUBTOTAL (VND)"), "vat_row": True}),
    (r"AIR FREIGHT", {**_VT_IMX_COMMON, "header": "AWB.NO",
                      "fees": [("A/F CHARGE", None), ("FUEL", None), ("=P/U USD", None), ("=HD USD", None)],
                      "total": "=TOTAL USD"}),
]


def detect_imx(text: str) -> bool:
    norm = _vt_norm(text)
    # 舊檔 (2025) 沒有公司名稱：發票符號 1C25TMX，或「Mã bảng kê」+ AWB.No + D/O FEE
    return ("IMX EXPRESS" in norm or "0315868311" in norm or re.search(r"\dC\d{2}TMX\b", norm) is not None
            or ("MA BANG KE" in norm and "AWB.NO" in norm and "D/O FEE" in norm))


_vt_register("IMX", "CÔNG TY TNHH DỊCH VỤ CHUYỂN PHÁT NHANH IMX", detect_imx,
             "CÔNG TY TNHH DỊCH VỤ CHUYỂN PHÁT NHANH IMX", _VT_IMX)


# --- VN15 EXPEDITORS 的 Excel 對帳表 (AIR IMP) -----------------------------------
# 「單價」(11~17 欄) 與「金額」(20~26 欄) 兩組費用欄，只取金額；另有運費 CƯỚC、
# 代收代付 THU HỘ ĐẠI LÝ NƯỚC NGOÀI、VAT、合計。發票號碼 = 每列的 TAX INVOICE NUMBER。
# PDF 規則在上面 (parse_expeditors)；這裡重新註冊一次，把 Excel 規則加上去。
_LH = "LAM HANG - (THUE 8%)"
_VT_EXPEDITORS = [
    (r".", {"header": "EI REF", "key": ["TAX INVOICE NUMBER", "VAN DON PHU"],
            "fees": [(f"{_LH} PHI LENH", None), (f"{_LH} PHI XEP DO", None), (f"{_LH} PHI PHU TROI", None),
                     (f"{_LH} PHI VE SINH", None), (f"{_LH} PHI PCS", None), (f"{_LH} PHI E-MANIFEST", None),
                     (f"{_LH} PHI HANDLING", None), ("=CUOC", "CƯỚC"),
                     ("=THU HO DAI LY NUOC NGOAI", "THU HỘ ĐẠI LÝ NƯỚC NGOÀI")],
            "vat": "=VAT", "vat_desc": "VAT 8%", "grand": "TONG CONG (VND)", "date": "TAX INVOICE DATE",
            "boxes": {"20": "KHOI LUONG 20'", "40": "KHOI LUONG 40'"}, "kg": "CBM/C.W",
            "fields": {"mbl_no": "VAN DON CHINH", "bl_no": "VAN DON PHU", "order_no": "EI REF",
                       "arrive_date": "NGAY DEN", "port_of_loading": "NOI DI", "port_of_discharge": "NOI DEN"}}),
]
_exp_sheet_parse, _exp_sheet_totals = _vt_make("CÔNG TY TNHH EXPEDITORS VIỆT NAM", _VT_EXPEDITORS)


def detect_expeditors(text: str) -> bool:      # noqa: F811  (加上 Excel 對帳表)
    upper = text.upper()
    return ("EXPEDITORS VIETNAM COMPANY LIMITED" in upper and "TAX No:" in text) or (
        "EXPEDITORS VIET" in upper and "EI REF" in upper)


register_supplier(
    "EXPEDITORS", "EXPEDITORS VIETNAM COMPANY LIMITED", detect_expeditors, parse_expeditors,
    multi_page=True, extract_total=_exp_extract_total, extract_grand_total=_exp_extract_grand_total,
    sheet_parse=_exp_sheet_parse, sheet_totals=_exp_sheet_totals,
)


# --- VN05 CONG THANH 的 POU CHEN 對帳表 (2026-10-28 抽樣，工作表 xuât HĐ / P.CHEN / PHI CH / luu cont) --
# "P.CHEN …"：每票一列 (INVOICE # = 寶成出貨單號)，PHÍ DỊCH VỤ、PHÍ VẬN TẢI、PHÍ LƯU CONT，
#   VAT 0%，表尾 TOTAL 列勾稽整份 (使用者：運費表核對整份合計)。拖車行：車牌照 CONG THANH
#   慣例填船名與貨櫃號碼，港口 VIETNAM。
# "PHI …"：代墊吊櫃費，每列一張第三方紅色發票 (RED INV. NO 當發票號碼)，未稅 + VAT = AMOUNT。
_VT_CT_POUCHEN = [
    (r"^P\.?\s?CHEN", {
        "header": "INVOICE #", "two_rows": False, "key": ["INVOICE #"],
        "fees": [("PHI DICH VU", None), ("PHI VAN TAI", None), ("PHI LUU CONT", None)],
        "sum_row": r"^TOTAL", "date": "DATE OF SHIP OUT", "cbm": "=CBM", "kg": "WEIGHT",
        "boxes": {"45": "=45'", "40": "=40'", "20": "=20'"},
        # 這個版面就是寶成 POU CHEN 的對帳表 (工作表名稱 P.CHEN)
        "const": {"port_of_loading": "VIETNAM", "port_of_discharge": "VIETNAM", "consignee": "POU CHEN"},
        "ffill_date": True,
        # FCRB/L 欄寫的是 "EFL/SCSC"、"DONG NAI PORT" 這類地點，不當提單號
        "fields": {"order_no": "ORDER NO", "vessel": "BIEN SO", "container_no": "BIEN SO"}}),
    (r"^PHI", {
        "header": "DISCRIPTION", "key": ["RED /INV. NO", "RED INV"],
        "fees": [("EXCLUDING", "@DISCRIPTION")], "vat": "=VAT", "vat_desc": "VAT",
        "grand": "=AMOUNT", "date": "ISSUED RED", "ffill_date": True,
        "const": {"port_of_loading": "VIETNAM", "port_of_discharge": "VIETNAM", "consignee": "POU CHEN"},
        "fields": {"order_no": "INV. NO", "container_no": "CONT./ TRUCK"}}),
]
_ct_pc_parse, _ct_pc_totals = _vt_make("CTY TNHHDV GN VT VA TM CONG THANH", _VT_CT_POUCHEN)


def _ct_is_pouchen(book) -> bool:
    return any(re.match(r"^P\.?\s?CHEN", _vt_norm(n)) for n, _ in book.sheets)


def detect_cong_thanh_any(text: str) -> bool:
    # 2025 年的 POU CHEN 對帳表抬頭沒有 CONG THANH 稅號：用壓車表標題辨識
    return detect_cong_thanh(text) or "SAVED CHASSIS TRAILER LIST" in text.upper()


register_supplier(
    "CONG_THANH", "CTY TNHHDV GN VT VA TM CONG THANH", detect_cong_thanh_any, None,
    sheet_parse=lambda book: _ct_pc_parse(book) if _ct_is_pouchen(book) else parse_cong_thanh_sheets(book),
    sheet_totals=lambda book: _ct_pc_totals(book) if _ct_is_pouchen(book) else _cong_thanh_totals(book),
)


# --- VN03 RỒNG Á CHÂU 的 ADC 對帳表 (ASIA DRAGON FORWARDING，寶成 POU CHEN，2026-10-28 抽樣) --
# 工作表 "FCL/LCL/AIR POU CHEN"：每批 (Lot No) 一列，費用分欄 (報關、服務、拖車、吊櫃、
# 當地費用、倉租…)，沒有 VAT；表頭在第一列 (下一列就是資料)。"TẠI CHỖ"：每張報關單一筆
# 服務費。"RF" 是代墊費明細 (主表已有同樣的金額)，不取。沒有逐列合計，只有表尾欄合計
# (由 excel_formula_issues 檢查 SUM 範圍)。
_VT_ADC = [
    (r"(FCL|LCL|AIR).*POU", {
        "header": "DECLERATION NO", "two_rows": False, "key": ["=LOT NO"], "carry": True,
        "fees": [(x, None) for x in ("EXTRA CUSTOMS SHEET", "SERVICE FEE", "TRUCKING FEE", "CLEARING EXPENSE",
                                     "LIFT ON FEE", "LIFT OFF FEE", "LOCAL CHARGE", "STORAGE FEE", "DEMURRAGE",
                                     "EXTRA FEE")],
        "date": "DELIVERY", "boxes": {"20": "=20'", "40": "=40'"}, "cbm": "=CBM", "kg": "KGS OF",
        "fields": {"bl_no": "AWB/ B/L NO", "order_no": "=INVOICE NO", "container_no": "CONT NO"}}),
    (r"TAI CHO", {
        "header": "DECLARATION #", "two_rows": False, "key": ["=NO"],
        "fees": [("SERVICE CHARGE", None)], "date": "DECLARATION/ DATE",
        "fields": {"order_no": "=INVOICE NO"}}),
]
_adc_sheet_parse, _adc_sheet_totals = _vt_make("CÔNG TY TNHH GIAO NHẬN THƯƠNG MẠI RỒNG Á CHÂU", _VT_ADC)


def _rac_is_adc(book) -> bool:
    return any(re.search(r"(FCL|LCL|AIR).*POU|TAI CHO", _vt_norm(n)) for n, _ in book.sheets)


def _rac_sheet_parse(book):
    return _adc_sheet_parse(book) if _rac_is_adc(book) else parse_rong_a_chau_sheets(book)


def _rac_sheet_totals(book):
    return _adc_sheet_totals(book) if _rac_is_adc(book) else _rong_a_chau_totals(book)


def detect_rong_a_chau_any(text: str) -> bool:
    return detect_rong_a_chau(text) or "ASIA DRAGON FORWARDING" in text.upper()


register_supplier(
    "RONG_A_CHAU", "CÔNG TY TNHH GIAO NHẬN THƯƠNG MẠI RỒNG Á CHÂU", detect_rong_a_chau_any, None,
    sheet_parse=_rac_sheet_parse, sheet_totals=_rac_sheet_totals,
)


# ===========================================================================
# 越南 PDF 帳單 (2026-10-27 新增)：VN07 MAERSK、VN10 KN、VN04 DHL EXPRESS
# ===========================================================================
# 越南的正式發票是電子增值稅發票 (HÓA ĐƠN GIÁ TRỊ GIA TĂNG / VAT INVOICE)，每家
# 版面不同。INVOICE NO = 發票號碼 (Số / Invoice No.，不含發票符號 1C26TSG，待使用者
# 確認)。數字是英式千分位 (12,556,673)。


def _vn_pdf_num(text: Optional[str]):
    """'12,556,673' / '3,587.5' -> 數字。"""
    if not text:
        return None
    try:
        value = float(text.replace(",", ""))
    except ValueError:
        return None
    return int(value) if value.is_integer() else value


def _vn_pdf_date(text: Optional[str]) -> Optional[str]:
    m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", text or "")
    return _format_date_en(*m.groups()) if m else None


# --- VN07 MAERSK LOGISTICS & SERVICES VIỆT NAM ---------------------------------
# PDF：電子發票，費用行 "Phí xếp dỡ hàng hóa FCR001555329 VND M3 29.682 423,040 8% 12,556,673"
# (名稱、FCR、幣別、單位、數量、單價、稅率、未稅金額)，稅額只有總數 -> 單獨一筆 VAT。
# Excel：Late come charge 對帳表，每列有自己的發票號碼 (INVOICE NO.)。
_MAERSK_VN_ITEM_RE = re.compile(
    # 參考號碼也可能數字開頭 ("26HCM0018482"，2026-10-28 抽樣)
    r"^(?P<desc>.+?)\s+(?P<ref>FCR\d+|[A-Z]{3,}\d{5,}|\d{2}[A-Z]{2,4}\d{5,})\s+(?P<cur>[A-Z]{3})\s+(?P<unit>\S+)\s+"
    r"(?P<qty>[\d,.]+)\s+(?P<price>[\d,.]+)\s+(?P<rate>\d+%|KCT)\s+(?P<amount>[\d,.]+)$", re.MULTILINE)


def detect_maersk_vn(text: str) -> bool:
    upper = text.upper()
    return "MAERSK LOGISTICS & SERVICES VI" in upper or (
        "MAERSK" in upper and "LATE COME CHARGE" in upper)


def parse_maersk_vn(text: str) -> Dict:
    h: Dict[str, Optional[str]] = {}
    series = _search(r"K[ýy] hi[ệe]u[^:]*:\s*(\S+)", text)
    number = _search(r"S[ốo] \(Invoice No\.\):\s*(\d+)", text)
    h["invoice_no"] = number   # 只用號碼 (發票符號 series 待使用者確認要不要加)
    h["invoice_date"] = _vn_pdf_date(_search(r"\(Invoice Date\):\s*(\S+)", text))
    h["supplier"] = "CÔNG TY TNHH MAERSK LOGISTICS & SERVICES VIỆT NAM"
    h["consignee"] = _search(r"S[ốo] \(Invoice No\.\):[^\n]*\n(.+)", text)
    h["port_of_discharge"] = _search(r"\(POD\):\s*(\S+)", text)
    h["port_of_loading"] = _search(r"\(POL\):\s*(\S+)", text)
    h["order_no"] = _search(r"\(Ref No\):\s*(\S+)", text)
    items, cbm = [], None
    for m in _MAERSK_VN_ITEM_RE.finditer(text):
        h.setdefault("bl_no", m.group("ref"))
        if m.group("unit").upper() == "M3":
            cbm = m.group("qty")
        items.append({"description": m.group("desc").strip(), "amount": _vn_pdf_num(m.group("amount")),
                      "vat": None})
    h["volume"] = _fmt_volume(cbm=cbm)
    vat = _search(r"\(Total VAT amount\):\s*([\d,.]+)", text)
    rate = _search(r"\s(\d+)%\s+[\d,.]+$", text)
    if vat:
        items.append({"description": f"VAT {rate}%" if rate else "VAT", "amount": None,
                      "vat": _vn_pdf_num(vat)})
    return {"header": h, "items": items or [{"description": None, "amount": None, "vat": None}]}


def _maersk_vn_total(text: str):
    return _vn_pdf_num(_search(r"\(Subtotal\):\s*([\d,.]+)", text))


def _maersk_vn_grand(text: str):
    return _vn_pdf_num(_search(r"\(Total amount\):\s*([\d,.]+)", text))


_MAERSK_VN_SHEET = [
    (r".", {"header": "BILL NO", "key": ["=INVOICE NO.", "BILL NO"],
            "fees": [("AMOUNT (NOT VAT)", "@REMARK")], "vat": "=VAT", "vat_desc": "VAT 8%",
            "total": "AMOUNT (NOT VAT)", "grand": "TOTAL (VND)", "cbm": "=CBM", "two_rows": False,
            "fields": {"bl_no": "BILL NO", "order_no": "SO BOOKING", "onboard_date": "NGAY XUAT"}}),
]
_maersk_vn_sheet_parse, _maersk_vn_sheet_totals = _vt_make(
    "CÔNG TY TNHH MAERSK LOGISTICS & SERVICES VIỆT NAM", _MAERSK_VN_SHEET)

register_supplier(
    # 月結 PDF (LOCAL CHARGES) 是好幾張電子發票合在一起：每頁一張 (2026-10-28 抽樣)
    "MAERSK_VN", "CÔNG TY TNHH MAERSK LOGISTICS & SERVICES VIỆT NAM", detect_maersk_vn, parse_maersk_vn,
    multi_page=True,
    extract_total=_maersk_vn_total, extract_grand_total=_maersk_vn_grand,
    sheet_parse=_maersk_vn_sheet_parse, sheet_totals=_maersk_vn_sheet_totals,
)


# --- VN10 KUEHNE + NAGEL 越南 -------------------------------------------------
# 兩種 PDF：
#  (a) 電子發票：費用行 "1 Phí khai manifest tại nơi đi 1 SDG 650,000 650,000 8% 52,000 702,000"
#      (序號、名稱、數量、單位、單價、未稅、稅率、稅額、含稅)，每筆自帶稅額。
#  (b) 附件清單 (DANH SÁCH ĐÍNH KÈM HOÁ ĐƠN)：每列一票 (Số File)，10 個費用欄 + 合計 + VAT。
#      表頭在 PDF 裡兩行交錯，費用名稱依欄位順序整理在 _KN_VN_LIST_FEES。
#      使用者 2026-10-27 確認名稱 (不含 "- 215" 這類代碼；PDF 字型把 biển/cảng 讀成 bi^n/cang)。
_KN_VN_ITEM_RE = re.compile(
    r"^\d+\s+(?P<desc>.+?)\s+(?P<qty>[\d,.]+)\s+(?P<unit>[A-Z0-9]{2,5})\s+(?P<price>[\d,.]+)\s+"
    r"(?P<amount>[\d,.]+)\s+(?P<rate>\d+%|KCT)\s+(?P<tax>[\d,.]+)\s+(?P<total>[\d,.]+)$", re.MULTILINE)
_KN_VN_ITEM_NODESC_RE = re.compile(
    r"^\d+\s+(?P<qty>[\d,.]+)\s+(?P<unit>[A-Z0-9]{2,5})\s+(?P<price>[\d,.]+)\s+"
    r"(?P<amount>[\d,.]+)\s+(?P<rate>\d+%|KCT)\s+(?P<tax>[\d,.]+)\s+(?P<total>[\d,.]+)$")
_KN_VN_LIST_FEES = (
    "Phí biên nhận hàng đường biển", "CY monitoring", "Phí xếp dỡ tại cảng đi",
    "Phí làm hàng lb", "Phí e-VGM", "Phí Cơ sở hạ tầng cảng xuất",
    "Phí làm hàng ngoài giờ", "Phí sắp xếp hàng xuất", "Phí lưu kho hàng xuất",
    "Phí DEM/DET")
_KN_VN_LIST_ROW_RE = re.compile(
    r"^(?P<file>\d{8,})\s+(?P<dn>\S+)\s+(?P<so>\d+)\s+(?P<cbm>[\d.]+)\s+(?P<kg>[\d.]+?)"
    r"(?P<dest>[A-Z][A-Z ]+?)\s+(?P<money>(?:[\d,]+ VND\s*){12})$", re.MULTILINE)


def detect_kn_vn(text: str) -> bool:
    upper = text.upper()
    return "KUEHNE + NAGEL" in upper and ("HÓA ĐƠN" in upper or "HOÁ ĐƠN" in upper or "VIỆT NAM" in upper)


# 精簡版附件清單 (2026-10-28 抽樣 "KEUHNE+NAGEL_SEP.2026_BY WEEKLY" 第 1 頁)：空白的費用欄、SO#、
# 材積都不印，每列只剩 "1075399881 ED SEP 17 2026 6,147,130 VND 6,147,130 VND 491,770 VND"。
# 用 Grand Total 列 (12 個 VND 值都有印) 判斷哪幾個費用欄有金額，依序對應。
_KN_VN_SHORT_ROW_RE = re.compile(
    r"^(?P<file>\d{8,})\s+(?P<dn>.+?)\s+(?P<money>(?:[\d,]+ VND\s*){3,11})$", re.MULTILINE)


def _kn_vn_list_rows(text: str):
    rows = []
    for m in _KN_VN_LIST_ROW_RE.finditer(text):
        money = [_vn_pdf_num(v) for v in re.findall(r"([\d,]+) VND", m.group("money"))]
        rows.append({"file": m.group("file"), "dn": m.group("dn"), "so": m.group("so"),
                     "cbm": m.group("cbm"), "kg": m.group("kg"), "dest": m.group("dest").strip(),
                     "fees": money[:10], "total": money[10], "vat": money[11]})
    grand = _search(r"^Grand Total\s+(.+)$", text)
    grand_money = [_vn_pdf_num(v) for v in re.findall(r"([\d,]+) VND", grand or "")]
    if rows or len(grand_money) != 12:
        return rows
    used = [i for i, v in enumerate(grand_money[:10]) if v]
    for m in _KN_VN_SHORT_ROW_RE.finditer(text):
        money = [_vn_pdf_num(v) for v in re.findall(r"([\d,]+) VND", m.group("money"))]
        if len(money) != len(used) + 2:
            continue
        fees = [None] * 10
        for i, v in zip(used, money):
            fees[i] = v
        rows.append({"file": m.group("file"), "dn": m.group("dn").strip(), "so": None, "cbm": None,
                     "kg": None, "dest": "", "fees": fees, "total": money[-2], "vat": money[-1]})
    return rows


def parse_kn_vn(text: str) -> Dict:
    h: Dict[str, Optional[str]] = {"supplier": "CÔNG TY TNHH KUEHNE + NAGEL"}
    list_rows = _kn_vn_list_rows(text)
    if list_rows:
        h["invoice_date"] = _vn_pdf_date(_search(r"ĐƠN NGÀY\s*(\d{1,2}/\d{1,2}/\d{4})", text.upper()))
        h["consignee"] = "CÔNG TY TNHH POUYUEN VIỆT NAM" if "POUYUEN" in text.upper() else None
        items = []
        # 第 10 欄在精簡版清單是 "Phí thị thực" (不是 DEM/DET)
        names = _KN_VN_LIST_FEES[:9] + (("Phí thị thực",) if "Phí thị thực" in text else _KN_VN_LIST_FEES[9:])
        for r in list_rows:
            common = {"invoice_no": r["file"], "order_no": r["so"], "port_of_discharge": r["dest"] or None,
                      "volume": _fmt_volume(cbm=r["cbm"], kg=r["kg"])}
            items += [{"description": name, "amount": amt, "vat": None, **common}
                      for name, amt in zip(names, r["fees"]) if amt]
            items.append({"description": "VAT 8%", "amount": None, "vat": r["vat"], **common})
        return {"header": h, "items": items or [{"description": None, "amount": None, "vat": None}]}
    series = _search(r"\(Series\):\s*(\S+)", text)
    number = _search(r"\(Invoice No\.\):\s*(\S+)", text)
    h["invoice_no"] = number   # 只用號碼 (發票符號 series 待使用者確認要不要加)
    h["invoice_date"] = _vn_pdf_date(_search(r"\(Invoice date\):\s*(\S+)", text))
    h["consignee"] = _search(r"\(Customer.s name\):\s*(.+)", text)
    h["onboard_date"] = _vn_pdf_date(_search(r"\(Departure date\):\s*(\S+)", text))
    h["arrive_date"] = _vn_pdf_date(_search(r"\(Arrival date\):\s*(\S+)", text))
    # 欄位空白時，同一行接著是下一個標籤 ("Số lô hàng (File number): Số chứng từ nội bộ…"、
    # "(AWB/SWB): Tiền tệ…")，不可當成值 (2026-10-28 抽樣 KEUHNE+NAGEL_JUL.2026_BY SHIPMENT 第 44 頁)
    blank = r"(?!Số\b|Tiền\b|Ngày\b)"
    h["order_no"] = _search(r"\(File number\):[ \t]*" + blank + r"(\S+)", text)
    m = re.search(r"\(AWB/SWB\):[ \t]*" + blank + r"(\S+)[^\n]*\n(\d+)\s*$", text, re.MULTILINE)
    h["bl_no"] = (m.group(1) + m.group(2)) if m and m.group(1).endswith("-") else _search(
        r"\(AWB/SWB\):[ \t]*" + blank + r"(\S+)", text)
    items, cw, cbm, boxes = [], None, None, {}
    lines = text.splitlines()
    for i, line in enumerate(lines):
        line = line.strip()
        m = _KN_VN_ITEM_RE.match(line)
        desc = m.group("desc").strip() if m else None
        if not m:
            # 名稱太長折成上下兩行，中間那行只有序號與數字 (2026-10-28 抽樣：
            # "Phí Khai Trọng Lượng(VGM thông" / "2 1 SH 650,000 …" / "thường)")
            m = _KN_VN_ITEM_NODESC_RE.match(line)
            if not m:
                continue
            desc = " ".join(x.strip() for x in (lines[i - 1] if i else "", lines[i + 1] if i + 1 < len(lines) else "")
                            if x.strip()).strip() or None
        if m.group("unit") == "CW":
            cw = m.group("qty").replace(",", "")
        if re.fullmatch(r"\d{2}[A-Z]{2}", m.group("unit")):      # 櫃型 (40HC)：整櫃
            boxes[m.group("unit")] = max(boxes.get(m.group("unit"), 0), float(m.group("qty").replace(",", "")))
        elif m.group("unit") == "VO":                            # 散貨的 CBM
            cbm = m.group("qty")
        items.append({"description": desc, "amount": _vn_pdf_num(m.group("amount")),
                      "vat": _vn_pdf_num(m.group("tax"))})
    h["volume"] = _fmt_volume(containers=[(q, u) for u, q in boxes.items()], cbm=cbm, kg=cw)
    return {"header": h, "items": items or [{"description": None, "amount": None, "vat": None}]}


def _kn_vn_totals(text: str):
    """(未稅, 含稅)：電子發票的 "Tổng cộng tiền thanh toán" 下一行；附件清單每票一組。"""
    rows = _kn_vn_list_rows(text)
    if rows:
        return ({r["file"]: r["total"] for r in rows}, {r["file"]: r["total"] + r["vat"] for r in rows})
    m = re.search(r"Tổng cộng tiền thanh toán\s*\n\s*([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)", text)
    return (_vn_pdf_num(m.group(1)), _vn_pdf_num(m.group(3))) if m else (None, None)


register_supplier(
    # 月結 PDF (BY SHIPMENT) 是好幾張電子發票合在一起：每頁一張 (2026-10-28 抽樣)
    "KN_VN", "CÔNG TY TNHH KUEHNE + NAGEL (VIỆT NAM)", detect_kn_vn, parse_kn_vn, multi_page=True,
    extract_total=lambda t: _kn_vn_totals(t)[0], extract_grand_total=lambda t: _kn_vn_totals(t)[1],
)


# --- VN04 DHL - VNPT EXPRESS (CÔNG TY TNHH CHUYỂN PHÁT NHANH DHL-VNPT) ---------
# 進口運費明細 (IB TARIFF)：一張發票 (Invoice no. SGNIR…) 好幾頁、好幾十票快遞。
# 每票第一行 "26/08/2026 PUS 1973098271 SGN 3 6.00 1,365,383 1,365,383" (日期、起運地、
# HAWB、目的地、件數、重量、運費、淨額)，接著附加費 (FUEL SURCHARGE、GO GREEN…)，
# 再來 "PROD CODE EXPRESS WORLDWIDE nondoc"、"Sender: … Receiver: …"。
# VAT 只有整張一個；稽核：費用加總 = Sub-total，加 VAT = Total VND。
_DHL_VN_MAIN_RE = re.compile(
    # 機場代碼可能有數字 ("S2X"，掃描檔)
    r"^(?P<date>\d{2}/\d{2}/\d{4})\s+(?P<org>[A-Za-z0-9]{3})\s+(?P<hawb>\d{8,})\s+(?P<dest>[A-Za-z0-9]{3})\s+"
    r"(?P<pcs>\d+)\s+(?P<kg>[\d.]+)\s+(?P<regular>[\d,]+)(?:\s+[\d,]+)*\s+(?P<net>[\d,]+)$")
# 附加費名稱有大小寫混合 ("Direct Signature 150,000 150,000")
# 只有一個金額的附加費 ("SHIPMENT VALUE PROTECTION 354,532") 也算；金額要有千分位或 3 位以內，
# 不會誤抓 "YOUR VAT CODE 3900911265" 這種代碼
_DHL_VN_SUR_RE = re.compile(
    r"^(?P<desc>[A-Z][A-Za-z \-/]+?)\s+(?:(?:(?P<regular>[\d,]+)\s+)(?:[\d,]+\s+)*(?P<net>\d{1,3}(?:,\d{3})*)"
    # 只有一個金額時一定要有千分位 (不然 "Page 2" 會被當成 2 元)
    r"|(?P<net1>\d{1,3}(?:,\d{3})+))$")


def detect_dhl_express_vn(text: str) -> bool:
    return _is_dhl_vnpt(text)


def _dhl_vn_invoice_no(text: str) -> Optional[str]:
    """Invoice no.：掃描字型會把 0 讀成 O、G 讀成 C ("SCNIRO0310381")，每頁都有一份，
    先把數字旁的 O 換成 0，再優先取 SGN 開頭的那一份。"""
    found = [re.sub(r"(?<=[\dR])O(?=[\dO])|(?<=\d)O", "0", v) for v in re.findall(r"Invoice no\.\s*(\S+)", text)]
    found = [re.sub(r"(?<=\d)O|O(?=\d)", "0", v) for v in found]
    return next((v for v in found if v.startswith("SGN")), found[0] if found else None)


def _dhl_vn_dot_date(value, month_first: bool) -> Optional[str]:
    """'31.03.2026' (日.月.年，通知日) 或 '02.27.2026' (月.日.年，出貨日)。"""
    m = re.match(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", _vn_text(value))
    if not m:
        return _vn_date(value)
    a, b, year = m.groups()
    day, month = (b, a) if month_first else (a, b)
    if int(month) > 12:          # 欄位順序寫反時自動對調
        day, month = month, day
    return _format_date_en(day, month, year)


def _dhl_vn_sheets(book):
    """DHL-EXP 出口 Excel (2026-10-28 抽樣，DHL-EXP2025xx.xls)：每個部門 (Shipper referent)
    一張工作表，同一個 NOTIFICATION NUMBER。每票一列 (出貨日、提單、國家、重量、運費)，
    下面接該票附加費；表尾 Fuel Surcharge、VAT 8%、Grand Total 是整張工作表的。
    回傳 (sheet_parse 結果, sheet_totals 結果)：稽核 = 各表 Grand Total 加總。"""
    items, totals = [], {}
    for name, rows in book.sheets:
        head = next((i for i, r in enumerate(rows[:30]) if any(_vt_norm(c) == "AIRWAYBILL" for c in r)), None)
        if head is None:
            continue
        info = {}
        for r in rows[:head]:
            cells = [c for c in r if c not in (None, "")]
            for j, c in enumerate(cells[:-1]):
                label = _vt_norm(c)
                for key in ("NOTIFICATION NUMBER", "NOTIFICATION DATE", "ACCOUNT NUMBER", "SHIPPER REFERENT"):
                    if label.startswith(key):
                        info[key] = cells[j + 1]
                        if key == "ACCOUNT NUMBER":
                            info["CONSIGNEE"] = cells[0] if j else None
        cols = {_vt_norm(c): j for j, c in enumerate(rows[head]) if c not in (None, "")}
        c_date, c_awb, c_country = cols.get("SHIPMENT DATE"), cols.get("AIRWAYBILL"), cols.get("COUNTRY")
        c_weight, c_charge = cols.get("WEIGHT"), cols.get("CHARGE")
        inv = _vn_text(info.get("NOTIFICATION NUMBER")) or name.strip()
        common = {"invoice_no": inv, "supplier": "CÔNG TY TNHH CHUYỂN PHÁT NHANH DHL - VNPT",
                  "consignee": _vn_text(info.get("CONSIGNEE")) or None,
                  "invoice_date": _dhl_vn_dot_date(info.get("NOTIFICATION DATE"), month_first=False),
                  "order_no": _vn_text(info.get("SHIPPER REFERENT")) or None, "port_of_loading": "VIET NAM"}
        current = None
        for r in rows[head + 1:]:
            nums = [c for c in r if isinstance(c, (int, float)) and not isinstance(c, bool)]
            label = _vt_norm(_vn_cell(r, c_awb)) if c_awb is not None else ""
            date = _vn_cell(r, c_date)
            if re.match(r"\d{1,2}\.\d{1,2}\.\d{4}", _vn_text(date)):
                current = {**common, "bl_no": _vn_text(_vn_cell(r, c_awb)),
                           "onboard_date": _dhl_vn_dot_date(date, month_first=True),
                           "port_of_discharge": _vn_text(_vn_cell(r, c_country)) or None,
                           "volume": _fmt_volume(kg=f"{_vn_num(_vn_cell(r, c_weight)):g}")
                           if _vn_num(_vn_cell(r, c_weight)) else None}
                items.append({"description": "FREIGHT CHARGE", "amount": _vn_num(_vn_cell(r, c_charge)),
                              "vat": None, **current})
            elif label.startswith("TOTAL FOR SHIPMENT") or not nums or not label:
                continue
            elif label.startswith("GRAND TOTAL"):
                totals[inv] = totals.get(inv, 0) + nums[-1]
            elif label.startswith("VAT"):
                items.append({"description": _vn_text(_vn_cell(r, c_awb)), "amount": None,
                              "vat": nums[-1], **common})
            elif label.startswith("FUEL SURCHARGE") and "%" in label:
                # 表尾燃油附加費 (整張工作表)："Fuel Surcharge 28.75%"
                items.append({"description": _vn_text(_vn_cell(r, c_awb)), "amount": nums[-1], "vat": None,
                              **common})
            elif current and nums[-1]:
                items.append({"description": _vn_text(_vn_cell(r, c_awb)), "amount": nums[-1], "vat": None,
                              **current})
    parsed = [{"header": {}, "items": items or [{"description": None, "amount": None}]}]
    return parsed, [{"invoice_no": k, "total": None, "grand_total": v} for k, v in totals.items()]


def _dhl_vn_clean(text: str) -> str:
    """掃描檔的數字常在千分位逗號後多一個空白 ("914, 328")，接回去。"""
    return re.sub(r"(\d),\s+(\d{3})(?!\d)", r"\1,\2", text)


def parse_dhl_express_vn(text: str) -> Dict:
    text = _dhl_vn_clean(text)
    h: Dict[str, Optional[str]] = {
        "supplier": "CÔNG TY TNHH CHUYỂN PHÁT NHANH DHL - VNPT",
        "invoice_no": _dhl_vn_invoice_no(text),
        "invoice_date": _vn_pdf_date(_search(r"^Date\s+(\S+)", text)),
        # 掃描字型會把 "DHL contact" 讀成 "DEL contact"
        "consignee": _search(r"^(C[OÔ]NG TY .+?)\s+D[HE]L contact", text)}
    items, current, start = [], None, False
    for line in text.splitlines():
        line = line.strip()
        if line.startswith("DATE ORG HAWB"):
            start = True
            continue
        # 每種產品 (DOC / nondoc) 各有一行 Sub-total，讀到 TOTAL FOR SHIPMENT 才結束
        if not start or line.startswith(("Sub-total", "TOTAL", "VAT ")):
            if line.startswith("Sub-total"):
                current = None
            elif line.startswith("TOTAL"):
                start = False
            continue
        m = _DHL_VN_MAIN_RE.match(line)
        if m:
            current = {"bl_no": m.group("hawb"), "onboard_date": _vn_pdf_date(m.group("date")),
                       "port_of_loading": m.group("org").upper(), "port_of_discharge": m.group("dest").upper(),
                       "volume": _fmt_volume(kg=m.group("kg"))}
            # 使用者 2026-10-27：ORG 這一行的運費名稱固定 "FREIGHT CHARGE" (不用產品代碼)
            items.append({"description": "FREIGHT CHARGE", "amount": _vn_pdf_num(m.group("net")),
                          "vat": None, **current})
            continue
        if current is None or line.startswith("PROD CODE"):
            continue
        m = re.match(r"Sender:\s*(.+?)\s+Receiver:", line)
        if m:
            for it in items:
                if it.get("bl_no") == current["bl_no"]:
                    it["shipper"] = m.group(1).strip()
            continue
        m = _DHL_VN_SUR_RE.match(line)
        net = _vn_pdf_num(m.group("net") or m.group("net1")) if m else None
        if net:
            items.append({"description": m.group("desc").strip(), "amount": net, "vat": None, **current})
    # 一份 PDF 可能分好幾段 (每個部門一段)，每段各有 VAT 與 Total VND
    for vat in re.finditer(r"^VAT\s+([\d,]+)\s+([\d,]+)$", text, re.MULTILINE):
        items.append({"description": "VAT", "amount": None, "vat": _vn_pdf_num(vat.group(2))})
    return {"header": h, "items": items or [{"description": None, "amount": None, "vat": None}]}


def _dhl_express_vn_total(text: str):
    """未稅合計：每種產品一行 Sub-total (最後一個數字是淨額)，全部加總。"""
    subs = re.findall(r"^Sub-total:.*?([\d,]+)\s+[\d,]+\s+([\d,]+)$", _dhl_vn_clean(text), re.MULTILINE)
    return sum(_vn_pdf_num(net) or 0 for _, net in subs) if subs else None


def _dhl_express_vn_grand(text: str):
    totals = re.findall(r"^Total VND\s+([\d,]+)", _dhl_vn_clean(text), re.MULTILINE)
    return sum(_vn_pdf_num(v) or 0 for v in totals) if totals else None


# --- IB TARIFF OF DHL SERVICES (進口代墊關稅帳單，2026-10-28 抽樣 "342 EXP 052025.pdf" 第 4 頁) ----
# 同一份 PDF 前面附提單 (WAYBILL DOC) 與中國海關單據，只擷取帳單頁。每票一段：
# "18/03 SGN 5990639594 CAN" (日/月 起運 HAWB 目的)，下面是費用行 "IMPORT EXPORT DUTIES 389591 389591"
# (最後一個數字是淨額)；整張 VAT 單獨一列；"Sub-total" 勾稽未稅、"TOTAL VND" 勾稽含稅。
def _is_dhl_ib_tariff(text: str) -> bool:
    # 一般月結帳單 (有 "DATE ORG HAWB" 表頭) 的頁尾也會印 "IB TARIFF OF DHL SERVICES"
    # (2026-10-31 越南資料1030：POU HUNG 09-2026 PHÍ CPN DHL-VNPT IMPORT.pdf 等 182 份被誤判、擷取不到)
    up = text.upper()
    return "IB TARIFF OF DHL SERVICES" in up and "DATE ORG HAWB" not in up


def _parse_dhl_ib_tariff(text: str) -> Dict:
    date = re.search(r"\bDATE\s+(\d{1,2})/(\d{1,2})/(\d{4})", text)
    year = date.group(3) if date else None
    h = {"supplier": "CÔNG TY TNHH CHUYỂN PHÁT NHANH DHL - VNPT",
         "invoice_no": _search(r"INVOICE NO\s+(\S+)", text),
         "invoice_date": _format_date_en(*date.groups()) if date else None,
         # "CONG TY TNHH POUHUNG VIET NAM DHL CONTACT INVOICING HOTLINE" (左右兩欄同一行)
         "consignee": _search(r"^(C[OÔ]NG TY [A-Z ]+?)\s+DHL CONTACT", text)}
    items, common = [], {}
    for line in text.splitlines():
        line = line.strip()
        m = re.match(r"^(\d{1,2})/(\d{1,2})\s+([A-Z]{3})\s+(\d{8,})\s+([A-Z]{3})$", line)
        if m:
            common = {"onboard_date": _format_date_en(m.group(1), m.group(2), year) if year else None,
                      "port_of_loading": m.group(3), "bl_no": m.group(4), "port_of_discharge": m.group(5)}
            continue
        # 費用行至少兩個數字 (費用、淨額)；地址行 "TAY NINH GAOYAO 526100" 只有郵遞區號
        m = re.match(r"^([A-Z][A-Z &/().-]*?[A-Z)])\s+((?:\d[\d.]*\s+)+\d+)$", line)
        if m and common and not re.match(r"(SUB-TOTAL|VAT|TOTAL|PROD CODE|PAGE|DATE)\b", m.group(1)):
            items.append({"description": m.group(1), "amount": int(m.group(2).split()[-1]), "vat": None, **common})
    vat = _search(r"^VAT\s+[\d\s]*?(\d+)\s*$", text)
    if items and vat:
        items.append({"description": "VAT", "amount": None, "vat": int(vat), **common})
    return {"header": h, "items": items or [{"description": None, "amount": None, "vat": None}]}


def _dhl_express_vn_parse(text: str) -> Dict:
    return _parse_dhl_ib_tariff(text) if _is_dhl_ib_tariff(text) else parse_dhl_express_vn(text)


def _dhl_express_vn_total2(text: str):
    if _is_dhl_ib_tariff(text):
        v = _search(r"^Sub-total\b.*?(\d+)\s*$", text)
        return int(v) if v else None
    return _dhl_express_vn_total(text)


def _dhl_express_vn_grand2(text: str):
    if _is_dhl_ib_tariff(text):
        v = _search(r"^TOTAL VND\s+(\d+)", text)
        return int(v) if v else None
    return _dhl_express_vn_grand(text)


register_supplier(
    "DHL_EXPRESS_VN", "CÔNG TY TNHH CHUYỂN PHÁT NHANH DHL - VNPT", detect_dhl_express_vn,
    _dhl_express_vn_parse, extract_total=_dhl_express_vn_total2, extract_grand_total=_dhl_express_vn_grand2,
    # 合輯前面的提單 (WAYBILL DOC) 與海關單據不是帳單頁；沒有任何頁符合時 invoice_text 會用全部頁
    invoice_page=lambda page: "WAYBILL DOC" not in page.upper() and "海关" not in page,
    sheet_parse=lambda book: _dhl_vn_sheets(book)[0], sheet_totals=lambda book: _dhl_vn_sheets(book)[1],
)


# --- VN11 APL LOGISTICS VIỆT NAM -------------------------------------------------
# PDF：一頁一張電子發票 (multi_page)，越南式數字 (13.025.090、數量 15,15)。
#   "Chứng từ số: 149767, SINAR SANUR 149S - HO CHI MINH CITY, VIETNAM - SANTOS,SP, BRAZIL"
#   = 單號、船名航次 - 啟運港 - 目的港 (卡車是 "BINH DUONG - DONG NAI-50H13141")。
#   費用行 "1 Phụ phí xăng dầu CBM 15,15 26.594 402.899 8% 32.232 435.131" 每筆自帶稅額；
#   整櫃費用 "Phí điều hành bến bãi công-tơ-nơ 40DC CTR 1" 用來算櫃量。
# Excel：延滯費 (NEO CONT/NEO KHO) 對帳表，每列一個貨櫃，發票號碼 = 每列的 Inv# (使用者
# 2026-10-27)；整張一個 VAT，發票號碼用表尾的 Invoice No.。
_APL_VN_ITEM_RE = re.compile(
    r"^\d+\s+(?P<desc>.+?)\s+(?P<unit>[A-Z]{2,4})\s+(?P<qty>[\d.,]+)\s+(?P<price>[\d.]+)\s+"
    # 稅率欄也可能印 "X" (代收代付 Thu hộ，2026-10-28 抽樣 "APL_APR.2026_LOCAL CHARGES" 第 1 頁)
    r"(?P<amount>[\d.]+)\s+(?P<rate>\d+%|KCT|X)\s+(?P<tax>[\d.]+)\s+(?P<total>[\d.]+)$", re.MULTILINE)


def detect_apl_vn(text: str) -> bool:
    upper = text.upper()
    # 舊字型的卡車表寫成 "APL LOGISTICS VIEÄT NAM" (2026-10-28 抽樣)
    return "APL LOGISTICS VIỆT NAM" in upper or "APL LOGISTICS V" in upper and (
        "POUHUNG" in upper or "POUYUEN" in upper or "CONTSEAL" in upper)


def parse_apl_vn(text: str) -> Dict:
    h: Dict[str, Optional[str]] = {"supplier": "CÔNG TY TNHH APL LOGISTICS VIỆT NAM"}
    series = _search(r"\(Serial\):\s*(\S+)", text)
    number = _search(r"S[ốo] \(No\.\):\s*(\d+)", text)
    h["invoice_no"] = number   # 只用號碼 (發票符號 series 待使用者確認要不要加)
    m = re.search(r"Ngày \(date\)\s*(\d{1,2})\s*tháng \(month\)\s*(\d{1,2})\s*năm \(year\)\s*(\d{4})", text)
    h["invoice_date"] = _format_date_en(*m.groups()) if m else None
    h["consignee"] = _search(r"\(Company name\):\s*(.+)", text)
    h["order_no"] = _search(r"Chứng từ số:\s*(\d+)", text)
    # 第二個 "Chứng từ số: N, …" 到第一筆費用之間是船名 - 啟運港 - 目的港
    m = re.search(r"Chứng từ số:\s*\d+,\s*(.*?)\n\s*1\s", text, re.DOTALL)
    route = re.sub(r"\s+", " ", m.group(1)).strip() if m else ""
    parts = [p.strip() for p in route.split(" - ")]
    if len(parts) == 3:
        h["vessel"], h["port_of_loading"], h["port_of_discharge"] = parts
    items, cbm, boxes = [], None, []
    for im in _APL_VN_ITEM_RE.finditer(text):
        desc = im.group("desc").strip()
        qty = im.group("qty").replace(".", "").replace(",", ".")
        if im.group("unit") == "CBM":
            cbm = qty
        box = re.search(r"công-tơ-nơ\s+(\d{2}\w*)$", desc)
        if box and im.group("unit") == "CTR":
            boxes.append((qty, box.group(1)))
        items.append({"description": desc, "amount": _dsv_vn_dot_amount(im.group("amount")),
                      "vat": _dsv_vn_dot_amount(im.group("tax"))})
    h["volume"] = _fmt_volume(containers=boxes, cbm=cbm)
    return {"header": h, "items": items or [{"description": None, "amount": None, "vat": None}]}


def _apl_vn_totals(text: str):
    m = re.search(r"Tổng cộng \(Total\):\s*([\d.]+)\s+([\d.]+)\s+([\d.]+)", text)
    return (_dsv_vn_dot_amount(m.group(1)), _dsv_vn_dot_amount(m.group(3))) if m else (None, None)


def apl_vn_invoice_page(page: str) -> bool:
    return "HÓA ĐƠN GIÁ TRỊ GIA TĂNG" in page.upper() and "APL" in page.upper()


_APL_VN_SHEET = [
    (r".", {"header": "BOOKING NO", "key": ["=INV#", "BOOKING NO"], "vat_doc_only": True,
            "fees": [("CHASSIS CONTSEAL", None), ("DROP OFF", None), ("CHARGE DET", "DET"),
                     ("CUSTOMS / OTHERS", None)],
            "total": "TOTAL ( VND )", "doc_key": r"Invoice No\.:\s*(\d+)", "doc_date": "INVOICE DATE",
            "vat_total": "TOTAL VAT AMOUNT", "vat_desc": "VAT 8%",
            "fields": {"container_no": "CONTAINER NUMBER", "order_no": "BOOKING NO",
                       "onboard_date": "PICK-UP MT DATE"}}),
]
_apl_vn_sheet_parse, _apl_vn_sheet_totals = _vt_make("CÔNG TY TNHH APL LOGISTICS VIỆT NAM", _APL_VN_SHEET)

register_supplier(
    "APL_VN", "CÔNG TY TNHH APL LOGISTICS VIỆT NAM", detect_apl_vn, parse_apl_vn, multi_page=True,
    extract_total=lambda t: _apl_vn_totals(t)[0], extract_grand_total=lambda t: _apl_vn_totals(t)[1],
    sheet_parse=_apl_vn_sheet_parse, sheet_totals=_apl_vn_sheet_totals,
)


SUPPLIER_NUMBERS = {
    # 供應商 key: (編號, 廠商代號, 進度管控表上的公司名稱)
    "YJE_TATA": ("ID01", "100019748", "PT. TATA HARMONI SARANATAMA"),
    "MAERSK": ("ID02", "100019226", "PT. MAERSK LOGISTICS INDONESIA"),
    "HANDAL": ("ID03", "100019606", "PT.HANDAL MANDIRI TRANSINDO"),
    "DHL": ("ID04", "100019563", "PT. BIROTIKA SEMESTA"),
    "KUEHNE_NAGEL": ("ID05", "100019143", "PT. KUEHNE NAGEL INDONESIA"),
    "DSV": ("ID06", "100023151", "PT. DSV TRANSPORT INDONESIA"),
    "INDOPROSTIME_TRANSPORT": ("ID07", "100019300", "PT. INDOPROSTIME TRANSPORT"),
    "OPL": ("ID08", "100004377", "ORIENTAL POWER LOGISTICS"),
    "OPL_USD": ("ID08-OPL(USD)", "ID08-OPL(USD)", "OPL曜捷"),
    "DWIHARTA": ("ID09", "100019242", "PT. DWIHARTA LOGISTINDO"),
    "TRANS_DAYA_PRIMA": ("ID10", "100019488", "PT. TRANS DAYA PRIMA"),
    "GALUH": ("ID11", "100019600", "PT.GALUH INTI BAHARI"),
    "ECOMET": ("ID12", "100019489", "PT. TRANS ECOMET GLOBAL"),
    "PUTRA_SEJAHTERA_SENTOSA": ("ID13", "100030053", "PT. PUTRA SEJAHTERA SENTOSA"),
    "APL": ("ID14", "100019535", "PT. APL LOGISTICS"),
    "PANCARAN": ("ID15", "100019386", "PT. PANCARAN SRIKANDI LOGISTIK"),
    "INDOPROSTIME": ("ID16", "100030991", "PT INDO PROSTIME EXPRESS"),
    "FEMARIA": ("ID18", "100019597", "PT. FEMARIA BUANA CARGO"),
    "EXPRESS_MAXIMUM": ("ID21", "100019046", "PT EXPRESS MAXIMUM"),
    "PAN_EKSPRES": ("ID28", "100030970", "PT PAN EKSPRES INTERNATIONAL"),
    "HYPER_MEGA": ("ID31", "100019295", "PT. HYPER MEGA SHIPPING"),
    "EVERGREEN_LOGISTICS": ("ID38", "100019246", "PT. EVERGREEN LOGISTICS INDONESIA"),
    "RONG_A_CHAU": ("VN03", "100008978", "CÔNG TY TNHH GIAO NHẬN THƯƠNG MẠI RỒNG Á CHÂU"),
    "CONG_THANH": ("VN05", "100008048", "CTY TNHHDV GN VT VA TM CONG THANH"),
    "TRANS_WAGON": ("VN01", "100009281", "CÔNG TY TNHH LIÊN VẬN QUỐC TẾ"),
    "DHL_EXPRESS_VN": ("VN04", "100008681", "CÔNG TY TNHH CHUYỂN PHÁT NHANH DHL-VNPT"),
    "MAERSK_VN": ("VN07", "100008773", "CÔNG TY TNHH MAERSK LOGISTICS & SERVICES VIỆT NAM"),
    "KN_VN": ("VN10", "100009241", "CÔNG TY TNHH KUEHNE+NAGEL"),
    "APL_VN": ("VN11", "100008592", "CÔNG TY TNHH APL LOGISTICS VIỆT NAM"),
    "VINH_HUNG_LONG": ("VN02", "100008024", "CÔNG TY TNHH VĨNH HƯNG LONG"),
    "BINH_TAN": ("VN06", "100012345", "HỢP TÁC XÃ VẬN TẢI DỊCH VỤ DU LỊCH BÌNH TÂN"),
    "NGUYEN_HUONG": ("VN12", "100009955", "CÔNG TY TNHH NGUYÊN HƯỞNG"),
    "ORIENTAL_MULTIMODAL": ("VN14", "100024393", "Công Ty TNHH Logistics Vận Tải Đa Phương Thức"),
    "SONG_TRIEU": ("VN16", "100010180", "CÔNG TY TNHH SONG TRIỀU"),
    "TIMES_CARGO": ("VN17", "100008880", "CÔNG TY TNHH DỊCH VỤ GIAO NHẬN HÀNG HÓA"),
    "SAIGON_TRANSPORT": ("VN18", "100007723", "CÔNG TY TNHH DỊCH VỤ VẬN CHUYỂN SÀI GÒN"),
    "IMX": ("VN19", "100025119", "CÔNG TY TNHH DỊCH VỤ CHUYỂN PHÁT NHANH IMX"),
    # 越南廠商代號：越南進度管控\By Vietnam supplier (Top 20) 的資料夾名稱
    "DSV_VN": ("VN08", "100007866", "CÔNG TY TNHH DSV AIR & SEA"),
    "MOL": ("VN09", "100009335", "CÔNG TY TNHH MOL LOGISTICS VIỆT NAM"),
    "DHL_VN": ("VN13", "100007618", "CTY CP GIAO NHẬN TOÀN CẦU DHL (VIỆTNAM)"),
    "EXPEDITORS": ("VN15", "100008945", "CÔNG TY TNHH EXPEDITORS VIỆT NAM"),
    "YUSEN": ("VN20", "100007349", "CHI NHÁNH CÔNG TY TNHH YUSEN LOGISTICS (VIỆTNAM) TẠI THÀNH PHỐ HỒ CHÍ MINH"),
}

for _key, (_number, _vendor_code, _name) in SUPPLIER_NUMBERS.items():
    set_supplier_number(_key, _number, _vendor_code, _name)
