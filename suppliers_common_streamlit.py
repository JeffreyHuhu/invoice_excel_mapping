# -*- coding: utf-8 -*-
"""
suppliers_common_streamlit.py
==============================
所有供應商共用的小工具函式 (日期格式、材積格式、越南 Excel 儲存格讀取等)，給
suppliers_indonesia_streamlit.py、suppliers_vietnam_streamlit.py 共用。

日後新增國家 (柬埔寨/孟加拉/緬甸/中國) 時，新檔案一樣從這裡 import 共用函式。
只放跨供應商共用的邏輯，不要加任何供應商專屬的 detect_xxx()/parse_xxx()。

註：_vn_text/_vn_num/_vn_amount/_vn_cell、_is_dhl_vnpt、_dsv_parse_date 雖然
名稱帶 "vn"/"dsv"，但原本就被印尼供應商的 Excel 對帳表 (DWIHARTA/ECOMET/
HANDAL/PUTRA rekap) 或越南供應商 (DSV_VN) 跨國重複使用，所以放在這裡，
不要搬回個別國家檔案。
"""

import re
from typing import Dict, Optional

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
#   2. 散貨 (LCL) 沒有貨櫃數：CBM 數字 + " CBM"，例如 "7.010 CBM"。(2026-11-01 由 M3 統一改成 CBM)
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
    for value, unit in ((cbm, "CBM"), (kg, "KG")):
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




# 英文月份縮寫對照表 (後面其他供應商也共用這份)；也收印尼文縮寫
# (Mei/Agt/Okt/Des，例如 TRANS DAYA PRIMA "21 Agt 2025")
_EN_MONTHS_ABBR = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
    "mei": 5, "agt": 8, "agu": 8, "okt": 10, "des": 12,
}


def _is_dhl_vnpt(text: str) -> bool:
    """越南 DHL-VNPT 快遞 (VN04) 的帳單：別家的 detect 都要排除 (附件裡常出現 APL、DSV
    等託運人名稱，2026-10-28 抽樣)。"""
    return re.search(r"DHL\s*-?\s*VNPT", text.upper()) is not None


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


def _vn_text(value) -> str:
    """儲存格 -> 文字：整數型的浮點數 (例如報關單號 308933256500.0) 去掉 .0，
    連續空白壓成一個。"""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return re.sub(r"\s+", " ", str(value)).strip()


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


def _vn_cell(row: list, col: Optional[int]):
    return row[col] if col is not None and col < len(row) else None
