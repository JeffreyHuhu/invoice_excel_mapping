# -*- coding: utf-8 -*-
"""
suppliers_china_streamlit.py
=============================
中國供應商 (CH01~) 的 detect_xxx()/parse_xxx()，用 register_supplier()
註冊到 extract_utils_streamlit.SUPPLIER_REGISTRY。寫法仿照
suppliers_cambodia_streamlit.py。

跟其他國家供應商共用 suppliers_common_streamlit.py 的工具函式。
"""

import re
from typing import Dict, List, Optional

from extract_utils_streamlit import register_supplier
from suppliers_common_streamlit import _search, _format_date_en_from_month_name


# ===========================================================================
# 供應商 CH01：SALEM LOGISTICS COMPANY LTD. (UNIT C) 跟它的關係企業
# ORIENTAL LOGISTICS GROUP LTD. (UNIT A)，兩家版面相同、都歸 CH01 (香港，
# SALESMAN 都是 ORIENTAL POWER LOGISTICS CO LTD)，美金 INVOICE
# ===========================================================================
#
# 版面跟越南 OPL 的美金 DEBIT NOTE (ORIENTAL_USD) 很像 (NAME & ADDRESS/HBL
# INFO/CONTAINER 等欄位命名相同)，但頁首是「SALEM LOGISTICS COMPANY LTD.」
# 或「ORIENTAL LOGISTICS GROUP LTD.」＋「INVOICE」(不是 DEBIT NOTE)，用公司
# 名稱辨認即可，不會跟 ORIENTAL_USD 混到 (ORIENTAL_USD 的 detect 要求文字裡
# 有 "DEBIT NOTE" 或越南 ORIENTAL LOGISTICS MULTIMODAL 專用字樣)。
#
# 帳單幣別固定 USD (TOTAL: USD)，沒有 VAT 欄。費用列原幣可能是 HKD 或 USD，
# 每列各自有匯率換算成 USD，只需要抓最右邊換算後的 USD 金額 (原幣/匯率不填，
# 使用者規則文件 CH01_Claude_Code_修改說明.md 確認)。
#
# NAME & ADDRESS 是雙欄版面：「公司名稱」跟右欄的 INVOICE NO/ISSUE DATE
# 同一行印出，公司名稱常跨 2 行，第 2 行 (TAIWAN BRANCH) 才整行乾淨，見
# parse_ch01() 的 consignee 處理。
#
# 有些 SAMPLE 是掃描檔 (沒有文字層)，用 register_supplier(ocr_psm=6) 指定：
# 預設 OCR 會整段漏讀「NAME & ADDRESS: INVOICE NO : XXXXX」這一行 (INVOICE
# NO 整個抓不到)，psm=6 (整頁當一個文字區塊) 才讀得到，跟 GALUH 用同一個
# 機制。detect_ch01() 只看公司名稱 (不強制要求 "INVOICE" 字樣)，因為部分
# 掃描頁的預設 OCR 會把這一行整行漏掉，連 "INVOICE" 字樣都不見——只有公司
# 名稱還在，用公司名稱才能穩定判斷「這一頁是不是 CH01」、觸發 psm=6 重新
# OCR (2026-11-26 黃江裕祥抽樣答案)。
#
# 即使用 psm=6，掃描頁的數字仍常見雜訊：千分位逗點/小數點後面多一個空格
# ("21. 050"、"4, 300. 00")，小數點偶爾被認成逗點 ("44, 65" 應為 "44.65")，
# 扣款金額前面的負號偶爾被認成波浪號 ("~300. 00" 應為 "-300.00")。
# _ch01_num() 統一處理這些雜訊，不要另外在各欄位各寫一套清理。


def detect_ch01(text: str) -> bool:
    return "SALEM LOGISTICS COMPANY LTD" in text or "ORIENTAL LOGISTICS GROUP LTD" in text


# 掃描頁數字常見的雜訊：千分位逗點/小數點後面多一個空格 ("4, 300. 00")，
# 小數點被誤讀成逗點 ("44, 65" 應為 "44.65")。逗點後面剛好 2 位數字時，當成
# 小數點 (兩位數金額不會真的需要千分位)；3 位數字時才是千分位逗點。
_CH01_NUM_RE = r"\d+(?:[.,] ?\d+)*"


def _ch01_num(value: Optional[str]):
    """'219.84'/'2,000.000'/'4, 300. 00'/'44, 65'/'~300. 00' -> 數字
    (整數時回傳 int)。"""
    if not value:
        return None
    s = re.sub(r"\s+", "", value.strip())
    neg = s[:1] in ("-", "~")
    if neg:
        s = s[1:]
    if "," in s:
        left, _, right = s.partition(",")
        if len(right) == 2 and "." not in right:
            s = f"{left}.{right}"  # 小數點被誤讀成逗點
        else:
            s = left + right  # 千分位逗點
    try:
        num = float(s)
    except ValueError:
        return None
    if neg:
        num = -num
    return int(num) if num == int(num) else round(num, 2)


def _ch01_trim(value: Optional[str]) -> Optional[str]:
    """CBM 數字去掉多餘的小數尾數 0 ('21.670' -> '21.67')。"""
    num = _ch01_num(value)
    if num is None:
        return None
    return f"{num:g}" if isinstance(num, float) else str(num)


# 費用列："名稱 原幣 單價 / 單位 x 數量 匯率 USD 金額"，例如
# "CFS CHARGES HKD 165.000 / R/T x 3.120 0.128205 USD 66.00"、
# "SOLAS FEE USD 15.000 / SHPT x 1.000 1.000000 USD 15.00"。
# 只取最後的 USD 金額，原幣/匯率不填 (使用者規則文件確認)。
_CH01_ITEM_RE = re.compile(
    rf"^(?P<desc>.+?)\s+[A-Z]{{3}}\s+{_CH01_NUM_RE}\s*/\s*\S+\s+x\s+{_CH01_NUM_RE}\s+{_CH01_NUM_RE}\s+"
    rf"USD\s+(?P<amt>[-~]?{_CH01_NUM_RE})\s*$", re.MULTILINE)

# 掃描頁有一筆扣款費用「第6條取消艙位/甩櫃」的中文名稱，OCR (只認英文)
# 整段讀成亂碼 (例如 "FB OTR AGA Re (it / FA")，照使用者 ANSWER 還原。只套
# 用在這組已知的亂碼，不要用在其他看起來像亂碼但沒有確認過的費用名稱
# (2026-11-26 黃江裕祥 17092026105151.pdf 答案)。
_CH01_KNOWN_OCR_DESC = {
    "FB OTR AGA Re (it / FA": "第6條取消艙位/甩櫃",
}

# "HBL INFO: HKGJKT25120890 1X20'GP 37 PACKAGE(S) 5917.800 KGS 21.670 CBM"
# (有櫃型) 或 "HBL INFO: HKGJKT25120888 18 PACKAGE(S) 770.700 KGS 3.120 CBM"
# (散貨，沒有櫃型)。掃描頁 OCR 偶爾把 HBL 號碼中間多切一個空格、尾巴幾位數
# 字漏在櫃型數字前面 (例如 "OSHKMAA260901 44 3X40' HQ" 應該是
# "OSHKMAA26090144 3X40' HQ")，用 blext 這個可選群組接住 (只有後面緊接著
# 「數字+X」的櫃型寫法時才當成號碼的延續，避免跟「5 PACKAGE(S)」這種真的
# 只有件數、沒有被切斷的版面搞混)。
_CH01_HBL_RE = re.compile(
    r"HBL INFO:\s*(?P<bl>\S+)(?:\s+(?P<blext>\d+)(?=\s+\d+\s*X))?\s+"
    r"(?P<boxinfo>.+?)\s+KGS\s+(?P<cbm>[\d]+(?:[.,] ?\d+)?)\s*CBM")
# 櫃型後面的符號 OCR 常認錯：直引號 ' / 彎引號 ' / 度symbol ° 都可能出現，
# 中間也可能多一個空格 ("40°HQ"、"40' HQ")，一律還原成標準寫法 "40'HQ"。
_CH01_BOX_RE = re.compile(r"^(\d+)\s*X\s*(\d+)\s*[\'’°]?\s*([A-Z]+)")

# 到達日 (ETA) 可能空白 (沒有印日期，照原文填 N/A)，日期前後的連字號掃描頁
# 偶爾被 OCR 認成全角/em dash，一律用 _DASH 類別接住。
_DASH = r"[-‐‑‒–—―]"
_CH01_PORT_RE = re.compile(
    rf"PORT OF LOADING/ETD\s*:\s*(?P<pol>.+?)\s*/\s*\d{{1,2}}{_DASH}[A-Za-z]{{3}}{_DASH}\d{{4}}\s+"
    rf"PORT OF DISCHARGE/ETA\s*:\s*(?P<pod>.+?)"
    rf"(?:\s*/\s*(?P<arrive>\d{{1,2}}{_DASH}[A-Za-z]{{3}}{_DASH}\d{{4}}))?(?=\n|$)")
_CH01_DATE_RE = re.compile(rf"(\d{{1,2}}){_DASH}([A-Za-z]{{3}}){_DASH}(\d{{4}})")

# 貨櫃號碼：第 1 個在 "CONTAINER:" 這一行，其餘各自一行、每行開頭就是
# 「4 個字母 + 7 個數字」的標準櫃號 (ISO 6346)，直到不符合格式的下一行
# (例如 "DESCRIPTION AMOUNT") 才停止。多個櫃用「, 」連接。
_CH01_CONTAINER_LINE_RE = re.compile(r"^([A-Z]{4}\d{7})\s*/")


def _ch01_containers(text: str) -> Optional[str]:
    m = re.search(r"CONTAINER:\s*([A-Z]{4}\d{7})", text)
    if not m:
        return None
    codes = [m.group(1)]
    nl = text.find("\n", m.end())
    rest = text[nl + 1:] if nl != -1 else ""
    for line in rest.splitlines():
        line_m = _CH01_CONTAINER_LINE_RE.match(line.strip())
        if not line_m:
            break
        codes.append(line_m.group(1))
    return ", ".join(codes)


def _ch01_volume(text: str) -> Optional[str]:
    """有櫃型 (例 1X20'GP) 時寫「1 x 20'GP / 21.67 CBM」，散貨只寫 CBM
    (例「3.12 CBM」)，跟柬埔寨 KH01 的寫法一樣 (使用者規則文件指定格式，
    這家供應商不適用一般『FCL 數量X櫃型 / LCL CBM』二選一的材積慣例)。"""
    m = _CH01_HBL_RE.search(text)
    if not m:
        return None
    cbm = _ch01_trim(m.group("cbm"))
    box_m = _CH01_BOX_RE.match(m.group("boxinfo").strip())
    if box_m and cbm:
        count, size, ctype = box_m.groups()
        return f"{int(count)} x {size}'{ctype} / {cbm} CBM"
    return f"{cbm} CBM" if cbm else None


def parse_ch01(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {
        "supplier": ("ORIENTAL LOGISTICS GROUP LTD." if "ORIENTAL LOGISTICS GROUP LTD" in text
                     else "SALEM LOGISTICS COMPANY LTD."),
    }
    header["invoice_no"] = _search(r"INVOICE NO\s*:\s*(\S+)", text)

    date_m = re.search(rf"ISSUE DATE\s*:\s*{_CH01_DATE_RE.pattern}", text)
    header["invoice_date"] = _format_date_en_from_month_name(*date_m.groups()) if date_m else None

    # NAME & ADDRESS 公司名稱跨 2 行、跟右欄的 INVOICE NO/ISSUE DATE 疊在
    # 同一行 (雙欄合併)，先取整段再把右欄的字樣拿掉、剩下的字合併成一行。
    name_m = re.search(r"NAME & ADDRESS:(.*?)CST REF NO", text, re.DOTALL)
    if name_m:
        block = re.sub(r"INVOICE NO\s*:\s*\S+", "", name_m.group(1))
        block = re.sub(r"ISSUE DATE\s*:\s*\S+", "", block)
        header["consignee"] = " ".join(block.split()) or None
    else:
        header["consignee"] = None

    order_no = _search(r"JOB NO\s*:\s*(\S+)", text)
    # 掃描頁 OCR 偶爾把開頭的 "S" 認成節符號 "§" (2026-11-26 答案)
    header["order_no"] = order_no.replace("§", "S") if order_no else None
    header["vessel"] = _search(r"VESSEL/VOYAGE\s*:\s*(.+?)\s+ON BOARD DATE", text)

    onboard_m = re.search(rf"ON BOARD DATE\s*:\s*{_CH01_DATE_RE.pattern}", text)
    header["onboard_date"] = _format_date_en_from_month_name(*onboard_m.groups()) if onboard_m else None

    header["arrive_date"] = "N/A"
    port_m = _CH01_PORT_RE.search(text)
    if port_m:
        header["port_of_loading"] = port_m.group("pol").strip()
        header["port_of_discharge"] = port_m.group("pod").strip()
        if port_m.group("arrive"):
            arrive_m = _CH01_DATE_RE.match(port_m.group("arrive"))
            header["arrive_date"] = _format_date_en_from_month_name(*arrive_m.groups()) if arrive_m else "N/A"

    # MBL NO 空白時後面緊接著 DESTINATION 標籤 (不是真的號碼)，排除掉
    header["mbl_no"] = _search(r"MBL NO\s*:\s*(?!DESTINATION)(\S+)", text) or "N/A"
    bl_m = _CH01_HBL_RE.search(text)
    header["bl_no"] = (bl_m.group("bl") + (bl_m.group("blext") or "")) if bl_m else None
    header["container_no"] = _ch01_containers(text)
    header["volume"] = _ch01_volume(text)
    header["shipper"] = "N/A"  # 帳單沒有 Shipper (使用者規則文件確認)

    items: List[Dict] = []
    for m in _CH01_ITEM_RE.finditer(text):
        desc = m.group("desc").strip()
        desc = _CH01_KNOWN_OCR_DESC.get(desc, desc)
        items.append({"description": desc, "amount": _ch01_num(m.group("amt")), "vat": None})
    return {"header": header, "items": items or [{"description": None, "amount": None, "vat": None}]}


def _ch01_total(text: str) -> Optional[float]:
    return _ch01_num(_search(rf"^TOTAL:\s*USD\s*({_CH01_NUM_RE})", text))


register_supplier(
    "CH01", "SALEM LOGISTICS COMPANY LTD.", detect_ch01, parse_ch01,
    extract_total=_ch01_total,
    ocr_psm=6,
)
