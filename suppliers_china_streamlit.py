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
# 供應商 CH01：SALEM LOGISTICS COMPANY LTD. (香港，SALESMAN 為 ORIENTAL POWER
# LOGISTICS CO LTD)，美金 INVOICE
# ===========================================================================
#
# 版面跟越南 OPL 的美金 DEBIT NOTE (ORIENTAL_USD) 很像 (NAME & ADDRESS/HBL
# INFO/CONTAINER 等欄位命名相同)，但頁首是「SALEM LOGISTICS COMPANY LTD.」
# ＋「INVOICE」(不是 DEBIT NOTE)，用公司名稱辨認即可，不會跟 ORIENTAL_USD
# 混到 (ORIENTAL_USD 的 detect 要求文字裡有 "DEBIT NOTE")。
#
# 帳單幣別固定 USD (TOTAL: USD)，沒有 VAT 欄。費用列原幣可能是 HKD 或 USD，
# 每列各自有匯率換算成 USD，只需要抓最右邊換算後的 USD 金額 (原幣/匯率不填，
# 使用者規則文件 CH01_Claude_Code_修改說明.md 確認)。
#
# NAME & ADDRESS 是雙欄版面：「公司名稱」跟右欄的 INVOICE NO/ISSUE DATE
# 同一行印出，公司名稱常跨 2 行，第 2 行 (TAIWAN BRANCH) 才整行乾淨，見
# _parse_ch01() 的 consignee 處理。


def detect_ch01(text: str) -> bool:
    upper = text.upper()
    return "SALEM LOGISTICS COMPANY LTD" in upper and "INVOICE" in upper


def _ch01_num(value: Optional[str]):
    """'219.84'/'2,000.000' -> 數字 (整數時回傳 int)。"""
    if not value:
        return None
    try:
        num = float(value.replace(",", ""))
    except ValueError:
        return None
    return int(num) if num == int(num) else num


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
    r"^(?P<desc>.+?)\s+[A-Z]{3}\s+[\d.,]+\s*/\s*\S+\s+x\s+[\d.,]+\s+[\d.]+\s+"
    r"USD\s+(?P<amt>[\d.,]+)\s*$", re.MULTILINE)

# "HBL INFO: HKGJKT25120890 1X20'GP 37 PACKAGE(S) 5917.800 KGS 21.670 CBM"
# (有櫃型) 或 "HBL INFO: HKGJKT25120888 18 PACKAGE(S) 770.700 KGS 3.120 CBM"
# (散貨，沒有櫃型)。
_CH01_HBL_RE = re.compile(r"HBL INFO:\s*(\S+)\s+(.+?)\s+KGS\s+([\d.]+)\s*CBM")
_CH01_BOX_RE = re.compile(r"^(\d+)X(\d+'[A-Z]+)")

_CH01_PORT_RE = re.compile(
    r"PORT OF LOADING/ETD\s*:\s*(.+?)/\d{1,2}-[A-Za-z]{3}-\d{4}\s+"
    r"PORT OF DISCHARGE/ETA\s*:\s*(.+?)/(\d{1,2}-[A-Za-z]{3}-\d{4})")


def _ch01_volume(text: str) -> Optional[str]:
    """有櫃型 (例 1X20'GP) 時寫「1 x 20'GP / 21.67 CBM」，散貨只寫 CBM
    (例「3.12 CBM」)，跟柬埔寨 KH01 的寫法一樣 (使用者規則文件指定格式，
    這家供應商不適用一般『FCL 數量X櫃型 / LCL CBM』二選一的材積慣例)。"""
    m = _CH01_HBL_RE.search(text)
    if not m:
        return None
    cbm = _ch01_trim(m.group(3))
    box_m = _CH01_BOX_RE.match(m.group(2).strip())
    if box_m and cbm:
        return f"{int(box_m.group(1))} x {box_m.group(2)} / {cbm} CBM"
    return f"{cbm} CBM" if cbm else None


def parse_ch01(text: str) -> Dict:
    header: Dict[str, Optional[str]] = {"supplier": "SALEM LOGISTICS COMPANY LTD."}
    header["invoice_no"] = _search(r"INVOICE NO\s*:\s*(\S+)", text)

    date_m = re.search(r"ISSUE DATE\s*:\s*(\d{1,2})-([A-Za-z]{3})-(\d{4})", text)
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

    header["order_no"] = _search(r"JOB NO\s*:\s*(\S+)", text)
    header["vessel"] = _search(r"VESSEL/VOYAGE\s*:\s*(.+?)\s+ON BOARD DATE", text)

    onboard_m = re.search(r"ON BOARD DATE\s*:\s*(\d{1,2})-([A-Za-z]{3})-(\d{4})", text)
    header["onboard_date"] = _format_date_en_from_month_name(*onboard_m.groups()) if onboard_m else None

    port_m = _CH01_PORT_RE.search(text)
    if port_m:
        header["port_of_loading"] = port_m.group(1).strip()
        header["port_of_discharge"] = port_m.group(2).strip()
        arrive_m = re.match(r"(\d{1,2})-([A-Za-z]{3})-(\d{4})", port_m.group(3))
        header["arrive_date"] = _format_date_en_from_month_name(*arrive_m.groups()) if arrive_m else None

    # MBL NO 空白時後面緊接著 DESTINATION 標籤 (不是真的號碼)，排除掉
    header["mbl_no"] = _search(r"MBL NO\s*:\s*(?!DESTINATION)(\S+)", text) or "N/A"
    header["bl_no"] = _search(r"HBL INFO:\s*(\S+)", text)
    header["container_no"] = _search(r"CONTAINER:\s*([A-Z0-9]+)", text)
    header["volume"] = _ch01_volume(text)
    header["shipper"] = "N/A"  # 帳單沒有 Shipper (使用者規則文件確認)

    items: List[Dict] = [
        {"description": m.group("desc").strip(), "amount": _ch01_num(m.group("amt")), "vat": None}
        for m in _CH01_ITEM_RE.finditer(text)
    ]
    return {"header": header, "items": items or [{"description": None, "amount": None, "vat": None}]}


def _ch01_total(text: str) -> Optional[float]:
    return _ch01_num(_search(r"^TOTAL:\s*USD\s*([\d.,]+)", text))


register_supplier(
    "CH01", "SALEM LOGISTICS COMPANY LTD.", detect_ch01, parse_ch01,
    extract_total=_ch01_total,
)
