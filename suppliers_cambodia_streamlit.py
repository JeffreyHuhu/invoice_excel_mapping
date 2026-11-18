# -*- coding: utf-8 -*-
"""
suppliers_cambodia_streamlit.py
================================
柬埔寨供應商 (KH01~) 的 detect_xxx()/parse_xxx()，用 register_supplier()
註冊到 extract_utils_streamlit.SUPPLIER_REGISTRY。寫法仿照
suppliers_indonesia_streamlit.py / suppliers_vietnam_streamlit.py。

跟印尼/越南供應商共用 suppliers_common_streamlit.py 的工具函式。
"""

import re
from typing import Dict, List, Optional

from extract_utils_streamlit import register_supplier
from suppliers_common_streamlit import _search, _format_date_en_from_month_name


# ===========================================================================
# 供應商 KH01：ORIENTAL LOGISTICS & DISTRIBUTION (CAMBODIA) CO., LTD
# ===========================================================================
#
# 版面：高棉文與英文並列的 TAX INVOICE，高棉文字型缺字會變成亂碼
# (pdfplumber 讀出 (cid:NN))，只取英文標籤/數值。公司名稱本身只在
# logo 圖片裡 (文字層沒有)，固定填使用者確認的英文全名，不要從文字擷取。
# 幣別固定 USD (使用者 2025 規則文件確認)。
#
# 一份 PDF 可能有好幾張各自獨立的 TAX INVOICE (不同 TAX INV NO.)；同一張
# 可能跨兩頁 (Page: 1/2、2/2)，續頁沒有 HBL INFO/Container No，沿用第 1
# 頁 (見 _kh01_invoice_blocks())。Grand Total Charge 只印在最後一頁。

# 識別用的公司稅籙號碼 (Our VAT)，所有樣本一致，比高棉文公司名稱穩定。
_KH01_TAX_ID = "L001-103007369"


def detect_kh01(text: str) -> bool:
    return "TAX INVOICE" in text.upper() and _KH01_TAX_ID in text


def _kh01_num(value: Optional[str]):
    """'271.00'/'1,018' -> 數字 (整數時回傳 int)。"""
    if not value:
        return None
    try:
        num = float(value.replace(",", ""))
    except ValueError:
        return None
    return int(num) if num == int(num) else num


def _kh01_trim(value: Optional[str]) -> Optional[str]:
    """CBM 數字去掉多餘的小數尾數 0 ('26.100' -> '26.1'、'1.018' -> '1.018')。"""
    num = _kh01_num(value)
    if num is None:
        return None
    return f"{num:g}" if isinstance(num, float) else str(num)


# 一張 TAX INVOICE 的開頭一定有 "Page: N/M"，拿這個天然就存在的分頁點切開
# 整份 PDF 的文字 (不用另外註冊 page_text)。
_KH01_PAGE_RE = re.compile(r"(?=Page:\s*\d+\s*/\s*\d+)")
_KH01_TAX_INV_NO_RE = re.compile(r"TAX INV NO\.\s*:\s*(\S+)")
_KH01_GRAND_TOTAL_RE = re.compile(
    r"Grand Total Charge\s*:\s*([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)")
# 費用列："名稱 數量 x 單位 USD 單價 USD 金額 VAT 含稅小計"，例如
# "IMPORT CLEARANCE 1.000 x SET USD 90.000 USD 90.00 9.00 99.00"、
# "THC FEE 3.640 x CBM USD 9.000 USD 32.76 3.28 36.04"。
# 名稱可能帶括號/連字號，例如 "DOCUMENT (B/L) FEE"、"X-RAY"；也可能大小寫
# 混用，例如 "EIB Fee"、"filing Fee" (2026-11-18)。
_KH01_ITEM_RE = re.compile(
    r"^(?P<desc>[A-Za-z][A-Za-z0-9 ()/\-]*?)\s+(?P<qty>[\d.,]+)\s*x\s*(?P<unit>\S+)\s+"
    r"USD\s+(?P<price>[\d.,]+)\s+USD\s+(?P<amount>[\d.,]+)\s+"
    r"(?P<vat>[\d.,]+)\s+(?P<total>[\d.,]+)\s*$")
_KH01_HBL_RE = re.compile(r"HBL INFO:\s*(\S+)\s+(.+)")
_KH01_CONTAINER_RE = re.compile(r"Container No\s*:\s*([A-Z0-9]+)\s*//\s*([^;\n]+)")
_KH01_CBM_RE = re.compile(r"KGS\s+([\d,.]+)\s*CBM")
_KH01_FCL_RE = re.compile(r"^(\d+)X(\d+'?[A-Z]+)$")


def _kh01_invoice_blocks(text: str) -> List[str]:
    """依 'TAX INV NO.' 把連續同號的頁面合併成一個區塊，這樣續頁 (沒有
    HBL INFO/Container No) 也能在合併後的區塊文字裡找到第 1 頁的值。"""
    pages = [p for p in _KH01_PAGE_RE.split(text) if p.strip()]
    if not pages:
        pages = [text]
    blocks: List[str] = []
    cur_no: Optional[str] = None
    buf: List[str] = []
    for page in pages:
        m = _KH01_TAX_INV_NO_RE.search(page)
        no = m.group(1) if m else None
        if no and no != cur_no and buf:
            blocks.append("\n".join(buf))
            buf = []
        cur_no = no or cur_no
        buf.append(page)
    if buf:
        blocks.append("\n".join(buf))
    return blocks


def _kh01_volume_bl_container(block: str):
    """回傳 (B/L NO., Volume, Container No.)。整櫃 (HBL INFO 帶 "1X40'HQ")
    寫「1 x 40'HQ / 26.1 CBM」；散貨只有 CBM 數字時寫「1.018 CBM」。"""
    m = _KH01_HBL_RE.search(block)
    bl_no = m.group(1) if m else None
    rest = m.group(2).strip() if m else ""
    cbm_m = _KH01_CBM_RE.search(block)
    cbm = _kh01_trim(cbm_m.group(1)) if cbm_m else None
    fcl_m = _KH01_FCL_RE.match(rest.split()[0]) if rest else None
    if fcl_m and cbm:
        volume = f"{int(fcl_m.group(1))} x {fcl_m.group(2)} / {cbm} CBM"
    elif cbm:
        volume = f"{cbm} CBM"
    else:
        volume = None
    container_m = _KH01_CONTAINER_RE.search(block)
    container = container_m.group(1) if container_m else None
    return bl_no, volume, container


def _parse_kh01_block(block: str) -> List[Dict]:
    no_m = _KH01_TAX_INV_NO_RE.search(block)
    date_m = re.search(r"/DATE\s*:\s*(\d{1,2})-([A-Za-z]{3})-(\d{4})", block)
    eta_m = re.search(r"ETA\s*:\s*(\d{1,2})-([A-Za-z]{3})-(\d{4})", block)
    mbl_m = re.search(r"MBL NO\s*:\s*(\S*)", block)
    consignee_m = re.search(r"^(.+?)\s+JOB NO\s*:\s*(\S+)", block, re.MULTILINE)
    # 用行首 "^:" 錨定，避免 \s* 跨行吃到更前面不相關的 ":" (例如 "Ex. RATE :")
    vessel_m = re.search(r"^:\s*([^\n:]+?)\s*$\n\s*VESSEL/VOYAGE", block, re.MULTILINE)
    port_m = re.search(r"^:\s*([^\n:]+?)\s*:\s*([^\n]+?)\s*$\n\s*FROM\s+DESTINATION", block, re.MULTILINE)
    bl_no, volume, container_no = _kh01_volume_bl_container(block)

    header = {
        "invoice_no": no_m.group(1) if no_m else None,
        "invoice_date": _format_date_en_from_month_name(*date_m.groups()) if date_m else None,
        "supplier": "ORIENTAL LOGISTICS & DISTRIBUTION (CAMBODIA) CO., LTD",
        "consignee": consignee_m.group(1).strip() if consignee_m else None,
        "order_no": consignee_m.group(2) if consignee_m else None,
        "arrive_date": _format_date_en_from_month_name(*eta_m.groups()) if eta_m else None,
        "onboard_date": "N/A",
        "bl_no": bl_no,
        "mbl_no": (mbl_m.group(1) if mbl_m and mbl_m.group(1) else "N/A"),
        "port_of_loading": port_m.group(1).strip() if port_m else None,
        "port_of_discharge": port_m.group(2).strip() if port_m else None,
        "volume": volume,
        "vessel": vessel_m.group(1).strip() if vessel_m else None,
        "container_no": container_no,
    }

    items = []
    for line in block.splitlines():
        m = _KH01_ITEM_RE.match(line.strip())
        if not m:
            continue
        items.append({
            "description": m.group("desc").strip(),
            "amount": _kh01_num(m.group("amount")),
            "vat": _kh01_num(m.group("vat")),
        })
    if not items:
        items = [{}]
    return [{**header, **it} for it in items]


def parse_kh01(text: str) -> Dict:
    rows = []
    for block in _kh01_invoice_blocks(text):
        rows.extend(_parse_kh01_block(block))
    return {"header": {}, "items": rows}


def _kh01_totals(text: str, which: str) -> Dict[str, float]:
    """{TAX INV NO.: 總額}；which='net' 取 Grand Total Charge 的 AMOUNT
    小計 (未稅)，which='gross' 取第 3 個數字 (含稅總額)。"""
    out: Dict[str, float] = {}
    for block in _kh01_invoice_blocks(text):
        no_m = _KH01_TAX_INV_NO_RE.search(block)
        total_m = _KH01_GRAND_TOTAL_RE.search(block)
        if no_m and total_m:
            out[no_m.group(1)] = _kh01_num(
                total_m.group(1) if which == "net" else total_m.group(3))
    return out


register_supplier(
    "KH01", "ORIENTAL LOGISTICS & DISTRIBUTION (CAMBODIA) CO., LTD",
    detect_kh01, parse_kh01,
    extract_total=lambda t: _kh01_totals(t, "net"),
    extract_grand_total=lambda t: _kh01_totals(t, "gross"),
)
