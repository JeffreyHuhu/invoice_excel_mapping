# -*- coding: utf-8 -*-
"""
extract_utils.py
=================
供應商帳單 PDF -> 結構化資料 解析模組 (全新重建版)

這個版本鎖定 KUEHNE NAGEL 這種帳單格式：同一份 PDF 裡包含
  - 第1頁：印尼稅務局格式的「Faktur Pajak」(稅務發票)
  - 第2頁起：Kuehne Nagel 自己的「KN Sales Invoice」

兩份文件內容互補：供應商名稱／費用名稱(Description)／金額(Amount) 取自
Faktur Pajak；其餘船務相關欄位 (Invoice No/Date、日期、港口、船名、追蹤
號碼...) 取自 Sales Invoice。

核心設計：
  1. extract_text_from_pdf()：用 pdfplumber 抽取文字，可傳入不同的
     x_tolerance / y_tolerance / layout 參數。
  2. parse_kn_invoice() + build_row()：用正規表示式抓出每個欄位，抓不到
     的欄位一律填 'N/A' (使用者需求)。
  3. EXTRACTION_CONFIGS + find_best_extraction()：對同一份 PDF「重複執行」
     多組擷取參數，每一組都拿去跟使用者上傳的正確答案 Excel 逐欄比對算出
     正確率，一旦達到 100% 就提早停止，沒有的話就回傳嘗試過最高的分數。
  4. compare_row() / compute_accuracy()：逐欄比對邏輯，回傳比對明細與正確率。
"""

import re
import io
import datetime
from typing import Dict, List, Optional, Tuple

import pdfplumber
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

NA = "N/A"  # 抓不到值時，一律填這個 (使用者明確要求)


# ---------------------------------------------------------------------------
# 1. 欄位定義 (對應 KUEHNE_NAGEL_資料模板.xlsx 的 A~P 欄，欄位順序完全一致)
# ---------------------------------------------------------------------------

FIELD_CODES: List[str] = [
    "invoice_date", "invoice_no", "supplier", "consignee", "description",
    "amount", "order_no", "arrive_date", "onboard_date", "bl_no", "mbl_no",
    "port_of_loading", "port_of_discharge", "volume", "vessel", "container_no",
]

# Excel 範本第2列的顯示欄名，逐字對照 KUEHNE_NAGEL_資料模板.xlsx
DISPLAY_HEADERS: Dict[str, str] = {
    "invoice_date":      "INVOICE DATE\n發票日期",
    "invoice_no":        "INVOICE NO\n發票號碼",
    "supplier":          "SUPPLIER\n供應商名稱",
    "consignee":         "CONSIGNEE\n集團公司名稱",
    "description":       "Description\n費用名稱",
    "amount":            "AMOUNT\n金額",
    "order_no":          "Order No.",
    "arrive_date":       "Arrive Date\n到達日",
    "onboard_date":      "on board date\n開船日",
    "bl_no":             "B/L NO.\n提單號碼",
    "mbl_no":            "MBL NO.\n主提單號碼",
    "port_of_loading":   "Port of Loading\n啟運港",
    "port_of_discharge": "Port of discharge\n目的港",
    "volume":            "Volume\n材積",
    "vessel":             "Vessel\n船名",
    "container_no":      "Container no.\n貨櫃號碼",
}


# ---------------------------------------------------------------------------
# 2. PDF 文字擷取 (可調參數，供迴圈嘗試不同設定用)
# ---------------------------------------------------------------------------

def extract_text_from_pdf(pdf_path: str, **pdfplumber_kwargs) -> str:
    """讀取 PDF 全部頁面文字並合併成一個字串。

    pdfplumber_kwargs 會轉給 page.extract_text()，例如 x_tolerance、
    y_tolerance、layout。不同 PDF 的版面留白不一定用同一組參數解析最準，
    這裡開放參數化讓 find_best_extraction() 可以多試幾組。
    """
    parts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            parts.append(page.extract_text(**pdfplumber_kwargs) or "")
    return "\n".join(parts)


# 一組一組不同的擷取參數，讓「迴圈找最佳解析結果」有得選
EXTRACTION_CONFIGS: List[Dict] = [
    {},                                    # pdfplumber 預設值
    {"x_tolerance": 1, "y_tolerance": 1},
    {"x_tolerance": 1, "y_tolerance": 3},
    {"x_tolerance": 2, "y_tolerance": 3},
    {"x_tolerance": 3, "y_tolerance": 5},
    {"layout": True},                       # 盡量保留原始版面座標間距
]


# ---------------------------------------------------------------------------
# 3. 格式清理小工具
# ---------------------------------------------------------------------------

def clean_amount_id_style(text: Optional[str]) -> Optional[int]:
    """印尼式數字 '1.521.098,00' (句點=千分位、逗號=小數) -> 1521098 (int)。"""
    if not text:
        return None
    text = text.strip().replace(".", "")
    text = text.split(",")[0]
    return int(text) if text.isdigit() else None


def clean_port(value: Optional[str]) -> Optional[str]:
    """'LOS ANGELES, CA' -> 'LOS ANGELES'：只取城市名，逗號後的州/國碼截掉。
    沒有逗號的港口 (例如 'JAKARTA') 原樣保留。
    """
    if not value:
        return value
    return value.split(",")[0].strip()


def clean_tracking_no(value: Optional[str]) -> Optional[str]:
    """'1076 510 528' -> '1076510528'：把數字群組間的空白去掉。"""
    if not value:
        return value
    return re.sub(r"\s+", "", value.strip())


def clean_volume(text: Optional[str]) -> Optional[float]:
    """'11.810' -> 11.81 (float)。"""
    if not text:
        return None
    text = text.strip()
    try:
        return float(text)
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# 4. 正規表示式規則
# ---------------------------------------------------------------------------
#
# KN Sales Invoice 是左右兩欄並排的版面 (例如同一列左邊是
# "P. OF LOADING : JAKARTA"，右邊緊接著 "ETD/ATD : 23.08.2026")。
# pdfplumber 抽取文字時，同一 y 座標高度的內容會被接成同一行字串，所以
# 每個欄位都用「遇到下一個已知欄位關鍵字、或連續兩個以上空白、或換行
# 就停止擷取」的方式抓值，避免把右欄內容一起抓進來。
#
# 註：第2頁左側有一段直式浮水印 ("ORIGINAL") 文字，pdfplumber 抽取時會
# 跟旁邊的免責聲明段落攪在一起、順序顛倒，但我們需要的欄位都在那段亂碼
# 之後，不受影響。

_LABEL_WORDS = [
    r"INVOICE TO", r"SALES INVOICE", r"KN TRACKING NO\.?", r"KN ACCOUNTING NO\.?",
    r"ACCOUNT NO\.?", r"INVOICE NO\.?\s*/\s*DATE", r"DUE DATE", r"CONSIGNEE", r"SHIPPER",
    r"NOTIFY", r"FIRST VESSEL NA", r"VESSEL NAME", r"FIRST VOYAGE NU", r"VOYAGE",
    r"PL\.\s*OF RECEIPT", r"MOVEMENT TYPE", r"P\.\s*OF LOADING", r"ETD/ATD",
    r"P\.\s*OF DISCHARGE", r"ETA/ATA", r"PL\.\s*OF DELIVERY", r"DANGEROUS GOODS",
    r"TERMS OF TRADE", r"INSURAN\.\s*STATUS", r"MARKS\s*&\s*NOS", r"CHARGE",
    r"GTN SHIPPING ORDER NUMBER", r"COMMERCIAL INVOICE NUMBER", r"ESI TOKEN",
]
_NEXT_LABEL_ALT = "(?:" + "|".join(_LABEL_WORDS) + ")"
_STOP = rf"(?=\s{{2,}}|\n|$|\s*{_NEXT_LABEL_ALT}\b)"


def _search(pattern: str, text: str) -> Optional[str]:
    m = re.search(pattern, text, flags=re.MULTILINE)
    if not m:
        return None
    return m.group(1).strip().rstrip(",").strip()


def parse_kn_invoice(text: str) -> Dict[str, Optional[str]]:
    """從 PDF 全文抓出每個欄位的「原始」字串 (還沒做格式清理、還沒補 N/A)。"""
    raw: Dict[str, Optional[str]] = {}

    # --- 來自 Faktur Pajak (稅務發票，第1頁) ---
    raw["supplier"] = _search(
        r"Pengusaha Kena Pajak:\s*\n\s*Nama\s*:\s*(.+)", text
    )
    raw["description"] = _search(
        r"\n([A-Za-z][^\n]+)\nRp\s+[\d.,]+\s*x", text
    )
    raw["amount_raw"] = _search(
        r"Harga Jual\s*/\s*Penggantian\s*/\s*Uang Muka\s*/\s*Termin\s+([\d.,]+)", text
    )

    # --- 來自 KN Sales Invoice (第2頁起) ---
    raw["consignee"] = _search(rf"INVOICE TO\s*\n\s*(.+?){_STOP}", text)

    m = re.search(r"INVOICE NO\.\s*/\s*DATE\s+(\S+)\s+(\d{1,2}\.\d{1,2}\.\d{4})", text)
    if m:
        raw["invoice_no"], raw["invoice_date"] = m.group(1), m.group(2)
    else:
        raw["invoice_no"], raw["invoice_date"] = None, None

    raw["bl_no_raw"] = _search(rf"KN TRACKING NO\.\s*(.+?){_STOP}", text)
    raw["order_no"] = _search(rf"GTN SHIPPING ORDER NUMBER\s*\n\s*(.+?){_STOP}", text)
    raw["vessel"] = _search(rf"VESSEL NAME\s*:\s*(.+?){_STOP}", text)
    raw["port_of_loading_raw"] = _search(rf"P\.\s*OF LOADING\s*:\s*(.+?){_STOP}", text)
    raw["port_of_discharge_raw"] = _search(rf"P\.\s*OF DISCHARGE\s*:\s*(.+?){_STOP}", text)
    raw["onboard_date"] = _search(rf"ETD/ATD\s*:\s*(.+?){_STOP}", text)
    raw["arrive_date"] = _search(rf"ETA/ATA\s*:\s*(.+?){_STOP}", text)

    m2 = re.search(r"AS PER ATTACHED\s+([\d,\.]+)\s+([\d,\.]+)", text)
    raw["volume_raw"] = m2.group(2) if m2 else None

    # --- 這種帳單本來就沒有海運提單/貨櫃資訊，直接標記缺漏 ---
    raw["mbl_no"] = None
    raw["container_no"] = None

    return raw


def build_row(raw: Dict[str, Optional[str]]) -> Dict[str, str]:
    """把 parse_kn_invoice() 的原始結果做格式清理，抓不到的欄位補 'N/A'。"""
    row = {
        "invoice_date": raw.get("invoice_date"),
        "invoice_no": raw.get("invoice_no"),
        "supplier": raw.get("supplier"),
        "consignee": raw.get("consignee"),
        "description": raw.get("description"),
        "amount": clean_amount_id_style(raw.get("amount_raw")),
        "order_no": raw.get("order_no"),
        "arrive_date": raw.get("arrive_date"),
        "onboard_date": raw.get("onboard_date"),
        "bl_no": clean_tracking_no(raw.get("bl_no_raw")),
        "mbl_no": raw.get("mbl_no"),
        "port_of_loading": clean_port(raw.get("port_of_loading_raw")),
        "port_of_discharge": clean_port(raw.get("port_of_discharge_raw")),
        "volume": clean_volume(raw.get("volume_raw")),
        "vessel": raw.get("vessel"),
        "container_no": raw.get("container_no"),
    }
    for k, v in row.items():
        if v is None or (isinstance(v, str) and not v.strip()):
            row[k] = NA
    return row


# ---------------------------------------------------------------------------
# 5. 讀取「正確答案」Excel (欄位順序固定同 FIELD_CODES，起始列/欄自動偵測)
# ---------------------------------------------------------------------------

_HEADER_ANCHOR = "INVOICE DATE"  # 用這個關鍵字定位表頭在哪一列/哪一欄


def read_reference_excel(source) -> pd.DataFrame:
    """source 可以是檔案路徑 (str) 或檔案 bytes。
    自動找出含有 'INVOICE DATE' 字樣的儲存格當作表頭錨點，藉此定位資料
    真正的起始列與起始欄，不用替不同版本的範本各寫一次死板的座標。
    """
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(source)
    wb = load_workbook(source, data_only=True)
    ws = wb.active

    header_row, start_col = None, None
    for r in range(1, 5):
        for c in range(1, 5):
            val = ws.cell(row=r, column=c).value
            if val and _HEADER_ANCHOR in str(val).upper():
                header_row, start_col = r, c
                break
        if header_row:
            break
    if header_row is None:
        header_row, start_col = 2, 1  # 找不到就退回本範本預設位置 (A2)

    records = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=False):
        rec = {}
        empty = True
        for idx, code in enumerate(FIELD_CODES):
            cell = row[start_col - 1 + idx]
            rec[code] = cell.value
            if cell.value not in (None, ""):
                empty = False
        if not empty:
            records.append(rec)
    return pd.DataFrame(records)


# ---------------------------------------------------------------------------
# 6. 比對邏輯：擷取結果 vs 正確答案
# ---------------------------------------------------------------------------

def normalize_value(value) -> str:
    """統一轉成好比對的字串：去除空白差異、大小寫差異、千分位逗號。"""
    if value is None:
        return ""
    if isinstance(value, (datetime.date, datetime.datetime)):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).strip().upper()
    text = re.sub(r"\s+", " ", text)
    return text.rstrip(",")


def compare_row(extracted: Dict[str, str], reference: Optional[Dict]) -> List[Dict]:
    """逐欄比對擷取結果與正確答案，回傳每個欄位的比對紀錄 (供網頁紅綠燈顯示)。"""
    results = []
    for code in FIELD_CODES:
        ext_val = extracted.get(code, NA)
        ref_val = (reference or {}).get(code, NA)
        match = normalize_value(ext_val) == normalize_value(ref_val)
        results.append({
            "欄位": DISPLAY_HEADERS[code].split("\n")[0],
            "擷取結果": ext_val,
            "正確答案": ref_val,
            "結果": "✅ 相符" if match else "❌ 不相符",
        })
    return results


def compute_accuracy(results: List[Dict]) -> float:
    if not results:
        return 0.0
    matched = sum(1 for r in results if r["結果"].startswith("✅"))
    return matched / len(results)


def quality_score(row: Dict[str, str]) -> float:
    """沒有正確答案 Excel 時的自我檢查分數：欄位有抓到值(不是N/A)的比例。
    僅供迴圈挑選較好的擷取設定用，不是跟人工核對過的正確率。
    """
    if not row:
        return 0.0
    filled = sum(1 for v in row.values() if v and v != NA)
    return filled / len(row)


# ---------------------------------------------------------------------------
# 7. 迴圈：反覆嘗試不同擷取設定，直到 100% 正確率 (或找出目前能達到的最高值)
# ---------------------------------------------------------------------------

def find_best_extraction(
    pdf_path: str,
    reference: Optional[Dict] = None,
    configs: Optional[List[Dict]] = None,
) -> Tuple[Dict[str, str], float, Dict, Optional[List[Dict]], List[Dict]]:
    """對同一份 PDF 用多組擷取參數各解析一次，每次都評分，
    一旦達到 100% 正確率就提早停止；跑完所有設定仍未達 100% 的話，
    回傳嘗試過的最高分數與對應結果。

    回傳：(最佳row, 最佳分數0~1, 使用的設定, 該設定的逐欄比對明細或None, 全部嘗試紀錄)
    """
    configs = configs or EXTRACTION_CONFIGS
    attempts: List[Dict] = []
    best_row, best_score, best_cfg, best_results = None, -1.0, None, None

    for cfg in configs:
        text = extract_text_from_pdf(pdf_path, **cfg)
        raw = parse_kn_invoice(text)
        row = build_row(raw)

        if reference is not None:
            results = compare_row(row, reference)
            score = compute_accuracy(results)
        else:
            results = None
            score = quality_score(row)

        attempts.append({"設定": cfg or "預設", "分數(%)": round(score * 100, 1)})

        if score > best_score:
            best_row, best_score, best_cfg, best_results = row, score, cfg, results

        if best_score >= 1.0:
            break  # 已經達到 100% 正確率，不用再嘗試其他設定

    return best_row, best_score, best_cfg, best_results, attempts


# ---------------------------------------------------------------------------
# 8. 輸出 Excel (欄位/表頭格式對齊 KUEHNE_NAGEL_資料模板.xlsx：
#    單一表頭列在第2列、資料從第3列開始、A欄起算)
# ---------------------------------------------------------------------------

def build_excel(rows: List[Dict[str, str]]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "KUEHNE_NAGEL"

    bold = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="DDEBF7")
    wrap_center = Alignment(wrap_text=True, vertical="center", horizontal="center")

    for idx, code in enumerate(FIELD_CODES, start=1):
        cell = ws.cell(row=2, column=idx, value=DISPLAY_HEADERS[code])
        cell.font = bold
        cell.fill = header_fill
        cell.alignment = wrap_center
        ws.column_dimensions[get_column_letter(idx)].width = 20

    for r, row in enumerate(rows, start=3):
        for idx, code in enumerate(FIELD_CODES, start=1):
            ws.cell(row=r, column=idx, value=row.get(code, NA))

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
