# -*- coding: utf-8 -*-
"""
extract_utils.py
=================
多供應商帳單擷取系統：共用核心引擎 (供應商無關)

流程：辨識哪間供應商帳單 → 選取對應的帳單辨識系統 → 擷取欄位 → 跟正確答案比對算正確率

架構：
  - 這支檔案只放「跟供應商無關」的共用邏輯：PDF文字擷取、供應商辨識/註冊表、
    迴圈找最佳擷取設定、比對評分、Excel讀寫。
  - 每個供應商專屬的辨識規則 (detect) 與擷取規則 (parse) 放在 suppliers.py，
    透過 register_supplier() 註冊進本檔案的 SUPPLIER_REGISTRY。
  - 新增第 21 家、22 家...供應商時，只需要在 suppliers.py 多寫一組
    detect()/parse() 並呼叫 register_supplier()，不用改這支檔案，
    也不用改 app.py。

標準欄位 (不管哪個供應商，擷取結果都要對應同一套欄位代碼，這樣才能放進
同一份 Excel、用同一套比對/查核邏輯)：
  invoice_date, invoice_no, supplier, consignee, order_no, arrive_date,
  onboard_date, bl_no, mbl_no, port_of_loading, port_of_discharge, volume,
  vessel, container_no, shipper  → 每張帳單「共用」一次 (抬頭欄位)
  description, amount   → 每筆費用明細各一列 (一張帳單可能有多筆費用)
"""

import re
import io
import os
import sys
import json
import math
import datetime
from collections import OrderedDict
from typing import Callable, Dict, List, Optional, Tuple

import pdfplumber
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

try:
    import pytesseract
    from PIL import ImageOps
    _OCR_AVAILABLE = True
except ImportError:
    pytesseract = None
    ImageOps = None
    _OCR_AVAILABLE = False

# 本機版打包成 .exe (PyInstaller) 時，會把精簡版 Tesseract 一起包進去
# (放在解壓目錄的 tesseract/ 底下，見 D:\資誠local執行\*.spec)，有的話優先
# 使用，讓沒有安裝 Tesseract 的電腦也能辨識掃描檔。雲端 Streamlit 版與
# 直接執行 .py 時不受影響，照舊從 PATH 找系統安裝的 Tesseract。
_BUNDLED_TESSERACT_DIR = os.path.join(getattr(sys, "_MEIPASS", ""), "tesseract")
if pytesseract is not None and getattr(sys, "frozen", False) and \
        os.path.isfile(os.path.join(_BUNDLED_TESSERACT_DIR, "tesseract.exe")):
    pytesseract.pytesseract.tesseract_cmd = os.path.join(_BUNDLED_TESSERACT_DIR, "tesseract.exe")
    os.environ["TESSDATA_PREFIX"] = os.path.join(_BUNDLED_TESSERACT_DIR, "tessdata")

NA = "N/A"  # 抓不到值時，一律填這個 (使用者需求)

# ---------------------------------------------------------------------------
# 1. 標準欄位定義
# ---------------------------------------------------------------------------

HEADER_FIELD_CODES: List[str] = [
    "invoice_date", "invoice_no", "supplier", "consignee", "order_no",
    "arrive_date", "onboard_date", "bl_no", "mbl_no", "port_of_loading",
    "port_of_discharge", "volume", "vessel", "container_no", "shipper",
]
# 注意：vat 是「每筆費用明細各自可能不同」的欄位 (同一張帳單裡，有些費用
# 有 VAT、有些沒有)，所以歸在 ITEM_FIELD_CODES (跟 description/amount 一樣
# 逐筆列出)，不是 HEADER_FIELD_CODES (整張帳單共用一次)。
ITEM_FIELD_CODES: List[str] = ["description", "amount", "vat"]

# 完整欄位順序 (輸出 Excel 用，對應使用者最新範本的欄位順序：
# INVOICE DATE, INVOICE NO, SUPPLIER, CONSIGNEE, Description, AMOUNT, VAT,
# Order No., Arrive Date, on board date, B/L NO., MBL NO., Port of Loading,
# Port of discharge, Volume, Vessel, Container no., Shipper)
FIELD_CODES: List[str] = [
    "invoice_date", "invoice_no", "supplier", "consignee", "description",
    "amount", "vat", "order_no", "arrive_date", "onboard_date", "bl_no",
    "mbl_no", "port_of_loading", "port_of_discharge", "volume", "vessel",
    "container_no", "shipper",
]

# 標準欄位裡「值的語意是日期」的欄位（不管哪個供應商，parse_xxx() 吐出來
# 的字串格式可能不一樣，但這三欄的值都代表一個日期）。使用者要求：轉檔
# 結果 Excel 裡這幾欄一律要是「真正的 Excel 日期儲存格」(不是純文字)，
# 顯示成 "Aug-10-2026" 這種「英文月縮寫-日-年」格式，見 _coerce_to_date()/
# DATE_DISPLAY_NUMBER_FORMAT。
DATE_FIELD_CODES: Tuple[str, ...] = ("invoice_date", "arrive_date", "onboard_date")

# Excel 自訂數字格式字串：mmm=英文月份縮寫 (Jan/Feb/…)、dd=兩位數日、
# yyyy=四位數年，用減號分隔，對應使用者要求的顯示方式 "Aug-10-2026"。
DATE_DISPLAY_NUMBER_FORMAT = "mmm-dd-yyyy"

DISPLAY_HEADERS: Dict[str, str] = {
    "invoice_date":      "INVOICE DATE\n發票日期",
    "invoice_no":        "INVOICE NO\n發票號碼",
    "supplier":          "SUPPLIER\n供應商名稱",
    "consignee":         "CONSIGNEE\n集團公司名稱",
    "description":       "Description\n費用名稱",
    "amount":            "AMOUNT\n金額",
    "vat":               "VAT\n稅額",
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
    "shipper":           "Shipper\n出貨方",
}


# ---------------------------------------------------------------------------
# 2. PDF 文字擷取 (可調參數，供迴圈嘗試不同設定用)
# ---------------------------------------------------------------------------

# 掃描檔 (沒有文字層的 PDF，例如 DHL 帳單是掃描件轉存的 PDF) OCR 快取：
# key 是 (pdf_path, page_index)，value 是 OCR 出來的文字。OCR 本身不受
# pdfplumber_kwargs (x_tolerance 等只對「有文字層」的 PDF 有意義) 影響，
# find_best_extraction() 會用不同的 kwargs 重複呼叫這支函式最多 20 次，
# 用快取避免同一頁被重複 OCR 20 次、白白浪費時間。
_OCR_TEXT_CACHE: Dict[Tuple[str, int], str] = {}


def _ocr_page_text(pdf_path: str, page_index: int, page) -> str:
    """對「沒有文字層」的單一頁面做 OCR，回傳辨識出來的文字。

    掃描件常常整頁被掃描成歪的或轉了 90/180/270 度 (例如 DHL 帳單第 2 頁)，
    直接 OCR 轉錯方向的圖片只會得到亂碼，所以先用 pytesseract 的方向偵測
    (image_to_osd) 抓出建議的旋轉角度，轉正後再正式辨識文字。方向偵測本身
    有時會失敗 (太乾淨的圖、文字太少)，失敗就當作不用轉直接 OCR，不要讓
    整個擷取流程掛掉。
    """
    cache_key = (pdf_path, page_index)
    if cache_key in _OCR_TEXT_CACHE:
        return _OCR_TEXT_CACHE[cache_key]

    text = ""
    if _OCR_AVAILABLE:
        try:
            image = page.to_image(resolution=300).original
            try:
                osd = pytesseract.image_to_osd(image)
                m = re.search(r"Rotate:\s*(\d+)", osd)
                rotate_by = int(m.group(1)) if m else 0
            except Exception:
                rotate_by = 0
            if rotate_by:
                image = image.rotate(-rotate_by, expand=True)
            text = pytesseract.image_to_string(image) or ""
        except Exception:
            text = ""

    _OCR_TEXT_CACHE[cache_key] = text
    return text


# 掃描檔省算力 (使用者規定，所有供應商適用)：掃描頁先用低解析度、不做轉向
# 偵測快速 OCR；認得出是帳單頁的才用 _ocr_page_text() 高解析度重做，其他頁
# (提單、條款、領料單等附件) 直接用低解析度的文字。
_OCR_FAST_DPI = 150
_OCR_FAST_CACHE: Dict[Tuple[str, int], Tuple[str, bool]] = {}


def _ocr_page_fast(pdf_path: str, page_index: int, page) -> Tuple[str, bool]:
    """低解析度 OCR，回傳 (文字, 是否轉向)。轉向的頁 (例如 DHL 第 2 頁) 低解析度
    讀出來是亂碼、判斷不了，呼叫端一律改用高解析度。"""
    cache_key = (pdf_path, page_index)
    if cache_key not in _OCR_FAST_CACHE:
        text, rotated = "", False
        if _OCR_AVAILABLE:
            try:
                image = page.to_image(resolution=_OCR_FAST_DPI).original
                text = pytesseract.image_to_string(image) or ""
                # 轉向偵測每頁約 1 秒，比低解析度 OCR 還貴：只有讀出來像亂碼
                # 的頁才做
                if _looks_garbled(text):
                    try:
                        m = re.search(r"Rotate:\s*(\d+)", pytesseract.image_to_osd(image))
                        rotated = bool(m and int(m.group(1)))
                    except Exception:
                        rotated = False
            except Exception:
                text = ""
        _OCR_FAST_CACHE[cache_key] = (text, rotated)
    return _OCR_FAST_CACHE[cache_key]


_OCR_PSM_CACHE: Dict[Tuple[str, int, int], str] = {}
# 指定 psm 的頁用 400 dpi：GALUH 掃描帳單在 300 dpi 會把 117.580.000 讀成
# 417.580.000、整欄金額讀不到 (使用者 2026-10-25 要求掃描檔也要到等級高)
_OCR_PSM_DPI = 400


def _ocr_page_with_psm(pdf_path: str, page_index: int, psm: int, contrast: bool = False) -> str:
    """用指定的版面分析模式 (tesseract --psm) 高解析度重新 OCR 一頁。例如
    GALUH 的金額欄跟費用名稱分得很開，預設模式會漏掉整欄金額，psm 6 (整頁
    當成一個文字區塊) 才讀得到。供應商用 register_supplier(ocr_psm=6) 指定。"""
    cache_key = (pdf_path, page_index, psm, contrast)
    if cache_key not in _OCR_PSM_CACHE:
        text = ""
        if _OCR_AVAILABLE:
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    image = pdf.pages[page_index].to_image(resolution=_OCR_PSM_DPI).original
                if contrast:
                    # 灰階 + 自動對比 (contrast=True)：見 invoice_text() 的自我檢查
                    from PIL import ImageOps
                    image = ImageOps.autocontrast(image.convert("L"))
                text = pytesseract.image_to_string(image, config=f"--psm {psm}") or ""
            except Exception:
                text = ""
        _OCR_PSM_CACHE[cache_key] = text
    return _OCR_PSM_CACHE[cache_key]


def limited_pages(pdf):
    """抽樣/診斷用：環境變數 INVOICE_MAX_PAGES (例如 50) 有設定時只讀前 N 頁
    (使用者 2026-10-28：超過 50 頁的 PDF 抽樣時只跑約 50 頁)。正式轉檔不設定，讀整份。"""
    limit = os.environ.get("INVOICE_MAX_PAGES", "").strip()
    return pdf.pages[:int(limit)] if limit.isdigit() and int(limit) > 0 else pdf.pages


def _scanned_page_indexes(pdf_path: str) -> List[int]:
    """沒有文字層的頁 (掃描頁)。"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            return [i for i, p in enumerate(limited_pages(pdf)) if not (p.extract_text() or "").strip()]
    except Exception:
        return []


def _looks_garbled(text: str) -> bool:
    """OCR 文字像亂碼：有一定字數，但「3 個字母以上的英文字」不到一半 (轉向頁
    讀出來多是 "S3XV1"、"O6L'E" 這種片段)。"""
    tokens = text.split()
    if len(tokens) < 20:
        return False
    words = sum(1 for t in tokens if re.fullmatch(r"[A-Za-z]{3,}[.,:;]?", t))
    return words / len(tokens) < 0.5


def _page_needs_full_ocr(fast: Tuple[str, bool]) -> bool:
    """要用高解析度重做的頁：轉向的頁 (低解析度判斷不了)，以及低解析度就認得
    出供應商或抬頭是 Faktur Pajak 的頁 (擷取與稅務發票判斷要用)。"""
    text, rotated = fast
    key = detect_supplier(text)
    if key and SUPPLIER_REGISTRY.get(key, {}).get("page_text"):
        # 這家供應商的 page_text 會自己 OCR (例如 CargoWise 掃描頁用 400 dpi)，
        # 偵測用低解析度就夠，不要再多做一次高解析度 (YUSEN 29 頁省約 50 秒)
        return rotated
    return (rotated or key is not None
            or text.upper().lstrip().startswith("FAKTUR PAJAK"))


def extract_text_from_pdf(pdf_path: str, **pdfplumber_kwargs) -> str:
    """讀取 PDF 全部頁面文字並合併成一個字串。

    帳單大多是「可編輯 PDF」，pdfplumber 可以直接抓到文字層。若遇到沒有
    文字層的掃描檔 (例如某一頁整頁都是圖片)，該頁會改用 OCR (見
    _ocr_page_text()) 辨識文字，不用另外設定，系統會自動判斷每一頁要不要
    用 OCR。

    pdfplumber_kwargs 轉給 page.extract_text()，例如 x_tolerance、
    y_tolerance、layout。不同供應商/不同份 PDF 的版面留白不盡相同，同一組
    參數不一定每份都是最佳解，因此開放參數化，讓 find_best_extraction()
    可以多試幾組，挑出效果最好的。

    全文 = 各頁文字 (extract_pages_from_pdf()) 用換行合併，兩者共用同一份
    快取，同一份 PDF 不管先呼叫哪一支都只解析一次。
    """
    return "\n".join(extract_pages_from_pdf(pdf_path, **pdfplumber_kwargs))


_OCR_ENHANCED_TEXT_CACHE: Dict[Tuple[str, int], str] = {}


def _ocr_page_text_enhanced(pdf_path: str, page_index: int, page) -> str:
    """跟 _ocr_page_text() 一樣會先偵測/校正旋轉角度，但額外先把圖片轉成
    灰階 + 自動對比 (PIL ImageOps.autocontrast) 再做 OCR。

    實測這組前處理可以明顯降低數字誤判的機率 (例如 HANDAL 掃描版帳單原本
    OCR 常把 "14" 誤讀成 "44"，加上這個前處理後正確率明顯提升)。只給
    force_ocr=True 的供應商用獨立的函式/快取，不影響 _ocr_page_text()
    原本給『沒有文字層』的一般掃描件供應商 (例如 DHL) 使用的既有行為。
    """
    cache_key = (pdf_path, page_index)
    if cache_key in _OCR_ENHANCED_TEXT_CACHE:
        return _OCR_ENHANCED_TEXT_CACHE[cache_key]

    text = ""
    if _OCR_AVAILABLE:
        try:
            image = page.to_image(resolution=300).original
            try:
                osd = pytesseract.image_to_osd(image)
                m = re.search(r"Rotate:\s*(\d+)", osd)
                rotate_by = int(m.group(1)) if m else 0
            except Exception:
                rotate_by = 0
            if rotate_by:
                image = image.rotate(-rotate_by, expand=True)
            image = ImageOps.autocontrast(image.convert("L"))
            text = pytesseract.image_to_string(image) or ""
        except Exception:
            text = ""

    _OCR_ENHANCED_TEXT_CACHE[cache_key] = text
    return text


def extract_pages_from_pdf_force_ocr(pdf_path: str) -> List[str]:
    """跟 extract_pages_from_pdf() 類似，但『不管 pdfplumber 有沒有抓到文字
    層，每一頁一律強制用 OCR 辨識』，不看文字層內容。

    用途：少數供應商的 PDF 雖然有文字層 (不是空白)，但那個文字層本身就是
    亂碼 (字型編碼有問題，數字/字母會被讀錯，例如 "3,892,040" 被讀成
    "3,g92,o4o"，8 被讀成 g、0 被讀成 o)。一般流程 (extract_pages_from_pdf())
    只在『文字層完全是空白』時才會改用 OCR，遇到這種『有文字層但是亂碼』
    的狀況不會觸發，會悄悄把錯誤的數字寫進 Excel，不會有任何警告。

    第一次發現這個狀況的供應商是 HANDAL (PT. Handal Mandiri Transindo) 的
    掃描版帳單。透過 register_supplier(..., force_ocr=True) 讓特定供應商
    一律呼叫這個函式，不受一般『先試文字層』的邏輯影響；其他供應商不受
    影響。

    ⚠️ 同一家供應商可能同時有「真正掃描的版型」跟「原生可編輯 PDF 的版型」
    兩種帳單 (HANDAL 就是這樣：月結單是掃描件，但部分單張發票是乾淨的
    原生 PDF，完全沒有文字層亂碼問題)。如果整頁都強制 OCR，乾淨版型反而
    會被 OCR 的版面重新切行方式搞亂 (OCR 看到的是『渲染後的圖片』，欄位
    斷行方式跟 pdfplumber 原生文字層不一樣，容易把不同欄位的文字黏在一起)。
    因此這裡逐頁判斷：只有『整頁幾乎被一張大圖片蓋住』(page.images 裡有
    覆蓋面積 >= 90% 頁面的圖片，也就是真正的掃描頁) 才強制 OCR；沒有這種
    滿版圖片的頁面 (原生文字 PDF) 一律直接用 pdfplumber 的文字層，不做
    OCR，避免把乾淨版型的文字搞亂。
    """
    pages_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(limited_pages(pdf)):
            if _page_is_scanned_image(page):
                text = _ocr_page_text_enhanced(pdf_path, i, page)
            else:
                text = page.extract_text() or ""
            pages_text.append(text)
    return pages_text


def _page_is_scanned_image(page, coverage_threshold: float = 0.9) -> bool:
    """判斷這一頁是不是『整頁被一張掃描圖片蓋住』(真正的掃描件)，而不是
    『原生文字 PDF 單純也有插圖』。用頁面裡最大那張圖片的面積占頁面總面積
    的比例來判斷，預設 90% 以上才算掃描頁。"""
    try:
        page_area = float(page.width) * float(page.height)
        if page_area <= 0:
            return False
        for im in page.images:
            img_w = float(im.get("x1", 0)) - float(im.get("x0", 0))
            img_h = float(im.get("y1", 0)) - float(im.get("y0", 0))
            if img_w <= 0 or img_h <= 0:
                continue
            if (img_w * img_h) / page_area >= coverage_threshold:
                return True
    except Exception:
        return False
    return False


# 同一份 PDF 在一次轉檔裡會被讀好幾次 (擷取 parse_invoice_pdf()、金額稽核
# audit_invoice_totals()、稅務發票判斷 standalone_tax_invoice_owner()、
# multi_page 供應商逐頁擷取)。pdfplumber 解析是整個流程最慢的一步 (實測
# MAERSK 每份約 180ms，佔總時間 4 成以上是在重複解析)，所以把每頁文字快取
# 起來，同一份 PDF 只解析一次。只留最近幾份，批次處理上萬份也不會佔記憶體；
# key 含檔案修改時間與大小，同一路徑換成別的檔案 (例如網頁上傳的暫存檔)
# 不會讀到舊內容。
_PAGES_CACHE: "OrderedDict[tuple, List[str]]" = OrderedDict()
_PAGES_CACHE_SIZE = 8


def _pages_cache_key(pdf_path: str, pdfplumber_kwargs: Dict) -> tuple:
    try:
        st = os.stat(pdf_path)
        identity = (st.st_mtime_ns, st.st_size)
    except OSError:
        identity = None
    kwargs = tuple(sorted((k, repr(v)) for k, v in pdfplumber_kwargs.items()))
    return os.path.abspath(pdf_path), identity, kwargs


def extract_pages_from_pdf(pdf_path: str, **pdfplumber_kwargs) -> List[str]:
    """跟 extract_text_from_pdf() 一樣，但個別回傳『每一頁』的文字 (不合併成
    一個字串)。給 multi_page 供應商 (一份 PDF 裡有好幾張各自獨立的發票，
    一頁一張) 逐頁辨識/擷取用。同樣會對沒有文字層的頁面自動改用 OCR。

    Excel 帳單 (見 is_spreadsheet()) 回傳每個「顯示中」工作表的文字，第一
    段以 SHEET_TEXT_MARK 開頭，detect_supplier() 據此只比對 Excel 供應商。
    """
    if is_spreadsheet(pdf_path):
        return [sheet_text(name, rows) for name, rows in read_spreadsheet(pdf_path).sheets]

    key = _pages_cache_key(pdf_path, pdfplumber_kwargs)
    if key in _PAGES_CACHE:
        _PAGES_CACHE.move_to_end(key)
        return list(_PAGES_CACHE[key])

    pages_text = []
    with pdfplumber.open(pdf_path) as pdf:
        scanned = []
        for i, page in enumerate(limited_pages(pdf)):
            text = page.extract_text(**pdfplumber_kwargs) or ""
            if not text.strip():
                scanned.append(i)
            pages_text.append(text)
        text_key = detect_supplier("\n".join(pages_text)) if scanned and len(scanned) < len(pages_text) else None
        if text_key and SUPPLIER_REGISTRY.get(text_key, {}).get("page_text"):
            # 文字層已經認得出供應商，而且這家的 page_text 會自己 OCR 掃描頁：
            # 這裡不用再 OCR (掃描頁留空字串，由 supplier_pages() 補上)
            scanned = []
        if scanned:
            # 沒有文字層的掃描頁：先低解析度，認得出是帳單頁的再高解析度重做；
            # 低解析度一頁都認不出 (例如整份掃歪) 時全部用高解析度，跟以前一樣
            fast = {i: _ocr_page_fast(pdf_path, i, pdf.pages[i]) for i in scanned}
            full = [i for i in scanned if _page_needs_full_ocr(fast[i])]
            if not any(detect_supplier(fast[i][0]) for i in scanned):
                full = scanned
            for i in scanned:
                pages_text[i] = (_ocr_page_text(pdf_path, i, pdf.pages[i]) if i in full
                                 else fast[i][0])

    _PAGES_CACHE[key] = pages_text
    while len(_PAGES_CACHE) > _PAGES_CACHE_SIZE:
        _PAGES_CACHE.popitem(last=False)
    return list(pages_text)


# ---------------------------------------------------------------------------
# 2b. Excel 帳單 (越南供應商常用 Excel 對帳表請款，不是 PDF)
# ---------------------------------------------------------------------------
# 只讀「顯示中」的工作表：供應商常把舊的工作副本 (例如 "(2)"、"CHECK XE"、
# "THEO BG"、"gốc") 隱藏起來留在同一個檔案裡，那些不是這次請款的內容。
# Excel 供應商用 register_supplier(..., sheet_parse=..., sheet_totals=...)
# 註冊：detect 一樣吃文字 (sheet_text() 的結果)，parse/稽核則直接拿結構化
# 的儲存格 (SheetBook)，不用從文字再切欄位。

SPREADSHEET_SUFFIXES = (".xlsx", ".xlsm", ".xls")
# Excel 帳單轉成的文字開頭會帶這個標記，detect_supplier() 看到就只比對
# Excel 供應商 (避免 PDF 供應商的關鍵字，例如備註欄的 "DHL"，誤判)
SHEET_TEXT_MARK = "[[SPREADSHEET]]"


class SheetBook:
    """一份 Excel 帳單：file_name 檔名、sheets [(工作表名稱, 列資料)] (只含
    顯示中的工作表；列資料是各儲存格的值，日期為 datetime)。"""

    def __init__(self, path: str, sheets: List[Tuple[str, List[list]]]):
        self.path = path
        self.file_name = os.path.basename(path)
        self.sheets = sheets


def is_spreadsheet(path) -> bool:
    return str(path).lower().endswith(SPREADSHEET_SUFFIXES)


_BOOK_CACHE: "OrderedDict[tuple, SheetBook]" = OrderedDict()


def read_spreadsheet(path: str) -> SheetBook:
    """讀 Excel 帳單的顯示中工作表 (xlsx/xls 都支援，用 python-calamine)。
    跟 PDF 一樣快取最近幾份，擷取、稽核、驗證時不會重複讀檔。"""
    from python_calamine import CalamineWorkbook, SheetVisibleEnum

    key = _pages_cache_key(path, {})
    if key in _BOOK_CACHE:
        _BOOK_CACHE.move_to_end(key)
        return _BOOK_CACHE[key]
    wb = CalamineWorkbook.from_path(path)
    sheets = []
    for meta in wb.sheets_metadata:
        if meta.visible != SheetVisibleEnum.Visible:
            continue
        rows = wb.get_sheet_by_name(meta.name).to_python(skip_empty_area=False)
        sheets.append((meta.name, rows))
    book = SheetBook(path, sheets)
    _BOOK_CACHE[key] = book
    while len(_BOOK_CACHE) > _PAGES_CACHE_SIZE:
        _BOOK_CACHE.popitem(last=False)
    return book


def sheet_text(name: str, rows: List[list]) -> str:
    """工作表轉成文字 (每列一行，儲存格用 ' | ' 隔開)，給 detect/驗證/顯示用。"""
    lines = [f"{SHEET_TEXT_MARK} [{name}]"]
    for row in rows:
        cells = [str(v).strip() for v in row if v not in (None, "")]
        if cells:
            lines.append(" | ".join(cells))
    return "\n".join(lines)


def excel_formula_issues(path: str) -> Optional[List[Dict]]:
    """檢查 Excel 帳單 (xlsx) 顯示中工作表的合計公式：欄位自己的直向
    =SUM(X3:X259) 範圍，有沒有漏掉下面的明細列 (例如明細到第 277 列，合計
    只加到第 259 列)。回傳問題清單 (同一列、同一個漏算範圍的欄合併成一筆)：
      {"sheet", "row", "end", "last", "formula", "missing": {欄: 漏算金額}}
    .xls 舊格式讀不到公式，回傳 None (呼叫端顯示「無法檢查」)。"""
    if not str(path).lower().endswith((".xlsx", ".xlsm")):
        return None
    from openpyxl import load_workbook as _load
    from openpyxl.utils import column_index_from_string

    sum_re = re.compile(r"^=SUM\(\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)\)$", re.I)
    values = dict(read_spreadsheet(path).sheets)
    wb = _load(path, read_only=True, data_only=False)
    issues: List[Dict] = []
    try:
        for ws in wb.worksheets:
            if ws.sheet_state != "visible" or ws.title not in values:
                continue
            filled: Dict[str, List[int]] = {}   # 欄 -> 有內容的列號
            sums = []                           # (列, 欄, 起, 迄, 公式)
            formulas: Dict[Tuple[int, str], str] = {}   # (列, 欄) -> 公式
            for row in ws.iter_rows():
                for cell in row:
                    v = getattr(cell, "value", None)
                    if v is None or v == "":
                        continue
                    col = cell.column_letter
                    filled.setdefault(col, []).append(cell.row)
                    if isinstance(v, str) and v.startswith("="):
                        formulas[(cell.row, col)] = v.replace(" ", "")
                    m = sum_re.match(v) if isinstance(v, str) else None
                    if m and m.group(1).upper() == m.group(3).upper() == col:
                        sums.append((cell.row, col, int(m.group(2)), int(m.group(4)), v))
            grid = values[ws.title]
            grouped: Dict[tuple, Dict] = {}
            for f_row, col, start, end, formula in sums:
                data_rows = [r for r in filled.get(col, []) if start <= r < f_row]
                last = max(data_rows, default=end)
                if last <= end:
                    continue
                c = column_index_from_string(col) - 1
                missing = 0.0
                for r in range(end + 1, last + 1):
                    v = grid[r - 1][c] if r - 1 < len(grid) and c < len(grid[r - 1]) else None
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        missing += v
                g = grouped.setdefault((f_row, end, last), {
                    "sheet": ws.title, "row": f_row, "end": end, "last": last,
                    "formula": formula, "missing": {}})
                g["missing"][col] = missing
            issues.extend(grouped.values())
            issues.extend(_rate_total_issues(ws.title, formulas, sums, grid))
            # 公式沒有存計算結果 (用程式產生、沒在 Excel 開過的 xlsx)：讀不到金額，
            # 使用者 2026-10-27 規定標 ERROR
            blank = [f"{col}{r}" for (r, col) in formulas
                     if (lambda v: v is None or v == "")(
                         grid[r - 1][column_index_from_string(col) - 1]
                         if r - 1 < len(grid) and column_index_from_string(col) - 1 < len(grid[r - 1])
                         else None)]
            # 少數公式本來就算出空白 (=IF(…,"",…))：一半以上的公式都沒有值才算
            if len(blank) >= 3 and len(blank) * 2 >= len(formulas):
                issues.append({"kind": "uncached", "sheet": ws.title, "cells": blank})
    finally:
        wb.close()
    return issues


# 合計列用「同一列其他欄 × 稅率」算的公式，例如 E360 '=D360*8%'
_RATE_RE = re.compile(r"^=\+?(?:ROUND\()?\$?([A-Z]{1,3})\$?(\d+)\*([\d.]+%?)", re.I)


def _rate_total_issues(sheet: str, formulas: Dict[Tuple[int, str], str],
                       sums: List[tuple], grid: List[list]) -> List[Dict]:
    """合計列的稅額用「淨額合計 × 稅率」算 (=D360*8%)，不是加總明細稅額：
    明細有一列沒填稅額 (例如 E357 空白、D357 有金額) 時，合計稅額就跟
    明細加總對不上。回傳 {"kind": "rate", "sheet", "row", "col", "formula",
    "detail_sum", "total", "blank": [沒填稅額的儲存格]}。"""
    from openpyxl.utils import column_index_from_string

    def num(r, c):
        v = grid[r - 1][c] if r - 1 < len(grid) and c < len(grid[r - 1]) else None
        return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None

    issues = []
    sum_range = {(r, col): (start, end) for r, col, start, end, _f in sums}
    for (f_row, col), formula in formulas.items():
        m = _RATE_RE.match(formula)
        if not m or int(m.group(2)) != f_row or m.group(1).upper() == col:
            continue
        base = m.group(1).upper()
        if (f_row, base) not in sum_range:
            continue                      # 只看合計列 (淨額欄是 =SUM 的那一列)
        start, end = sum_range[(f_row, base)]
        c, b = column_index_from_string(col) - 1, column_index_from_string(base) - 1
        detail = sum(num(r, c) or 0 for r in range(start, end + 1))
        total = num(f_row, c)
        if total is None or abs(total - detail) <= 1:
            continue
        blank = [f"{col}{r}" for r in range(start, end + 1)
                 if num(r, b) and num(r, c) is None and (r, col) not in formulas]
        issues.append({"kind": "rate", "sheet": sheet, "row": f_row, "col": col,
                       "formula": formula, "detail_sum": detail, "total": total,
                       "blank": blank, "start": start, "end": end})
    return issues


def format_formula_issue(issue: Dict) -> str:
    """公式問題 -> 說明文字 (給驗證異常、轉檔總表、查驗結果用)。"""
    if issue.get("kind") == "uncached":
        cells = issue["cells"]
        return (f"原帳單問題 (Excel 公式沒有存計算結果)：[{issue['sheet'].strip()}] {len(cells)} 個公式"
                f"儲存格讀不到金額 (例 {'、'.join(cells[:3])})；請用 Excel 開啟後另存，再重新轉檔")
    if issue.get("kind") == "rate":
        blank = "、".join(issue["blank"][:5]) + (" 等" if len(issue["blank"]) > 5 else "")
        col, start, end = issue["col"], issue.get("start"), issue.get("end")
        fix = f"，應為 =SUM({col}{start}:{col}{end})" if start and end else ""
        return (f"原帳單問題 (Excel 公式有誤)：[{issue['sheet'].strip()}] {col}{issue['row']} "
                f"用 {issue['formula']} 計算，不是加總明細稅額{fix}；合計 {issue['total']:,.2f}，"
                f"明細加總 {issue['detail_sum']:,.2f}，差額 {issue['total'] - issue['detail_sum']:,.2f}"
                + (f"；明細 {blank} 空白 (有淨額，沒有稅額)" if blank else ""))
    cols = list(issue["missing"])
    amounts = "、".join(f"{c} 欄 {v:,.2f}" for c, v in issue["missing"].items() if v)
    return (f"原帳單問題 (Excel 公式有誤)：[{issue['sheet'].strip()}] 第 {issue['row']} 列 "
            f"{'/'.join(cols)} 欄 (例 {issue['formula']}) 只加到第 {issue['end']} 列，"
            f"明細到第 {issue['last']} 列，第 {issue['end'] + 1}~{issue['last']} 列沒算進合計"
            + (f"；漏算 {amounts}" if amounts else ""))


# ---------------------------------------------------------------------------
# 2c. 非英文文字加英文備註 (使用者最高原則：擷取出來不是英文就要加英文翻譯)
# ---------------------------------------------------------------------------
# 只在「寫入轉檔結果 Excel」時套用 (跟日期格式一樣)，不改 parse 結果，比對
# 正確答案與回歸測試不受影響。比對用「去掉聲調、轉大寫」的文字；先套整句
# 對照，再套詞組對照；越南文沒有對照的部分至少附上去聲調拼音。新詞直接加進
# 下面兩個表即可。

# CONSIGNEE (集團公司名稱) 不加英文備註 (使用者 2026-10-27 規定)
TEXT_NOTE_FIELDS = ("supplier", "description", "port_of_loading",
                    "port_of_discharge", "shipper", "vessel")

# 整句對照 (公司名稱等)
_EN_FULL = {
    "CONG TY TNHH GIAO NHAN THUONG MAI RONG A CHAU":
        "Rong A Chau Trading Forwarding Co., Ltd. / Asia Dragon Cargo",
    "CONG TY TNHH DICH VU GIAO NHAN VAN TAI VA THUONG MAI CONG THANH":
        "Cong Thanh Forwarding, Transport and Trading Service Co., Ltd.",
    "CONG TY TRACH NHIEM HUU HAN POU SUNG VIET NAM": "Pou Sung Vietnam Co., Ltd.",
    # CONG THANH 對帳單的額外費用說明 (DEC 欄)
    "10 KIEN HANG BI UOT HU HET GIAY, POUSUNG KHAC PHUC HAU QUA TOI DA CON PHAT SINH THEM "
    "63.5USD (TY GIA: 26.241 NGAY 15/1/2026). CONG THANH TRA PHI PHAT SINH NAY (SU DUNG "
    "INVOICE PSV0810666665 DE CAN TRU)":
        "10 packages got wet and the shoes were damaged; after Pousung's mitigation an extra "
        "63.5 USD remained (rate 26.241 on 15/1/2026). Cong Thanh pays this cost (offset "
        "against invoice PSV0810666665)",
    "XIN KIEM TAI KHO RIENG MAY CAT, MAY CANG LIEU : PS251205NO; VP26111Y11":
        "Request for inspection at a private warehouse for cutting and fabric-stretching "
        "machines: PS251205NO; VP26111Y11",
    "DE XUAT HAI QUAN KIEM MAY XUAT BAN TAI NHA MAY INV VP26110N3S ; PS251205NO":
        "Request for customs inspection at the factory of machines exported for sale, "
        "inv VP26110N3S; PS251205NO",
    "DIEU CHINH TO KHAI GIAM SO LUONG, SO KIEN, TRONG LUONG, TRI GIA...PSV0810686870":
        "Customs declaration amendment to reduce quantity, packages, weight and value... "
        "PSV0810686870",
}
# 詞組對照 (由長到短套用；越南文、印尼文常見帳單用語)
_EN_PHRASES = (
    # 越南文
    ("PHU PHI XANG DAU", "Fuel surcharge"), ("PHI VAN CHUYEN", "Transportation fee"),
    ("PHI DICH VU", "Service fee"), ("PHAT SINH", "Incurred charge"),
    ("PT HA BAI SOM", "Early drop-off fee"), ("NANG HA", "Lift on/off"),
    ("LUU BAI", "Storage"), ("LUU CONT", "Container detention"),
    ("VE SINH CONT", "Container cleaning"),
    # CONG THANH 額外費用項目
    ("XIN KIEM MAY TAI NHA MAY", "Machine inspection at factory request"),
    ("XIN KIEM NHA MAY", "Factory inspection request"),
    ("KHAI BAO MAY MOC", "Machinery declaration"), ("KIEM MAY MOC", "Machinery inspection"),
    ("KHAI BO SUNG", "Supplementary declaration"), ("DAN TEM", "Label sticking"),
    ("XIN GIU GAC VA DAN TEM", "Guard keeping and label sticking"),
    ("GIU GAC", "Guard keeping"),
    # 越南 DSV 費用名稱 (使用者 2026-10 正確答案)
    ("PHI LAO VU - HANG XUAT", "Export Service Fee"),
    ("PHI QUAN LY HANG XUAT", "Export Cargo Management Fee"),
    ("PHI KHAI BAO AMS", "AMS Filing Fee"), ("PHI VAN DON", "Bill of Lading Fee"),
    # 越南 MOL 費用名稱
    ("PHI NIEM CHI", "Seal Fee"), ("PHI KHAI AFR", "AFR Filing Fee"),
    ("PHI XEP DO", "Terminal Handling / Loading-Unloading Fee"), ("PHI LAM HANG", "Handling Fee"),
    ("PHI SOI HANG", "X-ray Screening Fee"), ("PHI PHAT HANH VAN DON", "B/L Issuance Fee"),
    ("PHI VGM", "VGM Fee"), ("PHI CFS", "CFS Fee"), ("PHI THC", "Terminal Handling Charge"),
    ("THUE 8% VAT", "VAT 8%"), ("THUE 10% VAT", "VAT 10%"),
    # 印尼文
    ("TOTAL BIAYA EKSTERNAL", "Total external cost"), ("TOTAL BIAYA REIMBURSEMENT",
                                                       "Total reimbursement cost"),
    ("BIAYA LOKASI BELUM OPEN", "Location-not-yet-open fee"),
    ("BIAYA SEGEL CONTAINER", "Container seal fee"), ("BIAYA SEWA MOBIL", "Car rental fee"),
    ("BIAYA PENUMPUKAN", "Storage (stacking) fee"), ("BIAYA LOLO", "LOLO fee"),
    ("BIAYA EXTERNAL", "External cost"), ("BIAYA EKSTERNAL", "External cost"),
    ("JASA TRANSPORTASI", "Transportation service"), ("JASA PPJK", "Customs broker (PPJK) service"),
    ("JASA HANDLING", "Handling service"), ("GUDANG FEE", "Warehouse fee"),
    ("FIAT DOKUMEN", "Document approval"), ("ADMINISTRASI", "Administration"),
    ("MATERAI", "Stamp duty"), ("NON PPH", "excl. income tax (PPh)"),
    ("BIAYA", "Fee"), ("JASA", "Service"), ("GUDANG", "Warehouse"),
)
_EN_PHRASE_RES = tuple((re.compile(rf"(?<![A-Z]){re.escape(k)}(?![A-Z])"), v)
                       for k, v in sorted(_EN_PHRASES, key=lambda kv: -len(kv[0])))


def _strip_tones(text: str) -> str:
    import unicodedata
    text = text.replace("Đ", "D").replace("đ", "d")
    return "".join(ch for ch in unicodedata.normalize("NFKD", text)
                   if not unicodedata.combining(ch))


def english_note(value):
    """不是英文的文字後面加上英文翻譯括號備註，例如
    'PHỤ PHÍ XĂNG DẦU UP 15%' -> 'PHỤ PHÍ XĂNG DẦU UP 15% (Fuel surcharge UP 15%)'。
    已經是英文 (或對照表沒有、也沒有越南文聲調) 就原樣回傳；已經有同樣備註
    的不重複加。"""
    if not isinstance(value, str) or not value.strip() or value == NA:
        return value
    plain = _strip_tones(value)
    key = re.sub(r"\s+", " ", plain).strip().upper()
    english = _EN_FULL.get(key)
    if english is None:
        translated, used = key, []
        for pattern, repl in _EN_PHRASE_RES:
            if pattern.search(translated):
                translated = pattern.sub(repl, translated)
                used.append(repl)
        # 已經帶英文備註 (例如 'Nang ha (Lift on/off)') 就不重複加
        if used and all(u.upper() in value.upper() for u in used):
            return value
        if used:
            english = translated
        elif plain != value:
            # 越南文地名等沒有對照的：附上去聲調的拼音
            english = re.sub(r"\s+", " ", plain).strip().title()
    if not english or english.upper() in value.upper():
        return value
    return f"{value} ({english})"


def _generate_extraction_configs(max_attempts: int = 100) -> List[Dict]:
    """產生一組(最多 max_attempts 個)不同的 pdfplumber 擷取參數組合，
    給「重新嘗試擷取」的迴圈依序測試。第一組永遠是預設值 (最常見、最快)，
    之後依序嘗試不同的 x_tolerance / y_tolerance 組合 (影響同一列文字判斷
    的容忍度，容忍度不同時，左右並排的欄位有機會被切開或黏在一起)，
    最後補上 layout=True (盡量保留原始版面座標間距) 這個較特殊的模式。
    """
    configs: List[Dict] = [{}]  # 第1次一定先試預設值
    x_tolerances = [1, 1.5, 2, 2.5, 3, 3.5, 4, 4.5, 5]
    y_tolerances = [1, 1.5, 2, 2.5, 3, 3.5, 4, 4.5, 5, 5.5, 6, 7, 8, 9, 10, 12, 15]
    for y_tol in y_tolerances:
        for x_tol in x_tolerances:
            if len(configs) >= max_attempts - 1:  # 留最後一格給 layout=True
                break
            cfg = {"x_tolerance": x_tol, "y_tolerance": y_tol}
            if cfg not in configs:
                configs.append(cfg)
        if len(configs) >= max_attempts - 1:
            break
    configs.append({"layout": True})
    return configs[:max_attempts]


# 預設的「最多嘗試次數」：正確率沒有達到 100% 時，最多重新嘗試這麼多組
# 不同的擷取參數，每組都算一次正確率，一達到 100% 就提早停止。
MAX_EXTRACTION_ATTEMPTS = 5
EXTRACTION_CONFIGS: List[Dict] = _generate_extraction_configs(MAX_EXTRACTION_ATTEMPTS)


# ---------------------------------------------------------------------------
# 2.5 「學習記憶」：記住每家供應商目前已知能得到最高分的擷取設定
# ---------------------------------------------------------------------------
#
# 迴圈每次找到「比目前記錄更高分」的設定時，就寫回這個 JSON 檔，下次同一家
# 供應商的帳單進來時，會優先套用這組已知最佳設定當作第 1 次嘗試 (而不是
# 每次都要從頭試 100 組)，等於系統會隨著使用次數愈多、愈快找到高分設定。
#
# 注意（重要限制）：這只是「同一次部署期間」的記憶。Streamlit Cloud 每次
# reboot / 重新部署都會用 GitHub 上的原始碼重新建立環境，這個檔案若沒有
# 一併提交回 GitHub，記憶就會被重置。若要讓記憶永久保留，需要把
# learned_configs.json 這個檔案也上傳/提交到 repo 裡。
LEARNED_CONFIGS_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "learned_configs.json"
)


def _load_learned_configs() -> Dict[str, Dict]:
    try:
        with open(LEARNED_CONFIGS_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_learned_config(supplier_key: Optional[str], config: Optional[Dict], score: float) -> None:
    """只有『這次分數比之前記錄的更高』時才更新，避免學到比較差的設定。"""
    if not supplier_key:
        return
    data = _load_learned_configs()
    prev = data.get(supplier_key)
    if prev is not None and score <= prev.get("score", -1):
        return
    data[supplier_key] = {"config": config or {}, "score": score}
    try:
        with open(LEARNED_CONFIGS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass  # 寫入失敗 (例如唯讀環境) 不影響主流程，只是這次沒學到


# ---------------------------------------------------------------------------
# 3. 供應商辨識與註冊表
# ---------------------------------------------------------------------------
#
# 每個供應商模組 (suppliers.py) 呼叫 register_supplier() 註冊：
#   - key        內部代碼，例如 "DWIHARTA"
#   - label      顯示名稱
#   - detect_fn  detect(text) -> bool，判斷這份PDF文字是不是這家供應商
#   - parse_fn   parse(text) -> {"header": {...}, "items": [{...}, ...]}
#                header 裡的 key 對應 HEADER_FIELD_CODES，
#                items 是一個 list，每筆是 {"description":.., "amount":..}，
#                至少要回傳一筆 (抓不到費用明細就回傳一筆全空的)。
#   - multi_page 選填，預設 False。極少數供應商會把「好幾張不同發票號碼的
#                獨立帳單」放在同一份 PDF 裡、一頁一張 (例如 HYPER MEGA)，
#                跟其他供應商「不管幾頁都是同一張發票」完全不同。這種情況
#                設 multi_page=True，detect_fn/parse_fn 改成逐頁呼叫 (每次
#                傳入『這一頁』的文字，而不是整份 PDF 合併後的文字)，
#                找出每一頁各自的發票，分別展開成列。

SUPPLIER_REGISTRY: Dict[str, Dict] = {}


def register_supplier(key: str, label: str,
                       detect_fn: Callable[[str], bool],
                       parse_fn: Optional[Callable[[str], Dict]],
                       multi_page: bool = False,
                       extract_total: Optional[Callable[[str], Optional[float]]] = None,
                       extract_grand_total: Optional[Callable[[str], Optional[float]]] = None,
                       extract_final_total: Optional[Callable[[str], Optional[float]]] = None,
                       deduction_prefixes: Tuple[str, ...] = (),
                       audit_exclude: Tuple[str, ...] = (),
                       force_ocr: bool = False,
                       sheet_parse: Optional[Callable[["SheetBook"], List[Dict]]] = None,
                       sheet_totals: Optional[Callable[["SheetBook"], List[Dict]]] = None,
                       extract_vat_total: Optional[Callable[[str], Optional[float]]] = None,
                       invoice_page: Optional[Callable[[str], bool]] = None,
                       ocr_psm: Optional[int] = None,
                       page_text: Optional[Callable[[str], List[str]]] = None) -> None:
    """註冊一家供應商。

    page_text (可選)：這家供應商自己的逐頁文字擷取 (傳入 PDF 路徑，回傳每頁
    文字)，取代預設的 extract_pages_from_pdf()。用在需要看字元大小/位置才分得
    開的版面 (例如越南 CargoWise 帳單的 DRAFT 浮水印、疊在費用名稱上的代碼)。
    回傳空字串的頁 (掃描頁) 會退回預設擷取 (含 OCR)。

    ocr_psm (可選，非 multi_page 供應商用)：這家的掃描帳單頁改用指定的
    tesseract 版面分析模式重新 OCR (例如 GALUH 用 6，預設模式會漏掉金額欄)。

    invoice_page (可選，非 multi_page 供應商用)：判斷某一頁是不是這家的帳單
    頁。有給時只擷取/勾稽帳單頁，其他頁 (提單、條款、其他公司的發票等雜訊)
    當附件，見 invoice_text()；查驗結果會列出附件頁。

    extract_vat_total (可選，目前只用於單張發票的 PDF)：從原始文字抓帳單本身
    寫的 VAT 合計 (例如 DWIHARTA 的 "V A T 201,348")，跟擷取到的 VAT 加總
    勾稽 (vat_match)。

    sheet_parse (Excel 帳單供應商用)：有給就代表這家供應商的帳單是 Excel
    (見 read_spreadsheet())。detect_fn 一樣吃文字 (sheet_text() 的結果)，
    但擷取改呼叫 sheet_parse(book)，回傳「每張對帳單/發票各一個」
    {"header":.., "items":[..]} 的 list；parse_fn 可傳 None。

    sheet_totals (Excel 帳單供應商用，可選)：從帳單本身的合計列抓總額，回傳
    [{"invoice_no":.., "total":.., "grand_total":.., "final_total":..}, ..]
    (不適用的項目給 None)，audit_invoice_totals() 拿來跟擷取到的明細加總勾稽；
    可另帶 "grand_tolerance" / "final_tolerance" 放寬含稅總額 / 實付金額的容許
    誤差 (預設 ±1)。invoice_no 給
    None 代表「整份檔案的合計」，用全部明細勾稽並套用到檔案裡的每張發票；
    同時也有該發票自己的合計時，含稅總額 (grand) 以發票自己的為準。

    extract_total (可選)：從『一張發票』的原始文字裡抓出帳單本身寫的
    『未稅總額』(Total (Excl. VAT) / Sub Total / SUBTOTAL / Net value… /
    DPP，依供應商命名不同)，用來跟系統擷取到的費用明細加總互相稽核 (見
    audit_invoice_totals())。沒有實作 (預設 None) 的供應商就不做這項稽核，
    不會被當成「不一致」。

    extract_grand_total (可選)：從『一張發票』的原始文字裡抓出帳單本身寫的
    『含稅總額』(Total Amount Due / Amount Due / Total / GRAND TOTAL…，依
    供應商命名不同)。有實作的話，audit_invoice_totals() 會多做第二項稽核：
    系統擷取到的費用明細加總 (未稅) + 每筆費用的 VAT 加總 (排除
    deduction_prefixes 指定的扣除項，見下)，應該要等於這個含稅總額；不
    一致就標示出來。這一項跟 extract_total 互相獨立，各自沒有實作的供應商
    就不做那一項稽核 (不算「不一致」，只是『還沒支援』)。

    extract_final_total (可選)：從『一張發票』的原始文字裡抓出帳單本身寫的
    『實際支付金額』(例如印尼供應商常見的 "TOTAL YANG DIBAYAR"，代表含稅
    總額再扣掉代扣稅之後真正要付的錢)。有實作的話，audit_invoice_totals()
    會多做第三項稽核：含稅總額 (未稅加總 + 非扣除類 VAT 加總) 再減掉
    deduction_prefixes 指定的扣除類 VAT 加總，應該要等於這個實付金額；
    不一致就標示出來。這一項也跟前兩項互相獨立。

    deduction_prefixes (可選)：這家供應商『每筆費用自帶 VAT 欄』或『獨立
    一列稅金』裡，哪些 description 開頭字樣代表的不是要『加』進含稅總額
    的稅金 (例如一般的 VAT/PPN)，而是要從含稅總額『扣』掉的代扣稅/預扣稅
    (例如印尼的 PPH 2% 這種由付款方預扣、不是加在帳單金額上的稅)。比對時
    用 description 正規化後 (大寫) 是否以這裡列出的字樣開頭來判斷，不要求
    完全相等 (因為稅率常常隨發票不同，例如 "PPH 2%"/"PPH 1%")。預設空
    tuple 代表這家供應商沒有這種『扣除類』VAT，所有逐筆 VAT 都算加總
    (維持舊行為不變)。

    audit_exclude (可選)：這家供應商的 items 裡，有哪些 description 是
    『不是真正費用、不該算進未稅總額加總』的合成列 (例如 YJE_TATA 的
    "DPP" 是額外附加的未稅金額備註列，不是一筆實際費用)，稽核加總時要
    跳過這些列，避免誤判成不一致。

    force_ocr (可選，預設 False)：這家供應商的 PDF 文字層『有內容但是
    亂碼』(不是空白，一般的『沒有文字層才用 OCR』判斷不會觸發)，設成
    True 後，不管文字層有沒有內容，一律強制用 OCR 重新辨識每一頁 (見
    extract_pages_from_pdf_force_ocr())，取代文字層，避免把亂碼數字
    悄悄寫進 Excel。目前只對 multi_page=True 的供應商生效 (透過
    _parse_multi_page_pdf())。
    """
    SUPPLIER_REGISTRY[key] = {
        "label": label, "detect": detect_fn, "parse": parse_fn, "multi_page": multi_page,
        "extract_total": extract_total, "extract_grand_total": extract_grand_total,
        "extract_final_total": extract_final_total, "deduction_prefixes": deduction_prefixes,
        "audit_exclude": audit_exclude, "force_ocr": force_ocr,
        "sheet_parse": sheet_parse, "sheet_totals": sheet_totals,
        "extract_vat_total": extract_vat_total, "invoice_page": invoice_page,
        "ocr_psm": ocr_psm, "page_text": page_text, "multi_page_pdf": multi_page,
        # Excel 供應商一份檔案可能有好幾張對帳單，比對時跟 multi_page 一樣用
        # invoice_no 分組
        "multi_page": multi_page or sheet_parse is not None,
    }


def detect_supplier(text: str) -> Optional[str]:
    """依序問過每個已註冊的供應商『這是你的帳單嗎』，回傳第一個承認的 key。
    都沒有承認就回傳 None，呼叫端應提示『無法辨識供應商』。
    Excel 帳單的文字 (以 SHEET_TEXT_MARK 開頭) 只比對 Excel 供應商，PDF 的
    文字只比對 PDF 供應商。
    """
    from_sheet = (text or "").startswith(SHEET_TEXT_MARK)
    for key, entry in SUPPLIER_REGISTRY.items():
        # 同一家可能 PDF 與 Excel 帳單都有 (例如越南 MAERSK)：有對應的擷取規則才比對
        if from_sheet and entry.get("sheet_parse") is None:
            continue
        if not from_sheet and entry.get("parse") is None:
            continue
        try:
            if entry["detect"](text):
                return key
        except Exception:
            continue
    return None


def supplier_label(key: Optional[str]) -> str:
    if key and key in SUPPLIER_REGISTRY:
        return SUPPLIER_REGISTRY[key]["label"]
    return "未知供應商 (無法辨識)"


def set_supplier_number(key: str, number: str, vendor_code: str = "",
                        official_name: str = "") -> None:
    """登記供應商編號 (例如 'ID02') 與廠商代號 (例如 '100019226')，編號依
    進度管控的 Top 供應商清單 (見 suppliers_streamlit.py 的 SUPPLIER_NUMBERS)。
    只是對照資料，不影響辨識順序與擷取規則。"""
    SUPPLIER_REGISTRY[key].update(
        number=number, vendor_code=vendor_code, official_name=official_name)


def supplier_number(key: Optional[str]) -> str:
    """供應商編號，沒有登記時回傳空字串。"""
    return (SUPPLIER_REGISTRY.get(key) or {}).get("number", "") if key else ""


def supplier_display(key: Optional[str]) -> str:
    """畫面/報表顯示用：'ID02 MAERSK'，沒有編號時就是 key 本身。"""
    number = supplier_number(key)
    return f"{number} {key}" if number else (key or "")


def list_registered_suppliers() -> List[Dict[str, str]]:
    """回傳目前已註冊的供應商清單 (給網頁上顯示用)，依編號排序，沒有編號
    的排在最後 (維持註冊順序)。辨識時仍依註冊順序比對 (detect_supplier())。"""
    suppliers = [
        {"key": k, "label": v["label"], "number": v.get("number", ""),
         "vendor_code": v.get("vendor_code", "")}
        for k, v in SUPPLIER_REGISTRY.items()
    ]
    return sorted(suppliers, key=lambda s: (not s["number"], s["number"]))


def is_multi_page_supplier(key: Optional[str]) -> bool:
    """這家供應商是不是『一份 PDF 裡有好幾張各自獨立的發票，一頁一張』
    (參見 register_supplier() 的 multi_page 說明)。找不到這家供應商時視為
    False (走一般的單張發票流程)。
    """
    return bool(key and key in SUPPLIER_REGISTRY and SUPPLIER_REGISTRY[key].get("multi_page"))


# ---------------------------------------------------------------------------
# 3b. 單獨一份的印尼稅務發票 (Faktur Pajak / e-Faktur)
# ---------------------------------------------------------------------------
# 印尼供應商常把電子稅務發票 (檔名常見 'OutputTaxInvoice-...' 或賣方稅籍
# 編號開頭) 跟帳單分開寄。這種 PDF 不是帳單，版面也跟任何供應商的帳單
# 不同，硬套供應商規則只會抓出錯誤資料，所以不轉 Excel，另外歸到
# 「<供應商> 稅務發票」資料夾。
#
# 「單獨一份」= 第 1 頁開頭就是 Faktur Pajak，而且『每一頁』都是稅務發票
# 的內容 (續頁是一筆筆費用 + Potongan Harga/PPnBM，最後一頁是 DPP/PPN 合計
# 與電子簽章)。KUEHNE NAGEL 那種「第 1 頁 Faktur Pajak + 後面接 KN Sales
# Invoice」的混合檔，後面的帳單頁沒有這些字樣，不算單獨稅務發票，照常轉檔。

TAX_INVOICE_FOLDER_SUFFIX = "稅務發票"
_FAKTUR_PAGE_MARKERS = (
    "FAKTUR PAJAK", "PENGUSAHA KENA PAJAK", "POTONGAN HARGA", "PPNBM",
    "DASAR PENGENAAN PAJAK", "HARGA JUAL", "JUMLAH PPN",
    "DITANDATANGANI SECARA ELEKTRONIK",
)
# 賣方 = 「Pengusaha Kena Pajak:」下一行的「Nama : ...」
_FAKTUR_SELLER_RE = re.compile(
    r"PENGUSAHA\s+KENA\s+PAJAK\s*:?\s*NAMA\s*:\s*([^\n]+)", re.IGNORECASE)
_FOLDER_UNSAFE_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def _is_faktur_first_page(text: str) -> bool:
    """第一頁的抬頭就是 Faktur Pajak。只看抬頭：掃描檔 OCR 常把其他字樣讀錯
    (例如 'Kode dan Nomor Seri' 讀成 'Kode dan Nomet Se')，不能靠它們判斷。"""
    return (text or "").upper().lstrip().startswith("FAKTUR PAJAK")


def _faktur_seller_by_line(text: str) -> Optional[str]:
    """OCR 讀不出「Pengusaha Kena Pajak」標籤時的備援：抬頭附近 (前 10 行)
    第一個能被已註冊供應商認出的行就是賣方 (例如 'Narna: MAERSK LOGISTICS
    INDONESIA')。逐行判斷，避開 detect 對整份稅務發票的排除規則。"""
    for line in (text or "").splitlines()[1:11]:
        key = detect_supplier(line)
        if key:
            return key
    return None


def standalone_tax_invoice_owner(pdf_path: str, text: str) -> Optional[str]:
    """判斷這份 PDF 是不是「單獨一份的稅務發票」。

    不是：回傳 None (照常轉檔)。
    是：回傳要歸到哪家供應商 —— 賣方名稱能被已註冊的供應商認出來就回傳
    供應商 key (例如 'MAERSK')，認不出來就回傳賣方名稱本身，連賣方名稱都
    抓不到時回傳空字串 (呼叫端自行補「未知供應商」)。

    text 是 extract_text_from_pdf() 合併後的全文；開頭不是 Faktur Pajak 時
    直接回傳 None，不會重讀 PDF，一般帳單不多花時間。
    """
    if not _is_faktur_first_page(text):
        return None
    # 頁面已經在快取裡 (擷取時讀過/OCR 過)，不會重讀。每一頁都要是稅務發票的
    # 內容，或至少不是供應商帳單 (掃描檔背面的條款頁 OCR 常是亂碼，沒有任何
    # 稅務發票關鍵字)；有一頁認得出是供應商帳單就不是「單獨一份」
    pages = extract_pages_from_pdf(pdf_path)
    for page in pages:
        has_marker = any(m in page.upper() for m in _FAKTUR_PAGE_MARKERS)
        if not has_marker and detect_supplier(page):
            return None

    m = _FAKTUR_SELLER_RE.search(text)
    seller = m.group(1).strip() if m else ""
    if seller:
        key = detect_supplier(seller)
        if key:
            return key
    else:
        key = _faktur_seller_by_line(text)
        if key:
            return key
    return _FOLDER_UNSAFE_RE.sub(" ", seller).strip(" .")


# ---------------------------------------------------------------------------
# 4. 擷取結果組裝：header + items -> 完整列 (一筆費用一列)，缺漏補 N/A
# ---------------------------------------------------------------------------

def _fill_na(value):
    if value is None or (isinstance(value, str) and not value.strip()):
        return NA
    return value


def expand_to_rows(parsed: Dict) -> List[Dict[str, object]]:
    """把供應商 parse() 回傳的 {"header":.., "items":[..]} 展開成
    每筆費用一列的完整 row (欄位對應 FIELD_CODES，缺漏一律補 N/A)。

    大部分供應商的抬頭欄位 (例如 Order No.) 整張帳單只有一個值，每筆費用
    明細都套用同一個值即可。但少數供應商 (例如 PT. TRANS DAYA PRIMA 有多
    筆費用明細時) 是「每一筆費用明細各自有自己的 Order No.」，不是整張
    帳單共用一個。為了不用改動每個供應商的資料結構，這裡允許 items 裡的
    每一筆 item dict『順便』帶上任何一個 HEADER_FIELD_CODES 欄位 (例如
    {"description":.., "amount":.., "order_no": "這筆專屬的單號"})，
    有帶的話就覆蓋掉整張帳單共用的抬頭值，只套用在這一列；沒有帶的欄位
    仍然沿用抬頭值，其他供應商 (item 裡只有 description/amount) 完全不受
    影響。
    """
    header = parsed.get("header", {}) or {}
    items = parsed.get("items") or [{}]

    header_row = {code: _fill_na(header.get(code)) for code in HEADER_FIELD_CODES}

    rows = []
    for item in items:
        row = dict(header_row)
        for code in HEADER_FIELD_CODES:
            if code in item:
                row[code] = _fill_na(item.get(code))
        for code in ITEM_FIELD_CODES:
            row[code] = _fill_na(item.get(code))
        rows.append(row)
    return rows


def _parse_multi_page_pdf(pdf_path: str, key: str, **pdfplumber_kwargs) -> List[Dict]:
    """multi_page 供應商專用：逐頁辨識+擷取，每一頁『承認』的就各自展開成
    一組列，回傳所有頁面的列攤平在一起 (一個 list，裡面可能混著好幾張
    不同發票的列，用 invoice_no 欄位分辨屬於哪一張)。一頁都沒認出來時，
    回傳一筆全 N/A 的列，跟其他供應商『抓不到就整張 N/A』的行為一致。
    """
    entry = SUPPLIER_REGISTRY[key]
    if entry.get("force_ocr"):
        pages = extract_pages_from_pdf_force_ocr(pdf_path)
    else:
        pages = supplier_pages(pdf_path, key, **pdfplumber_kwargs)
    rows: List[Dict] = []
    for page_text in pages:
        try:
            if not entry["detect"](page_text):
                continue
        except Exception:
            continue
        parsed = entry["parse"](page_text)
        rows.extend(expand_to_rows(parsed))
    if not rows:
        rows = expand_to_rows({"header": {}, "items": [{}]})
    return rows


_CUSTOM_PAGES_CACHE: "OrderedDict[tuple, List[str]]" = OrderedDict()


def supplier_pages(pdf_path: str, key: Optional[str], **pdfplumber_kwargs) -> List[str]:
    """逐頁文字：供應商有登記 page_text 時用它 (空白頁退回預設擷取)，否則用
    extract_pages_from_pdf()。page_text 的結果有快取：同一份 PDF 擷取、稽核、
    列附件頁會各呼叫一次，不快取每次都要重新解析版面 (YUSEN 29 頁一次約 30 秒)。"""
    entry = SUPPLIER_REGISTRY.get(key) if key else None
    if not (entry and entry.get("page_text")) or is_spreadsheet(pdf_path):
        return extract_pages_from_pdf(pdf_path, **pdfplumber_kwargs)
    cache_key = (key,) + _pages_cache_key(pdf_path, pdfplumber_kwargs)
    if cache_key in _CUSTOM_PAGES_CACHE:
        _CUSTOM_PAGES_CACHE.move_to_end(cache_key)
        return list(_CUSTOM_PAGES_CACHE[cache_key])
    try:
        pages = entry["page_text"](pdf_path)
    except Exception:
        # 自訂擷取失敗就用預設擷取，不中斷流程
        return extract_pages_from_pdf(pdf_path, **pdfplumber_kwargs)
    if any(not p.strip() for p in pages):
        default = extract_pages_from_pdf(pdf_path, **pdfplumber_kwargs)
        if len(default) == len(pages):
            pages = [c if c.strip() else d for c, d in zip(pages, default)]
    _CUSTOM_PAGES_CACHE[cache_key] = pages
    while len(_CUSTOM_PAGES_CACHE) > _PAGES_CACHE_SIZE:
        _CUSTOM_PAGES_CACHE.popitem(last=False)
    return list(pages)


def _page_balances(entry: Dict, text: str) -> Optional[bool]:
    """一頁文字自己算得平嗎：明細加總 (含稅額) = 含稅總額，或明細金額 = 未稅總額。
    沒有總額規則或抓不到總額時回傳 None (不判斷)。"""
    try:
        items = entry["parse"](text).get("items") or []
        amount = sum(i.get("amount") or 0 for i in items)
        vat = sum(i.get("vat") or 0 for i in items)
        if entry.get("extract_grand_total"):
            total = entry["extract_grand_total"](text)
            if isinstance(total, (int, float)):
                return abs(amount + vat - total) <= 1
        if entry.get("extract_total"):
            total = entry["extract_total"](text)
            if isinstance(total, (int, float)):
                return abs(amount - total) <= 1
    except Exception:
        return None
    return None


def invoice_text(pdf_path: str, key: Optional[str],
                 **pdfplumber_kwargs) -> Tuple[str, List[int]]:
    """帳單常夾帶雜訊頁 (提單、條款、其他公司的發票…)。供應商有登記
    invoice_page (register_supplier()) 時，只把「帳單頁」合併成擷取用的文字，
    回傳 (文字, 附件頁的頁碼 (1 起算))；沒有登記或一頁都不符合時用全部頁面。"""
    pages = supplier_pages(pdf_path, key, **pdfplumber_kwargs)
    entry = SUPPLIER_REGISTRY.get(key) if key else None
    psm = entry.get("ocr_psm") if entry else None
    if psm and not is_spreadsheet(pdf_path):
        # 這家供應商的掃描頁要用指定的 OCR 模式 (只重做認得出供應商的頁)
        pages = list(pages)
        for i in _scanned_page_indexes(pdf_path):
            if detect_supplier(pages[i]) == key:
                pages[i] = _ocr_page_with_psm(pdf_path, i, psm)
                # 自我檢查：這一頁的明細加總對不上帳單總額時，改用灰階 + 自動對比
                # 再 OCR 一次，對得上才採用 (GALUH 抽樣 0500047、0500071、0500086：
                # 原圖漏掉 TRUCKING 金額、把 9 讀成 1；有些帳單反而原圖才讀得對)
                if _page_balances(entry, pages[i]) is False:
                    alt = _ocr_page_with_psm(pdf_path, i, psm, contrast=True)
                    if _page_balances(entry, alt):
                        pages[i] = alt
    check = entry.get("invoice_page") if entry else None
    keep = [i for i, p in enumerate(pages) if check(p)] if check else []
    if not keep:
        return "\n".join(pages), []
    noise = [i + 1 for i in range(len(pages)) if i not in keep]
    return "\n".join(pages[i] for i in keep), noise


# 前後不能接英文字母 (不用 \b：抽樣檔名 "f007_POU HUNG…" 的底線算字元)
_POU_FACTORY_RE = re.compile(r"(?<![A-Z])POU\s*-?\s*(YUEN|HUNG|CHEN|SUNG|LI|SHENG|YI)(?![A-Z])")


def _consignee_from_filename(path: str) -> Optional[str]:
    """檔名裡的寶成廠別 ("POUYUEN"、"POU HUNG"、"POULI") -> "POU YUEN" 這種寫法。
    檔名沒有時，看上兩層資料夾名稱 (越南資料夾 "…\\POU LI\\2025\\THANG 09\\…")。"""
    parts = os.path.normpath(path or "").split(os.sep)
    for part in reversed(parts[-3:]):
        m = _POU_FACTORY_RE.search(part.upper())
        if m:
            return f"POU {m.group(1)}"
    return None


def parse_invoice_pdf(pdf_path: str, supplier_key: Optional[str] = None,
                       **pdfplumber_kwargs) -> Tuple[Optional[str], List[Dict], str]:
    """讀取PDF、自動判斷供應商 (若未指定)、用對應規則擷取欄位。
    回傳 (辨識到的供應商key或None, 展開後的rows列表, 原始文字)。

    對 multi_page 供應商 (參見 register_supplier())，rows 可能包含好幾張
    不同發票各自的列 (一頁一張發票)，呼叫端要用 invoice_no 分組。
    """
    text = extract_text_from_pdf(pdf_path, **pdfplumber_kwargs)
    key = supplier_key or detect_supplier(text)

    if is_spreadsheet(pdf_path):
        entry = SUPPLIER_REGISTRY.get(key) if key else None
        if entry and entry.get("sheet_parse"):
            rows = [row for parsed in entry["sheet_parse"](read_spreadsheet(pdf_path))
                    for row in expand_to_rows(parsed)]
        else:
            rows = []
        # 帳單內沒有客戶名稱時，用檔名裡的寶成廠別補 (使用者 2026-10-27：
        # "POU HUNG 09-26 … .xlsx" -> POU HUNG)
        customer = _consignee_from_filename(pdf_path)
        if customer:
            for row in rows:
                if row.get("consignee") in (None, "", NA):
                    row["consignee"] = customer
        return key, rows or expand_to_rows({"header": {}, "items": [{}]}), text

    # PDF 只看註冊時的 multi_page (同時有 Excel 規則的供應商，multi_page 只是比對分組用)
    if key in SUPPLIER_REGISTRY and SUPPLIER_REGISTRY[key].get("multi_page_pdf"):
        rows = _parse_multi_page_pdf(pdf_path, key, **pdfplumber_kwargs)
        return key, rows, text

    if key and key in SUPPLIER_REGISTRY:
        parsed = SUPPLIER_REGISTRY[key]["parse"](invoice_text(pdf_path, key, **pdfplumber_kwargs)[0])
    else:
        parsed = {"header": {}, "items": [{}]}
    rows = expand_to_rows(parsed)
    return key, rows, text


# ---------------------------------------------------------------------------
# 5. 比對評分：擷取結果 vs 正確答案
# ---------------------------------------------------------------------------

# 全系統目前只有這三種日期字串格式會被 parse_xxx() 吐出來（見
# suppliers_streamlit.py 裡各家 _xxx_format_date()/_xxx_parse_date() 最終
# 的回傳格式）：
#   1. "DD Mon YYYY"（例如 "08 Apr 2025"）：絕大多數供應商共用的
#      _format_date_en()/_format_date_en_from_month_name()。
#   2. "YYYY-MON-DD"（例如 "2026-JUN-11"）：INDOPROSTIME 系列專用的
#      _indo_format_date()。
#   3. "DD/MM/YYYY"（例如 "08/04/2025"）：PUTRA_SEJAHTERA_SENTOSA 專用的
#      _pss_format_date()。
#   4. "DD.Mon.YYYY"（例如 "31.Aug.2026"）：MAERSK INVOICE DATE 專用的
#      _maersk_extract_invoice_date() (為了跟正確答案的文字儲存格比對，刻意
#      保留點分隔)；跟格式 1 只差在分隔符號，共用同一個 regex。
# _coerce_to_date() 把這幾種格式統一解析成真正的 datetime.date，供輸出
# Excel 時寫成日期型別儲存格用；往後若有供應商新增其他格式，在這裡
# 多加一個 regex 分支即可，不用改任何供應商自己的 parse_xxx()。
_DATE_FMT_DD_MON_YYYY = re.compile(
    r"^(\d{1,2})(?:\s+|\.)([A-Za-z]{3})(?:\s+|\.)(\d{4})$"
)
_DATE_FMT_YYYY_MON_DD = re.compile(
    r"^(\d{4})-([A-Za-z]{3})-(\d{1,2})$"
)
_DATE_FMT_DD_SLASH_MM_SLASH_YYYY = re.compile(
    r"^(\d{1,2})/(\d{1,2})/(\d{4})$"
)
_DATE_MONTH_ABBR_TO_NUM = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12,
}


def _coerce_to_date(value) -> Optional[datetime.date]:
    """把擷取結果裡「語意是日期」的值轉成真正的 datetime.date，轉不出來
    (None、"N/A"、空字串、格式不認得) 就回傳 None，呼叫端視為「這欄不是
    日期，照原樣輸出」，不要讓這個函式拋例外中斷整個輸出流程。
    """
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    if not isinstance(value, str):
        return None
    text = value.strip()
    if not text or _is_na(text):
        return None

    m = _DATE_FMT_DD_MON_YYYY.match(text)
    if m:
        day_s, month_abbr, year_s = m.groups()
        month = _DATE_MONTH_ABBR_TO_NUM.get(month_abbr.upper())
        if month:
            try:
                return datetime.date(int(year_s), month, int(day_s))
            except ValueError:
                return None
        return None

    m = _DATE_FMT_YYYY_MON_DD.match(text)
    if m:
        year_s, month_abbr, day_s = m.groups()
        month = _DATE_MONTH_ABBR_TO_NUM.get(month_abbr.upper())
        if month:
            try:
                return datetime.date(int(year_s), month, int(day_s))
            except ValueError:
                return None
        return None

    m = _DATE_FMT_DD_SLASH_MM_SLASH_YYYY.match(text)
    if m:
        day_s, month_s, year_s = m.groups()
        try:
            return datetime.date(int(year_s), int(month_s), int(day_s))
        except ValueError:
            return None

    return None


def normalize_value(value) -> str:
    """統一轉成好比對的字串：去除空白/大小寫差異，日期物件轉成固定格式，
    數字不管原本是 int/float/字串、有沒有多餘的小數位 (例如 '11.810' 對
    11.81、'1076510528' 對 1076510528)，都統一轉成同一種數字字串比較，
    這樣不會因為型態或多餘的零而誤判成不相符。
    """
    if value is None:
        return ""
    if isinstance(value, (datetime.date, datetime.datetime)):
        # 日期一律用「DD Mon YYYY」(月份英文縮寫) 表示，和 suppliers.py 的
        # _format_date_en() 統一；大小寫交給下面文字分支一致處理 (轉大寫)，
        # 這裡先轉成和文字分支相同的大寫，確保兩邊比對時格式完全一致。
        return value.strftime("%d %b %Y").upper()

    # 先嘗試當數字比對 (int/float本身，或看起來像數字的字串)
    if isinstance(value, (int, float)):
        num = float(value)
        return f"{num:.6f}".rstrip("0").rstrip(".")

    text = str(value).strip()
    if re.fullmatch(r"-?\d+(\.\d+)?", text):
        num = float(text)
        return f"{num:.6f}".rstrip("0").rstrip(".")

    # 使用者的正確答案 Excel 有時會把金額打成「英式千分位逗號」文字
    # (例如 '100,000') 而不是真正的 Excel 數字儲存格，去掉逗號後如果變成
    # 純數字，一樣當數字比對，不要因為多了逗號就跟擷取結果 (int 100000)
    # 判定成不相符。
    no_comma = text.replace(",", "")
    if re.fullmatch(r"-?\d+(\.\d+)?", no_comma):
        num = float(no_comma)
        return f"{num:.6f}".rstrip("0").rstrip(".")

    # 有些供應商的正確答案 Excel 日期欄位是純文字 "DD/MM/YYYY"（不是 Excel
    # 日期型別儲存格），但同一家供應商不同月份的帳單，使用者準備正確答案
    # 時有時候又存成真正的 Excel 日期型別（上面的 datetime 分支會轉成
    # "DD Mon YYYY"）。兩種情況都要能跟系統擷取出來的日期值比對成功，所以
    # 純文字的 "DD/MM/YYYY" 也統一轉成跟 datetime 分支一樣的 "DD Mon YYYY"
    # 格式，不要只是單純大寫比較文字。句點分隔的 "DD.MM.YYYY" (例如 KN 空運
    # 答案 "10.06.2025") 也一樣。
    slash_date_m = re.fullmatch(r"(\d{1,2})[/.](\d{1,2})[/.](\d{4})", text)
    if slash_date_m:
        try:
            day_i, month_i, year_i = (int(g) for g in slash_date_m.groups())
            parsed = datetime.date(year_i, month_i, day_i)
        except ValueError:
            parsed = None
        if parsed is not None:
            return parsed.strftime("%d %b %Y").upper()

    # 其他供應商專用的日期文字格式 ("2026-SEP-02"、"31.Aug.2026"，見
    # _coerce_to_date() 上方說明) 也統一轉成 "DD MON YYYY"：正確答案 Excel
    # 改成真正的日期型別儲存格後 (上面 datetime 分支)，擷取出來的日期文字
    # 才比對得起來。
    date_value = _coerce_to_date(text)
    if date_value is not None:
        return date_value.strftime("%d %b %Y").upper()

    text = text.upper()
    text = re.sub(r"\s+", " ", text)
    return text.rstrip(",")


def _is_na(value) -> bool:
    """判斷『正確答案』這一格是不是 N/A (代表這欄位對這張帳單不適用)。
    涵蓋 None、空字串、不分大小寫/前後空白的 'N/A' 字樣，以及 pandas 幫
    「某幾列有值、某幾列整欄都空白」的欄位 (例如 vat 這種逐筆明細欄位，
    不是每一筆費用都有稅額) 自動補的 NaN (浮點數)——一整欄混著數字跟空白
    時，pandas 會把空白轉成 float('nan')，不是 None，字面轉成字串是
    'nan'，不先特判會被誤判成「有值、只是文字是 nan」而算不相符。
    """
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    text = str(value).strip()
    return text == "" or text.upper() == "N/A"


def compare_rows(rows: List[Dict], reference_rows: pd.DataFrame) -> List[Dict]:
    """比對「一張帳單」擷取出來的 rows (可能有多筆費用) 與正確答案
    (reference_rows：同一張發票在正確答案 Excel 裡的所有列)。

    抬頭欄位只比對一次 (取第一列)；費用明細 (description+amount) 用集合
    比對，不要求擷取順序跟正確答案一致。
    """
    if not rows or reference_rows is None or reference_rows.empty:
        return []

    results = []
    first_row = rows[0]
    ref_first = reference_rows.iloc[0].to_dict()

    for code in HEADER_FIELD_CODES:
        ext_val = first_row.get(code, NA)
        ref_val = ref_first.get(code, NA)
        if _is_na(ref_val):
            # 正確答案本來就是 N/A：代表這個欄位對這張帳單不適用 (例如沒有
            # 提單的帳單，MBL NO./Container no. 本來就沒有值)，不管擷取結果
            # 抓到什麼都不該算錯，也不計入正確率的分母 (使用者需求)。
            status = "⚪ 略過（正確答案為 N/A）"
        else:
            match = normalize_value(ext_val) == normalize_value(ref_val)
            status = "✅ 相符" if match else "❌ 不相符"
        results.append({
            "欄位": DISPLAY_HEADERS[code].split("\n")[0],
            "擷取結果": ext_val,
            "正確答案": ref_val,
            "結果": status,
        })

    # 大部分供應商每張帳單的抬頭欄位 (發票號碼、Order No.、船名…等) 整張
    # 帳單只有一個值，已經在上面「只用第一列」比對過一次。但少數供應商
    # (例如 PT. TRANS DAYA PRIMA / YJE_TATA 的 Order No.，或 DHL 一張發票
    # 裡有好幾個提單/航次時的 B/L NO.、開船日、啟運港、目的港) 是「每一筆
    # 費用明細各自有自己的抬頭值」，光比對第一列沒辦法驗證到第二、第三筆
    # 費用明細的單號/日期/港口對不對。這裡動態偵測：只要擷取結果或正確
    # 答案裡，同一張帳單的某個抬頭欄位出現超過一種不同的值，就自動把該
    # 欄位也一起納入「逐筆費用明細」的比對 key (不限於 order_no，未來任何
    # 供應商只要有「每筆費用各自不同」的抬頭欄位都會自動被抓進來)；否則
    # 沿用原本「只比 Description/Amount」的做法，對其餘欄位整張帳單只有
    # 一個值的供應商完全不影響既有的比對結果跟顯示格式。
    varying_codes = []
    for code in HEADER_FIELD_CODES:
        ext_vals = {normalize_value(r.get(code)) for r in rows}
        ref_vals = (
            {normalize_value(v) for v in reference_rows[code]}
            if code in reference_rows.columns else set()
        )
        if len(ext_vals) > 1 or len(ref_vals) > 1:
            varying_codes.append(code)

    # 逐筆費用明細比對的欄位：description/amount 一定都要比對，但像 vat
    # 這種「比較晚才加進來」的明細欄位，舊版供應商範本 (還沒加這欄) 整份
    # 正確答案根本沒有這個資料，不是「這筆費用剛好沒有稅額」。如果不分青
    # 紅皂白把 vat 一律納入比對 key，擷取結果真的抓到稅額時反而會被這種
    # 舊範本誤判成不相符 (正確答案那邊因為整欄都沒資料被當成空白，兩邊就
    # 對不起來)。這裡只有當「這份正確答案裡這個欄位確實有填過至少一筆非
    # 空白值」才納入比對，沒有填過就當作這個欄位對這張帳單不適用，不納入
    # 比對 key (跟 HEADER_FIELD_CODES 的 varying_codes 是同一種「選擇性
    # 啟用」精神，只是判斷條件不同：varying_codes 看「是否不只一種值」，
    # 這裡看「正確答案是否真的用過這個欄位」)。
    def _norm_item_field(value):
        return "" if _is_na(value) else normalize_value(value)

    item_codes = ["description", "amount"]
    for code in ITEM_FIELD_CODES:
        if code in item_codes:
            continue
        ref_has_data = code in reference_rows.columns and reference_rows[code].apply(
            lambda v: not _is_na(v)
        ).any()
        if ref_has_data:
            item_codes.append(code)
    item_labels = [DISPLAY_HEADERS[c].split("\n")[0] for c in item_codes]

    if varying_codes:
        extra_labels = [DISPLAY_HEADERS[c].split("\n")[0] for c in varying_codes]
        item_label = "/".join(item_labels) + "/" + "/".join(extra_labels)

        # varying_codes 裡的欄位也要用 _norm_item_field (不是 normalize_value)
        # 把 N/A 統一轉成同一個空字串再比較：像 KUEHNE NAGEL 的 Volume 這種
        # 「只有某幾筆費用才有值 (例如逐櫃費用)、其他筆留空」的欄位，擷取結果
        # 留空時填的是 NA 常數 "N/A"，但正確答案 Excel 整欄裡只要「有些列有
        # 值、有些列空白」就會被 pandas 自動補 float('nan')——這兩種「空白」
        # 字面轉成字串不一樣 ("N/A" vs "nan")，如果直接用 normalize_value()
        # 逐字比較會被誤判成不相符，所以這裡跟 item_codes 用同一套、對 N/A
        # 有特判的 _norm_item_field()。
        def _item_key(r):
            return tuple(_norm_item_field(r.get(c)) for c in item_codes) + \
                tuple(_norm_item_field(r.get(c)) for c in varying_codes)
    else:
        item_label = "/".join(item_labels)

        def _item_key(r):
            return tuple(_norm_item_field(r.get(c)) for c in item_codes)

    def _item_fmt(t):
        return " / ".join(str(x) for x in t)

    ext_items = {_item_key(r) for r in rows}
    ref_items = {_item_key(r) for _, r in reference_rows.iterrows()}

    for key in sorted(ext_items & ref_items):
        results.append({"欄位": item_label, "擷取結果": _item_fmt(key),
                         "正確答案": _item_fmt(key), "結果": "✅ 相符"})
    for key in sorted(ext_items - ref_items):
        results.append({"欄位": item_label, "擷取結果": _item_fmt(key),
                         "正確答案": "(正確答案沒有)", "結果": "❌ 不相符"})
    for key in sorted(ref_items - ext_items):
        results.append({"欄位": item_label, "擷取結果": "(未擷取到)",
                         "正確答案": _item_fmt(key), "結果": "❌ 不相符"})

    return results


def _group_rows_by_invoice(rows: List[Dict]) -> Dict[str, List[Dict]]:
    groups: Dict[str, List[Dict]] = {}
    for row in rows:
        inv_no = normalize_value(row.get("invoice_no"))
        groups.setdefault(inv_no, []).append(row)
    return groups


def compare_multi_invoice_rows(rows: List[Dict], reference_df: Optional[pd.DataFrame]) -> List[Dict]:
    """multi_page 供應商專用比對：rows 裡可能混著好幾張不同發票各自的列
    (一份 PDF 有好幾頁、每頁是一張獨立發票)，所以不能像一般供應商一樣只
    取第一列當抬頭比對——這裡先用 invoice_no 把 rows 分組，每組各自去
    reference_df (完整的正確答案表格，不是預先篩過某一張發票的子集) 找出
    同一個 invoice_no 的正確答案列，分別呼叫 compare_rows() 後把所有結果
    串起來，這樣每張發票的抬頭欄位都會各自比對一次。

    ⚠️ 重要防呆 (2026-09 修正，適用「所有」multi_page 供應商，包含未來新增
    的供應商，不用每加一家就重寫一次)：
    如果某張發票的 invoice_no 在 reference_df 裡完全找不到對應的列 (最常見
    情境：使用者一次上傳好幾家供應商的 PDF，但『正確答案 Excel』只包含其中
    一部分供應商的資料，或是漏放了某張發票)，舊版做法是這張發票直接被
    compare_rows() 回傳空 list、整個跳過不計入比對結果。這樣會造成一個很
    隱密的假象：compute_accuracy() 看到「這份 PDF 完全沒有可比對的欄位」時
    會依照『正確答案全是 N/A』的防禦邏輯回傳 1.0 (100%)，介面上看起來像是
    滿分過關，但「逐欄比對明細」卻完全沒有這張發票的任何一列——使用者只會
    看到正確率 100% 卻找不到比對明細，很容易誤以為是程式漏寫了這家供應商
    的比對邏輯 (但其實只是正確答案 Excel 裡沒有這張發票的資料)。
    這裡改成：找不到對應正確答案時，明確產生一筆「❌ 不相符」的提示列，
    這樣 (1) compute_accuracy() 會正確反映『沒有比對到』不是『滿分』，
    (2) 「逐欄比對明細」一定會顯示這張發票、並清楚寫出原因，不會整張消失。
    """
    if not rows or reference_df is None or reference_df.empty or "invoice_no" not in reference_df.columns:
        return []
    results: List[Dict] = []
    groups = _group_rows_by_invoice(rows)
    ref_inv_norm = reference_df["invoice_no"].map(normalize_value)
    for inv_no, group_rows in groups.items():
        ref_subset = reference_df[ref_inv_norm == inv_no]
        if ref_subset.empty:
            results.append({
                "欄位": "⚠️ Invoice No",
                "擷取結果": group_rows[0].get("invoice_no", NA) if group_rows else NA,
                "正確答案": "(正確答案 Excel 裡找不到這個發票號碼，請確認有沒有漏放這張發票的資料)",
                "結果": "❌ 不相符",
            })
            continue
        results.extend(compare_rows(group_rows, ref_subset))
    return results


def compute_accuracy(results: List[Dict]) -> float:
    """正確率 = 相符欄位數 / 有意義的比對欄位數。

    正確答案是 N/A 的欄位 (結果以 "⚪" 開頭) 完全不計入分母，符合「正確答案
    是 N/A 時，不管擷取結果如何都不影響正確率」的需求；如果一張帳單所有
    欄位的正確答案都是 N/A (理論上不會發生，但防禦性處理)，視為 100%。
    """
    counted = [r for r in results if not r["結果"].startswith("⚪")]
    if not counted:
        return 1.0
    matched = sum(1 for r in counted if r["結果"].startswith("✅"))
    return matched / len(counted)


def quality_score(rows: List[Dict]) -> float:
    """沒有正確答案 Excel 時的自我檢查分數：欄位有抓到值(不是N/A)的比例。
    僅供迴圈挑選較好的擷取設定用，不是跟人工核對過的正確率。
    """
    if not rows:
        return 0.0
    total = ok = 0
    for row in rows:
        for v in row.values():
            total += 1
            if v and v != NA:
                ok += 1
    return ok / total if total else 0.0


def audit_invoice_totals(
    supplier_key: Optional[str],
    pdf_path: str,
    rows: List[Dict],
    **pdfplumber_kwargs,
) -> List[Dict]:
    """稽核『系統擷取出來的費用總額』是否跟帳單本身寫的總額一致，有三項
    各自獨立的稽核 (哪一項有做，取決於這家供應商有沒有實作對應的擷取
    規則，見 register_supplier())：

    1. 未稅總額稽核 (需要 extract_total)：系統擷取到的費用明細加總 (不含
       VAT) 要跟帳單本身寫的『未稅總額』(Total (Excl. VAT) / Sub Total /
       DPP…，依供應商命名不同) 一致。
    2. 含稅總額稽核 (需要 extract_grand_total)：未稅費用加總 + 每筆費用
       的『加總類』VAT 加總 (排除 deduction_prefixes 指定的扣除類稅金，
       見下)，要等於帳單本身寫的『含稅總額』(Total Amount Due / Amount
       Due / Total / GRAND TOTAL…，依供應商命名不同)。這一項不要求供應商
       一定要有逐筆 VAT (沒有 VAT 欄位的供應商，VAT 加總自然是 0)。
    3. 實付金額稽核 (需要 extract_final_total)：含稅總額 (第 2 項的計算
       結果) 再減掉 deduction_prefixes 指定的『扣除類』VAT 加總 (例如
       印尼供應商常見的 PPH 代扣稅)，要等於帳單本身寫的『實付金額』
       (例如 "TOTAL YANG DIBAYAR")。沒有扣除類稅金的供應商，扣除加總
       自然是 0 (等同『含稅總額 = 實付金額』)。

    三項稽核互相獨立：只實作其中幾項的供應商，只會做對應的那幾項；一項都
    沒實作的供應商回傳空列表，代表『這家供應商還沒辦法自動稽核』，不是
    「一致」也不是「不一致」，呼叫端不應顯示警示。

    回傳一份清單，每張發票 (multi_page 供應商一份 PDF 可能有好幾張) 各一筆
    (視有實作哪些規則，以下欄位可能是 None，代表『這項稽核不適用』)：
      invoice_no      發票號碼 (抓不到時是 N/A)
      extracted_sum   系統擷取到的費用明細加總 (不含 VAT，跳過
                       audit_exclude 指定的合成列，例如 YJE_TATA 的 "DPP")
      invoice_total   帳單原文抓到的未稅總額 (第 1 項稽核用)
      diff            extracted_sum - invoice_total
      match           第 1 項稽核是否一致 (容許 ±1 的四捨五入誤差)
      vat_sum         系統擷取到的『加總類』VAT 加總 (第 2 項稽核用，不含
                       deduction_prefixes 指定的扣除類稅金)
      grand_total     帳單原文抓到的含稅總額 (第 2 項稽核用)
      grand_diff      (extracted_sum + vat_sum) - grand_total
      grand_match     第 2 項稽核是否一致 (容許 ±1 的四捨五入誤差)
      deduction_sum   系統擷取到的『扣除類』VAT 加總 (第 3 項稽核用，例如
                       PPH 代扣稅，deduction_prefixes 沒設定的供應商固定是
                       0，None 表示沒有足夠資訊可以計算)
      final_total     帳單原文抓到的實付金額 (第 3 項稽核用)
      extracted_final (extracted_sum + vat_sum) - deduction_sum
      final_diff      extracted_final - final_total
      final_match     第 3 項稽核是否一致 (容許 ±1 的四捨五入誤差)

    呼叫端 (app.py / verify_app.py) 只需要對 match=False / grand_match=False
    / final_match=False 的項目顯示警示；對應欄位是 None 時代表這項稽核
    不適用，不要顯示警示。
    """
    entry = SUPPLIER_REGISTRY.get(supplier_key) if supplier_key else None
    if not entry:
        return []
    extract_total_fn = entry.get("extract_total")
    extract_grand_total_fn = entry.get("extract_grand_total")
    extract_final_total_fn = entry.get("extract_final_total")
    sheet_totals_fn = entry.get("sheet_totals") if is_spreadsheet(pdf_path) else None
    if not (extract_total_fn or extract_grand_total_fn or extract_final_total_fn
            or sheet_totals_fn):
        return []

    exclude_descs = {normalize_value(d) for d in entry.get("audit_exclude", ())}
    deduction_prefixes = tuple(p.upper() for p in entry.get("deduction_prefixes", ()))
    need_vat_split = bool(extract_grand_total_fn or extract_final_total_fn or sheet_totals_fn)

    def _sum_rows(group_rows: List[Dict]) -> float:
        total = 0.0
        for r in group_rows:
            if normalize_value(r.get("description")) in exclude_descs:
                continue
            amt = r.get("amount")
            if isinstance(amt, (int, float)) and not isinstance(amt, bool):
                total += amt
        return total

    def _sum_vat_add_and_deduct(group_rows: List[Dict]):
        """把每筆費用的 VAT 分成『加總類』(一般 VAT/PPN，計入含稅總額) 跟
        『扣除類』(deduction_prefixes 指定的代扣稅，例如 PPH，從含稅總額
        扣掉才等於實付金額) 兩份加總。沒有設定 deduction_prefixes 的
        供應商，所有 VAT 都算加總類，等同舊版行為 (vat_sum = 全部 VAT
        加總、deduction_sum = 0)，不影響既有供應商的稽核結果。
        """
        add_total = 0.0
        deduct_total = 0.0
        for r in group_rows:
            vat = r.get("vat")
            if not isinstance(vat, (int, float)) or isinstance(vat, bool):
                continue
            desc_norm = normalize_value(r.get("description") or "").upper()
            if deduction_prefixes and any(desc_norm.startswith(p) for p in deduction_prefixes):
                deduct_total += vat
            else:
                add_total += vat
        return add_total, deduct_total

    def _build_result(group_rows: List[Dict], invoice_total, grand_total, final_total,
                      grand_tolerance: float = 1, final_tolerance: float = 1,
                      vat_total=None) -> Dict:
        """grand_tolerance / final_tolerance：含稅總額 / 實付金額允許的誤差
        (預設 ±1)。帳單合計的稅額若是「未稅合計 × 稅率」、明細稅額卻是逐列四捨
        五入 (例如 CONG THANH 代墊款)，Excel 供應商可透過 sheet_totals 放寬到
        「每列 0.5」。vat_total：帳單本身的 VAT 合計 (Excel 帳單的合計列)，
        跟擷取到的 VAT 加總勾稽 (容許誤差同 final_tolerance)。"""
        extracted_sum = _sum_rows(group_rows)
        vat_sum = deduction_sum = None
        if need_vat_split:
            vat_sum, deduction_sum = _sum_vat_add_and_deduct(group_rows)
        extracted_grand = (
            extracted_sum + vat_sum if vat_sum is not None else None
        )
        extracted_final = (
            extracted_grand - deduction_sum
            if extracted_grand is not None and deduction_sum is not None
            and (extract_final_total_fn or sheet_totals_fn) else None
        )
        result = {
            "invoice_no": group_rows[0].get("invoice_no", NA),
            "extracted_sum": extracted_sum,
            "invoice_total": invoice_total,
            "diff": extracted_sum - invoice_total if invoice_total is not None else None,
            "match": abs(extracted_sum - invoice_total) <= 1 if invoice_total is not None else None,
            "vat_sum": vat_sum,
            "vat_total": vat_total,
            "vat_diff": (vat_sum - vat_total
                         if vat_sum is not None and vat_total is not None else None),
            "vat_match": (abs(vat_sum - vat_total) <= final_tolerance
                          if vat_sum is not None and vat_total is not None else None),
            "grand_total": grand_total,
            "grand_diff": (
                extracted_grand - grand_total
                if extracted_grand is not None and grand_total is not None else None
            ),
            "grand_match": (
                abs(extracted_grand - grand_total) <= grand_tolerance
                if extracted_grand is not None and grand_total is not None else None
            ),
            "deduction_sum": deduction_sum,
            "final_total": final_total,
            "extracted_final": extracted_final,
            "final_diff": (
                extracted_final - final_total
                if extracted_final is not None and final_total is not None else None
            ),
            "final_match": (
                abs(extracted_final - final_total) <= final_tolerance
                if extracted_final is not None and final_total is not None else None
            ),
        }
        return result

    results: List[Dict] = []

    if sheet_totals_fn:
        # Excel 帳單：總額直接取自帳單本身的合計列 (sheet_totals)，依
        # invoice_no 對回擷取到的明細
        groups = _group_rows_by_invoice(rows)
        whole = None
        per_invoice: Dict[str, Dict] = {}
        for t in sheet_totals_fn(read_spreadsheet(pdf_path)):
            args = (t.get("total"), t.get("grand_total"), t.get("final_total"),
                    t.get("grand_tolerance", 1), t.get("final_tolerance", 1),
                    t.get("vat_total"))
            if t.get("invoice_no") is None:
                # 整份檔案的合計 (例如表格最後的合計列)：用全部明細勾稽
                whole = _build_result(rows, *args)
                continue
            key = normalize_value(t.get("invoice_no"))
            if groups.get(key):
                per_invoice[key] = _build_result(groups[key], *args)
        # 每張發票的結果：有整份檔案的勾稽就套用到每張發票；這張發票自己的
        # 含稅總額勾稽 (grand_*) 優先 (例如 RỒNG Á CHÂU：每張發票的費用 +
        # VAT PASS THROUGH = 該發票的 TOTAL)
        for key, group_rows in groups.items():
            if not group_rows:
                continue
            own = per_invoice.get(key)
            if whole is None and own is None:
                continue
            result = dict(whole or own)
            if whole is not None and own is not None:
                result.update({k: own[k] for k in ("vat_sum", "grand_total", "grand_diff",
                                                     "grand_match")})
            result["invoice_no"] = group_rows[0].get("invoice_no", NA)
            results.append(result)
        return results

    if entry.get("multi_page_pdf"):
        pages = supplier_pages(pdf_path, supplier_key, **pdfplumber_kwargs)
        groups = _group_rows_by_invoice(rows)
        # key(正規化發票號碼) -> 原始發票號碼字串，用來在「不是發票抬頭頁」
        # 的頁面裡，用文字比對的方式判斷這一頁屬於哪一張發票 (見下方)。
        raw_invoice_no = {
            normalize_value(gr[0].get("invoice_no")): gr[0].get("invoice_no")
            for gr in groups.values() if gr
        }

        totals_found: Dict[str, float] = {}
        grand_totals_found: Dict[str, float] = {}
        final_totals_found: Dict[str, float] = {}

        def _assign(found: Dict[str, float], value, page_upper: str) -> None:
            """這一頁抓到的總額歸給頁面上印的那張發票 (整個字比對，避免 "4521"
            誤配 "45217")；擷取函式回傳 {發票號碼: 總額} 時逐票歸屬 (越南 KN 附件清單)。"""
            if isinstance(value, dict):
                for k, v in value.items():
                    nk = normalize_value(k)
                    if nk in raw_invoice_no and v is not None:
                        found.setdefault(nk, v)
                return
            for inv_key, inv_raw in raw_invoice_no.items():
                if inv_key in found or not inv_raw:
                    continue
                if re.search(rf"(?<![A-Z0-9]){re.escape(str(inv_raw).upper())}(?![A-Z0-9])", page_upper):
                    found[inv_key] = value
                    break

        for page_text in pages:
            # 大部分供應商的『未稅總額』就印在發票抬頭頁本身 (detect_fn 認得
            # 的那一頁)；但少數供應商 (例如 DSV) 是「一張發票的資料跨好幾頁」
            # ，未稅總額印在後面的附屬頁 (例如 Page 2 of 2 的條款/簽收頁)，
            # 那一頁不會被 detect_fn 認出來，所以這裡不限定只看 detect_fn
            # 認得的頁面，而是每一頁都試著抓抓看未稅總額；抓到之後，再用
            # 『這一頁的文字裡有沒有出現某張發票的發票號碼』來判斷這個未稅
            # 總額屬於哪一張發票 (發票抬頭頁跟附屬頁通常都會重複印一次發票
            # 號碼，例如 DSV 附屬頁開頭就是 "INVOICE ID610413182")。
            if extract_total_fn:
                try:
                    invoice_total = extract_total_fn(page_text)
                except Exception:
                    invoice_total = None
                if invoice_total is not None:
                    _assign(totals_found, invoice_total, page_text.upper())
            if extract_grand_total_fn:
                try:
                    grand_total = extract_grand_total_fn(page_text)
                except Exception:
                    grand_total = None
                if grand_total is not None:
                    _assign(grand_totals_found, grand_total, page_text.upper())
            if extract_final_total_fn:
                try:
                    final_total = extract_final_total_fn(page_text)
                except Exception:
                    final_total = None
                if final_total is not None:
                    _assign(final_totals_found, final_total, page_text.upper())

        for inv_key in set(totals_found) | set(grand_totals_found) | set(final_totals_found):
            group_rows = groups.get(inv_key) or []
            if not group_rows:
                continue
            results.append(_build_result(
                group_rows, totals_found.get(inv_key), grand_totals_found.get(inv_key),
                final_totals_found.get(inv_key),
            ))
    else:
        # 只用帳單頁勾稽 (附件頁的金額不算，見 invoice_text())
        text = invoice_text(pdf_path, supplier_key, **pdfplumber_kwargs)[0]
        invoice_total = None
        grand_total = None
        final_total = None
        if extract_total_fn:
            try:
                invoice_total = extract_total_fn(text)
            except Exception:
                invoice_total = None
        if extract_grand_total_fn:
            try:
                grand_total = extract_grand_total_fn(text)
            except Exception:
                grand_total = None
        if extract_final_total_fn:
            try:
                final_total = extract_final_total_fn(text)
            except Exception:
                final_total = None
        vat_total = None
        if entry.get("extract_vat_total"):
            try:
                vat_total = entry["extract_vat_total"](text)
            except Exception:
                vat_total = None
        totals = (invoice_total, grand_total, final_total)
        if any(isinstance(t, dict) for t in totals):
            # 一份帳單多票、每票各自有總額 (例如越南 DSV 對帳表)：擷取函式回傳
            # {發票號碼: 總額}，逐票勾稽
            norm = [{normalize_value(k): v for k, v in t.items()} if isinstance(t, dict)
                    else {} for t in totals]
            for key, group_rows in _group_rows_by_invoice(rows).items():
                picked = [n.get(key) for n in norm]
                if group_rows and any(v is not None for v in picked):
                    results.append(_build_result(group_rows, *picked))
            return results
        if (invoice_total is not None or grand_total is not None or final_total is not None) and rows:
            results.append(_build_result(rows, invoice_total, grand_total, final_total,
                                         vat_total=vat_total))
    return results


def mismatched_invoice_numbers(audit_results: List[Dict]) -> set:
    """從 audit_invoice_totals() 的結果裡，挑出『任何一項稽核不一致』的
    發票號碼 (正規化後的字串)，給輸出 Excel 時標紅字用 (見 build_excel()
    的 red_invoice_nos 參數、verify_app.py 的 _build_result_excel())。
    三項稽核 (未稅/含稅/實付) 只要有一項明確判定為不一致 (False，不是
    None／不適用) 就算這張發票有問題。
    """
    bad: set = set()
    for ar in audit_results:
        if ar.get("match") is False or ar.get("grand_match") is False or ar.get("final_match") is False:
            inv_no = ar.get("invoice_no")
            if inv_no not in (None, NA):
                bad.add(normalize_value(inv_no))
    return bad


# ---------------------------------------------------------------------------
# 6. 迴圈：反覆嘗試不同擷取設定，直到 100% 正確率 (或找出目前能達到的最高值)
# ---------------------------------------------------------------------------

def find_best_extraction(
    pdf_path: str,
    supplier_key: Optional[str] = None,
    reference_rows: Optional[pd.DataFrame] = None,
    configs: Optional[List[Dict]] = None,
    max_attempts: int = MAX_EXTRACTION_ATTEMPTS,
):
    """對同一份 PDF 重複嘗試擷取，直到正確率達到 100% 或用完 max_attempts
    次嘗試為止 (預設最多 100 次，每次用不同的 pdfplumber 擷取參數)。
    每嘗試一次都重新評分，一旦達到 100% 就立刻停止 (不用浪費剩餘次數)；
    如果 100 次都試完仍未達到 100%，回傳「這 100 次裡分數最高」的那次結果，
    呼叫端可以據此產出核對報告 (書面記錄哪些欄位仍然對不上)。

    「學習記憶」：如果這家供應商之前已經學到某組設定分數最高，這裡會把
    那組設定排到第 1 次嘗試 (參見 _load_learned_configs())，通常能讓正確
    答案在第 1 次就命中，不用每次都從頭試。跑完之後，只要這次找到的最佳
    分數比之前記錄的更高，就會呼叫 _save_learned_config() 更新記憶，讓
    系統隨著使用次數增加、愈來愈快抓到高分設定。

    回傳一個 dict，包含：
      supplier_key    辨識到的供應商 key (或 None)
      rows            最佳一次的展開列 (List[Dict])
      score           最佳分數 0~1
      config          最佳一次採用的擷取參數
      results         該次的逐欄比對明細 (沒有正確答案時為 None)
      attempts        全部嘗試紀錄 (每筆含第幾次、用的設定、當次分數)
      attempts_used   實際用了幾次嘗試 (達到100%會提早停止，< max_attempts)
      reached_100     是否有達到 100% 正確率 (沒有正確答案可比對時恆為 False)
      learned_applied 這次是否有套用「之前學到的最佳設定」當第 1 次嘗試

    注意：對 multi_page 供應商 (一份 PDF 裡有好幾張各自獨立的發票)，
    reference_rows 這裡預期是「完整」的正確答案表格 (不要預先篩成某一張
    發票)，內部會用 invoice_no 自己分組比對；其他一般供應商則維持原本的
    用法：reference_rows 是呼叫端已經篩好、只屬於這張發票的那幾列。
    """
    configs = list(configs or EXTRACTION_CONFIGS)[:max_attempts]
    if is_spreadsheet(pdf_path):
        configs = [{}]  # Excel 帳單沒有 pdfplumber 參數可調，讀一次就好

    learned_entry = _load_learned_configs().get(supplier_key) if supplier_key else None
    learned_config = learned_entry["config"] if learned_entry else None
    if learned_config is not None:
        # 把「學習到的最佳設定」排到第一個嘗試，其餘設定照原順序排在後面
        # (去除重複，避免同一組設定被試兩次)。
        configs = [learned_config] + [c for c in configs if c != learned_config]
    configs = configs[:max_attempts]

    is_multi = is_multi_page_supplier(supplier_key)

    attempts: List[Dict] = []
    best = None
    detected_key = supplier_key
    reached_100 = False

    for i, cfg in enumerate(configs, start=1):
        if is_spreadsheet(pdf_path):
            key, rows, _ = parse_invoice_pdf(pdf_path, supplier_key)
            is_multi = True  # Excel 帳單一份檔案可能有好幾張對帳單
        elif is_multi:
            # multi_page 供應商：這份 PDF 可能是好幾張各自獨立的發票 (一頁
            # 一張)，rows 會混著所有頁面/發票的列，reference_rows 這裡預期
            # 是「完整」的正確答案表格 (呼叫端不要預先篩成某一張發票)，
            # 靠 compare_multi_invoice_rows() 自己用 invoice_no 分組比對。
            key = supplier_key
            rows = _parse_multi_page_pdf(pdf_path, key, **cfg)
        else:
            text = extract_text_from_pdf(pdf_path, **cfg)
            key = supplier_key or detect_supplier(text)
            if key and key in SUPPLIER_REGISTRY:
                parsed = SUPPLIER_REGISTRY[key]["parse"](invoice_text(pdf_path, key, **cfg)[0])
            else:
                parsed = {"header": {}, "items": [{}]}
            rows = expand_to_rows(parsed)
        detected_key = detected_key or key

        has_ref = reference_rows is not None and not reference_rows.empty
        if has_ref:
            results = compare_multi_invoice_rows(rows, reference_rows) if is_multi \
                else compare_rows(rows, reference_rows)
            score = compute_accuracy(results)
        else:
            results = None
            score = quality_score(rows)

        is_learned = (i == 1 and learned_config is not None and cfg == learned_config)
        attempts.append({
            "第幾次嘗試": i,
            "設定": (str(cfg) if cfg else "預設") + ("（沿用學習記憶）" if is_learned else ""),
            "分數(%)": round(score * 100, 1),
        })

        if best is None or score > best["score"]:
            best = {"rows": rows, "score": score, "config": cfg, "results": results}

        if best["score"] >= 1.0:
            reached_100 = True
            break  # 已經 100% 正確率，不用再嘗試剩餘次數

    _save_learned_config(detected_key, best["config"], best["score"])

    return {
        "supplier_key": detected_key,
        "rows": best["rows"],
        "score": best["score"],
        "config": best["config"],
        "results": best["results"],
        "attempts": attempts,
        "attempts_used": len(attempts),
        "reached_100": reached_100,
        "learned_applied": learned_config is not None,
    }


# ---------------------------------------------------------------------------
# 7. Excel 讀寫
# ---------------------------------------------------------------------------

_HEADER_ANCHOR = "INVOICE DATE"


def _field_code_for_header_text(text) -> Optional[str]:
    """把範本表頭儲存格的文字 (可能是 "INVOICE DATE\n發票日期" 這種雙行，
    也可能使用者範本只打了單行 "VAT") 對應回標準欄位代碼，只比對「第一行」
    (跟 DISPLAY_HEADERS 的英文標籤一致)，不分大小寫、去除頭尾空白。
    """
    if text is None:
        return None
    first_line = str(text).split("\n")[0].strip().upper()
    if not first_line:
        return None
    for code in FIELD_CODES:
        if DISPLAY_HEADERS[code].split("\n")[0].strip().upper() == first_line:
            return code
    return None


def read_reference_excel(source) -> pd.DataFrame:
    """讀取『正確答案』Excel。不同供應商範本的表頭可能在第1列或第2列、
    欄位可能從A欄或B欄開始，這裡自動搜尋含有 'INVOICE DATE' 字樣的儲存格
    當作錨點定位，不用替每個供應商範本各寫一次死板的座標。

    欄位對應不是用「第幾欄」這種死板位置算的，而是逐欄讀表頭文字、拿去跟
    DISPLAY_HEADERS 比對找出代表哪個標準欄位 (_field_code_for_header_text())
    ——這樣不管使用者的範本把欄位加在中間 (例如 VAT 插在 AMOUNT 跟
    Order No. 之間) 還是最後面 (例如 Shipper)，都能正確對應，不用每次使用
    者調整範本欄位順序/新增欄位時，連同所有舊範本的欄位位置都要跟著改。
    舊範本沒有的欄位 (例如還沒加 VAT/Shipper 的範本) 單純比對不到、值是
    None，後續 _is_na() 會視為「這欄不適用」略過不計入正確率。
    """
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(source)
    wb = load_workbook(source, data_only=True)
    ws = wb.active

    # 2026-10 起使用者範本在正式表頭上面多加了一列「抓取欄位」(給自己看的
    # 內部欄位命名備忘，例如 "Invoice Date"/"Payment can be Transfer to")，
    # 這一列的文字恰好也可能包含 "INVOICE DATE" 這個錨點字樣，如果不排除
    # 會被誤認成表頭列，比正式的「顯示欄位」列多抓了一列、害後面全部資料
    # 錯位一行。真正的表頭列一定是我們自己 DISPLAY_HEADERS 的雙行格式
    # ("INVOICE DATE\n發票日期" 這種英文+中文兩行)，所以這裡額外要求儲存
    # 格文字裡要有換行符號，排除掉那種只有英文單行的「備忘列」。
    header_row, start_col = None, None
    for r in range(1, 5):
        for c in range(1, 5):
            val = ws.cell(row=r, column=c).value
            if val and "\n" in str(val) and _HEADER_ANCHOR in str(val).upper():
                header_row, start_col = r, c
                break
        if header_row:
            break
    if header_row is None:
        header_row, start_col = 1, 1

    col_to_code: Dict[int, str] = {}
    for c in range(start_col, ws.max_column + 1):
        code = _field_code_for_header_text(ws.cell(row=header_row, column=c).value)
        if code and code not in col_to_code.values():
            col_to_code[c] = code

    records = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=False):
        rec = {code: None for code in FIELD_CODES}
        empty = True
        for col_idx, code in col_to_code.items():
            idx0 = col_idx - 1
            value = row[idx0].value if idx0 < len(row) else None
            rec[code] = value
            if value not in (None, ""):
                empty = False
        if not empty:
            records.append(rec)
    return pd.DataFrame(records, columns=FIELD_CODES)


def build_excel(
    all_rows: List[Dict[str, object]],
    red_invoice_nos: Optional[set] = None,
) -> bytes:
    """輸出成統一格式的 Excel：A欄是辨識到的供應商名稱，B欄起是標準欄位，
    表頭在第1列，資料從第2列開始。不管來源是哪個供應商都放在同一張表，
    方便一次彙整多家供應商、多份帳單的轉檔結果。

    red_invoice_nos (可選)：金額查核機制 (見 audit_invoice_totals()/
    mismatched_invoice_numbers()) 判定『不一致』的發票號碼集合 (正規化後
    的字串)。這是使用者明確要求的通用規則 (不限特定供應商)：只要這張
    發票的查核結果不符，輸出 Excel 裡屬於這張發票的所有列，整列文字都
    改成紅字，方便人工一眼看出哪些發票需要複查。沒有查核結果 (None 或
    沒做查核) 的供應商不受影響，維持原本黑字。
    """
    wb = Workbook()
    ws = wb.active
    # 分頁名稱統一固定寫死 "工作表"（使用者明確要求的全系統通用原則，不限
    # 特定供應商）：不要用供應商名稱或其他描述性字樣，避免像本機執行工具
    # 曾經發生的錯誤 (沿用內嵌範本殘留的 "DWIHARTA" 分頁名稱，導致任何
    # 供應商轉出來的 Excel 分頁名稱都錯誤顯示成 "DWIHARTA")。
    ws.title = "工作表"

    bold = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="DDEBF7")
    wrap_center = Alignment(wrap_text=True, vertical="center", horizontal="center")
    red_font = Font(color="FFFF0000")
    red_bold_font = Font(color="FFFF0000", bold=True)

    ws.cell(row=1, column=1, value="供應商").font = bold
    ws.cell(row=1, column=1).fill = header_fill
    ws.column_dimensions["A"].width = 22

    for idx, code in enumerate(FIELD_CODES, start=2):
        cell = ws.cell(row=1, column=idx, value=DISPLAY_HEADERS[code])
        cell.font = bold
        cell.fill = header_fill
        cell.alignment = wrap_center
        ws.column_dimensions[get_column_letter(idx)].width = 20

    red_invoice_nos = red_invoice_nos or set()
    for r, row in enumerate(all_rows, start=2):
        is_mismatch = normalize_value(row.get("invoice_no")) in red_invoice_nos
        cell_a = ws.cell(row=r, column=1, value=row.get("__supplier_label", NA))
        if is_mismatch:
            cell_a.font = red_bold_font
        for idx, code in enumerate(FIELD_CODES, start=2):
            raw_value = row.get(code, NA)
            if code in TEXT_NOTE_FIELDS:
                raw_value = english_note(raw_value)  # 非英文加英文備註
            cell = ws.cell(row=r, column=idx, value=raw_value)
            if code in DATE_FIELD_CODES:
                date_value = _coerce_to_date(raw_value)
                if date_value is not None:
                    cell.value = date_value
                    cell.number_format = DATE_DISPLAY_NUMBER_FORMAT
            if is_mismatch:
                cell.font = red_font

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_verification_report_excel(score_rows: List[Dict], compare_rows_all: List[Dict]) -> bytes:
    """產出「核對報告」Excel，給下載核對報告按鈕用。

    分頁1「正確率總覽」：每份 PDF 一列，顯示辨識到的供應商、最終正確率、
    用了幾次嘗試 (最多 20 次)、是否達到 100%。
    分頁2「逐欄比對明細」：每個欄位一列，相符/不相符用綠/紅底色標示，
    正確答案是 N/A (不計入正確率) 的欄位用灰底標示，方便直接找出哪些欄位
    還對不上，人工只需要複查標紅的部分。
    """
    wb = Workbook()
    bold = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="DDEBF7")
    green = PatternFill("solid", fgColor="C6EFCE")
    red = PatternFill("solid", fgColor="FFC7CE")
    grey = PatternFill("solid", fgColor="E7E6E6")  # 正確答案為 N/A、不計入正確率的欄位

    ws1 = wb.active
    ws1.title = "正確率總覽"
    cols1 = ["來源檔案", "辨識供應商", "Invoice No", "正確率", "嘗試次數", "是否達到100%"]
    ws1.append(cols1)
    for cell in ws1[1]:
        cell.font = bold
        cell.fill = header_fill
    for row in score_rows:
        ws1.append([row.get(c, NA) for c in cols1])
    for col_letter, width in zip("ABCDEF", (24, 22, 16, 10, 10, 12)):
        ws1.column_dimensions[col_letter].width = width

    ws2 = wb.create_sheet("逐欄比對明細")
    cols2 = ["來源檔案", "供應商", "欄位", "擷取結果", "正確答案", "結果"]
    ws2.append(cols2)
    for cell in ws2[1]:
        cell.font = bold
        cell.fill = header_fill
    for r in compare_rows_all:
        ws2.append([r.get(c, NA) for c in cols2])
        status = str(r.get("結果", ""))
        if status.startswith("✅"):
            fill = green
        elif status.startswith("⚪"):
            fill = grey
        else:
            fill = red
        ws2.cell(row=ws2.max_row, column=6).fill = fill
    for col_letter, width in zip("ABCDEF", (24, 22, 18, 30, 30, 10)):
        ws2.column_dimensions[col_letter].width = width

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
