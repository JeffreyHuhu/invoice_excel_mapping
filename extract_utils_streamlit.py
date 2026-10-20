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
import json
import math
import datetime
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
    """
    parts = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text(**pdfplumber_kwargs) or ""
            if not text.strip():
                text = _ocr_page_text(pdf_path, i, page)
            parts.append(text)
    return "\n".join(parts)


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
        for i, page in enumerate(pdf.pages):
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


def extract_pages_from_pdf(pdf_path: str, **pdfplumber_kwargs) -> List[str]:
    """跟 extract_text_from_pdf() 一樣，但個別回傳『每一頁』的文字 (不合併成
    一個字串)。給 multi_page 供應商 (一份 PDF 裡有好幾張各自獨立的發票，
    一頁一張) 逐頁辨識/擷取用。同樣會對沒有文字層的頁面自動改用 OCR。
    """
    pages_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text(**pdfplumber_kwargs) or ""
            if not text.strip():
                text = _ocr_page_text(pdf_path, i, page)
            pages_text.append(text)
    return pages_text


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
                       parse_fn: Callable[[str], Dict],
                       multi_page: bool = False,
                       extract_total: Optional[Callable[[str], Optional[float]]] = None,
                       extract_grand_total: Optional[Callable[[str], Optional[float]]] = None,
                       extract_final_total: Optional[Callable[[str], Optional[float]]] = None,
                       deduction_prefixes: Tuple[str, ...] = (),
                       audit_exclude: Tuple[str, ...] = (),
                       force_ocr: bool = False) -> None:
    """註冊一家供應商。

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
    }


def detect_supplier(text: str) -> Optional[str]:
    """依序問過每個已註冊的供應商『這是你的帳單嗎』，回傳第一個承認的 key。
    都沒有承認就回傳 None，呼叫端應提示『無法辨識供應商』。
    """
    for key, entry in SUPPLIER_REGISTRY.items():
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


def list_registered_suppliers() -> List[Dict[str, str]]:
    """回傳目前已註冊的供應商清單 (給網頁上顯示用)。"""
    return [{"key": k, "label": v["label"]} for k, v in SUPPLIER_REGISTRY.items()]


def is_multi_page_supplier(key: Optional[str]) -> bool:
    """這家供應商是不是『一份 PDF 裡有好幾張各自獨立的發票，一頁一張』
    (參見 register_supplier() 的 multi_page 說明)。找不到這家供應商時視為
    False (走一般的單張發票流程)。
    """
    return bool(key and key in SUPPLIER_REGISTRY and SUPPLIER_REGISTRY[key].get("multi_page"))


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
        pages = extract_pages_from_pdf(pdf_path, **pdfplumber_kwargs)
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


def parse_invoice_pdf(pdf_path: str, supplier_key: Optional[str] = None,
                       **pdfplumber_kwargs) -> Tuple[Optional[str], List[Dict], str]:
    """讀取PDF、自動判斷供應商 (若未指定)、用對應規則擷取欄位。
    回傳 (辨識到的供應商key或None, 展開後的rows列表, 原始文字)。

    對 multi_page 供應商 (參見 register_supplier())，rows 可能包含好幾張
    不同發票各自的列 (一頁一張發票)，呼叫端要用 invoice_no 分組。
    """
    text = extract_text_from_pdf(pdf_path, **pdfplumber_kwargs)
    key = supplier_key or detect_supplier(text)

    if is_multi_page_supplier(key):
        rows = _parse_multi_page_pdf(pdf_path, key, **pdfplumber_kwargs)
        return key, rows, text

    if key and key in SUPPLIER_REGISTRY:
        parsed = SUPPLIER_REGISTRY[key]["parse"](text)
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
#      _maersk_format_invoice_date() (為了跟正確答案的文字儲存格比對，刻意
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
    # 格式，不要只是單純大寫比較文字。
    slash_date_m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", text)
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
    if not extract_total_fn and not extract_grand_total_fn and not extract_final_total_fn:
        return []

    exclude_descs = {normalize_value(d) for d in entry.get("audit_exclude", ())}
    deduction_prefixes = tuple(p.upper() for p in entry.get("deduction_prefixes", ()))
    need_vat_split = bool(extract_grand_total_fn or extract_final_total_fn)

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

    def _build_result(group_rows: List[Dict], invoice_total, grand_total, final_total) -> Dict:
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
            and extract_final_total_fn else None
        )
        result = {
            "invoice_no": group_rows[0].get("invoice_no", NA),
            "extracted_sum": extracted_sum,
            "invoice_total": invoice_total,
            "diff": extracted_sum - invoice_total if invoice_total is not None else None,
            "match": abs(extracted_sum - invoice_total) <= 1 if invoice_total is not None else None,
            "vat_sum": vat_sum,
            "grand_total": grand_total,
            "grand_diff": (
                extracted_grand - grand_total
                if extracted_grand is not None and grand_total is not None else None
            ),
            "grand_match": (
                abs(extracted_grand - grand_total) <= 1
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
                abs(extracted_final - final_total) <= 1
                if extracted_final is not None and final_total is not None else None
            ),
        }
        return result

    results: List[Dict] = []

    if entry.get("multi_page"):
        pages = extract_pages_from_pdf(pdf_path, **pdfplumber_kwargs)
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
        for page_text in pages:
            # 大部分供應商的『未稅總額』就印在發票抬頭頁本身 (detect_fn 認得
            # 的那一頁)；但少數供應商 (例如 DSV) 是「一張發票的資料跨好幾頁」
            # ，未稅總額印在後面的附屬頁 (例如 Page 2 of 2 的條款/簽收頁)，
            # 那一頁不會被 detect_fn 認出來，所以這裡不限定只看 detect_fn
            # 認得的頁面，而是每一頁都試著抓抓看未稅總額；抓到之後，再用
            # 『這一頁的文字裡有沒有出現某張發票的發票號碼』來判斷這個未稅
            # 總額屬於哪一張發票 (發票抬頭頁跟附屬頁通常都會重複印一次發票
            # 號碼，例如 DSV 附屬頁開頭就是 "INVOICE ID610413182")。
            page_upper = None
            if extract_total_fn:
                try:
                    invoice_total = extract_total_fn(page_text)
                except Exception:
                    invoice_total = None
                if invoice_total is not None:
                    page_upper = page_text.upper()
                    for inv_key, inv_raw in raw_invoice_no.items():
                        if inv_key in totals_found or not inv_raw:
                            continue
                        if str(inv_raw).upper() in page_upper:
                            totals_found[inv_key] = invoice_total
                            break
            if extract_grand_total_fn:
                try:
                    grand_total = extract_grand_total_fn(page_text)
                except Exception:
                    grand_total = None
                if grand_total is not None:
                    if page_upper is None:
                        page_upper = page_text.upper()
                    for inv_key, inv_raw in raw_invoice_no.items():
                        if inv_key in grand_totals_found or not inv_raw:
                            continue
                        if str(inv_raw).upper() in page_upper:
                            grand_totals_found[inv_key] = grand_total
                            break
            if extract_final_total_fn:
                try:
                    final_total = extract_final_total_fn(page_text)
                except Exception:
                    final_total = None
                if final_total is not None:
                    if page_upper is None:
                        page_upper = page_text.upper()
                    for inv_key, inv_raw in raw_invoice_no.items():
                        if inv_key in final_totals_found or not inv_raw:
                            continue
                        if str(inv_raw).upper() in page_upper:
                            final_totals_found[inv_key] = final_total
                            break

        for inv_key in set(totals_found) | set(grand_totals_found) | set(final_totals_found):
            group_rows = groups.get(inv_key) or []
            if not group_rows:
                continue
            results.append(_build_result(
                group_rows, totals_found.get(inv_key), grand_totals_found.get(inv_key),
                final_totals_found.get(inv_key),
            ))
    else:
        text = extract_text_from_pdf(pdf_path, **pdfplumber_kwargs)
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
        if (invoice_total is not None or grand_total is not None or final_total is not None) and rows:
            results.append(_build_result(rows, invoice_total, grand_total, final_total))
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
        if is_multi:
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
                parsed = SUPPLIER_REGISTRY[key]["parse"](text)
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
