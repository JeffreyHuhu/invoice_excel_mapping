# -*- coding: utf-8 -*-
"""
suppliers_vietnam_streamlit.py
===============================
越南供應商 (VN01~VN20) 的 detect_xxx()/parse_xxx() 與 Excel 帳單
sheet_parse()/sheet_totals()，全部在這裡用 register_supplier() 註冊到
extract_utils_streamlit.SUPPLIER_REGISTRY。

新增越南供應商的作法跟以前一樣：
  1. 仿照下面任一家供應商的寫法，寫 detect_xxx()/parse_xxx()
     (PDF 帳單) 或 sheet_parse()/sheet_totals() (Excel 帳單)。
  2. 檔案最後呼叫一次 register_supplier("XXX_KEY", "顯示名稱", ...)。
  3. detect() 的判斷順序很重要：越有可能誤判的規則要放越後面。

印尼供應商在 suppliers_indonesia_streamlit.py；共用小工具在
suppliers_common_streamlit.py。兩支檔案都由 suppliers_streamlit.py 統一
import 觸發註冊，其他程式 (app.py/verify_app.py/run_regression.py 等) 不用改。
"""

import calendar
import collections
import difflib
import math
import os
import re
import datetime
from typing import Dict, List, Optional, Tuple

from extract_utils_streamlit import (
    _ocr_page_with_psm, _scanned_page_indexes, extract_pages_from_pdf,
    limited_pages, register_supplier, set_supplier_number, SUPPLIER_REGISTRY, SheetBook,
)
from suppliers_common_streamlit import (
    INDO_MONTHS, _search, _fmt_volume, _volume_from_text,
    _format_date_en, _format_date_en_from_month_name,
    _MONTH_ABBR_EN, _MONTH_NAME_TO_NUM, _VOLUME_BOX_RE,
    _EN_MONTHS_ABBR, _is_dhl_vnpt, _dsv_parse_date,
    _vn_text, _vn_num, _vn_amount, _vn_cell,
)

# ===========================================================================
# 越南 Excel 帳單共用工具 (VN03 RỒNG Á CHÂU、VN05 CONG THANH)
# ===========================================================================
# 越南供應商用 Excel 對帳表 (Bảng kê / Statement) 請款，不是 PDF。核心引擎
# 只交出「顯示中」的工作表 (read_spreadsheet())，這裡直接讀儲存格，不從文字
# 切欄位。金額欄位是數字型態，不用處理千分位。


def _vn_label(value) -> str:
    """表頭文字比對用：轉大寫、換行與連續空白壓成一個空白。"""
    return _vn_text(value).upper()


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


_VN_CONTAINER_RE = re.compile(r"^[A-Z]{4}\d{7}$")


def _vn_container(value) -> Optional[str]:
    """「CONT / XE」欄常混著車牌 (51C.91515) 或簡寫櫃號 (CTTU642)，只收完整的
    貨櫃號碼 (4 碼英文 + 7 碼數字)。一格裡換行塞好幾個櫃號時 (使用者 2026-11-09
    VN03 答案)，各自驗證後用 ", " 連接。"""
    codes = []
    for part in re.split(r"[\n,]+", "" if value is None else str(value)):
        code = _vn_text(part).upper().replace(" ", "")
        if _VN_CONTAINER_RE.match(code):
            codes.append(code)
    return ", ".join(codes) if codes else None


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
            # 材積統一寫法 (見 _fmt_volume)：CBM 欄 -> "0.26 CBM"
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

# SUPPLIER 一律用英文全名 (使用者 2026-11-09 規定)
_RAC_SUPPLIER = "Asia Dragon Forwarding Trading Company Limited"
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
    # "POUPHONG\…")：全名使用者 2026-10-28 確認。帳單上確實是 POU PHONG 時算成功辨識，
    # 不計入錯誤 (使用者 2026-11-10 規定，取消先前的不計入/中 規則)
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
    "TOLL",  # NHAP FCL DOT 1 的過路費 (使用者 2026-11-09 答案)
}
# SERVICE ALL 存在時，它的組成欄不另外轉 (避免重複計算)
_RAC_SERVICE_ALL_PARTS = {"SERVICE FEE", "EXTRA CUSTOM SHEET"}
_RAC_VAT_LABELS = {"VAT PASS THROUGH"}
# 小計/合計欄，不轉成費用也不列入勾稽
_RAC_SUBTOTAL_LABELS = {"TỔNG CỘNG", "TOTAL", "CLEARING EXPENSE PER P/L", "CLEARING EXPENSE"}
# 描述欄位：表頭 -> 代號 (同名欄位以第一次出現的為準，左右表內容相同)
_RAC_META_LABELS = {
    "INVOICE": "invoice_no", "SO INVOICE": "invoice_no", "INVOICE NO": "invoice_no",
    "SỐ BILL": "bl_no", "AWB,B/L NO": "bl_no", "BILL": "bl_no",
    "DECLARATION": "mbl_no", "SO TO KHAI": "mbl_no", "EXTRA CUSTOMS": "mbl_no",
    "KGS OF BL": "kgs_of_bl", "CONT": "container_no",
    "SL CONT": "sl_cont", "CONT 20/40": "cont_size",
    "DATE C/D": "date_cd", "DELIVERY DATE FCL": "delivery_date", "DATE": "invoice_date_col",
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
    charges, vat_cols, vat_cols_left, components = [], [], [], set()
    for j, label in enumerate(labels):
        if not label:
            continue
        field = _RAC_META_LABELS.get(label)
        if field and field not in meta:
            meta[field] = j
        if j < block_start:
            # 左表 (整張 B/L) 的 VAT PASS THROUGH 留著備援：少數列 (例如 B/L 是
            # "LOCAL" 的列) 右邊「依 KGS 分攤後」的 VAT 欄沒有算，要退回左表的值
            # (使用者 2026-11-09 答案)
            if label in _RAC_VAT_LABELS:
                vat_cols_left.append(j)
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
            # 欄名去掉 "PER P/L" (使用者 2026-11-09 答案，費用名稱不要帶這個技術性
            # 標記)；"CFS,THC,CIC PER P/L" 左表的乾淨欄名是 "CFS,THC,CIC FEE"
            # (多一個 "FEE")，這裡補上，其他欄 (STORAGE、D/O, H/L FEE…) 單純去掉
            # 字尾就是乾淨欄名，不用再處理
            charge = re.sub(r"\s*PER\s*P/L\s*$", "", original, flags=re.I).strip() or original
            if charge.upper() == "CFS,THC,CIC":
                charge += " FEE"
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
            "vat_cols_left": vat_cols_left, "billing": billing, "sub_header": is_sub, "row_total": row_total}


def _rac_volume(row: list, meta: Dict[str, int], kind: str) -> Optional[str]:
    """材積統一寫法：FCL 櫃量寫成「1X20」「2X40」(多櫃用 ", " 連接，
    使用者 2026-11-10 統一改為 1X40 格式，同 APL_VN)；LCL 沒有櫃：CBM 欄，沒有 CBM 才寫重量。"""
    # "SL CONT"（數量）x "CONT 20/40"（櫃型，例如 40）兩欄合併 (CÔNG NỢ POU PHONG
    # NHAP FCL 版面，使用者 2026-11-09 答案)
    sl_cont = _vn_num(_vn_cell(row, meta.get("sl_cont")))
    cont_size = _vn_num(_vn_cell(row, meta.get("cont_size")))
    if sl_cont and cont_size:
        return f"{int(sl_cont)}X{int(cont_size)}"
    containers = []
    for field, size in (("cont_20", "20"), ("cont_40", "40")):
        n = _vn_num(_vn_cell(row, meta.get(field)))
        if n:
            containers.append((n, size))
    if containers:
        return ", ".join(f"{int(n)}X{size}" for n, size in containers)
    if kind == "fcl_import":
        return None
    kgs = _vn_num(_vn_cell(row, meta.get("kgs")))
    if kind == "lcl_import":
        # LCL+LOCAL - IMPORT 一律用 KGS，不管 CBM 欄有沒有值 (使用者 2026-11-09 答案)
        return _fmt_volume(kg=_vn_text(kgs) if kgs else None)
    cbm = _vn_num(_vn_cell(row, meta.get("cbm")))
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
    last_file_no, last_file_box = None, {"container_no": None, "volume": None}
    last_bl_no, last_kgs_of_bl = None, None
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
        container_no = _vn_container(_vn_cell(row, meta.get("container_no")))
        volume = _rac_volume(row, meta, kind)
        if kind == "fcl_import":
            # 同一票 (File No，A 欄) 有多張 B/L 時，櫃數/櫃號只寫在第 1 列，要往下
            # 填到同一票的其他列 (使用者 2026-11-09 答案)
            file_no = _vn_text(_vn_cell(row, 0))
            if file_no != last_file_no:
                last_file_no, last_file_box = file_no, {"container_no": None, "volume": None}
            container_no = container_no or last_file_box["container_no"]
            volume = volume or last_file_box["volume"]
            last_file_box = {"container_no": container_no, "volume": volume}
        bl_no = _vn_text(_vn_cell(row, meta.get("bl_no"))) or None
        if kind == "lcl_import" and meta.get("kgs_of_bl") is not None:
            # Bill 空白、KGS OF BL 跟前一票相同 (同一張 B/L 分攤成好幾票)：B/L NO.
            # 用前一票的 Bill (使用者 2026-11-09 答案)
            kgs_of_bl = _vn_num(_vn_cell(row, meta["kgs_of_bl"]))
            if not bl_no and kgs_of_bl is not None and kgs_of_bl == last_kgs_of_bl:
                bl_no = last_bl_no
            last_bl_no, last_kgs_of_bl = bl_no, kgs_of_bl
        shared = {
            "invoice_no": invoice_no,
            "bl_no": bl_no,
            "mbl_no": _vn_text(_vn_cell(row, meta.get("mbl_no"))) or None,
            "container_no": container_no,
            "volume": volume,
        }
        # 有獨立的 "DATE" 欄 (新格式) 才當 INVOICE DATE；沒有這欄的版面 (例如
        # FCL - IMPORT) 照舊維持 N/A (使用者 2026-11-09 答案)
        if meta.get("invoice_date_col") is not None:
            shared["invoice_date"] = _vn_date(_vn_cell(row, meta["invoice_date_col"])) or "N/A"
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
        net = 0.0
        for j, charge in table["charges"]:
            amount = _vn_num(_vn_cell(row, j))
            if amount:
                items.append({"description": charge, "amount": amount, "vat": None, **shared})
                net += amount
        vat = sum(_vn_num(_vn_cell(row, j)) or 0 for j in table["vat_cols"]) or \
            sum(_vn_num(_vn_cell(row, j)) or 0 for j in table["vat_cols_left"])
        # Description 固定「VAT PASS THROUGH (VAT x%)」(使用者 2026-11-09 規定，取代舊的
        # 「VAT」/「VAT PASS THROUGH」沒有稅率寫法)；稅率 = VAT ÷ 費用加總，四捨五入到整數
        vat_desc = f"VAT PASS THROUGH (VAT {round(vat / net * 100)}%)" if net and vat else "VAT PASS THROUGH"
        if kind != "export":
            items.append({"description": vat_desc, "amount": None, "vat": vat or None, **shared})
        elif vat:
            # 出口表也有 VAT PASS THROUGH 欄時 (2026-10-28 抽樣 POU PHONG "XKTC" 工作表)，
            # 有金額才列；原本沒有 VAT 欄的出口表不受影響
            items.append({"description": vat_desc, "amount": None, "vat": vat, **shared})
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
                  "bl_no": r["fcr"], "mbl_no": r["swb"], "volume": f"{r['cbm']} CBM"}
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
def _dsv_vn_grid_lines(page, year: Optional[str]) -> str:
    """year：這個版面沒有任何欄位寫年份 (DATE 只有「27-Mar」)，從檔名補 (使用者
    2026-11-09 抽查 "...BY WEEKLY" 低等級批次；_cargowise_pages() 從檔名抓年份)。
    用額外一欄 "YEAR" 夾進每一列，_dsv_vn_grid_rows() 當備援年份來源。"""
    out = ""
    for table in page.extract_tables() or []:
        # "SOA" 是舊字型編碼 ("SOÁ SWB")；新樣本的表頭是乾淨的 "Số SWB"，用 "SWB"
        # 本身比對兩種編碼都抓得到
        if not table or not any("SWB" in (c or "") for c in table[0]):
            continue
        cell = lambda c: re.sub(r"\s+", " ", c or "").strip()  # noqa: E731
        header = [cell(c) for c in table[0]] + (["YEAR"] if year else [])
        out += "\n##DSVHDR## " + " | ".join(header)
        for row in table[1:]:
            cells = [cell(c) for c in row] + ([year] if year else [])
            out += "\n##DSVROW## " + " | ".join(cells)
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
        # SOA #/SO# 通常是空的，這個版面的年份要靠檔名 (YEAR 欄，_dsv_vn_grid_lines() 補的)
        soa = re.search(r"(\d{4})$", cells.get("SOA #", "") or cells.get("SO#", ""))
        year = soa.group(1) if soa else cells.get("YEAR")
        day = re.fullmatch(r"(\d{1,2})-([A-Za-z]{3})", cells.get("DATE", ""))
        rows.append({
            "swb": swb, "fcr": cells.get("FCR") or None, "po": cells.get("PO#") or None,
            "date": _format_date_en_from_month_name(day.group(1), day.group(2), year)
            if day and year else None,
            # CBM 偶爾被表格線切成兩段 ("1 .609")，去掉中間的空白接回去
            "cbm": re.sub(r"(?<=\d)\s+(?=[.,]?\d)", "", cells.get("CBM") or "") or None,
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
                  "bl_no": r["fcr"], "mbl_no": r["swb"], "volume": f"{r['cbm']} CBM" if r["cbm"] else None}
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


def _dsv_vn_sum_by_swb(rows: List[Dict], field: str) -> Dict[str, float]:
    """同一個 SWB 在「每週合輯」可能分兩週出現 (同票貨分兩次收費，日期不同)，
    明細 (_group_rows_by_invoice) 會把兩次的費用列合併成一組，這裡的總額也要
    用加總、不能互相覆蓋，才會跟合併後的明細加總一致 (使用者 2026-11-09 抽查
    DSV_VN「BY WEEKLY」批次，JUL.2026/OCT.2025/SEP.2025 各有一組 SWB 這樣分兩週)。"""
    totals: Dict[str, float] = {}
    for r in rows:
        value = r.get(field)
        if value is None:
            continue
        totals[r["swb"]] = totals.get(r["swb"], 0) + value
    return totals


def _dsv_vn_extract_total(text: str):
    """未稅總額：對帳表是每票 Thành tiền trước thuế；PRO FORMA 是每張 SUBTOTAL。"""
    if "##DSVROW## " in text:
        return _dsv_vn_sum_by_swb(_dsv_vn_grid_rows(text), "subtotal")
    if _is_dsv_vn_einvoice(text):
        return _dsv_vn_dot_amount(_search(r"Cộng doanh thu trên hóa đơn \(Amount\)\s*:\s*([\d.]+)", text))
    rows = _dsv_vn_table_rows(text)
    return {r["swb"]: r["subtotal"] for r in rows} if rows else _cw_extract_total(text)


def _dsv_vn_extract_grand_total(text: str):
    if "##DSVROW## " in text:
        return _dsv_vn_sum_by_swb(_dsv_vn_grid_rows(text), "total")
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
    # DSV "BY WEEKLY" 版面的年份要從檔名補 (見 _dsv_vn_grid_lines())
    dsv_year_m = re.search(r"(?<!\d)(20\d{2})(?!\d)", os.path.basename(pdf_path))
    dsv_year = dsv_year_m.group(1) if dsv_year_m else None
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
            # "SOA #" 是舊字型編碼的 "SO#"；新樣本是乾淨編碼，只用 "SWB" 判斷就夠
            # (使用者 2026-11-09 抽查 "...BY WEEKLY" 低等級批次)
            if "SWB" in text:
                text += _dsv_vn_grid_lines(clean, dsv_year)
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
    # OCR 頁的原因編號可能讀成 "*" 或 "1." 後面多一個空白 ("0% * 46.695.655")；
    # 原因編號也可能整個不印，"0%" 後面直接接金額 ("0% 74.815.872"，2026-11-10
    # 抽查 DHL_VN BY SHIPMENT 批次 S2501377271)
    r"|0%\s+(?:(?:\d\.\s?|\*\s))?(?=\d{1,3}(?:\.\d{3})*(?:,\d{1,2})?$))\s*"
    r"(?P<amount>\d{1,3}(?:\.\d{3})*(?:,\d{1,2})?)$", re.MULTILINE)
# 費用名稱裡夾著外幣單價跟匯率 ("Org. Terminal Handling-Forwarder USD 108,96 @"、
# "Freight USD 2.911,35 @ 25698,000000"：匯率可能接在同一行的 "@" 後面，也可能印在
# 下一行，兩種都被 _CW_CHARGE_RE 的 desc 群組整段吃進去)，答案要求外幣單價移到名稱
# 後面的括號、去掉 "@" 和匯率 (VN13 DHL 空運票，2026-11-10)
_CW_FX_DESC_RE = re.compile(r"^(?P<desc>.+?)\s+(?P<cur>[A-Z]{3})\s+(?P<amt>[\d.,]+)\s+@(?:\s+[\d.,]+)?$")


def _cw_clean_fx_desc(desc: str) -> str:
    m = _CW_FX_DESC_RE.match(desc)
    return f"{m.group('desc')} ({m.group('cur')} {m.group('amt')})" if m else desc


def _cw_clean_draft_watermark(desc: str) -> str:
    """"PRO FORMA" 頁常有一層 "DRAFT" 浮水印，字元位置固定，疊到真正文字上會插進一個
    字母 ("Security FDiling Fee" = "Filing"；金額裡插一個字母 "3.5D84,10" = "3.584,10")
    (VN13 DHL 答案，2026-11-10)。"""
    desc = desc.replace("FDiling", "Filing")
    return re.sub(r"(?<=[\d.])[A-Z](?=[\d,])", "", desc)


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
    """SHIPMENT 號碼；掃描頁 OCR 會在字母跟數字中間多讀一個 O ("SSGNO0315882")。
    掃描頁抬頭的 "SHIPMENT xxx" 有時被切成兩段 ("SHIPMENT S2602884 158"，中間
    多一個空白把號碼截短)：頁面上通常還有條碼的文字層 "^SHP=S2602884191;INV;|"
    沒有這個問題，優先用它 (使用者 2026-11-10 抽查 DHL_VN BY SHIPMENT 批次)。"""
    barcode = _search(r"\^SHP=([A-Z][A-Z0-9]{5,});INV", text)
    if barcode:
        return barcode
    inv = _search(_CW_INV_RE, text)
    if inv:
        fixed = re.sub(r"^([A-Z]+?)O+(?=\d{7,}$)", r"\1", inv)
        inv = fixed if fixed != inv and re.fullmatch(r"[A-Z]{3,5}\d{7,}", fixed) else inv
        return inv
    # 畫質太差的掃描頁，整行 "SHIPMENT SEDC1846196" 常被 OCR 讀成亂碼，連
    # SHIPMENT 這個字都認不出來：FLIGHT / DATE & REFERENCE 那一行的 HAWB 通常
    # 就是 SHIPMENT 號碼去掉開頭的 "S" (使用者 2026-11-11 回報 DSV_VN JUN.2026/
    # JUL.2026 混掃描檔的批次，SEDC1846196~99 這幾頁的抬頭整行讀壞)
    hawb = _search(r"FLIGHT / DATE[^\n]*\n[^\n]*?\s([A-Z]{2,5}\d{6,})\s*$", text)
    return f"S{hawb}" if hawb else None


def _cw_docs(text: str, page_nums: Optional[List[int]] = None):
    """一份 PDF 有好幾張發票時切開：每頁抬頭都有 SHIPMENT 號碼，連續幾頁號碼
    相同 (或續頁讀不到號碼) 就是同一張；"Page 1 of N" 一定是新的一張 (同一個
    SHIPMENT 可能開兩張，例如 YUSEN SSGN0316461 另有一張貨櫃延滯費)。沒有號碼
    也沒有頁碼的頁 (疊印的附件頁) 不算。沒有分頁記號 (非 _cargowise_pages
    擷取) 時整份當一張。先修 OCR 的數字空白 ("8%=64 . 000"、"1100000 ,00")。

    page_nums (可選)：跟 text 分頁記號一一對應的原始 PDF 頁碼 (0 起算)，給
    cargowise_invoice_pages() 用，算出每張發票對應到原始 PDF 的頁碼範圍 (本機
    轉檔工具拆檔用)。沒有傳就跟以前一樣只回傳文字 list，不影響既有行為；有傳
    時回傳 (文字 list, 對應頁碼 list)。"""
    text = re.sub(r"(?<=\d) *([.,]) *(?=\d)", r"\1", text)
    raw = text.split(_CW_PAGE_MARK)
    if page_nums is not None:
        pairs = [(p, n) for p, n in zip(raw, page_nums) if p.strip()]
        pages = [p for p, _ in pairs]
        nums: List[Optional[int]] = [n for _, n in pairs]
    else:
        pages = [p for p in raw if p.strip()]
        nums = [None] * len(pages)
    if len(pages) <= 1:
        single = [text.replace(_CW_PAGE_MARK, "")]
        return (single, [nums]) if page_nums is not None else single
    docs: List[List[str]] = []
    doc_pages: List[List[Optional[int]]] = []
    doc_inv: List[Optional[str]] = []
    current = None
    for page, num in zip(pages, nums):
        inv = _cw_invoice_no(page)
        # 畫質太差的掃描頁常連 "Page 1 of N" 都讀不出來，但 "SHIPMENT DETAILS
        # PRINTED BY" / "GOODS DESCRIPTION" 這兩個區塊只會出現在第一頁 (總額頁
        # 沒有)，可以當作沒讀到頁碼時的備援判斷 (使用者 2026-11-11 回報 DSV_VN
        # JUN.2026/JUL.2026 混掃描檔批次：同一個 SHIPMENT 開兩張發票，沒有這個
        # 判斷，第二張會被誤併進第一張)
        first_page = bool(re.search(r"\bPage 1 of \d+", page)) or (
            not re.search(r"\bPage \d+ of \d+", page)
            and ("SHIPMENT DETAILS PRINTED BY" in page or "GOODS DESCRIPTION" in page))
        if inv is None and not re.search(r"\bPage \d+ of \d+", page):
            continue
        # "Page 2 of N" 以後一定是上一張的續頁 (掃描頁 OCR 號碼可能讀錯，不能只靠號碼)
        later_page = re.search(r"\bPage (?:[2-9]|\d{2,}) of \d+", page)
        # 跳頁：續頁的 SHIPMENT 號碼跟目前這一張不同，但前面已經開過同號碼的一張
        # (中間夾著別的發票)，要接回那一張，不能誤接到目前這一張 (使用者 2026-11-11
        # 回報 DSV_VN APR.2026 SSGN0439661 第 1 頁在第 113 頁，第 2 頁 (TOTAL) 被排到
        # 第 131 頁，中間夾了 17 頁其他發票；沒接回去的話 SSGN0439661 抓不到 TOTAL，
        # 第 131 頁又被誤併進第 131 頁前一張 SSGN0441871)
        jump_target = (next((i for i in range(len(docs) - 1, -1, -1) if doc_inv[i] == inv), None)
                       if inv and inv != current and not first_page else None)
        if jump_target is not None:
            docs[jump_target].append(page)
            doc_pages[jump_target].append(num)
        elif docs and not first_page and (inv is None or inv == current or later_page):
            docs[-1].append(page)
            doc_pages[-1].append(num)
        else:
            docs.append([page])
            doc_pages.append([num])
            doc_inv.append(inv)
            current = inv
    # 同一張單據掃描兩次 (號碼與 TOTAL 都相同) 只留一張 (使用者 2026-10-28：DHL 合輯第 19、20 頁
    # S2601526262)。同號碼但總額不同的是另一張發票 (YUSEN SSGN0316461)，照常保留
    out, out_pages, seen = [], [], set()
    for d, nums_d in zip(("\n".join(d) for d in docs), doc_pages):
        total = re.findall(r"(?m)^[\s|\"'`]*TOTAL\b[^\n]*?(\d[\d.,]*\d)\s*$", d)
        key = (_cw_invoice_no(d), total[-1] if total else None)
        if key[0] and key[1] and key in seen:
            continue
        seen.add(key)
        out.append(d)
        out_pages.append(nums_d)
    if not out:
        out, out_pages = [text], [nums]
    return (out, out_pages) if page_nums is not None else out


def cargowise_invoice_pages(pdf_path: str, key: str) -> Dict[str, List[int]]:
    """DSV_VN/MOL/DHL_VN/YUSEN 專用：回傳這份 PDF 每張發票 (SHIPMENT 號碼)
    對應到原始 PDF 的頁碼 (1 起算)。跟 invoice_text()/_cw_docs() 用同一套切法，
    只是多記頁碼，給本機轉檔工具 (pdf_to_excel.exe) 判斷能不能把信心等級低的
    發票拆成另一份 PDF，不影響原本的擷取結果。算不出頁碼對應時回傳空 dict
    (呼叫端退回整份一起歸檔的舊行為)。"""
    from extract_utils_streamlit import supplier_pages
    entry = SUPPLIER_REGISTRY.get(key)
    if not entry:
        return {}
    pages = supplier_pages(pdf_path, key)
    check = entry.get("invoice_page")
    keep = [i for i, p in enumerate(pages) if check(p)] if check else list(range(len(pages)))
    if not keep:
        keep = list(range(len(pages)))
    text = "\n".join(pages[i] for i in keep)
    docs, doc_pages = _cw_docs(text, page_nums=keep)
    result: Dict[str, List[int]] = {}
    for doc_text, nums in zip(docs, doc_pages):
        inv = _cw_invoice_no(doc_text)
        if not inv:
            continue
        bucket = result.setdefault(inv, [])
        for n in nums:
            if n is not None and (n + 1) not in bucket:
                bucket.append(n + 1)  # 1 起算，給 pypdf 拆檔用
    for inv in result:
        result[inv].sort()
    return result


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


_CW_CHARGE_SHEET_FEE_RE = re.compile(
    r"^(?P<desc>\S.+?)\s+(?P<charge>\d{1,3}(?:\.\d{3})*)\s+"
    r"(?P<tax>\d{1,3}(?:\.\d{3})*)\s+(?P<amount>\d{1,3}(?:\.\d{3})*)$")


def _cw_charge_sheet_items(doc: str) -> List[Dict]:
    """DSV 越南「Charge Sheet」版面 (BY SHIPMENT 合輯裡偶爾夾著的另一種版型，使用者
    2026-11-09 抽查 BY SHIPMENT 批次，FEB.2026/JAN.2025 各有一張)。費用行第一行是
    「名稱片段 CHARGE(VND) TAX(VND) AMOUNT(VND)」，名稱剩下的部分 (單價/單位說明)
    折到接下來幾行，直到下一筆費用或 "Job Total" 才結束 (Job Total 是整張小計，
    不當品項)。"""
    if "CHARGE DESCRIPTION" not in doc or "Job Total" not in doc:
        return []
    items: List[Dict] = []
    pending: Optional[Dict] = None
    for line in doc.splitlines():
        m = _CW_CHARGE_SHEET_FEE_RE.match(line.strip())
        if m and m.group("desc") == "Job Total":
            break
        if m:
            if pending:
                items.append(pending)
            pending = {"description": m.group("desc"),
                       "amount": _dsv_vn_dot_amount(m.group("charge")),
                       "vat": _dsv_vn_dot_amount(m.group("tax"))}
        elif pending and line.strip():
            pending["description"] += " " + line.strip()
    if pending:
        items.append(pending)
    for it in items:
        it["description"] = re.sub(r"\s+", " ", it["description"]).strip()
    return items


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
    # OCEAN BILL OF LADING (MBL) 欄可能空白，只剩 IMO 和 HOUSE BILL OF LADING 兩個值
    # (VN13 DHL 答案：MBL 空白填 N/A，2026-11-10)；前後不能同時要求 \s+ (CLAUDE.md
    # Regex 陷阱)，ocean 群組整個選填
    m = _cw_after(doc, r"OCEAN BILL OF LADING.*HOUSE BILL OF LADING",
                  r"(?P<v>.+?) / (?P<voy>\S+) / \S+(?:\s+(?P<ocean>\S+))?\s+(?P<house>\S+)$")
    if m:
        h["vessel"] = f"{m.group('v')} / {m.group('voy')}"
        h["mbl_no"], h["bl_no"] = m.group("ocean") or "N/A", m.group("house")
    else:
        m = _cw_after(doc, r"FLIGHT / DATE & REFERENCE\s+MAWB\s+HAWB",
                      r"(?P<flight>\S+) / .*?(?P<mawb>\S+) (?P<hawb>\S+)$")
        if m:
            h["vessel"], h["mbl_no"], h["bl_no"] = m.group("flight", "mawb", "hawb")
        else:
            # DHL 版面的表頭沒有 "& REFERENCE"、也沒有中間那個單號，船名多接日期
            # (VN13 答案，2026-11-10)
            m = _cw_after(doc, r"FLIGHT / DATE\s+MAWB\s+HAWB",
                          r"(?P<flight>\S+) / (?P<date>\S+) (?P<mawb>\S+) (?P<hawb>\S+)$")
            if m:
                h["vessel"] = f"{m.group('flight')} / {m.group('date')} (Flight)"
                h["mbl_no"], h["bl_no"] = m.group("mawb"), m.group("hawb")
    m = _cw_after(doc, r"^ORIGIN\b",
                  rf"(?P<pol>.+?)\s+(?P<etd>{_CW_DATE})\s+(?P<pod>.+?)(?:\s+(?P<eta>{_CW_DATE}))?$")
    h["port_of_loading"] = m.group("pol").strip(" ,") if m else None
    h["port_of_discharge"] = m.group("pod").strip(" ,") if m else None
    h["onboard_date"] = _cw_date(m.group("etd")) if m else None
    h["arrive_date"] = _cw_date(m.group("eta")) if m else None
    sec = re.search(r"^CONTAINERS\s*\n(.*?)(?:\n\s*(?:Total CO|CO₂|CHARGES)|\Z)",
                    doc, re.DOTALL | re.MULTILINE)
    # 櫃號不一定在這一行開頭 (VN13 DHL："- 40HC, OOCU6556180 - 40HC"，前面多一個沒有
    # 櫃號的櫃型，2026-11-10)：改成整段找「櫃號 - 櫃型」，不要求從行首開始，其他
    # 供應商既有答案要保留「櫃號 - 櫃型」(例如 "FFAU3490244 - 40HC") 不受影響
    conts = re.findall(r"[A-Z]{4}\d{7}\s*-\s*\w+", sec.group(1)) if sec else []
    h["container_no"] = ", ".join(conts) or None

    cur = _search(r"CHARGES IN ([A-Z]{3})", doc)
    suffix = f" ({cur})" if cur and cur != "VND" else ""      # 外幣發票的費用名稱加幣別
    items = [{"description": _cw_clean_fx_desc(_cw_clean_draft_watermark(im.group("desc").strip())) + suffix,
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
        items = _cw_charge_sheet_items(doc)
    if not items:
        items = [{"description": None, "amount": None, "vat": None}]
    return [{**h, **it} for it in items]


def _cw_parse(text: str, supplier: Optional[str]) -> Dict:
    items = [row for doc in _cw_docs(text) for row in _cw_parse_doc(doc, supplier)]
    return {"header": {}, "items": items}


def _cw_totals(text: str, patterns, conv=_cw_vn_number, key_fn=_cw_invoice_no):
    """每張發票的總額 {SHIPMENT 號碼: 金額}。patterns 依序試，第一個抓得到的
    為準 (同一個 pattern 抓到好幾個就加總，例如 MOL 每個幣別一行 GRAND TOTAL)。
    key_fn：發票號碼不是 SHIPMENT 號碼時可以換一個 (VN13 DHL 的「INVOICE」頁用
    F 號碼，2026-11-10)。"""
    out, missing = {}, set()
    for doc in _cw_docs(text):
        inv = key_fn(doc)
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
    invoice_page_map=lambda p: cargowise_invoice_pages(p, "DSV_VN"),
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


_DHL_VN_SUPPLIER = "DHL GLOBAL FORWARDING (VIETNAM) CORPORATION"


def _dhl_vn_doc_key(doc: str) -> Optional[str]:
    """一般 "INVOICE" 頁 (不是 PRO FORMA) 的發票號碼是 F 號碼，不是 SHIPMENT 號碼
    (VN13 答案，2026-11-10)。"""
    return _search(r"\bINVOICE\s+(F\d+)\b", doc) or _cw_invoice_no(doc)


def _is_dhl_vn_vat_invoice(page: str) -> bool:
    return "HÓA ĐƠN GIÁ TRỊ GIA TĂNG" in page.upper()


def _dhl_vn_vat_docs(text: str) -> List[str]:
    """VAT INVOICE 頁可能跨頁：第 1 頁結尾 "Trang 1/2"，續頁開頭
    "tiếp theo trang trước"，Tổng cộng 在最後一頁 (VN13 答案，2026-11-10)。"""
    merged: List[str] = []
    for page in text.split(_CW_PAGE_MARK):
        if not _is_dhl_vn_vat_invoice(page):
            continue
        if merged and re.search(r"ti[eế]p theo trang tr[uướ]+c", page, re.IGNORECASE):
            merged[-1] += "\n" + page
        else:
            merged.append(page)
    return merged


def _dhl_vn_cw_text(text: str) -> str:
    """拿掉 VAT INVOICE 頁，剩下的頁交給共用的 CargoWise 解析 (PRO FORMA / 一般
    INVOICE)。"""
    return _CW_PAGE_MARK.join(p for p in text.split(_CW_PAGE_MARK) if not _is_dhl_vn_vat_invoice(p))


# VAT INVOICE 費用列："STT 號 Số vận đơn 品名 Tỷ giá 數量 單價 未稅 稅率% 稅額 含稅"；
# Số vận đơn、品名都可能整欄空白 (前後不能同時要求 \s+，CLAUDE.md Regex 陷阱)
_DHL_VN_VATINV_ITEM_RE = re.compile(
    r"^(?P<stt>\d+)\s+(?:(?P<hbl>[A-Z]{2,4}\d{5,7})\s+)?(?P<desc>.*?)\s*"
    r"(?P<rate>[\d.,]+)\s+(?P<qty>[\d.,]+)\s+(?P<price>[\d.,]+)\s+(?P<net>[\d.,]+)\s+"
    r"(?P<pct>\d+)%\s+(?P<vat>[\d.,]+)\s+(?P<total>[\d.,]+)$", re.MULTILINE)
# 英文備註 (使用者同意先暫用，之後依統一費用對照表更新，2026-11-10)
_DHL_VN_EN_NOTE = {
    "PHI XEP DO HANG TAI CANG DI": "Origin Cargo Handling Fee",
    "PHI QUAN LY CHUNG TU TAI CANG DI": "Origin Documentation Management Fee",
    "PHI CHUNG TU GOC TAI CANG DI": "Origin Original Documents Fee",
    "PHI KHAI BAO HAI QUAN TAI CANG DI": "Origin Customs Declaration Fee",
    "PHI VAN CHUYEN BOC XEP TAI KHO TAI CANG DI": "Origin Warehouse Trucking & Handling Fee",
    "PHI DICH VU KHAI BAO TRONG LUONG TAI CANG DI": "Origin VGM Declaration Service Fee",
}


def _dhl_vn_vat_key(doc: str) -> Optional[str]:
    # _cargowise_pages 會把這頁兩欄的版面重排，"Số hóa đơn(Invoice No.):" 後面常常
    # 變空白，真正的號碼被排到上一欄「Ký hiệu(Serial No.): 1C25TSG」的下一行，自己
    # 獨立一行 (VN13 答案，2026-11-10)：優先認這個排列，抓不到才退回原本的標籤比對
    m = re.search(r"Ký hiệu\(Serial No\.\):\s*(\S+)\s*\n\s*(\d+)\b", doc)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    series = _search(r"Ký hiệu\(Serial No\.\):\s*(\S+)", doc)
    number = _search(r"Số hóa đơn\(Invoice No\.\):\s*(\S+)", doc)
    return f"{series}-{number}" if series and number else number


def _dhl_vn_vat_totals(doc: str):
    m = re.search(r"Tổng cộng\(Grand total\):\s*([\d.]+)\s+([\d.]+)\s+([\d.]+)", doc)
    return (_dsv_vn_dot_amount(m.group(1)), _dsv_vn_dot_amount(m.group(3))) if m else (None, None)


def _parse_dhl_vn_vat_invoice(doc: str) -> List[Dict]:
    h: Dict[str, Optional[str]] = {"supplier": _DHL_VN_SUPPLIER, "invoice_no": _dhl_vn_vat_key(doc)}
    m = re.search(r"Ngày hóa đơn\(Invoice date\):\s*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})", doc)
    h["invoice_date"] = _format_date_en(*m.groups()) if m else None
    h["consignee"] = _search(r"Người mua\(Buyer Name\):\s*(.+)", doc)
    if h["consignee"]:
        h["consignee"] = _vn_full_consignee(h["consignee"])
    h["order_no"] = _search(r"Số hóa đơn nội bộ\(DGF Inv\.No\.\):\s*(\S+)", doc) or "N/A"
    items = []
    lines = doc.splitlines()
    for i, line in enumerate(lines):
        m = _DHL_VN_VATINV_ITEM_RE.match(line.strip())
        if not m:
            continue
        desc = m.group("desc").strip()
        if not desc:
            # 品名折成 3 行：費用列上方第 1 行 = 前半段，下方第 1 行 = 後半段
            # (VN13 答案，2026-11-10)
            prefix = lines[i - 1].strip() if i else ""
            suffix = lines[i + 1].strip() if i + 1 < len(lines) else ""
            desc = " ".join(x for x in (prefix, suffix) if x)
        en = _DHL_VN_EN_NOTE.get(_vt_norm(desc))
        items.append({"description": f"{desc} ({en})" if en else desc,
                      "amount": _dsv_vn_dot_amount(m.group("net")), "vat": _dsv_vn_dot_amount(m.group("vat")),
                      "bl_no": m.group("hbl") or "N/A", **h})
    return items or [{"description": None, "amount": None, "vat": None, **h}]


def parse_dhl_vn(text: str) -> Dict:
    if _is_dhl_vn_charge_sheet(text):
        return _parse_dhl_vn_charge_sheet(text)
    items = []
    for doc in _cw_docs(_dhl_vn_cw_text(text)):
        rows = _cw_parse_doc(doc, _DHL_VN_SUPPLIER)
        # 一般 "INVOICE" 頁用 F 號碼；材積照帳單原文寫 M3，不是 CBM (VN13 答案，2026-11-10)
        f_no = _search(r"\bINVOICE\s+(F\d+)\b", doc)
        for r in rows:
            if f_no:
                r["invoice_no"] = f_no
            if r.get("volume"):
                r["volume"] = r["volume"].replace(" CBM", " M3")
            if r.get("consignee"):
                r["consignee"] = _vn_full_consignee(r["consignee"])
        items += rows
    for doc in _dhl_vn_vat_docs(text):
        items += _parse_dhl_vn_vat_invoice(doc)
    return {"header": {}, "items": items}


def _dhl_vn_extract_total(text: str):
    if _is_dhl_vn_charge_sheet(text):
        return None
    out = _cw_totals(text, [r"SUBTOTAL[ \t]+(\d[\d.]*(?:,\d+)?)",
                            r"discrepancies\.[ \t]+(\d{1,3}(?:\.\d{3})+(?:,\d+)?)[ \t]*$"],
                     key_fn=_dhl_vn_doc_key)
    for doc in _dhl_vn_vat_docs(text):
        net, _grand = _dhl_vn_vat_totals(doc)
        key = _dhl_vn_vat_key(doc)
        if key and net is not None:
            out[key] = net
    return out


def _dhl_vn_extract_grand_total(text: str):
    if _is_dhl_vn_charge_sheet(text):
        v = _search(r"Total \(VND\)\s*([\d.]+,\d{2})", text)
        return _dhl_vn_num(v) if v else None
    out = (_cw_totals(text, [r"^TOTAL[ \t]+(?:VND|USD)[ \t]+(\d[\d.]*(?:,\d+)?)",
                             r"^(\d[\d.]*(?:,\d+)?)[ \t]*\nTOTAL[ \t]+(?:VND|USD)\s*$",
                             r"Invoiced[ \t]+(?:VND|USD)[ \t]+(\d{1,3}(?:\.\d{3})+(?:,\d+)?)\b",
                             # DHL 版面 "TOTAL VND" 常跟左欄的銀行轉帳說明擠在同一行，不是行首
                             # (VN13 答案，2026-11-10)
                             r"TOTAL[ \t]+(?:VND|USD)[ \t]+(\d[\d.]*(?:,\d+)?)\s*$"],
                      key_fn=_dhl_vn_doc_key)
           or _cw_totals(text, [r"GRAND TOTAL for [A-Z]{3}[ \t]+(\d[\d,]*(?:\.\d+)?)"],
                         conv=_cw_comma_number, key_fn=_dhl_vn_doc_key))
    for doc in _dhl_vn_vat_docs(text):
        _net, grand = _dhl_vn_vat_totals(doc)
        key = _dhl_vn_vat_key(doc)
        if key and grand is not None:
            out[key] = grand
    return out


register_supplier(
    "DHL_VN", _DHL_VN_SUPPLIER, detect_dhl_vn, parse_dhl_vn,
    extract_total=_dhl_vn_extract_total,
    extract_grand_total=_dhl_vn_extract_grand_total, page_text=_cargowise_pages,
    invoice_page_map=lambda p: cargowise_invoice_pages(p, "DHL_VN"),
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
    invoice_page_map=lambda p: cargowise_invoice_pages(p, "YUSEN"),
)


# VN09 MOL：Debit Note，公司名稱只在 logo 圖片裡 (文字層沒有)，固定填
# "MOL Logistics Vietnam Inc" (使用者 2026-10 正確答案)；House B/L 是 "MLG…"
def detect_mol(text: str) -> bool:
    return bool(re.match(r"\s*Debit Note - ", text)
                and (re.search(r"\bMLG[A-Z]{2}\d{6,8}\b", text) or "MOL LOGISTICS" in text.upper())) \
        or _is_molcs(text)


# --- MOL CONSOLIDATION SERVICE 越南增值稅發票 (2026-10-28 抽樣 "MOL LOGICTICS_SEP.2025" 等，63 頁) --
# 一頁一張 HÓA ĐƠN GIÁ TRỊ GIA TĂNG；這批 "Số (No)" 與日期都空白 (草稿) 的張數占多數 (119 張裡
# 117 張，使用者 2026-11-10 VN09 答案)，INVOICE NO 這種情況只填 Ký hiệu「1C25TMC」(不要用備註
# 裡的訂艙號碼)。明細 "1 PHÍ CHỨNG TỪ (FCR fee) Bill 1.00 1,000,000 1,000,000"，名稱太長時折到
# 上下兩行；整張一個 VAT 單獨一列。
_MOLCS_ITEM_RE = re.compile(
    r"^\d+\s+(?:(?P<desc>.+?)\s+)?(?P<unit>[A-Za-z]{2,6})\s+(?P<qty>[\d.,]+)\s+[\d.,]+\s+(?P<amount>[\d.,]+)$")
# 訂艙號碼 (Order No.)／括號內的號碼 (B/L NO.)，例如 "VN2085282 (620112501501)"；多筆用「/」分隔
# 時 Order No. 用「, 」連接 (使用者 2026-11-10 答案)。整行獨立，不要跟其他文字黏在一起判斷。
_MOLCS_REF_RE = re.compile(r"^((?:VN\d{7}/)*VN\d{7})(?:\s*\((\S+)\))?\s*$", re.MULTILINE)


def _is_molcs(text: str) -> bool:
    upper = text.upper()
    return "MOL CONSOLIDATION SERVICE" in upper and "HÓA ĐƠN GIÁ TRỊ GIA TĂNG" in upper


def _molcs_pages(text: str) -> List[str]:
    """一張發票可能跨 2 頁 (第 2 頁開頭印 "Tiep theo trang truoc")：從發票頁開始累積，
    遇到有 "Cộng tiền hàng" (Total before VAT) 的頁才收尾成一張完整發票的文字 (使用者
    2026-11-10 VN09 答案；之前只挑『當頁已經有 Cộng tiền hàng』會把跨頁發票的品項頁漏掉)。
    MOL 改成 multi_page=True 逐頁呼叫後，這支函式常常只拿到『一頁』：有合計就直接收尾；
    只有品項、沒有合計的那一頁 (例如跨頁發票的第 1 頁) 也要在迴圈結束時原樣吐出來，不要
    因為沒看到合計就整頁丟掉——另一頁 (只有合計、沒有品項) 會在各自的呼叫裡吐出 VAT 列，
    最後兩頁的列用同一個 invoice_no 合併 (見 _parse_multi_page_pdf())。"""
    out: List[str] = []
    buf: List[str] = []
    for p in text.split(_CW_PAGE_MARK):
        if not _is_molcs(p):
            continue
        buf.append(p)
        if "Cộng tiền hàng" in p:
            out.append("\n".join(buf))
            buf = []
    if buf:
        out.append("\n".join(buf))
    return out


def _molcs_num(value: Optional[str]):
    """'1,000,000'／'1.000.000'／'13.08'／'3,00' -> 數字。規則 (使用者 2026-11-10 答案)：
    最後一個 '.'/',' 後面只有 1~2 位數字時，那個符號是小數點，其餘的 '.'/',' 都是千分位。"""
    if not value:
        return None
    text = value.strip()
    m = re.search(r"[.,](\d{1,2})$", text)
    if m:
        intpart = re.sub(r"[.,]", "", text[:m.start()])
        try:
            num = float(f"{intpart or '0'}.{m.group(1)}")
        except ValueError:
            return None
    else:
        digits = re.sub(r"[.,]", "", text)
        if not digits.isdigit():
            return None
        num = float(digits)
    return int(num) if num == int(num) else num


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
        # Volume = 這一列自己的 Số lượng + ĐVT (每個品項各自的單位不同，不是整張發票共用一個值，
        # 使用者 2026-11-10 答案：「13.08 CBM」「1 SHPT」「3 Cont」)
        items.append({"description": desc, "amount": _molcs_num(m.group("amount")), "vat": None,
                      "volume": f"{_molcs_num(m.group('qty'))} {m.group('unit')}"})
        last = i + (0 if m.group("desc") else 1)
    # 有些月結檔的版型 (例如電子發票檢視器 "Bản thể hiện của hóa đơn điện tử") 欄位是
    # 「數值在前、英文標籤括號在下一行」(例如 "Ký hiệu 1C25TMC\n(Sign):")，跟一般版型
    # 「標籤在前、數值接在後面」(例如 "Ký hiệu (Sign): 1C25TMC") 相反；兩種都要認得
    # (使用者 2026-11-10 VN09 答案，同一份 PDF 兩種版型混用)。
    series = (_search(r"Ký hiệu \(Sign\):\s*(\S+)", page)
              or _search(r"Ký hiệu\s+(\S+)\s*\n\s*\(Sign\):", page))
    no = (_search(r"Số\s*\(No\):\s*(\d+)", page)
          or _search(r"Số\s+(\d+)\s*\n\s*\(No\):", page))
    d = (re.search(r"Ngày\s*\(Date\)\s*(\d{1,2})\s*tháng\s*\(month\)\s*(\d{1,2})\s*năm\s*\(year\)\s*(\d{4})", page)
         or re.search(r"Ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})\s*\n\s*\(Date\)", page))
    ref_m = _MOLCS_REF_RE.search(page)
    rate = _search(r"VAT rate\):\s*(\d+)\s*%", page)
    consignee_raw = (_search(r"Tên đơn vị \(Company's name\):\s*(.+)", page)
                      or _search(r"Tên đơn vị\s+(.+?)\s*\n\s*\(Company's name\):", page))
    common = {
        # Số 空白時只填 Ký hiệu (使用者 2026-11-10 答案，不要用訂艙號碼頂替)
        "invoice_no": f"{series}-{no}" if series and no else series,
        "invoice_date": _format_date_en(*d.groups()) if d else None,
        "supplier": "MOL CONSOLIDATION SERVICE (VIETNAM) CO., LTD.",   # 英文名 (使用者 2026-11-10 答案)
        "consignee": _vn_full_consignee(consignee_raw),
        "order_no": ", ".join(ref_m.group(1).split("/")) if ref_m else "N/A",
        "bl_no": (ref_m.group(2) if ref_m else None) or "N/A",
    }
    rows = [{**common, **it} for it in items]
    vat = (_molcs_num(_search(r"VAT amount\):\s*([\d.,]+)", page))
           or _molcs_num(_search(r"([\d.,]+)\s*\n\s*\(VAT amount\):", page)))
    if vat is not None:
        # 不要求這一頁本身要有品項列才加 VAT 列：MOL 現在逐頁辨識/擷取 (multi_page=True)，
        # 跨頁發票的合計常常印在「只有合計、沒有品項」的那一頁，兩頁各自的列最後靠
        # invoice_no 合併 (使用者 2026-11-10 答案；見 _parse_multi_page_pdf())。
        rows.append({**common, "description": f"VAT {rate}%" if rate else "VAT", "amount": None,
                     "vat": vat, "volume": "N/A", "order_no": "N/A", "bl_no": "N/A"})
    return rows


# --- VN09 MOL LOGISTICS 空運 Debit Note (SAEM11Y… 開頭；跟上面海運 CargoWise Debit Note
# (SSEM11Z… 開頭、走 _cw_parse()) 完全是兩種版型，不要共用規則；使用者 2026-11-10 答案) --------
# 費用名稱跟代碼 (OAWBFEE/WAMS/OXRAY/OTHC) 疊字 (兩層文字幾乎同位置，pdfplumber 逐字抽出時交錯
# 排列)，這份檔案目前只出現這 4 種，直接整句替換成乾淨文字；之後遇到新代碼再補進這個字典。
_MOL_DN_CHARGE_GARBLE = {
    "OPAhWí vBậFnE đEơn hàng không": "Phí vận đơn hàng không",
    "WPAhMí ASMS": "Phí AMS",
    "OPXhRí AsoYi chiếu": "Phí soi chiếu",
    "OPThHí CXếp Dỡ": "Phí Xếp Dỡ",
}
_MOL_DN_EN_NOTE = {
    "Phí vận đơn hàng không": "Air Waybill Fee",
    "Phí AMS": "AMS Filing Fee",
    "Phí soi chiếu": "X-ray Screening Fee",
    "Phí Xếp Dỡ": "Terminal Handling / Loading-Unloading Fee",
}
_MOL_DN_ITEM_RE = re.compile(
    r"^(?P<desc>.+?)\s+(?P<qty>[\d.]+)\s*-\s*(?P<unit>\S+)\s+(?P<rate>[\d.]+)\s+(?P<vatpct>\d+)\s*%\s+"
    r"(?P<cur>[A-Z]{3})\.?\s+(?P<pdfamt>[\d.]+)$", re.MULTILINE)
_MOL_DN_GROUP_RE = re.compile(
    r"^(POUYUEN VIETNAM COMPANY LIMITED|POU (?:LI|HUNG|SUNG|CHEN) VIETNAM COMPANY LIMITED)\b")
_MOL_DN_FLIGHT_RE = re.compile(r"^(\S+)\s*/\s*(\d{1,2}-[A-Za-z]{3})\s+(\d{5,})\s+(\S+)$", re.MULTILINE)


def _is_mol_air_dn(text: str) -> bool:
    return bool(re.match(r"\s*Debit Note - ", text)) and "FLIGHT / DATE" in text


def _mol_air_dn_pages(text: str) -> List[str]:
    """空運 Debit Note 每張都是 "Page 1 of 1" (使用者 2026-11-10 樣本 5 張都是)，不用像
    molcs 那樣跨頁累積，一頁就是一張完整發票。"""
    return [p for p in text.split(_CW_PAGE_MARK) if _is_mol_air_dn(p)]


def _mol_dn_date(text: Optional[str]) -> Optional[str]:
    """'02-Apr-25' (13:50 等時間丟棄) -> 標準日期字串。"""
    m = re.match(r"(\d{1,2})-([A-Za-z]{3})-(\d{2})", text or "")
    if not m:
        return None
    day, mon, yy = m.groups()
    return _format_date_en_from_month_name(day, mon, 2000 + int(yy))


def _parse_mol_air_dn(page: str) -> List[Dict]:
    for garbled, clean in _MOL_DN_CHARGE_GARBLE.items():
        page = page.replace(garbled, clean)
    lines = page.splitlines()
    h: Dict[str, Optional[str]] = {
        "invoice_no": _search(r"SHIPMENT(\S+)", page),
        # TOP20 英文名稱 (使用者指定)
        "supplier": "MOL LOGISTICS TRANSPORTATION VIETNAM INC",
        "invoice_date": _mol_dn_date(_search(r"DATE(\S+)", page)),
    }
    # CONSIGNEE = 帳單抬頭 ("公司名稱 SHIPMENT…" 兩欄黏在一起，取 SHIPMENT 前的部分)。用整段文字
    # 搜尋這一行，不要用固定行號——多頁文件合併時，第 1 頁以外的每一頁開頭都會多一個空白行
    # (invoice_text() 用 "\n".join() 接上 _cargowise_pages() 自帶的換行分頁記號)，行號會位移。
    consignee_m = re.search(r"^(.+?)\s*SHIPMENT\S", page, re.MULTILINE)
    h["consignee"] = _vn_full_consignee(consignee_m.group(1).strip()) if consignee_m else None
    h["order_no"] = _search(r"P\.O\.#(\S+)", page) or "N/A"

    flight_m = _MOL_DN_FLIGHT_RE.search(page)
    if flight_m:
        h["vessel"] = f"{flight_m.group(1)} / {flight_m.group(2)} (Flight)"
        h["mbl_no"], h["bl_no"] = flight_m.group(3), flight_m.group(4)

    vol_m = re.search(r"([\d.]+)\s*M3", page)
    h["volume"] = f"{vol_m.group(1)} M3" if vol_m else None

    # ORIGIN/ETD/DESTINATION/ETA 這一列有時因為目的港名稱太長換行 (例如 "...International Apt,"
    # 换行後接 "China")，先把標籤列到 "CHARGE DESCRIPTION" 之間的文字合併成一行再抓。
    label_idx = next((i for i, ln in enumerate(lines) if ln.startswith("ORIGIN") and "ETD" in ln), None)
    charge_idx = next((i for i, ln in enumerate(lines) if ln.startswith("CHARGE DESCRIPTION")), len(lines))
    if label_idx is not None:
        blob = " ".join(ln.strip() for ln in lines[label_idx + 1:charge_idx])
        m = re.match(r"^(.+?)\s+(\d{1,2}-[A-Za-z]{3}-\d{2})\s+(.+?)\s+(\d{1,2}-[A-Za-z]{3}-\d{2})$", blob)
        if m:
            h["port_of_loading"], h["onboard_date"] = m.group(1), _mol_dn_date(m.group(2))
            h["port_of_discharge"], h["arrive_date"] = m.group(3), _mol_dn_date(m.group(4))

    shipper_idx = next((i for i, ln in enumerate(lines) if ln.strip() == "SHIPPER CONSIGNEE"), None)
    if shipper_idx is not None and shipper_idx + 1 < len(lines):
        sm = _MOL_DN_GROUP_RE.match(lines[shipper_idx + 1].strip())
        if sm:
            h["shipper"] = sm.group(1)

    rows = []
    for m in _MOL_DN_ITEM_RE.finditer(page):
        desc = m.group("desc").strip()
        qty, rate, pdf_amt = float(m.group("qty")), float(m.group("rate")), float(m.group("pdfamt"))
        amount = round(qty * rate, 2)
        vat = round(pdf_amt - amount, 2)
        en = _MOL_DN_EN_NOTE.get(desc)
        # 英文備註 + 幣別備註都直接寫在 description 裡 (位置不是句尾，不是 english_note() 自動
        # 加得出來的格式，使用者 2026-11-10 答案："Phí AMS (AMS Filing Fee) (USD)")
        full_desc = f"{desc} ({en}) (USD)" if en else f"{desc} (USD)"
        rows.append({**h, "description": full_desc, "amount": amount, "vat": vat})
    return rows


def parse_mol(text: str) -> Dict:
    # 一份 PDF 可能同時混 VAT 發票 (molcs) 跟空運 Debit Note 兩種頁面 (使用者 2026-11-10
    # VN09 答案)，不能用單一 if/elif 二選一，要各自找出自己的頁面分別處理、合併結果；完全沒有
    # 這兩種頁面時 (舊樣本只有海運 Debit Note)，照舊退回 _cw_parse()。
    rows = [r for p in _molcs_pages(text) for r in _molcs_parse_page(p)]
    rows += [r for p in _mol_air_dn_pages(text) for r in _parse_mol_air_dn(p)]
    if rows:
        return {"header": {}, "items": rows}
    return _cw_parse(text, "MOL Logistics Vietnam Inc")


def _mol_totals(text: str, label: str):
    out = {}
    for p in _molcs_pages(text):
        rows = _molcs_parse_page(p)
        if rows:
            # 標籤/數值順序可能相反 (見 _molcs_parse_page() 說明)，兩種都要試。
            value = (_search(label + r"\):\s*([\d.,]+)", p)
                     or _search(r"([\d.,]+)\s*\n\s*\(" + label + r"\):", p))
            out[rows[0]["invoice_no"]] = _molcs_num(value)
    return out


def _mol_grand_total(text: str):
    # 一份 PDF 可能同時混 molcs VAT 發票跟 Debit Note (海運/空運)，兩種各自的含稅總額都要
    # 合併進同一份勾稽結果，不能只挑其中一種 (使用者 2026-11-10 VN09 答案)。
    out = dict(_mol_totals(text, r"Total amount")) if _is_molcs(text) else {}
    out.update(_cw_extract_grand_total(text) or {})
    return out


register_supplier(
    "MOL", "MOL LOGISTICS VIETNAM INC", detect_mol, parse_mol,
    # 一份 PDF 可能混好幾張各自獨立的發票 (VAT 發票 + 空運/海運 Debit Note)，改成逐頁
    # 辨識/擷取，比對時才會用 invoice_no 各自分組 (使用者 2026-11-10 VN09 答案)。
    multi_page=True,
    extract_total=lambda t: _mol_totals(t, r"Total before VAT") if _is_molcs(t) else None,
    extract_grand_total=_mol_grand_total,
    page_text=_cargowise_pages,
    invoice_page_map=lambda p: cargowise_invoice_pages(p, "MOL"),
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
    # 舊版型 (2025 年初樣本) 把字母拆開印出 "N U M B ER"、"D A TE"：每個字母間
    # 允許空白，兼容新舊兩種版型 (2026-11-10 抽樣 EXPEDITORS_JAN/FEB.2025)
    h["invoice_no"] = _search(r"TAX No:\s*N\s*U\s*M\s*B\s*E\s*R\s+(\S+)", text)
    m = re.search(r"CLIENT NO:\s*\S+\s+D\s*A\s*T\s*E\s+(\d{2})/(\d{2})/(\d{2})", text)
    h["invoice_date"] = _format_date_en(m.group(2), m.group(1), f"20{m.group(3)}") if m else None
    # 中間可能夾空行 (樣本 E3Z3101331)
    h["consignee"] = _search(r"CLIENT NO:[^\n]*\n\s*(.+?)\s+INVOICE\s*$", text)
    h["bl_no"] = _search(r"AWB/BL:\s*(\S+)", text)
    h["port_of_discharge"] = _search(r"DESTINATION:\s*(\S+)", text)
    h["vessel"] = _search(r"CARRIER:\s*(.+?)\s*$", text)
    # 材積統一寫法 (見 _fmt_volume)：海運 "5.120 CBM" -> "5.120 CBM"；空運只有
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
    # 單一費用列的金額被字型錯誤吃掉 (例如 "AIR FREIGHT E &5"，正確應為 85)：
    # 這一列抓不到金額、整張又只有這一列費用時，改用 SUB TOTAL 當金額
    # (2026-11-10 抽樣 EXPEDITORS_APR.2026 E3Z2939002)
    if not items and block:
        lines = [l.strip() for l in block.group(1).split("\n")
                 if l.strip() and not l.strip().startswith("BAS ")]
        if len(lines) == 1:
            sub_total = _exp_extract_total(text)
            dm = re.match(r"^(?P<desc>.+?)\s+(?:\d{1,2}|E)\s+\S+$", lines[0])
            if dm and sub_total is not None:
                items.append({"description": dm.group("desc"), "amount": sub_total, "vat": None})
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


def _vt_bill_month(rows: List[list]) -> Optional[Tuple[int, int]]:
    """帳單月份："HÀNG NHẬP THÁNG 1/2025"、"HÀNG XUẤT THÁNG 02/2026" -> (年, 月)。"""
    text = _vt_norm(" ".join(_vn_text(c) for r in rows[:6] for c in r if c not in (None, "")))
    m = re.search(r"THANG\s+(\d{1,2})[/.](\d{4})", text)
    return (int(m.group(2)), int(m.group(1))) if m else None


def _vt_dm_swap(value, bill_month: Optional[Tuple[int, int]]):
    """Excel 日期值日、月放反 (VN02 BK/công nợ/POU HUNG 對帳表)：帳單月份是 1 月，
    卻出現 2025-08-01 這種月份對不上帳單月 (也不是上個月) 的日期，對調日、月後
    (2025-01-08) 就落在帳單月或上個月範圍內，改用對調後的日期。"""
    if not bill_month or not isinstance(value, (datetime.date, datetime.datetime)):
        return value
    y, m, d = value.year, value.month, value.day
    by, bm = bill_month
    py, pm = (by, bm - 1) if bm > 1 else (by - 1, 12)
    if (y, m) in ((by, bm), (py, pm)):
        return value
    if 1 <= m <= 12 and 1 <= d <= 12 and (y, d) in ((by, bm), (py, pm)):
        return datetime.date(y, d, m)
    return value


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
    # "17-Aug"、"17 Aug" (日在前、沒有年份，年份用同一張表的年份) (2026-11-05 VN07 MAERSK Excel)
    m = re.match(r"(\d{1,2})[-\s]([A-Za-z]{3})$", text.strip())
    if m and m.group(2).upper() in months and year:
        return _format_date_en(m.group(1), str(months.index(m.group(2).upper()) + 1), year)
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
           if k in ("total", "total_add", "total_check", "grand", "vat", "date", "cbm", "kg", "port",
                     "group_col", "base_key")}
    keys = [_vt_col(labels, k) for k in cfg["key"]]
    # inv_col：整張沒有單一發票號碼 (doc_key) 時，改用這一欄當每一列的發票號碼 (VN APL
    # Inv# 欄，原文照填包含 "+" 連接符號)；這一欄不存在 (例如 POUYUEN DN 沒有 Inv# 欄) 時
    # 答案要求填 N/A，不要退回用 BOOKING NO 當發票號碼 (2026-11-10 VN11 答案)
    inv_col = _vt_col(labels, cfg["inv_col"]) if cfg.get("inv_col") else None
    # volume_raw：材積直接照表上的數字 (樣本答案 VN06 IMP/EXP，2026-11-04)，依序取第一個有值的欄
    raw_vol_cols = [_vt_col(labels, k) for k in cfg.get("volume_raw", ())]
    # pct_col：費用名稱裡 "{pct}" 代入的百分比欄 (樣本答案 VN06 NB 的 FUEL SURCHARGE 16%，2026-11-05)
    pct_col = _vt_col(labels, cfg["pct_col"]) if cfg.get("pct_col") else None

    def _fee_desc(desc, row, col=None):
        if "{desc_col}" in desc and cfg.get("desc_col"):
            # desc_col：費用名稱接上另一欄的文字 (VN APL：DET - 備註欄)
            dc = find(cfg["desc_col"])
            desc = desc.replace("{desc_col}", _vn_text(_vn_cell(row, dc)) if dc is not None else "")
        if "{days}" in desc and cfg.get("days_label"):
            # days_label：費用名稱裡的 {days} 代入天數欄 (VN02 NEO-DD：PHÍ NEO XE x 天數 ngày)
            dc = find(cfg["days_label"])
            days = _vn_num(_vn_cell(row, dc)) if dc is not None else None
            desc = desc.replace("{days}", str(int(days or 0)))
        if "{shd}" in desc and col:
            # {shd}：費用欄左邊一欄的 SHĐ 憑證號碼 (VN16 Cont 表 NÂNG/HẠ/CÂN)，空白時整段括號不要
            shd = _vn_text(_vn_cell(row, col - 1))
            return desc.replace(" (SHĐ {shd})", f" (SHĐ {shd})" if shd else "").replace("{shd}", shd)
        if "{rmb}" in desc and col:
            # {rmb}：費用欄左邊一欄的 RMB 原幣金額，千分位逗號、小數照原樣 (VN12 NGUYÊN HƯỞNG
            # 答案："air freight (RMB 6,408)"，2026-11-10)
            rmb = _vn_num(_vn_cell(row, col - 1))
            text = (f"{int(rmb):,}" if float(rmb).is_integer() else f"{rmb:,}") if rmb else ""
            return desc.replace("{rmb}", f"RMB {text}" if text else "")
        if "{pct}" not in desc:
            return desc
        v = _vn_num(_vn_cell(row, pct_col)) if pct_col is not None else None
        if v is None and pct_col is not None:
            # 百分比是文字 "19%" (VN16 SONG TRIỀU 表)：_vn_num() 不認百分號，這裡另外讀
            m_pct = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*%\s*", _vn_text(_vn_cell(row, pct_col)))
            v = float(m_pct.group(1)) if m_pct else None
        pct = (round(v * 100) if v is not None and v <= 1 else round(v or 0))
        return desc.replace("{pct}", str(pct))
    # 發票號碼由好幾欄用逗號連起來 (VN02 代墊吊櫃費：SỐ HĐ NÂNG, SỐ HĐ HẠ，使用者 2026-10-27)
    join_keys = [_vt_col(labels, k) for k in cfg.get("key_join", ())]
    fees = [(find(f[0]), f[1]) for f in cfg["fees"]]
    # 費用自己的發票號碼欄 (第三個元素)：費用欄右邊第一個符合的欄。例如 VN18 空運費
    # 與 Local Charge 各有一張紅色發票 (兩欄都叫 "Red invoice no.")
    fee_keys = {}
    for f in cfg["fees"]:
        c = find(f[0])
        if len(f) > 2 and c is not None:
            if f[2].startswith("<"):
                # 第三元素開頭加 "<"：改找費用欄左邊最近的符合欄 (VN18 Local Charge 用 A/F 那張發票，2026-11-05)
                left = _vt_col(labels[:c], "$" + f[2][1:])
                fee_keys[c] = left
            else:
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
    # 欄位可以直接給欄號 (int)：表頭沒有標題的欄 (VN12 NGUYÊN HƯỞNG 的 SHENZHEN 欄，2026-11-06)
    fields = {code: (lab if isinstance(lab, int) else find(lab)) for code, lab in cfg.get("fields", {}).items()}
    boxes = {size: _vt_col(labels, *(lab if isinstance(lab, (list, tuple)) else (lab,)))
             for size, lab in cfg.get("boxes", {}).items()}
    if cfg.get("boxes_adjacent"):
        # boxes_adjacent：20'/40' 兩欄因為三列表頭只比對了前兩列，欄名完全相同
        # 分不出來 (VN14 BANGKE IMP "VOL (CBM) CY QTY")，用緊接著右邊的欄位取代
        # ({"base": 共同欄名, "20": 0, "40": 1})
        ba = cfg["boxes_adjacent"]
        base_c = find(ba["base"])
        if base_c is not None:
            boxes.update({size: base_c + off for size, off in ba.items() if size != "base"})
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
        if cfg.get("doc_key_col") is not None:
            # doc_key_col：標籤跟數值隔著空白欄，數值固定在同一列的某一欄 (VN APL 表尾
            # "Invoice No.:" 那一列)。不可用整本攤平後「下一個非空白儲存格」找——數值欄
            # 剛好空白時會撈到隔壁無關的計算數字 (例如 "Local invoice No.:" 那一列隔壁
            # 欄剛好有個不相干的小計 140.6316，2026-11-10 VN11 TRUCK EXPORT 答案)
            match_row = next((r for r in rows
                              if any(isinstance(c, str) and re.search(cfg["doc_key"], c) for c in r)), None)
            doc_key = (_vn_text(_vn_cell(match_row, cfg["doc_key_col"])) or None) if match_row else None
        else:
            # 整張對帳表一個號碼 (例如 VN19 "Mã bảng kê: PYV260901"、VN11 表尾
            # "Invoice No.: 28433")：當發票號碼
            all_text = " ".join(_vn_text(c) for r in rows for c in r if c not in (None, ""))
            m_doc = re.search(cfg["doc_key"], all_text)
            doc_key = m_doc.group(1) if m_doc else None
    if cfg.get("order_doc"):
        if cfg.get("order_doc_col") is not None:
            # order_doc_col：同 doc_key_col，數值固定欄位，不用攤平找「下一個非空白」
            match_row = next((r for r in rows
                              if any(isinstance(c, str) and re.search(cfg["order_doc"], c) for c in r)), None)
            common["order_no"] = (_vn_text(_vn_cell(match_row, cfg["order_doc_col"])) or "N/A") \
                if match_row else "N/A"
        else:
            # order_doc：整張工作表一個 Order No. (VN APL 答案：Local invoice No.，2026-11-06)
            m_order = re.search(cfg["order_doc"], " ".join(_vn_text(c) for r in rows for c in r if c not in (None, "")))
            common["order_no"] = m_order.group(1) if m_order else None
    if "doc_const" in cfg:
        # doc_const：整張工作表固定的發票號碼 (VN02 CHI HỘ 工作表沒有字軌，填 N/A)
        doc_key = cfg["doc_const"]
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
                # vat_amount_last：VAT 列有好幾個金額時取最右邊的合計欄 (VN16 SONG TRIỀU，2026-11-05)
                vat_amt = nums[-1] if cfg.get("vat_amount_last") else nums[0]
                items.append({"description": cfg.get("vat_label") or _vn_text(cells[0]), "amount": None, "vat": vat_amt,
                              **common, "invoice_no": doc_key or f"{name.strip()} VAT"})
                # vat_no_date：整張 VAT 列沒有日期 (樣本答案 VN06 IMP/EXP，2026-11-04)
                # 其他欄位也填 N/A，避免表頭層的共用欄位 (提單、材積) 被補進 VAT 列
                if cfg.get("vat_no_date"):
                    # 有 doc_key (整張一個發票號碼) 時，VAT 列仍用該號碼 (VN06 NB)
                    items[-1].update({k: "N/A" for k in (
                        ("invoice_no",) if not doc_key else ()) + (
                        "invoice_date", "bl_no", "order_no", "volume", "port_of_loading", "port_of_discharge",
                        "vessel", "container_no", "shipper", "onboard_date", "arrive_date")})
            break
    if cfg.get("vat_total"):
        # 表尾的整張 VAT ("Total VAT amount | 2590843.48")，取該列最後一個數字
        for row in rows[end:]:
            if any(cfg["vat_total"] in _vt_norm(c) for c in row):
                nums = [_vn_amount(c) for c in row if _vn_amount(c)]
                if nums:
                    items.append({"description": cfg.get("vat_desc", "VAT"), "amount": None,
                                  "vat": nums[-1], **common, "invoice_no": doc_key or f"{name.strip()} VAT"})
                    if cfg.get("vat_na_fields"):
                        # vat_na_fields：整張 VAT 列指定欄位填 N/A (VN APL 答案)
                        items[-1].update({k: "N/A" for k in cfg["vat_na_fields"]})
                    elif cfg.get("vat_no_date"):
                        # 整張 VAT 列只有發票號碼與金額 (VN12 NGUYÊN HƯỞNG 答案)：其他欄填 N/A
                        items[-1].update({k: "N/A" for k in (
                            "invoice_date", "bl_no", "order_no", "volume", "port_of_loading", "port_of_discharge",
                            "vessel", "container_no", "shipper", "onboard_date", "arrive_date")})
                break
    last_key = last_date = None
    cur_grp, grp_parts = None, []      # merge_no_fee 用：同一台車的列
    last_base, group_n, last_fields = "", 0, {}
    sheet_year = None
    if col.get("date") is not None:
        years = collections.Counter((_vn_date(_vn_cell(r, col["date"])) or "")[-4:] for r in rows[start:end])
        years.pop("", None)
        sheet_year = years.most_common(1)[0][0] if years else (common["invoice_date"] or "")[-4:] or None
    bill_month = _vt_bill_month(rows) if cfg.get("date_dm_swap") else None
    for row in rows[start:end]:
        key = next((_vn_text(_vn_cell(row, k)) for k in keys if _vn_text(_vn_cell(row, k))), "")
        multi_key = False
        if cfg.get("key_comma") and key:
            # key_comma：一格裡換行的多個號碼用 ", " 連接 (VN16 SONG TRIỀU 答案，2026-11-05)
            raw_key = next(str(_vn_cell(row, k)) for k in keys if _vn_text(_vn_cell(row, k)))
            lines = [ln.strip() for ln in raw_key.split("\n") if ln.strip()]
            key = ", ".join(lines) if lines else key
            multi_key = len(lines) > 1
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
        # 電子發票保留完整發票符號 ("1C26TSG-7734"、"1-C26TOD-28667"、"C26TOD/17995")，
        # 不去掉符號 (使用者 2026-11-05 決定：所有供應商都保留完整符號，取代 2026-10-27 的「只用號碼」)
        if cfg.get("inv_series") and str(key).isdigit():
            # inv_series：Excel 只有流水號，發票符號 = 系列 1C + 年份 YY + TAA + "-" + 8 位流水號
            # (VN17 TIMES CARGO 新答案，2026-11-06)；年份取發票日期
            d_cell = _vn_cell(row, col["date"]) if col.get("date") is not None else None
            yy = d_cell.year % 100 if isinstance(d_cell, (datetime.date, datetime.datetime)) else 0
            key = cfg["inv_series"].format(yy=f"{yy:02d}") + "-" + str(key).zfill(8)
        last_key = key
        if not doc_key and cfg.get("date_key"):
            # 沒有整張發票符號時，改用這一列自己的 Pickup Date (yyyy/mm/dd) 當發票號碼
            # (VN12 NGUYÊN HƯỞNG 舊版答案，2026-11-10)
            d_cell = _vn_cell(row, col["date"]) if col.get("date") is not None else None
            key = d_cell.strftime("%Y/%m/%d") if isinstance(d_cell, (datetime.date, datetime.datetime)) else "N/A"
        elif not doc_key and cfg.get("inv_col"):
            # 沒有整張的發票號碼 (doc_key) 時，改用 Inv# 欄當每一列的發票號碼 (不是 BOOKING
            # NO)；這張表沒有 Inv# 欄就填 N/A (VN11 答案：01-2025 TRUCK EXPORT 有 Inv# 欄、
            # POUYUEN DN 沒有，2026-11-10)
            raw_inv = _vn_text(_vn_cell(row, inv_col)) if inv_col is not None else None
            key = raw_inv.lstrip("'") if raw_inv else "N/A"
        # vat_doc_only：對帳表號只給整張的 VAT 列用，費用列用自己的號碼 (VN11 Inv#)
        else:
            key = key if cfg.get("vat_doc_only") else (doc_key or key)
        if cfg.get("key_strip_space") and isinstance(key, str):
            # 發票號碼去掉所有空白 ("1-C26TOD- 40183" -> "1-C26TOD-40183"，VN14 答案 2026-11-10)
            key = re.sub(r"\s+", "", key)
        head = dict(common, invoice_no=key)
        for code, c in fields.items():
            if c is not None and _vn_text(_vn_cell(row, c)):
                value = _vn_cell(row, c)
                if code.endswith("_date"):
                    # 日期前面可能接船務慣用縮寫 ("ATD 02/09/2026"，VN14 POUYUEN EXP on board date)
                    stripped = re.sub(r"^(ATD|ETD|ATA|ETA)\s+", "", _vn_text(value))
                    head[code] = _vn_date(value) or _vn_date(stripped) or _vn_text(value)
                else:
                    head[code] = _vn_text(value)
            # field_suffix：欄位值後面加固定文字 (VN19 IMX 船名/航班加 " (Flight)"，答案 2026-11-05)
            if head.get(code) and code in cfg.get("field_suffix", {}):
                head[code] += cfg["field_suffix"][code]
        # order_is_key / bl_is_key：發票號碼同時當 Order No. 或 B/L NO. (VN16 SONG TRIỀU 答案，2026-11-05)
        if cfg.get("order_is_key"):
            head["order_no"] = key or None
        if cfg.get("bl_is_key"):
            head["bl_no"] = key or None
        if cfg.get("volume_prefix") and head.get("volume"):
            # volume_prefix：櫃型前面加字，例如 "1 x " + 40HC (VN APL 答案)
            head["volume"] = cfg["volume_prefix"] + str(head["volume"])
        for code in cfg.get("carry_fields", ()):
            # 續行沿用同一票上一列的欄位 (提單號、港口)
            if head.get(code):
                last_fields[code] = head[code]
            elif last_fields.get(code):
                head[code] = last_fields[code]
        if col.get("date") is not None:
            raw_date = _vn_cell(row, col["date"])
            if bill_month is not None:
                raw_date = _vt_dm_swap(raw_date, bill_month)
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
        # port_arrow：起迄港用 "->" 分隔 (VN19 IMX ROUTINE "BUSAN / KOREA -> HCM /VN")
        port_sep = "->" if cfg.get("port_arrow") else " - "
        if col.get("port") is not None and port_sep in _vn_text(_vn_cell(row, col["port"])):
            pol, pod = _vn_text(_vn_cell(row, col["port"])).split(port_sep, 1)
            head.setdefault("port_of_loading", pol.strip())
            head.setdefault("port_of_discharge", pod.strip())
        def measure(name_):
            # Excel 浮點誤差 (6.912000000000001) 取到小數 3 位
            value = _vn_num(_vn_cell(row, col.get(name_))) if col.get(name_) is not None else None
            return f"{round(value, 3):g}" if value else None

        box_containers = [(_vn_num(_vn_cell(row, c)), size) for size, c in boxes.items() if c is not None]
        box_vol = _fmt_volume(containers=box_containers)
        cbm_val = measure("cbm")
        if cfg.get("volume_both") and box_vol and cbm_val and float(cbm_val) > 0:
            # volume_both：櫃型與 CBM 都要列出 (VN14 TRUCK/BANGKE EXP 表頭 "CY QUANTITY
            # 20／40 + CFS CBM"，答案 2026-11-10：散貨裝進整櫃，兩個數字都要留)
            head["volume"] = f"{box_vol}, {cbm_val} CBM"
        else:
            head["volume"] = box_vol or _fmt_volume(cbm=cbm_val, kg=measure("kg"))
        if cfg.get("carry_volume"):
            # carry_volume：續行 (同一個 Bill 的 Ocean Freight 發票) 沒有自己的材積欄，
            # 沿用上一張發票的材積 (VN14 CN POULI 答案，2026-11-10)
            if head.get("volume"):
                last_fields["volume"] = head["volume"]
            elif last_fields.get("volume"):
                head["volume"] = last_fields["volume"]
        if cfg.get("volume_col") is not None:
            # volume_col：材積直接取一欄的櫃型文字，前面加 volume_prefix (VN APL：1 x 40HC)
            vraw = _vn_text(_vn_cell(row, cfg["volume_col"]))
            if vraw:
                head["volume"] = (cfg.get("volume_prefix") or "") + vraw
        raw_vol = next((_vn_cell(row, c) for c in raw_vol_cols if c is not None and _vn_text(_vn_cell(row, c))), None)
        if raw_vol is not None:
            # 數字照原樣 (1.129、20)，文字照原文 ("40HC")
            if isinstance(raw_vol, (int, float)):
                head["volume"] = int(raw_vol) if float(raw_vol).is_integer() else round(raw_vol, 3)
            else:
                head["volume"] = _vn_text(raw_vol)
        if cfg.get("cbm_always") and col.get("cbm") is not None and _vn_cell(row, col["cbm"]) not in (None, ""):
            # 有 CBM 欄就一律寫 CBM，即使是 0 (樣本答案 08-CFS：CBM 0 寫 "0 CBM"，不改用重量)
            cbm_val = _vn_num(_vn_cell(row, col["cbm"])) or 0
            head["volume"] = f"{round(cbm_val, 3):g} CBM"
        if cfg.get("kg_always") and col.get("kg") is not None and _vn_cell(row, col["kg"]) not in (None, ""):
            # 同上，有 KG 欄就一律寫 KG (樣本答案 08-AIR：KG 0 寫 "0 KG")
            kg_val = _vn_num(_vn_cell(row, col["kg"])) or 0
            head["volume"] = f"{round(kg_val, 3):g} KG"
        if cfg.get("merge_no_fee"):
            # 同一台車的續行列 (沒有費用、只有單號)：併入上一張 (有費用的那一列)，
            # 單號用 ", " 連接、材積用 " + " 連接 (樣本答案 VN06 NB，2026-11-04)
            if any(_vn_amount(_vn_cell(row, c)) for c, _ in fees) or cur_grp is None:
                cur_grp = len(grp_parts)
                grp_parts.append([])
            grp_parts[cur_grp].append({"bl_no": head.get("bl_no"), "order_no": head.get("order_no"),
                                       "volume": head.get("volume")})
            head["_grp"] = cur_grp
        # 費用名稱 "@欄位" 代表取該欄的文字 (例如 VN07 Remark 欄 "Late come charge")
        row_items = [{"description": (_vn_text(_vn_cell(row, find(desc[1:]))) or desc[1:])
                      if desc.startswith("@") else _fee_desc(desc, row, c),
                      "amount": _vn_amount(_vn_cell(row, c)), "vat": None, **head}
                     for c, desc in fees if _vn_amount(_vn_cell(row, c))]
        if cfg.get("vnd_alloc") and row_items:
            # vnd_alloc：費用欄是原幣 (USD)，金額改成越南盾：正常情況下把這一列的未稅
            # 小計 (pretax) 依 USD 比例分攤到每個費用、尾差放最後一列；小計欄跟目前列出的
            # 費用對不上 (來源帳單本身的小計算錯，例如少了某筆費用的附件) 時，改用
            # 「USD ÷ (含 VAT 時先 ÷1.08) × 匯率」各自獨立換算，不勉強湊出小計，讓後續的
            # 金額稽核自然抓出不一致 (VN14 格式 D2 BANGKE EXP C26TOD-21469，2026-11-10 答案)
            ac = cfg["vnd_alloc"]
            pretax_c = find(ac["pretax"])
            rate_c = find(ac["rate"])
            usd_total_c = find(ac["usd_total"]) if ac.get("usd_total") else None
            pretax_val = _vn_num(_vn_cell(row, pretax_c)) if pretax_c is not None else None
            rate_val = _vn_num(_vn_cell(row, rate_c)) if rate_c is not None else None
            recorded_usd = _vn_num(_vn_cell(row, usd_total_c)) if usd_total_c is not None else None
            usd_vals = [it["amount"] for it in row_items]
            usd_sum = sum(usd_vals)
            divisor = 1.08 if ac.get("vat_included") else 1.0
            mismatch = recorded_usd is not None and abs(recorded_usd - usd_sum) > 0.5
            if mismatch and rate_val:
                allocated = [round(uv / divisor * rate_val) for uv in usd_vals]
            elif pretax_val is not None and usd_sum:
                allocated, running = [], 0
                for i, uv in enumerate(usd_vals):
                    if i < len(usd_vals) - 1:
                        amt = round(pretax_val * uv / usd_sum)
                        running += amt
                        allocated.append(amt)
                    else:
                        allocated.append(pretax_val - running)
            else:
                allocated = usd_vals
            for it, uv, amt in zip(row_items, usd_vals, allocated):
                # 欄位表頭本身寫了幣別 (子標題 "USD") 時，前面已經自動加過裸的 "(USD)"，
                # 這裡要換成含金額的版本，不要兩個都留 ("TRUCKING FEE (USD) (USD 442.8)")
                base_desc = re.sub(r"\s*\(USD\)$", "", it["description"])
                it["description"] = f"{base_desc} (USD {uv:g})"
                it["amount"] = amt
        if cfg.get("vnd_rate") and row_items:
            # vnd_rate：費用欄是原幣 (USD)，沒有分攤用的小計欄：金額 = USD × 匯率欄，VAT =
            # 含稅總額欄 − 這一列換算後的費用加總 (VN14 格式 B BANGKE IMP，2026-11-10 答案)
            rc = cfg["vnd_rate"]
            rate_c = find(rc["rate"])
            grand_c = find(rc["grand"])
            rate_val = _vn_num(_vn_cell(row, rate_c)) if rate_c is not None else None
            grand_val = _vn_num(_vn_cell(row, grand_c)) if grand_c is not None else None
            for it in row_items:
                uv = it["amount"]
                base_desc = re.sub(r"\s*\(USD\)$", "", it["description"])
                it["description"] = f"{base_desc} (USD {uv:g})"
                it["amount"] = round(uv * rate_val) if rate_val else uv
            if grand_val is not None:
                vat_amt = grand_val - sum(it["amount"] for it in row_items)
                row_items.append({"description": f"VAT {0 if not round(vat_amt, 2) else 8}%",
                                   "amount": None, "vat": vat_amt, **head})
        if cfg.get("gross_vat"):
            # gross_vat：帳單金額是含 8% VAT 的總額 -> 未稅 = 總額 ÷ 1.08 (四捨五入)，VAT = 總額 − 未稅
            # (VN02 YUEDE 新答案，2026-11-05)
            for it in row_items:
                gross = it["amount"]
                net = math.floor(gross / 1.08 + 0.5)
                it["amount"], it["vat"] = net, gross - net
        if cfg.get("vat_zero_desc"):
            # vat_zero_desc：這些費用名稱的 VAT 寫 0 (VN12 空運費免稅，2026-11-06 新答案)；
            # 用開頭比對 (不要求完全相等)，因為費用名稱後面可能接原幣金額備註
            # (VN12 "air freight (RMB 6,408)"，2026-11-10 新答案)
            for it in row_items:
                if str(it["description"]).upper().startswith(cfg["vat_zero_desc"]):
                    it["vat"] = 0
        if cfg.get("vat_of_amount"):
            # vat_of_amount：未稅金額 × 8% 寫在同一列 (VN02 NEO-DD 答案)
            for it in row_items:
                it["vat"] = round(it["amount"] * 0.08)
        for it in row_items:
            if isinstance(it["description"], str) and it["description"].startswith("-") and it["amount"]:
                # 減項 (費用名稱前加 "-")：金額取負號
                it["description"], it["amount"] = it["description"][1:], -abs(it["amount"])
        for it, (c, _) in zip(row_items, [f for f in fees if _vn_amount(_vn_cell(row, f[0]))]):
            own = _vn_text(_vn_cell(row, fee_keys.get(c))) if fee_keys.get(c) is not None else ""
            if own:
                # 電子發票保留完整發票符號 (使用者 2026-11-05 決定)，照帳單寫入 Red invoice no.
                it["invoice_no"] = own
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
        # desc_usd：費用名稱後面加括號註記 J 欄的外幣金額 ("Late come charge (USD 13.824)")；
        # 數字加 "USD " 前綴，文字照抄 ("Min charge 50USD")。(2026-11-05 VN07 答案提示列)
        usd_col = find(cfg["desc_usd"]) if cfg.get("desc_usd") else None
        if usd_col is not None and row_items:
            raw_usd = _vn_cell(row, usd_col)
            usd_num = _vn_num(raw_usd)
            # 千分位逗號 (答案 "(USD 1,473)"、"(USD 1,000)"，VN19 IMX / TRANS_WAGON 2026-11-05)；小數照原樣
            usd_note = (f"USD {usd_num:,.3f}".rstrip("0").rstrip(".") if usd_num is not None
                        else _vn_text(raw_usd))
            only = cfg.get("desc_usd_only")  # 只加在指定費用名稱的列 (VN18 只加 A/F)
            if usd_note:
                for it in row_items:
                    if only is None or it["description"] == only:
                        it["description"] = f"{it['description']} ({usd_note})"
        # vat_zero_ok：VAT 是 0 也要寫一列 ("VAT 0%"，VN14 Ocean Freight 免稅發票，
        # 2026-11-10 答案)；一般供應商 VAT 欄空白/0 當沒有，維持原行為
        vat_zero_ok = cfg.get("vat_zero_ok")
        vat = (_vn_num(_vn_cell(row, col.get("vat"))) if vat_zero_ok
               else _vn_amount(_vn_cell(row, col.get("vat")))) if col.get("vat") is not None else None
        has_vat = vat is not None if vat_zero_ok else bool(vat)
        if has_vat and cfg.get("vat_same_row") and row_items:
            # 每列都有自己的 VAT：寫在費用同一列的 vat 欄，不另外拆成 VAT 列 (2026-11-05 VN07 MAERSK Excel)
            row_items[0]["vat"] = vat
        elif has_vat:
            desc = f"VAT {0 if not vat else 8}%" if cfg.get("vat_pct_desc") else cfg.get("vat_desc", "VAT")
            row_items.append({"description": desc, "amount": None, "vat": vat, **head})
        if not row_items:
            continue
        items += row_items
        # total_skip_multi_key：合併好幾個發票號碼的列 (key_comma 的多行儲存格)，
        # 帳單自己的 TỔNG 常常是舊值沒跟著更新 (VN16 Chuyển kho xe cont 沒有
        # TOTAL PLV CHECK 核對欄可以交叉比對，使用者 2026-11-11 確認不稽核這種列，
        # 避免把帳單本身的殘留舊值誤判成擷取錯誤)
        if multi_key and cfg.get("total_skip_multi_key"):
            continue
        # 合計欄空白 (公式沒有存計算結果) 時當成沒有合計，不是 0
        t = totals.setdefault(key, [None, None])
        # total_add：另一個幣別的合計 (VN14 費用有 USD 也有 VND，使用者 2026-10-27 要求
        # 兩種都稽核)：USD 合計 + VND 合計 = 兩種費用的加總，任一幣別對不上都會不通過
        for i, name_ in ((0, "total"), (0, "total_add"), (1, "grand")):
            value = _vn_num(_vn_cell(row, col[name_])) if col.get(name_) is not None else None
            if name_ == "total" and col.get("total_check") is not None:
                # total_check：帳單另外有一欄自己核對明細加總用 (VN16 TOTAL PLV
                # CHECK)，優先採用；跟 TỔNG 常常各自有缺漏、彼此對不上
                # (2026-11-11 使用者確認)。費用加總符合核對欄或 TỔNG 任一個就算
                # 通過；核對欄這一列空白 (沒有可信基準) 就不稽核這張發票；核對欄
                # 有值但兩邊都對不上，保留 TỔNG 讓稽核照常抓出不一致 (真的異常)
                row_sum = round(sum(it.get("amount") or 0 for it in row_items), 2)
                check_val = _vn_num(_vn_cell(row, col["total_check"]))
                if check_val is not None:
                    if abs(row_sum - check_val) <= 1 or (value is not None and abs(row_sum - value) <= 1):
                        value = row_sum
                    else:
                        value = check_val
                else:
                    value = None
            if value is not None:
                t[i] = (t[i] or 0) + value
    if cfg.get("merge_no_fee"):
        for it in items:
            grp = it.pop("_grp", None)
            if grp is None:
                continue
            parts = grp_parts[grp]

            def joined(field, sep=", "):
                vals = [str(p[field]) for p in parts if p.get(field) not in (None, "", "N/A")]
                return sep.join(dict.fromkeys(vals)) or "N/A"
            it["bl_no"] = joined("bl_no")
            it["order_no"] = joined("order_no")
            it["volume"] = joined("volume", " + ")
    # 整張表一個 VAT 的列沒有自己的日期，抬頭也沒寫日期時：用這張表最晚的日期
    dated = [datetime.datetime.strptime(it["invoice_date"], "%d %b %Y") for it in items
             if it.get("invoice_date") and re.fullmatch(r"\d{2} \w{3} \d{4}", it["invoice_date"])]
    for it in items:
        if not it.get("invoice_date") and it.get("amount") is None and dated:
            it["invoice_date"] = max(dated).strftime("%d %b %Y")
    if cfg.get("vat_last"):
        # 整張 VAT 列放在最後一列 (樣本答案 VN06 IMP/EXP，2026-11-04)
        items = [it for it in items if it.get("amount") is not None] + \
                [it for it in items if it.get("amount") is None]
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
    # 簡稱後面加國家備註 ("YUE DE (VN)"，VN14 BANGKE IMP TO: 欄，2026-11-10 答案)：
    # 先去掉這個備註才比對，避免因為多出的字母比對不到
    bare = re.sub(r"\s*\((?:VN|VIETNAM)\)\s*$", "", _vt_norm(name or ""))
    flat = re.sub(r"[^A-Z]", "", bare)
    full = next((full for key, full in _VN_CONSIGNEES.items() if key.replace(" ", "") == flat), None)
    if full:
        return full
    # VN02 YUEDE 客戶 (DỤ ĐỨC)：答案用英文名稱 (2026-11-05 新答案)
    if re.search(r"\bDU\s*DUC\b", _vt_norm(name or "")) and "VIET" in _vt_norm(name or ""):
        return "YUE DE VIETNAM COMPANY LIMITED"
    # 整句越南公司名 ("CÔNG TY TNHH POU HUNG VIỆT NAM"，樣本答案 VN06 NB 表)：含 "POU 寶成廠別" 且是公司名才換
    norm = _vt_norm(name or "")
    m = re.search(r"\bPOU\s*(SUNG|CHEN|YUEN|LI|HUNG)\b", norm)
    if m and ("CONG TY" in norm or "VIET" in norm):
        return _VN_CONSIGNEES.get(f"POU {m.group(1)}", name)
    return name


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
            # sum_row 也會算出一個可勾稽的未稅合計 (表尾 TOTAL 列各費用欄加總)，不是只有
            # "total"/"sheet_total" 才算有總額可勾稽 (2026-11-12 發現：VN18 SAIGON_TRANSPORT
            # 等好幾家供應商只設定 sum_row，這裡漏了判斷，整份一直顯示「沒有可勾稽的總額」)
            has_total |= "total" in cfg or "sheet_total" in cfg or "sum_row" in cfg
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
    # 簡化版表頭沒有 "TOTAL(VND) NOT VAT LOCAL" / "TOTAL(VND) VAT LOCAL CHARGE" 欄，
    # 只有裸的 "VND" (費用欄本身) 與 "Total Amount included VAT" (2026-11-11 抽樣
    # "CONG NO T01-25 POU LI"：原本找不到這兩欄，整張沒有總額可勾稽)
    (r"(?!)", {"header": "BOOKING NO", "key": ["=INVOICE NO", "BOOKING NO"], "carry": True,
               "fees": [("=VND", "@RATE")], "vat": "VAT (8%)", "vat_desc": "VAT 8%",
               "total": ("TOTAL(VND) NOT VAT LOCAL", "=VND"),
               "grand": ("=TOTAL(VND) VAT LOCAL CHARGE", "=TOTAL AMOUNT INCLUDED VAT"),
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
    # POU SUNG 版 (2026-11-04 答案)：發票號碼保留完整 "1C25TYY-3700" (不去前綴)
    (r"(?!)", {"header": "RED INV NO", "two_rows": False, "key": ["RED INV NO"], "einv_key": True,
               "keep_einv_prefix": True,
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
    # 供應商名稱用英文帳單抬頭 (使用者 2026-11-04 規定，與 Excel 版一致)
    h: Dict[str, Optional[str]] = {"supplier": _TW_EXCEL_SUPPLIER}
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


# --- VN01 TRANS WAGON：CONG NO 對帳表 (POU CHEN 版，2026-11-04 使用者答案) -------
# 單一工作表 (例 "POU CHEN")，每列一張發票 (INVOICE NO 在 T 欄)。費用欄 DO / CFS/THC /
# OTHER (DRA/LOAD/HD/CIC) / O/FRT / EXW/FCA/OTHER 是 USD，匯率在 "Tel:" 那一列的數字欄
# (J7 = 25,549)；每個費用 × 匯率 四捨五入成 VND (答案檔每項費用各一列)。
# 每張發票最後一列 VAT (W 欄，本檔為 0)。BOOKING NO 與材積只寫在一票的第一列，往下沿用。
# 材積：有櫃數 (D/E 欄) 寫 "1 x 20' + 1 x 40'"；只有 CBM (F 欄) 寫 "8.001 CBM"。
# 供應商名稱照答案檔的帳單抬頭 "TRANS WAGON INTERNATIONAL CO.,LTD"。
_TW_EXCEL_SUPPLIER = "TRANS WAGON INTERNATIONAL CO.,LTD"
_TW_CONG_NO_FEES = ["DO", "CFS/THC", "OTHER (DRA/LOAD/HD/CIC)", "O/FRT", "EXW/FCA/OTHER"]


def _tw_key(value) -> str:
    """表頭比對用：轉大寫並去掉所有空白 (含換行)。"""
    return re.sub(r"\s", "", _vn_label(value))


def _tw_cong_no_head(rows: List[list]) -> Tuple[Optional[int], List[str]]:
    """找 CONG NO 版的表頭列：回傳 (列索引, 去空白的欄位標籤)；不是這版回傳 (None, [])。"""
    for i, row in enumerate(rows[:20]):
        labels = [_tw_key(c) for c in row]
        if {"INVOICENO", "BOOKINGNO", "AMOUNT(NOTVAT)", "INVOICEDATE", "CFS/THC", "O/FRT",
            "EXW/FCA/OTHER"} <= set(labels):
            return i, labels
    return None, []


def _tw_cong_no_rate(rows: List[list]) -> Optional[float]:
    """匯率：「Tel:」那一列的第一個數字欄 (J7)。"""
    for row in rows[:12]:
        if _vn_text(_vn_cell(row, 0)).upper().startswith("TEL"):
            return next((v for v in (_vn_num(c) for c in row[1:]) if v), None)
    return None


def _tw_cong_no_date(value) -> Optional[str]:
    """發票日期 "16/01/2025" 或 Excel 日期 -> "16 Jan 2025" (同其他 VN 規則的格式)。"""
    if isinstance(value, (datetime.date, datetime.datetime)):
        return value.strftime("%d %b %Y")
    m = re.fullmatch(r"\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*", _vn_text(value))
    return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1))).strftime("%d %b %Y") if m else None


_TW_CONST_FORMULA = re.compile(r"^=[\d.+\-*/() ]+$")


def _tw_const_formulas(path: str, sheet: str) -> Dict[Tuple[int, int], float]:
    """純數字算式儲存格 (例如 =220+370、=350+650) 若檔案沒存計算結果，calamine 讀成空白。
    這裡直接算出 Excel 會顯示的值。只接受數字與 + - * / ( ) 的算式。
    回傳 {(列索引, 欄索引): 值}，索引從 0 起算，與 read_spreadsheet() 的列欄一致。"""
    import openpyxl
    out: Dict[Tuple[int, int], float] = {}
    if not str(path).lower().endswith((".xlsx", ".xlsm")):
        return out
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        if sheet not in wb.sheetnames:
            return out
        for r_i, row in enumerate(wb[sheet].iter_rows(min_row=1, min_col=1)):
            for c_i, cell in enumerate(row):
                if isinstance(cell.value, str) and _TW_CONST_FORMULA.match(cell.value):
                    out[(r_i, c_i)] = float(eval(cell.value[1:], {"__builtins__": {}}))
    finally:
        wb.close()
    return out


def _tw_cong_no_sheets(book) -> Tuple[List[Dict], List[Dict]]:
    """CONG NO 版：回傳 (明細 items, 每張發票的合計 totals)。沒有這版表頭時回傳空值。"""
    items: List[Dict] = []
    totals: List[Dict] = []
    for name, rows in book.sheets:
        head, labels = _tw_cong_no_head(rows)
        rate = _tw_cong_no_rate(rows) if head is not None else None
        if head is None or rate is None:
            continue
        col = {k: labels.index(k) for k in ("BOOKINGNO", "INVOICENO", "INVOICEDATE", "AMOUNT(NOTVAT)", "VAT",
                                             "TOTALAMOUNT")
               if k in labels}
        fees = [(name, labels.index(_tw_key(name))) for name in _TW_CONG_NO_FEES if _tw_key(name) in labels]
        sub = [_tw_key(c) for c in rows[head + 1]] if head + 1 < len(rows) else []   # 第二列：20' / 40' / CMB
        c20, c40, ccbm = (sub.index(k) if k in sub else None for k in ("20'", "40'", "CMB"))
        consignee = _vt_consignee(rows)
        booking = None
        consts = _tw_const_formulas(book.path, name)
        for ridx in range(head + 2, len(rows)):
            row = rows[ridx]
            if _vn_text(_vn_cell(row, 0)).upper().startswith("TOTAL"):
                break
            invoice = _vn_text(_vn_cell(row, col["INVOICENO"]))
            if not invoice:
                continue                                  # 空白列或費用續行 (沒有發票號碼)
            bk = _vn_text(_vn_cell(row, col["BOOKINGNO"]))
            if bk:                                        # 新的一票：材積跟著這一列重讀
                booking = bk
                n20, n40 = _vn_num(_vn_cell(row, c20)), _vn_num(_vn_cell(row, c40))
                cbm = _vn_num(_vn_cell(row, ccbm))
                boxes = [f"{int(n)}X{size}" for n, size in ((n20, "20"), (n40, "40")) if n]   # 統一寫法 1X40 (2026-11-10)
                volume = ", ".join(boxes) if boxes else (f"{cbm:g} CBM" if cbm else None)
            base = {"supplier": _TW_EXCEL_SUPPLIER, "consignee": consignee, "invoice_no": invoice,
                    "invoice_date": _tw_cong_no_date(_vn_cell(row, col["INVOICEDATE"])),
                    "bl_no": booking, "volume": volume}
            for name, c in fees:
                usd = _vn_num(_vn_cell(row, c))
                if usd is None:                           # 沒存計算結果的算式 (例如 =220+370)
                    usd = consts.get((ridx, c))
                if usd:                                   # 0 或空白不列
                    # Excel ROUND 是四捨五入 (金額皆為正數)
                    # 費用名稱加註 USD 金額 (答案檔：DO (USD 65)、O/FRT (USD 1,000)、CFS/THC (USD 35.84))
                    usd_text = f"{usd:,.2f}".rstrip("0").rstrip(".")
                    items.append({**base, "description": f"{name} (USD {usd_text})",
                                  "amount": int(usd * rate + 0.5),
                                  "vat": None})
            # 每張發票最後一列 VAT (本檔為 0)
            items.append({**base, "description": "VAT", "amount": None,
                          "vat": _vn_num(_vn_cell(row, col.get("VAT"))) or 0})
            totals.append({"invoice_no": invoice, "total": _vn_num(_vn_cell(row, col["AMOUNT(NOTVAT)"])),
                           "grand_total": _vn_num(_vn_cell(row, col.get("TOTALAMOUNT")))})
    return items, totals


def _tw_is_cong_no(book) -> bool:
    return any(_tw_cong_no_head(rows)[0] is not None for _, rows in book.sheets)


_tw_vt_parse, _tw_vt_totals = _vt_make("CÔNG TY TNHH LIÊN VẬN QUỐC TẾ", _VT_TRANS_WAGON)


# --- VN01 TRANS WAGON：PAYMENT FOR LOCAL CHARGE (工作表 BẢNG KÊ，2026-11-09 使用者答案) ---
# 一列一票：STT/Invoice No/B/L/ATD/XRATE + 費用欄 (USD，欄位依月份不同，動態抓
# XRATE 右邊到 Total 左邊的每一欄) + Total/VAT/Sum Total (USD)/Sum Total (VND)。
# 金額 = USD × XRATE 四捨五入；每票最後加一列 VAT 8% = Sum Total(VND) − 費用 VND 加總。
# 答案的公司全名含逗號後空格 "CO., LTD" (跟既有 CONG_NO 版 "CO.,LTD" 不同，照答案)。
_TW_SUPPLIER_SPACE = "TRANS WAGON INTERNATIONAL CO., LTD"


def _tw_header_consignee(rows: List[list]) -> Optional[str]:
    """抬頭客戶公司名稱："To : …" 或沒有標籤、直接寫公司名的那一格 (例如
    "CONG TY POU PHONG VIETNAM")。簡稱/越南文名稱用 _VN_CONSIGNEES 換成英文全名。"""
    text = _vt_consignee(rows)
    if not text:
        for row in rows[:10]:
            for c in row:
                if not c:
                    continue
                first = str(c).splitlines()[0].strip()
                if first and re.match(r"^C[ÔO]NG\s+TY\b", _vt_norm(first)) and "TRANS WAGON" not in first.upper():
                    text = first
                    break
            if text:
                break
    if not text:
        return None
    flat = re.sub(r"[^A-Z]", "", _vt_norm(text))
    for key, full in _VN_CONSIGNEES.items():
        if re.sub(r"[^A-Z]", "", key) in flat:
            return full
    return text


def _tw_usd_text(usd: float) -> str:
    """費用名稱括號裡的美金金額：75 -> "75"、40.6 -> "40.6"。"""
    return f"{usd:,.2f}".rstrip("0").rstrip(".")


def _tw_bang_ke_head(rows: List[list]) -> Tuple[Optional[int], List[str]]:
    for i, row in enumerate(rows[:20]):
        labels = [_tw_key(c) for c in row]
        if {"B/L", "ATD", "XRATE", "SUMTOTAL(VND)"} <= set(labels):
            return i, labels
    return None, []


def _tw_is_bang_ke(book) -> bool:
    return any(_tw_bang_ke_head(rows)[0] is not None for _, rows in book.sheets)


def _tw_bang_ke_sheets(book) -> Tuple[List[Dict], List[Dict]]:
    """「PAYMENT FOR LOCAL CHARGE」版：回傳 (費用列, 每張發票的合計)。"""
    items: List[Dict] = []
    totals: List[Dict] = []
    for name, rows in book.sheets:
        head, labels = _tw_bang_ke_head(rows)
        if head is None:
            continue
        col = {k: labels.index(k) for k in (
            "INVOICENO", "B/L", "ATD", "XRATE", "TOTAL", "SUMTOTAL(VND)", "VATINVDATE",
            "NOTEPORTOFLOADING/DISCHARGE") if k in labels}
        fee_cols = list(range(col["XRATE"] + 1, col["TOTAL"]))
        fee_names = {c: re.sub(r"\s+", " ", _vn_text(rows[head][c])).strip() for c in fee_cols}
        consignee = _tw_header_consignee(rows)
        for row in rows[head + 1:]:
            bl = _vn_text(_vn_cell(row, col["B/L"]))
            if bl.upper() == "TOTAL":
                break
            invoice = _vn_text(_vn_cell(row, col["INVOICENO"]))
            if not invoice:
                continue
            rate = _vn_num(_vn_cell(row, col["XRATE"])) or 0
            note = _vn_text(_vn_cell(row, col.get("NOTEPORTOFLOADING/DISCHARGE")))
            pol = pod = None
            if " - " in note:
                pol, pod = (p.strip() for p in note.split(" - ", 1))
            base = {"supplier": _TW_SUPPLIER_SPACE, "consignee": consignee, "invoice_no": invoice,
                    "invoice_date": _vn_date(_vn_cell(row, col.get("VATINVDATE"))),
                    "onboard_date": _vn_date(_vn_cell(row, col["ATD"])),
                    "bl_no": bl, "port_of_loading": pol, "port_of_discharge": pod}
            fee_vnd_total = 0
            for c in fee_cols:
                usd = _vn_num(_vn_cell(row, c))
                if not usd:
                    continue
                vnd = int(usd * rate + 0.5)
                fee_vnd_total += vnd
                items.append({**base, "description": f"{fee_names[c]} (USD {_tw_usd_text(usd)})",
                              "amount": vnd, "vat": None})
            sum_total_vnd = _vn_num(_vn_cell(row, col.get("SUMTOTAL(VND)")))
            vat_amt = (sum_total_vnd - fee_vnd_total) if sum_total_vnd is not None else None
            items.append({**base, "description": "VAT 8%", "amount": None, "vat": vat_amt})
            totals.append({"invoice_no": invoice, "total": fee_vnd_total, "grand_total": sum_total_vnd})
    return items, totals


# --- VN01 TRANS WAGON：FREIGHT STATEMENT (EXPORT/IMPORT) (工作表 TONG，2026-11-09 使用者
# 答案) --------------------------------------------------------------------------------
# 一列一票 (BILL No.)：所有 TOTAL(VND) 欄的公式都沒有存值，VND 自己用 USD × XRATE 算。
# D/O、H/D、CIC、CFS、THC、VSC、EQUIP、(出口多一欄 LC、進口多一欄 OTHER) 是有 VAT 的
# Local Charge；A/F、O/F、EXWORK 沒有 VAT。VAT 用每個 Local Charge 各自 ROUND(VND×8%)，
# 不用「總額先四捨五入再分攤尾差」：驗算發現來源公式偶有手動 +1 (例如 08SEA2609077)，
# 逐項各自四捨五入剛好跟來源的手動調整值一致，不需另外處理。
# 進口版有 EXW 工作表：這票的 EXWORK 改用 EXW 明細逐筆列出 (不列 TONG 的 EXWORK 一列)。
# 2026-11-12：出口版多一欄 "LC" (跟 D/O 等同屬 Local Charge，不是 _TW_NOVAT_FEES)，沒加
# 進這個清單時會整欄被跳過、不列入 total/grand_total，但本身自我比對又會「剛好一致」
# 不會被稽核抓到 (因為 total 是用同一份已濾掉 LC 的欄位重新加總，不是讀帳單自己的小計
# 格)；驗算來源 "TOTAL(VND) not VAT Local Charge"/VAT/GRAND TOTAL 三欄，加回 LC 後三欄
# 都剛好對得上 (sample7 9HCMHKG604N3：636704+50936=687640)，才確認 LC 要當 Local Charge。
# 2026-11-12：這批資料月份、船務不同，同一種費用表頭名稱有好幾種寫法 (例如
# "DRAYAGE CHARGE"/"DRAYAGE CHARGES"、"LOADING"/"LOADING CHARGE")，另外出現
# "TCS"、"LOCAL CHARGE"、"OTHERS" (複數，跟舊的 "OTHER" 單數不同欄)。都用帳單
# 自己的 VAT/GRAND TOTAL 稽核驗證過屬於 Local Charge (有 VAT)：
# SBS-80010897：D/O 649250+LOCAL CHARGE 908950=1558200；
# LPEX2502638：D/O 1664000+THC 4608000+LOADING 14336000+DRAYAGE CHARGE
# 6400000=27008000；LPEX2608630：D/O 1438800+THC 6278400+LOADING CHARGE
# 7713014+DRAYAGE CHARGES 9156000=24586214，都剛好等於帳單
# "TOTAL(VND) not VAT Local Charge"/"TOTAL(VND) Local Charge" 欄。
_TW_LOCAL_FEES = ("D/O", "H/D", "CIC", "CFS", "THC", "VSC", "EQUIP", "OTHER", "OTHERS", "LC",
                  "TCS", "LOCAL CHARGE", "DRAYAGE CHARGE", "DRAYAGE CHARGES",
                  "LOADING", "LOADING CHARGE", "INSPECTION FEE")
_TW_NOVAT_FEES = ("A/F", "O/F", "EXWORK")


def _tw_header_first_line(cell) -> str:
    """表頭儲存格第一行 (換行在 _vn_text() 壓成空白前先切開，否則永遠只有一行)。"""
    if not cell:
        return ""
    return re.sub(r"\s+", " ", str(cell).split("\n", 1)[0]).strip()


def _tw_bare_fee_name(cell) -> str:
    """表頭第一行去掉結尾的 "(USD)"，用來比對欄位代號 (例如 "O/F (USD)" -> "O/F")。"""
    return re.sub(r"\s*\(USD\)\s*$", "", _tw_header_first_line(cell), flags=re.I).strip().upper()


def _tw_fee_label(cell) -> str:
    """費用名稱照抄表頭第一行原文 (例如 "D/O"、"O/F (USD)")。"""
    return _tw_header_first_line(cell)


def _tw_header_col(header_row: list, label) -> Optional[int]:
    """label 可以是單一字串，也可以是多個候選字串 (同一欄位各月份標籤不同時用，
    例如出口版 "ETD"、進口版 "Ngay tau den"、舊樣本 "ATD")。比對去聲調、不分大小寫，
    並忽略儲存格結尾的括號註記 (例如 "Ngay tau den (ETA)"、"VAT(8%)"，
    2026-11-12：新批次這個欄位多了 "(ETA)" 後綴，固定比對抓不到)。"""
    labels = {label} if isinstance(label, str) else set(label)
    norm_labels = {_vt_norm(lb) for lb in labels}
    for c, cell in enumerate(header_row):
        norm_cell = _vt_norm(cell)
        norm_cell_bare = re.sub(r"\s*\([^()]*\)\s*$", "", norm_cell).strip()
        if (_tw_header_first_line(cell).upper() in labels or norm_cell in norm_labels
                or norm_cell_bare in norm_labels):
            return c
    return None


def _tw_box_cbm_volume(n20, n40, cbm) -> Optional[str]:
    """材積寫成 "1X40" (使用者 2026-11-10 統一格式，同 APL_VN)；多櫃用 ", " 連接；
    沒有櫃數時用 CBM，例如 "1.296 CBM"。"""
    boxes = [f"{int(n)}X{size}" for n, size in ((n20, "20"), (n40, "40")) if n]
    if boxes:
        return ", ".join(boxes)
    if cbm:
        return f"{cbm:g} CBM"
    return None


def _tw_tong_head(rows: List[list]) -> Optional[int]:
    for i, row in enumerate(rows[:15]):
        names = {_tw_bare_fee_name(c) for c in row}
        if {"D/O", "THC", "EXWORK"} <= names:
            return i
    return None


def _tw_is_tong(book) -> bool:
    return any(_tw_tong_head(rows) is not None for _, rows in book.sheets)


def _tw_tong_fee_cols(header_row: list) -> Dict[str, List[int]]:
    """同一個費用代號可能對到兩欄 (例如 "A/F (USD)" 出現兩次，一欄給 "cuoc VC đường bay"、
    一欄給另一種路線，每張發票通常只有一欄有值)：全部記下來，不要只取第一欄
    (2026-11-12：BK 122025 AIR IMP 這張發票的值在第二欄，取第一欄抓到 0，整筆漏算)。"""
    out: Dict[str, List[int]] = {}
    for c, cell in enumerate(header_row):
        name = _tw_bare_fee_name(cell)
        if name in _TW_LOCAL_FEES or name in _TW_NOVAT_FEES:
            out.setdefault(name, []).append(c)
    return out


def _tw_exw_sheet(rows: List[list]) -> Optional[int]:
    for i, row in enumerate(rows[:5]):
        if any(_tw_key(c) == "EXW/FCADETAIL" for c in row):
            return i
    return None


def _tw_exw_details(book) -> Dict[str, List[Tuple[str, float]]]:
    """進口版 EXW 工作表：{Bill No: [(明細名稱, USD), …]}；沒有這張表時回傳空字典。
    欄位位置用表頭文字找，不固定寫欄號：有些月份 "Bill No" 前面多一欄序號
    (例如 "No:")，固定欄號會整欄錯位抓到空白/文字 (2026-11 批次抽樣)。"""
    details: Dict[str, List[Tuple[str, float]]] = {}
    for _, rows in book.sheets:
        h = _tw_exw_sheet(rows)
        if h is None:
            continue
        header_row = rows[h]
        bill_c_found = _tw_header_col(header_row, "BILL NO")
        bill_c = bill_c_found if bill_c_found is not None else 0
        label_c = next((c for c, cell in enumerate(header_row) if _tw_key(cell) == "EXW/FCADETAIL"), 2)
        usd_c = label_c + 1
        current = None
        for row in rows[h + 1:]:
            bill = _vn_text(_vn_cell(row, bill_c))
            label = _vn_text(_vn_cell(row, label_c))
            usd = _vn_num(_vn_cell(row, usd_c))
            if bill:
                current = bill
            if current and label and usd:
                details.setdefault(current, []).append((label, usd))
    return details


def _tw_tong_sheets(book) -> Tuple[List[Dict], List[Dict]]:
    """「FREIGHT STATEMENT (EXPORT/IMPORT)」版：回傳 (費用列, 每張發票的合計)。"""
    items: List[Dict] = []
    totals: List[Dict] = []
    exw = _tw_exw_details(book)
    seen_bills: set = set()   # 同一活頁簿可能有第二張 TONG 版面的工作表複本 (同一批 Bill No.
    # 多加幾欄查核用的欄位，例如 "Sheet1" 比 "TONG" 多了 "Lot No."/"check" 欄，內容其餘
    # 相同)，兩張都抓會整票金額算兩倍；同一張發票只採用先遇到的那張表 (2026-11-12)。
    for name, rows in book.sheets:
        head = _tw_tong_head(rows)
        if head is None:
            continue
        header_row = rows[head]
        fee_cols = _tw_tong_fee_cols(header_row)
        bill_c = _tw_header_col(header_row, "BILL NO.")
        # 日期欄標籤依月份不同："ATD"(舊樣本)、"ETD"(出口版)、"Ngay tau den"/"ETA"(進口版)、
        # "et"(2026-11-12：APRIL 2026 IMPORT-YUEDE 打字少打一個 D，照抄原文的誤植)
        atd_c = _tw_header_col(header_row, ("ATD", "ETD", "ET", "ETA", "NGAY TAU DEN"))
        rate_c = _tw_header_col(header_row, "XRATE")
        note_c = _tw_header_col(header_row, "NOTE PORT OF LOADING/DISCHARGE")
        # 稽核用帳單自己的 VAT / GRAND TOTAL 欄 (不是逐項重新加總)，才能抓到「表頭
        # 多了一種沒處理過的費用欄位」這種漏算 (2026-11-12：出口版 "LC" 欄漏掉時，
        # 逐項重新加總自己跟自己比對一定一致，抓不到少算了一欄)
        vat_c = _tw_header_col(header_row, "VAT")
        grand_c = _tw_header_col(header_row, "GRAND TOTAL")
        # 表頭完全沒有獨立的 VAT 欄 (例如 POU CHEN 版只有 "TOTAL(VND) Local Charge"、
        # "GRAND TOTAL"，沒有 "VAT" 欄) 代表這批 Local Charge 本身沒有稅
        # (2026-11-12：套固定 8% 算出來的 VAT 跟帳單 GRAND TOTAL 對不起來，差額剛好
        # 等於多算的 VAT，確認這個版面 Local Charge 不含稅)。
        has_vat = vat_c is not None
        sub = rows[head + 1] if head + 1 < len(rows) else []
        col20 = next((c for c, v in enumerate(sub) if _vn_text(v).strip("'").upper() == "20"), None)
        col40 = next((c for c, v in enumerate(sub) if _vn_text(v).strip("'").upper() == "40"), None)
        colcmb = next((c for c, v in enumerate(sub) if _vn_text(v).upper() == "CMB"), None)
        consignee = _tw_header_consignee(rows)
        for row in rows[head + 2:]:
            if _vn_text(_vn_cell(row, 0)).upper() == "TOTAL":
                break
            bill = _vn_text(_vn_cell(row, bill_c))
            if not bill:
                continue
            if bill in seen_bills:
                continue
            seen_bills.add(bill)
            rate = _vn_num(_vn_cell(row, rate_c)) or 0
            note = _vn_text(_vn_cell(row, note_c))
            pol = pod = None
            if " - " in note:
                pol, pod = (p.strip() for p in note.split(" - ", 1))
            date_text = _vn_date(_vn_cell(row, atd_c))
            volume = _tw_box_cbm_volume(_vn_num(_vn_cell(row, col20)), _vn_num(_vn_cell(row, col40)),
                                         _vn_num(_vn_cell(row, colcmb)))
            base = {"supplier": _TW_SUPPLIER_SPACE, "consignee": consignee, "invoice_no": bill,
                    "invoice_date": date_text, "onboard_date": date_text, "bl_no": bill,
                    "port_of_loading": pol, "port_of_discharge": pod, "volume": volume}
            exw_bill = exw.get(bill)
            local_vnd = vat_vnd = novat_vnd = 0
            for fee in _TW_LOCAL_FEES:
                for c in fee_cols.get(fee, []):
                    usd = _vn_num(_vn_cell(row, c))
                    if not usd:
                        continue
                    vnd = int(usd * rate + 0.5)
                    vat = int(vnd * 0.08 + 0.5) if has_vat else 0
                    items.append({**base,
                                  "description": f"{_tw_fee_label(header_row[c])} (USD {_tw_usd_text(usd)})",
                                  "amount": vnd, "vat": vat})
                    local_vnd += vnd
                    vat_vnd += vat
            for fee in _TW_NOVAT_FEES:
                if fee == "EXWORK" and exw_bill:
                    continue   # 這票的 EXWORK 改用 EXW 明細拆列
                for c in fee_cols.get(fee, []):
                    usd = _vn_num(_vn_cell(row, c))
                    if not usd:
                        continue
                    vnd = int(usd * rate + 0.5)
                    items.append({**base,
                                  "description": f"{_tw_fee_label(header_row[c])} (USD {_tw_usd_text(usd)})",
                                  "amount": vnd, "vat": 0})
                    novat_vnd += vnd
            if exw_bill:
                for label, usd in exw_bill:
                    vnd = int(usd * rate + 0.5)
                    items.append({**base, "description": f"{label} (USD {_tw_usd_text(usd)})",
                                  "amount": vnd, "vat": 0})
                    novat_vnd += vnd
            grand_cell = _vn_num(_vn_cell(row, grand_c)) if grand_c is not None else None
            vat_cell = _vn_num(_vn_cell(row, vat_c)) if vat_c is not None else None
            total_cell = (grand_cell - vat_cell) if grand_cell is not None and vat_cell is not None else None
            totals.append({
                "invoice_no": bill,
                "total": total_cell if total_cell is not None else local_vnd + novat_vnd,
                "grand_total": grand_cell if grand_cell is not None else local_vnd + vat_vnd + novat_vnd,
                # 帳單自己的 GRAND TOTAL 偶有幾塊 VND 尾差 (逐欄個別四捨五入後加總，
                # 跟帳單自己的 GRAND TOTAL 公式路徑不同)，容許到 ±5
                # (2026-11-12：10SEA2603104 差 1.04、08SEA2501023 差 4、
                # 11SEA2412066 差 5)。
                "grand_tolerance": 5})
    return items, totals


def _tw_sheet_parse(book) -> List[Dict]:
    # CONG NO / BẢNG KÊ / TONG 版另外處理；其餘版型 (S1 / FREIGHT STATEMENT / TỔNG) 照原規則
    if _tw_is_cong_no(book):
        items, _ = _tw_cong_no_sheets(book)
        return [{"header": {}, "items": items or [{"description": None, "amount": None}]}]
    if _tw_is_bang_ke(book):
        items, _ = _tw_bang_ke_sheets(book)
        return [{"header": {}, "items": items or [{"description": None, "amount": None}]}]
    if _tw_is_tong(book):
        items, _ = _tw_tong_sheets(book)
        return [{"header": {}, "items": items or [{"description": None, "amount": None}]}]
    docs = _tw_vt_parse(book)
    # Excel 版的供應商一律用英文帳單抬頭名稱 (使用者 2026-11-04 規定：輸出英文公司名稱)
    for doc in docs:
        for it in doc["items"]:
            if it.get("supplier"):
                it["supplier"] = _TW_EXCEL_SUPPLIER
    return docs


def _tw_sheet_totals(book) -> List[Dict]:
    if _tw_is_cong_no(book):
        return _tw_cong_no_sheets(book)[1]
    if _tw_is_bang_ke(book):
        return _tw_bang_ke_sheets(book)[1]
    if _tw_is_tong(book):
        return _tw_tong_sheets(book)[1]
    return _tw_vt_totals(book)


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
    # "SOA CONTAINER"：舊字型編碼 (TCVN3) 讀出來變成 "SOÁ"，不是聲調問題，要另外列
    # (使用者 2026-11-09 答案 BK NEO 表)
    "header": ("SO CONTAINER", "SOA CONTAINER"), "key": ["SO LO MOI", "SO LO CU", "=SO LO", "SO CONTAINER"],
    # 發票號碼 = 表頭「Đính kèm hóa đơn số」下方的字軌 (1C26TYY 這種，只取字軌)，2026-11-05 依新答案
    "doc_key": r"\b(\d[A-Z]\d{2}[A-Z]{3})\b",
    "gross_vat": True,
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
    (r"(?!)", {**_VT_VHL_COMMON, "gross_vat": False, "date": "NGAY VAO XUONG", "require": "PHI NEO XE",
               "fees": [("=TONG CONG", "PHÍ NEO XE")], "grand": "=TONG CONG"}),
    # 磅櫃 / 附加費表 (CÂN CONT、cân-phụ thu)
    (r"(?!)", {**_VT_VHL_COMMON, "date": ("=NGAY", "NGAY LAY"), "require": ("PHI CAN CONT", "CHI HO CAN"),
               "fields": {**_VT_VHL_COMMON["fields"], "bl_no": "SO LO"},
               "fees": [(x, None) for x in ("PT CAN CONT", "=PTCC", "PHI CAN CONT", "=PHU THU", "PHU THU CAN",
                                            "CHI HO CAN")],
               # 代墊過磅費已含 VAT：TỔNG 是含稅合計 (跟同一批的運費表合併勾稽)
               "grand": ("=TONG", "=TONG CONG")}),
    # 2026-10-28 抽樣補：散貨 (HÀNG LẺ：TIỀN CƯỚC VC、PHÍ BOT、代墊) 與超大櫃 (CONT QUÁ KHỔ：
    # VAT 欄叫 "VAT 8%"，同列另有代墊吊櫃費 PHÍ NÂNG/HẠ (GỒM VAT)、VSC)
    # name_require：跟 2026-11-09 新答案的簡單版 "hang le" (CƯỚC/PHÍ NÂNG/PHÍ HẠ、
    # 合計欄 "TỔNG CỘNG") 表頭不一樣，要用這張表自己的合計欄名稱排除，不然名稱一樣
    # 會先被這份比對掉，簡單版就進不了下面「靠表頭配對」的 XUẤT 版面
    (r"CUOC VC|HANG LE", {**_VT_VHL_COMMON,
                  "name_require": ("TONG CONG SO TIEN", "TONG TIEN CUOC VAN CHUYEN"),
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
    (r"NEO CONT", {**_VT_VHL_COMMON, "gross_vat": False, "date": None,
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
# 2026-11-05 YUEDE 新答案 (sample1)：每張工作表費用欄固定 (答案提示列)，費用名稱照答案 (寫入時加英文備註)。
# 費用 = 含稅總額，AMOUNT = 總額 ÷ 1.08、VAT = 總額 − AMOUNT (gross_vat)；發票號碼 = 表頭字軌 (doc_key)；
# B/L NO. = SỐ LÔ (批號)。CHI HỘ 工作表沒有發票字軌 (N/A)。
_VT_VHL_ANSWER = [
    # 工作表名稱每個月、每個帳本都不一樣 (DU DUC-NHẬP、6070、X1-1、nhap1…)，一律用表頭
    # 判斷 (require)，不要靠名稱比對；require_all 的欄位是這個版面獨有、不會跟其他版面
    # 混到的欄 (NHẬP 有「TỔNG CỘNG CƯỚC VC」中間小計欄，XUẤT 沒有)。
    #
    # Excel 日期日、月放反只出現在 BK/công nợ/POU HUNG 對帳表 (靠表頭判斷配對的版面)，
    # JANUARY／FEBRUARY-YUEDE (靠工作表名稱 "DU DUC-NHẬP"/"DU DUC -XUẤT" 直接配對) 不會
    # 放反；所以另外複製兩份「只靠表頭配對、永遠不靠名稱」的版本 (名稱寫 "(?!)"：
    # 用查表配對，加 date_dm_swap)，放在最前面優先比對，名稱有符合的工作表 (1月/2月)
    # 一律先被下面「靠名稱」的版本比對掉，不會用到這兩份。
    # BK 的 "6070" 表頭日期欄沒有 "LẤY" 只寫 "NGÀY"，"NGAY LAY" 找不到時退回 "=NGAY"
    (r"(?!)", {**_VT_VHL_COMMON, "key": ["SO LO"], "date": ("NGAY LAY", "=NGAY"), "date_dm_swap": True,
               "require_all": ("TONG CONG CUOC VC",),
               "fields": {**_VT_VHL_COMMON["fields"], "bl_no": "SO LO"},
               # công nợ VHL "6010" 的 CƯỚC 拆成市內/市外兩欄、VSC 併進 "PHÍ D/O VSC"；
               # POU HUNG "xuat" 表 (雖然叫 xuat，但有 TỔNG CỘNG CƯỚC VC 中間小計欄，結構
               # 跟 NHẬP 一樣) 另有「hỗ trợ phí cơ sở hạ tầng」費與「Phí hạ bãi sớm」
               "fees": [("=CUOC VC", "CƯỚC VC"), ("=CUOC VC TRONG TPHCM", "CƯỚC VC TRONG TPHCM"),
                        ("=CUOC VC NGOAI TPHCM", "CƯỚC VC NGOÀI TPHCM"), ("=BD NANG, HA", "BD NÂNG, HẠ"),
                        ("=PHI NANG", "PHÍ NÂNG"), ("=PHI HA", "PHÍ HẠ"), ("=VSC", "VSC"),
                        ("=PHI D/O VSC", "PHÍ D/O VSC"), ("=HO TRO PHI CO SO HA TANG", None),
                        ("=PHI HA BAI SOM", "Phí hạ bãi sớm")],
               "grand": "=TONG CONG"}),
    (r"(?!)", {**_VT_VHL_COMMON, "key": ["SO LO"], "date": "NGAY", "date_dm_swap": True,
               "require": "CUOC", "_xuat_variant": True,
               "fields": {**_VT_VHL_COMMON["fields"], "bl_no": "SO LO"},
               "fees": [("CUOC", None), ("=PHI NANG", "PHÍ NÂNG"), ("=PHI HA", "PHÍ HẠ"),
                        ("=PHI HA SOM", "Phí hạ sớm")],
               "grand": "=TONG CONG"}),
    (r"DU DUC-?NHAP", {**_VT_VHL_COMMON, "key": ["SO LO"], "date": "NGAY LAY",
                       "require_all": ("TONG CONG CUOC VC",),
                       "fields": {**_VT_VHL_COMMON["fields"], "bl_no": "SO LO"},
                       # 一定要用完全比對 "=" (2 月的表頭把代墊發票號欄直接寫成
                       # "SỐ HĐ \nPHÍ NÂNG"，用片語比對會先抓到這欄的字軌號碼當成金額)
                       "fees": [("=CUOC VC", "CƯỚC VC"), ("=BD NANG, HA", "BD NÂNG, HẠ"),
                                ("=PHI NANG", "PHÍ NÂNG"), ("=PHI HA", "PHÍ HẠ"), ("=VSC", "VSC")],
                       "grand": "=TONG CONG"}),
    (r"DU DUC ?-?XUAT", {**_VT_VHL_COMMON, "key": ["SO LO"], "date": "NGAY",
                         "require": "CUOC", "_xuat_variant": True,
                         "fields": {**_VT_VHL_COMMON["fields"], "bl_no": "SO LO"},
                         # 費用名稱照抄表頭原文 (desc=None)：1 月表頭是 "CƯỚC"，2 月是
                         # "CƯỚC VẬN CHUYỂN"，兩種都要照抄，不能寫死成同一個字串。
                         # BD NÂNG HẠ (BDNH) 要不要列，_vhl_cfgs_for_book() 動態補上
                         # "CUOC" 不能用 "=" 完全比對：1 月表頭是 "CƯỚC"，2 月是
                         # "CƯỚC VẬN CHUYỂN"，片語比對才能兩種都抓到
                         "fees": [("CUOC", None), ("=PHI NANG", "PHÍ NÂNG"), ("=PHI HA", "PHÍ HẠ"),
                                  ("=PHI HA SOM", "Phí hạ sớm")],
                         "grand": "=TONG CONG"}),
    (r"CAN CONT", {**_VT_VHL_COMMON, "key": ["SO LO"], "date": "NGAY", "require_all": ("PTCC",),
                   "fields": {**_VT_VHL_COMMON["fields"], "bl_no": "SO LO"},
                   "fees": [("=PTCC", "PTCC"), ("=CAN", "CÂN")], "grand": "=TONG"}),
    (r"NEO-?DD", {**_VT_VHL_COMMON, "gross_vat": False, "vat_of_amount": True, "date": None,
                  "require": "PHI NEO XE", "key": ["SO LO", "SOA LOA"],
                  "const": {"invoice_date": "N/A"},  # 工作表沒有提櫃日，填 N/A (答案提示列)
                  # BK 的 "NEO" 工作表是舊字型編碼 ("SOÁ LÔ"、"SOÁ CONTAINER"、"SOÁ NGÀY NEO"、
                  # "LOAÏI CONT" 讀出來變成 "SOA LOA"、"SOA CONTAINER"、"SOA NGAØY NEO"、
                  # "LOAII CONT")，相關欄位都要多列一個舊編碼寫法
                  "fields": {**_VT_VHL_COMMON["fields"], "bl_no": ("SO LO", "SOA LOA"),
                             "container_no": ("SO CONTAINER", "SOA CONTAINER")},
                  "boxes": {"20": ("LOAI CONT 20'", "LOAII CONT 20'"), "40": ("LOAI CONT 40'", "LOAII CONT 40'"),
                            "45": ("LOAI CONT 45'", "LOAII CONT 45'")},
                  # 英文備註位置在 "PHÍ NEO XE" 後面、天數前面 (使用者指定的固定格式)，
                  # 跟其他費用的備註加在最後面不一樣，寫入層的 english_note() 沒辦法產生
                  # 這種中間插入的格式，所以直接在這裡寫死 (_EN_PHRASES 另有 "PHI NEO XE"
                  # 片語對照，讓 english_note() 認得這裡已經加過備註，不會重複加)
                  "fees": [("=TONG CONG", "PHÍ NEO XE (Detention Fee-Truck) x {days} ngày")],
                  # 舊編碼 "SOÁ NGÀY NEO" 讀出來的 "Ø" 不是一般的 "O"，用結尾 "Y NEO" 來配對
                  "days_label": ("NGAY NEO", "Y NEO"),
                  "grand": "=TONG CONG"}),
    (r"CHI HO", {**_VT_VHL_COMMON, "date": "NGAY LAY", "doc_const": "N/A",
                 "fees": [("CUOC VC", "CƯỚC VC")], "grand": None}),
]
_VT_VINH_HUNG_LONG[0:0] = _VT_VHL_ANSWER
_VT_VINH_HUNG_LONG.insert(0, (r"(?!)", {
    **_VT_VINH_HUNG_LONG[0][1], "require_all": ("TONG CONG HD NANG HA", "=TONG CONG", "=CAN CONT"),
    "fees": _VT_VINH_HUNG_LONG[0][1]["fees"] + [("=CAN CONT", None)]}))


def _vhl_xuat_sheet(rows: List[list]) -> bool:
    """是不是「XUẤT (出口)」版面：有 SỐ CONTAINER + CƯỚC，但沒有 NHẬP 版面獨有的
    「TỔNG CỘNG CƯỚC VC」中間小計欄。"""
    h = _vt_header_row(rows, "SO CONTAINER")
    if h is None:
        return False
    labels = _vt_labels(rows, h, True)
    return _vt_col(labels, "CUOC") is not None and _vt_col(labels, "=TONG CONG CUOC VC") is None


def _vhl_xuat_has_bdnh(rows: List[list]) -> bool:
    """「XUẤT」版面的 BDNH (BD NÂNG HẠ) 欄要不要列進費用，看每櫃小計公式有沒有
    算進去 (使用者 2026-11-09 確認：1 月 TỔNG CỘNG = CƯỚC + 吊櫃合計，沒有算 BDNH；
    2 月 TỔNG CỘNG = CƯỚC + BDNH + 吊櫃合計，有算)。用金額關係判斷，不用月份硬寫，
    之後月份版面只要公式邏輯一樣就會自動判斷正確。"""
    h = _vt_header_row(rows, "SO CONTAINER")
    if h is None:
        return False
    labels = _vt_labels(rows, h, True)
    c_cuoc, c_bdnh = _vt_col(labels, "CUOC"), _vt_col(labels, "=BDNH", "=BD NANG HA")
    c_hd, c_grand = _vt_col(labels, "=TONG CONG HD NANG HA"), _vt_col(labels, "=TONG CONG")
    if c_cuoc is None or c_bdnh is None or c_hd is None or c_grand is None:
        return False
    for row in rows[h + 2:h + 30]:
        cuoc, bdnh = _vn_num(_vn_cell(row, c_cuoc)), _vn_num(_vn_cell(row, c_bdnh))
        hd, grand = _vn_num(_vn_cell(row, c_hd)), _vn_num(_vn_cell(row, c_grand))
        if cuoc is None or grand is None:
            continue
        return abs(grand - (cuoc + (bdnh or 0) + (hd or 0))) < 1
    return False


def _vhl_cfgs_for_book(book) -> List[Tuple[str, Dict]]:
    """「XUẤT」版面的費用設定 (_xuat_variant 標記的兩份：靠名稱配對、靠表頭配對)
    動態決定要不要加 BD NÂNG HẠ，依這本活頁簿裡任一張 XUẤT 版面工作表的金額關係判斷。"""
    cfgs = list(_VT_VINH_HUNG_LONG)
    has_bdnh = any(_vhl_xuat_sheet(rows) and _vhl_xuat_has_bdnh(rows) for _, rows in book.sheets)
    if not has_bdnh:
        return cfgs
    return [(pat, {**cfg, "fees": cfg["fees"] + [(("=BDNH", "=BD NANG HA"), "BD NÂNG HẠ")]})
            if cfg.get("_xuat_variant") else (pat, cfg) for pat, cfg in cfgs]


def _vhl_sheet_parse(book) -> List[Dict]:
    parse, _ = _vt_make("VINH HUNG LONG CO., LTD", _vhl_cfgs_for_book(book))
    docs = parse(book)
    for doc in docs:
        for it in doc["items"]:
            # SỐ LÔ 有時候寫的不是批號 (使用者指定)：KIEN GIANG 是地名 -> B/L NO. 不填，
            # 改填目的港；HANG ROI (散貨) 只是標記 -> B/L NO. 不填
            if it.get("bl_no") == "KIEN GIANG":
                it["bl_no"], it["port_of_discharge"] = None, "KIEN GIANG"
            elif it.get("bl_no") == "HANG ROI":
                it["bl_no"] = None
    return docs


def _vhl_sheet_totals(book) -> List[Dict]:
    _, totals = _vt_make("VINH HUNG LONG CO., LTD", _vhl_cfgs_for_book(book))
    return totals(book)


register_supplier(
    "VINH_HUNG_LONG", "CÔNG TY TNHH VĨNH HƯNG LONG", detect_vinh_hung_long, None,
    sheet_parse=_vhl_sheet_parse, sheet_totals=_vhl_sheet_totals,
)


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
    # 2026-11-04 樣本答案 (IMP/EXP 兩張表)：發票號碼 = PAY NO (PFQ-…)，提單 = LOT NO / DEC NO，
    # 費用名稱照表頭 (Price、2%)，VAT 列也用同一個 PAY NO
    (r"^(IMP|EXP)$", {"header": "PAY NO", "two_rows": False, "key": ["PAY NO"], "name_require": "=PRICE",
                      "fees": [("MISCELL", None), ("=OT.", None), ("=OTHER", None),
                               ("=PRICE", None), ("=2%", "2%")], "rest_fee": "2%",
                      "total": "=TOTAL", "vat_row": True, "vat_last": True, "vat_no_date": True,
                      "date": "DEC DATE",
                      "fields": {"order_no": "=INV", "bl_no": ("LOT NO", "DEC NO")}, "cbm": "TOTAL CFS/AIR",
                      "volume_raw": ("TOTAL CON", "TOTAL CFS/AIR")}),
    # NB / PHV (2026-11-05 樣本答案，兩張表格式相同)：費用名稱帶英文說明 (寫入前在解析時加上)，FUEL SURCHARGE 的百分比
    # 取 "TỶ LỆ TĂNG GIÁ…" 欄 (0.16 -> 16%)；VAT 列名稱 "Thuế VAT (8%) (VAT 8%)"
    (r"^(NB|PHV)$", {"header": "SO LO", "key": ["=SO LO"],
               "fees": [("CUOC VC CHUA VAT", "CƯỚC VC CHƯA VAT (TRUCKING FREIGHT)"),
                        ("SO TIEN PHU THU", "SỐ TIỀN PHỤ THU (FUEL SURCHARGE {pct}%)")],
               "pct_col": "TY LE TANG GIA",
               "vat_label": "Thuế VAT (8%) (VAT 8%)", "vat_last": True,
               "total": "TONG TIEN THANH TOAN", "vat_row": True, "date": "=NGAY",
               "cbm": "SO LUONG CBM",
               # 2026-11-04 樣本答案：發票號碼 = 抬頭 "Đính kèm hóa đơn số" (整張一個)，SỐ LÔ 放 B/L NO.，
               # 船名與貨櫃號碼留空
               "doc_key": r"(\d[A-Z]\d{2}[A-Z]{3} \d{3,5})", "merge_no_fee": True, "vat_no_date": True,
               "fields": {"order_no": "SO INV", "bl_no": "=SO LO"}}),
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
    (r"CFS", {"header": "SO LO", "two_rows": False, "key": ["SO LO"], "name_require": "ST/1 LO",
                  # 2026-11-05 樣本答案 (08-CFS)：費用名稱照表頭 + 英文說明；SỐ LÔ 同時當發票號碼與 B/L NO.；
                  # 船名與貨櫃號碼留空
                  "fees": [("ST/1 LO CHUA VAT", "ST/1 lô chưa VAT (TRUCKING FREIGHT)"),
                           ("PHU PHI XANG DAU", "PHỤ PHÍ XĂNG DẦU (16%) (FUEL SURCHARGE)")],
                  "total": "TONG TIEN/1 LO", "vat": "VAT 8%", "vat_desc": "VAT 8%",
                  "grand": ("=TONG CONG", "=TOTAL"), "date": "NGAY NHAN HANG", "ffill_date": True,
                  "cbm": "CMB", "kg": "=KG", "cbm_always": True,
                  "fields": {"bl_no": "=SO LO", "port_of_loading": "=CANG"}}),
    (r"AIR", {"header": "SO LO", "two_rows": False, "key": ["SO LO"], "name_require": "ST/1 LO",
                  # 2026-11-05 樣本答案 (08-CFS)：費用名稱照表頭 + 英文說明；SỐ LÔ 同時當發票號碼與 B/L NO.；
                  # 船名與貨櫃號碼留空
                  "fees": [("ST/1 LO CHUA VAT", "ST/1 lô chưa VAT (TRUCKING FREIGHT)"),
                           ("PHU PHI XANG DAU", "PHỤ PHÍ XĂNG DẦU (16%) (FUEL SURCHARGE)")],
                  "total": "TONG TIEN/1 LO", "vat": "VAT 8%", "vat_desc": "VAT 8%",
                  "grand": ("=TONG CONG", "=TOTAL"), "date": "NGAY NHAN HANG", "ffill_date": True,
                  "cbm": "CMB", "kg": "=KG", "kg_always": True,
                  "fields": {"bl_no": "SO LO", "port_of_loading": "=CANG"}}),
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
             # 供應商欄用英文名稱 (樣本答案，2026-11-04)
             "HOP TAC XA VAN TAI DICH VU DU LICH BINH TAN", _VT_BINH_TAN)


# --- VN12 NGUYÊN HƯỞNG (空運請款) ---------------------------------------------
# 每項費用都有 RMB 與 VND 兩欄，只取 VND；Description 接上 RMB 原幣金額 (使用者
# 2026-11-10 新答案："air freight (RMB 6,408)"、"Customs fees (RMB 150 / bill)")。
# 舊版 (T01-2025) 比新版少 1 欄 (沒有城市欄)，用 "SAP" 欄位區分版型 (name_require)，
# 一定要排在新版前面：舊版沒有「ĐÍNH KÈM HĐ SỐ」字軌，INVOICE NO 改用 Pickup Date
# (yyyy/mm/dd)；Port of Loading 填 N/A (沒有城市欄，不可套新版寫死的欄號 2)。
_VT_NGUYEN_HUONG_FEES = [
    ("AIR FREIGHT VND", "air freight ({rmb})"), ("CUSTOMS FEES (BILL) VND", "Customs fees ({rmb} / bill)"),
    ("BILL OF LADING FEE (BILL) VND", "Bill of lading fee ({rmb} / bill)"),
    ("ADDITIONAL FEE (BILL) VND", "Additional fee ({rmb} / bill)"),
    ("WAREHOUSING FEE (BILL) VND", "Warehousing fee ({rmb} / bill)"),
    ("PAD EXPENSE VND", "Pad expense ({rmb})"), ("D/O (BILL) VND", "D/O (VND / bill)"),
]
_VT_NGUYEN_HUONG = [
    # 只讀「空运请款」：sap、cps 是空白舊模板 (合計列都是 0，不會真的匯出資料，但
    # 名稱比對限制到這張也比較保險)；Sheet1 有時整張複製「空运请款」(例如
    # T3-2026)，但沒有表頭的「ĐÍNH KÈM HĐ SỐ」字軌，發票號碼會變成逐票的
    # Waybill number，跟「空运请款」(整張一個字軌) 比對不出內容相同而被誤判成
    # 兩份不同資料 (使用者 2026-11-10 答案：Sheet1 內容相同只能算 1 份)。
    (r"空运请款", {"header": "WAYBILL", "key": ["WAYBILL"], "name_require": "SAP",
            "date_key": True, "vat_zero_desc": ("AIR FREIGHT",),
            "vat_total": "VAT 8%", "vat_desc": "VAT 8%", "vat_last": True,
            # 舊版 VAT 列連 INVOICE NO 也填 N/A (跟新版不同，新版用字軌)，2026-11-10 答案
            "vat_na_fields": ("invoice_no", "invoice_date", "bl_no", "order_no", "volume",
                              "port_of_loading", "port_of_discharge", "vessel", "container_no",
                              "shipper", "onboard_date", "arrive_date"),
            "fees": _VT_NGUYEN_HUONG_FEES, "const": {"port_of_loading": "N/A"},
            "total": "TOTAL AMOUNT VND", "date": "PICKUP DATE", "kg": "WEIGHT",
            "fields": {"bl_no": "WAYBILL", "shipper": "=SHIPPER"}}),
    (r"空运请款", {"header": "WAYBILL", "key": ["WAYBILL"],
            # 發票號碼 = 表頭「ĐÍNH KÈM HĐ SỐ」的字軌 (1C26TNH-，只取字軌) (VN12 新答案，2026-11-06)
            "doc_key": r"(\d[A-Z]\d{2}[A-Z]{3})-",
            "vat_zero_desc": ("AIR FREIGHT",),
            # 表尾整張 VAT 8% 列：發票號碼 = 字軌，日期/提單等填 N/A (2026-11-06 新答案)
            "vat_total": "VAT 8%", "vat_desc": "VAT 8%", "vat_no_date": True, "vat_last": True,
            "fees": _VT_NGUYEN_HUONG_FEES,
            "total": "TOTAL AMOUNT VND", "date": "PICKUP DATE", "kg": "WEIGHT",
            # port_of_loading 欄位 (城市) 沒有表頭文字 (合併儲存格留白)，只能寫死欄號；
            # 新版固定排在 Waybill number 前一欄 (使用者 2026-11-06 答案)
            "fields": {"bl_no": "WAYBILL", "shipper": "=SHIPPER", "port_of_loading": 2}}),
]


def detect_nguyen_huong(text: str) -> bool:
    return "NGUYEN HUONG" in _vt_norm(text)


_vt_register("NGUYEN_HUONG", "CÔNG TY TNHH NGUYÊN HƯỞNG", detect_nguyen_huong,
             "NGUYEN HUONG COMPANY LTD.", _VT_NGUYEN_HUONG)


# --- VN14 ORIENTAL LOGISTICS MULTIMODAL (Công Ty TNHH Logistics Vận Tải Đa Phương Thức) --
# 5 種 Excel 對帳表 (CN 進口、IMP 進口、TRUCK、SEA 出口、SEA (2) 出口)。費用欄是 USD，
# 帳單本身另有每張發票的 SỐ TIỀN TRƯỚC THUẾ (未稅)、VAT、TỔNG TIỀN (含稅，VND)、
# EXCHAGRE RATE (匯率) (IMP 版沒有未稅小計欄)。AMOUNT 改成越南盾 (2026-11-10 使用者答案)：
#  - 有未稅小計欄的版型 (CN/TRUCK/SEA/SEA (2))：正常情況依 USD 比例分攤未稅小計、尾差
#    放最後一列；小計跟目前列出的費用 (TOTAL (USD) 欄) 對不上時 (帳單本身的問題，例如少了
#    某筆費用的附件)，改用「USD (這幾家含 VAT，先 ÷1.08) × 匯率」各自換算，不勉強湊出
#    小計，讓金額稽核自然抓出不一致 (BANGKE EXP C26TOD-21469)。
#  - IMP 版沒有未稅小計欄：直接 USD × 匯率，VAT = 含稅總額 (TỔNG TIỀN) − 换算後費用加總。
# 發票號碼 = SỐ HD，去掉所有空白。VAT 欄 0 也要留一列 ("VAT 0%")。報價比較工作表
# (Compare quotation) 不讀。
_VT_OLM_CN_FEES = [(x, None) for x in (
    "OCEAN FREIGHT", "OTHER CHARGES", "CFS FEE", "=CIC", "D/O FEE", "=HANDLING FEE", "CLEANING FEE",
    "=THC", "=CMC", "=LSS", "=LCC", "=BAF", "AFR AMENDMENT", "=HANDLING",
    # 空運進口 (2026-10-28 抽樣 "POU HUNG 06-2026 PHƯƠNG ĐÔNG air import")
    "=AIR FREIGHT")]
_VT_OLM_SEA_FEES = [(x, None) for x in (
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
    "=EFS")]
_VT_ORIENTAL = [
    # 2024 PAYMENT FOR LOCAL CHARGE：Invoice No | B/L | 各項費用 (USD) | Total | VAT | Sum Total (USD/VND)
    (r"(?!)", {"header": "SUM TOTAL", "two_rows": False, "require": "SUM TOTAL", "key": ["INVOICE NO", "=B/L"],
               "fees": [], "fees_between": ("=B/L", ("=TOTAL",)), "fee_cur": "USD",
               "vat": "=VAT", "vat_desc": "VAT (USD)", "total": "=TOTAL", "grand": "SUM TOTAL (USD)",
               "date": "VAT INV DATE", "fields": {"bl_no": "=B/L"}}),
    # 每月發票清單 (ORIENTEL 2026-06：No | Invoice date | Invoice No | Bill No | Amount)
    (r"(?!)", {"header": "BILL NO", "two_rows": False, "require": "=AMOUNT", "key": ["INVOICE NO"],
               "fees": [("=AMOUNT", "Amount")], "date": "INVOICE DATE",
               "sum_row": r"^TOTAL", "fields": {"bl_no": "BILL NO"}}),
    # BANGKE IMP (2026-11-10 答案，工作表 "T5-2026")：沒有未稅小計欄，用 BOOKING NO 區分。
    # 三列表頭 ("VOL (CBM)" 合併格佔滿整列)：費用欄的名稱其實是在表頭下一列 (D/O、THC…)，
    # 跟兩列表頭的供應商一樣用兩列合併比對時，每個費用欄名前面都會多出被合併儲存格帶
    # 進來的 "VOL (CBM) " (兩列表頭合併的結果)，改用 fees_re 吃掉這個固定前綴。CY QTY
    # 20'/40' 兩欄因為同一個理由沒有子標題可分，兩欄名稱相同 ("VOL (CBM) CY QTY")，
    # 用第一個/最後一個符合取代。
    (r"(?!)", {"header": "H-B/L", "require": "BOOKING NO", "key": ["=SO HD"], "key_strip_space": True,
               "carry_fields": ("bl_no",), "carry_volume": True,
               "fees": [], "fees_re": r"^VOL \(CBM\) (?!CY QTY)(.+)$", "fees_skip": r"CHECK",
               "vnd_rate": {"rate": "EXCHAGRE RATE", "grand": "TONG TIEN (VND)"},
               "grand": "TONG TIEN (VND)", "date": "NGAY HD",
               "boxes_adjacent": {"base": "VOL (CBM) CY QTY", "20": 0, "40": 1},
               "fields": {"bl_no": "H-B/L", "mbl_no": "M-B/L", "order_no": "BOOKING NO"}}),
    (r"^CN", {"header": "BILL NO", "above": True, "key": ["SO HD", "BILL NO"], "fees": _VT_OLM_CN_FEES,
              "key_strip_space": True, "ffill_date": True,   # 同一批第二張發票的 NGÀY HĐ 合併在上一列
              # Ocean Freight 那張發票 (同一個 Bill 的第 2 列) 沒有自己的 Bill No/LOT NO/
              # ORIGIN PORT/材積，沿用上一列 (答案 2026-11-10)
              "carry_fields": ("bl_no", "order_no", "port_of_loading"), "carry_volume": True,
              "vnd_alloc": {"pretax": "SO TIEN TRUOC THUE", "rate": "EXCHAGRE RATE",
                            "usd_total": "=TOTAL (USD)", "vat_included": False},
              "total": "SO TIEN TRUOC THUE", "grand": "TONG TIEN (VND)",
              "vat": "=VAT", "vat_pct_desc": True, "vat_zero_ok": True, "date": "NGAY HD",
              "boxes": {"20": "CY QTY 20'", "40": "CY QTY 40'"}, "cbm": "CFS CBM",
              "fields": {"bl_no": "BILL NO", "order_no": "LOT NO", "port_of_loading": "ORIGIN PORT"}}),
    (r"TRUCK", {"header": "H-B/L", "key": ["SO HD", "H-B/L", "M-B/L"], "key_strip_space": True,
                "fees": [(x, None) for x in ("TRUCKING FEE", "BORDER GATE FEE", "CUSTOMS CLEARANCE FEE",
                                             "TOLL FEE", "LOLO FEE", "FUEL SURCHARGE", "FCR TRUCKING")],
                "vnd_alloc": {"pretax": "SO TIEN TRUOC THUE", "rate": "EXCHAGRE RATE",
                              "usd_total": "$=TOTAL USD", "vat_included": True},
                "total": "SO TIEN TRUOC THUE", "grand": "TONG TIEN (VND)",
                "vat": "=VAT", "vat_pct_desc": True, "vat_zero_ok": True,
                "date": "NGAY HD", "boxes": {"20": "CY QUANTITY 20", "40": "CY QUANTITY 40"}, "cbm": "CFS CBM",
                "volume_both": True,
                "fields": {"bl_no": "H-B/L", "mbl_no": "M-B/L", "order_no": "LOT NO",
                           "port_of_loading": "ORIGIN PORT"}}),
    (r"SEA", {"header": "H-B/L", "key": ["SO HD", "H-B/L", "M-B/L"], "key_strip_space": True,
              "skip_sheet": r"COMPARE", "volume_both": True,
              "fees": _VT_OLM_SEA_FEES,
              "vnd_alloc": {"pretax": "SO TIEN TRUOC THUE", "rate": "EXCHAGRE RATE",
                            "usd_total": "$=TOTAL USD", "vat_included": True},
              "total": "SO TIEN TRUOC THUE", "grand": "TONG TIEN (VND)",
              "vat": "=VAT", "vat_pct_desc": True, "vat_zero_ok": True,
              "date": "NGAY HD", "boxes": {"20": "CY QUANTITY 20", "40": "CY QUANTITY 40"}, "cbm": "CFS CBM",
              "port": "NOTE PORT", "port_arrow": True,
              "fields": {"bl_no": "H-B/L", "mbl_no": "M-B/L", "order_no": "LOT NO",
                         "onboard_date": "=DATE"}}),
]


def detect_oriental_multimodal(text: str) -> bool:
    norm = _vt_norm(text)
    if "DEBIT NOTE" in norm:
        return False    # 美金 DEBIT NOTE 是獨立供應商 VN14_USD，見 detect_oriental_usd()
    # 2026-10-28 抽樣：有的對帳表只寫越南文名稱 (… ĐA PHƯƠNG THỨC PHƯƠNG ĐÔNG) 或只有發票符號 C26TOD
    return ("ORIENTAL LOGISTICS MULTIMODA" in norm or "DA PHUONG THUC PHUONG DONG" in norm
            or "VAN TAI DA PHUONG THUC" in norm or re.search(r"\d-?C\d{2}TOD\b", norm) is not None)


# PDF VAT INVOICE (HÓA ĐƠN GIÁ TRỊ GIA TĂNG, FAST e-invoice, Serial C26TOD，2026-11-10 答案)：
# 費用行 "1 Cước vận tải quốc tế (O/F) XMNSGN26050461 2 40'HQ 17.076.150 34.152.300"，
# 表尾 Sub total / VAT / Total payment。VAT 只在合計、沒有逐筆 VAT -> 最後加 1 列 VAT。
_OLM_EINV_ITEM_RE = re.compile(
    r"^\d+\s+(?P<desc>.+?)\s+(?P<bl>\S*\d{6,}\S*)\s+(?P<qty>[\d,]+)\s+(?P<unit>\S+)\s+"
    r"(?P<price>[\d.]+)\s+(?P<amt>[\d.]+)$", re.MULTILINE)


def _olm_num(text: str) -> float:
    return float(text.replace(",", ""))


def parse_oriental_pdf(text: str) -> Dict:
    h: Dict[str, Optional[str]] = {"supplier": "ORIENTAL LOGISTICS MULTIMODAL TRANSPORT COMPANY LIMITED."}
    items: List[Dict] = []
    # INVOICE NO = Mẫu số + "-" + Ký hiệu + "-" + Số (例 1-C25TOD-699，使用者 2026-11-10 確認)
    mau = _search(r"\(Form No\.\):\s*(\S+)", text)
    hieu = _search(r"\(Serial No\.\):\s*(\S+)", text)
    so = _search(r"S[ốo] \(Invoice No\.\):\s*(\d+)", text)
    h["invoice_no"] = "-".join(x for x in (mau, hieu, so) if x) or so
    h["invoice_date"] = _vn_pdf_date(_search(r"\(Date\):\s*(\S+)", text))
    h["consignee"] = _vn_full_consignee(_search(r"\(Company name\):\s*(.+)", text))
    m = re.search(r"HÀNH TRÌNH:\s*(.+?)\s+-\s+(.+?)\s*/\s*\n?\s*(.+)", text)
    if m:
        h["port_of_loading"], h["port_of_discharge"], h["vessel"] = (x.strip() for x in m.groups())
    cbm = 0.0
    for m in _OLM_EINV_ITEM_RE.finditer(text):
        h.setdefault("bl_no", m.group("bl"))
        # 數量單位是 KHỐI (= CBM，「,」是小數)；單位 LÔ (批次) 不計入材積 (2026-11-10 答案)
        if _vt_norm(m.group("unit")).startswith("KH"):
            cbm = max(cbm, float(m.group("qty").replace(",", ".")))
        items.append({"description": m.group("desc").strip(),
                      "amount": _dsv_vn_dot_amount(m.group("amt")), "vat": None})
    vat = _dsv_vn_dot_amount(_search(r"\(VAT amount\):\s*([\d.]+)", text))
    rate = _search(r"\(Tax rate\):\s*(\d+%)", text)
    if vat:
        # 整張只有一筆合計 VAT，不是某張提單的費用：B/L NO.、Volume 不沿用表頭 (答案 2026-11-10)
        items.append({"description": f"VAT {rate}" if rate else "VAT", "amount": None, "vat": vat,
                      "bl_no": "N/A", "volume": "N/A"})
    h["volume"] = f"{cbm:g} CBM" if cbm else None
    return {"header": h, "items": items or [{"description": None, "amount": None, "vat": None}]}


def _olm_pdf_total(text: str):
    return _dsv_vn_dot_amount(_search(r"\(Sub total\):\s*([\d.]+)", text))


def _olm_pdf_grand(text: str):
    return _dsv_vn_dot_amount(_search(r"\(Total payment\):\s*([\d.]+)", text))


# --- VN14_USD ORIENTAL LOGISTICS (美金 DEBIT NOTE，獨立供應商，使用者 2026-11-10 指定) ----
# 帳單沒有匯率和越南盾，金額全部是 USD。費用行 "THC 9.000 / CBM x 5.400 8 USD 52.49"
# (單價/單位 × 數量、稅率、幣別、含稅金額)，未稅 = 單價 × 數量，VAT = 差額 (同一列)；
# Description 照抄「費用名稱 + 單價/單位 x 數量」；表尾 "AMOUNT: USD 238.14" 是含稅合計。
def detect_oriental_usd(text: str) -> bool:
    norm = _vt_norm(text)
    return "DEBIT NOTE" in norm and (
        "ORIENTAL LOGISTICS MULTIMODA" in norm or "DA PHUONG THUC PHUONG DONG" in norm
        or "VAN TAI DA PHUONG THUC" in norm or re.search(r"\d-?C\d{2}TOD\b", norm) is not None)


_OLM_USD_ITEM_RE = re.compile(
    r"^(?P<desc>.+?)\s+(?P<price>[\d.,]+)\s*/\s*(?P<unit>\S+)\s+x\s+(?P<qty>[\d.,]+)\s+(?P<rate>\d+)\s+"
    r"(?P<cur>[A-Z]{3})\s+(?P<amt>[\d.,]+)$", re.MULTILINE)


def parse_oriental_usd_pdf(text: str) -> Dict:
    h: Dict[str, Optional[str]] = {"supplier": "ORIENTAL LOGISTICS MULTIMODAL TRANSPORT COMPANY LIMITED."}
    h["invoice_no"] = _search(r"REF NO\.\s*:\s*(\S+)", text)
    h["invoice_date"] = _vn_pdf_date(_search(r"ISSUE DATE\s*:\s*(\S+)", text))
    h["consignee"] = _vn_full_consignee(_search(r"NAME & ADDRESS:[^\n]*\n(.+?)\s+ISSUE DATE", text))
    h["vessel"] = _search(r"VESSEL/VOYAGE\s*:\s*(.+?)\s+ONBOARD DATE", text)
    h["onboard_date"] = _vn_pdf_date(_search(r"ONBOARD DATE\s*:\s*(\S+)", text))
    # 啟運港/目的港這幾個字被 pdfplumber 切成三行 ("PORT OF xxx, PORT OF" /
    # ": : yyy" / "LOADING/ETD 國名/日期 DISCHARGE/ETA")，國名要接回啟運港後面
    m = re.search(r"PORT OF\s+(.+?),\s*PORT OF\s*:\s*:\s*(.+?)\s*LOADING/ETD\s+(.+?)/\d{1,2}/\d{1,2}/\d{4}",
                  text, re.DOTALL)
    if m:
        h["port_of_loading"] = f"{m.group(1).strip()}, {m.group(3).strip()}"
        h["port_of_discharge"] = m.group(2).strip()
    h["mbl_no"] = _search(r"MBL NO\s*:\s*(?!DESTINATION)(\S+)", text)   # MBL 空白時後面是 DESTINATION
    h["bl_no"] = _search(r"HBL INFO:\s*(\S+)", text)
    h["container_no"] = _search(r"CONTAINER:\s*([A-Z]{4}\d{7})", text)
    items: List[Dict] = []
    units = set()
    for m in _OLM_USD_ITEM_RE.finditer(text):
        price, qty = m.group("price"), m.group("qty")
        net = round(_olm_num(price) * _olm_num(qty), 2)
        amt = _olm_num(m.group("amt"))
        units.add(m.group("unit").upper())
        items.append({"description": f"{m.group('desc').strip()} (USD {price} / {m.group('unit')} x {qty})",
                      "amount": net, "vat": round(amt - net, 2) or None})
    # 材積：按 CBM 計價的是散貨 (CBM)，按櫃計價的是整櫃 (櫃型)
    cbm = _search(r"HBL INFO:.*?([\d.]+)\s*CBM", text)
    cbm = f"{float(cbm):g}" if cbm else None
    boxes = re.findall(r"/(\d{2}'[A-Z]{2})\s", text) if "CBM" not in units else []
    h["volume"] = _fmt_volume(containers=[(1, b) for b in boxes], cbm=cbm)
    return {"header": h, "items": items or [{"description": None, "amount": None, "vat": None}]}


def _olm_usd_pdf_grand(text: str):
    v = _search(r"^AMOUNT:\s*[A-Z]{3}\s*([\d.,]+)", text)
    return _olm_num(v) if v else None


register_supplier(
    "ORIENTAL_USD", "ORIENTAL LOGISTICS (DEBIT NOTE, USD)", detect_oriental_usd, parse_oriental_usd_pdf,
    extract_grand_total=_olm_usd_pdf_grand,
)


# 工作表叫 "CN T 03" 但版面是 SEA 對帳表 (表頭 H-B/L，沒有 BILL NO)：放在 CN 設定前面，
# 用必要欄位區分 (2026-10-28 抽樣 "POU HUNG 03-2026 PHƯƠNG ĐÔNG export")。"name_require"
# 才是名稱已符合時還要檢查的必要欄位 ("require" 只給名稱沒對到時的表頭比對用)；否則一般
# "CN T9.2026" 這種真的是 CN 版面的工作表，會被這個放在前面的設定誤先攔截 (2026-11-10)
_VT_ORIENTAL.insert(next(i for i, (p, _c) in enumerate(_VT_ORIENTAL) if p == r"^CN"),
                    (r"^CN", {**next(c for p, c in _VT_ORIENTAL if p == r"SEA"), "name_require": "H-B/L"}))
_olm_sheet_parse, _olm_sheet_totals = _vt_make("ORIENTAL LOGISTICS MULTIMODAL TRANSPORT COMPANY LIMITED.", _VT_ORIENTAL)
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
                 "total": "=TONG", "vat_row": True, "vat_last": True, "vat_amount_last": True, "pct_col": "% TANG", "key_comma": True,
                 "fields": {"vessel": "SO XE", "container_no": ("SO CONTAINER", "SO XE"),
                            "port_of_loading": ("DIEM DI", "DIEM DI"), "port_of_discharge": "DIEM DEN"}}
_VT_SONG_TRIEU = [
    # 新版 (2026 年起)：有 SO XE/DIEM DI/DIEM DEN/% TANG 欄。name_require 擋住
    # 舊版 (2025 年，沒有 SO XE)，才會往下試舊版設定 (2026-11-11 抽樣 T03.2025)
    (r"XE TAI", {**_VT_ST_COMMON, "header": "SO XE", "key": ["SO INV"], "order_is_key": True,
                 "name_require": "SO XE",
                 "fees": [("PHI BOC", None), ("NEO XE", None), ("PHAT SINH", None),
                          ("CUOC VAN CHUYEN CHUA TANG", None), ("SO TIEN % TANG", "SỐ TIỀN % TĂNG {pct}%")]}),
    # 舊版：沒有車號/起迄點/漲價欄，運費就是最終金額，沒有 TỔNG 欄可以勾稽
    # (2026-11-11 抽樣 T03.2025 "Hang Xe Tai")
    (r"XE TAI", {**_VT_ST_COMMON, "header": "SO INV", "key": ["SO INV"], "order_is_key": True,
                 "total": None,
                 "fees": [("PHI BOC", None), ("PHAT SINH", None), ("CUOC VAN CHUYEN", None)]}),
    (r"CHUYEN KHO", {**_VT_ST_COMMON, "header": "SO CONTAINER", "key": ["SO INV", "SO CONTAINER"], "order_is_key": True,
                     "total_skip_multi_key": True,
                     "fees": [("TROUBLE", None), ("=NEO", None), ("CUOC VAN CHUYEN CHUA TANG", None),
                              ("SO TIEN % TANG", "SỐ TIỀN % TĂNG {pct}%")]}),
    (r"CONT (XUAT|NHAP)", {**_VT_ST_COMMON, "header": "SO CONTAINER", "key": ["=SO LO", "SO CONTAINER"], "bl_is_key": True,
                           # total_check / TOTAL PLV CHECK：帳單自己核對明細加總用的欄 (2026
                           # 年版才有，舊版 2025 樣本沒有這欄)；跟 TỔNG 常常各自缺漏、對不上，
                           # 符合任一欄就算通過 (2026-11-11 使用者確認，見 _vt_parse_sheet)
                           "total_check": "TOTAL PLV CHECK",
                           "fees": [("CUOC VC CHUA TANG", None), ("SO TIEN % TANG", "SỐ TIỀN % TĂNG {pct}%"),
                                    ("BD NANG", None), (("PHI KHAC PHI", "PHI KHAC SO TIEN"), None),
                                    (("PHI NANG/HA NANG", "PHI NANG / HA NANG"), "PHÍ NÂNG (SHĐ {shd})"),
                                    (("PHI NANG/HA HA", "PHI NANG / HA HA"), "PHÍ HẠ (SHĐ {shd})"),
                                    ("PHI CAN CAN", "PHÍ CÂN (SHĐ {shd})"),
                                    # 沖銷 (TỔNG 要減掉，2026-10-28 抽樣 T7.2026)；名稱前的 "-" 代表減項
                                    ("CAN TRU", "-Cấn trừ phí phát sinh")]}),
]


def detect_song_trieu(text: str) -> bool:
    norm = _vt_norm(text)
    return "SONG TRIEU" in norm or "033 1003828 194" in norm


_vt_register("SONG_TRIEU", "CÔNG TY TNHH SONG TRIỀU", detect_song_trieu,
             "SONG TRIEU COMPANY LIMITED", _VT_SONG_TRIEU)


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
            "inv_series": "1C{yy}TAA",
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


_tcl_sheet_parse, _tcl_sheet_totals = _vt_make("Times Cargo Logistic Co., Ltd - TCL", _VT_TIMES_CARGO)
register_supplier(
    "TIMES_CARGO", "CÔNG TY TNHH DỊCH VỤ GIAO NHẬN HÀNG HÓA (TIMES CARGO)", detect_times_cargo,
    parse_times_cargo_pdf, extract_total=_tcl_pdf_total, extract_grand_total=_tcl_pdf_grand,
    sheet_parse=_tcl_sheet_parse, sheet_totals=_tcl_sheet_totals,
)


# --- VN18 DỊCH VỤ VẬN CHUYỂN SÀI GÒN (空運) -------------------------------------
# 每列一票：空運費 A/F (USD 與換算後 VND)、Local Charge、Break Bulk (都是 VND)，
# 各有自己的紅色發票號碼 (Red invoice no.)。使用者 2026-10-27：發票號碼用 Red invoice no.
# (每項費用取自己右邊那一欄；空白才用 HAWB)。日期欄是
# "26/08/2026 - 146 KGS" 這種文字。
# 整張表沒有逐列合計 (每列只是一票)，但表尾有 "TOTAL" 列 (各費用欄合計) +
# "TOTAL AMOUNT :" (全部合計)：用 sum_row 跟整張表費用加總勾稽
# (2026-11-12 使用者回報「沒有可勾稽的總額」才發現這張表其實有合計列)。
_VT_SAIGON = [
    # OBC 快遞的 BẢNG KÊ CÔNG NỢ (2026-10-28 抽樣，工作表 "POU SUNG-寶崧-NHAP")：每列一張提單 (Số Bill)，
    # THÀNH TIỀN VND 是金額 (已含其他費用)
    (r"(?!)", {"header": "SO BILL", "require": "THANH TIEN VND", "key": ["SO BILL"],
               "fees": [("THANH TIEN VND", "Cước vận chuyển (OBC)")], "date": "NGAY", "kg": "SO KG",
               "sum_row": r"^(TONG|TOTAL|THANH TIEN$)", "port": None,
               "fields": {"bl_no": "SO BILL", "port_of_loading": "CHUYEN DI"}}),
    (r".", {"header": "HAWB", "key": ["HAWB NO"], "name_require": "HAWB NO",
            # A/F 費用名稱 "A/F (USD 59.52)"：USD 原幣金額取 A/F (USD) 欄，只加在 A/F 列 (2026-11-05 POUYUEN 答案)
            "fees": [("AMOUNT (VND)", "A/F", "RED INVOICE NO"),
                     ("LOCAL CHARGE (VND)", "Local Charge", "<RED INVOICE NO"),
                     # 有些月份這欄標題打成 "Break AWB" (不是 "Break Bulk")，意思相同，欄位名稱要兩個都認
                     # (2026-11-12 抽樣：APRIL-2026/FEBRUARY-2026/BK 012025 都是 "Break AWB"，
                     # 原本只認 "Break Bulk" 會漏抓這筆費用，整張金額少了這一欄)
                     (("BREAK BULK", "BREAK AWB"), "Break Bulk", "RED INVOICE NO")],
            "desc_usd": "USD", "desc_usd_only": "A/F",
            # 集團公司名稱用英文名稱列表的寫法 (答案 "POUYUEN VIETNAM COMPANY LIMITED"，帳單抬頭 "POU YUEN VN ENT LTD")
            "const": {"consignee": "POUYUEN VIETNAM COMPANY LIMITED"},
            # 同一 MAWB 的後續列日期空白，沿用上一列日期 (答案表頭說明，2026-11-05)
            "ffill_date": True,
            "date": ("INV DATE", "DATE INV"), "kg": "CHARGEABLE WEIGHT",
            "sum_row": r"^TOTAL$",
            "fields": {"bl_no": "HAWB NO", "mbl_no": "MAWB NO"}}),
]


def detect_saigon_transport(text: str) -> bool:
    norm = _vt_norm(text)
    # 舊字型抬頭讀不出公司名稱：用 MST 0301440413 或表頭組合 (Hawb No + Red Invoice No + Local Charge)
    return ("DICH VU VAN CHUYEN SAI GON" in norm or "0301440413" in norm
            or all(k in norm for k in ("HAWB NO", "RED INVOICE", "LOCAL CHARGE (VND)")))


def _saigon_fill_vnd_gap(rows: List[list]) -> List[list]:
    """OBC 對帳表 (NHAP/XUAT) 的 "THÀNH TIỀN VND" 欄是公式 (USD 金額 × 匯率)；
    帳單沒存計算結果時 calamine 讀成空白 (2026-11-11 POU SUNG 09-SAIF 抽樣，兩張
    工作表都是空白，整張發票讀不到金額)。自己用表頭上方 "EX RATE" 小表的 USD
    匯率算出來，跟 Excel 開啟後會顯示的值一樣。沒有這個版面、或已經有金額時
    不改動 (維持其他樣本的既有行為)。"""
    h = _vt_header_row(rows, "SO BILL")
    if h is None:
        return rows
    labels = _vt_labels(rows, h)
    vnd_c, usd_c = _vt_col(labels, "THANH TIEN VND"), _vt_col(labels, "THANH TIEN USD")
    if vnd_c is None or usd_c is None:
        return rows
    data_rows = rows[h + 2:]
    if any(_vn_num(_vn_cell(row, vnd_c)) for row in data_rows):
        return rows      # 已經有金額 (公式有存檔)，不用補
    rate = None
    for row in rows[:h]:
        for i, cell in enumerate(row):
            if _vt_norm(cell) == "USD":
                rate = next((v for v in (_vn_num(c) for c in row[i + 1:]) if v), None)
        if rate:
            break
    if not rate:
        return rows
    out = [list(row) for row in rows]
    for row in out[h + 2:]:
        usd_val = _vn_num(_vn_cell(row, usd_c))
        if usd_val and vnd_c < len(row):
            row[vnd_c] = round(usd_val * rate)
    return out


def _saigon_sheets(book: SheetBook) -> SheetBook:
    return SheetBook(book.path, [(name, _saigon_fill_vnd_gap(rows)) for name, rows in book.sheets])


_saigon_parse, _saigon_totals = _vt_make("SAI GON AIR & OCEAN FREIGHT SCE.CO.LTD", _VT_SAIGON)
register_supplier(
    "SAIGON_TRANSPORT", "CÔNG TY TNHH DỊCH VỤ VẬN CHUYỂN SÀI GÒN", detect_saigon_transport, None,
    sheet_parse=lambda book: _saigon_parse(_saigon_sheets(book)),
    sheet_totals=lambda book: _saigon_totals(_saigon_sheets(book)))


# --- VN19 IMX EXPRESS (CÔNG TY TNHH DỊCH VỤ CHUYỂN PHÁT NHANH IMX，MST 0315868311) --
# 兩張對帳表：D/O FEE (USD 與 VND 兩欄，取 VND) 與 AIR FREIGHT。
# 2026-11-05 使用者答案 (VN-19 sample)：
#   發票號碼 = Ký Hiệu (例如 1C26TMX)，兩張表同一號碼 (Hóa Đơn Số 空白)；
#   AIR 金額改寫 VND (TOTAL VND = USD × 26,180，含 A/F/FUEL/P/U/HD 全部費用)，
#   描述保留 "(USD 原幣)"；集團公司、供應商用英文名稱；ROUTINE 拆成起迄港；
#   船名/航班加 " (Flight)"；VAT 列描述 "V.A.T TAX (8%) (VAT 8%)"，日期 N/A。
# 稽核：D/O FEE 用表尾 TOTAL (VND) / SUBTOTAL (VND)；AIR 用表尾 "TOTAL (USD)" 那列 (內容是 VND 合計)。
# 兩張表共用同一發票號碼，合計會自動相加 (sheet_totals 合併)。
_VT_IMX_COMMON = {"key": ["AWB.NO"], "date": "=DATE", "kg": "WEIGHT",
                  "doc_key": r"K\S{0,3}\s*Hi\S{0,4}\s*:\s*(\d[A-Z]\d{2}[A-Z]+)",
                  "desc_usd": "USD", "port": "ROUTINE", "port_arrow": True,
                  "const": {"consignee": "POUYUEN VIETNAM COMPANY LIMITED"},
                  "fields": {"bl_no": "AWB.NO", "shipper": "=SHIPPER", "vessel": "FLIGHT"},
                  "field_suffix": {"vessel": " (Flight)"}}
_VT_IMX = [
    (r"D\.?O FEE", {**_VT_IMX_COMMON, "header": "AWB.NO", "fees": [("D/O FEE VND", "D/O FEE")],
                    "sheet_total": ("TOTAL (VND)", "SUBTOTAL (VND)"), "vat_row": True,
                    "vat_label": "V.A.T TAX (8%) (VAT 8%)", "vat_no_date": True, "vat_last": True}),
    (r"AIR FREIGHT", {**_VT_IMX_COMMON, "header": "AWB.NO",
                      "fees": [("TOTAL VND", "A/F CHARGE")],
                      "sheet_total": ("TOTAL (USD)", "TOTAL (USD)")}),
]


def detect_imx(text: str) -> bool:
    norm = _vt_norm(text)
    # 舊檔 (2025) 沒有公司名稱：發票符號 1C25TMX，或「Mã bảng kê」+ AWB.No + D/O FEE
    return ("IMX EXPRESS" in norm or "0315868311" in norm or re.search(r"\dC\d{2}TMX\b", norm) is not None
            or ("MA BANG KE" in norm and "AWB.NO" in norm and "D/O FEE" in norm))


# --- VN19 IMX 舊版 (2025) AIR FREIGHT CHARGE INVOICE：單一工作表，費用只有 USD，沒有
# VAT、沒有發票號碼 (用 AWB.No)。整張表只有 1 個 EX. RATE 與 1 個 TOTAL (VND)：每列每個
# 費用 = ROUND(USD × EX.RATE)，四捨五入的尾差加在最後一筆費用，讓合計 = TOTAL (VND)
# (答案 POU HUNG-AIR-APR-2025，2026-11-12 使用者確認)。版型跟新版 (D.O FEE/AIR FREIGHT
# 兩張表、VND 直接列出) 完全不同，不套用 _VT_IMX 那組 DSL 設定。
def _imx_old_sheet(book) -> Optional[List[list]]:
    for _, rows in book.sheets:
        if any(_vn_text(_vn_cell(row, 0)).upper() == "AIR FREIGHT CHARGE INVOICE" for row in rows[:10]):
            return rows
    return None


def _imx_old_items(rows: List[list]) -> Tuple[List[Dict], Optional[float]]:
    header_idx = next((i for i, row in enumerate(rows) if _vn_text(_vn_cell(row, 0)) == "No."), None)
    if header_idx is None:
        return [], None
    header = rows[header_idx]

    def col(*keywords):
        return next((i for i, v in enumerate(header) if all(k in _vn_label(v) for k in keywords)), None)

    c_date, c_awb, c_shipper, c_consignee = col("DATE"), col("AWB"), col("SHIPPER"), col("CONSIGNEE")
    c_weight = col("WEIGHT")
    c_af, c_fuel, c_pu, c_handling = col("A/F", "CHARGE"), col("FUEL"), col("P/U"), col("HANDLING")
    c_port_load = c_shipper + 1 if c_shipper is not None else None
    c_port_disc = c_consignee + 1 if c_consignee is not None else None

    # 舊版表頭 CONSIGNEE 是舊字型編碼 ("COÂNG TY TNHH POU HUNG VIEÄT NAM")，重音字母
    # 常normalize 不掉：直接比對明文關鍵字 "POU HUNG"/"POUYUEN" (使用者 2026-11-12 確認)
    consignee_raw = (_vt_consignee(rows) or "").upper()
    consignee = ("POU HUNG VIETNAM COMPANY LIMITED" if "POU HUNG" in consignee_raw
                 else "POUYUEN VIETNAM COMPANY LIMITED" if "POUYUEN" in consignee_raw else None)
    rate = total_vnd = None
    data_rows = []
    for row in rows[header_idx + 2:]:
        label = _vn_label(_vn_cell(row, 0))
        if label.startswith("EX. RATE") or label.startswith("EX RATE"):
            rate = next((v for v in row if isinstance(v, (int, float))), None)
        elif label.startswith("TOTAL (VND)"):
            total_vnd = next((v for v in row if isinstance(v, (int, float))), None)
        elif isinstance(_vn_cell(row, 0), (int, float)):
            data_rows.append(row)

    items: List[Dict] = []
    for row in data_rows:
        awb = _vn_text(_vn_cell(row, c_awb))
        kg_val = _vn_num(_vn_cell(row, c_weight)) if c_weight is not None else None
        head = {
            "supplier": "IMX EXPRESS SERVICE COMPANY LIMITED", "consignee": consignee,
            "invoice_no": awb, "invoice_date": _vn_date(_vn_cell(row, c_date)), "bl_no": awb,
            "shipper": _vn_text(_vn_cell(row, c_shipper)) or None,
            "port_of_loading": _vn_text(_vn_cell(row, c_port_load)) or None,
            "port_of_discharge": _vn_text(_vn_cell(row, c_port_disc)) or None,
            "volume": _fmt_volume(kg=f"{round(kg_val, 3):g}" if kg_val else None),
        }
        for c, label in ((c_af, "A/F CHARGE"), (c_fuel, "FUEL"), (c_pu, "P/U"), (c_handling, "HANDLING")):
            usd = _vn_num(_vn_cell(row, c)) if c is not None else None
            if not usd:
                continue
            # 千分位逗號、小數照原樣 (同 desc_usd 的慣例，答案 "(USD 2,872.55)")；
            # ":,.3f" 順便去掉 Excel 浮點誤差 (7.199999999999999 -> "7.200" -> "7.2")
            usd_note = f"USD {usd:,.3f}".rstrip("0").rstrip(".")
            items.append({**head, "description": f"{label} ({usd_note})",
                          "amount": round(usd * rate) if rate else None, "vat": None})
    if rate and total_vnd is not None and items:
        items[-1]["amount"] += round(total_vnd) - sum(it["amount"] for it in items)
    return items, (round(total_vnd) if total_vnd is not None else None)


def _parse_imx_old(book) -> List[Dict]:
    rows = _imx_old_sheet(book)
    items, _ = _imx_old_items(rows) if rows else ([], None)
    return [{"header": {}, "items": items or [{"description": None, "amount": None}]}]


def _imx_old_totals(book) -> List[Dict]:
    rows = _imx_old_sheet(book)
    if not rows:
        return []
    _, total_vnd = _imx_old_items(rows)
    return [{"invoice_no": None, "total": total_vnd}] if total_vnd is not None else []


_imx_new_parse, _imx_new_totals = _vt_make("IMX EXPRESS SERVICE COMPANY LIMITED", _VT_IMX)


def _imx_sheet_parse(book):
    return _parse_imx_old(book) if _imx_old_sheet(book) else _imx_new_parse(book)


def _imx_sheet_totals(book):
    return _imx_old_totals(book) if _imx_old_sheet(book) else _imx_new_totals(book)


register_supplier(
    "IMX", "CÔNG TY TNHH DỊCH VỤ CHUYỂN PHÁT NHANH IMX", detect_imx, None,
    sheet_parse=_imx_sheet_parse, sheet_totals=_imx_sheet_totals,
)  # 供應商用英文名稱 (答案 2026-11-05)


# --- VN15 EXPEDITORS 的 Excel 對帳表 (AIR IMP) -----------------------------------
# 「單價」(11~17 欄) 與「金額」(20~26 欄) 兩組費用欄，只取金額；另有運費 CƯỚC、
# 代收代付 THU HỘ ĐẠI LÝ NƯỚC NGOÀI、VAT、合計。發票號碼 = 每列的 TAX INVOICE NUMBER。
# PDF 規則在上面 (parse_expeditors)；這裡重新註冊一次，把 Excel 規則加上去。
_LH = "LAM HANG - (THUE 8%)"
_VT_EXPEDITORS = [
    (r".", {"header": "EI REF", "key": ["TAX INVOICE NUMBER", "VAN DON PHU"],
            # 表頭有「單價」與「金額」兩組重複欄名 (CƯỚC、THU HỘ ĐẠI LÝ NƯỚC NGOÀI 各
            # 出現兩次)，單價組恆為 0：用 "$=" 取最後一個 (金額組) 才是真正的金額
            # (2026-11-10 抽樣 AIR IMP 08.30.2025 CƯỚC 欄被單價組的 0 蓋掉)。
            # OCEAN IMP 版型第 4/5 欄標題不同 (PHÍ BỐC XẾP / PHÍ CFS，不是 AIR IMP 的
            # PHÍ VỆ SINH CONT / PHÍ PCS/CFS)：兩種都要比對 (2026-11-10 抽樣
            # OCEAN IMP 07.30.2025 PHÍ BỐC XẾP 欄漏抓)
            "fees": [(f"{_LH} PHI LENH", None), (f"{_LH} PHI XEP DO", None), (f"{_LH} PHI PHU TROI", None),
                     ((f"{_LH} PHI VE SINH", f"{_LH} PHI BOC XEP"), None),
                     ((f"{_LH} PHI PCS", f"{_LH} PHI CFS"), None), (f"{_LH} PHI E-MANIFEST", None),
                     (f"{_LH} PHI HANDLING", None), ("$=CUOC", "CƯỚC"),
                     ("$=THU HO DAI LY NUOC NGOAI", "THU HỘ ĐẠI LÝ NƯỚC NGOÀI")],
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


def _ct_pc_invoice_list_sheets(book) -> List[Tuple[str, list]]:
    """POU CHEN 批次裡，代墊吊櫃費工作表有時叫 "PHI…" (配 _VT_CT_POUCHEN 的
    "^PHI" 設定)，有時叫別的名稱 (例如 "CH PC")：名稱比對抓不到，整張代墊
    資料 (含可勾稽的每票 AMOUNT) 被跳過 (2026-11-12 使用者回報 TU 12-17.10.26
    「沒有可勾稽的總額」才發現：P.CHEN 表本身這期沒有 TOTAL 列可勾稽，但真正
    的問題是 CH PC 這張代墊表整張被漏掉，裡面本來就有逐票 AMOUNT 可查核)。
    改用內容判斷 (跟 _ct_paid_on_behalf 一致的 "INVOICE LIST" 標記)，不限制
    工作表名稱；"^PHI" 名稱比對到的表已經由 _VT_CT_POUCHEN 處理，這裡排除。"""
    out = []
    for name, rows in book.sheets:
        if re.match(r"^PHI", _vt_norm(name)):
            continue
        top = " ".join(_vn_label(v) for r in rows[:15] for v in r if _vn_label(v))
        if "INVOICE LIST" in top:
            out.append((name, rows))
    return out


def _ct_pc_batch_ref(rows: List[list]) -> Optional[str]:
    """P.CHEN 表抬頭的對帳單批次號 ("Số BK: 12-17/10/26")，給沒有發票號碼的
    整批調整費用當發票號碼用 (見 _ct_pc_extra_notes)。"""
    text = "\n".join(" | ".join(_vn_text(v) for v in r if _vn_text(v)) for r in rows[:15])
    return _search(r"S[ốO]\s*BK\s*:\s*([\w./-]+)", text)


def _ct_pc_extra_notes(book) -> List[Dict]:
    """P.CHEN 表明細列結束、TOTAL 列之前，偶爾有沒有發票號碼、只用一段文字
    說明原因的整批調整費用 (例如 "17/10: vào kho TBS thay nắp hộp giày…
    2,000,000")，不屬於任何一張出貨發票。_vt_make 的明細列比對要靠 INVOICE #
    欄分組，這種列抓不到發票號碼就被跳過 (2026-11-12 使用者回報 TU 12-17.10.26
    金額查核差了剛好這兩筆的合計才發現：P.CHEN 自己的 TOTAL 列已經把這兩筆
    算進去，明細卻漏了，所以 whole-file 勾稽對不起來)。
    使用者 2026-11-12 確認：這種列算在對帳單批次號 (Số BK) 底下，不併入任何
    一張出貨發票。"""
    items: List[Dict] = []
    for name, rows in book.sheets:
        if not re.match(r"^P\.?\s?CHEN", _vt_norm(name)):
            continue
        head_i = next((i for i, row in enumerate(rows[:20])
                        if _vt_col(_vt_labels(rows, i, two_rows=False), "INVOICE") is not None), None)
        if head_i is None:
            continue
        labels = _vt_labels(rows, head_i, two_rows=False)
        idx_inv = _vt_col(labels, "INVOICE")
        idx_fees = [(c, fee_name) for c, fee_name in (
            (_vt_col(labels, "PHI DICH VU"), "PHÍ DỊCH VỤ"),
            (_vt_col(labels, "PHI VAN TAI"), "PHÍ VẬN TẢI"),
            (_vt_col(labels, "PHI LUU CONT"), "Phi Luu cont")) if c is not None]
        batch_ref = _ct_pc_batch_ref(rows)
        for row in rows[head_i + 1:]:
            if _vn_label(_vn_cell(row, 0)).startswith("TOTAL"):
                break
            if _vn_text(_vn_cell(row, idx_inv)):
                continue   # 一般明細列 (有發票號碼)，_ct_pc_parse 已經處理
            note = _vn_text(_vn_cell(row, 0))
            for c, fee_name in idx_fees:
                amount = _vn_num(_vn_cell(row, c))
                if amount:
                    items.append({
                        "description": note or fee_name, "amount": amount, "vat": None,
                        "invoice_no": batch_ref, "supplier": "CTY TNHHDV GN VT VA TM CONG THANH",
                        "port_of_loading": "VIETNAM", "port_of_discharge": "VIETNAM",
                        "consignee": "POU CHEN",
                    })
    return items


def _ct_pc_parse_all(book) -> List[Dict]:
    parsed = _ct_pc_parse(book)
    extra_notes = _ct_pc_extra_notes(book)
    if extra_notes:
        parsed.append({"header": {}, "items": extra_notes})
    for _name, rows in _ct_pc_invoice_list_sheets(book):
        extra = _ct_paid_on_behalf(rows)
        if extra and extra["items"]:
            parsed.append({"header": extra["header"], "items": extra["items"]})
    return parsed


def _ct_pc_totals_all(book) -> List[Dict]:
    """P.CHEN 本身的 TOTAL 列 (whole-file 勾稽) 沒有把代墊吊櫃費 (CH PC/PHI
    表) 算進去——兩者是獨立費用，各自開票，只是同一份檔案。直接把 CH PC 的
    稽核結果加進來會讓 P.CHEN 原本就對得上的 whole-file 勾稽多算一筆、變成
    「不一致」(2026-11-12)：whole-file 的未稅合計要把 CH PC 的未稅合計也加
    進去，才會等於全部明細 (P.CHEN + CH PC) 的加總。"""
    totals = _ct_pc_totals(book)
    extra_entries: List[Dict] = []
    extra_net = 0.0
    has_extra = False
    for _name, rows in _ct_pc_invoice_list_sheets(book):
        extra = _ct_paid_on_behalf(rows)
        if not extra:
            continue
        extra_entries += extra.get("audit", [])
        extra_net += sum(it.get("amount") or 0 for it in extra["items"])
        has_extra = True
    if has_extra:
        for t in totals:
            if t.get("invoice_no") is None and t.get("total") is not None:
                t["total"] += extra_net
    return totals + extra_entries


def detect_cong_thanh_any(text: str) -> bool:
    # 2025 年的 POU CHEN 對帳表抬頭沒有 CONG THANH 稅號：用壓車表標題辨識
    return detect_cong_thanh(text) or "SAVED CHASSIS TRAILER LIST" in text.upper()


register_supplier(
    "CONG_THANH", "CTY TNHHDV GN VT VA TM CONG THANH", detect_cong_thanh_any, None,
    sheet_parse=lambda book: _ct_pc_parse_all(book) if _ct_is_pouchen(book) else parse_cong_thanh_sheets(book),
    sheet_totals=lambda book: _ct_pc_totals_all(book) if _ct_is_pouchen(book) else _cong_thanh_totals(book),
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
# 一份大檔 PDF (396 頁) 混 3 種頁面 (使用者 2026-11-09 VN07_Claude_Code_修改說明)：
#  (A) 一般電子發票，費用行 "Phí xếp dỡ hàng hóa SGN5049893 VND M3 3.896 280,709 8% 1,093,642"
#      (名稱、FCR、幣別、單位、數量、單價、稅率、未稅金額)，稅額只有總數 -> 單獨一筆 VAT。
#  (B) 字型缺字頁：越南文字母被 pdfplumber 讀成 "(cid:NN)" (例如 "Phí x(cid:1)p d(cid:2) hàng
#      hóa")，但英文數字、ASCII 字母、括號內英文標籤都正常。表頭/合計一律只認英文標籤 (不要用
#      越南文標籤，會因為缺字對不到)；品名用「去掉 (cid:NN) 後只留 A-Z/a-z」當 key 對應回正確
#      品名 (A/B 兩種頁面同一套邏輯，不用另外分支判斷)。
#  (C) MAERSK VIỆT NAM (船公司代收運費) 頁，表頭/費用列完全不同版型，見 _maersk_vn_carrier()。
# Excel：Late come charge 對帳表，每列有自己的發票號碼 (INVOICE NO.)。
_MAERSK_VN_ITEM_RE = re.compile(
    # 參考號碼也可能數字開頭 ("26HCM0018482"，2026-10-28 抽樣)
    r"^(?P<desc>.+?)\s+(?P<ref>FCR\d+|[A-Z]{3,}\d{5,}|\d{2}[A-Z]{2,4}\d{5,})\s+(?P<cur>[A-Z]{3})\s+(?P<unit>\S+)\s+"
    r"(?P<qty>[\d,.]+)\s+(?P<price>[\d,.]+)\s+(?P<rate>\d+%|KCT)\s+(?P<amount>[\d,.]+)$", re.MULTILINE)

# (B) 字型缺字頁的品名修正字典：key 是「去掉 (cid:NN)、只留字母」的比對鍵 (例如 "Phí xếp dỡ
# hàng hóa" 跟缺字的 "Phí x(cid:1)p d(cid:2) hàng hóa" 都會得到同一個鍵 "Phxpdhngha")，value 是
# 正常頁的完整品名。對不到就照原文 (之後遇到新品名，依使用者確認的統一費用對照表補進這份清單)。
_MAERSK_VN_KNOWN_DESCS = (
    "Phí xếp dỡ hàng hóa", "Phí xếp dỡ container tại cảng đi", "Phí phát hành FCR",
    "Phí cầu cảng", "Phí khai báo tổng trọng lượng (VGM)", "Phí quản lý chứng từ",
    "Phí niêm chì hãng tàu", "Phí chứng từ - cảng đi", "Phí phát hành vận đơn",
    "Phí dịch vụ quản lý khách hàng",
)


def _maersk_vn_desc_key(text: str) -> str:
    cleaned = re.sub(r"\(cid:\d+\)", "", text or "")
    return re.sub(r"[^A-Za-z]", "", cleaned)


_MAERSK_VN_DESC_BY_KEY = {_maersk_vn_desc_key(d): d for d in _MAERSK_VN_KNOWN_DESCS}


def _maersk_vn_fix_desc(desc: str) -> str:
    if "(cid:" not in desc:
        return desc
    return _MAERSK_VN_DESC_BY_KEY.get(_maersk_vn_desc_key(desc), desc)


def _maersk_vn_is_carrier(text: str) -> bool:
    """(C) MAERSK VIỆT NAM 船公司代收頁 (跟 (A)/(B) MAERSK LOGISTICS & SERVICES 電子發票完全不同版型)。"""
    upper = text.upper()
    return "MAERSK VIETNAM LIMITED" in upper or "AS AGENT FOR CARRIER MAERSK" in upper


def detect_maersk_vn(text: str) -> bool:
    upper = text.upper()
    if "MAERSK LOGISTICS & SERVICES VI" in upper or _maersk_vn_is_carrier(text):
        return True
    return "MAERSK" in upper and "LATE COME CHARGE" in upper


def parse_maersk_vn(text: str) -> Dict:
    if _maersk_vn_is_carrier(text):
        return _parse_maersk_vn_carrier(text)
    return _parse_maersk_vn_logistics(text)


def _parse_maersk_vn_logistics(text: str) -> Dict:
    h: Dict[str, Optional[str]] = {}
    # 表頭只用括號內的英文標籤找值 (字型缺字頁的越南文標籤常常缺字對不到，但英文標籤跟數字
    # 不受影響；使用者 2026-11-09 確認 A/B 兩種頁面共用這套規則)。
    series = _search(r"\(Symbol No\. & Serial No\.\):\s*(\S+)", text)
    number = _search(r"\(Invoice No\.\):\s*(\d+)", text)
    # 發票號碼 = 發票符號 + "-" + 號碼 (樣本答案 VN07 "1C26TSG-00164419"，使用者 2026-11-05 確認)
    h["invoice_no"] = f"{series}-{number}" if series and number else number
    h["invoice_date"] = _vn_pdf_date(_search(r"\(Invoice [Dd]ate\):\s*(\S+)", text))
    h["supplier"] = "MAERSK LOGISTICS & SERVICES VIETNAM COMPANY LIMITED"   # 英文名 (使用者確認)
    h["consignee"] = _vn_full_consignee(_search(r"\(Invoice No\.\):[^\n]*\n(.+)", text))
    h["port_of_discharge"] = _search(r"\(POD\):\s*(\S+)", text)
    h["port_of_loading"] = _search(r"\(POL\):\s*(\S+)", text)
    h["order_no"] = _search(r"\(Ref No\):\s*(\S+)", text)

    # 每筆費用各自的 FCR/HBL/Job number 就是這一列的 B/L NO.；Volume = 同一個 FCR 的 M3
    # 數量 (DOC 列沿用同一 FCR 的 M3 列)，沒有 M3 列時填「數量 單位」(使用者 2026-11-09 確認，
    # VN07 的 Volume 寫法沿用帳單原印的 M3 單位，不走其他供應商的 _fmt_volume() CBM 慣例)。
    groups: "collections.OrderedDict[str, list]" = collections.OrderedDict()
    for m in _MAERSK_VN_ITEM_RE.finditer(text):
        groups.setdefault(m.group("ref"), []).append(m)

    items = []
    for ref, matches in groups.items():
        cbm = next((mm.group("qty") for mm in matches if mm.group("unit").upper() == "M3"), None)
        for mm in matches:
            volume = f"{cbm} M3" if cbm else f"{mm.group('qty')} {mm.group('unit')}"
            items.append({
                "description": _maersk_vn_fix_desc(mm.group("desc").strip()),
                "amount": _vn_pdf_num(mm.group("amount")), "vat": None,
                "bl_no": ref, "volume": volume,
            })

    vat = _search(r"\(Total VAT amount\):\s*([\d,.]+)", text)
    # 稅率印在這張發票的 Subtotal 區塊 ("chịu thuế suất 8%")，少數發票這個區塊跟
    # "(Total VAT amount)" 剛好被分頁切到不同頁 (使用者 2026-11-09 樣本
    # 1C25TSG-00056964：Subtotal/8% 在上一頁，Total VAT amount 在下一頁)，這一頁
    # 抓不到稅率時沿用 VN07 目前唯一出現過的稅率 8%。
    rate = _search(r"\s(\d+)%\s+[\d,.]+$", text) or "8"
    if vat:
        # 整張 VAT 列：抬頭的單號、港口、材積填 N/A (樣本答案 VN07 PDF，2026-11-05)
        items.append({"description": f"VAT {rate}%" if rate else "VAT", "amount": None,
                      "vat": _vn_pdf_num(vat), "order_no": "N/A", "bl_no": "N/A", "mbl_no": "N/A",
                      "port_of_loading": "N/A", "port_of_discharge": "N/A", "volume": "N/A"})
    return {"header": h, "items": items or [{"description": None, "amount": None, "vat": None}]}


# (C) MAERSK VIỆT NAM 船公司代收頁：費用行 "PHÍ XẾP DỠ HÀNG - CẢNG ĐI 1 CÔNG TEN NƠ VND
# 3,050,000 1 3,050,000KHAC:5.26% 160,526 3,210,526" (名稱、數量、單位、幣別、單價、匯率、
# 未稅金額[跟稅率黏在一起無空白]、稅率、稅額、含稅金額)，每筆費用自帶稅額 (不另外產生 VAT 列)。
# 單位只有 2 種 (使用者指定)：CÔNG TEN NƠ = Container、VẬN ĐƠN = B/L。
_MAERSK_VN_CARRIER_ITEM_RE = re.compile(
    r"^(?P<desc>[^\n]+?)\s+(?P<qty>[\d.,]+)\s+(?P<unit>CÔNG TEN NƠ|VẬN ĐƠN)\s+(?P<cur>VND|USD)\s+"
    r"(?P<price>[\d,.]+)\s+(?P<roe>[\d,.]+)\s+(?P<amount>[\d,.]+)KHAC:(?P<rate>[\d.]+)%\s+"
    r"(?P<tax>[\d,.]+)\s+(?P<total>[\d,.]+)$", re.MULTILINE)
_MAERSK_VN_CARRIER_POR_POD_RE = re.compile(
    r"^(?P<por>.+?,[A-Z]{2})\s(?P<pod>.+?,[A-Z]{2})\s(?P<vessel>.+)$", re.MULTILINE)
_MAERSK_VN_CARRIER_ETD_ETA_RE = re.compile(
    r"^(?P<etd>\d{2}/\d{2}/\d{4})\s(?P<eta>\d{2}/\d{2}/\d{4})\s(?P<voy>\S+)$", re.MULTILINE)


def _parse_maersk_vn_carrier(text: str) -> Dict:
    h: Dict[str, Optional[str]] = {}
    series = _search(r"\(Symbol No\. & Serial No\.\):\s*(\S+)", text)
    number = _search(r"\(Invoice No\.\):\s*(\d+)", text)
    h["invoice_no"] = f"{series}-{number}" if series and number else number
    h["invoice_date"] = _vn_pdf_date(_search(r"\(Invoice [Dd]ate\):\s*(\S+)", text))
    h["supplier"] = "MAERSK VIETNAM LIMITED"   # 船公司 (使用者指定，同一份 PDF 一起處理)
    h["consignee"] = _vn_full_consignee(_search(r"\(Name, address of Customer\):[^\n]*\n(.+)", text))
    h["order_no"] = _search(r"\(Reference\):\s*(\S+)", text)
    h["bl_no"] = _search(r"\(B/L No / Container No\.\):\s*(\S+)", text) or "N/A"

    vessel = None
    m = _MAERSK_VN_CARRIER_POR_POD_RE.search(text)
    if m:
        h["port_of_loading"] = m.group("por")
        h["port_of_discharge"] = m.group("pod")
        vessel = m.group("vessel").strip()
    m2 = _MAERSK_VN_CARRIER_ETD_ETA_RE.search(text)
    if m2:
        h["onboard_date"] = _vn_pdf_date(m2.group("etd"))
        h["arrive_date"] = _vn_pdf_date(m2.group("eta"))
        # Vessel = 船名 + Main Voy No. (例如 "MAERSK DHAKA 512B"，使用者指定)
        h["vessel"] = f"{vessel} {m2.group('voy')}" if vessel else None
    elif vessel:
        h["vessel"] = vessel

    items = []
    for m3 in _MAERSK_VN_CARRIER_ITEM_RE.finditer(text):
        unit_en = "Container" if "CÔNG TEN" in m3.group("unit") else "B/L"
        items.append({
            "description": m3.group("desc").strip(),
            "amount": _vn_pdf_num(m3.group("amount")), "vat": _vn_pdf_num(m3.group("tax")),
            "volume": f"{m3.group('qty')} x {unit_en}",
        })
    return {"header": h, "items": items or [{"description": None, "amount": None, "vat": None}]}


_MAERSK_VN_CARRIER_TOTAL_RE = re.compile(r"\(Total\):\s*([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)")


def _maersk_vn_invoice_no(text: str) -> Optional[str]:
    series = _search(r"\(Symbol No\. & Serial No\.\):\s*(\S+)", text)
    number = _search(r"\(Invoice No\.\):\s*(\d+)", text)
    return f"{series}-{number}" if series and number else number


def _maersk_vn_total(text: str):
    if _maersk_vn_is_carrier(text):
        m = _MAERSK_VN_CARRIER_TOTAL_RE.search(text)
        value = _vn_pdf_num(m.group(1)) if m else None
    else:
        value = _vn_pdf_num(_search(r"\(Subtotal\):\s*([\d,.]+)", text))
    if value is None:
        return None
    # 這一頁的總額要歸給這一頁的發票：發票號碼 = Symbol No. & Serial No. + Invoice No.，
    # 原文把兩者印在不同行，合併字串找不到完整比對，用字典 {發票號碼: 總額} 直接指定歸屬
    # (2026-11-11 修正：171 張全部誤判成「沒有可勾稽的總額」)
    inv_no = _maersk_vn_invoice_no(text)
    return {inv_no: value} if inv_no else value


def _maersk_vn_grand(text: str):
    if _maersk_vn_is_carrier(text):
        m = _MAERSK_VN_CARRIER_TOTAL_RE.search(text)
        value = _vn_pdf_num(m.group(3)) if m else None
    else:
        value = _vn_pdf_num(_search(r"\(Total amount\):\s*([\d,.]+)", text))
    if value is None:
        return None
    inv_no = _maersk_vn_invoice_no(text)
    return {inv_no: value} if inv_no else value


# 每列的 VAT 寫在費用同一列 (vat_same_row)；日期用每列 INVOICE DATE 欄 (只有日月時年份取工作表標題)
# 供應商名稱用英文 (2026-11-05 使用者確認 VN07 Excel 用英文名)
_MAERSK_VN_SHEET = [
    (r".", {"header": "BILL NO", "key": ["=INVOICE NO.", "BILL NO"],
            "fees": [("AMOUNT (NOT VAT)", "@REMARK")], "vat": "=VAT", "vat_desc": "VAT 8%", "vat_same_row": True,
            "desc_usd": "USD",
            # 檔案沒有客戶名稱；使用者指定依地點 Đồng Tháp 為 YUE DE (2026-11-05)
            "const": {"consignee": "YUE DE VIETNAM COMPANY LIMITED"},
            "total": "AMOUNT (NOT VAT)", "grand": "TOTAL (VND)", "cbm": "=CBM", "two_rows": False,
            "date": "INVOICE DATE",
            "fields": {"bl_no": "BILL NO", "order_no": "SO BOOKING", "onboard_date": "NGAY XUAT",
                       "mbl_no": "SO TO KHAI"}}),
]
_maersk_vn_sheet_parse, _maersk_vn_sheet_totals = _vt_make(
    "MAERSK LOGISTICS & SERVICES VIETNAM COMPANY LIMITED", _MAERSK_VN_SHEET)

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
    # 費用名稱後的代碼 "- 215" 等照答案寫 (2026-11-06)
    "Phí biên nhận hàng đường biển - 215", "CY monitoring - 376", "Phí xếp dỡ tại cảng đi - 400",
    "Phí làm hàng lẻ - 428", "Phí e-VGM - 347", "Phí Cơ sở hạ tầng cảng xuất",
    "Phí làm hàng ngoài giờ", "Phí sắp xếp hàng xuất", "Phí lưu kho hàng xuất",
    "Phí DEM/DET")
#   FCR 欄 (dn) 有時整欄空白 (連 "0" 佔位都沒有，只留 Combined SO# 一個數字)，
#   前後不能同時要求 \s+ (見 CLAUDE.md Regex 陷阱)，改成整組 FCR 選填 (使用者
#   2026-11-10 抽查 KN_VN BY WEEKLY 批次，FEB/JAN/MAR/OCT.2025 第 1 頁漏抓)
_KN_VN_LIST_ROW_RE = re.compile(
    r"^(?P<file>\d{8,})\s+(?:(?P<dn>\S+)\s+)?(?P<so>\d+)\s+(?P<cbm>[\d.]+)\s+(?P<kg>[\d.]+?)"
    r"(?P<dest>[A-Z][A-Z ]+?)\s+(?P<money>(?:[\d,]+ VND\s*){12})$", re.MULTILINE)


def _kn_vn_pages(pdf_path: str) -> List[str]:
    """附件清單常跨頁 (使用者 2026-11-10 抽查 BY WEEKLY 批次)：續頁沒有公司抬頭、
    也沒有 "ĐƠN NGÀY" 這個清單日期，只剩費用欄表頭 + 明細列，併回前一頁同一張
    清單，明細才能套用同一個發票日期 (否則續頁那幾票 INVOICE DATE 會是 N/A)。"""
    out: List[str] = []
    for page in extract_pages_from_pdf(pdf_path):
        upper = page.upper()
        is_continuation = (out and "KUEHNE + NAGEL" not in upper
                           and (_KN_VN_LIST_ROW_RE.search(page) or _KN_VN_SHORT_ROW_RE.search(page)))
        if is_continuation:
            out[-1] += "\n" + page
        else:
            out.append(page)
    return out


def detect_kn_vn(text: str) -> bool:
    upper = text.upper()
    if "KUEHNE + NAGEL" in upper and ("HÓA ĐƠN" in upper or "HOÁ ĐƠN" in upper or "VIỆT NAM" in upper):
        return True
    # 附件清單續頁：公司抬頭只印在清單第一頁，續頁只剩費用欄表頭 + 明細列，沒有
    # 公司名稱或任何關鍵字，整頁會被當成「不是這家」漏掉 (使用者 2026-11-10 抽查
    # KN_VN BY WEEKLY 批次，FEB/JAN/MAR/OCT.2025 各漏了一頁的明細)
    return bool(_KN_VN_LIST_ROW_RE.search(text) or _KN_VN_SHORT_ROW_RE.search(text))


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
    h: Dict[str, Optional[str]] = {"supplier": "Kuehne + Nagel Co., Ltd"}
    list_rows = _kn_vn_list_rows(text)
    # 附件清單表格常跨頁：公司抬頭頁本身的明細列印在下一頁，這一頁單獨解析時
    # list_rows 是空的，不要掉進下面的「單張發票」分支產生一筆全 N/A 的假列
    # (使用者 2026-11-10 抽查 KN_VN BY WEEKLY 批次)
    if not list_rows and "ĐÍNH KÈM HOÁ ĐƠN" in text.upper():
        return {"header": h, "items": []}
    if list_rows:
        h["invoice_date"] = _vn_pdf_date(_search(r"ĐƠN NGÀY\s*(\d{1,2}/\d{1,2}/\d{4})", text.upper()))
        h["consignee"] = "POUYUEN VIETNAM COMPANY LIMITED" if "POUYUEN" in text.upper() else None
        items = []
        # 第 10 欄在精簡版清單是 "Phí thị thực" (不是 DEM/DET)
        names = _KN_VN_LIST_FEES[:9] + (("Phí thị thực",) if "Phí thị thực" in text else _KN_VN_LIST_FEES[9:])
        for r in list_rows:
            common = {"invoice_no": r["file"], "bl_no": r["file"], "order_no": r["so"], "port_of_discharge": r["dest"] or None,
                      "volume": _fmt_volume(cbm=r["cbm"], kg=r["kg"])}
            items += [{"description": name, "amount": amt, "vat": None, **common}
                      for name, amt in zip(names, r["fees"]) if amt]
            items.append({"description": "VAT amount (8%)", "amount": None, "vat": r["vat"], **common})
        return {"header": h, "items": items or [{"description": None, "amount": None, "vat": None}]}
    series = _search(r"\(Series\):\s*(\S+)", text)
    number = _search(r"\(Invoice No\.\):\s*(\S+)", text)
    # 發票符號 + "-" + 號碼 (例如 1C26TKS-134757，2026-11-06 新答案)
    h["invoice_no"] = f"{series}-{number}" if series and number else number
    h["invoice_date"] = _vn_pdf_date(_search(r"\(Invoice date\):\s*(\S+)", text))
    h["consignee"] = _search(r"\(Customer.s name\):\s*(.+)", text)
    if h["consignee"] and "POUYUEN" in h["consignee"].upper():
        h["consignee"] = "POUYUEN VIETNAM COMPANY LIMITED"   # 答案用英文名稱 (2026-11-06)
    h["onboard_date"] = _vn_pdf_date(_search(r"\(Departure date\):\s*(\S+)", text))
    h["arrive_date"] = _vn_pdf_date(_search(r"\(Arrival date\):\s*(\S+)", text))
    # 欄位空白時，同一行接著是下一個標籤 ("Số lô hàng (File number): Số chứng từ nội bộ…"、
    # "(AWB/SWB): Tiền tệ…")，不可當成值 (2026-10-28 抽樣 KEUHNE+NAGEL_JUL.2026_BY SHIPMENT 第 44 頁)
    blank = r"(?!Số\b|Tiền\b|Ngày\b)"
    # 答案 (2026-11-06)：Order No. = KN No. (內部單號)；B/L = 批號 File number；MBL = AWB 的 "/" 後半
    h["order_no"] = _search(r"\(KN No\.\):[ \t]*" + blank + r"(\S+)", text)
    m = re.search(r"\(AWB/SWB\):[ \t]*" + blank + r"(\S+)[^\n]*\n(\d+)\s*$", text, re.MULTILINE)
    awb = (m.group(1) + m.group(2)) if m and m.group(1).endswith("-") else _search(
        r"\(AWB/SWB\):[ \t]*" + blank + r"(\S+)", text)
    h["bl_no"] = _search(r"\(File number\):[ \t]*" + blank + r"(\S+)", text)
    if awb and "/" in awb:
        h["mbl_no"] = awb.split("/", 1)[1]
    elif awb and not h["bl_no"]:
        h["bl_no"] = awb
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
    if cw and not boxes:
        # 空運：照答案寫計費重 CW，千分位逗號 (例如 3,587.5 CW，2026-11-06 新答案)
        h["volume"] = f"{float(cw):,g} CW"
    return {"header": h, "items": items or [{"description": None, "amount": None, "vat": None}]}


def _kn_vn_totals(text: str):
    """(未稅, 含稅)：電子發票的 "Tổng cộng tiền thanh toán" 下一行；附件清單每票一組。"""
    rows = _kn_vn_list_rows(text)
    if rows:
        return ({r["file"]: r["total"] for r in rows}, {r["file"]: r["total"] + r["vat"] for r in rows})
    m = re.search(r"Tổng cộng tiền thanh toán\s*\n\s*([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)", text)
    if not m:
        return (None, None)
    total, grand = _vn_pdf_num(m.group(1)), _vn_pdf_num(m.group(3))
    # 這一頁的總額要歸給這一頁的發票：發票號碼 = Series + Invoice No.，原文把兩者印在
    # 不同行，合併字串找不到完整比對，用字典 {發票號碼: 總額} 直接指定歸屬
    # (2026-11-11 修正：108 張全部誤判成「沒有可勾稽的總額」)
    series = _search(r"\(Series\):\s*(\S+)", text)
    number = _search(r"\(Invoice No\.\):\s*(\S+)", text)
    inv_no = f"{series}-{number}" if series and number else number
    if not inv_no:
        return (total, grand)
    return ({inv_no: total}, {inv_no: grand})


register_supplier(
    # 月結 PDF (BY SHIPMENT) 是好幾張電子發票合在一起：每頁一張 (2026-10-28 抽樣)
    "KN_VN", "CÔNG TY TNHH KUEHNE + NAGEL (VIỆT NAM)", detect_kn_vn, parse_kn_vn, multi_page=True,
    extract_total=lambda t: _kn_vn_totals(t)[0], extract_grand_total=lambda t: _kn_vn_totals(t)[1],
    page_text=_kn_vn_pages,
)


# --- VN04 DHL - VNPT EXPRESS (CÔNG TY TNHH CHUYỂN PHÁT NHANH DHL-VNPT) ---------
# 進口運費明細 (IB TARIFF)：一張發票 (Invoice no. SGNIR…) 好幾頁、好幾十票快遞。
# 每票第一行 "26/08/2026 PUS 1973098271 SGN 3 6.00 1,365,383 1,365,383" (日期、起運地、
# HAWB、目的地、件數、重量、運費、淨額)，接著附加費 (FUEL SURCHARGE、GO GREEN…)，
# 再來 "PROD CODE EXPRESS WORLDWIDE nondoc"、"Sender: … Receiver: …"。
# VAT 只有整張一個；稽核：費用加總 = Sub-total，加 VAT = Total VND。
_DHL_VN_MAIN_RE = re.compile(
    # 機場代碼可能有數字 ("S2X"，掃描檔)。件數/重量 0 的稅費單 (例如 LOW VALUE
    # REGULATORY TAX) 這一行本身沒有 REGULAR CHARGE/NET CHARGE 數字，金額只在下面
    # 的附加費行，所以這段要設成選填，不然整行比對失敗、current 沒設到，下面的
    # 附加費行也會被跳過 (使用者 2026-11-09 抽查 "152016 EXP.pdf" SGNIR00379176)
    # PIECE 欄偶爾是字母 "L" 不是數字 (2026-11-09 抽查 "SGNIR00269258-IMPORT.pdf")；
    # ORG 偶爾掉一個字變成只有 2 碼 ("CA"，2026-11-09 抽查 "SGNIR00317248 IMPORT.pdf"
    # 的 "CA 2946902066"，照抄帳單原樣，不猜第 3 碼)，或多一個句點 ("DG."，2026-11-09
    # 抽查 "SGNIR00402876.pdf"；句點不算進代碼本身)
    r"^(?P<date>\d{2}/\d{2}/\d{4})\s+(?P<org>[A-Za-z0-9]{2,3})\.?\s+(?P<hawb>\d{8,})\s+(?P<dest>[A-Za-z0-9]{3})\s+"
    r"(?P<pcs>[A-Z\d]+)\s+(?P<kg>[\d.]+)(?:\s+(?P<regular>[\d,]+)(?:\s+[\d,]+)*\s+(?P<net>[\d,]+))?$")
# 附加費名稱有大小寫混合 ("Direct Signature 150,000 150,000")
# 只有一個金額的附加費 ("SHIPMENT VALUE PROTECTION 354,532") 也算；金額要有千分位或 3 位以內，
# 不會誤抓 "YOUR VAT CODE 3900911265" 這種代碼
_DHL_VN_SUR_RE = re.compile(
    # 名稱本身帶數字/冒號 ("PREMIUM 12:00")：允許 desc 裡出現數字和冒號
    r"^(?P<desc>[A-Z][A-Za-z0-9 \-/:]+?)\s+(?:(?:(?P<regular>[\d,]+)\s+)(?:[\d,]+\s+)*(?P<net>\d{1,3}(?:,\d{3})*)"
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
        common = {"invoice_no": inv, "supplier": "DHL - VNPT EXPRESS LTD.",
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
        "supplier": "DHL - VNPT EXPRESS LTD.",
        "invoice_no": _dhl_vn_invoice_no(text),
        "invoice_date": _vn_pdf_date(_search(r"^Date\s+(\S+)", text)),
        # 掃描字型會把 "DHL contact" 讀成 "DEL contact"；集團公司名稱換成英文
        # (使用者 2026-11-09 規定)
        "consignee": _vn_full_consignee(_search(r"^(C[OÔ]NG TY .+?)\s+D[HE]L contact", text))}
    items, current, start, main_item = [], None, False, None
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
            # 件數/重量 0 的稅費單這一行沒有金額，金額在下面的附加費行 (不建立
            # FREIGHT CHARGE 項目；PROD CODE 改套到那筆附加費，見下方)
            if m.group("net"):
                # 運費名稱 (REGULAR CHARGE) 要等同一票的 PROD CODE 行出現才會填入 (見下方)
                items.append({"description": "FREIGHT CHARGE", "amount": _vn_pdf_num(m.group("net")),
                              "vat": None, **current})
                main_item = items[-1]
            else:
                main_item = None
            continue
        if current is None:
            continue
        # 使用者 2026-11-04 規則 (樣本答案提示列)：REGULAR CHARGE 填 "PROD CODE (Regular Charge)"
        m = re.match(r"PROD CODE\s+(.+)$", line)
        if m:
            if main_item is not None:
                main_item["description"] = f"{m.group(1).strip()} (Regular Charge)"
            continue
        m = re.match(r"Sender:\s*(.+?)\s+Receiver:", line)
        if m:
            for it in items:
                if it.get("bl_no") == current["bl_no"]:
                    it["shipper"] = m.group(1).strip()
            continue
        # Order No. = Shipper ref.，空白填 N/A (樣本答案提示列)；同一票的每一列都帶同一個值
        m = re.match(r"Shipper ref\.\s*(.*)$", line)
        if m:
            for it in items:
                if it.get("bl_no") == current["bl_no"]:
                    it["order_no"] = m.group(1).strip() or None
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
    text = _dhl_vn_clean(text)
    totals = []
    for m in re.finditer(r"^Total VND\b[ \t]*(.*)$", text, re.MULTILINE):
        tail = m.group(1).strip()
        if re.fullmatch(r"[\d,]+", tail):
            totals.append(tail)
            continue
        # 頁尾版面偶爾跟地址欄交錯，金額不在同一行 (2026-11-09 抽查 "SGNIR00324691.pdf"：
        # "Total VND" 後面接著 "37-10, 37-11," 等地址碎片，金額隔幾行才出現)，往下
        # 一小段範圍內找第一個純數字列
        num = re.search(r"^[ \t]*([\d,]+)[ \t]*$", text[m.end():m.end() + 300], re.MULTILINE)
        if num:
            totals.append(num.group(1))
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
    h = {"supplier": "DHL - VNPT EXPRESS LTD.",
         "invoice_no": _search(r"INVOICE NO\s+(\S+)", text),
         "invoice_date": _format_date_en(*date.groups()) if date else None,
         # "CONG TY TNHH POUHUNG VIET NAM DHL CONTACT INVOICING HOTLINE" (左右兩欄同一行)；
         # 集團公司名稱換成英文 (使用者 2026-11-09 規定)
         "consignee": _vn_full_consignee(_search(r"^(C[OÔ]NG TY [A-Z ]+?)\s+DHL CONTACT", text))}
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
# 描述太長會換成三行：描述前半 (獨立一行)、序號+單位+數字 (描述整個漏掉)、描述後半
# (獨立一行)，例如 "Phí điều hành bến bãi cho nguyên công-" / "4 CTR 1 ..." / "tơ-nơ"
# (2026-11-11 修正：這筆費用原本整筆漏抓，造成該張發票未稅總額勾稽對不上)。
# 比對前先補回成一行，前後兩行不可以是 "Chứng từ số:"/"Thu hộ" 這類標籤行。
_APL_VN_WRAPPED_DESC_RE = re.compile(
    r"\n(?P<prefix>(?!Chứng từ|Thu hộ)[^\n\d][^\n]*)\n(?P<num>\d+)\s+(?P<unit>CTR|FCR|CBM|KGS)\s+"
    r"(?P<rest>[\d.,]+\s+[\d.,]+\s+[\d.,]+\s+(?:\d+%|KCT|X)\s+[\d.,]+\s+[\d.,]+)\n"
    r"(?P<suffix>(?!Chứng từ|Thu hộ)[^\n\d]+)\n")


def _apl_vn_fix_wrapped_desc(text: str) -> str:
    return _APL_VN_WRAPPED_DESC_RE.sub(
        lambda m: f"\n{m.group('num')} {m.group('prefix').strip()}{m.group('suffix').strip()} "
                  f"{m.group('unit')} {m.group('rest')}\n", text)


def detect_apl_vn(text: str) -> bool:
    upper = text.upper()
    # 舊字型的卡車表寫成 "APL LOGISTICS VIEÄT NAM" (2026-10-28 抽樣)
    return "APL LOGISTICS VIỆT NAM" in upper or "APL LOGISTICS V" in upper and (
        "POUHUNG" in upper or "POUYUEN" in upper or "CONTSEAL" in upper)


def parse_apl_vn(text: str) -> Dict:
    text = _apl_vn_fix_wrapped_desc(text)
    h: Dict[str, Optional[str]] = {"supplier": "APL Logistics Vietnam Co., Ltd"}
    series = _search(r"\(Serial\):\s*(\S+)", text)
    number = _search(r"S[ốo] \(No\.\):\s*(\d+)", text)
    # 發票符號 + "-" + 號碼 (例如 1C26TAT-26716，2026-11-06 新答案)
    h["invoice_no"] = f"{series}-{number}" if series and number else number
    m = re.search(r"Ngày \(date\)\s*(\d{1,2})\s*tháng \(month\)\s*(\d{1,2})\s*năm \(year\)\s*(\d{4})", text)
    h["invoice_date"] = _format_date_en(*m.groups()) if m else None
    h["consignee"] = _search(r"\(Company name\):\s*(.+)", text)
    if h["consignee"] and "POUYUEN" in h["consignee"].upper():
        h["consignee"] = "POUYUEN VIETNAM COMPANY LIMITED"   # 答案用英文 (2026-11-06)
    if h["consignee"] and "POU HUNG" in h["consignee"].upper():
        h["consignee"] = "POU HUNG VIETNAM COMPANY LIMITED"
    if h["consignee"] and "POU LI" in h["consignee"].upper():
        h["consignee"] = "POU LI VIETNAM COMPANY LIMITED"
    if h["consignee"] and "POU SUNG" in h["consignee"].upper():
        h["consignee"] = "POU SUNG VIETNAM COMPANY LIMITED"
    if h["consignee"] and "POU CHEN" in h["consignee"].upper():
        h["consignee"] = "POU CHEN VIETNAM ENTERPRISE LTD."
    air = "HAWB" in text
    hawb = _search(r"HAWB:\s*(\d+)", text)
    # 一張發票可能有好幾個運送單 ("Chứng từ số: 956057, BINH DUONG - …")，每個運送單各自的
    # 單號、船名、港口、櫃號、材積，寫在該運送單的每一列費用上 (2026-11-06 新答案)
    # 運送單標題：「Chứng từ số: N, 船名…」或單行「Chứng từ số: N」(接著是費用)；列表標題 (N; N;) 不算
    starts = list(re.finditer(r"Chứng từ số:\s*(\S+?)(?:,\s*|\s*\n)", text))
    # "Thu hộ APL Logistics Ltd: Chứng từ số: xxx (FCR : nnn)" 區塊：號碼跟緊接著的下一個
    # 正常 "Chứng từ số: xxx" 區塊相同，那個區塊底下的費用要加註代收代付 (暫定文字，
    # 2026-11-10 VN11 FEB.2025 答案)；(FCR : nnn) 不填，這裡故意不擷取
    thu_ho_orders = set(m.group(1) for m in re.finditer(
        r"Thu hộ APL Logistics Ltd:\s*Chứng từ số:\s*(\S+?)(?:\s*\(|,\s*|\s*\n)", text))
    items = []
    for i, sm in enumerate(starts or [None]):
        if sm is None:
            order, rest = None, text
        else:
            order = sm.group(1)
            end = starts[i + 1].start() if i + 1 < len(starts) else len(text)
            rest = text[sm.end():end]
        item_m = re.search(r"\n\s*\d+\s+[^\W\d]", rest)
        route_txt = rest[:item_m.start()] if item_m else rest
        route = " ".join(route_txt.split())
        container = None
        cm = re.search(r"\b(\d{2}[A-Z])\s*-\s*(\d{5})\b", route) or re.search(r"\b(\d{2}[A-Z]\d{5,6})\b", route)
        if cm:
            container = f"{cm.group(1)}-{cm.group(2)}" if cm.lastindex == 2 else cm.group(1)
            route = route[:cm.start()] + route[cm.end():]
        parts = [re.sub(r"^[\s\-.]+|[\s\-.]+$", "", p) for p in route.split(" - ")]
        parts = [p for p in parts if p]
        vessel = pol = pod = None
        if len(parts) >= 3:
            vessel, pol, pod = parts[0], parts[1], parts[2]
            if air:
                vessel = f"{vessel} (Flight)"
        elif len(parts) == 2:
            pol, pod = parts[0], parts[1].split("-")[0].strip()
            if not container:
                pass
        # 清理：港口前面的 "ROUTE:" 字樣、港口後面黏著的櫃號 (版面合併時會黏在一起)
        def _clean_port(v):
            if not v:
                return v
            v = v.split("ROUTE:")[-1]
            v = re.sub(r"-?\s*\d{2}[A-Z]-?\d{5,6}.*$", "", v)
            return v.strip(" -.") or None
        pol, pod = _clean_port(pol), _clean_port(pod)
        if pol and not container:
            mc2 = re.search(r"\d{2}[A-Z]\d{5,6}", (route_txt or ""))
            container = mc2.group(0) if mc2 else None
        ship = {"order_no": order, "vessel": vessel, "port_of_loading": pol,
                "port_of_discharge": pod, "container_no": container,
                "bl_no": hawb if air else None}
        suffix = " - Thu hộ APL Logistics Ltd" if order in thu_ho_orders else ""
        cbm = kgs = None
        sizes: Dict[str, int] = {}
        block_items = []
        for im in _APL_VN_ITEM_RE.finditer(rest):
            desc = im.group("desc").strip() + suffix
            qty = im.group("qty").replace(".", "").replace(",", ".")
            if im.group("unit") == "CBM":
                cbm = qty
            elif im.group("unit") == "KGS":
                kgs = qty
            # 櫃型後面可能接一個單引號 ("45'")，不算進櫃型文字裡 (2026-11-10 FEB.2025 答案 7144)
            box = re.search(r"công-tơ-nơ\s+(\d{2}\w*)'?$", im.group("desc").strip())
            if box and im.group("unit") == "CTR":
                # 櫃數用這一列自己的 Số lượng，不可寫死 1 個 (2026-11-10 答案修正
                # 27435/27438/28051/28171 這 4 張：同種櫃型不只 1 個)
                sizes[box.group(1)] = sizes.get(box.group(1), 0) + int(float(qty) or 1)
            block_items.append({"description": desc, "amount": _dsv_vn_dot_amount(im.group("amount")),
                                "vat": _dsv_vn_dot_amount(im.group("tax"))})
        if cbm:
            volume = f"{cbm} CBM"
        elif kgs:
            volume = f"{kgs} KG"
        elif sizes:
            volume = ", ".join(f"{n}X{sz}" for sz, n in sizes.items())
        else:
            volume = None
        for it in block_items:
            items.append({**it, **ship, "volume": volume})
    return {"header": h, "items": items or [{"description": None, "amount": None, "vat": None}]}


def _apl_vn_totals(text: str):
    m = re.search(r"Tổng cộng \(Total\):\s*([\d.]+)\s+([\d.]+)\s+([\d.]+)", text)
    if not m:
        return (None, None)
    total, grand = _dsv_vn_dot_amount(m.group(1)), _dsv_vn_dot_amount(m.group(3))
    # 這一頁的總額要歸給這一頁的發票：發票號碼 = 字軌(Serial) + 號碼(No.)，但原文把
    # 兩者印在不同行，合併後的字串找不到完整比對，用字典 {發票號碼: 總額} 直接指定
    # 歸屬，不靠原文字串比對 (2026-11-11 修正：117 張全部誤判成「沒有可勾稽的總額」)
    series = _search(r"\(Serial\):\s*(\S+)", text)
    number = _search(r"S[ốo] \(No\.\):\s*(\d+)", text)
    inv_no = f"{series}-{number}" if series and number else number
    if not inv_no:
        return (total, grand)
    return ({inv_no: total}, {inv_no: grand})


def _apl_vn_pages(pdf_path: str) -> List[str]:
    """續頁 ("tiep theo trang truoc - Trang 2 /2") 併回前一頁：同一張發票的最後幾筆運送單
    印在第二頁，沒有發票號碼，逐頁辨識會漏掉 (2026-11-06)。"""
    merged: List[str] = []
    for page in extract_pages_from_pdf(pdf_path):
        if merged and re.search(r"tiep theo trang truoc", page, re.IGNORECASE):
            merged[-1] = merged[-1] + "\n" + page
        else:
            merged.append(page)
    return merged


def apl_vn_invoice_page(page: str) -> bool:
    return "HÓA ĐƠN GIÁ TRỊ GIA TĂNG" in page.upper() and "APL" in page.upper()


_APL_VN_SHEET = [
    # POU HUNG 對帳表 (VND 工作表，2026-11-06 新答案)：只取 DET 欄，費用名稱 = "DET - " + 備註；
    # 發票號碼 = Invoice No. (整張，表尾固定在同一欄，見 doc_key_col)；沒有 Invoice No. 時改用
    # 每一列自己的 Inv# 欄 (原文照填含 "+")，連 Inv# 欄都沒有就填 N/A，不要退回用 Booking No.
    # 當發票號碼 (2026-11-10 VN11 答案：01-2025 TRUCK EXPORT 有 Inv# 欄、POUYUEN DN 沒有)；
    # Order No. = Local invoice No.；B/L = Booking No.；櫃型寫成 "1X40HC" (2026-11-10 使用者
    # 決定跟其他供應商統一用這個格式，不用 VN11 答案原文的「1 x 40HC」)
    (r".", {"header": "BOOKING NO", "key": ["BOOKING NO"], "inv_col": "INV#",
            # CHARGE 底下不只 DET 一欄：Chassis contseal、Drop off TCIT/VICT、Customs/others
            # 或 Waiting/others 都要各自拆一筆費用 (2026-11-10 VN11 答案：01-2025 TRUCK EXPORT
            # 同一列可能同時有 Drop off TCIT/VICT 跟 DET 兩筆費用)
            "fees": [("CHASSIS CONTSEAL", "Chassis contseal - {desc_col}"),
                     ("DROP OFF TCIT/VICT", "Drop off TCIT/VICT - {desc_col}"),
                     ("DET", "DET - {desc_col}"),
                     ("CUSTOMS / OTHERS", "Customs / others - {desc_col}"),
                     ("WAITING / OTHERS", "Waiting / others - {desc_col}")],
            "desc_col": "REMARKS",
            "total": "TOTAL ( VND )", "doc_key": r"Invoice No\.:", "doc_key_col": 12,
            "order_doc": r"Local invoice No\.:", "order_doc_col": 12, "doc_date": "INVOICE DATE",
            "vat_total": "TOTAL VAT AMOUNT", "vat_desc": "TOTAL VAT AMOUNT IN VND",
            "fields": {"container_no": "CONTAINER NUMBER", "bl_no": "BOOKING NO"},
            "volume_col": 3, "volume_prefix": "1X", "vat_last": True,
            "vat_na_fields": ("bl_no", "volume", "container_no", "port_of_loading", "port_of_discharge",
                              "vessel", "shipper", "onboard_date", "arrive_date")}),
]
_apl_vn_sheet_parse, _apl_vn_sheet_totals = _vt_make("APL Logistics Vietnam Co., Ltd", _APL_VN_SHEET)

register_supplier(
    "APL_VN", "CÔNG TY TNHH APL LOGISTICS VIỆT NAM", detect_apl_vn, parse_apl_vn, multi_page=True,
    extract_total=lambda t: _apl_vn_totals(t)[0], extract_grand_total=lambda t: _apl_vn_totals(t)[1],
    sheet_parse=_apl_vn_sheet_parse, sheet_totals=_apl_vn_sheet_totals,
    page_text=lambda pdf_path: _apl_vn_pages(pdf_path),
)


