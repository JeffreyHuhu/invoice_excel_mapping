# -*- coding: utf-8 -*-
"""
validation_streamlit.py
=======================
擷取結果的驗證層 (Validation)

問題背景：供應商 PDF 格式千變萬化，規則不可能 100% 抓對，而且不可能替
每一份帳單都人工做正確答案 Excel。所以這裡改用「帳單本身就該成立的規則」
自動檢查擷取結果合不合理，不需要正確答案：

  1. 必填欄位缺失 (ERROR)：INVOICE DATE、INVOICE NO、SUPPLIER、CONSIGNEE、
     Description、AMOUNT (使用者指定的六個必要欄位)。
  2. 格式檢查：金額/稅額不是數字、日期格式認不得、日期不合理 (太舊或在
     未來)、到達日早於開船日；選填欄位的基本格式 (單號要有數字、貨櫃號碼
     要有 4 碼英文 + 7 碼數字、材積要有數字、港口/船名/出貨方要有文字)。
  3. 金額勾稽：沿用 audit_invoice_totals() 的三項稽核 (未稅加總=未稅總額、
     未稅+稅額=含稅總額、含稅-代扣稅=實付金額)，任何一項不一致就是錯誤。

每個問題分三種嚴重度：
  ERROR  幾乎確定抓錯，一定要人工複查 (信心等級直接降為「低」)
  WARN   可疑，建議複查
  INFO   不是錯誤，但降低信心 (例如這家供應商還沒有總額稽核規則可以交叉確認)

每張發票算出一個 0~1 的信心度與「高/中/低」信心等級，低信心的發票會觸發
fallback_streamlit.py 的備援擷取，並列入異常清單 (Exceptions)。

這支檔案只讀取擷取結果，不修改任何供應商規則或核心引擎的行為。
"""

import datetime
import io
import re
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from extract_utils_streamlit import (
    NA,
    DATE_FIELD_CODES,
    SUPPLIER_REGISTRY,
    _coerce_to_date,
    _group_rows_by_invoice,
    _is_na,
    normalize_value,
)
from logger_streamlit import get_logger

ERROR = "ERROR"
WARN = "WARN"
INFO = "INFO"

# 每個問題從信心度 1.0 扣掉的分數
_SEVERITY_PENALTY = {ERROR: 0.35, WARN: 0.1, INFO: 0.05}

# 使用者指定的必要欄位：抬頭 4 個 + 費用明細 2 個 (description/amount，見
# _check_items())。缺任何一個都是 ERROR。其餘 12 個欄位為選填，缺值不算
# 異常，有值時才做基本格式檢查 (_check_optional_formats())。
REQUIRED_HEADER_FIELDS = ("invoice_date", "invoice_no", "supplier", "consignee")

# 選填欄位的基本格式規則 (WARN)：(欄位, 規則代碼, 檢查函式, 說明)
_CONTAINER_NO_RE = re.compile(r"[A-Z]{4}\d{6,7}")
_OPTIONAL_FORMAT_RULES = (
    ("order_no", "ID_NO_DIGIT", lambda v: re.search(r"\d", v), "單號裡沒有任何數字"),
    ("bl_no", "ID_NO_DIGIT", lambda v: re.search(r"\d", v), "提單號碼裡沒有任何數字"),
    ("mbl_no", "ID_NO_DIGIT", lambda v: re.search(r"\d", v), "主提單號碼裡沒有任何數字"),
    # 貨櫃號碼可能是多個、中間有空白或換行 (例如 "BMOU 4921920")，去掉空白後
    # 只要找得到至少一組「4 碼英文 + 6~7 碼數字」就算合格
    ("container_no", "CONTAINER_FORMAT",
     lambda v: _CONTAINER_NO_RE.search(re.sub(r"\s", "", v.upper())),
     "找不到「4 碼英文 + 7 碼數字」格式的貨櫃號碼"),
    ("volume", "VOLUME_NO_DIGIT", lambda v: re.search(r"\d", v), "材積裡沒有任何數字"),
    ("port_of_loading", "TEXT_NO_LETTER", lambda v: re.search(r"[A-Za-z]", v), "啟運港沒有任何文字"),
    ("port_of_discharge", "TEXT_NO_LETTER", lambda v: re.search(r"[A-Za-z]", v), "目的港沒有任何文字"),
    ("vessel", "TEXT_NO_LETTER", lambda v: re.search(r"[A-Za-z]", v), "船名/航班沒有任何文字"),
    ("shipper", "TEXT_NO_LETTER", lambda v: re.search(r"[A-Za-z]", v), "出貨方沒有任何文字"),
)

# 正確答案裡常見的「不適用」佔位字元，視同空值 (不算缺漏，也不做格式檢查)
_PLACEHOLDERS = {"-", "--", "/"}

# 已知的「原帳單本身的問題」：使用者核對過原帳單，確認不是擷取錯誤。使用者
# 規定原帳單的問題不標示為 WARN、不反紅：改成 INFO (只當附註、不扣信心分數、
# 不影響信心等級)，說明後面註記屬原帳單問題。
#   供應商 key -> {(欄位代碼, 規則代碼): (說明, 只在這些原始值時才算；None=不限)}
KNOWN_SOURCE_ISSUES: Dict[str, Dict[Tuple[str, str], Tuple[str, Optional[set]]]] = {
    "MAERSK": {
        ("arrive_date", "DATE_FORMAT"): ("MAERSK 原帳單的 ETA 本身就是「--」", {"--"}),
        ("arrive_date", "DATE_ORDER"): ("MAERSK 原帳單的 ETA 本身就早於 ETD", None),
    },
    "CONG_THANH": {
        ("container_no", "CONTAINER_FORMAT"): ("CONG THANH 對帳單的 CONT/XE 欄是車牌或簡寫櫃號，照抄原文", None),
    },
    "RONG_A_CHAU": {
        ("invoice_date", "MISSING_REQUIRED"): ("RỒNG Á CHÂU 的 Excel 對帳表本身沒有發票日期", None),
        # 擷取規則只在費用金額不為 0 時才轉成一列，金額空白的只會是使用者規定
        # 「進口每筆出貨都要列」的 VAT / VAT PASS THROUGH 列 (稅額 0 時金額、稅額都 N/A)
        ("amount", "MISSING_REQUIRED"): ("進口每筆出貨都列 VAT 列，VAT PASS THROUGH 為 0 時金額與稅額皆為 N/A", None),
        ("bl_no", "ID_NO_DIGIT"): ("國內段 (Local) 出貨的提單欄本來就寫 LOCAL", {"LOCAL", "Local", "local"}),
    },
}

# 日期合理範圍：早於這一年，或晚於今天 + 這麼多天，視為可疑
_MIN_REASONABLE_YEAR = 2015
_MAX_FUTURE_DAYS = 60

LEVEL_HIGH = "高"
LEVEL_MEDIUM = "中"
LEVEL_LOW = "低"

FIELD_LABELS: Dict[str, str] = {
    "invoice_date": "INVOICE DATE 發票日期", "invoice_no": "INVOICE NO 發票號碼",
    "supplier": "SUPPLIER 供應商名稱", "consignee": "CONSIGNEE 集團公司名稱",
    "description": "Description 費用名稱", "amount": "AMOUNT 金額", "vat": "VAT 營業稅額",
    "order_no": "Order No. 訂單號碼", "arrive_date": "Arrive Date 到達日",
    "onboard_date": "on board date 開船日", "bl_no": "B/L NO. 提單號碼",
    "mbl_no": "MBL NO. 主提單號碼", "port_of_loading": "Port of Loading 啟運港",
    "port_of_discharge": "Port of discharge 目的港", "volume": "Volume 材積",
    "vessel": "Vessel 船名", "container_no": "Container no. 貨櫃號碼",
    "shipper": "Shipper 出貨方",
    "total": "金額總計", "supplier_key": "供應商辨識", "extraction": "擷取方式",
}


def _issue(severity: str, field: str, rule: str, message: str,
           row: Optional[int] = None) -> Dict:
    """row：問題出在這張發票的第幾筆費用 (0 起算)；抬頭/整張發票層級的問題
    不帶 row。本機轉檔工具用它把 Excel 裡對應的儲存格反紅。"""
    issue = {"severity": severity, "field": field, "rule": rule, "message": message}
    if row is not None:
        issue["row"] = row
    return issue


def _is_blank(value) -> bool:
    """N/A、空值，或「-」這類不適用佔位字元。"""
    return _is_na(value) or str(value).strip() in _PLACEHOLDERS


def _check_optional_formats(row: Dict) -> List[Dict]:
    """選填欄位有值時的基本格式檢查。"""
    issues = []
    for code, rule, check, message in _OPTIONAL_FORMAT_RULES:
        raw = row.get(code)
        if _is_blank(raw):
            continue
        if not check(str(raw)):
            issues.append(_issue(WARN, code, rule, f"「{FIELD_LABELS[code]}」{message}：{raw}"))
    return issues


def _to_number(value) -> Optional[float]:
    """金額/稅額轉成數字；N/A 回傳 None；轉不出來 (不是數字) 回傳 NaN 標記。"""
    if _is_na(value):
        return None
    if isinstance(value, bool):
        return float("nan")
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").strip())
    except ValueError:
        return float("nan")


def _check_header(row: Dict) -> List[Dict]:
    issues = []
    for code in REQUIRED_HEADER_FIELDS:
        if _is_blank(row.get(code)):
            issues.append(_issue(ERROR, code, "MISSING_REQUIRED",
                                 f"必填欄位「{FIELD_LABELS[code]}」抓不到值"))
    issues.extend(_check_optional_formats(row))

    today = datetime.date.today()
    parsed_dates: Dict[str, datetime.date] = {}
    for code in DATE_FIELD_CODES:
        raw = row.get(code)
        if _is_na(raw):
            continue
        date_value = _coerce_to_date(raw)
        if date_value is None:
            issues.append(_issue(WARN, code, "DATE_FORMAT",
                                 f"「{FIELD_LABELS[code]}」日期格式無法辨識：{raw}"))
            continue
        parsed_dates[code] = date_value
        if date_value.year < _MIN_REASONABLE_YEAR or \
                date_value > today + datetime.timedelta(days=_MAX_FUTURE_DAYS):
            issues.append(_issue(WARN, code, "DATE_OUT_OF_RANGE",
                                 f"「{FIELD_LABELS[code]}」日期不合理：{raw}"))

    onboard, arrive = parsed_dates.get("onboard_date"), parsed_dates.get("arrive_date")
    if onboard and arrive and arrive < onboard:
        issues.append(_issue(WARN, "arrive_date", "DATE_ORDER",
                             f"到達日 ({row.get('arrive_date')}) 早於開船日 ({row.get('onboard_date')})"))
    return issues


# Excel 帳單的合計列公式錯誤 (呼叫端檢查後放進 extra_issues，見
# extract_utils.excel_formula_issues())
RULE_EXCEL_FORMULA = "EXCEL_FORMULA_RANGE"


def _explain_by_formula_issue(issues: List[Dict]) -> List[Dict]:
    """Excel 帳單的查驗順序：先查 Excel 公式，再查每票，最後查整份帳單合計。
    已經查到合計列公式錯誤時，整份帳單合計對不上就是這個錯誤造成的：保留
    公式錯誤 (ERROR)，合計不一致改成 INFO 並註明原因，不重複算成兩個錯誤。"""
    if not any(i["rule"] == RULE_EXCEL_FORMULA for i in issues):
        return issues
    explained = []
    for issue in issues:
        if issue["rule"] == "TOTAL_FINAL_MISMATCH":
            issue = dict(issue, severity=INFO,
                         message=f"{issue['message']}（原因：帳單 Excel 合計列公式漏算，見公式錯誤）")
        explained.append(issue)
    return explained


def _mark_known_source_issues(issues: List[Dict], supplier_key: Optional[str],
                              row: Dict) -> List[Dict]:
    """符合 KNOWN_SOURCE_ISSUES 的問題改成 INFO，說明後面註記屬原帳單問題
    (另帶 known_source_issue=True，_confidence() 不扣分)。回傳新的 list，
    不改動原本的 dict。"""
    known = KNOWN_SOURCE_ISSUES.get(supplier_key or "", {})
    if not known:
        return issues
    marked = []
    for issue in issues:
        entry = known.get((issue["field"], issue["rule"]))
        if entry:
            note, only_values = entry
            if only_values is None or str(row.get(issue["field"], "")).strip() in only_values:
                issue = dict(issue, severity=INFO, known_source_issue=True,
                             message=f"{issue['message']}（{note}，屬原帳單問題，非擷取錯誤）")
        marked.append(issue)
    return marked


def _check_items(group_rows: List[Dict]) -> List[Dict]:
    issues = []
    has_any_item = False
    for idx, row in enumerate(group_rows, start=1):
        desc, amount_raw, vat_raw = row.get("description"), row.get("amount"), row.get("vat")
        if _is_na(desc) and _is_na(amount_raw):
            continue
        has_any_item = True
        label = f"第 {idx} 筆費用「{desc if not _is_na(desc) else NA}」"
        row = idx - 1
        if _is_blank(desc):
            issues.append(_issue(ERROR, "description", "MISSING_REQUIRED",
                                 f"第 {idx} 筆費用 (金額 {amount_raw}) 沒有「{FIELD_LABELS['description']}」",
                                 row))

        amount = _to_number(amount_raw)
        vat = _to_number(vat_raw)
        if amount is None:
            # 稅金列 (PPN/VAT/PPh) 的慣例是金額 N/A、稅額填在 VAT 欄，不算缺漏
            if vat is None:
                issues.append(_issue(ERROR, "amount", "MISSING_REQUIRED",
                                     f"{label}沒有「{FIELD_LABELS['amount']}」", row))
        elif amount != amount:  # NaN
            issues.append(_issue(ERROR, "amount", "AMOUNT_NOT_NUMERIC",
                                 f"{label}的金額不是數字：{amount_raw}", row))
        elif amount == 0:
            # 金額 0 的費用很常見 (例如 MATERAI 印花稅免收)，只當附註
            issues.append(_issue(INFO, "amount", "AMOUNT_ZERO", f"{label}金額為 0", row))
        elif amount < 0:
            issues.append(_issue(WARN, "amount", "AMOUNT_NEGATIVE",
                                 f"{label}金額為負數 ({amount_raw})，請確認是否為折讓/扣款", row))

        if vat is not None and vat != vat:
            issues.append(_issue(WARN, "vat", "VAT_NOT_NUMERIC",
                                 f"{label}的稅額不是數字：{vat_raw}", row))
        elif vat is not None and amount and amount == amount and abs(vat) > abs(amount):
            issues.append(_issue(WARN, "vat", "VAT_EXCEEDS_AMOUNT",
                                 f"{label}稅額 ({vat_raw}) 大於金額 ({amount_raw})", row))

    if not has_any_item:
        # Description/AMOUNT 是必填欄位，整張發票沒有任何費用明細就是 ERROR
        issues.append(_issue(ERROR, "description", "NO_ITEMS",
                             "沒有抓到任何費用明細 (Description/AMOUNT 為必填)"))
    return issues


def _check_audit(audit: Optional[Dict], supplier_key: Optional[str]) -> List[Dict]:
    """把 audit_invoice_totals() 的結果轉成問題清單。"""
    if audit is None:
        entry = SUPPLIER_REGISTRY.get(supplier_key) if supplier_key else None
        has_audit_rule = bool(entry and (
            entry.get("extract_total") or entry.get("extract_grand_total")
            or entry.get("extract_final_total")))
        if has_audit_rule:
            message = "帳單上找不到總額，無法做金額勾稽"
        else:
            message = "此供應商尚無總額稽核規則，無法做金額勾稽"
        return [_issue(INFO, "total", "NO_TOTAL_AUDIT", message)]

    issues = []
    checks = (
        ("match", "diff", "TOTAL_NET_MISMATCH", "費用加總 (未稅) 與帳單未稅總額不一致"),
        ("grand_match", "grand_diff", "TOTAL_GRAND_MISMATCH", "未稅 + 稅額與帳單含稅總額不一致"),
        ("final_match", "final_diff", "TOTAL_FINAL_MISMATCH",
         "含稅 - 代扣稅與帳單實付金額 (Excel 帳單為表格合計列) 不一致"),
    )
    for match_key, diff_key, rule, message in checks:
        if audit.get(match_key) is False:
            issues.append(_issue(ERROR, "total", rule, f"{message} (差額 {audit.get(diff_key)})"))
    return issues


def _confidence(issues: List[Dict]) -> float:
    # 已知原帳單問題不是擷取錯誤，不扣分
    score = 1.0 - sum(_SEVERITY_PENALTY[i["severity"]] for i in issues
                      if not i.get("known_source_issue"))
    return round(max(score, 0.0), 2)


def _level(issues: List[Dict], confidence: float) -> str:
    if any(i["severity"] == ERROR for i in issues):
        return LEVEL_LOW
    # 有任何 WARN (含 LLM 擷取的「需人工複查」) 就至少是「中」，才會列入異常清單
    if any(i["severity"] == WARN for i in issues):
        return LEVEL_MEDIUM
    return LEVEL_HIGH if confidence >= 0.9 else LEVEL_MEDIUM


def validate_extraction(
    rows: List[Dict],
    supplier_key: Optional[str],
    audit_results: Optional[List[Dict]] = None,
    extra_issues: Optional[List[Dict]] = None,
) -> List[Dict]:
    """驗證一份 PDF 的擷取結果，回傳每張發票各一筆：
      invoice_no  發票號碼 (抓不到時為 N/A)
      issues      問題清單 (每筆含 severity/field/rule/message，費用明細的
                  問題另帶 row)
      confidence  信心度 0~1
      level       信心等級 高/中/低
      rows        這張發票的 rows (跟傳入的 rows 是同一批物件，供對回 Excel 列)

    rows 可能包含好幾張發票 (multi_page 供應商或 LLM 擷取)，依 invoice_no
    分組各自驗證。extra_issues 會附加到每張發票上 (例如「本次為 LLM 擷取」)。
    """
    audit_by_invoice = {}
    for ar in audit_results or []:
        # 跟 _group_rows_by_invoice() 用同一套正規化，才能把稽核結果對回發票
        audit_by_invoice[normalize_value(ar.get("invoice_no"))] = ar

    base_issues = list(extra_issues or [])
    if not supplier_key:
        base_issues.append(_issue(ERROR, "supplier_key", "UNKNOWN_SUPPLIER",
                                  "無法辨識供應商，沒有對應的擷取規則"))

    results = []
    groups = _group_rows_by_invoice(rows) if rows else {"": [{}]}
    for inv_key, group_rows in groups.items():
        issues = list(base_issues)
        issues.extend(_check_header(group_rows[0]))
        issues.extend(_check_items(group_rows))
        issues.extend(_check_audit(audit_by_invoice.get(inv_key), supplier_key))
        issues = _explain_by_formula_issue(issues)
        issues = _mark_known_source_issues(issues, supplier_key, group_rows[0])
        confidence = _confidence(issues)
        results.append({
            "invoice_no": group_rows[0].get("invoice_no", NA),
            "issues": issues,
            "confidence": confidence,
            "level": _level(issues, confidence),
            "rows": group_rows if rows else [],
        })
    return results


def worst_level(validations: List[Dict]) -> str:
    """一份 PDF 裡所有發票中最差的信心等級。"""
    levels = {v["level"] for v in validations}
    for level in (LEVEL_LOW, LEVEL_MEDIUM):
        if level in levels:
            return level
    return LEVEL_HIGH


def error_count(validations: List[Dict]) -> int:
    return sum(1 for v in validations for i in v["issues"] if i["severity"] == ERROR)


def log_issues(file_name: str, supplier_lbl: str, source: str, validations: List[Dict]) -> None:
    """把 ERROR/WARN 等級的異常欄位逐筆寫進 log 檔 (INFO 只是附註，不寫)。"""
    logger = get_logger()
    for v in validations:
        for i in v["issues"]:
            if i["severity"] == INFO:
                continue
            log = logger.error if i["severity"] == ERROR else logger.warning
            log("%s | %s | invoice=%s | 信心=%s | 來源=%s | 欄位=%s | %s | %s",
                file_name, supplier_lbl, v["invoice_no"], v["level"], source,
                FIELD_LABELS.get(i["field"], i["field"]), i["rule"], i["message"])


# ---------------------------------------------------------------------------
# 異常清單 (Exceptions) 輸出
# ---------------------------------------------------------------------------

EXCEPTION_COLUMNS = [
    "來源檔案", "供應商", "Invoice No", "信心等級", "信心度", "擷取來源",
    "嚴重度", "欄位", "檢查項目", "說明",
]


def build_exception_rows(file_name: str, supplier_lbl: str, source: str,
                         validations: List[Dict]) -> List[Dict]:
    """把信心等級不是「高」的發票攤平成異常清單 (每個問題一列)。
    INFO 等級的問題只是附註，不單獨成列，除非這張發票沒有其他問題。
    """
    rows = []
    for v in validations:
        if v["level"] == LEVEL_HIGH:
            continue
        issues = [i for i in v["issues"] if i["severity"] != INFO] or v["issues"]
        for i in issues:
            rows.append({
                "來源檔案": file_name,
                "供應商": supplier_lbl,
                "Invoice No": v["invoice_no"],
                "信心等級": v["level"],
                "信心度": f"{v['confidence'] * 100:.0f}%",
                "擷取來源": source,
                "嚴重度": i["severity"],
                "欄位": FIELD_LABELS.get(i["field"], i["field"]),
                "檢查項目": i["rule"],
                "說明": i["message"],
            })
    return rows


def add_exceptions_sheet(excel_bytes: bytes, exception_rows: List[Dict]) -> bytes:
    """在既有的轉檔結果 Excel 後面加一個「Exceptions」分頁，第一個分頁
    (工作表) 完全不動。沒有異常時也會加上分頁並註明「無異常」。
    """
    wb = load_workbook(io.BytesIO(excel_bytes))
    _write_exceptions_sheet(wb.create_sheet("Exceptions"), exception_rows)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_exceptions_excel(exception_rows: List[Dict]) -> bytes:
    """單獨輸出異常清單 Excel (只有 Exceptions 一個分頁)。"""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Exceptions"
    _write_exceptions_sheet(ws, exception_rows)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _write_exceptions_sheet(ws, exception_rows: List[Dict]) -> None:
    bold = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="DDEBF7")
    fills = {
        ERROR: PatternFill("solid", fgColor="FFC7CE"),
        WARN: PatternFill("solid", fgColor="FFEB9C"),
        INFO: PatternFill("solid", fgColor="E7E6E6"),
    }

    ws.append(EXCEPTION_COLUMNS)
    for cell in ws[1]:
        cell.font = bold
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    if not exception_rows:
        ws.append(["(無異常：所有發票信心等級皆為「高」)"])
    for r in exception_rows:
        ws.append([r.get(c, NA) for c in EXCEPTION_COLUMNS])
        ws.cell(row=ws.max_row, column=EXCEPTION_COLUMNS.index("嚴重度") + 1).fill = \
            fills.get(r.get("嚴重度"), fills[INFO])

    for col_letter, width in zip("ABCDEFGHIJ", (24, 24, 16, 9, 8, 16, 8, 12, 22, 60)):
        ws.column_dimensions[col_letter].width = width
    ws.freeze_panes = "A2"
