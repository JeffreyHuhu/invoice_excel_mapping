# -*- coding: utf-8 -*-
"""
validation_streamlit.py
=======================
擷取結果的驗證層 (Validation)

問題背景：供應商 PDF 格式千變萬化，規則不可能 100% 抓對，而且不可能替
每一份帳單都人工做正確答案 Excel。所以這裡改用「帳單本身就該成立的規則」
自動檢查擷取結果合不合理，不需要正確答案：

  1. 關鍵欄位缺失：發票號碼/發票日期/供應商名稱抓不到、沒有任何費用明細、
     有費用名稱卻沒有金額。
  2. 格式檢查：金額/稅額不是數字、日期格式認不得、日期不合理 (太舊或在
     未來)、到達日早於開船日。
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
from typing import Dict, List, Optional

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

ERROR = "ERROR"
WARN = "WARN"
INFO = "INFO"

# 每個問題從信心度 1.0 扣掉的分數
_SEVERITY_PENALTY = {ERROR: 0.35, WARN: 0.1, INFO: 0.05}

# 少了就幾乎無法入帳的抬頭欄位 (ERROR)，以及缺了值得注意的欄位 (WARN)。
# 發票日期歸在 WARN：少數供應商的帳單本身就沒有印發票日期 (例如 ECOMET)。
REQUIRED_HEADER_FIELDS = ("invoice_no", "supplier")
RECOMMENDED_HEADER_FIELDS = ("invoice_date", "consignee")

# 日期合理範圍：早於這一年，或晚於今天 + 這麼多天，視為可疑
_MIN_REASONABLE_YEAR = 2015
_MAX_FUTURE_DAYS = 60

LEVEL_HIGH = "高"
LEVEL_MEDIUM = "中"
LEVEL_LOW = "低"

FIELD_LABELS: Dict[str, str] = {
    "invoice_no": "發票號碼", "invoice_date": "發票日期", "supplier": "供應商名稱",
    "consignee": "集團公司名稱", "description": "費用名稱", "amount": "金額",
    "vat": "稅額", "arrive_date": "到達日", "onboard_date": "開船日",
    "total": "金額總計", "supplier_key": "供應商辨識", "extraction": "擷取方式",
}


def _issue(severity: str, field: str, rule: str, message: str) -> Dict[str, str]:
    return {"severity": severity, "field": field, "rule": rule, "message": message}


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
        if _is_na(row.get(code)):
            issues.append(_issue(ERROR, code, "MISSING_KEY_FIELD",
                                 f"關鍵欄位「{FIELD_LABELS[code]}」抓不到值"))
    for code in RECOMMENDED_HEADER_FIELDS:
        if _is_na(row.get(code)):
            issues.append(_issue(WARN, code, "MISSING_FIELD",
                                 f"欄位「{FIELD_LABELS[code]}」抓不到值"))

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


def _check_items(group_rows: List[Dict]) -> List[Dict]:
    issues = []
    has_any_item = False
    for idx, row in enumerate(group_rows, start=1):
        desc, amount_raw, vat_raw = row.get("description"), row.get("amount"), row.get("vat")
        if _is_na(desc) and _is_na(amount_raw):
            continue
        has_any_item = True
        label = f"第 {idx} 筆費用「{desc if not _is_na(desc) else NA}」"

        amount = _to_number(amount_raw)
        vat = _to_number(vat_raw)
        if amount is None:
            # 稅金列 (PPN/VAT/PPh) 的慣例是金額 N/A、稅額填在 VAT 欄，不算缺漏
            if vat is None:
                issues.append(_issue(ERROR, "amount", "AMOUNT_MISSING", f"{label}沒有金額"))
        elif amount != amount:  # NaN
            issues.append(_issue(ERROR, "amount", "AMOUNT_NOT_NUMERIC",
                                 f"{label}的金額不是數字：{amount_raw}"))
        elif amount == 0:
            # 金額 0 的費用很常見 (例如 MATERAI 印花稅免收)，只當附註
            issues.append(_issue(INFO, "amount", "AMOUNT_ZERO", f"{label}金額為 0"))
        elif amount < 0:
            issues.append(_issue(WARN, "amount", "AMOUNT_NEGATIVE",
                                 f"{label}金額為負數 ({amount_raw})，請確認是否為折讓/扣款"))

        if vat is not None and vat != vat:
            issues.append(_issue(WARN, "vat", "VAT_NOT_NUMERIC",
                                 f"{label}的稅額不是數字：{vat_raw}"))
        elif vat is not None and amount and amount == amount and abs(vat) > abs(amount):
            issues.append(_issue(WARN, "vat", "VAT_EXCEEDS_AMOUNT",
                                 f"{label}稅額 ({vat_raw}) 大於金額 ({amount_raw})"))

    if not has_any_item:
        # WARN 而非 ERROR：少數帳單 (例如 TATA A/F 明細表) 有些發票本來就沒有費用列
        issues.append(_issue(WARN, "description", "NO_ITEMS", "沒有抓到任何費用明細"))
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
        ("final_match", "final_diff", "TOTAL_FINAL_MISMATCH", "含稅 - 代扣稅與帳單實付金額不一致"),
    )
    for match_key, diff_key, rule, message in checks:
        if audit.get(match_key) is False:
            issues.append(_issue(ERROR, "total", rule, f"{message} (差額 {audit.get(diff_key)})"))
    return issues


def _confidence(issues: List[Dict]) -> float:
    score = 1.0 - sum(_SEVERITY_PENALTY[i["severity"]] for i in issues)
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
      issues      問題清單 (每筆含 severity/field/rule/message)
      confidence  信心度 0~1
      level       信心等級 高/中/低

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
        confidence = _confidence(issues)
        results.append({
            "invoice_no": group_rows[0].get("invoice_no", NA),
            "issues": issues,
            "confidence": confidence,
            "level": _level(issues, confidence),
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
