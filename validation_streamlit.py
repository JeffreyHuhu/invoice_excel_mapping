# -*- coding: utf-8 -*-
"""
validation_streamlit.py
=======================
擷取結果的驗證層 (Validation)

問題背景：供應商 PDF 格式千變萬化，規則不可能 100% 抓對，而且不可能替
每一份帳單都人工做正確答案 Excel。所以這裡改用「帳單本身就該成立的規則」
自動檢查擷取結果合不合理，不需要正確答案：

  1. 必填欄位缺失 (ERROR)：INVOICE NO、SUPPLIER、CONSIGNEE、Description、AMOUNT。
     INVOICE DATE 缺值例外，只標 WARN，不算重大錯誤 (2026-11-12 使用者最高規則，
     所有供應商適用：原帳單本身常漏填發票日期，不該整張發票被打成低信心)。
  2. 格式檢查：金額/稅額不是數字、日期格式認不得、金額為負數 (帳單有
     CREDIT 字樣的折讓不計入)。
  3. 金額勾稽：沿用 audit_invoice_totals() 的三項稽核 (未稅加總=未稅總額、
     未稅+稅額=含稅總額、含稅-代扣稅=實付金額)，任何一項不一致就是錯誤。
     完全沒有勾稽 (NO_TOTAL_AUDIT，帳單找不到總額或供應商還沒有稽核規則)：
     金額沒有被查核過，信心等級直接降為「低」(2026-11-12 使用者規定，
     所有供應商適用)。

每個問題分三種嚴重度：
  ERROR  幾乎確定抓錯，一定要人工複查 (信心等級直接降為「低」)
  WARN   可疑，建議複查
  INFO   不是錯誤，但降低信心 (NO_TOTAL_AUDIT 例外：見上方第 3 點，直接降為「低」)

每張發票算出一個 0~1 的信心度與「高/中/低」信心等級，低信心的發票會觸發
fallback_streamlit.py 的備援擷取，並列入異常清單 (Exceptions)。

這支檔案只讀取擷取結果，不修改任何供應商規則或核心引擎的行為。
"""

import io
import re
from typing import Dict, List, Optional, Pattern, Tuple, Union

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
# 異常 (選填欄位不做格式檢查)。
REQUIRED_HEADER_FIELDS = ("invoice_date", "invoice_no", "supplier", "consignee")

# 個別供應商另外的必要欄位 (使用者在答案 Excel 提示列標「(必要)」的欄位)：
# {供應商: {欄位: 缺值時的等級}}。"container_no" 只在整櫃 (材積是「數量X櫃型」) 時才必要。
# - KUEHNE_NAGEL (使用者 2026-10-31，樣本 JKT1521476)：Volume 缺值先列 WARN；貨櫃號碼非必要 (2026-11-02 使用者規定：
#   沒抓到櫃號不計入 ERROR，有值仍檢查格式)；
#   沒有 VAT 的帳單只要通過金額稽核就可以，所以 VAT 不列入。
# - DHL / INDOPROSTIME_TRANSPORT / OPL (使用者 2026-10-31，KEY 值範例)：缺值列 WARN (等級最多「中」)。
#   VAT 一律不列入：這幾家都有含稅總額勾稽，帳單有印稅金卻漏抓會勾稽不符 (ERROR)，沒有稅金不扣分。
# 新的 KEY 值範例：用 invoice-answer-check skill 的 required_fields.py 讀出「(必要)」欄位。
_W = "WARN"
SUPPLIER_REQUIRED_FIELDS = {
    "KUEHNE_NAGEL": {
        "arrive_date": "ERROR", "onboard_date": "ERROR", "bl_no": "ERROR",
        "port_of_loading": "ERROR", "port_of_discharge": "ERROR", "vessel": "ERROR",
        "volume": "WARN",
    },
    "DHL": {"onboard_date": _W, "bl_no": _W, "port_of_loading": _W, "port_of_discharge": _W,
            "volume": _W},
    "INDOPROSTIME_TRANSPORT": {"order_no": _W, "bl_no": _W, "mbl_no": _W, "port_of_loading": _W,
                               "volume": _W, "vessel": _W, "shipper": _W},
    "OPL": {"arrive_date": _W, "onboard_date": _W, "bl_no": _W, "mbl_no": _W,
            "port_of_loading": _W, "port_of_discharge": _W, "volume": _W, "vessel": _W,
            "container_no": _W, "shipper": _W},
}
_FCL_VOLUME_RE = re.compile(r"\d+\s*X\s*\d{2}", re.IGNORECASE)
# 必要欄位只檢查 KEY 值範例那種版型 (使用者 2026-10-31)：
# OPL 只檢查 NCLR 開頭的進口帳單；CT-/IP-/AT- 等其他版型本來就沒有主提單、出貨方。
SUPPLIER_REQUIRED_WHEN = {
    "OPL": lambda row: str(row.get("invoice_no") or "").upper().startswith("NCLR"),
}

# 正確答案裡常見的「不適用」佔位字元，視同空值 (不算缺漏)
_PLACEHOLDERS = {"-", "--", "/"}

# 解析器判定「原帳單無法判讀」時，把 description 設成這個標記，驗證層判 ERROR (原因：無法判讀)
UNREADABLE_DESC = "UNREADABLE"

# 已知的「原帳單本身的問題」：使用者核對過原帳單，確認不是擷取錯誤。使用者
# 規定原帳單的問題不標示為 WARN、不反紅：改成 INFO (只當附註、不扣信心分數、
# 不影響信心等級)，說明後面註記屬原帳單問題。
#   供應商 key -> {(欄位代碼, 規則代碼): (說明, 只在這些原始值時才算；None=不限)}
#   原始值也可以是 re.Pattern (整個值符合才算)，給「每張不同、但格式固定」的值用
KNOWN_SOURCE_ISSUES: Dict[str, Dict[Tuple[str, str], Tuple[str, Optional[Union[set, Pattern]]]]] = {
    "MAERSK": {
        ("arrive_date", "DATE_FORMAT"): ("MAERSK 原帳單的 ETA 本身就是「--」", {"--"}),
    },
    # KUEHNE_NAGEL 原帳單的 ETD/ATD、ETA/ATA 欄本來就是空白 (使用者 2026-11-01 確認，
    # 例如 JKT1485198、JKT1517926)：只在值為空白時才算，有值但抓不到仍列 ERROR
    "KUEHNE_NAGEL": {
        ("onboard_date", "MISSING_REQUIRED"): ("KUEHNE NAGEL 原帳單的 ETD/ATD 欄本來就是空白", {"N/A", "None", ""}),
        ("arrive_date", "MISSING_REQUIRED"): ("KUEHNE NAGEL 原帳單的 ETA/ATA 欄本來就是空白", {"N/A", "None", ""}),
    },
    # DHL 印尼帳單：DUTIES & TAXES / INBOUND 版型原帳單沒有開船日、啟運港、目的港、材積、提單號碼
    # (使用者 2026-11-02 確認：列備註，不算錯誤)；只限擷取不到的空值，有值仍照常檢查
    "DHL": {
        ("onboard_date", "MISSING_REQUIRED"): ("DHL 原帳單沒有開船日", {"N/A", "None", ""}),
        ("port_of_loading", "MISSING_REQUIRED"): ("DHL 原帳單沒有啟運港", {"N/A", "None", ""}),
        ("port_of_discharge", "MISSING_REQUIRED"): ("DHL 原帳單沒有目的港", {"N/A", "None", ""}),
        ("volume", "MISSING_REQUIRED"): ("DHL 原帳單沒有材積", {"N/A", "None", ""}),
        ("bl_no", "MISSING_REQUIRED"): ("DHL 原帳單沒有提單號碼", {"N/A", "None", ""}),
    },
    "RONG_A_CHAU": {
        ("invoice_date", "MISSING_REQUIRED"): ("RỒNG Á CHÂU 的 Excel 對帳表本身沒有發票日期", None),
        # 擷取規則只在費用金額不為 0 時才轉成一列，金額空白的只會是使用者規定
        # 「進口每筆出貨都要列」的 VAT / VAT PASS THROUGH 列 (稅額 0 時金額、稅額都 N/A)
        ("amount", "MISSING_REQUIRED"): ("進口每筆出貨都列 VAT 列，VAT PASS THROUGH 為 0 時金額與稅額皆為 N/A", None),
    },
    "INDOPROSTIME_TRANSPORT": {
        # 帳單印 "2026-JUN-31" (6 月沒有 31 日)；[POUCHEN]_IPE JUNI 的 FLIGHT 印
        # "CI761/287052025" (日期多一碼，得到 2025-MAY-87，使用者 2026-10-30 確認)
        ("arrive_date", "DATE_FORMAT"): ("INDOPROSTIME 原帳單的日期本身就不存在",
                                         {"2026-JUN-31", "2025-MAY-87"}),
        ("onboard_date", "DATE_FORMAT"): ("INDOPROSTIME 原帳單的日期本身就不存在",
                                          {"2026-JUN-31", "2025-MAY-87"}),
        # 原帳單本來就沒有 B/L NO. 值 (使用者 2026-11-01 確認，不計入)
        ("bl_no", "MISSING_REQUIRED"): ("INDOPROSTIME 原帳單本來就沒有 B/L NO. 值", None),
    },
}
# INDO PROSTIME EXPRESS (ID16) 與 TRANSPORT (ID07) 是兩個供應商 (使用者 2026-11-01 要求分開)。
# EXPRESS 先沿用 TRANSPORT 的原帳單問題清單當起點，之後依 EXPRESS 樣本各自調整，不再共用同一份。
KNOWN_SOURCE_ISSUES["INDOPROSTIME"] = dict(KNOWN_SOURCE_ISSUES["INDOPROSTIME_TRANSPORT"])
# 上面 B/L NO. 缺值的原帳單規則只限 TRANSPORT，EXPRESS 不沿用 (使用者 2026-11-01)
KNOWN_SOURCE_ISSUES["INDOPROSTIME"].pop(("bl_no", "MISSING_REQUIRED"), None)

# 供應商例外：這些 (欄位, 規則) 對該供應商不是問題，連附註都不列 (使用者規定)
SUPPLIER_EXEMPTIONS: Dict[str, set] = {
    # CONG THANH 是拖車行：CONT/XE 欄是車牌或簡寫櫃號、Số xe (車號) 可能只有
    # 數字、帳單上原本就是 0 (或空白視為 0) 的費用，都照抄原文；負數是對帳單
    # 層級的扣款 (例如貨損)，使用者確認正確
    "CONG_THANH": {("amount", "AMOUNT_ZERO"), ("amount", "AMOUNT_NEGATIVE")},
    # PSS Excel 帳單的 AMOUNT 為負數 = 折讓/扣款 (使用者 2026-10-31 確認：中 folder 的 -500,000)
    "PUTRA_SEJAHTERA_SENTOSA": {("amount", "AMOUNT_NEGATIVE")},
    # ECOMET 費用 DISCOUNT TERLAMBAT 為負數 = 折讓 (使用者確認，D26-03-0263 -500,000)
    "ECOMET": {("amount", "AMOUNT_NEGATIVE")},
}


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


def _check_header(row: Dict, supplier_key: Optional[str] = None) -> List[Dict]:
    issues = []
    for code in REQUIRED_HEADER_FIELDS:
        if _is_blank(row.get(code)):
            # 缺發票日期不算重大錯誤，只標 WARN (信心等級最多「中」)，不直接判定「低」。
            # 2026-11-12 使用者最高規則：所有供應商都適用，原帳單本身常漏填發票日期，
            # 不該整張發票被打成低信心。其餘必填欄位 (INVOICE NO/SUPPLIER/CONSIGNEE)
            # 不受影響，仍是 ERROR。
            severity = WARN if code == "invoice_date" else ERROR
            issues.append(_issue(severity, code, "MISSING_REQUIRED",
                                 f"必填欄位「{FIELD_LABELS[code]}」抓不到值"))
    applies = SUPPLIER_REQUIRED_WHEN.get(supplier_key or "", lambda _row: True)(row)
    for code, level in (SUPPLIER_REQUIRED_FIELDS.get(supplier_key or "", {}) if applies else {}).items():
        if code == "container_no" and not _FCL_VOLUME_RE.search(str(row.get("volume") or "")):
            continue    # 散貨 (LCL)、空運沒有櫃號
        if _is_blank(row.get(code)):
            issues.append(_issue(ERROR if level == "ERROR" else WARN, code, "MISSING_REQUIRED",
                                 f"必要欄位「{FIELD_LABELS[code]}」抓不到值"))
    for code in DATE_FIELD_CODES:
        raw = row.get(code)
        if _is_na(raw):
            continue
        if _coerce_to_date(raw) is None:
            issues.append(_issue(WARN, code, "DATE_FORMAT",
                                 f"「{FIELD_LABELS[code]}」日期格式無法辨識：{raw}"))
    return issues


# Excel 帳單的合計列公式錯誤 (呼叫端檢查後放進 extra_issues，見
# extract_utils.excel_formula_issues())
RULE_EXCEL_FORMULA = "EXCEL_FORMULA_RANGE"


def _explain_by_formula_issue(issues: List[Dict]) -> List[Dict]:
    """Excel 帳單的查驗順序：先查 Excel 公式，再查每票，最後查整份帳單合計。
    已經查到合計列公式錯誤時，整份帳單合計對不上就是這個錯誤造成的：合計
    不一致註明原因。使用者原則：沒通過金額稽核一律是重大 ERROR，原因是原
    帳單也一樣，所以不降級。"""
    if not any(i["rule"] == RULE_EXCEL_FORMULA and i["severity"] != INFO for i in issues):
        return issues
    explained = []
    for issue in issues:
        if issue["rule"] in ("TOTAL_NET_MISMATCH", "TOTAL_FINAL_MISMATCH", "TOTAL_VAT_MISMATCH",
                             "TOTAL_GRAND_MISMATCH"):
            issue = dict(issue,
                         message=f"{issue['message']}（原因：原帳單 Excel 合計列公式有誤，見公式錯誤）")
        explained.append(issue)
    return explained


# 只看每票勾稽的帳單：合計列公式錯誤只列成 INFO，不扣分 (使用者 2026-10-28：
# CONG THANH 代墊款清單「每票通過就 PASS」，例如 CHI HO 清單 E360 =D360*8% 把 8%
# 套在不課稅的費用上)。有整份合計勾稽的版面 (運費對帳單、POU CHEN 對帳表)，
# 公式錯誤仍是 ERROR (使用者 2026-10-28 確認)。
PER_TICKET_FORMULA_INFO_SUPPLIERS = {"CONG_THANH", "DWIHARTA"}


def _formula_issue_as_info(issues: List[Dict], supplier_key: Optional[str],
                           has_whole_audit: bool) -> List[Dict]:
    if supplier_key not in PER_TICKET_FORMULA_INFO_SUPPLIERS or has_whole_audit:
        return issues
    return [dict(i, severity=INFO, no_penalty=True,
                 message=f"{i['message']}（這家供應商只看每票勾稽，合計列公式錯誤屬原帳單問題）")
            if i["rule"] == RULE_EXCEL_FORMULA else i for i in issues]


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
            value = str(row.get(issue["field"], "")).strip()
            if only_values is None or (only_values.fullmatch(value) if isinstance(only_values, re.Pattern)
                                       else value in only_values):
                issue = dict(issue, severity=INFO, known_source_issue=True,
                             message=f"{issue['message']}（{note}，屬原帳單問題，非擷取錯誤）")
        marked.append(issue)
    return marked


# 單號欄寫的是文字 (沒有數字) 代表「額外費用項目」的供應商：不報單號格式，
# 改成逐筆標示人工查核 (費用名稱就是那段文字，使用者規定)
EXTRA_CHARGE_ID_SUPPLIERS = {"CONG_THANH"}
RULE_EXTRA_CHARGE = "EXTRA_CHARGE_REVIEW"


def _check_extra_charges(group_rows: List[Dict], audit_passed: bool = False) -> List[Dict]:
    """額外費用項目：金額勾稽沒通過 (或無法勾稽) 時 WARN 請人工查核；
    已通過金額勾稽就只當附註 (使用者規定)。"""
    issues = []
    for idx, row in enumerate(group_rows, start=1):
        texts = [str(row.get(code)) for code in ("order_no", "bl_no")
                 if not _is_blank(row.get(code)) and not re.search(r"\d", str(row.get(code)))]
        if not texts:
            continue
        message = (f"第 {idx} 筆費用「{row.get('description')}」是額外費用項目"
                   f" (單號欄寫的是費用名稱：{texts[0]})")
        if audit_passed:
            # 已通過金額勾稽：只是附註，也不扣信心分數 (使用者規定)
            issues.append(dict(_issue(INFO, "description", RULE_EXTRA_CHARGE,
                                      message + "，已通過金額勾稽", idx - 1), no_penalty=True))
        else:
            issues.append(_issue(WARN, "description", RULE_EXTRA_CHARGE,
                                 message + "，請人工查核", idx - 1))
    return issues


def _has_credit_text(row: Dict) -> bool:
    """這一筆費用名稱是否有 CREDIT 字樣 (折讓)。只看費用名稱，不看公司名稱。"""
    return "CREDIT" in str(row.get("description") or "").upper()


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
        if desc == UNREADABLE_DESC:
            # 原帳單無法判讀 (解析器放入標記，例如海關文件誤當帳單)：直接判 ERROR，
            # 不再逐欄檢查 (使用者 2026-10-31 確認：BDOIR00051508)
            issues.append(_issue(ERROR, "description", "UNREADABLE", "無法判讀", row))
            continue
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
            # 金額 0 是帳單原文 (例如 MATERAI 印花稅免收、ECOMET "CUSTOM Rp 0.00")：
            # 屬原帳單問題，不扣分 (使用者 2026-10-25 規定，所有供應商適用)
            issues.append(dict(_issue(INFO, "amount", "AMOUNT_ZERO",
                                      f"{label}金額為 0（帳單原文就是 0，屬原帳單問題，非擷取錯誤）", row),
                               known_source_issue=True))
        elif amount < 0 and not _has_credit_text(group_rows[idx - 1]):
            # 帳單有 CREDIT 字樣代表折讓，不計入 (使用者 2026-11-03)
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
            or entry.get("extract_final_total") or entry.get("sheet_totals")))
        if has_audit_rule:
            message = "帳單上找不到總額，無法做金額勾稽"
        else:
            message = "此供應商尚無總額稽核規則，無法做金額勾稽"
        return [_issue(INFO, "total", "NO_TOTAL_AUDIT", message)]

    issues = []
    checks = (
        ("match", "diff", "TOTAL_NET_MISMATCH", "費用加總 (未稅) 與帳單未稅總額不一致"),
        ("vat_match", "vat_diff", "TOTAL_VAT_MISMATCH", "VAT 加總與帳單 VAT 合計不一致"),
        ("grand_match", "grand_diff", "TOTAL_GRAND_MISMATCH", "未稅 + 稅額與帳單含稅總額不一致"),
        ("final_match", "final_diff", "TOTAL_FINAL_MISMATCH",
         "含稅 - 代扣稅與帳單實付金額 (Excel 帳單為表格合計列) 不一致"),
    )
    for match_key, diff_key, rule, message in checks:
        if audit.get(match_key) is False:
            diff = audit.get(diff_key)
            diff_text = f"{diff:,.2f}" if isinstance(diff, (int, float)) else diff
            # 原帳單問題備註 (例如發票重複列出)：仍是 ERROR，只附上原因
            note = f"（{audit['source_note']}）" if audit.get("source_note") else ""
            issues.append(_issue(ERROR, "total", rule, f"{message} (差額 {diff_text}){note}"))
    return issues


def _confidence(issues: List[Dict]) -> float:
    # 已知原帳單問題不是擷取錯誤，不扣分；no_penalty 的附註也不扣分
    score = 1.0 - sum(_SEVERITY_PENALTY[i["severity"]] for i in issues
                      if not (i.get("known_source_issue") or i.get("no_penalty")))
    return round(max(score, 0.0), 2)


def _level(issues: List[Dict], confidence: float) -> str:
    if any(i["severity"] == ERROR for i in issues):
        return LEVEL_LOW
    # 未勾稽代表金額沒有被查核過，不管其他問題為何都列為「低」
    # (使用者 2026-11-12 規定，所有供應商適用)
    if any(i["rule"] == "NO_TOTAL_AUDIT" for i in issues):
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
    # 整份帳單合計的勾稽 (invoice_no 為 None)
    has_whole_audit = any(ar.get("invoice_no") is None for ar in audit_results or [])

    base_issues = list(extra_issues or [])
    if not supplier_key:
        base_issues.append(_issue(ERROR, "supplier_key", "UNKNOWN_SUPPLIER",
                                  "無法辨識供應商，沒有對應的擷取規則"))

    results = []
    groups = _group_rows_by_invoice(rows) if rows else {"": [{}]}
    for inv_key, group_rows in groups.items():
        issues = list(base_issues)
        issues.extend(_check_header(group_rows[0], supplier_key))
        issues.extend(_check_items(group_rows))
        audit_issues = _check_audit(audit_by_invoice.get(inv_key), supplier_key)
        if supplier_key in EXTRA_CHARGE_ID_SUPPLIERS:
            issues.extend(_check_extra_charges(group_rows, audit_passed=(
                inv_key in audit_by_invoice and not audit_issues)))
        issues.extend(audit_issues)
        exempt = SUPPLIER_EXEMPTIONS.get(supplier_key or "", set())
        issues = [i for i in issues if (i["field"], i["rule"]) not in exempt]
        issues = _explain_by_formula_issue(issues)
        issues = _formula_issue_as_info(issues, supplier_key, has_whole_audit)
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
