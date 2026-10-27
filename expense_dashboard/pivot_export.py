# -*- coding: utf-8 -*-
"""
pivot_export.py — 匯出「可樞紐分析」的費用彙總 Excel

【做什麼】
  把資料庫明細填進樞紐範本 templates/pivot_template.xlsx (由
  build_pivot_template.py 用 Excel 做一次)。範本已有：
    兩頁儀表板 (KPI、交叉分析篩選器、樞紐分析圖)、各角度樞紐分析表、
    明細表格「明細表」、讀取失敗表格「讀取失敗表」。版面見 build_pivot_template.py。
  樞紐分析表設定「開檔時自動重新整理」，主管打開檔案就看到最新資料。

【欄位】(使用者 2026-10-27 指定)
  前 18 欄跟轉檔結果模板一致 (extract_utils_streamlit.DISPLAY_HEADERS 的
  順序與名稱，雙行改成單行「英文 中文」)。分析欄位接在右邊。

【為什麼直接改 xlsx 的 XML】
  openpyxl 讀寫既有檔案時會丟掉圖表與交叉分析篩選器；用 Excel 自動化又
  要求執行的電腦有 Excel。所以只換掉範本裡兩張表格的資料列 (sheetData)
  與表格範圍，其他零件原封不動複製。10 萬列約數十秒，不需要 Excel。
"""

import io
import re
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union
from xml.sax.saxutils import escape

import pandas as pd

import fx
from expense_normalize import (CONTAINER_TYPES, COUNTRY_FOLDER_CURRENCY, parse_volume,
                               supplier_country)

TEMPLATE = Path(__file__).with_name("templates") / "pivot_template.xlsx"
DETAIL_SHEET, DETAIL_TABLE = "明細", "明細表"
FAILED_SHEET, FAILED_TABLE = "讀取失敗", "讀取失敗表"
# Excel 一張工作表最多 1,048,576 列 (含標題列)
EXCEL_MAX_ROWS = 1_048_575

# (資料欄位, Excel 標題, 型別)。型別：date / num / rate / int / text
TEMPLATE_COLUMNS: List[Tuple[str, str, str]] = [
    ("invoice_date", "INVOICE DATE 發票日期", "date"),
    ("invoice_no", "INVOICE NO 發票號碼", "text"),
    ("supplier", "SUPPLIER 供應商名稱", "text"),
    ("consignee", "CONSIGNEE 集團公司名稱", "text"),
    ("description_raw", "Description 費用名稱", "text"),
    ("amount_orig", "AMOUNT 金額", "num"),
    ("vat_orig", "VAT 稅額", "num"),
    ("order_no", "Order No.", "text"),
    ("arrive_date", "Arrive Date 到達日", "date"),
    ("onboard_date", "on board date 開船日", "date"),
    ("bl_no", "B/L NO. 提單號碼", "text"),
    ("mbl_no", "MBL NO. 主提單號碼", "text"),
    ("port_of_loading", "Port of Loading 啟運港", "text"),
    ("port_of_discharge", "Port of discharge 目的港", "text"),
    ("volume", "Volume 材積", "text"),
    ("vessel", "Vessel 船名", "text"),
    ("container_no", "Container no. 貨櫃號碼", "text"),
    ("shipper", "Shipper 出貨方", "text"),
]
ANALYSIS_COLUMNS: List[Tuple[str, str, str]] = [
    ("currency", "原幣別", "text"),
    ("fx_rate", "匯率 (1 USD = ?)", "rate"),
    ("amount", "AMOUNT (USD)", "num"),
    ("vat", "VAT (USD)", "num"),
    ("expense_category", "費用大類", "text"),
    ("expense_item", "費用項目 (標準化)", "text"),
    ("country", "國家", "text"),
    ("year", "年", "int"),
    ("quarter", "季", "text"),
    ("month", "月", "int"),
    ("year_month", "年月", "text"),
    ("route", "航線", "text"),
    # 運量 (parse_volume 拆材積)：同一供應商的同一批貨只在第一列有值，樞紐加總不重複
    ("ship_mode", "運送型態", "text"),
] + [(f"cnt_{t}", f"{t} 櫃數", "count") for t in CONTAINER_TYPES] + [
    ("teu", "TEU", "count"),
    ("cbm", "CBM", "qty"),
    ("kg", "重量 KG", "qty"),
    # 單位成本的分子：FCL 費用只除以 TEU、LCL 費用只除以 CBM
    ("fcl_usd", "FCL 費用 (USD)", "num"),
    ("lcl_usd", "LCL 費用 (USD)", "num"),
    ("invoice_count", "發票張數", "count"),
    ("source_folder", "來源資料夾", "text"),
    ("source_file", "來源檔案", "text"),
]
COLUMNS = TEMPLATE_COLUMNS + ANALYSIS_COLUMNS
FAILED_COLUMNS: List[Tuple[str, str, str]] = [
    ("source_file", "檔案", "text"), ("error", "原因", "text")]
# add_helper_columns() 算的欄位；其餘從資料庫讀
HELPER_COLUMNS = {"country", "quarter", "year_month", "route", "ship_mode", "teu", "cbm", "kg",
                  "fcl_usd", "lcl_usd", "invoice_count"} | {f"cnt_{t}" for t in CONTAINER_TYPES}
DB_COLUMNS = [c for c, _, _ in COLUMNS if c not in HELPER_COLUMNS]

NUMBER_FORMATS = {"date": "mmm-dd-yyyy", "num": "#,##0.00", "rate": "#,##0.0000", "int": "0",
                  "count": "#,##0", "qty": "#,##0.000"}

_COUNTRY_NAMES = {
    "CN": "中國", "TW": "台灣", "IN": "印度", "ID": "印尼", "VN": "越南", "TH": "泰國",
    "MY": "馬來西亞", "SG": "新加坡", "PH": "菲律賓", "KH": "柬埔寨", "LA": "寮國",
    "MM": "緬甸", "BN": "汶萊", "TL": "東帝汶",
}
_FOLDER_ALIAS = {"中国": "中國", "大陸": "中國", "臺灣": "台灣"}
# XML 不允許的控制字元
_ILLEGAL_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def items_cte(fixed: Optional[Dict[str, float]] = None) -> Tuple[str, list]:
    """換算 USD 的 items CTE (fx.rate_sql)，另外保留原幣 VAT (vat_orig)。"""
    sql, params = fx.rate_sql(fixed)
    return "WITH " + sql.replace("amount AS amount_orig,",
                                 "amount AS amount_orig, vat AS vat_orig,"), params


def detail_select(where: str = "WHERE NOT is_duplicate") -> str:
    """查明細的 SELECT (接在 items_cte() 後面)。"""
    return (f"SELECT {', '.join(DB_COLUMNS)} FROM items {where} "
            "ORDER BY supplier, invoice_date, invoice_no")


def _country(source_file, supplier) -> Optional[str]:
    """國家：來源路徑裡的國家資料夾優先，其次看供應商名稱的國別特徵。"""
    for part in reversed(re.split(r"[\\/]", str(source_file or ""))):
        name = part.strip()
        if name in COUNTRY_FOLDER_CURRENCY:
            return _FOLDER_ALIAS.get(name, name)
    return _COUNTRY_NAMES.get(supplier_country(supplier))


def add_helper_columns(df: pd.DataFrame) -> pd.DataFrame:
    """加上國家、季、年月、航線、運量、發票張數 (樞紐分析分組與加總用)。
    df 要依 detail_select() 的順序排好 (每批貨的運量記在最早那一列)。"""
    pairs = df[["source_file", "supplier"]].drop_duplicates()
    lookup = {(f, s): _country(f, s) for f, s in pairs.itertuples(index=False)}
    df["country"] = [lookup[k] for k in zip(df["source_file"], df["supplier"])]
    year = pd.to_numeric(df["year"], errors="coerce").astype("Int64")
    month = pd.to_numeric(df["month"], errors="coerce").astype("Int64")
    df["quarter"] = ("Q" + ((month - 1) // 3 + 1).astype("string")).where(month.notna())
    df["year_month"] = (year.astype("string") + "-"
                        + month.astype("string").str.zfill(2)).where(year.notna() & month.notna())
    pol = df["port_of_loading"].fillna("").astype(str).str.strip()
    pod = df["port_of_discharge"].fillna("").astype(str).str.strip()
    df["route"] = (pol.where(pol != "", "?") + " → " + pod.where(pod != "", "?")).where(
        (pol != "") | (pod != ""))

    # 運量：材積原文 -> 數字。同一供應商 + 提單號碼 (沒有就用發票號碼) + 材積 = 同一批貨，
    # 只記在第一列 (使用者 2026-10-27：每家供應商各算一次)
    parsed = {v: parse_volume(v) for v in df["volume"].dropna().unique()}
    mode = df["volume"].map(pd.Series({v: r["mode"] for v, r in parsed.items()}, dtype=object))
    df["ship_mode"] = mode
    first = ~pd.DataFrame({"s": df["supplier"], "b": df["bl_no"].fillna(df["invoice_no"]),
                           "v": df["volume"]}).duplicated()
    for key, col in [(t, f"cnt_{t}") for t in CONTAINER_TYPES] + [
            ("teu", "teu"), ("cbm", "cbm"), ("kg", "kg")]:
        values = pd.Series({v: r[key] for v, r in parsed.items()}, dtype=float)
        df[col] = df["volume"].map(values).where(first)
    amount = pd.to_numeric(df["amount"], errors="coerce")
    df["fcl_usd"] = amount.where(mode == "FCL")
    df["lcl_usd"] = amount.where(mode == "LCL")
    # 發票張數：每張發票 (供應商 + 發票號碼) 只在第一列記 1
    df["invoice_count"] = (~pd.DataFrame({"s": df["supplier"], "i": df["invoice_no"].fillna(
        df["source_file"])}).duplicated()).astype(int)
    return df


# ---------------------------------------------------------------------------
# xlsx (zip) 處理
# ---------------------------------------------------------------------------

def _col_letter(i: int) -> str:
    """0 起算的欄號 -> Excel 欄名 (A, B, ..., AA)。"""
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _resolve(base: str, target: str) -> str:
    """rels 的相對路徑 -> zip 內的完整路徑。"""
    if target.startswith("/"):
        return target.lstrip("/")
    parts = base.split("/")[:-1]
    for p in target.split("/"):
        if p == "..":
            parts.pop()
        elif p != ".":
            parts.append(p)
    return "/".join(parts)


def _rels(z: zipfile.ZipFile, part: str) -> Dict[str, str]:
    """某個零件的關聯：{rId: 目標完整路徑}。"""
    folder, name = part.rsplit("/", 1)
    rels_path = f"{folder}/_rels/{name}.rels"
    if rels_path not in z.namelist():
        return {}
    xml = z.read(rels_path).decode("utf-8")
    return {m.group(1): _resolve(part, m.group(2))
            for m in re.finditer(r'<Relationship [^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"', xml)}


def _sheet_part(z: zipfile.ZipFile, sheet_name: str) -> str:
    """工作表名稱 -> zip 內的 sheetN.xml 路徑。"""
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = _rels(z, "xl/workbook.xml")
    for m in re.finditer(r"<sheet [^>]*?/>", wb):
        tag = m.group(0)
        name = re.search(r'name="([^"]+)"', tag).group(1)
        if name == escape(sheet_name):
            return rels[re.search(r'r:id="([^"]+)"', tag).group(1)]
    raise KeyError(f"範本沒有工作表：{sheet_name}")


def _cell_fragments(series: pd.Series, kind: str, style: str) -> List[Optional[str]]:
    """一欄資料 -> 每格的 XML 片段 (不含 <c r="..."，空值為 None)。"""
    s_attr = f' s="{style}"' if style else ""
    if kind == "date":
        days = (pd.to_datetime(series, errors="coerce") - pd.Timestamp("1899-12-30")) \
            / pd.Timedelta(days=1)
        return [None if pd.isna(v) else f"{s_attr}><v>{v:.10g}</v></c>" for v in days]
    if kind in NUMBER_FORMATS:   # date 以外有數字格式的型別都是數字
        nums = pd.to_numeric(series, errors="coerce")
        return [None if pd.isna(v) else f"{s_attr}><v>{float(v):.15g}</v></c>" for v in nums]
    out: List[Optional[str]] = []
    for v in series:
        if v is None or (not isinstance(v, str) and pd.isna(v)):
            out.append(None)
            continue
        text = _ILLEGAL_XML.sub("", str(v))
        out.append(None if text == "" else
                   f' t="inlineStr"{s_attr}><is><t xml:space="preserve">'
                   f"{escape(text)}</t></is></c>")
    return out


def _shared_strings(z: zipfile.ZipFile) -> List[str]:
    """共用字串表 (範本的標題列存在這裡)。"""
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    xml = z.read("xl/sharedStrings.xml").decode("utf-8")
    return ["".join(re.findall(r"<t[^>]*>([^<]*)</t>", si))
            for si in re.findall(r"<si>(.*?)</si>", xml, re.S)]


def _write_sheet(zout: zipfile.ZipFile, name: str, template_xml: str, shared: List[str],
                 df: pd.DataFrame, columns: List[Tuple[str, str, str]]) -> str:
    """把 df 寫成工作表 XML (沿用範本的標題列與第 2 列的儲存格格式)。回傳新範圍。"""
    head, rest = template_xml.split("<sheetData>", 1)
    rows_xml, tail = rest.split("</sheetData>", 1)
    header = re.search(r'<row r="1"[ >].*?</row>', rows_xml, re.S).group(0)
    titles = ([shared[int(i)] for i in re.findall(r'<c [^>]*?t="s"[^>]*><v>(\d+)</v>', header)]
              or re.findall(r"<t[^>]*>([^<]*)</t>", header))
    expected = [escape(t) for _, t, _ in columns]
    if titles != expected:
        raise ValueError(f"範本標題跟程式欄位不一致，請重跑 build_pivot_template.py：{titles}")
    row2 = re.search(r'<row r="2"[ >].*?</row>', rows_xml, re.S)
    styles = dict(re.findall(r'<c r="([A-Z]+)2"[^>]*?s="(\d+)"', row2.group(0))) if row2 else {}
    letters = [_col_letter(i) for i in range(len(columns))]
    frags = [_cell_fragments(df[c], kind, styles.get(letters[i], ""))
             for i, (c, _, kind) in enumerate(columns)]
    last_row = max(len(df) + 1, 2)  # 表格至少要有一列資料列
    ref = f"A1:{letters[-1]}{last_row}"
    head = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="{ref}"/>', head)

    with zout.open(name, "w") as f:
        f.write((head + "<sheetData>" + header).encode("utf-8"))
        buf: List[str] = []
        for n, row in enumerate(zip(*frags), start=2):
            cells = "".join(f'<c r="{letters[i]}{n}"{frag}'
                            for i, frag in enumerate(row) if frag is not None)
            buf.append(f'<row r="{n}">{cells}</row>')
            if len(buf) >= 5000:
                f.write("".join(buf).encode("utf-8"))
                buf.clear()
        if not len(df):
            buf.append('<row r="2"/>')
        f.write(("".join(buf) + "</sheetData>" + tail).encode("utf-8"))
    return ref


def write_workbook(detail: pd.DataFrame, failed: Optional[pd.DataFrame],
                   out: Union[Path, io.BytesIO]) -> None:
    """把明細 (DB_COLUMNS，未加輔助欄也可) 與讀取失敗清單填進樞紐範本，存到 out。"""
    if len(detail) > EXCEL_MAX_ROWS:
        raise ValueError(f"明細 {len(detail):,} 列，超過 Excel 上限 {EXCEL_MAX_ROWS:,} 列")
    if "route" not in detail.columns:
        detail = add_helper_columns(detail.copy())
    if failed is None:
        failed = pd.DataFrame(columns=[c for c, _, _ in FAILED_COLUMNS])
    with zipfile.ZipFile(TEMPLATE) as zin:
        sheets = {_sheet_part(zin, DETAIL_SHEET): (detail, COLUMNS),
                  _sheet_part(zin, FAILED_SHEET): (failed, FAILED_COLUMNS)}
        tables = {t: s for s in sheets for t in _rels(zin, s).values() if "/tables/" in t}
        shared = _shared_strings(zin)
        refs: Dict[str, str] = {}
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            # 先寫工作表，才知道表格的新範圍
            for name, (df, cols) in sheets.items():
                refs[name] = _write_sheet(zout, name, zin.read(name).decode("utf-8"),
                                          shared, df, cols)
            for item in zin.infolist():
                name = item.filename
                if name in sheets:
                    continue
                data = zin.read(name)
                if name in tables:
                    xml = data.decode("utf-8")
                    ref = refs[tables[name]]
                    xml = re.sub(r'(<table [^>]*?\bref=")[^"]*"', rf'\g<1>{ref}"', xml, count=1)
                    xml = re.sub(r'(<autoFilter [^>]*?\bref=")[^"]*"', rf'\g<1>{ref}"', xml, count=1)
                    data = xml.encode("utf-8")
                elif re.match(r"xl/pivotCache/pivotCacheDefinition\d+\.xml$", name):
                    # 開檔時自動重新整理 (範本已設定，這裡再確保一次)
                    xml = data.decode("utf-8")
                    if "refreshOnLoad=" not in xml:
                        xml = xml.replace("<pivotCacheDefinition ",
                                          '<pivotCacheDefinition refreshOnLoad="1" ', 1)
                    data = xml.encode("utf-8")
                zout.writestr(item, data)
