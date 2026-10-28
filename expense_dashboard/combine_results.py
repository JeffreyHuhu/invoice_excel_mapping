# -*- coding: utf-8 -*-
"""
combine_results.py — 把所有轉檔結果 Excel 整合成一份 (欄位跟轉檔結果模板一樣)

【做什麼】
  1. 掃描資料夾底下「XXX-辨識等級(高)」「XXX-辨識等級(中)」裡的轉檔結果 Excel
     (沿用 process_data.scan_excel_files，規則一樣：跳過轉檔總表、原始 Excel 帳單)。
  2. 依標題列找出 18 個標準欄位，順序與抬頭跟轉檔結果模板一樣；最右邊加一欄
     「辨識等級」(高/中)，讓主管分辨可能有錯的資料 (使用者 2026-10-28 指定)。
  3. 同一張發票 (供應商 + 發票號碼) 出現在多份 Excel 時，只保留第一份
     (等級高優先，再依路徑排序)，跟資料庫的重複判斷一樣 (含逐列比對的供應商)。
  4. 寫成一個工作表「工作表」，日期欄位用 mmm-dd-yyyy 顯示。
     讀不了的檔案記到 log，不中斷整批處理。
  5. 有給 totals_fx_folder (個別供應商的「<KEY>轉檔結果整合_資料模板.xlsx」) 時加第二個
     工作表「費用總和」：依幣別列出費用 (AMOUNT)、VAT、TOTAL (= 費用 + VAT) 的
     總和，當地幣別與美金各一組 (使用者 2026-10-28 指定)。匯率用作帳匯率檔
     (fx.latest_book_file / read_book_rates，跟資料庫一樣，每幣別最新一期)。

【怎麼用】
  python combine_results.py <輸出.xlsx> [<資料夾1> <資料夾2> ...] [--levels 高 中]
"""

import argparse
import sys
from datetime import date, datetime
from pathlib import Path
from typing import List, Sequence

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font

from expense_normalize import detect_currency, folder_currency, is_no_own_invoice, parse_amount
from process_data import _find_header, _read_rows, get_logger, scan_excel_files

# 抬頭跟 extract_utils_streamlit.DISPLAY_HEADERS 一致 (那支檔案要載入 PDF 套件，這裡不 import)
TEMPLATE_HEADERS = {
    "invoice_date": "INVOICE DATE\n發票日期",
    "invoice_no": "INVOICE NO\n發票號碼",
    "supplier": "SUPPLIER\n供應商名稱",
    "consignee": "CONSIGNEE\n集團公司名稱",
    "description": "Description\n費用名稱",
    "amount": "AMOUNT\n金額",
    "vat": "VAT\n稅額",
    "order_no": "Order No.",
    "arrive_date": "Arrive Date\n到達日",
    "onboard_date": "on board date\n開船日",
    "bl_no": "B/L NO.\n提單號碼",
    "mbl_no": "MBL NO.\n主提單號碼",
    "port_of_loading": "Port of Loading\n啟運港",
    "port_of_discharge": "Port of discharge\n目的港",
    "volume": "Volume\n材積",
    "vessel": "Vessel\n船名",
    "container_no": "Container no.\n貨櫃號碼",
    "shipper": "Shipper\n出貨方",
}
FIELDS = list(TEMPLATE_HEADERS)
LEVEL_HEADER = "辨識等級"
DATE_FIELDS = {"invoice_date", "arrive_date", "onboard_date"}
DATE_FORMAT = "mmm-dd-yyyy"
SHEET_NAME = "工作表"
TOTALS_SHEET_NAME = "費用總和"
EXCEL_MAX_ROWS = 1_048_575


def _invoice_key(row: dict):
    """(供應商, 發票號碼)；沒有發票號碼就不做重複判斷。SUPPLIER_NO_OWN_INVOICE
    的供應商要整列 (發票號碼+費用名稱+金額+提單號碼) 相同才算重複 (跟資料庫一樣)。"""
    no = row.get("invoice_no")
    if no in (None, ""):
        return None
    key = (str(row.get("supplier") or "").strip().upper(), str(no).strip().upper())
    if is_no_own_invoice(row.get("supplier")):
        key += tuple(str(row.get(f)) for f in ("description", "amount", "bl_no"))
    return key


def collect_rows(roots: Sequence[Path], levels: Sequence[str]):
    """回傳 (明細列, 統計)。明細列是 {欄位代碼: 值}。"""
    log = get_logger()
    files = []
    for rank, level in enumerate(levels):
        files += [(rank, level, path) for path, _, _, _ in scan_excel_files(list(roots), level, None)]
    files.sort(key=lambda t: (t[0], t[2]))

    rows: List[dict] = []
    seen = set()
    stats = {lv: {"files": 0, "rows": 0} for lv in levels}
    stats.update(failed=0, dup_rows=0)
    for _, level, path in files:
        try:
            raw = _read_rows(path)
            header_idx, mapping = _find_header(raw)
        except Exception as e:  # 損壞或格式不符：記 log，繼續下一份
            log.warning("整合 Excel 讀取失敗 %s：%s", path, e)
            stats["failed"] += 1
            continue
        file_rows = []
        for values in raw[header_idx + 1:]:
            row = {code: values[col] for col, code in mapping.items() if col < len(values)}
            row = {k: (None if v == "" else v) for k, v in row.items()}
            if any(v is not None for v in row.values()):
                row["level"] = level
                row["_path"] = path          # 判斷幣別用 (國家資料夾)，不寫進 Excel
                file_rows.append(row)
        keys = {_invoice_key(r) for r in file_rows} - {None}
        dup_keys = keys & seen
        if dup_keys:
            before = len(file_rows)
            file_rows = [r for r in file_rows if _invoice_key(r) not in dup_keys]
            stats["dup_rows"] += before - len(file_rows)
        seen |= keys
        rows += file_rows
        stats[level]["files"] += 1
        stats[level]["rows"] += len(file_rows)
    return rows, stats


def compute_totals(rows: List[dict], fx_folder: Path):
    """依幣別加總 -> ([(幣別, 列數, 費用, VAT, TOTAL, 1 USD 兌換數或 None)], 匯率檔名)。
    幣別判斷跟資料庫一樣：帳單幣別 > 供應商 > 國家資料夾。"""
    from collections import OrderedDict
    import fx
    sums = OrderedDict()
    for row in rows:
        cur = detect_currency(row.get("amount"), row.get("supplier"),
                              folder_currency(row.get("_path"))) or "(未知)"
        amount = parse_amount(row.get("amount")) or 0.0
        vat = parse_amount(row.get("vat")) or 0.0
        s = sums.setdefault(cur, [0, 0.0, 0.0])
        s[0] += 1
        s[1] += amount
        s[2] += vat
    book = fx.latest_book_file(fx_folder)
    rates = fx.read_book_rates(book) if book else {}
    out = []
    for cur, (n, amount, vat) in sums.items():
        rate = 1.0 if cur == "USD" else (rates[cur][2] if cur in rates else None)
        out.append((cur, n, amount, vat, amount + vat, rate))
    return out, (book.name if book else None)


def _write_totals_sheet(wb: Workbook, totals, book_name) -> None:
    ws = wb.create_sheet(TOTALS_SHEET_NAME)
    bold = Font(bold=True)
    money = "#,##0.00"
    for col, width in zip("ABCDEFGHIJ", (12, 10, 20, 20, 20, 14, 18, 18, 18)):
        ws.column_dimensions[col].width = width
    title = WriteOnlyCell(ws, value="費用總和 (當地幣別 & 美金)")
    title.font = Font(bold=True, size=14)
    ws.append([title])
    ws.append([f"匯率：{book_name or '找不到作帳匯率檔'} (每幣別最新一期；USD = 當地幣別 ÷ 1 USD 兌換數)"])
    ws.append(["TOTAL = 費用 + VAT；收「辨識等級」高與中的全部明細 (重複發票已略過)"])
    ws.append([])
    header = ["幣別", "明細列數", "費用總和", "VAT 總和", "TOTAL 總和", "1 USD 兌換數",
              "費用總和 (USD)", "VAT 總和 (USD)", "TOTAL 總和 (USD)"]
    cells = []
    for text in header:
        c = WriteOnlyCell(ws, value=text)
        c.font = bold
        cells.append(c)
    ws.append(cells)
    usd_sum = [0.0, 0.0, 0.0]
    missing = []

    def num(v, fmt=money):
        c = WriteOnlyCell(ws, value=v)
        c.number_format = fmt
        return c

    for cur, n, amount, vat, total, rate in totals:
        usd = [v / rate for v in (amount, vat, total)] if rate else [None, None, None]
        if rate:
            usd_sum = [a + b for a, b in zip(usd_sum, usd)]
        else:
            missing.append(cur)
        ws.append([cur, n, num(amount), num(vat), num(total), num(rate, "#,##0.0000") if rate else "沒有匯率",
                   *(num(v) for v in usd)])
    label = WriteOnlyCell(ws, value="合計 (USD)")
    label.font = bold
    ws.append([label, sum(t[1] for t in totals), None, None, None, None, *(num(v) for v in usd_sum)])
    if missing:
        ws.append([])
        ws.append([f"注意：{'、'.join(missing)} 沒有匯率，沒有算進 USD 合計"])


def write_workbook(rows: List[dict], out_path: Path, totals=None) -> None:
    """write_only 模式寫檔，大量資料也不佔太多記憶體。"""
    if len(rows) > EXCEL_MAX_ROWS:
        raise SystemExit(f"明細 {len(rows):,} 列，超過 Excel 上限 {EXCEL_MAX_ROWS:,} 列。")
    wb = Workbook(write_only=True)
    ws = wb.create_sheet(SHEET_NAME)
    ws.freeze_panes = "A2"
    bold, wrap = Font(bold=True), Alignment(wrap_text=True, vertical="center")
    header = []
    for text in [*TEMPLATE_HEADERS.values(), LEVEL_HEADER]:
        cell = WriteOnlyCell(ws, value=text)
        cell.font, cell.alignment = bold, wrap
        header.append(cell)
    ws.append(header)
    for row in rows:
        out = []
        for field in FIELDS:
            value = row.get(field)
            if field in DATE_FIELDS and isinstance(value, (date, datetime)):
                cell = WriteOnlyCell(ws, value=value)
                cell.number_format = DATE_FORMAT
                out.append(cell)
            else:
                out.append(value)
        out.append(row["level"])
        ws.append(out)
    if totals is not None:
        _write_totals_sheet(wb, *totals)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def run(out_path: Path, roots: Sequence[Path], levels: Sequence[str] = ("高", "中"),
        totals_fx_folder: Path = None) -> Path:
    """totals_fx_folder：放作帳匯率檔的資料夾；有給就加「費用總和」工作表。"""
    rows, stats = collect_rows(roots, levels)
    summary = compute_totals(rows, totals_fx_folder) if totals_fx_folder else None
    write_workbook(rows, out_path, summary)
    print(f"已整合轉檔結果 Excel：{out_path}")
    for lv in levels:
        print(f"  等級({lv})：{stats[lv]['files']:,} 份、{stats[lv]['rows']:,} 列")
    print(f"  共 {len(rows):,} 列；重複發票略過 {stats['dup_rows']:,} 列；讀取失敗 {stats['failed']:,} 份")
    for cur, n, amount, vat, total, rate in (summary[0] if summary else []):
        usd = f"{total / rate:,.2f} USD" if rate else "沒有匯率"
        print(f"  {cur}：費用 {amount:,.2f}、VAT {vat:,.2f}、TOTAL {total:,.2f} ({usd})")
    return out_path


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("out", type=Path)
    parser.add_argument("roots", nargs="+", type=Path)
    parser.add_argument("--levels", nargs="+", default=["高", "中"])
    args = parser.parse_args()
    run(args.out, args.roots, args.levels)
    return 0


if __name__ == "__main__":
    sys.exit(main())
