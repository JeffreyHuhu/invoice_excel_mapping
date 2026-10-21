# -*- coding: utf-8 -*-
"""
process_data.py — 第一階段：把上萬份「辨識等級(高)」轉檔結果 Excel 彙整進 DuckDB

【做什麼】
  1. 掃描指定資料夾 (可多個) 底下所有「XXX-辨識等級(高)」資料夾裡的 Excel
     (轉檔工具的輸出)，略過轉檔總表與 Excel 暫存檔 (~$)。
  2. 多個 process 平行讀檔 (python-calamine，沒有的話退回 openpyxl 唯讀
     模式)，每份檔案讀完只回傳清洗後的明細列；主程式每累積 BATCH_ROWS 列
     就寫進資料庫一次，記憶體用量固定，不會因為檔案多而爆掉。
  3. 清洗 (expense_normalize.py)：發票日期 -> 標準日期 + 年/月；AMOUNT/VAT
     -> 數值；Description -> 標準費用項目 + 費用大類；港口統一大小寫。
  4. 寫進 DuckDB (預設 data/expenses.duckdb)：
       expense_items  清洗後的完整明細 (原始費用名稱也保留在 description_raw)
       etl_files      每份 Excel 的處理紀錄 (修改時間、大小、列數、錯誤原因)
  5. 增量更新：再跑一次只處理新增或有修改的檔案；已經不存在的檔案，它的
     明細會從資料庫移除。加 --full 則整個重建。
  6. 同一張發票 (供應商 + 發票號碼) 出現在好幾份 Excel 時 (例如同一份 PDF
     轉了兩次)，只保留其中一份，其餘標記 is_duplicate，儀表板預設不計入。
  7. 讀不了/格式不符的 Excel 記進 etl_files 與 logs/expense_etl.log，
     不中斷整批處理。

【怎麼用】
  python process_data.py "D:\\發票\\02-Maersk Logistics Indonesia"
  python process_data.py <資料夾1> <資料夾2> --supplier-folder MAERSK
  python process_data.py <資料夾> --full --parquet data\\expenses.parquet
"""

import argparse
import logging
import os
import sys
import time
from concurrent.futures import ProcessPoolExecutor
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Dict, Iterator, List, Optional, Tuple

import duckdb
import pandas as pd

from expense_normalize import (
    clean_text,
    normalize_expense,
    normalize_port,
    parse_amount,
    parse_date,
    parse_float,
)

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_DB = BASE_DIR / "data" / "expenses.duckdb"
LOG_PATH = BASE_DIR / "logs" / "expense_etl.log"

# 每累積幾列寫進資料庫一次 (控制主程式記憶體用量)
BATCH_ROWS = 20_000
# 標題列最多往下找幾列 (pdf_to_excel 範本在第 1 列、批次範本在第 2 列)
HEADER_SEARCH_ROWS = 6
SUMMARY_REPORT_PREFIX = "轉檔總表"

# Excel 標題第一行 (換行前、轉大寫、去掉結尾的點) -> 欄位代碼
HEADER_FIELDS = {
    "INVOICE DATE": "invoice_date", "INVOICE NO": "invoice_no",
    "SUPPLIER": "supplier", "CONSIGNEE": "consignee", "DESCRIPTION": "description",
    "AMOUNT": "amount", "VAT": "vat", "ORDER NO": "order_no",
    "ARRIVE DATE": "arrive_date", "ON BOARD DATE": "onboard_date",
    "B/L NO": "bl_no", "MBL NO": "mbl_no", "PORT OF LOADING": "port_of_loading",
    "PORT OF DISCHARGE": "port_of_discharge", "VOLUME": "volume",
    "VESSEL": "vessel", "CONTAINER NO": "container_no", "SHIPPER": "shipper",
}

# expense_items 的欄位 (順序 = 建表順序 = 寫入 DataFrame 的欄位順序)
ITEM_COLUMNS = (
    ("source_file", "VARCHAR"), ("source_folder", "VARCHAR"),
    ("invoice_date", "DATE"), ("year", "INTEGER"), ("month", "INTEGER"),
    ("invoice_no", "VARCHAR"), ("supplier", "VARCHAR"), ("consignee", "VARCHAR"),
    ("description_raw", "VARCHAR"), ("expense_item", "VARCHAR"),
    ("expense_category", "VARCHAR"), ("amount", "DOUBLE"), ("vat", "DOUBLE"),
    ("order_no", "VARCHAR"), ("arrive_date", "DATE"), ("onboard_date", "DATE"),
    ("bl_no", "VARCHAR"), ("mbl_no", "VARCHAR"), ("port_of_loading", "VARCHAR"),
    ("port_of_discharge", "VARCHAR"), ("volume", "DOUBLE"), ("vessel", "VARCHAR"),
    ("container_no", "VARCHAR"), ("shipper", "VARCHAR"),
)
_DATE_COLUMNS = {name for name, typ in ITEM_COLUMNS if typ == "DATE"}

try:
    from python_calamine import CalamineWorkbook
except ImportError:  # 沒裝 calamine 時退回 openpyxl (較慢但結果相同)
    CalamineWorkbook = None


def get_logger() -> logging.Logger:
    logger = logging.getLogger("expense_etl")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    try:
        LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        handler: logging.Handler = RotatingFileHandler(
            LOG_PATH, maxBytes=1_000_000, backupCount=5, encoding="utf-8")
    except OSError:
        handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(handler)
    return logger


# ============================================================
# 掃描檔案
# ============================================================

def _is_target_folder(name: str, level: str, supplier_folder: Optional[str]) -> bool:
    if not name.endswith(f"-辨識等級({level})"):
        return False
    return supplier_folder is None or name.upper().startswith(supplier_folder.upper())


def scan_excel_files(roots: List[Path], level: str,
                     supplier_folder: Optional[str]) -> Iterator[Tuple[str, str, float, int]]:
    """逐一產生 (檔案路徑, 來源資料夾名稱, 修改時間, 大小)。用 os.scandir
    走訪，不一次把整個資料夾樹讀進記憶體。"""
    stack = [str(r) for r in roots]
    while stack:
        current = stack.pop()
        try:
            entries = list(os.scandir(current))
        except OSError as e:
            get_logger().warning("無法讀取資料夾 %s：%s", current, e)
            continue
        target = _is_target_folder(os.path.basename(current), level, supplier_folder)
        for entry in entries:
            if entry.is_dir(follow_symlinks=False):
                stack.append(entry.path)
            elif target and entry.name.lower().endswith((".xlsx", ".xlsm")) \
                    and not entry.name.startswith(("~$", SUMMARY_REPORT_PREFIX)):
                st = entry.stat()
                # 來源資料夾：辨識等級資料夾的上一層 (例如客戶/專案資料夾名稱)
                folder = os.path.basename(os.path.dirname(current))
                yield entry.path, folder, st.st_mtime, st.st_size


# ============================================================
# 讀檔 + 清洗 (在子 process 裡執行)
# ============================================================

def _read_rows(path: str) -> List[list]:
    """讀第一個工作表的所有列 (值)。"""
    if CalamineWorkbook is not None:
        wb = CalamineWorkbook.from_path(path)
        return wb.get_sheet_by_index(0).to_python(skip_empty_area=False)
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        return [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    finally:
        wb.close()


def _header_code(value) -> Optional[str]:
    if value is None:
        return None
    first_line = str(value).strip().split("\n")[0].strip().upper().rstrip(".").strip()
    return HEADER_FIELDS.get(first_line)


def _find_header(rows: List[list]) -> Tuple[int, Dict[int, str]]:
    """回傳 (標題列 index, {欄 index: 欄位代碼})；找不到丟出 ValueError。"""
    for i, row in enumerate(rows[:HEADER_SEARCH_ROWS]):
        mapping = {}
        for col, value in enumerate(row):
            code = _header_code(value)
            if code and code not in mapping.values():
                mapping[col] = code
        if {"description", "amount"} <= set(mapping.values()) and len(mapping) >= 4:
            return i, mapping
    raise ValueError("找不到標題列 (前幾列沒有 Description / AMOUNT 等欄位)")


def clean_record(raw: Dict[str, object]) -> Dict[str, object]:
    """一列原始值 -> 清洗後的明細 (不含 source_file/source_folder)。"""
    invoice_date = parse_date(raw.get("invoice_date"))
    item, category = normalize_expense(raw.get("description"))
    return {
        "invoice_date": invoice_date,
        "year": invoice_date.year if invoice_date else None,
        "month": invoice_date.month if invoice_date else None,
        "invoice_no": clean_text(raw.get("invoice_no")),
        "supplier": clean_text(raw.get("supplier")),
        "consignee": clean_text(raw.get("consignee")),
        "description_raw": clean_text(raw.get("description")),
        "expense_item": item,
        "expense_category": category,
        "amount": parse_amount(raw.get("amount")),
        "vat": parse_amount(raw.get("vat")),
        "order_no": clean_text(raw.get("order_no")),
        "arrive_date": parse_date(raw.get("arrive_date")),
        "onboard_date": parse_date(raw.get("onboard_date")),
        "bl_no": clean_text(raw.get("bl_no")),
        "mbl_no": clean_text(raw.get("mbl_no")),
        "port_of_loading": normalize_port(raw.get("port_of_loading")),
        "port_of_discharge": normalize_port(raw.get("port_of_discharge")),
        "volume": parse_float(raw.get("volume")),
        "vessel": clean_text(raw.get("vessel")),
        "container_no": clean_text(raw.get("container_no")),
        "shipper": clean_text(raw.get("shipper")),
    }


def load_file(task: Tuple[str, str]) -> Tuple[str, List[tuple], Optional[str]]:
    """讀一份 Excel 並清洗，回傳 (路徑, 明細列 tuple 清單, 錯誤原因)。
    任何例外都轉成錯誤原因回傳，不讓單一壞檔中斷整批。"""
    path, folder = task
    try:
        rows = _read_rows(path)
        header_index, mapping = _find_header(rows)
        records = []
        for row in rows[header_index + 1:]:
            raw = {code: row[col] for col, code in mapping.items() if col < len(row)}
            # 範本後面的空白列、或只剩框線的列直接略過
            if all(clean_text(v) is None for v in raw.values()):
                continue
            rec = clean_record(raw)
            rec["source_file"] = path
            rec["source_folder"] = folder
            records.append(tuple(rec[name] for name, _ in ITEM_COLUMNS))
        return path, records, None
    except Exception as e:  # noqa: BLE001  壞檔可能丟出各種例外
        return path, [], f"{type(e).__name__}: {e}"


# ============================================================
# 資料庫
# ============================================================

def init_db(con: duckdb.DuckDBPyConnection) -> None:
    cols = ", ".join(f"{name} {typ}" for name, typ in ITEM_COLUMNS)
    con.execute(f"CREATE TABLE IF NOT EXISTS expense_items ({cols}, "
                "is_duplicate BOOLEAN DEFAULT FALSE)")
    con.execute("""CREATE TABLE IF NOT EXISTS etl_files (
        source_file VARCHAR PRIMARY KEY, mtime DOUBLE, size BIGINT,
        row_count INTEGER, error VARCHAR, processed_at TIMESTAMP)""")


def _delete_files(con: duckdb.DuckDBPyConnection, paths: List[str]) -> None:
    """刪除這些檔案的明細與處理紀錄 (檔案有修改要重讀、或已不存在)。"""
    if not paths:
        return
    con.register("drop_files", pd.DataFrame({"source_file": paths}))
    con.execute("DELETE FROM expense_items WHERE source_file IN "
                "(SELECT source_file FROM drop_files)")
    con.execute("DELETE FROM etl_files WHERE source_file IN "
                "(SELECT source_file FROM drop_files)")
    con.unregister("drop_files")


def _flush(con: duckdb.DuckDBPyConnection, items: List[tuple], files: List[tuple]) -> None:
    """把累積的明細與檔案紀錄寫進資料庫，寫完清空緩衝。"""
    if items:
        df = pd.DataFrame(items, columns=[name for name, _ in ITEM_COLUMNS])
        for col in _DATE_COLUMNS:
            df[col] = pd.to_datetime(df[col])
        con.register("buf_items", df)
        select = ", ".join(f"CAST({name} AS {typ})" for name, typ in ITEM_COLUMNS)
        con.execute(f"INSERT INTO expense_items ({', '.join(n for n, _ in ITEM_COLUMNS)}) "
                    f"SELECT {select} FROM buf_items")
        con.unregister("buf_items")
        items.clear()
    if files:
        con.register("buf_files", pd.DataFrame(
            files, columns=["source_file", "mtime", "size", "row_count", "error", "processed_at"]))
        con.execute("INSERT OR REPLACE INTO etl_files SELECT * FROM buf_files")
        con.unregister("buf_files")
        files.clear()


def mark_duplicates(con: duckdb.DuckDBPyConnection) -> int:
    """同一個 (供應商, 發票號碼) 出現在好幾份 Excel 時，只保留路徑排序第一
    份，其餘的明細標記 is_duplicate。回傳被標記的列數。"""
    con.execute("UPDATE expense_items SET is_duplicate = FALSE WHERE is_duplicate")
    con.execute("""
        UPDATE expense_items AS e SET is_duplicate = TRUE
        FROM (
            SELECT coalesce(supplier, '') AS sup, invoice_no, min(source_file) AS keep
            FROM expense_items
            WHERE invoice_no IS NOT NULL
            GROUP BY 1, 2
            HAVING count(DISTINCT source_file) > 1
        ) AS d
        WHERE coalesce(e.supplier, '') = d.sup AND e.invoice_no = d.invoice_no
          AND e.source_file <> d.keep""")
    return con.execute("SELECT count(*) FROM expense_items WHERE is_duplicate").fetchone()[0]


# ============================================================
# 主流程
# ============================================================

def run(roots: List[Path], db_path: Path, level: str, supplier_folder: Optional[str],
        workers: int, full: bool, parquet: Optional[Path]) -> int:
    log = get_logger()
    started = time.time()
    db_path.parent.mkdir(parents=True, exist_ok=True)
    con = duckdb.connect(str(db_path))
    init_db(con)
    if full:
        con.execute("DELETE FROM expense_items")
        con.execute("DELETE FROM etl_files")

    print(f"掃描資料夾 (辨識等級({level})"
          + (f"、供應商資料夾 {supplier_folder}" if supplier_folder else "") + ") ...")
    scanned = {path: (folder, mtime, size)
               for path, folder, mtime, size in scan_excel_files(roots, level, supplier_folder)}
    known = {path: (mtime, size) for path, mtime, size in
             con.execute("SELECT source_file, mtime, size FROM etl_files").fetchall()}

    todo = [p for p, (_, mtime, size) in scanned.items() if known.get(p) != (mtime, size)]
    # 資料庫有紀錄但檔案已不存在 (被搬走/刪除) 的 -> 移除；只是這次沒掃到
    # (例如換了 --supplier-folder) 但檔案還在的保留
    removed = [p for p in known if p not in scanned and not os.path.exists(p)]
    _delete_files(con, [p for p in todo if p in known] + removed)
    print(f"找到 {len(scanned):,} 份 Excel：新增/修改 {len(todo):,} 份、"
          f"未變動 {len(scanned) - len(todo):,} 份、已移除 {len(removed):,} 份")

    items: List[tuple] = []
    files: List[tuple] = []
    total_rows = failed = 0
    tasks = [(p, scanned[p][0]) for p in todo]
    if tasks:
        with ProcessPoolExecutor(max_workers=workers) as pool:
            for done, (path, records, error) in enumerate(
                    pool.map(load_file, tasks, chunksize=16), start=1):
                _, mtime, size = scanned[path]
                if error:
                    failed += 1
                    log.error("%s | 讀取失敗 | %s", path, error)
                elif not records:
                    log.warning("%s | 沒有任何明細列", path)
                items.extend(records)
                total_rows += len(records)
                files.append((path, mtime, size, len(records), error, datetime.now()))
                if len(items) >= BATCH_ROWS:
                    _flush(con, items, files)
                if done % 500 == 0 or done == len(tasks):
                    rate = done / max(time.time() - started, 1e-6)
                    print(f"  已處理 {done:,}/{len(tasks):,} 份 ({rate:,.0f} 份/秒)，"
                          f"明細 {total_rows:,} 列，失敗 {failed:,} 份")
        _flush(con, items, files)

    duplicates = mark_duplicates(con)
    if parquet:
        parquet.parent.mkdir(parents=True, exist_ok=True)
        con.execute(f"COPY (SELECT * FROM expense_items) TO '{parquet.as_posix()}' "
                    "(FORMAT PARQUET, COMPRESSION ZSTD)")
        print(f"已匯出 Parquet：{parquet}")

    summary = con.execute("""
        SELECT coalesce(CAST(year AS VARCHAR), '日期無法辨識'), count(*),
               count(DISTINCT coalesce(supplier, '') || '|' || coalesce(invoice_no, source_file)),
               sum(amount)
        FROM expense_items WHERE NOT is_duplicate GROUP BY 1 ORDER BY 1""").fetchall()
    con.close()

    print("=" * 60)
    print(f"完成，耗時 {time.time() - started:,.1f} 秒；資料庫：{db_path}")
    print(f"本次寫入 {total_rows:,} 列，讀取失敗 {failed:,} 份 (原因見 {LOG_PATH})；"
          f"重複發票 {duplicates:,} 列不計入")
    for year, rows, invoices, amount in summary:
        print(f"  {year}：發票 {invoices:,} 張、明細 {rows:,} 列、金額 {amount or 0:,.0f}")
    print("=" * 60)
    log.info("ETL 完成 | 新增/修改 %d 份 | 寫入 %d 列 | 失敗 %d 份", len(todo), total_rows, failed)
    return 0


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="彙整轉檔結果 Excel 進 DuckDB")
    parser.add_argument("roots", nargs="+", type=Path, help="要掃描的資料夾 (可多個，含所有子資料夾)")
    parser.add_argument("--db", type=Path, default=DEFAULT_DB, help=f"DuckDB 檔案 (預設 {DEFAULT_DB})")
    parser.add_argument("--level", default="高", help="只讀這個辨識等級的資料夾 (預設 高)")
    parser.add_argument("--supplier-folder", help="只讀名稱以此開頭的辨識等級資料夾，例如 MAERSK")
    parser.add_argument("--workers", type=int, default=max((os.cpu_count() or 2) - 1, 1),
                        help="平行讀檔的 process 數 (預設 CPU 核心數 - 1)")
    parser.add_argument("--full", action="store_true", help="清空資料庫重新匯入")
    parser.add_argument("--parquet", type=Path, help="另外匯出一份 Parquet 檔")
    args = parser.parse_args(argv)

    missing = [str(r) for r in args.roots if not r.is_dir()]
    if missing:
        print(f"找不到資料夾：{', '.join(missing)}")
        return 1
    return run([r.resolve() for r in args.roots], args.db.resolve(), args.level,
               args.supplier_folder, args.workers, args.full, args.parquet)


if __name__ == "__main__":
    sys.exit(main())
