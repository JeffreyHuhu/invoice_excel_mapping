# -*- coding: utf-8 -*-
"""
build_pivot_template.py — 用本機 Excel 產生樞紐範本 templates/pivot_template.xlsx

只在改了 pivot_export.py 的欄位或想換樞紐版面時執行一次 (需要 Excel 與
pywin32)。匯出時 pivot_export.write_workbook() 只換範本裡的資料，不需要 Excel。

範本內容 (使用者 2026-10-27 指定)：
  儀表板        ：交叉分析篩選器 (國家/供應商/年/費用大類，連動全部樞紐) + 4 張樞紐分析圖
  樞紐_供應商月份：年、月 × 供應商
  樞紐_費用大類  ：費用大類 × 供應商
  樞紐_集團公司  ：集團公司 × 費用大類
  樞紐_航線      ：航線 (啟運港 → 目的港) 前 20 名
  明細 / 讀取失敗：Excel 表格，只放一列範例 (不含客戶資料，可 commit)
樞紐分析表不存快取資料 (SaveData=False)、開檔時自動重新整理。

用法：python build_pivot_template.py
"""

import re
import sys
import zipfile

import win32com.client

from pivot_export import (COLUMNS, DETAIL_SHEET, DETAIL_TABLE, FAILED_COLUMNS,
                          FAILED_SHEET, FAILED_TABLE, NUMBER_FORMATS, TEMPLATE)

# Excel 常數
XL_SRC_RANGE, XL_YES, XL_DATABASE = 1, 1, 1
XL_ROW, XL_COLUMN = 1, 2
XL_SUM = -4157
XL_DESCENDING, XL_TOP_COUNT = 2, 1
XL_COLUMN_STACKED, XL_BAR_STACKED, XL_BAR_CLUSTERED = 52, 58, 57
XL_PIVOT_VERSION = 5          # Excel 2013 以上
XL_MISSING_ITEMS_NONE = 0     # 重新整理後不保留已不存在的項目
XLSX = 51

PLACEHOLDER = "(請按「啟用編輯」更新資料)"
USD = "AMOUNT (USD)"
VALUE_CAPTION = "USD 金額"


def title(name: str) -> str:
    return dict((c, t) for c, t, _ in COLUMNS)[name]


def sample_row() -> list:
    """範本的範例列：文字欄放提示，數字為 0 (樞紐分析表需要至少一列資料)。"""
    values = {"date": "2026/1/1", "num": 0, "rate": 1, "int": 1}
    row = [values.get(kind, PLACEHOLDER) for _, _, kind in COLUMNS]
    row[[c for c, _, _ in COLUMNS].index("year")] = 2026
    row[[c for c, _, _ in COLUMNS].index("year_month")] = "2026-01"
    row[[c for c, _, _ in COLUMNS].index("quarter")] = "Q1"
    return row


SHEET_ORDER = ["儀表板", "樞紐_供應商月份", "樞紐_費用大類", "樞紐_集團公司", "樞紐_航線",
               DETAIL_SHEET, FAILED_SHEET]


def add_last(wb):
    """新增工作表 (pywin32 動態呼叫的 After= 無效，順序最後由 SHEET_ORDER 排)。"""
    return wb.Worksheets.Add()


def strip_personal_info(path) -> None:
    """清掉範本裡的作者與「重新整理者」(Excel 會寫入 Windows 使用者名稱；repo 是公開的)。"""
    tmp = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "docProps/core.xml" or "pivotCacheDefinition" in item.filename:
                xml = data.decode("utf-8")
                xml = re.sub(r'refreshedBy="[^"]*"', 'refreshedBy=""', xml)
                xml = re.sub(r"(<(dc:creator|cp:lastModifiedBy)>)[^<]*(</\2>)", r"\1\3", xml)
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    tmp.replace(path)


def make_table(ws, columns, row, table_name):
    """標題 + 一列範例 -> Excel 表格，設定欄寬與數字格式。"""
    n = len(columns)
    ws.Range(ws.Cells(1, 1), ws.Cells(1, n)).Value = [[t for _, t, _ in columns]]
    ws.Range(ws.Cells(2, 1), ws.Cells(2, n)).Value = [row]
    lo = ws.ListObjects.Add(XL_SRC_RANGE, ws.Range(ws.Cells(1, 1), ws.Cells(2, n)), None, XL_YES)
    lo.Name = table_name
    lo.TableStyle = "TableStyleMedium2"
    for j, (_, t, kind) in enumerate(columns, start=1):
        if kind in NUMBER_FORMATS:
            lo.ListColumns(j).DataBodyRange.NumberFormat = NUMBER_FORMATS[kind]
        ws.Columns(j).ColumnWidth = min(max(len(t) + 4, 12), 40)
    return lo


def add_pivot(wb, cache, sheet_name, heading, rows, cols, top_n=None):
    ws = add_last(wb)
    ws.Name = sheet_name
    ws.Range("A1").Value = heading
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    pt = cache.CreatePivotTable(TableDestination=ws.Range("A3"),
                                TableName="PT_" + sheet_name.split("_", 1)[1],
                                DefaultVersion=XL_PIVOT_VERSION)
    for i, f in enumerate(rows, start=1):
        pt.PivotFields(f).Orientation = XL_ROW
        pt.PivotFields(f).Position = i
    for i, f in enumerate(cols, start=1):
        pt.PivotFields(f).Orientation = XL_COLUMN
        pt.PivotFields(f).Position = i
    df = pt.AddDataField(pt.PivotFields(USD), VALUE_CAPTION, XL_SUM)
    df.NumberFormat = "#,##0.00"
    # 第一層列欄位依金額由大到小 (年/月維持時間順序)
    first = pt.PivotFields(rows[0])
    if rows[0] not in (title("year"), title("month")):
        first.AutoSort(XL_DESCENDING, VALUE_CAPTION)
        if top_n:
            first.PivotFilters.Add2(Type=XL_TOP_COUNT, DataField=df, Value1=top_n)
    pt.SaveData = False
    pt.HasAutoFormat = False   # 重新整理時不要自動改欄寬
    return pt


def add_chart(ws, pt, chart_type, chart_title, left, top, width=520, height=320):
    chart = ws.Shapes.AddChart2(-1, chart_type, left, top, width, height).Chart
    chart.SetSourceData(pt.TableRange1)
    chart.HasTitle = True
    chart.ChartTitle.Text = chart_title
    chart.ShowAllFieldButtons = False


def main() -> int:
    TEMPLATE.parent.mkdir(exist_ok=True)
    xl = win32com.client.DispatchEx("Excel.Application")   # 獨立程序，不影響使用者開著的 Excel
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count > 1:
            wb.Worksheets(2).Delete()
        dash = wb.Worksheets(1)
        dash.Name = "儀表板"

        detail = add_last(wb)
        detail.Name = DETAIL_SHEET
        make_table(detail, COLUMNS, sample_row(), DETAIL_TABLE)
        cache = wb.PivotCaches().Create(SourceType=XL_DATABASE, SourceData=DETAIL_TABLE,
                                        Version=XL_PIVOT_VERSION)

        supplier, category = title("supplier"), title("expense_category")
        pts = [
            add_pivot(wb, cache, "樞紐_供應商月份", "各供應商每月費用 (USD)",
                      [title("year"), title("month")], [supplier]),
            add_pivot(wb, cache, "樞紐_費用大類", "費用大類 × 供應商 (USD)",
                      [category], [supplier]),
            add_pivot(wb, cache, "樞紐_集團公司", "集團公司 × 費用大類 (USD)",
                      [title("consignee")], [category]),
            add_pivot(wb, cache, "樞紐_航線", "航線費用前 20 名 (USD)",
                      [title("route")], [], top_n=20),
        ]
        cache.RefreshOnFileOpen = True
        cache.MissingItemsLimit = XL_MISSING_ITEMS_NONE

        failed = add_last(wb)
        failed.Name = FAILED_SHEET
        make_table(failed, FAILED_COLUMNS, ["-", "-"], FAILED_TABLE)

        # 儀表板：標題、交叉分析篩選器 (連動全部樞紐)、4 張樞紐分析圖
        dash.Range("A1").Value = "物流費用分析儀表板 (金額：USD)"
        dash.Range("A1").Font.Bold = True
        dash.Range("A1").Font.Size = 16
        dash.Range("A2").Value = ("點選上方篩選器可同時篩選全部圖表與樞紐分析表。"
                                  "資料在「明細」頁；樞紐分析表在「樞紐_」各頁。")
        for i, field in enumerate([title("country"), supplier, title("year"), category]):
            sc = wb.SlicerCaches.Add2(pts[0], field)
            sc.Slicers.Add(SlicerDestination=dash, Name=f"Slicer_{i}", Caption=field,
                           Top=45, Left=10 + i * 262, Width=250, Height=150)
            for pt in pts[1:]:
                sc.PivotTables.AddPivotTable(pt)
        add_chart(dash, pts[0], XL_COLUMN_STACKED, "各供應商每月費用 (USD)", 10, 205)
        add_chart(dash, pts[1], XL_BAR_STACKED, "費用大類 × 供應商 (USD)", 540, 205)
        add_chart(dash, pts[2], XL_BAR_STACKED, "集團公司 × 費用大類 (USD)", 10, 535)
        add_chart(dash, pts[3], XL_BAR_CLUSTERED, "航線費用前 20 名 (USD)", 540, 535)

        # 依 SHEET_ORDER 排列：由後往前逐一移到最前面
        for name in reversed(SHEET_ORDER):
            wb.Worksheets(name).Move(wb.Worksheets(1))
        dash.Activate()
        dash.Range("A1").Select()
        if TEMPLATE.exists():
            TEMPLATE.unlink()
        wb.SaveAs(str(TEMPLATE), FileFormat=XLSX)
        wb.Close(False)
    finally:
        xl.Quit()
    strip_personal_info(TEMPLATE)
    print(f"已產生樞紐範本：{TEMPLATE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
