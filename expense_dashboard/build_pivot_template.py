# -*- coding: utf-8 -*-
"""
build_pivot_template.py — 用本機 Excel 產生樞紐範本 templates/pivot_template.xlsx

只在改了 pivot_export.py 的欄位或想換樞紐版面時執行一次 (需要 Excel 與
pywin32)。匯出時 pivot_export.write_workbook() 只換範本裡的資料，不需要 Excel。

範本內容 (使用者 2026-10-27 指定，給高階主管)：
  儀表板_總覽    ：KPI (總費用/發票張數/TEU/CBM/每 TEU 成本)、交叉分析篩選器、
                  每月 × 費用大類、今年 vs 去年、各供應商每月、費用大類 × 供應商
  儀表板_運量航線：交叉分析篩選器、每月櫃型櫃數、航線前 20 名、集團公司 × 費用大類
  樞紐_*         ：各圖表的樞紐分析表，另有供應商排名與占比、單位成本 (供應商/航線)
  明細 / 讀取失敗：Excel 表格，只放一列範例 (不含客戶資料，可 commit)
交叉分析篩選器 (國家/供應商/年/費用大類) 連動全部樞紐。樞紐分析表不存快取資料
(SaveData=False)、開檔時自動重新整理。每 TEU 成本 = FCL 費用 / TEU (樞紐計算欄位)；
不算每 CBM 成本 (使用者 2026-10-27)。

用法：python build_pivot_template.py
"""

import re
import sys
import zipfile

import win32com.client

from pivot_export import (COLUMNS, CONTAINER_TYPES, DETAIL_SHEET, DETAIL_TABLE, FAILED_COLUMNS,
                          FAILED_SHEET, FAILED_TABLE, NUMBER_FORMATS, TEMPLATE)

# Excel 常數
XL_SRC_RANGE, XL_YES, XL_DATABASE = 1, 1, 1
XL_ROW, XL_COLUMN = 1, 2
XL_SUM = -4157
XL_PERCENT_OF_TOTAL = 8
XL_DESCENDING, XL_TOP_COUNT = 2, 1
XL_LINE, XL_COLUMN_STACKED, XL_BAR_STACKED, XL_BAR_CLUSTERED = 4, 52, 58, 57
XL_PIVOT_VERSION = 5          # Excel 2013 以上
XL_MISSING_ITEMS_NONE = 0     # 重新整理後不保留已不存在的項目
XLSX = 51

PLACEHOLDER = "(請按「啟用編輯」更新資料)"
DASH_MAIN, DASH_VOLUME = "儀表板_總覽", "儀表板_運量航線"
MONEY, QTY, COUNT, PCT = "#,##0.00", "#,##0.000", "#,##0", "0.0%"


def title(name: str) -> str:
    return dict((c, t) for c, t, _ in COLUMNS)[name]


# 樞紐的值欄位：(來源欄位或計算欄位, 標題, 數字格式[, 計算方式])
USD_VALUE = (title("amount"), "USD 金額", MONEY)
INVOICES = (title("invoice_count"), "發票數", COUNT)
TEU = (title("teu"), "TEU 數", COUNT)
CBM = (title("cbm"), "CBM 數", QTY)
FCL_USD = (title("fcl_usd"), "FCL 費用", MONEY)
LCL_USD = (title("lcl_usd"), "LCL 費用", MONEY)
PER_TEU = ("每TEU成本", "每 TEU 成本 (USD)", MONEY)
# 樞紐計算欄位 (建在共用快取上，全部樞紐都能用)：sum(分子) / sum(分母)
CALCULATED = {"每TEU成本": f"='{title('fcl_usd')}'/'{title('teu')}'"}

PIVOT_SHEETS = ["樞紐_每月費用大類", "樞紐_年度比較", "樞紐_供應商排名", "樞紐_供應商月份",
                "樞紐_費用大類", "樞紐_集團公司", "樞紐_運量", "樞紐_單位成本_供應商",
                "樞紐_單位成本_航線", "樞紐_航線"]
SHEET_ORDER = [DASH_MAIN, DASH_VOLUME] + PIVOT_SHEETS + [DETAIL_SHEET, FAILED_SHEET]


def sample_row() -> list:
    """範本的範例列：文字欄放提示，數字為 0 (樞紐分析表需要至少一列資料)。"""
    values = {"date": "2026/1/1", "num": 0, "rate": 1, "int": 1, "count": 0, "qty": 0}
    row = [values.get(kind, PLACEHOLDER) for _, _, kind in COLUMNS]
    row[[c for c, _, _ in COLUMNS].index("year")] = 2026
    row[[c for c, _, _ in COLUMNS].index("year_month")] = "2026-01"
    row[[c for c, _, _ in COLUMNS].index("quarter")] = "Q1"
    return row


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


def add_pivot(wb, cache, sheet_name, heading, rows, cols, values, top_n=None, sort=True,
              dest=None, name=None):
    """建一張樞紐分析表。values：[(欄位, 標題, 格式[, 計算方式])]，第一個值用來排序/取前 N 名。
    dest：(工作表, 儲存格) 放在既有工作表 (例如儀表板的 KPI)；否則新增 sheet_name 工作表。"""
    if dest:
        ws, cell = dest
    else:
        ws, cell = add_last(wb), "A3"
        ws.Name = sheet_name
        ws.Range("A1").Value = heading
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Size = 14
    pt = cache.CreatePivotTable(TableDestination=ws.Range(cell),
                                TableName=name or "PT_" + sheet_name.split("_", 1)[1],
                                DefaultVersion=XL_PIVOT_VERSION)
    if pt.CalculatedFields().Count == 0:   # 計算欄位在共用快取上，只要建一次
        for calc_name, formula in CALCULATED.items():
            pt.CalculatedFields().Add(calc_name, formula, True)
    for i, f in enumerate(rows, start=1):
        pt.PivotFields(f).Orientation = XL_ROW
        pt.PivotFields(f).Position = i
    for i, f in enumerate(cols, start=1):
        pt.PivotFields(f).Orientation = XL_COLUMN
        pt.PivotFields(f).Position = i
    data_fields = []
    for field, caption, fmt, *calc in values:
        df = pt.AddDataField(pt.PivotFields(field), caption, XL_SUM)
        df.NumberFormat = fmt
        if calc:
            df.Calculation = calc[0]
        data_fields.append(df)
    # 第一層列欄位依第一個值由大到小 (時間欄位用 sort=False 維持時間順序)
    if rows and sort:
        first = pt.PivotFields(rows[0])
        first.AutoSort(XL_DESCENDING, values[0][1])
        if top_n:
            first.PivotFilters.Add2(Type=XL_TOP_COUNT, DataField=data_fields[0], Value1=top_n)
    pt.ErrorString = "-"          # 0 TEU 的單位成本顯示 "-"
    pt.DisplayErrorString = True
    pt.SaveData = False
    pt.HasAutoFormat = False      # 重新整理時不要自動改欄寬
    return pt


def add_chart(ws, pt, chart_type, chart_title, left, top, width=520, height=320):
    chart = ws.Shapes.AddChart2(-1, chart_type, left, top, width, height).Chart
    chart.SetSourceData(pt.TableRange1)
    chart.HasTitle = True
    chart.ChartTitle.Text = chart_title
    chart.ShowAllFieldButtons = False


def dashboard_heading(ws, text, note):
    ws.Range("A1").Value = text
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    ws.Range("A2").Value = note


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
        dash.Name = DASH_MAIN
        dash2 = add_last(wb)
        dash2.Name = DASH_VOLUME

        detail = add_last(wb)
        detail.Name = DETAIL_SHEET
        make_table(detail, COLUMNS, sample_row(), DETAIL_TABLE)
        cache = wb.PivotCaches().Create(SourceType=XL_DATABASE, SourceData=DETAIL_TABLE,
                                        Version=XL_PIVOT_VERSION)

        supplier, category = title("supplier"), title("expense_category")
        year, month, year_month = title("year"), title("month"), title("year_month")
        unit_cost = [FCL_USD, TEU, PER_TEU, LCL_USD, CBM]
        kpi = add_pivot(wb, cache, "", "", [], [],
                        [USD_VALUE, INVOICES, TEU, CBM, PER_TEU],
                        dest=(dash, "A4"), name="PT_KPI")
        pt = {
            "每月費用大類": add_pivot(wb, cache, "樞紐_每月費用大類", "每月 × 費用大類 (USD)",
                                [year_month], [category], [USD_VALUE], sort=False),
            "年度比較": add_pivot(wb, cache, "樞紐_年度比較", "今年 vs 去年：各月費用 (USD)",
                              [month], [year], [USD_VALUE], sort=False),
            "供應商排名": add_pivot(wb, cache, "樞紐_供應商排名", "供應商費用排名與占比 (USD)",
                               [supplier], [],
                               [USD_VALUE,
                                (title("amount"), "占總費用 %", PCT, XL_PERCENT_OF_TOTAL),
                                INVOICES]),
            "供應商月份": add_pivot(wb, cache, "樞紐_供應商月份", "各供應商每月費用 (USD)",
                               [year, month], [supplier], [USD_VALUE], sort=False),
            "費用大類": add_pivot(wb, cache, "樞紐_費用大類", "費用大類 × 供應商 (USD)",
                              [category], [supplier], [USD_VALUE]),
            "集團公司": add_pivot(wb, cache, "樞紐_集團公司", "集團公司 × 費用大類 (USD)",
                              [title("consignee")], [category], [USD_VALUE]),
            "運量": add_pivot(wb, cache, "樞紐_運量",
                            "每月櫃型櫃數 (同一供應商的同一批貨只算一次)", [year_month], [],
                            [(title(f"cnt_{t}"), t, COUNT) for t in CONTAINER_TYPES],
                            sort=False),
            "單位成本_供應商": add_pivot(
                wb, cache, "樞紐_單位成本_供應商",
                "單位成本 (供應商)：每 TEU 成本 = FCL 費用 / TEU",
                [supplier], [], unit_cost),
            "單位成本_航線": add_pivot(
                wb, cache, "樞紐_單位成本_航線", "單位成本 (航線，TEU 前 20 名)",
                [title("route")], [], [TEU] + [v for v in unit_cost if v is not TEU], top_n=20),
            "航線": add_pivot(wb, cache, "樞紐_航線", "航線費用前 20 名 (USD)",
                            [title("route")], [], [USD_VALUE], top_n=20),
        }
        cache.RefreshOnFileOpen = True
        cache.MissingItemsLimit = XL_MISSING_ITEMS_NONE

        failed = add_last(wb)
        failed.Name = FAILED_SHEET
        make_table(failed, FAILED_COLUMNS, ["-", "-"], FAILED_TABLE)

        note = ("點選篩選器可同時篩選兩頁儀表板的全部圖表與樞紐分析表。"
                "資料在「明細」頁；樞紐分析表在「樞紐_」各頁。")
        dashboard_heading(dash, "物流費用分析儀表板：總覽 (金額：USD)", note)
        dashboard_heading(dash2, "物流費用分析儀表板：運量與航線 (金額：USD)", note)
        # 交叉分析篩選器：兩頁儀表板各放一組，同一個篩選快取，連動全部樞紐
        for i, field in enumerate([title("country"), supplier, year, category]):
            sc = wb.SlicerCaches.Add2(kpi, field)
            sc.Slicers.Add(SlicerDestination=dash, Name=f"Slicer_{i}_main", Caption=field,
                           Top=100, Left=10 + i * 262, Width=250, Height=130)
            sc.Slicers.Add(SlicerDestination=dash2, Name=f"Slicer_{i}_volume", Caption=field,
                           Top=45, Left=10 + i * 262, Width=250, Height=130)
            for p in pt.values():
                sc.PivotTables.AddPivotTable(p)
        add_chart(dash, pt["每月費用大類"], XL_COLUMN_STACKED, "每月 × 費用大類 (USD)", 10, 245)
        add_chart(dash, pt["年度比較"], XL_LINE, "今年 vs 去年：各月費用 (USD)", 540, 245)
        add_chart(dash, pt["供應商月份"], XL_COLUMN_STACKED, "各供應商每月費用 (USD)", 10, 575)
        add_chart(dash, pt["費用大類"], XL_BAR_STACKED, "費用大類 × 供應商 (USD)", 540, 575)
        add_chart(dash2, pt["運量"], XL_COLUMN_STACKED, "每月櫃型櫃數", 10, 190)
        add_chart(dash2, pt["航線"], XL_BAR_CLUSTERED, "航線費用前 20 名 (USD)", 540, 190)
        add_chart(dash2, pt["集團公司"], XL_BAR_STACKED, "集團公司 × 費用大類 (USD)", 10, 520)

        # 依 SHEET_ORDER 排列：由後往前逐一移到最前面
        for sheet in reversed(SHEET_ORDER):
            wb.Worksheets(sheet).Move(wb.Worksheets(1))
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
