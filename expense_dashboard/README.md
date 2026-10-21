# 費用分析系統 (expense_dashboard)

把轉檔工具輸出的「辨識等級(高)」Excel 彙整成資料庫，給主管用網頁儀表板分析 2025–2026 年各項費用。

## 安裝

```
pip install -r requirements.txt
```

## 第一階段：彙整資料 (process_data.py)

```
python process_data.py "D:\發票\02-Maersk Logistics Indonesia"
python process_data.py <資料夾1> <資料夾2> --supplier-folder MAERSK
python process_data.py <資料夾> --full --parquet data\expenses.parquet
```

- 掃描資料夾底下所有 `XXX-辨識等級(高)` 子資料夾裡的 Excel (略過轉檔總表與 `~$` 暫存檔)。
- 多 process 平行讀檔、每 20,000 列寫入一次資料庫，記憶體用量固定 (實測 1 萬份約 16 秒)。
- 增量更新：再跑一次只處理新增/修改的檔案；檔案被刪除或搬走時，它的明細會從資料庫移除。`--full` 整個重建。
- 同一張發票 (供應商 + 發票號碼) 出現在好幾份 Excel 時只算一次 (其餘標 `is_duplicate`)。
- 讀不了的檔案記在 `etl_files` 資料表與 `logs/expense_etl.log`，不中斷整批。
- 結果存在 `data/expenses.duckdb` (`--db` 可改位置)。

### 清洗規則 (expense_normalize.py)

| 欄位 | 規則 |
|---|---|
| INVOICE DATE | `31.Aug.2026`、`2026-08-31`、`31/08/2026`、印尼文月份 (`9 Okt 2025`)、Excel 日期序號等都轉成標準日期，另存年份/月份 |
| AMOUNT / VAT | 去掉 `Rp`/`IDR`、千分位 (`1.234.567` 或 `1,234,567`)、括號負數，轉成數值 |
| Description | 轉大寫、去標點與追蹤號碼、展開縮寫 (THC、Doc、Chg…)，再對照 `EXPENSE_ALIASES` 合併成標準費用項目，並依關鍵字分到費用大類 (`EXPENSE_CATEGORIES`)；原始名稱保留在 `description_raw` |
| 港口 | 統一大小寫 (`LOS ANGELES` → `Los Angeles`) |

發現有同性質費用沒被合併時，在 `EXPENSE_ALIASES` 加一行對照，然後執行 `python process_data.py <資料夾> --full` 重建。

## 第二階段：儀表板 (app.py)

```
streamlit run app.py
```

- 側邊欄篩選：年份、供應商、費用大類、費用項目、啟運港、目的港、集團公司。
- KPI：總費用金額、總 VAT、總發票數、平均每張發票金額，以及最新年份跟前一年**同期** (1 月到最新月份) 比較的增減率。
- 圖表：費用項目排行榜、2025 vs 2026 月度趨勢、目的港 × 費用大類堆疊圖、航線排行。
- 費用項目年度比較表 (各年金額、增減金額、增減率)。
- 明細表格 (畫面顯示最新 2,000 筆) 與匯出全部符合條件的明細 (CSV / Excel)。

資料庫在別的位置時，可以在側邊欄輸入路徑，或設定環境變數 `EXPENSE_DB`。
儀表板只用唯讀連線讀資料庫，開著時也可以同時跑 process_data.py 更新資料。
