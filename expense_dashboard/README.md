# 費用分析系統 (expense_dashboard)

把轉檔工具輸出的「辨識等級(高)」Excel 彙整成資料庫，給主管用網頁儀表板分析各項費用 (可換算 USD)。

## 給使用者：expense_dashboard.exe

雙擊 `expense_dashboard.exe` → 自動打開瀏覽器顯示儀表板 (不用安裝 Python；程式只在自己電腦上執行，不對外開放)。
黑色視窗是程式本身，關掉它就結束。第一次執行防毒軟體會先掃描，可能要等 1~2 分鐘。

左側選單有兩頁：

- **費用分析儀表板**：篩選、KPI、圖表、明細匯出 (CSV / Excel)。側邊欄「幣別」可切換 **換算 USD** (每月平均匯率或自訂固定匯率) 或 **原幣**。
- **資料管理**：
  1. 資料庫位置：指向要看的 `expenses.duckdb` (例如公司共用資料夾)，記在 EXE 旁邊的 `dashboard_config.json`。
  2. 從轉檔結果更新資料：輸入資料夾、按「開始更新」(只處理新增/修改的 Excel)。
  3. 匯率：一鍵抓 ECB 月平均匯率，或在表格裡手動修改 (手動的不會被自動更新覆蓋)。
  4. 讀取失敗的檔案清單。

### 資料怎麼佈署

EXE 不含資料；資料是另一個檔案 `expenses.duckdb` (MAERSK 2025–2026 約 55 MB)。

| 方式 | 做法 | 適合 |
|---|---|---|
| **共用資料夾 (建議)** | 把 `expenses.duckdb` 放在公司共用磁碟/NAS；每個人第一次開 EXE 時在「資料管理」指定一次路徑。資料由一人更新，大家看到的都是最新的。 | 團隊多人使用 |
| **跟 EXE 一起發** | 資料夾裡放 `expense_dashboard.exe` + `_internal\` + `data\expenses.duckdb`，壓縮後發給同事。EXE 預設就讀旁邊的 `data\expenses.duckdb`。 | 離線、少數人 |

注意：拿到資料庫檔就能看到全部發票資料，只在公司內部流通，不要放進公開的 GitHub repo (repo 已 gitignore `data/`)。
同一時間只能有一個人更新資料；更新期間其他人的儀表板會提示「資料庫正在更新中」，稍後重新整理即可。

## 給維護者

```
pip install -r requirements.txt
streamlit run app.py                      # 開發時直接跑網頁
python process_data.py "<轉檔結果所在資料夾>" [--supplier-folder MAERSK] [--full]
python -m PyInstaller --noconfirm expense_dashboard.spec   # 打包 dist\expense_dashboard\ (EXE + _internal\)
```

打包用 onedir 模式：EXE 執行時不解壓縮到 `%TEMP%`，啟動較快，也不會留下 `_MEIxxxx` 暫存資料夾。
發佈時把 `dist\expense_dashboard\` 裡的內容整份複製過去，EXE 與 `_internal\` 要放同一層。

| 檔案 | 內容 |
|---|---|
| `app.py` | 入口，切換兩個頁面 |
| `dashboard.py` / `data_admin.py` | 儀表板頁 / 資料管理頁 |
| `process_data.py` | 彙整 (多 process 讀檔、每 2 萬列寫入一次、增量更新、重複發票標記、更新匯率) |
| `expense_normalize.py` | 清洗規則：日期、金額、幣別、費用名稱對照 (`EXPENSE_ALIASES`)、費用大類 (`EXPENSE_CATEGORIES`)、港口 |
| `fx.py` | 匯率表與換算 USD 的 SQL |
| `settings.py` | 路徑與設定檔 (EXE 時以 EXE 所在資料夾為準) |
| `launcher.py` / `dashboard_deps.py` / `expense_dashboard.spec` | EXE 打包 |

### 幣別與匯率

- 每筆明細記錄原幣別，判斷順序：1. 帳單上的幣別 (金額文字的代碼/符號：USD、Rp、₫、RM、฿、₹、NT$、¥…)；2. 供應商 (`SUPPLIER_CURRENCY` 指定名稱 → 公司名稱國別特徵，例如 CÔNG TY/TNHH → VND、PT → IDR、SDN BHD → MYR)；3. 國家資料夾 (`供應商轉檔結果\<國家>\`，`COUNTRY_FOLDER_CURRENCY`)；4. 預設 IDR。目前支援中國、東南亞、印度、台灣；新增國家要補 `expense_normalize.py` 的規則。
- 匯率表 `fx_rates`：每幣別每月一筆「1 USD = ?」，來源為歐洲央行 (Frankfurter API) 每日參考匯率的月平均；只送出幣別與日期，不送任何發票資料。當月沒有匯率時用最近一個月。
- 修改費用名稱對照規則後要「完整重建」(`--full`)；完整重建會開新檔 (避免資料庫越用越大)，手動輸入的匯率會保留。
