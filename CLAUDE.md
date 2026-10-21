# 專案目標 (Project Goals)
- 自動化解析各式 supplier PDF 發票並準確對應至 Excel 欄位。
- 提供高效、穩定的批次處理流程。

# 最高指導原則 (Guiding Principles)
1. **效能優先**：處理大量 PDF 時，必須優先考慮記憶體與速度，善用批次 processing，嚴禁將大量檔案內容直接讀入 LLM 上下文。
2. **程式碼品質**：採用模組化設計，遵循 PEP 8 規範，並加上清晰的繁體中文註解。
3. **錯誤處理**：若遇到格式不符或損壞的 PDF，必須妥善捕捉 exception，紀錄到 log 檔中，不得中斷整體流程。
4. **極簡修改**：修改程式碼時，只動必要的模組，盡量保持現有邏輯的穩定性。

# 專案架構 (Project Structure)
- `extract_utils_streamlit.py`：與供應商無關的共用核心引擎 (PDF 文字擷取/OCR、供應商註冊表、標準欄位、比對評分、VAT/總額稽核、Excel 讀寫)。
- `suppliers_streamlit.py`：各供應商專屬的 `detect_xxx()` / `parse_xxx()`，以 `register_supplier()` 註冊。新增或修正供應商只改這支檔案；`detect_supplier()` 依註冊順序比對，關鍵字較寬鬆的規則要放後面。
  - Excel 帳單供應商 (越南 VN03 RỒNG Á CHÂU、VN05 CONG THANH)：用 `register_supplier(..., sheet_parse=, sheet_totals=)` 註冊；核心只交出「顯示中」的工作表 (`read_spreadsheet()`，隱藏的是供應商的工作副本)，detect 吃 `sheet_text()` (以 `SHEET_TEXT_MARK` 開頭，只跟 Excel 供應商比對)，`sheet_totals` 取帳單本身合計列做金額勾稽。
  - 供應商編號 `SUPPLIER_NUMBERS`：印尼 ID01~ (Top15 + 同一份金額排名)、越南 VN01~VN20 (Vietnam Supplier Top 20)。
- `app.py`：Streamlit 自動比對網頁 (上傳 PDF + 正確答案 Excel → 逐欄比對、計算正確率)。
- `verify_app.py`：Streamlit 人工查核網頁 (只上傳 PDF → 逐欄標記錯誤、收集回報)。
- `run_regression.py`：回歸測試，樣本放在 `test_samples/<SUPPLIER_KEY>/sampleN.pdf` + `sampleN_ref.xlsx`。
- `learned_configs.json`：各供應商學到的最佳擷取參數。
- `validation_streamlit.py`：驗證層。不需正確答案，以必填欄位 (INVOICE DATE/INVOICE NO/SUPPLIER/CONSIGNEE/Description/AMOUNT，缺少即 ERROR)、格式 (金額/稅額數字、日期、選填欄位基本格式)、金額勾稽檢查擷取結果，給每張發票信心等級 (高/中/低)，產出異常清單 (Exceptions) 並把異常欄位逐筆寫入 log。
- `fallback_streamlit.py`：備援擷取。信心「低」時依序嘗試 OCR 重新擷取、LLM 結構化擷取 (Claude，需使用者勾選且設定 `ANTHROPIC_API_KEY`，預設關閉)。既有規則優先：單張發票只用 LLM 補缺值/異常欄位，其餘保留規則值；LLM 結果信心最高只到「中」。
- `logger_streamlit.py`：共用 log，寫入 `logs/invoice_processing.log` (已 gitignore)。

# 工作流程 (Workflow)
- 修改任一供應商規則後，必須執行回歸測試，確認該供應商既有樣本維持 100% 正確率：
  `python run_regression.py <SUPPLIER_KEY>`
- 新增/修正供應商依照 `add-invoice-supplier` skill 的標準流程 (含同步至本機執行資料夾)。
