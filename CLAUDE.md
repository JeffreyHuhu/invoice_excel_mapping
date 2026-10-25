# 專案目標 (Project Goals)
- 目標: 自動化解析各式 supplier請款單據, 並準確對應至 Excel 欄位。
- 手段: 提供高效、穩定的批次處理流程。要優先執行local端程式, 不可大量讀取檔案(超過10份檔案), 如果超過請尋求授權
- 目的: 將解析後的結果, 做各種物流項目分析，供高階主管查看。

使用最有效率，運算最少，消耗最少TOKEN的方式執行專案
請都用繁體中文回答，並盡量問我問題, 以便你更好達成目標


# 最高指導原則 (Guiding Principles)
1. **效能優先**：處理大量 PDF 時，必須優先考慮記憶體與速度，善用批次 processing，嚴禁將大量檔案內容直接讀入 LLM 上下文。
2. **程式碼品質**：採用模組化設計，遵循 PEP 8 規範，並加上清晰的繁體中文註解。
3. **錯誤處理**：若遇到格式不符或損壞的檔案，必須妥善捕捉 exception，紀錄到 log 檔中，不得中斷整體流程。
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

# 同步規則 (Sync Rules)
改完 `*_streamlit.py` 或 `D:\資誠local執行\local_*.py` 後，一律用 `invoice-release` skill。不要手動逐步做。skill 負責回歸測試、同步、打包 EXE、commit + push。

必須記住的重點：
- 三個位置：GitHub `main`、`D:\資誠local執行\`、`D:\資誠streamlit程式碼\`。兩個 D: 資料夾都要更新。少一個，就會「雲端正常，本機全部 N/A」。
- 本機版檔名不帶 `_streamlit` (`extract_utils.py`、`suppliers.py`、`validation.py`、`logger.py`)。`import xxx_streamlit` 改成 `import xxx`。本機版沒有用到 fallback，所以不需要 `fallback.py`。
- 本機版檔名不可加進 git。repo 是公開的，不可 commit 任何客戶帳單資料。
- `D:\資誠local執行\` 是基準。先寫入並驗證它，再寫入另一邊。兩邊不一致時，以它為準。不確定時，先問我。
- 只覆蓋 `.py` 不會更新 `.exe`。要重新打包 `pdf_to_excel.exe`、`local_batch_process.exe`。改到儀表板時，也要打包 `expense_dashboard.exe`。PyInstaller 一律用 onedir，不可用 onefile。
- 只改本機專用腳本 (`local_*.py`) 時，只同步 `D:\資誠local執行\`。
- 收尾時，要明確告訴我「兩個資料夾都已更新」。
- git commit 訊息用英文。回覆使用者一律用繁體中文。

# 輸出 Excel 規則 (Output Excel Rules)
這些規則適用所有「轉檔結果」Excel。包含雲端版 (`build_excel()`、`verify_app.py` 的 `_build_result_excel()`) 和本機版 (`D:\資誠local執行\local_pdf_to_excel_gui.py`、`local_batch_process_gui.py` 的 `write_result_to_excel()`)。

1. 工作表名稱固定是 `"工作表"`。不可用供應商名稱，也不可留下範本的名稱 (曾經全部變成 `DWIHARTA`)。例外：內部核對報告 `build_verification_report_excel()` 有兩個分頁，名稱不改。
2. 日期欄位 (`invoice_date`、`arrive_date`、`onboard_date`) 要寫成真正的 Excel 日期。顯示格式是 `mmm-dd-yyyy`，例如 `Aug-10-2026`。
3. 日期轉換只在「寫入 Excel」時做，用 `_coerce_to_date()`。不要改各供應商的 `parse_xxx()`，也不要改 `normalize_value()`、`compare_rows()`。轉不出日期 (例如 `N/A`) 時，保留原值。
4. 非英文的擷取文字，寫入時用 `english_note()` 加上「(English translation)」。這條規則適用所有帳單 (PDF 帳單和 Excel 帳單)，所有國家、所有供應商。欄位見 `TEXT_NOTE_FIELDS`。不要改 `parse_xxx()`。新的公司名稱加進 `_EN_FULL`，新的詞組加進 `_EN_PHRASES`。已經有英文備註的文字不重複加。例外：`verify_app.py` 的 `_build_result_excel()` 不加英文備註。這份匯出檔會直接當正確答案 Excel，而比對程式不會去除括號備註。
5. 幣別依供應商國家：越南 VND，印尼 IDR。新國家要加 `COUNTRY_CURRENCY` 規則。
6. 驗證方式：用 `openpyxl` 讀回輸出檔，檢查 `wb.active.title == "工作表"`，以及日期欄位的 `number_format` 和型別。

# 新增或修正供應商 (Supplier Rules)

## 標準步驟
1. 前置分析：逐頁看 PDF 原文。確認是否混有多種單據、是掃描檔還是可編輯 PDF、是一份一張發票還是多張發票。
2. 寫 `detect_xxx()`：用該供應商獨有的關鍵字。不要用太短的字。曾經 `"MAERSK"` 誤判船名 `JOSEPHINE MAERSK`。同一供應商有多種版型時，各寫一個 detect/parse，再用外層函式匯整。
3. 寫 `parse_xxx()`：回傳 `{"header": {...}, "items": [...]}`。抓不到的欄位填 `None`。
4. 在 `suppliers_streamlit.py` 最後用 `register_supplier()` 註冊。寬鬆的判斷式放後面。
5. 跑回歸測試到 100%：`python run_regression.py <SUPPLIER_KEY>`。同一供應商的所有樣本都要跑。
6. 改到共用邏輯 (`extract_utils_streamlit.py`、寬鬆的 detect、`normalize_value()`) 時，跑全部 `test_samples/*/`。
7. 依「同步規則」用 `invoice-release` skill 發佈。

## 沒有達到 100% 時
- 先判斷：是程式的問題，還是正確答案 Excel 的問題。
- 程式問題：修 `detect_xxx`/`parse_xxx`。
- 正確答案有錯：不要為了配合錯誤答案改程式。要列出「哪一列、哪一欄、PDF 寫什麼、Excel 寫什麼」，問我確認，再更新 `sampleN_ref.xlsx`。
- 「雲端正常，本機全部 N/A」：先確認兩個 D: 資料夾的 `suppliers.py` 是否缺新函式。

## 正確答案 Excel 的提示列
- 表頭上方可能有「提示列」，例如 `判斷帳單類型 key word`、`擷取key word`。這些只用來寫規則。
- 提示列不放進輸出檔，也不放進 `test_samples/*/sampleN_ref.xlsx`。`read_reference_excel()` 會自動跳過它們。

## 附件頁 (雜訊頁)
- 所有供應商都只擷取帳單頁。附件頁要列出，但不擷取。
- 用 `register_supplier(invoice_page=)` 指定帳單頁，用 `invoice_text()` 取帳單頁文字。擷取與勾稽都用它。
- 查驗結果列出「第 x~y 頁是附件，未擷取」。等級是 INFO，不扣分。遇到新的雜訊樣本，就幫該供應商補 `invoice_page`。

## Faktur Pajak
- Faktur Pajak 是印尼稅務發票，不是供應商帳單。這條規則只適用印尼。其他國家還沒有稅務發票樣本。
- 出現這個抬頭的頁面要跳過。不可從中擷取資料，也不可影響其他頁面。

## Excel 帳單
- 擷取和稽核前，先檢查原帳單的公式錯誤：SUM 範圍、合計列 VAT 是否等於「未稅 × 稅率」、合計是否等於明細加總。
- `excel_formula_issues()` 做這個檢查 (查驗結果第 1 節)。合計對不上時，先找原帳單公式的原因。新的錯誤類型加進這個檢查，再考慮改擷取規則。

## VAT 與總額稽核
先判斷這家供應商的 VAT 屬於哪一種：
1. 每個品項自帶 VAT 欄：把 `vat` 填在該品項。不要另外建假品項。VAT 通常有小數點，Qty/Rate/Amount 通常是整數，可用小數點區分。
2. 整張帳單只有一列稅金：建一個獨立品項，`amount` 為 `None`，`vat` 填金額。
3. 完全沒有稅金：`vat` 全部留 `None`，不傳 `extract_grand_total`。

稽核參數 (都是選填)：`extract_total`、`extract_grand_total`、`extract_final_total` + `deduction_prefixes`。
- 帳單自己的小計因四捨五入，和品項加總差超過 ±1 時，不要實作稽核參數 (例如 TATA A/F)。
- 有「非課稅但仍計入未稅總額」的費用時，`extract_total` 要把它加進去 (例如 PUTRA_SEJAHTERA_SENTOSA)。
- 某些月份才有的費用項目，只在金額大於 0 時才加入。

## 異常等級
- 金額稽核沒通過，一律是 ERROR。原因在原帳單時也不降級，只加上原因說明。
- 不可把 `TOTAL_*_MISMATCH` 或公式錯誤放進 `KNOWN_SOURCE_ISSUES` 或 `SUPPLIER_EXEMPTIONS`。通過金額稽核的項目不標錯誤。
- 使用者確認過的原帳單問題，放進 `validation_streamlit.py` 的 `KNOWN_SOURCE_ISSUES`。等級是 INFO，不是 WARN，也不標紅，不扣信心分。這只適用不影響金額的欄位。
- `KNOWN_SOURCE_ISSUES` 的限定原始值要盡量窄 (例如 ETA 只在值為 `--` 時才算)。也要在 `suppliers_streamlit.py` 該供應商區塊加註解，提醒不要改擷取規則。
- 使用者說某項目對某供應商「可以例外」時，登記到 `SUPPLIER_EXEMPTIONS`。這些項目完全不列出。

## 「獨立一列稅金」的 Description
照抄 PDF 上的標籤，百分比用 regex 動態抓出。每家格式不同，不要套用別家的格式：
- TRANS_DAYA_PRIMA：`VAT (11%)`
- APL：固定 `VAT`
- PUTRA_SEJAHTERA_SENTOSA：`PPN 1,1 %` 和 `PPH 2%`

處理新供應商時，先看使用者提供的正確答案寫了什麼，再補進這份清單。

## 其他經驗
- 描述文字預設照抄 PDF 原文。不要自行清理空白或格式。
- 日期比對：正確答案的日期可能是文字，也可能是真正的日期。`normalize_value()` 兩種都支援。`parse_xxx()` 固定輸出一種即可。
- 多品項供應商：不要因為第一份樣本每個品項的值都一樣 (甚至都空白)，就認定這個欄位整張發票共用。要回頭看 PDF 原文。如果每個品項各自有值，就在該品項 dict 加上對應的標準欄位 key (例如 `onboard_date`、`bl_no`、`port_of_loading`)。收到同供應商的多張發票時，要逐項再核對一次。
- Regex 陷阱：可能為空的群組，前後不要同時要求 `\s+`。否則整行會比對失敗，品項會悄悄少一筆。品項數量比預期少，又沒有錯誤時，先查這裡。
- 雙欄換行：pdfplumber 會把左右兩欄合成一行。一行可能同時有兩個欄位的殘字。要先拆成「屬於哪一欄的部分」，再分別處理。已知值清單 (例如國家名稱) 要隨新樣本擴充。

請用 ASD-STE100 的原則，以中文或英文回答問題。
用中文回答時，遵守以下規則：
1. 句子要短，每句盡量不超過 30 個字。
2. 一句只講一件事。
3. 同一個概念從頭到尾用同一個詞，不要換說法。
4. 用主動語態，每句的主詞要清楚。
5. 不要堆疊名詞，一個名詞片語最多三個詞。


