# -*- coding: utf-8 -*-
"""
app.py
======
多供應商帳單自動化系統 (Streamlit 網頁)

流程：辨識哪間供應商帳單 → 選取對應的帳單辨識系統 → 擷取資料 → 跟正確答案
       Excel 逐欄比對算出正確率。

重試機制：每份 PDF 最多重新嘗試 20 次不同的擷取參數，只要正確率還沒到
100% 就繼續換下一組參數再試，一達到 100% 立刻停止；20 次都試完仍未達
100% 的話，就用這 20 次裡分數最高的一次，並產出核對報告方便人工複查。

學習記憶：每次找到「比之前記錄更高分」的擷取設定時，會記錄到
learned_configs.json，下次同一家供應商的帳單進來就優先套用這組設定當
第 1 次嘗試，通常能一次就命中，不用每次都從頭試 20 組 (注意：這個記憶
只在同一次 Streamlit 部署期間有效，重新部署/reboot 後會重置，除非把
learned_configs.json 也提交回 GitHub)。

按鈕：
  ▶️ 開始執行比對：上傳檔案後不會自動跑，一定要按這個按鈕才會擷取＋比對
                    並顯示結果 (換了檔案但還沒按這個按鈕時，畫面上會提醒)
  🔄 重新查詢：清空目前的比對結果、快取，並清除已上傳的檔案，方便下一次重新查詢
  ⬇️ 下載核對報告：把正確率總覽 + 逐欄比對明細 (紅綠燈/灰) 匯出成 Excel

結果呈現順序：① 轉檔正確率(大字級顯示) → ② 逐欄比對明細 → ③ 轉檔結果預覽

特殊供應商：HYPER MEGA 一份 PDF 裡可能塞了好幾張各自獨立的發票 (一頁一
張，各自的發票號碼/日期/費用明細都不同)，跟其他供應商「不管幾頁都是同
一張發票」不一樣；這種 multi_page 供應商的比對是用 invoice_no 分組後各
自跟正確答案表格比對，詳見 extract_utils.py / suppliers.py 裡的說明。

用法：
    pip install streamlit pdfplumber openpyxl pandas
    streamlit run app.py

擴充：要支援第 21 家供應商，去 suppliers.py 加一組 detect_xxx()/parse_xxx()
並呼叫 register_supplier() 即可，這支 app.py 完全不用改。
"""

import datetime
import hashlib
import tempfile
import os

import streamlit as st

# ---------------------------------------------------------------------------
# 防禦性匯入：Streamlit Cloud 在「未捕捉的例外」發生時，會把真正的錯誤訊息
# 隱藏成一句通用的 "original error message is redacted"，讓人完全看不出
# 到底是哪個套件沒裝好。這裡改成自己 try/except 去匯入，一旦失敗就用
# st.error() 印出明確、不會被隱藏的訊息，並提示排解步驟。
# 最常見原因：requirements.txt 是後來才加上/修改的，但 Streamlit Cloud
# 沒有重新完整安裝套件 —— 到 App 右下角「Manage app」選單點
# 「Reboot app」(或「Clear cache」再重新部署) 就能強制重建環境解決。
# ---------------------------------------------------------------------------
try:
    import pandas as pd

    import suppliers  # noqa: F401  (只是為了觸發所有供應商的 register_supplier())
    from extract_utils import (
        FIELD_CODES,
        DISPLAY_HEADERS,
        NA,
        MAX_EXTRACTION_ATTEMPTS,
        find_best_extraction,
        read_reference_excel,
        build_excel,
        build_verification_report_excel,
        normalize_value,
        detect_supplier,
        extract_text_from_pdf,
        parse_invoice_pdf,
        supplier_label,
        list_registered_suppliers,
        is_multi_page_supplier,
    )
except ImportError as e:
    st.set_page_config(page_title="多供應商帳單自動化系統 - 啟動失敗", layout="wide")
    st.error(
        f"❌ 系統啟動失敗，缺少必要的 Python 套件：`{e}`\n\n"
        "這通常代表 Streamlit Cloud 沒有正確安裝 `requirements.txt` 裡列出的套件"
        "(常發生在 requirements.txt 是後來才新增/修改，但環境沒有重新完整建置時)。"
        "請依序嘗試：\n\n"
        "1. 確認 repo 根目錄的 `requirements.txt` 存在，且內容包含 "
        "`streamlit`、`pdfplumber`、`openpyxl`、`pandas`\n"
        "2. 到這個 App 頁面右下角點 **Manage app** → 選單裡點 **Reboot app**，"
        "強制重新安裝所有套件\n"
        "3. 若還是失敗，同樣在 **Manage app** 裡可以看到完整安裝 log，"
        "確認實際是哪個套件安裝失敗"
    )
    st.stop()

st.set_page_config(page_title="多供應商帳單自動化系統", layout="wide")
st.title("📄➡️📊 多供應商帳單自動化系統")
st.markdown(
    f"""
    <div style="font-size:19px; color:#000000; font-weight:500; line-height:1.6; margin-bottom:12px;">
        流程：① 自動辨識帳單屬於哪家供應商 → ② 套用該供應商專屬的擷取規則 →
        ③ 跟正確答案 Excel 逐欄比對算出正確率 (正確率未滿 100% 時，最多重新嘗試
        {MAX_EXTRACTION_ATTEMPTS} 次不同的擷取參數，一達到 100% 就停止)。
    </div>
    """,
    unsafe_allow_html=True,
)

registered = list_registered_suppliers()
with st.expander(f"🏷️ 目前系統已支援 {len(registered)} 家供應商的辨識規則"):
    for s in registered:
        st.write(f"- **{s['label']}** (代碼: `{s['key']}`)")
    st.caption("要新增供應商，請在 suppliers.py 增加一組辨識/擷取規則並註冊，不用改這個網頁程式。")

# ---------------------------------------------------------------------------
# 「重新查詢」按鈕：把檔案上傳元件的 key 綁定一個版本號，版本號加 1 之後
# Streamlit 會把它們視為全新的元件重新渲染，藉此讓已上傳的檔案一併被清掉
# (Streamlit 沒有直接清空 file_uploader 的 API，換 key 是官方建議的做法)。
# 同時清空比對結果快取，讓使用者可以直接開始下一次全新的查詢。
# ---------------------------------------------------------------------------
if "uploader_version" not in st.session_state:
    st.session_state["uploader_version"] = 0

st.markdown(
    """
    <style>
    div.stButton > button {
        font-size: 28px !important;
        font-weight: 800 !important;
        padding: 1.2em 1.8em !important;
        height: auto !important;
        border-radius: 12px !important;
    }
    /* 「重新查詢」按鈕塗成藍色，跟主要動作的「開始執行比對」區分開來。
       CSS 沒辦法直接用按鈕文字選取，改用「緊接在一個隱形標記元素後面的
       按鈕」這個常見手法：在按鈕前面放一個帶 id 的隱形標記，用 :has()
       選到「包含這個標記的區塊」，再用 + 選到它後面緊接著的區塊裡的按鈕。 */
    div:has(> #reset-btn-marker) + div button {
        background-color: #1565C0 !important;
        border-color: #1565C0 !important;
        color: #FFFFFF !important;
    }
    div:has(> #reset-btn-marker) + div button:hover {
        background-color: #0D47A1 !important;
        border-color: #0D47A1 !important;
        color: #FFFFFF !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# 「① 上傳 PDF 帳單」/「② 上傳正確答案 Excel」整個上傳區域放大兩倍：
# Streamlit 的 file_uploader 沒有內建大小參數，這裡直接針對它的內部元素
# (拖曳區塊、圖示、文字、按鈕) 把字體/圖示/內距都放大約 2 倍，讓整個
# 上傳介面 (包含虛線框範圍) 實際占用的空間也放大，而不是只放大文字。
st.markdown(
    """
    <style>
    /* 上傳元件標題文字 (「① 上傳 PDF 帳單...」/「② (選填) 上傳正確答案...」) */
    [data-testid="stFileUploader"] label p {
        font-size: 2em !important;
        font-weight: 700 !important;
    }
    /* 拖曳上傳的虛線框區塊本身：內距、最小高度都放大兩倍 */
    [data-testid="stFileUploaderDropzone"] {
        padding: 2.4em 2em !important;
        min-height: 220px !important;
    }
    /* 拖曳區塊裡的說明文字 ("Drag and drop file here"、檔案大小限制...) */
    [data-testid="stFileUploaderDropzoneInstructions"] span,
    [data-testid="stFileUploaderDropzoneInstructions"] small {
        font-size: 2em !important;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] small {
        font-size: 1.5em !important;
    }
    /* 雲朵上傳圖示放大兩倍 */
    [data-testid="stFileUploaderDropzoneInstructions"] svg {
        width: 2.4em !important;
        height: 2.4em !important;
    }
    /* 「Browse files」按鈕放大兩倍 */
    [data-testid="stFileUploaderDropzone"] button {
        font-size: 1.4em !important;
        padding: 0.8em 1.4em !important;
    }
    /* 已上傳檔案清單 (檔名、大小、刪除按鈕) 跟著放大，維持視覺比例一致 */
    [data-testid="stFileUploaderFile"] {
        font-size: 1.6em !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

uploader_key_suffix = st.session_state["uploader_version"]

col_pdf, col_ref = st.columns(2)
with col_pdf:
    uploaded_pdfs = st.file_uploader(
        "① 上傳 PDF 帳單 (可一次選取多個檔案、可混合不同供應商)",
        type=["pdf"], accept_multiple_files=True,
        key=f"pdf_upload_{uploader_key_suffix}",
    )
with col_ref:
    reference_excel = st.file_uploader(
        "② (選填) 上傳正確答案 Excel，用來自動比對並算出正確率",
        type=["xlsx"], accept_multiple_files=False,
        key=f"reference_upload_{uploader_key_suffix}",
    )
    st.caption("沒有上傳的話，會改用『資料完整度』(欄位是否成功抓到值，而非 N/A) 作為代理指標。")

if not uploaded_pdfs:
    st.info("請先上傳至少一份 PDF 帳單。")
    st.stop()

# ---------------------------------------------------------------------------
# 用檔案內容算出簽章，判斷「這批檔案是不是已經處理過」，用來提醒使用者
# 檔案換了但還沒重新執行 (擷取/比對本身一律等「執行比對」按鈕按下才會跑，
# 不會因為換檔案就自動觸發，符合『上傳完要按按鈕才會顯示結果』的需求)。
# ---------------------------------------------------------------------------

def _files_signature(files):
    return tuple((f.name, hashlib.md5(f.getvalue()).hexdigest()) for f in files) if files else None


current_signature = (
    _files_signature(uploaded_pdfs),
    hashlib.md5(reference_excel.getvalue()).hexdigest() if reference_excel else None,
)

st.write("")
col_run, col_reset = st.columns(2)
with col_run:
    run_clicked = st.button(
        "▶️ 開始執行比對",
        type="primary",
        use_container_width=True,
        help="上傳完 PDF (與選填的正確答案 Excel) 後，按這個按鈕才會開始擷取與比對，比對結果才會顯示在下方。",
    )
with col_reset:
    st.markdown('<span id="reset-btn-marker"></span>', unsafe_allow_html=True)
    reset_clicked = st.button(
        "🔄 重新查詢",
        use_container_width=True,
        help="清空比對結果、快取，並清除已上傳的 PDF / 正確答案 Excel，方便重新上傳新的一批檔案。",
    )
if reset_clicked:
    st.session_state.pop("result_bundle", None)
    st.session_state.pop("last_signature", None)
    st.session_state["uploader_version"] += 1
    st.rerun()

if run_clicked:
    reference_df = None
    if reference_excel is not None:
        try:
            reference_df = read_reference_excel(reference_excel.getvalue())
            if reference_df.empty:
                st.warning("⚠️ 正確答案 Excel 讀不到任何資料列，將改用資料完整度分數。")
                reference_df = None
        except Exception as e:  # noqa: BLE001
            st.warning(f"⚠️ 正確答案 Excel 讀取失敗 ({e})，將改用資料完整度分數。")

    has_reference = reference_df is not None
    score_label = "正確率" if has_reference else "資料完整度"

    all_export_rows, score_rows, compare_rows_all, attempts_log, unknown_files = [], [], [], [], []

    with st.spinner(
        f"解析中 (每份 PDF 先辨識供應商，正確率未滿 100% 時最多重試 "
        f"{MAX_EXTRACTION_ATTEMPTS} 次不同的擷取參數)..."
    ):
        for f in uploaded_pdfs:
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                tmp.write(f.getvalue())
                tmp_path = tmp.name
            try:
                quick_text = extract_text_from_pdf(tmp_path)
                supplier_key = detect_supplier(quick_text)
                if supplier_key is None:
                    unknown_files.append(f.name)

                is_multi = is_multi_page_supplier(supplier_key)

                reference_rows = None
                if has_reference and "invoice_no" in reference_df.columns:
                    if is_multi:
                        # multi_page 供應商 (例如 HYPER MEGA)：一份 PDF 裡
                        # 可能混著好幾張不同發票，沒辦法只靠「第一筆」的
                        # invoice_no 篩出對應答案；直接把完整的正確答案表
                        # 格傳進去，find_best_extraction() 內部會自己用
                        # invoice_no 分組、各自比對。
                        reference_rows = reference_df
                    else:
                        _, quick_rows, _ = parse_invoice_pdf(tmp_path, supplier_key=supplier_key)
                        inv_no = normalize_value(quick_rows[0].get("invoice_no"))
                        mask = reference_df["invoice_no"].map(normalize_value) == inv_no
                        candidate = reference_df[mask]
                        if not candidate.empty:
                            reference_rows = candidate

                result = find_best_extraction(
                    tmp_path, supplier_key=supplier_key, reference_rows=reference_rows
                )

                label = supplier_label(result["supplier_key"])
                for row in result["rows"]:
                    row["__supplier_label"] = label
                all_export_rows.extend(result["rows"])

                distinct_invoice_nos = sorted({
                    str(row.get("invoice_no", NA)) for row in result["rows"]
                })
                invoice_no_display = (
                    distinct_invoice_nos[0] if len(distinct_invoice_nos) == 1
                    else f"共 {len(distinct_invoice_nos)} 張發票: " + ", ".join(distinct_invoice_nos)
                )

                score_rows.append({
                    "來源檔案": f.name,
                    "辨識供應商": label,
                    "Invoice No": invoice_no_display,
                    score_label: f"{result['score'] * 100:.1f}%",
                    "嘗試次數": f"{result['attempts_used']}/{MAX_EXTRACTION_ATTEMPTS}",
                    "是否達到100%": "✅ 是" if result["reached_100"] else "❌ 否",
                    "採用設定": result["config"] or "預設",
                    "套用學習記憶": "✅ 是" if result.get("learned_applied") else "-",
                    "比對基準": "正確答案 Excel" if reference_rows is not None else "資料完整度自我檢查",
                })
                attempts_log.append({
                    "檔案": f.name, "供應商": label, "嘗試紀錄": result["attempts"],
                })

                if result["results"]:
                    for r in result["results"]:
                        compare_rows_all.append({"來源檔案": f.name, "供應商": label, **r})

            except Exception as e:  # noqa: BLE001
                st.error(f"❌ {f.name}：解析失敗 ({e})")
            finally:
                os.remove(tmp_path)

    st.session_state["result_bundle"] = {
        "has_reference": has_reference,
        "score_label": score_label,
        "all_export_rows": all_export_rows,
        "score_rows": score_rows,
        "compare_rows_all": compare_rows_all,
        "attempts_log": attempts_log,
        "unknown_files": unknown_files,
    }
    st.session_state["last_signature"] = current_signature

if "result_bundle" not in st.session_state:
    # 還沒按過「執行比對」：不顯示任何轉檔/比對結果，只顯示提示。
    st.info("⬆️ 上傳檔案後，請按上方「▶️ 開始執行比對」按鈕，比對結果才會顯示在這裡。")
    st.stop()

if st.session_state.get("last_signature") != current_signature:
    st.warning(
        "⚠️ 目前上傳的檔案跟上次按「執行比對」時不一樣了，以下仍是上一次的結果，"
        "請重新按上方「▶️ 開始執行比對」以取得最新結果。"
    )

bundle = st.session_state["result_bundle"]
has_reference = bundle["has_reference"]
score_label = bundle["score_label"]
all_export_rows = bundle["all_export_rows"]
score_rows = bundle["score_rows"]
compare_rows_all = bundle["compare_rows_all"]
attempts_log = bundle["attempts_log"]
unknown_files = bundle["unknown_files"]

if unknown_files:
    st.warning(
        "⚠️ 以下檔案無法辨識供應商 (未命中任何已註冊的規則)，欄位皆以 N/A 填入。"
        "請到 suppliers.py 新增這家供應商的辨識/擷取規則：\n\n"
        + "\n".join(f"- {name}" for name in unknown_files)
    )

if not all_export_rows:
    st.stop()

# ---------------------------------------------------------------------------
# ① 轉檔正確率 (%) —— 大字級數字放在結果的最前方
# ---------------------------------------------------------------------------

score_df = pd.DataFrame(score_rows)
avg_score = score_df[score_label].str.rstrip("%").astype(float).mean()
all_reached_100 = has_reference and all(r["是否達到100%"] == "✅ 是" for r in score_rows)
not_reached_files = [r["來源檔案"] for r in score_rows if r.get("是否達到100%") == "❌ 否"]

_big_color = "#1a7f37" if (not has_reference or all_reached_100) else "#c0392b"
st.markdown(
    f"""
    <div style="text-align:center; padding: 12px 0 4px 0;">
        <div style="font-size:96px; font-weight:800; line-height:1; color:{_big_color};">
            {avg_score:.1f}%
        </div>
        <div style="font-size:20px; color:#666; margin-top:4px;">
            整體{score_label}（所有 {len(score_rows)} 份 PDF 平均）
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.subheader(f"📊 轉檔{score_label}明細")

if not has_reference:
    st.info(
        "目前沒有上傳正確答案 Excel，以上是『資料完整度』：欄位有成功抓到值 "
        "(不是 N/A) 的比例，不是跟人工核對過的正確率；上傳正確答案 Excel "
        "後即可看到真正的逐欄比對正確率，並啟用重試迴圈與核對報告。"
    )

st.dataframe(score_df, use_container_width=True)

if has_reference and all_reached_100:
    st.success(f"🎉 全部 {len(score_rows)} 份帳單都在重試次數內達到 100% 正確率！")
elif has_reference and not_reached_files:
    st.warning(
        f"⚠️ 以下 {len(not_reached_files)} 份帳單嘗試了 {MAX_EXTRACTION_ATTEMPTS} 次"
        f"仍未達到 100% 正確率，已產出核對報告供人工複查不相符的欄位：\n\n"
        + "\n".join(f"- {name}" for name in not_reached_files)
    )

with st.expander(
    f"🔍 查看每份 PDF 重複嘗試 (最多 {MAX_EXTRACTION_ATTEMPTS} 次) 擷取設定時的分數"
):
    for log in attempts_log:
        st.write(f"**{log['檔案']}** (供應商: {log['供應商']})")
        st.dataframe(pd.DataFrame(log["嘗試紀錄"]), use_container_width=True)

st.divider()

# ---------------------------------------------------------------------------
# ② 逐欄比對明細
# ---------------------------------------------------------------------------

if has_reference and compare_rows_all:
    st.subheader("📋 逐欄比對明細（🟢相符 / 🔴不相符 / ⚪正確答案為N/A，不計入正確率）")
    cmp_df = pd.DataFrame(compare_rows_all)

    def _highlight(row):
        if row["結果"].startswith("✅"):
            color = "background-color:#C6EFCE"
        elif row["結果"].startswith("⚪"):
            color = "background-color:#E7E6E6"
        else:
            color = "background-color:#FFC7CE"
        return [color] * len(row)

    # 注意：st.dataframe 的 height 參數在部分 Streamlit 版本中，傳入 None
    # 或不合法的數值 (例如條件式算出負數/0) 會直接拋出
    # StreamlitInvalidHeightError，且錯誤訊息會被 Streamlit Cloud 隱藏成
    # 一句 "original error message is redacted"。這裡改成不指定 height，
    # 讓 Streamlit 自行依資料筆數決定高度，避免整支 App 崩潰。
    st.dataframe(
        cmp_df.style.apply(_highlight, axis=1),
        use_container_width=True,
    )

    report_bytes = build_verification_report_excel(score_rows, compare_rows_all)
    st.download_button(
        "⬇️ 下載核對報告 (正確率總覽 + 逐欄比對明細)",
        data=report_bytes,
        file_name=f"核對報告_{datetime.date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()

# ---------------------------------------------------------------------------
# ③ 轉檔結果預覽
# ---------------------------------------------------------------------------

preview_records = []
for row in all_export_rows:
    rec = {"供應商": row.get("__supplier_label", NA)}
    for c in FIELD_CODES:
        rec[DISPLAY_HEADERS[c].split("\n")[0]] = row.get(c, NA)
    preview_records.append(rec)
preview_df = pd.DataFrame(preview_records)

st.subheader("✏️ 轉檔結果預覽（可直接在表格中修正錯誤欄位；抓不到的欄位顯示 N/A）")
edited_df = st.data_editor(preview_df, num_rows="dynamic", use_container_width=True)

display_to_code = {DISPLAY_HEADERS[c].split("\n")[0]: c for c in FIELD_CODES}
export_rows = []
for _, r in edited_df.iterrows():
    row = {display_to_code[col]: r[col] for col in display_to_code}
    row["__supplier_label"] = r["供應商"]
    export_rows.append(row)

excel_bytes = build_excel(export_rows)
st.download_button(
    "⬇️ 下載 Excel (彙整所有供應商的轉檔結果)",
    data=excel_bytes,
    file_name=f"帳單轉檔結果_{datetime.date.today().isoformat()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
