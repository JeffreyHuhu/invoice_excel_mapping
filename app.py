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

按鈕：
  🗑️ 清除索引：清空目前的比對結果與快取，方便下一次重新搜尋/上傳
  ⬇️ 下載核對報告：把正確率總覽 + 逐欄比對明細 (紅綠燈) 匯出成 Excel

結果呈現順序：① 轉檔正確率 → ② 轉檔結果預覽 → ③ 逐欄比對明細

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
st.caption(
    "流程：① 自動辨識帳單屬於哪家供應商 → ② 套用該供應商專屬的擷取規則 → "
    f"③ 跟正確答案 Excel 逐欄比對算出正確率 (正確率未滿 100% 時，最多重新嘗試 "
    f"{MAX_EXTRACTION_ATTEMPTS} 次不同的擷取參數，一達到 100% 就停止)。"
)

registered = list_registered_suppliers()
with st.expander(f"🏷️ 目前系統已支援 {len(registered)} 家供應商的辨識規則"):
    for s in registered:
        st.write(f"- **{s['label']}** (代碼: `{s['key']}`)")
    st.caption("要新增供應商，請在 suppliers.py 增加一組辨識/擷取規則並註冊，不用改這個網頁程式。")

col_pdf, col_ref = st.columns(2)
with col_pdf:
    uploaded_pdfs = st.file_uploader(
        "① 上傳 PDF 帳單 (可一次選取多個檔案、可混合不同供應商)",
        type=["pdf"], accept_multiple_files=True, key="pdf_upload",
    )
with col_ref:
    reference_excel = st.file_uploader(
        "② (選填) 上傳正確答案 Excel，用來自動比對並算出正確率",
        type=["xlsx"], accept_multiple_files=False, key="reference_upload",
    )
    st.caption("沒有上傳的話，會改用『資料完整度』(欄位是否成功抓到值，而非 N/A) 作為代理指標。")

if not uploaded_pdfs:
    st.info("請先上傳至少一份 PDF 帳單。")
    st.stop()

clear_index_clicked = st.button(
    "🗑️ 清除索引 (清空目前的比對結果與快取，方便下一次重新搜尋/上傳)",
    help="更新過 suppliers.py 的擷取規則、換了一批檔案、或單純想清空重來時可以按這個按鈕。",
)
if clear_index_clicked:
    st.session_state.pop("result_bundle", None)
    st.session_state.pop("last_signature", None)
    st.rerun()

# ---------------------------------------------------------------------------
# 用檔案內容算出簽章，判斷「這批檔案是不是已經處理過」；簽章不同(新上傳
# 檔案或換了正確答案 Excel) 時，才重新整個跑一次擷取流程，避免 Streamlit
# 每次互動 (例如編輯表格) 都重新解析一次 PDF。按「清除索引」後快取會被
# 清空，下一次執行時一定會重新跑一次。
# ---------------------------------------------------------------------------

def _files_signature(files):
    return tuple((f.name, hashlib.md5(f.getvalue()).hexdigest()) for f in files) if files else None


current_signature = (
    _files_signature(uploaded_pdfs),
    hashlib.md5(reference_excel.getvalue()).hexdigest() if reference_excel else None,
)
need_recompute = (
    st.session_state.get("last_signature") != current_signature
    or "result_bundle" not in st.session_state
)

if need_recompute:
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

                reference_rows = None
                if has_reference and "invoice_no" in reference_df.columns:
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

                score_rows.append({
                    "來源檔案": f.name,
                    "辨識供應商": label,
                    "Invoice No": result["rows"][0].get("invoice_no", NA),
                    score_label: f"{result['score'] * 100:.1f}%",
                    "嘗試次數": f"{result['attempts_used']}/{MAX_EXTRACTION_ATTEMPTS}",
                    "是否達到100%": "✅ 是" if result["reached_100"] else "❌ 否",
                    "採用設定": result["config"] or "預設",
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
# ① 轉檔正確率 (%)
# ---------------------------------------------------------------------------

st.subheader(f"📊 轉檔{score_label}")

if not has_reference:
    st.info(
        "目前沒有上傳正確答案 Excel，以下是『資料完整度』：欄位有成功抓到值 "
        "(不是 N/A) 的比例，不是跟人工核對過的正確率；上傳正確答案 Excel "
        "後即可看到真正的逐欄比對正確率，並啟用重試迴圈與核對報告。"
    )

score_df = pd.DataFrame(score_rows)
st.dataframe(score_df, use_container_width=True)

avg_score = score_df[score_label].str.rstrip("%").astype(float).mean()
all_reached_100 = has_reference and all(r["是否達到100%"] == "✅ 是" for r in score_rows)
not_reached_files = [r["來源檔案"] for r in score_rows if r.get("是否達到100%") == "❌ 否"]

if has_reference and all_reached_100:
    st.success(f"🎉 全部 {len(score_rows)} 份帳單都在重試次數內達到 100% 正確率！")
elif has_reference and not_reached_files:
    st.warning(
        f"⚠️ 以下 {len(not_reached_files)} 份帳單嘗試了 {MAX_EXTRACTION_ATTEMPTS} 次"
        f"仍未達到 100% 正確率，已產出核對報告供人工複查不相符的欄位：\n\n"
        + "\n".join(f"- {name}" for name in not_reached_files)
    )
    st.metric(f"整體{score_label} (所有 PDF 平均)", f"{avg_score:.1f}%")
else:
    st.metric(f"整體{score_label} (所有 PDF 平均)", f"{avg_score:.1f}%")

with st.expander(
    f"🔍 查看每份 PDF 重複嘗試 (最多 {MAX_EXTRACTION_ATTEMPTS} 次) 擷取設定時的分數"
):
    for log in attempts_log:
        st.write(f"**{log['檔案']}** (供應商: {log['供應商']})")
        st.dataframe(pd.DataFrame(log["嘗試紀錄"]), use_container_width=True)

st.divider()

# ---------------------------------------------------------------------------
# ② 轉檔結果預覽
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

st.divider()

# ---------------------------------------------------------------------------
# ③ 逐欄比對明細
# ---------------------------------------------------------------------------

if has_reference and compare_rows_all:
    st.subheader("📋 逐欄比對明細（🟢相符 / 🔴不相符）")
    cmp_df = pd.DataFrame(compare_rows_all)

    def _highlight(row):
        color = "background-color:#C6EFCE" if row["結果"].startswith("✅") else "background-color:#FFC7CE"
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
