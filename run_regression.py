#!/usr/bin/env python3
"""供應商辨識規則的回歸測試工具。

用途：每次修改 suppliers_streamlit.py 之後，快速確認「這次改動的供應商」
過去所有已驗證樣本是否仍然維持 100% 正確率 (避免修好新樣本、卻改壞舊樣本)。

樣本存放慣例：test_samples/<SUPPLIER_KEY>/sampleN.pdf + sampleN_ref.xlsx
(N 從 1 開始)，每次新增/修正一個樣本，記得把 (PDF, 正確答案 Excel) 這一組
複製進對應資料夾，之後就會自動被這支工具納入回歸測試範圍。
Excel 帳單供應商 (越南) 的樣本是 sampleN.xlsx / sampleN.xls + sampleN_ref.xlsx。

用法：
    python3 run_regression.py                  # 測試全部供應商
    python3 run_regression.py DHL EXPRESS_MAXIMUM   # 只測試指定的供應商
"""
import sys
import glob
import os

import suppliers_streamlit  # noqa: F401  (觸發 register_supplier())
import extract_utils_streamlit as eu


def run_one(pdf_path: str, ref_path: str, expected_key: str):
    ref_rows = eu.read_reference_excel(ref_path)
    result = eu.find_best_extraction(pdf_path, supplier_key=expected_key, reference_rows=ref_rows)
    detected = eu.detect_supplier(eu.extract_text_from_pdf(pdf_path))
    return {
        "detected_ok": detected == expected_key,
        "detected": detected,
        "score": result["score"],
        "reached_100": result["reached_100"],
        "results": result["results"],
    }


def main():
    root = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_samples")
    requested = set(sys.argv[1:])

    supplier_dirs = sorted(
        d for d in os.listdir(root) if os.path.isdir(os.path.join(root, d))
    )
    if requested:
        supplier_dirs = [d for d in supplier_dirs if d in requested]

    overall_ok = True
    for key in supplier_dirs:
        sdir = os.path.join(root, key)
        pdfs = sorted(
            p for p in glob.glob(os.path.join(sdir, "sample*.*"))
            if p.lower().endswith((".pdf", ".xlsx", ".xls")) and not p.endswith("_ref.xlsx")
        )
        if not pdfs:
            continue
        print(f"\n=== {key} ({len(pdfs)} 個樣本) ===")
        for pdf_path in pdfs:
            ref_path = os.path.splitext(pdf_path)[0] + "_ref.xlsx"
            if not os.path.exists(ref_path):
                print(f"  [略過] 找不到對應的正確答案: {ref_path}")
                continue
            name = os.path.basename(pdf_path)
            try:
                r = run_one(pdf_path, ref_path, key)
            except Exception as e:  # noqa: BLE001
                overall_ok = False
                print(f"  ❌ {name}: 執行時發生錯誤 - {e}")
                continue
            ok = r["detected_ok"] and r["reached_100"]
            overall_ok = overall_ok and ok
            mark = "✅" if ok else "❌"
            print(f"  {mark} {name}: 辨識={r['detected']} 正確率={r['score']*100:.1f}%")
            if not ok:
                for row in r["results"]:
                    if row["結果"].startswith("❌"):
                        print(f"       - {row['欄位']}: 擷取={row['擷取結果']!r} 正確答案={row['正確答案']!r}")

    print("\n" + ("全部通過 ✅" if overall_ok else "有樣本沒通過，請往上檢查 ❌"))
    sys.exit(0 if overall_ok else 1)


if __name__ == "__main__":
    main()
