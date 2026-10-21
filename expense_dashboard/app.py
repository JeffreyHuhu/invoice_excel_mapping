# -*- coding: utf-8 -*-
"""
app.py — 費用分析系統入口 (streamlit run app.py，或執行打包好的 EXE)

兩個頁面：
  dashboard.py   費用分析儀表板 (篩選、KPI、圖表、明細匯出)
  data_admin.py  資料管理 (資料庫位置、從轉檔結果更新資料、匯率維護)
"""

from pathlib import Path

import streamlit as st

HERE = Path(__file__).resolve().parent

st.set_page_config(page_title="費用分析儀表板", page_icon="📊", layout="wide")
page = st.navigation([
    st.Page(HERE / "dashboard.py", title="費用分析儀表板", icon="📊", default=True),
    st.Page(HERE / "data_admin.py", title="資料管理", icon="🗂️"),
])
page.run()
