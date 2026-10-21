# -*- coding: utf-8 -*-
"""
dashboard_deps.py — 只給 PyInstaller 分析用：列出儀表板頁面 (以 .py 檔案形式
交給 Streamlit 執行，PyInstaller 看不到) 會 import 的模組，讓它們連同依賴的
套件一起打包進 EXE。
"""

import duckdb  # noqa: F401
import openpyxl  # noqa: F401
import pandas  # noqa: F401
import plotly.graph_objects  # noqa: F401
import python_calamine  # noqa: F401

import expense_normalize  # noqa: F401
import fx  # noqa: F401
import process_data  # noqa: F401
import settings  # noqa: F401
