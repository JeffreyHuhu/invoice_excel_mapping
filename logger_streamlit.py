# -*- coding: utf-8 -*-
"""
logger_streamlit.py
===================
全系統共用的 log 檔設定：批次處理時遇到格式不符/損壞的 PDF、驗證失敗、
LLM 備援呼叫失敗等狀況，一律寫進 logs/invoice_processing.log，不中斷整體
流程 (CLAUDE.md 最高指導原則第 3 條)。

log 檔超過 1MB 自動輪替，最多保留 5 份舊檔，避免無限長大。
"""

import logging
import os
import sys
from logging.handlers import RotatingFileHandler

# 打包成 EXE (PyInstaller) 時 __file__ 在執行完就刪除的暫存資料夾裡，
# log 改寫在 EXE 旁邊的 logs 資料夾，關掉程式後才查得到
if getattr(sys, "frozen", False):
    _BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    _BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(_BASE_DIR, "logs")
LOG_PATH = os.path.join(LOG_DIR, "invoice_processing.log")

_LOGGER_NAME = "invoice_excel_mapping"


def get_logger() -> logging.Logger:
    """取得共用 logger。重複呼叫不會重複掛 handler (Streamlit 每次互動都會
    重新執行整支 app.py，不檢查的話同一筆 log 會被寫好幾次)。
    log 資料夾無法建立 (例如雲端唯讀檔案系統) 時，退回只輸出到主控台，
    不讓 log 設定本身中斷流程。
    """
    logger = logging.getLogger(_LOGGER_NAME)
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
        handler: logging.Handler = RotatingFileHandler(
            LOG_PATH, maxBytes=1_000_000, backupCount=5, encoding="utf-8"
        )
    except OSError:
        handler = logging.StreamHandler()
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    return logger
