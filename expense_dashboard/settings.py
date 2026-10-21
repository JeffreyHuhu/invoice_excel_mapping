# -*- coding: utf-8 -*-
"""
settings.py — 費用分析系統的路徑與設定檔

  APP_DIR    ：程式所在資料夾。打包成 EXE 時是 EXE 旁邊 (不是 PyInstaller
               解壓縮的暫存資料夾)，log、設定檔、預設資料庫都放這裡，關掉
               程式後還在。
  資料庫位置 ：依序採用 環境變數 EXPENSE_DB > 設定檔 dashboard_config.json
               > APP_DIR\\data\\expenses.duckdb。在儀表板「資料管理」頁改過的
               路徑會存進設定檔，下次開啟沿用 (例如指向公司共用資料夾)。
"""

import json
import os
import sys
from pathlib import Path

if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent

CONFIG_PATH = APP_DIR / "dashboard_config.json"
DEFAULT_DB = APP_DIR / "data" / "expenses.duckdb"
LOG_DIR = APP_DIR / "logs"


def load_config() -> dict:
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_config(**values) -> None:
    config = load_config()
    config.update(values)
    CONFIG_PATH.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")


def get_db_path() -> Path:
    if os.environ.get("EXPENSE_DB"):
        return Path(os.environ["EXPENSE_DB"])
    configured = load_config().get("db_path")
    return Path(configured) if configured else DEFAULT_DB


def set_db_path(path: Path) -> None:
    save_config(db_path=str(path))
