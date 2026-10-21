# -*- coding: utf-8 -*-
"""
launcher.py — 打包成 expense_dashboard.exe 的進入點

雙擊 EXE：在這台電腦啟動儀表板 (只接受本機連線 127.0.0.1，不對外開放)，
等伺服器就緒後自動用預設瀏覽器打開。黑色視窗是伺服器本身，關掉它就結束
程式。資料庫位置見 settings.py (EXE 旁邊的 data\\expenses.duckdb，或在
「資料管理」頁指定的路徑，例如公司共用資料夾)。
"""

import multiprocessing
import socket
import sys
import threading
import time
import urllib.request
import webbrowser
from pathlib import Path

# 讓 PyInstaller 把頁面用到的模組 (與它們的套件) 一起打包；頁面本身是以
# .py 檔案的形式交給 Streamlit 執行
import dashboard_deps  # noqa: F401


def _free_port(start: int = 8501) -> int:
    for port in range(start, start + 50):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port
    raise RuntimeError("找不到可用的連接埠 (8501~8550 都被占用)")


def _open_browser_when_ready(port: int) -> None:
    url = f"http://localhost:{port}"
    # 新的 EXE 第一次執行時防毒軟體會先掃描，可能要 1~2 分鐘，最多等 5 分鐘
    for _ in range(600):
        try:
            with urllib.request.urlopen(f"{url}/_stcore/health", timeout=1) as resp:
                if resp.read() == b"ok":
                    webbrowser.open(url)
                    print(f"\n儀表板已開啟：{url}\n(關閉這個視窗即結束程式)\n")
                    return
        except OSError:
            pass
        time.sleep(0.5)
    print(f"\n瀏覽器沒有自動打開，請手動開啟 {url}\n")


def main() -> int:
    multiprocessing.freeze_support()
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    port = _free_port()
    print("費用分析儀表板啟動中，請稍候 (第一次執行可能要 1~2 分鐘)...")
    threading.Thread(target=_open_browser_when_ready, args=(port,), daemon=True).start()

    from streamlit.web import cli as stcli
    sys.argv = [
        "streamlit", "run", str(base / "app.py"),
        f"--server.port={port}", "--server.address=127.0.0.1",
        "--server.headless=true", "--global.developmentMode=false",
        "--browser.gatherUsageStats=false", "--server.fileWatcherType=none",
        "--client.toolbarMode=minimal",
    ]
    return stcli.main()


if __name__ == "__main__":
    sys.exit(main())
