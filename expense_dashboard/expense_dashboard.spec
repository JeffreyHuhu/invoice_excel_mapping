# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 打包設定：python -m PyInstaller --noconfirm expense_dashboard.spec
# 產生 dist\expense_dashboard.exe (單一檔案，內含 Python 與所有套件)。
from PyInstaller.utils.hooks import collect_data_files, collect_submodules, copy_metadata

# 頁面以 .py 檔案交給 Streamlit 執行，要放進 EXE 解壓縮後的根目錄
pages = [(f, '.') for f in ('app.py', 'dashboard.py', 'data_admin.py')]

datas = (
    pages
    + collect_data_files('streamlit')          # 前端網頁檔
    + copy_metadata('streamlit')               # streamlit 啟動時會查自己的版本
    + collect_data_files('plotly')             # plotly 的 validators / 樣板 json
)
hiddenimports = collect_submodules('streamlit')

a = Analysis(
    ['launcher.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # 用不到的大型套件不要打包進來，縮小 EXE
    excludes=['tkinter', 'matplotlib', 'IPython', 'jupyter', 'notebook', 'scipy',
              'torch', 'tensorflow', 'PIL.ImageQt', 'PyQt5', 'PySide6', 'pytest'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='expense_dashboard',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,   # 保留黑色視窗：顯示啟動狀態，關掉它就結束程式
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
