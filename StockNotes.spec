from PyInstaller.utils.hooks import collect_all


akshare_data, akshare_binaries, akshare_hiddenimports = collect_all("akshare")

analysis = Analysis(
    ["app.py"],
    pathex=[],
    binaries=akshare_binaries,
    datas=[
        ("templates", "templates"),
        ("static", "static"),
        *akshare_data,
    ],
    hiddenimports=akshare_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pytest", "IPython", "notebook", "jupyter"],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(analysis.pure)

exe = EXE(
    pyz,
    analysis.scripts,
    [],
    exclude_binaries=True,
    name="StockNotes",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
)

collect = COLLECT(
    exe,
    analysis.binaries,
    analysis.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="StockNotes",
)
