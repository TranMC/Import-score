# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['import_score.py'],
    pathex=[],
    binaries=[],
    datas=[('e:\\Open_source_project\\Nhập điểm\\app_config.json', '.')],
    hiddenimports=['pandas._libs.tslibs.base', 'pandas._libs.tslibs.np_datetime', 'pandas._libs.tslibs.timedeltas'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'scipy', 'PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'PIL', 'IPython', 'notebook', 'jedi', 'jupyter', 'zmq', 'cryptography', 'webbrowser', 'xml.dom.domreg', 'pycparser', 'sqlite3'],
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
    name='Quản lý điểm học sinh v4',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['app.ico'],
)
