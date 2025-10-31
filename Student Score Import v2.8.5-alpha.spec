# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all
from PyInstaller.utils.hooks import copy_metadata

datas = [('e:\\Open_source_project\\Nhập điểm\\app_config.json', '.'), ('e:\\Open_source_project\\Nhập điểm\\version.json', '.'), ('e:\\Open_source_project\\Nhập điểm\\changelog.json', '.')]
binaries = []
hiddenimports = ['pandas._libs.tslibs.base', 'pandas._libs.tslibs.np_datetime', 'pandas._libs.tslibs.timedeltas', 'matplotlib.backends.backend_tkagg', 'matplotlib.figure', 'matplotlib.backends.backend_pdf', 'PIL._tkinter_finder', 'themes', 'ui_utils', 'version_utils', 'check_for_updates', 'openpyxl.cell']
datas += copy_metadata('numpy')
tmp_ret = collect_all('numpy')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['import_score.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pkg_resources', 'pytest', 'setuptools', 'pip', 'torch', 'tensorflow', 'scipy', 'scipy.stats', 'scipy.spatial', 'scipy.special', 'scipy.ndimage', 'scipy.linalg', 'jax', 'IPython', 'notebook', 'jupyter', 'tkinter.test', 'unittest', 'test', 'lib2to3'],
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
    name='Student Score Import v2.8.5-alpha',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
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
