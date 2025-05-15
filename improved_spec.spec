# -*- mode: python ; coding: utf-8 -*-
import sys
import os
from PyInstaller.utils.hooks import collect_data_files

# Lấy phiên bản từ file
with open("version.txt", "r", encoding="utf-8") as f:
    VERSION = f.read().strip()

block_cipher = None

# Normalize paths để tránh vấn đề encoding
current_path = os.path.abspath(os.path.dirname(SPECPATH))
app_config = os.path.join(current_path, "app_config.json")
version_json = os.path.join(current_path, "version.json")
changelog_json = os.path.join(current_path, "changelog.json")

# Cấu hình riêng cho windows
if sys.platform.startswith('win'):
    sep = ';'
else:
    sep = ':'

# Bổ sung tùy chọn runtime hooks
a = Analysis(
    ['import_score.py'],
    pathex=[current_path],
    binaries=[],
    datas=[
        (app_config, '.'),
        (version_json, '.'),
        (changelog_json, '.'),
    ],
    hiddenimports=[
        'pandas._libs.tslibs.base',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.timedeltas',
        'matplotlib',
        'matplotlib.backends.backend_tkagg',
        'matplotlib.figure',
        'matplotlib.backends.backend_pdf',
        'matplotlib.colors',
        'matplotlib.pyplot',
        'matplotlib.rcsetup',
        'PIL',
        'PIL._tkinter_finder',
        'PIL.Image',
        'cryptography',
        'themes',
        'ui_utils',
        'version_utils',
        'check_for_updates',
        'openpyxl',
        'openpyxl.cell'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'scipy',
        'PyQt5',
        'PyQt6',
        'PySide2',
        'PySide6',
        'IPython',
        'notebook',
        'jedi',
        'jupyter',
        'zmq',
        'webbrowser',
        'xml.dom.domreg',
        'pycparser',
        'sqlite3',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Thêm cấu hình để sửa lỗi load dll
extra_options = {
    # Chỉ định đường dẫn tuyệt đối đến python DLL để tránh lỗi
    'dll_excludes': ['libcrypto-*.dll', 'MSVCP140.dll'],
    'bundle_files': 1,
    'optimize': 2,
}

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name=f'Student Score Import v{VERSION}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=True,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app.ico',
) 