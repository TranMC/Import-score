import PyInstaller.__main__
import os
import sys

# Đọc phiên bản từ version.txt
try:
    with open("version.txt", "r", encoding="utf-8") as f:
        version = f.read().strip()
except:
    version = "4.0"  # Phiên bản mặc định nếu không đọc được

print(f"Building version: {version}")

# Separator tuỳ hệ điều hành
sep = ';' if os.name == 'nt' else ':'

# Tự động tìm đường dẫn file config
data_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_config.json")

# Tạo thư mục dist nếu chưa tồn tại
if not os.path.exists("dist"):
    os.makedirs("dist")

# Các options tối ưu để giảm kích thước
options = [
    'import_score.py',
    f'--name=Quản lý điểm học sinh v{version}',
    '--onefile',
    '--windowed',
    '--icon=app.ico',
    '--clean',
    '--exclude-module=matplotlib',
    '--exclude-module=scipy',
    '--exclude-module=PyQt5',
    '--exclude-module=PyQt6',
    '--exclude-module=PySide2',
    '--exclude-module=PySide6',
    '--exclude-module=PIL',
    '--exclude-module=IPython',
    '--exclude-module=notebook',
    '--exclude-module=jedi',
    '--exclude-module=jupyter',
    '--exclude-module=zmq',
    '--exclude-module=cryptography',
    '--exclude-module=webbrowser',
    '--exclude-module=xml.dom.domreg',
    '--exclude-module=pycparser',
    '--exclude-module=sqlite3',
    f'--add-data={data_file}{sep}.',  # ✅ sửa lỗi dấu phân tách ở đây
    '--upx-dir=upx',
    '--hidden-import=pandas._libs.tslibs.base',
    '--hidden-import=pandas._libs.tslibs.np_datetime',
    '--hidden-import=pandas._libs.tslibs.timedeltas',
]

# Chạy PyInstaller với các options tối ưu
print(f"Running PyInstaller with options: {options}")
PyInstaller.__main__.run(options)

print(f"Build completed. Check the dist folder for the executable.")
