import os
import sys
import subprocess
import py_compile
import json
from datetime import datetime

print("Quá trình build bắt đầu...")

# Cập nhật build_date trong version.json
try:
    version_json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "version.json")
    with open(version_json_path, 'r', encoding='utf-8') as f:
        version_data = json.load(f)
    
    # Cập nhật build_date thành ngày hiện tại
    current_date = datetime.now().strftime("%Y-%m-%d")
    version_data['build_date'] = current_date
    
    # Lưu lại file version.json với build_date mới
    with open(version_json_path, 'w', encoding='utf-8') as f:
        json.dump(version_data, f, indent=4, ensure_ascii=False)
    
    print(f"Đã cập nhật build_date thành {current_date}")
except Exception as e:
    print(f"Lỗi khi cập nhật build_date: {str(e)}")

# Đọc phiên bản từ version.txt hoặc version.json
try:
    # Thử đọc phiên bản từ version.json bằng version_utils
    import version_utils
    version_info = version_utils.load_version_info()
    version = version_info.get('version', 'unknown')
except:
    # Thử đọc từ version.txt (cách cũ)
    try:
        with open("version.txt", "r", encoding="utf-8") as f:
            version = f.read().strip()
    except:
        version = "4.0"  # Phiên bản mặc định nếu không đọc được

print(f"Building version: {version}")

# Kiểm tra Python bytecode trước khi build
print("Kiểm tra lỗi cú pháp trong mã nguồn...")
try:
    py_compile.compile("import_score.py", doraise=True)
    py_compile.compile("write_log.py", doraise=True)
    py_compile.compile("themes.py", doraise=True)
    py_compile.compile("ui_utils.py", doraise=True)
    py_compile.compile("version_utils.py", doraise=True)
    py_compile.compile("check_for_updates.py", doraise=True)
    print("Không tìm thấy lỗi cú pháp trong mã nguồn.")
except py_compile.PyCompileError as e:
    print(f"Lỗi cú pháp trong mã nguồn: {str(e)}")
    print("Vui lòng sửa lỗi trước khi tiếp tục.")
    sys.exit(1)

# Kiểm tra và cài đặt PyInstaller nếu chưa cài đặt
try:
    import PyInstaller.__main__
except ImportError:
    print("PyInstaller chưa được cài đặt. Đang cài đặt...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "PyInstaller>=5.13.0"])
        import PyInstaller.__main__
    except Exception as e:
        print(f"Không thể cài đặt PyInstaller: {str(e)}")
        sys.exit(1)

# Separator tuỳ hệ điều hành
sep = ';' if os.name == 'nt' else ':'

# Tự động tìm đường dẫn file config
current_dir = os.path.dirname(os.path.abspath(__file__))
app_config_file = os.path.join(current_dir, "app_config.json")
version_json_file = os.path.join(current_dir, "version.json")
changelog_json_file = os.path.join(current_dir, "changelog.json")

# Tạo thư mục dist nếu chưa tồn tại
if not os.path.exists("dist"):
    os.makedirs("dist")

# Các options tối ưu để giảm kích thước
options = [
    'import_score.py',
    f'--name=Student Score Import v{version}',
    '--onefile',
    '--windowed',
    '--icon=app.ico',
    '--clean',
    '--exclude-module=scipy',
    '--exclude-module=PyQt5',
    '--exclude-module=PyQt6',
    '--exclude-module=PySide2',
    '--exclude-module=PySide6',
    '--exclude-module=IPython',
    '--exclude-module=notebook',
    '--exclude-module=jedi',
    '--exclude-module=jupyter',
    '--exclude-module=zmq',
    '--exclude-module=webbrowser',
    '--exclude-module=xml.dom.domreg',
    '--exclude-module=pycparser',
    '--exclude-module=sqlite3',
    f'--add-data={app_config_file}{sep}.',
    f'--add-data={version_json_file}{sep}.',
    f'--add-data={changelog_json_file}{sep}.',
    '--upx-dir=upx',
    '--hidden-import=pandas._libs.tslibs.base',
    '--hidden-import=pandas._libs.tslibs.np_datetime',
    '--hidden-import=pandas._libs.tslibs.timedeltas',
    '--hidden-import=matplotlib',
    '--hidden-import=matplotlib.backends.backend_tkagg',
    '--hidden-import=matplotlib.figure',
    '--hidden-import=matplotlib.backends.backend_pdf',
    '--hidden-import=matplotlib.colors',
    '--hidden-import=matplotlib.pyplot',
    '--hidden-import=matplotlib.rcsetup',
    '--hidden-import=PIL',
    '--hidden-import=PIL._tkinter_finder',
    '--hidden-import=PIL.Image',
    '--hidden-import=cryptography',
    '--hidden-import=themes',
    '--hidden-import=ui_utils',
    '--hidden-import=version_utils',
    '--hidden-import=check_for_updates',
    '--hidden-import=openpyxl',
    '--hidden-import=openpyxl.cell',
]

# Chạy PyInstaller với các options tối ưu
print(f"Chạy PyInstaller với các tùy chọn: {options}")
try:
    PyInstaller.__main__.run(options)
    print(f"Build hoàn thành. Kiểm tra thư mục dist để xem file thực thi.")
except Exception as e:
    print(f"Lỗi khi build: {str(e)}")
    sys.exit(1)
