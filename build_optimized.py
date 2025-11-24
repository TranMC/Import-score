"""
Công cụ xây dựng tối ưu hóa cho ứng dụng Quản Lý Điểm Học Sinh
Hỗ trợ tối ưu kích thước, kiểm tra lỗi và quản lý phụ thuộc
"""

import os
import sys
import subprocess
import json
import py_compile
from datetime import datetime
import shutil
import platform

def log(message, level="INFO"):
    """In thông điệp ra console với màu sắc"""
    colors = {
        "INFO": "\033[94m",  # Xanh dương
        "SUCCESS": "\033[92m",  # Xanh lá
        "WARNING": "\033[93m",  # Vàng
        "ERROR": "\033[91m",  # Đỏ
        "END": "\033[0m"  # Reset màu
    }
    
    # Kiểm tra nếu đang chạy trên Windows và không trong môi trường hỗ trợ màu ANSI
    if platform.system() == "Windows" and not os.environ.get('WT_SESSION'):
        print(f"[{level}] {message}")
    else:
        print(f"{colors.get(level, '')}{message}{colors['END']}")

def check_dependencies():
    """Kiểm tra và cài đặt các phụ thuộc cần thiết"""
    log("Kiểm tra các phụ thuộc cần thiết...")
    
    required_packages = [
        "pandas>=2.0.0",
        "matplotlib>=3.7.0",
        "numpy>=1.24.0",
        "requests>=2.28.0",
        "PyInstaller>=5.13.0",
        "cryptography>=41.0.0",
        "Pillow>=10.0.0",
        "openpyxl>=3.1.0"
    ]
    
    # Cài đặt các phụ thuộc thiếu
    missing_packages = []
    for package in required_packages:
        package_name = package.split('>=')[0]
        try:
            __import__(package_name)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        log(f"Đang cài đặt {len(missing_packages)} gói thiếu: {', '.join(missing_packages)}", "INFO")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install"] + missing_packages)
            log("Cài đặt phụ thuộc thành công.", "SUCCESS")
        except Exception as e:
            log(f"Lỗi khi cài đặt phụ thuộc: {str(e)}", "ERROR")
            sys.exit(1)
    else:
        log("Tất cả phụ thuộc đã được cài đặt.", "SUCCESS")
    
    # Kiểm tra PyInstaller đặc biệt
    try:
        import PyInstaller.__main__
        log("PyInstaller đã sẵn sàng.", "SUCCESS")
    except ImportError:
        log("Không thể import PyInstaller sau khi cài đặt. Đang thử cài đặt lại...", "WARNING")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "PyInstaller>=5.13.0"])
            import PyInstaller.__main__
            log("Cài đặt lại PyInstaller thành công.", "SUCCESS")
        except Exception as e:
            log(f"Không thể cài đặt PyInstaller: {str(e)}", "ERROR")
            sys.exit(1)

def check_source_code():
    """Kiểm tra lỗi cú pháp trong mã nguồn"""
    log("Kiểm tra lỗi cú pháp trong mã nguồn...")
    
    source_files = [
        "import_score.py",
        "write_log.py",
        "themes.py",
        "ui_utils.py",
        "version_utils.py",
        "check_for_updates.py"  # Sử dụng file chính thay vì file đã sửa
    ]
    
    try:
        for file in source_files:
            if os.path.exists(file):
                py_compile.compile(file, doraise=True)
                log(f"✓ {file} không có lỗi cú pháp", "INFO")
            else:
                log(f"⚠️ File {file} không tồn tại, bỏ qua", "WARNING")
        
        log("Kiểm tra mã nguồn hoàn tất. Không phát hiện lỗi cú pháp.", "SUCCESS")
    except py_compile.PyCompileError as e:
        log(f"Lỗi cú pháp trong mã nguồn: {str(e)}", "ERROR")
        log("Vui lòng sửa lỗi trước khi tiếp tục.", "ERROR")
        sys.exit(1)

def update_version_info():
    """Cập nhật thông tin phiên bản và build date"""
    log("Cập nhật thông tin phiên bản...")
    
    # Đường dẫn đến file version.json
    version_json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "version.json")
    
    try:
        # Đọc thông tin phiên bản hiện tại
        with open(version_json_path, 'r', encoding='utf-8') as f:
            version_data = json.load(f)
        
        # Cập nhật build_date thành ngày hiện tại
        current_date = datetime.now().strftime("%Y-%m-%d")
        version_data['build_date'] = current_date
        
        # Lưu lại file version.json với build_date mới
        with open(version_json_path, 'w', encoding='utf-8') as f:
            json.dump(version_data, f, indent=4, ensure_ascii=False)
        
        log(f"Đã cập nhật build_date thành {current_date}", "SUCCESS")
        
        # Trả về phiên bản hiện tại
        return version_data.get("version", "unknown")
    except Exception as e:
        log(f"Lỗi khi cập nhật build_date: {str(e)}", "ERROR")
        # Thử đọc từ version.txt (cách cũ)
        try:
            with open("version.txt", "r", encoding="utf-8") as f:
                version = f.read().strip()
                log(f"Đọc phiên bản từ version.txt: {version}", "WARNING")
                return version
        except:
            log("Không thể đọc thông tin phiên bản, sử dụng giá trị mặc định", "WARNING")
            return "4.0"  # Phiên bản mặc định nếu không đọc được

def prepare_build_environment():
    """Chuẩn bị môi trường build"""
    log("Chuẩn bị môi trường build...")
    
    # Tự động tìm đường dẫn file config
    current_dir = os.path.dirname(os.path.abspath(__file__))
    app_config_file = os.path.join(current_dir, "app_config.json")
    version_json_file = os.path.join(current_dir, "version.json")
    changelog_json_file = os.path.join(current_dir, "changelog.json")
    
    # QUAN TRỌNG: File app_config.json được bundle vào .exe sẽ làm MẪU
    # Khi app chạy lần đầu, nó sẽ copy config này vào AppData
    # Các lần sau, app sẽ đọc/ghi từ AppData (không ảnh hưởng file .exe)
    log("File app_config.json hiện tại sẽ được dùng làm mẫu cho app khi build", "INFO")
    
    # Kiểm tra xem các file cần thiết có tồn tại không
    missing_files = []
    for file_path, file_name in [
        (app_config_file, "app_config.json"),
        (version_json_file, "version.json"),
        (changelog_json_file, "changelog.json")
    ]:
        if not os.path.exists(file_path):
            missing_files.append(file_name)
    
    if missing_files:
        log(f"Không tìm thấy các file sau: {', '.join(missing_files)}. Vui lòng kiểm tra lại.", "ERROR")
        sys.exit(1)
    
    # Tạo thư mục dist nếu chưa tồn tại
    if not os.path.exists("dist"):
        os.makedirs("dist")
        log("Đã tạo thư mục dist", "INFO")
    
    # Kiểm tra UPX nếu có
    upx_dir = os.path.join(current_dir, "upx")
    if os.path.exists(upx_dir):
        log(f"Đã tìm thấy UPX trong thư mục: {upx_dir}", "INFO")
    else:
        log("Không tìm thấy UPX. Build sẽ không được nén. Bạn có thể tải UPX từ https://github.com/upx/upx/releases", "WARNING")
      # Sử dụng phiên bản fixed của file check_for_updates.py
    check_updates_fixed = os.path.join(current_dir, "check_for_updates_fixed.py")
    check_updates = os.path.join(current_dir, "check_for_updates.py")
    
    if os.path.exists(check_updates_fixed):
        log("Đã tìm thấy file check_for_updates_fixed.py, sẽ sử dụng phiên bản này", "INFO")
        # Tạo bản sao dự phòng của file gốc nếu chưa có
        if os.path.exists(check_updates) and not os.path.exists(check_updates + ".bak"):
            shutil.copy(check_updates, check_updates + ".bak")
            log("Đã tạo bản sao dự phòng của file check_for_updates.py", "INFO")
        
        # Copy file fixed để sử dụng
        shutil.copy(check_updates_fixed, check_updates)
        log("Đã sử dụng file check_for_updates_fixed.py cho build", "SUCCESS")
    else:
        log("Không tìm thấy file check_for_updates_fixed.py, sẽ sử dụng file check_for_updates.py hiện có", "INFO")
    
    log("Môi trường build đã sẵn sàng", "SUCCESS")
    
    # Trả về các file cần thêm vào build
    return app_config_file, version_json_file, changelog_json_file

def build_executable(version, config_files):
    """Build file thực thi với PyInstaller"""
    log(f"Bắt đầu build phiên bản {version}...")
    
    # Separator tuỳ hệ điều hành
    sep = ';' if os.name == 'nt' else ':'
    
    # Đảm bảo các đường dẫn file đều được normalize để tránh vấn đề với ký tự Unicode
    app_config_file, version_json_file, changelog_json_file = [
        os.path.normpath(os.path.abspath(file)) for file in config_files
    ]
    
    # Đặt biến môi trường để đảm bảo xử lý đúng encoding
    os.environ["PYTHONIOENCODING"] = "utf-8"
    
    # Các options cơ bản không bao gồm tối ưu hoá
    options = [
        'import_score.py',
        f'--name=Student Score Import v{version}',
        '--onefile',
        '--windowed',
        '--icon=app.ico',
        '--clean',
        '--noconfirm',
        '--noupx',  # Tắt UPX vì có thể gây lỗi với một số DLL
        '--exclude-module=pkg_resources',  # Loại bỏ pkg_resources để tránh lỗi jaraco
        f'--add-data={app_config_file}{sep}.',
        f'--add-data={version_json_file}{sep}.',
        f'--add-data={changelog_json_file}{sep}.',
        '--log-level=WARN',
        # Exclude các modules không cần thiết để giảm kích thước
        '--exclude-module=pytest',
        '--exclude-module=setuptools',
        '--exclude-module=pip',
        '--exclude-module=torch',
        '--exclude-module=tensorflow',
        '--exclude-module=scipy',
        '--exclude-module=scipy.stats',
        '--exclude-module=scipy.spatial',
        '--exclude-module=scipy.special',
        '--exclude-module=scipy.ndimage',
        '--exclude-module=scipy.linalg',
        '--exclude-module=jax',
        '--exclude-module=IPython',
        '--exclude-module=notebook',
        '--exclude-module=jupyter',
        '--exclude-module=tkinter.test',
        '--exclude-module=unittest',
        '--exclude-module=test',
        '--exclude-module=lib2to3',
        # Collect mode cho numpy để tránh lỗi import từ source
        '--collect-all=numpy',
        '--copy-metadata=numpy',
    ]
    
    # Thêm UPX nếu có
    upx_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "upx")
    if os.path.exists(upx_dir):
        options.append(f'--upx-dir={upx_dir}')
        log("Đã bật tính năng nén UPX", "INFO")

    # Thêm các hidden imports cần thiết (KHÔNG BAO GỒM NUMPY - đã dùng --collect-all)
    hidden_imports = [
        # Pandas và các phụ thuộc (numpy sẽ được collect bởi --collect-all)
        'pandas._libs.tslibs.base',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.timedeltas',
        
        # Matplotlib và các phụ thuộc
        'matplotlib.backends.backend_tkagg',
        'matplotlib.figure',
        'matplotlib.backends.backend_pdf',
        
        # Các module khác
        'PIL._tkinter_finder',
        'themes',
        'ui_utils',
        'version_utils',
        'check_for_updates',
        'openpyxl.cell',
    ]
    
    # Thêm hidden imports vào options
    for imp in hidden_imports:
        options.append(f'--hidden-import={imp}')
    
    # Chạy PyInstaller với các options đã chuẩn bị
    log(f"Chạy PyInstaller với numpy collect mode", "INFO")
    try:
        import PyInstaller.__main__
        PyInstaller.__main__.run(options)
        log(f"Build hoàn thành! Kiểm tra thư mục dist để xem file thực thi.", "SUCCESS")
        
        # Kiểm tra kết quả build
        exe_name = f"Student Score Import v{version}.exe"
        exe_path = os.path.join("dist", exe_name)
        
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            log(f"File thực thi đã được tạo: {exe_path}", "SUCCESS")
            log(f"Kích thước file: {size_mb:.2f}MB", "INFO")
        else:
            log(f"Không tìm thấy file thực thi trong thư mục dist", "ERROR")
        
        return True
    except Exception as e:
        log(f"Lỗi khi build: {str(e)}", "ERROR")
        sys.exit(1)

def main():
    """Hàm chính điều phối quá trình build"""
    print("\n" + "="*50)
    log("🚀 BẮT ĐẦU QUÁ TRÌNH BUILD", "INFO")
    print("="*50 + "\n")
    
    # Kiểm tra phụ thuộc
    check_dependencies()
    
    # Kiểm tra mã nguồn
    check_source_code()
    
    # Cập nhật thông tin phiên bản
    version = update_version_info()
    log(f"Phiên bản build: {version}", "INFO")
    
    # Chuẩn bị môi trường build
    config_files = prepare_build_environment()
    
    # Build ứng dụng với cấu hình đơn giản
    log("Sử dụng cấu hình build đơn giản", "INFO")
    success = build_executable(version, config_files)
    
    if success:
        print("\n" + "="*50)
        log("✅ QUÁ TRÌNH BUILD HOÀN TẤT THÀNH CÔNG", "SUCCESS")
        log(f"📦 File thực thi đã được tạo trong thư mục dist với tên: Student Score Import v{version}.exe", "SUCCESS")
        print("="*50 + "\n")
        return 0
    else:
        print("\n" + "="*50)
        log("❌ QUÁ TRÌNH BUILD THẤT BẠI", "ERROR")
        print("="*50 + "\n")
        return 1

if __name__ == "__main__":
    sys.exit(main())
