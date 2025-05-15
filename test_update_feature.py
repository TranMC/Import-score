"""
Script kiểm tra tính năng cập nhật tự động
"""

import tkinter as tk
from tkinter import ttk
import os
import json
import sys

# Thêm thư mục hiện tại vào đường dẫn để có thể import các module
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Đảm bảo chạy script từ đúng thư mục
os.chdir(os.path.dirname(os.path.abspath(__file__)))

def run_test():
    print("=== BẮT ĐẦU KIỂM TRA TÍNH NĂNG CẬP NHẬT TỰ ĐỘNG ===")
    
    # Kiểm tra file version.json
    try:
        with open('version.json', 'r', encoding='utf-8') as f:
            version_info = json.load(f)
            print(f"Phiên bản hiện tại: {version_info.get('version', 'Không xác định')}")
    except Exception as e:
        print(f"Lỗi khi đọc file version.json: {str(e)}")
    
    # Kiểm tra các file cần thiết
    required_files = [
        'check_for_updates.py',
        'version_utils.py'
    ]
    
    for file in required_files:
        if os.path.exists(file):
            print(f"✓ File {file} tồn tại")
        else:
            print(f"✗ File {file} không tồn tại")
    
    # Tạo cửa sổ test
    root = tk.Tk()
    root.title("Kiểm tra Cập nhật Tự động")
    root.geometry("600x400")
    
    # Tạo style
    style = ttk.Style()
    style.configure("StatusInfo.TLabel", foreground="blue")
    style.configure("StatusWarning.TLabel", foreground="orange")
    style.configure("StatusError.TLabel", foreground="red")
    style.configure("StatusGood.TLabel", foreground="green")
    style.configure("StatusCritical.TLabel", foreground="red", font=("Segoe UI", 10, "bold"))
    
    # Tạo UI cơ bản
    frame = ttk.Frame(root, padding=10)
    frame.pack(fill="both", expand=True)
    
    # Tạo status label để test cập nhật UI
    status_label = ttk.Label(frame, text="Chưa kiểm tra cập nhật", style="StatusInfo.TLabel")
    status_label.pack(pady=10)
    
    # Tạo config giả lập
    config = {
        'ui': {
            'theme': {
                'primary': '#1976D2',
                'background': '#FFFFFF',
                'card': '#F5F5F5',
                'text': '#000000'
            },
            'font_family': 'Segoe UI',
            'font_size': {
                'normal': 11
            }
        }
    }
    
    def save_config_mock():
        print("Giả lập lưu cấu hình")
    
    # Test check_for_updates
    def test_fixed_update():
        try:
            from check_for_updates import check_for_updates
            print("Đang kiểm tra check_for_updates.py...")
            result = check_for_updates(root, status_label, None, config, save_config_mock, True)
            print(f"Kết quả kiểm tra: {'Có phiên bản mới' if result else 'Không có phiên bản mới hoặc có lỗi'}")
        except Exception as e:
            print(f"Lỗi khi kiểm tra check_for_updates.py: {str(e)}")
    
    # Test check_for_updates original
    def test_original_update():
        try:
            from check_for_updates import check_for_updates
            print("Đang kiểm tra check_for_updates.py...")
            result = check_for_updates(root, status_label, None, config, save_config_mock, True)
            print(f"Kết quả kiểm tra: {'Có phiên bản mới' if result else 'Không có phiên bản mới hoặc có lỗi'}")
        except Exception as e:
            print(f"Lỗi khi kiểm tra check_for_updates.py: {str(e)}")
    
    # Test kiểm tra tự động không hiển thị thông báo
    def test_auto_update():
        try:
            from check_for_updates import check_for_updates
            print("Đang kiểm tra chế độ tự động (không hiển thị thông báo)...")
            result = check_for_updates(root, status_label, None, config, save_config_mock, False)
            print(f"Kết quả kiểm tra tự động: {'Có phiên bản mới' if result else 'Không có phiên bản mới hoặc có lỗi'}")
        except Exception as e:
            print(f"Lỗi khi kiểm tra chế độ tự động: {str(e)}")
    
    # Test kiểm tra cập nhật bất đồng bộ
    def test_async_update():
        try:
            from check_for_updates import check_updates_async
            print("Đang kiểm tra hàm check_updates_async()...")
            check_updates_async(root, status_label, None, config, save_config_mock)
            print("Đã gọi hàm check_updates_async() thành công")
        except Exception as e:
            print(f"Lỗi khi kiểm tra check_updates_async: {str(e)}")
    
    # Tạo các nút kiểm tra
    ttk.Button(frame, text="Kiểm tra phiên bản mới (Fixed)", command=test_fixed_update).pack(pady=5)
    ttk.Button(frame, text="Kiểm tra phiên bản mới (Original)", command=test_original_update).pack(pady=5)
    ttk.Button(frame, text="Kiểm tra tự động (Không thông báo)", command=test_auto_update).pack(pady=5)
    ttk.Button(frame, text="Kiểm tra bất đồng bộ", command=test_async_update).pack(pady=5)
    
    # Nút thoát
    ttk.Button(frame, text="Thoát", command=root.destroy).pack(pady=20)
    
    print("Giao diện kiểm tra đã sẵn sàng. Hãy nhấn các nút để kiểm tra tính năng.")
    root.mainloop()

if __name__ == "__main__":
    run_test()
