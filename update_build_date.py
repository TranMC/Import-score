#!/usr/bin/env python
"""
Script để cập nhật ngày build và trạng thái phiên bản
"""

import json
import os
import sys
from datetime import datetime

def main():
    """
    Hàm chính để cập nhật thông tin phiên bản
    """
    try:
        # Xác định đường dẫn tới file version.json
        script_dir = os.path.dirname(os.path.abspath(__file__))
        version_file = os.path.join(script_dir, 'version.json')
        
        # Đọc file version.json
        with open(version_file, 'r', encoding='utf-8') as f:
            version_info = json.load(f)
        
        # Lấy tham số từ dòng lệnh nếu có
        # update_build_date.py --set-channel dev|stable
        # update_build_date.py --toggle-dev
        if len(sys.argv) > 1:
            if sys.argv[1] == '--set-channel' and len(sys.argv) > 2:
                version_info['release_channel'] = sys.argv[2]
                print(f"Đã thiết lập kênh phát hành: {sys.argv[2]}")
                
            elif sys.argv[1] == '--toggle-dev':
                version_info['is_dev'] = not version_info.get('is_dev', False)
                print(f"Trạng thái is_dev: {version_info['is_dev']} " + 
                      f"({'Pre-release' if version_info['is_dev'] else 'Release'})")
                
            elif sys.argv[1] == '--set-dev' and len(sys.argv) > 2:
                is_dev = sys.argv[2].lower() in ('true', 'yes', '1', 't', 'y')
                version_info['is_dev'] = is_dev
                print(f"Đã thiết lập trạng thái is_dev: {is_dev}")
                
            elif sys.argv[1] == '--set-version' and len(sys.argv) > 2:
                version_info['version'] = sys.argv[2]
                print(f"Đã thiết lập phiên bản: {sys.argv[2]}")
        
        # Cập nhật ngày build
        version_info['build_date'] = datetime.now().strftime('%Y-%m-%d')
        print(f"Đã cập nhật ngày build: {version_info['build_date']}")
        
        # In thông tin phiên bản
        print(f"Phiên bản: {version_info['version']}")
        print(f"Trạng thái phát triển: {'Có' if version_info['is_dev'] else 'Không'}")
        print(f"Kênh phát hành: {version_info.get('release_channel', 'stable')}")
        print(f"Tên mã: {version_info.get('code_name', 'Unknown')}")
        
        # Lưu lại thông tin phiên bản
        with open(version_file, 'w', encoding='utf-8') as f:
            json.dump(version_info, f, indent=4, ensure_ascii=False)
        
        # Tạo file version.txt chứa phiên bản hiện tại
        with open(os.path.join(script_dir, 'version.txt'), 'w', encoding='utf-8') as f:
            f.write(version_info['version'])
            
        print("Cập nhật thành công!")
        return 0
    except Exception as e:
        print(f"Lỗi: {str(e)}")
        return 1

if __name__ == "__main__":
    sys.exit(main()) 