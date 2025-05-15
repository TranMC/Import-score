#!/usr/bin/env python
"""
Script cập nhật build_date trong version.json
Hoạt động tốt cả trong môi trường phát triển và CI/CD
"""

import os
import json
import sys
from datetime import datetime

def update_build_date(include_time=False):
    """
    Cập nhật build_date trong version.json với ngày hiện tại
    
    Args:
        include_time (bool): Nếu True, thêm giờ:phút:giây vào build_date
    
    Returns:
        bool: True nếu cập nhật thành công
    """
    try:
        # Đường dẫn tới file version.json
        script_dir = os.path.dirname(os.path.abspath(__file__))
        version_file = os.path.join(script_dir, "version.json")
        
        # Đọc file version.json
        with open(version_file, 'r', encoding='utf-8') as f:
            version_data = json.load(f)
        
        # Định dạng datetime tùy theo include_time
        date_format = "%Y-%m-%d %H:%M:%S" if include_time else "%Y-%m-%d"
        current_date = datetime.now().strftime(date_format)
        
        # Lưu giá trị cũ để kiểm tra thay đổi
        old_date = version_data.get('build_date', '')
        
        # Cập nhật build_date
        version_data['build_date'] = current_date
        
        # Lưu lại file version.json với build_date mới
        with open(version_file, 'w', encoding='utf-8') as f:
            json.dump(version_data, f, indent=4, ensure_ascii=False)
        
        # Log kết quả
        if old_date != current_date:
            print(f"✅ Đã cập nhật build_date: '{old_date}' -> '{current_date}'")
        else:
            print(f"ℹ️ build_date không thay đổi: '{current_date}'")
        
        # Kiểm tra môi trường CI/CD (GitHub Actions)
        if os.environ.get('GITHUB_ACTIONS') == 'true':
            print(f"BUILD_DATE={current_date}")
            # Nếu đang chạy trong GitHub Actions, thêm biến môi trường
            with open(os.environ.get('GITHUB_ENV', ''), 'a') as f:
                f.write(f"BUILD_DATE={current_date}\n")
        
        return True
    except Exception as e:
        print(f"❌ Lỗi khi cập nhật build_date: {str(e)}", file=sys.stderr)
        return False

if __name__ == "__main__":
    # Đọc tham số dòng lệnh nếu có
    include_time = "--include-time" in sys.argv
    success = update_build_date(include_time)
    if not success:
        sys.exit(1) 