"""
Module để quản lý phiên bản và changelog
"""

import os
import json
import pkg_resources

# Các hằng số cho kênh cập nhật
UPDATE_CHANNEL_STABLE = "stable"
UPDATE_CHANNEL_DEV = "dev"

def load_version_info():
    """
    Đọc thông tin phiên bản từ file version.json
    
    Returns:
        dict: Thông tin phiên bản
    """
    try:
        # Xác định đường dẫn tới file version.json
        script_dir = os.path.dirname(os.path.abspath(__file__))
        version_file = os.path.join(script_dir, 'version.json')
        
        # Thử đọc từ đường dẫn tuyệt đối
        if os.path.exists(version_file):
            with open(version_file, 'r', encoding='utf-8') as f:
                version_info = json.load(f)
                return version_info
        
        # Thử đọc từ thư mục gốc của ứng dụng
        base_dir = os.path.abspath(os.path.join(script_dir, os.pardir))
        version_file = os.path.join(base_dir, 'version.json')
        if os.path.exists(version_file):
            with open(version_file, 'r', encoding='utf-8') as f:
                version_info = json.load(f)
                return version_info
                
        # Thử đọc từ thư mục hiện tại
        version_file = 'version.json'
        if os.path.exists(version_file):
            with open(version_file, 'r', encoding='utf-8') as f:
                version_info = json.load(f)
                return version_info
    except Exception as e:
        print(f"Lỗi khi đọc file version.json: {str(e)}")
    
    # Giá trị mặc định nếu không đọc được file
    return {
        "version": "0.0.0",
        "is_dev": True,
        "build_date": "2000-01-01",
        "code_name": "Unknown",
        "release_channel": "stable"
    }

def save_version_info(version_info):
    """
    Lưu thông tin phiên bản vào file version.json
    
    Args:
        version_info (dict): Thông tin phiên bản cần lưu
    
    Returns:
        bool: True nếu lưu thành công, False nếu có lỗi
    """
    try:
        # Xác định đường dẫn tới file version.json
        script_dir = os.path.dirname(os.path.abspath(__file__))
        version_file = os.path.join(script_dir, 'version.json')
        
        # Kiểm tra xem file có tồn tại không
        if not os.path.exists(version_file):
            # Thử tìm file ở thư mục gốc của ứng dụng
            base_dir = os.path.abspath(os.path.join(script_dir, os.pardir))
            version_file = os.path.join(base_dir, 'version.json')
            if not os.path.exists(version_file):
                # Tạo mới trong thư mục hiện tại
                version_file = 'version.json'
        
        # Lưu thông tin phiên bản
        with open(version_file, 'w', encoding='utf-8') as f:
            json.dump(version_info, f, indent=4, ensure_ascii=False)
        
        return True
    except Exception as e:
        print(f"Lỗi khi lưu file version.json: {str(e)}")
        return False

def update_build_date():
    """
    Cập nhật ngày build trong file version.json
    
    Returns:
        bool: True nếu cập nhật thành công, False nếu có lỗi
    """
    try:
        # Đọc thông tin phiên bản hiện tại
        version_info = load_version_info()
        
        # Cập nhật ngày build
        from datetime import datetime
        version_info['build_date'] = datetime.now().strftime('%Y-%m-%d')
        
        # Lưu lại thông tin phiên bản
        return save_version_info(version_info)
    except Exception as e:
        print(f"Lỗi khi cập nhật ngày build: {str(e)}")
        return False

def toggle_dev_mode():
    """
    Đảo ngược trạng thái is_dev trong file version.json
    
    Returns:
        bool: Trạng thái is_dev mới
    """
    try:
        # Đọc thông tin phiên bản hiện tại
        version_info = load_version_info()
        
        # Đảo ngược trạng thái is_dev
        version_info['is_dev'] = not version_info.get('is_dev', False)
        
        # Lưu lại thông tin phiên bản
        save_version_info(version_info)
        
        return version_info['is_dev']
    except Exception as e:
        print(f"Lỗi khi thay đổi trạng thái dev: {str(e)}")
        return None

def set_release_channel(channel):
    """
    Thiết lập kênh phát hành trong file version.json
    
    Args:
        channel (str): Tên kênh phát hành ('stable' hoặc 'dev')
    
    Returns:
        bool: True nếu thành công, False nếu thất bại
    """
    try:
        # Đọc thông tin phiên bản hiện tại
        version_info = load_version_info()
        
        # Thiết lập kênh phát hành
        version_info['release_channel'] = channel
        
        # Lưu lại thông tin phiên bản
        return save_version_info(version_info)
    except Exception as e:
        print(f"Lỗi khi thiết lập kênh phát hành: {str(e)}")
        return False

def load_changelog():
    """
    Đọc changelog từ file changelog.json
    
    Returns:
        dict: Changelog
    """
    try:
        # Đường dẫn tới file changelog.json
        script_dir = os.path.dirname(os.path.abspath(__file__))
        changelog_file = os.path.join(script_dir, "changelog.json")
        
        # Đọc file changelog.json
        with open(changelog_file, 'r', encoding='utf-8') as f:
            changelog = json.load(f)
        
        return changelog
    except Exception as e:
        print(f"Lỗi khi đọc changelog: {str(e)}")
        return {}

def get_current_version():
    """
    Lấy phiên bản hiện tại
    
    Returns:
        str: Phiên bản hiện tại
    """
    return load_version_info().get('version', '1.0.0')

def is_dev_version():
    """
    Kiểm tra có phải phiên bản phát triển không
    
    Returns:
        bool: True nếu là phiên bản phát triển
    """
    return load_version_info().get('is_dev', False)

def get_version_display():
    """
    Lấy chuỗi hiển thị phiên bản
    
    Returns:
        str: Chuỗi hiển thị phiên bản
    """
    version_info = load_version_info()
    version = version_info.get('version', '1.0.0')
    is_dev = version_info.get('is_dev', False)
    
    if is_dev:
        return f"v{version} dev"
    else:
        return f"v{version}"

def get_changelog_for_version(version=None):
    """
    Lấy changelog cho một phiên bản cụ thể
    
    Args:
        version (str, optional): Phiên bản cần lấy changelog. 
                              Nếu None, lấy phiên bản hiện tại.
    
    Returns:
        list: Danh sách các thay đổi
    """
    if version is None:
        version = get_current_version()
        
    changelog = load_changelog()
    return changelog.get(version, [])

def get_update_channel():
    """
    Lấy kênh cập nhật hiện tại
    
    Returns:
        str: Kênh cập nhật hiện tại
    """
    return load_version_info().get('release_channel', UPDATE_CHANNEL_STABLE)

def get_all_versions():
    """
    Lấy danh sách tất cả các phiên bản từ changelog
    
    Returns:
        list: Danh sách các phiên bản, sắp xếp từ mới đến cũ
    """
    changelog = load_changelog()
    versions = list(changelog.keys())
    
    # Sắp xếp phiên bản từ mới đến cũ
    versions.sort(key=lambda v: [int(x) for x in v.split('.')], reverse=True)
    
    return versions 