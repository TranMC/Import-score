"""
Module để quản lý phiên bản và changelog
"""

import os
import json
import pkg_resources

def load_version_info():
    """
    Đọc thông tin phiên bản từ file version.json
    
    Returns:
        dict: Thông tin phiên bản
    """
    try:
        # Đường dẫn tới file version.json
        script_dir = os.path.dirname(os.path.abspath(__file__))
        version_file = os.path.join(script_dir, "version.json")
        
        # Đọc file version.json
        with open(version_file, 'r', encoding='utf-8') as f:
            version_info = json.load(f)
        
        return version_info
    except Exception as e:
        print(f"Lỗi khi đọc thông tin phiên bản: {str(e)}")
        return {
            "version": "1.0.0",
            "is_dev": False,
            "build_date": "",
            "code_name": "Unknown"
        }

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