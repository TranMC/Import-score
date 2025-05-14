import datetime
import os
import re
import json
import sys
import traceback

def setup_encoding():
    # Đảm bảo sử dụng UTF-8 cho output
    if sys.stdout.encoding != 'utf-8':
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except AttributeError:
            # Xử lý trường hợp Python phiên bản cũ không có reconfigure
            pass

def extract_version():
    """Trích xuất phiên bản từ file import_score.py"""
    try:
        with open("import_score.py", "r", encoding="utf-8") as f:
            content = f.read()
            
            # Tìm kiếm version trong cấu trúc config
            version_match = re.search(r"'version':\s*'([^']+)'", content)
            if version_match:
                version = version_match.group(1)
                print(f"Đã tìm thấy phiên bản: {version}")
                return version
            else:
                raise ValueError("Version không tìm thấy trong import_score.py")
    except Exception as e:
        print(f"Lỗi khi đọc version: {str(e)}")
        print(f"Chi tiết lỗi: {traceback.format_exc()}")
        return "unknown"

def write_version_file(version):
    """Ghi phiên bản vào file version.txt để CI/CD sử dụng"""
    try:
        with open("version.txt", "w", encoding="utf-8") as f:
            f.write(version)
        print(f"Đã ghi phiên bản {version} vào version.txt")
        return True
    except Exception as e:
        print(f"Lỗi khi ghi version.txt: {str(e)}")
        return False

def extract_changelog(version):
    """Trích xuất changelog từ file import_score.py"""
    try:
        with open("import_score.py", "r", encoding="utf-8") as f:
            content = f.read()
            print("Đang tìm changelog...")
            
            # Tìm changelog cụ thể cho phiên bản hiện tại trong config
            changelog_section = re.search(r"'changelog'\s*:\s*{([^}]*)}", content, re.DOTALL)
            
            if changelog_section:
                changelog_dict_text = '{' + changelog_section.group(1) + '}'
                # Chuyển đổi text thành json-compatible
                changelog_dict_text = changelog_dict_text.replace("'", '"')
                
                # Tìm và xử lý phiên bản hiện tại
                version_pattern = f'"{version}"\\s*:\\s*\\[(.*?)\\]'
                version_changelog_match = re.search(version_pattern, changelog_dict_text, re.DOTALL)
                
                if version_changelog_match:
                    # Lấy nội dung changelog
                    version_changelog_text = version_changelog_match.group(1)
                    # Tách các mục changelog
                    changelog_items = re.findall(r'"([^"]*)"', version_changelog_text)
                    if changelog_items:
                        changelog = "\n- " + "\n- ".join(changelog_items)
                        print(f"Đã tìm thấy {len(changelog_items)} mục changelog")
                        return changelog
                    else:
                        print("Không tìm thấy chi tiết changelog")
                        return "Không có thông tin changelog chi tiết"
                else:
                    print(f"Không tìm thấy changelog cho phiên bản {version}")
                    return f"Không tìm thấy changelog cho phiên bản {version}"
            else:
                print("Không tìm thấy phần changelog trong config")
                return "Không tìm thấy phần changelog trong config"
                
    except Exception as e:
        print(f"Lỗi khi đọc changelog: {str(e)}")
        print(f"Chi tiết lỗi: {traceback.format_exc()}")
        return f"Lỗi khi đọc changelog: {str(e)}"

def update_changelog_file(version, changelog):
    """Cập nhật file CHANGELOG.md với thông tin phiên bản mới"""
    now = datetime.datetime.now()
    
    # Tạo nội dung mới với định dạng markdown
    new_entry = f"## Phiên bản {version}\n"
    new_entry += f"### Ngày: {now.strftime('%d/%m/%Y %H:%M:%S')}\n"
    new_entry += f"### Thay đổi:{changelog}\n"
    new_entry += "\n" + "-" * 50 + "\n"

    # Đọc nội dung hiện tại nếu file đã tồn tại
    existing_content = ""
    version_exists = False
    if os.path.exists("CHANGELOG.md"):
        try:
            with open("CHANGELOG.md", "r", encoding="utf-8") as f:
                existing_content = f.read()
                print("Đọc file CHANGELOG.md hiện tại")
                # Tìm phiên bản trong nội dung
                if f"## Phiên bản {version}" in existing_content or f"Build Version: {version}" in existing_content:
                    version_exists = True
                    print(f"Phiên bản {version} đã tồn tại trong file")
        except Exception as e:
            print(f"Lỗi khi đọc CHANGELOG.md: {str(e)}")
            print(f"Chi tiết lỗi: {traceback.format_exc()}")
    else:
        print("File CHANGELOG.md chưa tồn tại, sẽ tạo mới")

    # Kiểm tra xem phiên bản này đã được ghi chưa
    if version_exists:
        print(f"Phiên bản {version} đã tồn tại trong CHANGELOG.md, không ghi lại")
        return False
    else:
        # Nếu file chưa tồn tại, thêm tiêu đề
        if not os.path.exists("CHANGELOG.md") or not existing_content:
            header = "# Lịch sử thay đổi\n\n"
            new_entry = header + new_entry
            print("Thêm tiêu đề vào CHANGELOG mới")
        
        # Ghi nội dung mới + nội dung cũ
        try:
            with open("CHANGELOG.md", "w", encoding="utf-8") as f:
                f.write(new_entry)
                if existing_content:
                    f.write(existing_content)
            
            print(f"Đã cập nhật CHANGELOG.md với phiên bản {version}")
            return True
        except Exception as e:
            print(f"Lỗi khi ghi CHANGELOG.md: {str(e)}")
            print(f"Chi tiết lỗi: {traceback.format_exc()}")
            return False

def main():
    """Hàm chính điều phối toàn bộ quá trình"""
    print("Bắt đầu tạo log phiên bản...")
    
    # Thiết lập mã hóa UTF-8
    setup_encoding()
    
    # Trích xuất phiên bản
    version = extract_version()
    
    # Ghi phiên bản vào file version.txt
    write_version_file(version)
    
    # Trích xuất changelog
    changelog = extract_changelog(version)
    
    # Cập nhật file CHANGELOG.md
    update_changelog_file(version, changelog)
    
    print("Hoàn thành!")
    return 0

if __name__ == "__main__":
    sys.exit(main())
