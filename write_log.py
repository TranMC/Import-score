import datetime
import os
import re
import json
import sys

print("Bắt đầu tạo log phiên bản...")

# Đảm bảo sử dụng UTF-8 cho output
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

now = datetime.datetime.now()
# Đọc version từ import_score.py
try:
    with open("import_score.py", "r", encoding="utf-8") as f:
        content = f.read()
        
        # Tìm kiếm version trong cấu trúc config
        version_match = re.search(r"'version':\s*'([^']+)'", content)
        if version_match:
            version = version_match.group(1)
            print(f"Đã tìm thấy phiên bản: {version}")
        else:
            raise ValueError("Version không tìm thấy trong import_score.py")
except Exception as e:
    version = "unknown"
    print(f"Lỗi khi đọc version: {str(e)}")

# Ghi vào version.txt để GitLab dùng
with open("version.txt", "w", encoding="utf-8") as f:
    f.write(version)
print(f"Đã ghi phiên bản {version} vào version.txt")

# Đọc changelog từ import_score.py
changelog = "Unknown"
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
                else:
                    changelog = "Không có thông tin changelog chi tiết"
                    print("Không tìm thấy chi tiết changelog")
            else:
                changelog = f"Không tìm thấy changelog cho phiên bản {version}"
                print(f"Không tìm thấy changelog cho phiên bản {version}")
        else:
            changelog = "Không tìm thấy phần changelog trong config"
            print("Không tìm thấy phần changelog trong config")
            
except Exception as e:
    changelog = f"Lỗi khi đọc changelog: {str(e)}"
    print(f"Lỗi khi đọc changelog: {str(e)}")

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
else:
    print("File CHANGELOG.md chưa tồn tại, sẽ tạo mới")

# Kiểm tra xem phiên bản này đã được ghi chưa
if version_exists:
    print(f"Phiên bản {version} đã tồn tại trong CHANGELOG.md, không ghi lại")
else:
    # Nếu file chưa tồn tại, thêm tiêu đề
    if not os.path.exists("CHANGELOG.md") or not existing_content:
        header = "# Lịch sử thay đổi\n\n"
        new_entry = header + new_entry
        print("Thêm tiêu đề vào CHANGELOG mới")
    
    # Ghi nội dung mới + nội dung cũ
    with open("CHANGELOG.md", "w", encoding="utf-8") as f:
        f.write(new_entry)
        if existing_content:
            f.write(existing_content)
    
    print(f"Đã cập nhật CHANGELOG.md với phiên bản {version}")

print("Hoàn thành!")
