import datetime
import os
import re
import json
import sys
import traceback
import version_utils

def setup_encoding():
    """Set up UTF-8 encoding for console output"""
    try:
        # Try to force UTF-8 encoding
        if sys.stdout.encoding != 'utf-8':
            try:
                sys.stdout.reconfigure(encoding='utf-8')
            except (AttributeError, IOError):
                # For older Python versions or when reconfigure fails
                pass
        
        # Set environment variables to force UTF-8
        os.environ["PYTHONIOENCODING"] = "utf-8"
    except Exception as e:
        print(f"Warning: Failed to set encoding: {str(e)}")

def extract_version():
    """Extract version from version.json file using version_utils"""
    try:
        version_info = version_utils.load_version_info()
        version = version_info.get('version', 'unknown')
        print(f"Found version: {version}")
        return version
    except Exception as e:
        print(f"Error reading version: {str(e)}")
        print(f"Error details: {traceback.format_exc()}")
        return "unknown"

def write_version_file(version):
    """Write version to version.txt file for CI/CD use"""
    try:
        with open("version.txt", "w", encoding="utf-8") as f:
            f.write(version)
        print(f"Wrote version {version} to version.txt")
        return True
    except Exception as e:
        print(f"Error writing version.txt: {str(e)}")
        return False

def extract_changelog(version):
    """Extract changelog from changelog.json file using version_utils"""
    try:
        # Sử dụng hàm từ version_utils để lấy changelog cho phiên bản cụ thể
        changelog_items = version_utils.get_changelog_for_version(version)
        
        if changelog_items:
            # Chuyển đổi các mục changelog thành định dạng markdown
            changelog = "\n- " + "\n- ".join(changelog_items)
            print(f"Found {len(changelog_items)} changelog items")
            return changelog
        else:
            print(f"No changelog found for version {version}")
            return f"No changelog found for version {version}"
                
    except Exception as e:
        print(f"Error reading changelog: {str(e)}")
        print(f"Error details: {traceback.format_exc()}")
        return f"Error reading changelog: {str(e)}"

def update_changelog_file(version, changelog):
    """Update CHANGELOG.md file with new version information"""
    now = datetime.datetime.now()
    
    # Đảm bảo năm là năm hiện tại, không phải năm tương lai
    if now.year > 2024:  # Kiểm tra năm có hợp lý không
        now = now.replace(year=2024)  # Sử dụng năm 2024 nếu có vấn đề
    
    # Create new content with markdown format
    new_entry = f"## Version {version}\n"
    new_entry += f"### Date: {now.strftime('%d/%m/%Y %H:%M:%S')}\n"
    new_entry += f"### Changes:{changelog}\n"
    new_entry += "\n" + "-" * 50 + "\n"

    # Read existing content if file exists
    existing_content = ""
    version_exists = False
    if os.path.exists("CHANGELOG.md"):
        try:
            with open("CHANGELOG.md", "r", encoding="utf-8") as f:
                existing_content = f.read()
                print("Reading existing CHANGELOG.md file")
                # Find version in content
                if f"## Version {version}" in existing_content or f"## Phiên bản {version}" in existing_content:
                    version_exists = True
                    print(f"Version {version} already exists in file")
        except Exception as e:
            print(f"Error reading CHANGELOG.md: {str(e)}")
            print(f"Error details: {traceback.format_exc()}")
    else:
        print("CHANGELOG.md file does not exist, will create new")

    # Check if this version has already been written
    if version_exists:
        print(f"Version {version} already exists in CHANGELOG.md, won't write again")
        return False
    else:
        # If file doesn't exist, add title
        if not os.path.exists("CHANGELOG.md") or not existing_content:
            header = "# Change History\n\n"
            new_entry = header + new_entry
            print("Adding title to new CHANGELOG")
        
        # Write new content + old content
        try:
            with open("CHANGELOG.md", "w", encoding="utf-8") as f:
                f.write(new_entry)
                if existing_content:
                    f.write(existing_content)
            
            print(f"Updated CHANGELOG.md with version {version}")
            return True
        except Exception as e:
            print(f"Error writing CHANGELOG.md: {str(e)}")
            print(f"Error details: {traceback.format_exc()}")
            return False

def main():
    """Main function coordinating the entire process"""
    print("Starting version log creation...")
    
    # Set up UTF-8 encoding
    setup_encoding()
    
    # Extract version
    version = extract_version()
    
    # Write version to version.txt
    write_version_file(version)
    
    # Extract changelog
    changelog = extract_changelog(version)
    
    # Update CHANGELOG.md file
    update_changelog_file(version, changelog)
    
    print("Completed!")
    return 0

if __name__ == "__main__":
    # Force ASCII output for error messages to avoid encoding issues
    try:
        sys.stderr.reconfigure(encoding='utf-8')
    except (AttributeError, IOError):
        pass
        
    sys.exit(main())
