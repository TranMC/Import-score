"""
Module để kiểm tra cập nhật phiên bản mới
"""

import requests
import tkinter as tk
from tkinter import ttk, messagebox
import os
from datetime import datetime
import threading
import json

def check_for_updates(root, status_label, file_path, config, save_config, show_notification=True):
    """
    Kiểm tra phiên bản mới trên GitHub
    
    Args:
        root: Cửa sổ chính của ứng dụng
        status_label: Label hiển thị trạng thái
        file_path: Đường dẫn file hiện tại đang mở
        config: Dictionary chứa cấu hình ứng dụng
        save_config: Hàm lưu cấu hình
        show_notification (bool): Có hiển thị thông báo khi không có phiên bản mới không
    
    Returns:
        bool: True nếu có phiên bản mới, False nếu không
    """
    try:
        # URL của GitHub API để kiểm tra phiên bản mới nhất
        github_api_url = "https://api.github.com/repos/TranMC/Import-score/releases/latest"
        
        # Hiển thị thông báo đang kiểm tra
        if show_notification:
            if status_label:
                status_label.config(text="Đang kiểm tra phiên bản mới...", style="StatusInfo.TLabel")
                root.update()
        
        # Gọi API để lấy thông tin phiên bản mới nhất với timeout bị tăng lên và headers
        headers = {
            'User-Agent': 'QuanLyDiemHocSinh-App',
            'Accept': 'application/vnd.github.v3+json'
        }
        response = requests.get(github_api_url, headers=headers, timeout=10)
        
        # Kiểm tra kết quả
        if response.status_code == 200:
            release_info = response.json()
            
            # Lấy phiên bản mới nhất
            latest_version = release_info.get('tag_name', '').lstrip('v')
            # Lấy phiên bản hiện tại từ version.json
            import version_utils
            version_info = version_utils.load_version_info()
            current_version = version_info.get('version', '0.0.0')
            is_dev = version_info.get('is_dev', False)
            
            # Kiểm tra nếu phiên bản hiện tại cao hơn phiên bản mới nhất, đánh dấu là dev
            if current_version > latest_version and not is_dev:
                # Thông báo cho người dùng biết họ đang sử dụng phiên bản phát triển
                if show_notification:
                    if file_path:
                        status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusGood.TLabel")
                    messagebox.showinfo("Phiên bản phát triển", f"Bạn đang sử dụng phiên bản phát triển ({current_version} dev version). Phiên bản phát hành mới nhất là {latest_version}.")
                return False
                
            # So sánh phiên bản
            elif latest_version > current_version:
                # Hiển thị thông báo có phiên bản mới
                if show_notification:
                    release_notes = release_info.get('body', 'Không có thông tin chi tiết.')
                    result = messagebox.askyesno(
                        "Có phiên bản mới",
                        f"Đã phát hiện phiên bản mới: {latest_version}\n"
                        f"Phiên bản hiện tại: {current_version}\n\n"
                        f"Tính năng mới:\n{release_notes[:200]}{'...' if len(release_notes) > 200 else ''}\n\n"
                        "Bạn có muốn tải phiên bản mới không?"
                    )
                    
                    if result:
                        # Tìm URL để tải xuống
                        download_url = ""
                        for asset in release_info.get('assets', []):
                            if asset.get('name', '').endswith('.exe'):
                                download_url = asset.get('browser_download_url', '')
                                break
                                
                        if download_url:
                            try:
                                # Hiển thị thông báo đang tải
                                status_label.config(text=f"Đang tải phiên bản mới {latest_version}...", 
                                                 style="StatusWarning.TLabel")
                                root.update()
                                
                                # Tạo thư mục tạm để tải xuống
                                temp_dir = os.path.join(os.path.expanduser("~"), "Downloads")
                                os.makedirs(temp_dir, exist_ok=True)
                                
                                # Tạo tên file tạm
                                file_name = os.path.basename(download_url)
                                temp_file = os.path.join(temp_dir, file_name)
                                
                                # Tải xuống tệp
                                with requests.get(download_url, stream=True) as r:
                                    r.raise_for_status()
                                    total_size = int(r.headers.get('content-length', 0))
                                    
                                    # Tạo cửa sổ hiển thị tiến trình
                                    progress_window = tk.Toplevel(root)
                                    progress_window.title("Đang tải xuống")
                                    progress_window.geometry("450x200")
                                    progress_window.transient(root)  # Đặt là cửa sổ con của root
                                    
                                    # Đặt cửa sổ ở giữa màn hình
                                    window_width = 450
                                    window_height = 200
                                    screen_width = root.winfo_screenwidth()
                                    screen_height = root.winfo_screenheight()
                                    position_x = int(screen_width/2 - window_width/2)
                                    position_y = int(screen_height/2 - window_height/2)
                                    progress_window.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")
                                    
                                    # Ngăn không cho đóng cửa sổ tiến trình
                                    progress_window.protocol("WM_DELETE_WINDOW", lambda: None)
                                    
                                    # Thiết lập màu nền
                                    progress_window_frame = ttk.Frame(progress_window, padding=15)
                                    progress_window_frame.pack(fill="both", expand=True)
                                    
                                    # Label hiển thị thông tin
                                    ttk.Label(progress_window_frame, 
                                           text=f"Đang tải xuống phiên bản {latest_version}...",
                                           font=(config['ui']['font_family'], 12, 'bold'),
                                           foreground=config['ui']['theme']['primary']).pack(pady=5)
                                    
                                    # Frame chứa thông tin chi tiết
                                    info_frame = ttk.Frame(progress_window_frame)
                                    info_frame.pack(fill="x", pady=5)
                                    
                                    # Label hiển thị tên file
                                    file_label = ttk.Label(info_frame, 
                                                        text=f"File: {file_name}",
                                                        font=(config['ui']['font_family'], 10))
                                    file_label.pack(anchor="w", pady=2)
                                    
                                    # Label hiển thị kích thước
                                    size_label = ttk.Label(info_frame, 
                                                        text=f"Kích thước: {total_size/1024/1024:.1f} MB",
                                                        font=(config['ui']['font_family'], 10))
                                    size_label.pack(anchor="w", pady=2)
                                    
                                    # Tạo style cho thanh tiến trình
                                    style = ttk.Style()
                                    style.configure("Download.Horizontal.TProgressbar", 
                                                  troughcolor=config['ui']['theme']['background'],
                                                  background=config['ui']['theme']['primary'],
                                                  thickness=15)
                                    
                                    # Thanh tiến trình
                                    progress = ttk.Progressbar(progress_window_frame, 
                                                            orient="horizontal", 
                                                            length=400, 
                                                            mode="determinate",
                                                            style="Download.Horizontal.TProgressbar")
                                    progress.pack(pady=10, fill="x")
                                    
                                    # Frame chứa thông tin tải xuống
                                    download_info_frame = ttk.Frame(progress_window_frame)
                                    download_info_frame.pack(fill="x", pady=5)
                                    
                                    # Label hiển thị phần trăm
                                    percent_label = ttk.Label(download_info_frame, 
                                                           text="0%", 
                                                           font=(config['ui']['font_family'], 10, 'bold'),
                                                           foreground=config['ui']['theme']['primary'])
                                    percent_label.pack(side="left", padx=5)
                                    
                                    # Label hiển thị thông tin tải
                                    download_detail_label = ttk.Label(download_info_frame, 
                                                                   text="(0.0/0.0 MB)", 
                                                                   font=(config['ui']['font_family'], 10))
                                    download_detail_label.pack(side="right", padx=5)
                                    
                                    downloaded = 0
                                    start_time = datetime.now()
                                    
                                    # Ghi tệp đã tải xuống
                                    with open(temp_file, 'wb') as f:
                                        for chunk in r.iter_content(chunk_size=8192):
                                            if chunk:  # Lọc ra các gói rỗng để giữ kết nối
                                                f.write(chunk)
                                                downloaded += len(chunk)
                                                
                                                # Cập nhật tiến trình
                                                if total_size > 0:
                                                    percent = int(100 * downloaded / total_size)
                                                    progress["value"] = percent
                                                    percent_label.config(text=f"{percent}%")
                                                    download_detail_label.config(text=f"({downloaded/1024/1024:.1f}/{total_size/1024/1024:.1f} MB)")
                                                    progress_window.update_idletasks()
                                    
                                    # Hoàn thành tải xuống
                                    progress["value"] = 100
                                    percent_label.config(text="100%")
                                    download_detail_label.config(text=f"({total_size/1024/1024:.1f}/{total_size/1024/1024:.1f} MB)")
                                    
                                    # Đóng cửa sổ tiến trình sau 1 giây
                                    progress_window.after(1000, progress_window.destroy)
                                
                                # Hỏi người dùng có muốn cài đặt không
                                install_now = messagebox.askyesno("Tải xuống hoàn tất", 
                                                             f"Đã tải phiên bản {latest_version} về:\n{temp_file}\n\nBạn có muốn cài đặt ngay bây giờ không?")
                                
                                if install_now:
                                    # Chạy tệp cài đặt và thoát ứng dụng hiện tại
                                    os.startfile(temp_file)
                                    root.after(1000, root.destroy)
                                
                                status_label.config(text=f"Đã tải phiên bản mới {latest_version}", 
                                                  style="StatusSuccess.TLabel")
                            except Exception as download_error:
                                messagebox.showerror("Lỗi tải xuống", 
                                                f"Không thể tải phiên bản mới: {str(download_error)}\n\nVui lòng tải thủ công từ trang web.")
                                import webbrowser
                                webbrowser.open(release_info.get('html_url', ''))
                        else:
                            import webbrowser
                            webbrowser.open(release_info.get('html_url', ''))
                
                if file_path:
                    status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusGood.TLabel")
                return True
            else:
                # Không có phiên bản mới
                if show_notification:
                    if file_path:
                        status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusGood.TLabel")
                    messagebox.showinfo("Cập nhật", f"Bạn đang sử dụng phiên bản mới nhất ({current_version}).")
                return False
        else:
            # Lỗi kết nối
            if show_notification:
                if file_path:
                    status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusError.TLabel")
                messagebox.showwarning("Lỗi kết nối", f"Không thể kết nối đến máy chủ GitHub. Mã lỗi: {response.status_code}")
            return False
    
    except requests.exceptions.ConnectionError as e:
        # Lỗi kết nối mạng cụ thể
        if show_notification:
            if file_path and status_label:
                status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusGood.TLabel")
            messagebox.showwarning("Lỗi kết nối", "Không thể kết nối đến máy chủ GitHub. Vui lòng kiểm tra kết nối mạng của bạn.")
        print(f"Lỗi kết nối: {str(e)}")
        return False
    
    except requests.exceptions.Timeout as e:
        # Lỗi timeout
        if show_notification:
            if file_path and status_label:
                status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusGood.TLabel")
            messagebox.showwarning("Lỗi kết nối", "Máy chủ GitHub phản hồi quá chậm. Vui lòng thử lại sau.")
        print(f"Lỗi timeout: {str(e)}")
        return False
        
    except requests.RequestException as e:
        # Xử lý lỗi request khác
        if show_notification:
            if file_path and status_label:
                status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusGood.TLabel")
            messagebox.showwarning("Lỗi kết nối", f"Không thể kiểm tra phiên bản mới: {str(e)}")
        print(f"Lỗi request: {str(e)}")
        return False
    
    except json.JSONDecodeError as e:
        # Lỗi khi parse JSON
        if show_notification:
            if file_path and status_label:
                status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusGood.TLabel")
            messagebox.showwarning("Lỗi dữ liệu", "Không thể đọc thông tin phiên bản từ máy chủ. Định dạng dữ liệu không hợp lệ.")
        print(f"Lỗi parse JSON: {str(e)}")
        return False
    
    except Exception as e:
        # Xử lý các lỗi khác
        if show_notification:
            if file_path and status_label:
                status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusGood.TLabel")
            messagebox.showwarning("Lỗi", f"Đã xảy ra lỗi khi kiểm tra phiên bản mới: {str(e)}")
        print(f"Lỗi không xác định: {str(e)}")
        return False

def check_updates_async(root, status_label, file_path, config, save_config):
    """
    Kiểm tra cập nhật trong luồng riêng biệt để không làm đóng băng giao diện
    
    Args:
        root: Cửa sổ chính của ứng dụng
        status_label: Label hiển thị trạng thái
        file_path: Đường dẫn file hiện tại đang mở
        config: Dictionary chứa cấu hình ứng dụng
        save_config: Hàm lưu cấu hình
    """
    print("Starting async check for updates...")
    threading.Thread(
        target=lambda: check_for_updates(root, status_label, file_path, config, save_config, False), 
        daemon=True
    ).start()
    print("Update check thread started")

# Thêm main function để module có thể chạy độc lập
if __name__ == "__main__":
    print("Running check_for_updates module directly")
    print("This module is designed to be imported, but we'll attempt to check for updates directly")
    
    # Tạo config đơn giản
    dummy_config = {
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
    
    # Hàm save_config giả
    def dummy_save_config():
        print("Dummy save_config called")
    
    # Kiểm tra cập nhật
    result = check_for_updates(None, None, None, dummy_config, dummy_save_config, True)
    print(f"Update check result: {'Updates available' if result else 'No updates available or error occurred'}") 