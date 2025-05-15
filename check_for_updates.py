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
import tempfile
import concurrent.futures
import time

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
        # Nếu không có root và cần hiển thị giao diện, tạo một root tạm thời
        temp_root = None
        if root is None and show_notification:
            # Tạo một cửa sổ ẩn để hiển thị các hộp thoại
            temp_root = tk.Tk()
            temp_root.withdraw()  # Ẩn cửa sổ
            # Sử dụng temp_root khi root là None
            root = temp_root
        
        # Cập nhật UI nếu có status_label
        if status_label:
            status_label.config(text="Đang kiểm tra phiên bản mới...", style="StatusInfo.TLabel")
            if root:
                root.update_idletasks()
                
        print("Bắt đầu kiểm tra cập nhật...")
        
        # URL của GitHub API để kiểm tra phiên bản mới nhất
        github_api_url = "https://api.github.com/repos/TranMC/Import-score/releases/latest"
        
        # Gọi API để lấy thông tin phiên bản mới nhất với timeout được tăng lên và headers
        headers = {
            'User-Agent': 'QuanLyDiemHocSinh-App',
            'Accept': 'application/vnd.github.v3+json'
        }
        print(f"Đang kết nối tới GitHub API: {github_api_url}")
        response = requests.get(github_api_url, headers=headers, timeout=30)  # Tăng timeout lên 30 giây
        
        # Kiểm tra kết quả
        print(f"Mã phản hồi từ GitHub API: {response.status_code}")
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
                    if status_label:
                        status_label.config(text=f"Đang sử dụng phiên bản phát triển ({current_version})", style="StatusWarning.TLabel")
                    messagebox.showinfo("Phiên bản phát triển", f"Bạn đang sử dụng phiên bản phát triển ({current_version} dev version). Phiên bản phát hành mới nhất là {latest_version}.")
                else:
                    # Khi tự động kiểm tra, cập nhật thanh trạng thái
                    if status_label:
                        status_label.config(text=f"Phiên bản phát triển ({current_version})", style="StatusWarning.TLabel")
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
                                if status_label:
                                    status_label.config(text=f"Đang tải phiên bản mới {latest_version}...", 
                                                      style="StatusWarning.TLabel")
                                    if root:
                                        root.update_idletasks()
                                
                                # Tạo thư mục tạm để tải xuống
                                temp_dir = os.path.join(os.path.expanduser("~"), "Downloads")
                                os.makedirs(temp_dir, exist_ok=True)
                                
                                # Tạo tên file tạm
                                file_name = os.path.basename(download_url)
                                temp_file = os.path.join(temp_dir, file_name)
                                
                                # Lấy kích thước file
                                resp = requests.head(download_url, allow_redirects=True)
                                total_size = int(resp.headers.get('content-length', 0))
                                
                                # Tạo cửa sổ hiển thị tiến trình
                                progress_window = tk.Toplevel(root) if root else tk.Tk()
                                progress_window.title("Đang tải xuống")
                                progress_window.geometry("450x250")
                                if root:
                                    progress_window.transient(root)  # Đặt là cửa sổ con của root
                                
                                # Đặt cửa sổ ở giữa màn hình
                                window_width = 450
                                window_height = 350
                                # Xử lý trường hợp root là None
                                if root:
                                    screen_width = root.winfo_screenwidth()
                                    screen_height = root.winfo_screenheight()
                                else:
                                    # Sử dụng kích thước màn hình mặc định nếu không có root
                                    screen_width = 1366  # Giá trị mặc định
                                    screen_height = 768  # Giá trị mặc định
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
                                
                                # Label hiển thị tốc độ tải
                                speed_label = ttk.Label(progress_window_frame,
                                                        text="Tốc độ: 0.0 KB/s",
                                                        font=(config['ui']['font_family'], 10))
                                speed_label.pack(anchor="w", pady=2)
                                
                                # Label hiển thị thời gian còn lại
                                time_left_label = ttk.Label(progress_window_frame,
                                                           text="Thời gian còn lại: ??:??",
                                                           font=(config['ui']['font_family'], 10))
                                time_left_label.pack(anchor="w", pady=2)
                                
                                # Frame chứa nút điều khiển
                                control_frame = ttk.Frame(progress_window_frame)
                                control_frame.pack(fill="x", pady=10)
                                
                                # Tạo style cho nút hủy
                                style.configure("Cancel.TButton",
                                                foreground=config['ui']['theme'].get('danger', '#f44336'),
                                                background=config['ui']['theme'].get('background', '#ffffff'))
                                
                                # Thêm nút hủy tải xuống
                                cancel_button = ttk.Button(control_frame, text="Hủy tải xuống",
                                                         command=lambda: setattr(progress_window, 'cancelled', True),
                                                         style="Cancel.TButton")
                                cancel_button.pack(side="right", padx=10, pady=5)
                                
                                # Tạo nút hoàn thành (ẩn ban đầu)
                                complete_button = ttk.Button(control_frame, text="Đóng",
                                                           command=progress_window.destroy)
                                # Ban đầu ẩn nút hoàn thành
                                
                                # Biến để kiểm tra trạng thái hủy
                                progress_window.cancelled = False
                                
                                # Tạo dictionary chứa các UI elements để truyền vào callback
                                ui_elements = {
                                    'percent_label': percent_label,
                                    'download_detail_label': download_detail_label,
                                    'progress': progress,
                                    'window': progress_window,
                                    'speed_label': speed_label,
                                    'time_left_label': time_left_label
                                }
                                
                                # Theo dõi thời gian và bytes đã tải
                                start_time = datetime.now()
                                downloaded = 0
                                last_downloaded = 0
                                last_time = start_time
                                current_speed = 0  # Tốc độ tải hiện tại (bytes/s)
                                
                                # Hàm cập nhật tốc độ và thời gian còn lại
                                def update_speed_and_time():
                                    nonlocal last_downloaded, last_time, current_speed
                                    
                                    # Nếu đã hủy tải xuống, không cập nhật nữa
                                    if progress_window.cancelled:
                                        return
                                    
                                    current_time = datetime.now()
                                    time_diff = (current_time - last_time).total_seconds()
                                    
                                    # Cập nhật tốc độ mỗi 0.5 giây
                                    if time_diff >= 0.5:
                                        # Tính tốc độ tải hiện tại
                                        bytes_diff = downloaded - last_downloaded
                                        if time_diff > 0:
                                            current_speed = bytes_diff / time_diff
                                        
                                        # Định dạng tốc độ
                                        if current_speed < 1024:
                                            speed_text = f"{current_speed:.1f} B/s"
                                        elif current_speed < 1024 * 1024:
                                            speed_text = f"{current_speed/1024:.1f} KB/s"
                                        else:
                                            speed_text = f"{current_speed/1024/1024:.1f} MB/s"
                                            
                                        # Tính thời gian còn lại
                                        if current_speed > 0:
                                            remaining_bytes = total_size - downloaded
                                            remaining_seconds = remaining_bytes / current_speed
                                            
                                            if remaining_seconds < 60:
                                                time_text = f"{int(remaining_seconds)} giây"
                                            elif remaining_seconds < 3600:
                                                time_text = f"{int(remaining_seconds // 60)} phút {int(remaining_seconds % 60)} giây"
                                            else:
                                                time_text = f"{int(remaining_seconds // 3600)} giờ {int((remaining_seconds % 3600) // 60)} phút"
                                        else:
                                            time_text = "Đang tính toán..."
                                            
                                        try:
                                            speed_label.config(text=f"Tốc độ hiện tại: {speed_text}")
                                            time_left_label.config(text=f"Thời gian còn lại: {time_text}")
                                            progress_window.update_idletasks()
                                        except Exception as e:
                                            print(f"Lỗi khi cập nhật tốc độ và thời gian: {str(e)}")
                                            
                                        # Cập nhật dữ liệu cuối
                                        last_downloaded = downloaded
                                        last_time = current_time
                                    
                                    # Lên lịch cập nhật tiếp theo
                                    if not progress_window.cancelled:
                                        progress_window.after(250, update_speed_and_time)
                                
                                # Bắt đầu cập nhật tốc độ và thời gian
                                update_speed_and_time()
                                
                                # Tải xuống tệp
                                # Ưu tiên dùng đa luồng nếu có thể
                                def update_progress(percent, downloaded_bytes, total_bytes, ui_elements):
                                    if total_bytes > 0:
                                        try:
                                            nonlocal downloaded
                                            downloaded = downloaded_bytes
                                            
                                            # Kiểm tra nếu đã hủy tải xuống
                                            if hasattr(ui_elements['window'], 'cancelled') and ui_elements['window'].cancelled:
                                                raise Exception("Đã hủy tải xuống")
                                            
                                            ui_elements['percent_label'].config(text=f"{percent}%")
                                            ui_elements['download_detail_label'].config(text=f"({downloaded_bytes/1024/1024:.1f}/{total_bytes/1024/1024:.1f} MB)")
                                            ui_elements['progress'].config(value=percent)
                                            ui_elements['window'].update_idletasks()
                                        except Exception as e:
                                            print(f"Lỗi khi cập nhật UI: {str(e)}")
                                            # Trả về False để báo lỗi cho hàm tải xuống
                                            return False
                                    return True
                                
                                # Sử dụng đa luồng nếu có thể
                                use_multipart = download_file_multipart(download_url, temp_file, num_workers=4, chunk_size=1024*1024,
                                                                      progress_callback=lambda percent, downloaded_bytes, total: update_progress(percent, downloaded_bytes, total, ui_elements))
                                if use_multipart:
                                    # Nếu tải đa luồng thành công, cập nhật biến downloaded
                                    downloaded = total_size
                                
                                if not use_multipart:
                                    # Nếu không dùng được đa luồng, fallback về tải thường
                                    with requests.get(download_url, stream=True) as r:
                                        r.raise_for_status()
                                        total_size = int(r.headers.get('content-length', 0))
                                        downloaded = 0
                                        start_time = datetime.now()
                                        last_update_time = start_time
                                        last_downloaded = 0
                                        with open(temp_file, 'wb') as f:
                                            for chunk in r.iter_content(chunk_size=32768):
                                                # Kiểm tra nếu đã hủy tải xuống
                                                if progress_window.cancelled:
                                                    break
                                                    
                                                if chunk:
                                                    f.write(chunk)
                                                    downloaded += len(chunk)
                                                    current_time = datetime.now()
                                                    time_diff = (current_time - last_update_time).total_seconds()
                                                    if time_diff >= 0.5 and total_size > 0:
                                                        percent = int(100 * downloaded / total_size)
                                                        try:
                                                            progress["value"] = percent
                                                            percent_label.config(text=f"{percent}%")
                                                            download_detail_label.config(text=f"({downloaded/1024/1024:.1f}/{total_size/1024/1024:.1f} MB)")
                                                            progress_window.update_idletasks()
                                                        except Exception as e:
                                                            print(f"Lỗi cập nhật UI tải thường: {str(e)}")
                                                        last_update_time = current_time
                                                        last_downloaded = downloaded
                                
                                # Kiểm tra nếu đã hủy tải xuống
                                if progress_window.cancelled:
                                    try:
                                        # Xóa file tạm nếu đã hủy tải xuống
                                        if os.path.exists(temp_file):
                                            os.remove(temp_file)
                                        
                                        # Đóng cửa sổ tiến trình
                                        progress_window.destroy()
                                        
                                        # Thông báo cho người dùng
                                        messagebox.showinfo("Đã hủy tải xuống", "Quá trình tải xuống đã bị hủy.")
                                        
                                        if status_label:
                                            status_label.config(text="Đã hủy tải xuống", style="StatusWarning.TLabel")
                                            
                                        return False
                                    except Exception as e:
                                        print(f"Lỗi khi xử lý hủy tải xuống: {str(e)}")
                                        return False

                                # Hoàn thành tải xuống
                                elapsed_time = (datetime.now() - start_time).total_seconds()
                                avg_speed = downloaded / elapsed_time if elapsed_time > 0 else 0
                                
                                # Định dạng tốc độ trung bình
                                if avg_speed < 1024:
                                    avg_speed_text = f"{avg_speed:.1f} B/s"
                                elif avg_speed < 1024 * 1024:
                                    avg_speed_text = f"{avg_speed/1024:.1f} KB/s"
                                else:
                                    avg_speed_text = f"{avg_speed/1024/1024:.1f} MB/s"
                                
                                try:
                                    progress["value"] = 100
                                    percent_label.config(text="100%")
                                    download_detail_label.config(text=f"({total_size/1024/1024:.1f}/{total_size/1024/1024:.1f} MB)")
                                    speed_label.config(text=f"Tốc độ trung bình: {avg_speed_text}")
                                    time_left_label.config(text=f"Đã hoàn thành trong: {int(elapsed_time//60)} phút {int(elapsed_time%60)} giây")
                                    
                                    # Ẩn nút hủy và hiển thị nút hoàn thành
                                    cancel_button.pack_forget()
                                    complete_button.pack(side="right", padx=10, pady=5)
                                    
                                    # Đổi tiêu đề cửa sổ
                                    progress_window.title("Tải xuống hoàn tất")
                                    
                                    # Đóng cửa sổ tiến trình sau một khoảng thời gian (nếu người dùng không đóng)
                                    progress_window.after(30000, progress_window.destroy)
                                except Exception as e:
                                    print(f"Lỗi khi cập nhật UI cuối cùng: {str(e)}")
                                    try:
                                        # Cố gắng đóng cửa sổ ngay lập tức nếu có lỗi
                                        progress_window.destroy()
                                    except:
                                        pass
                                
                                # Hỏi người dùng có muốn cài đặt không
                                install_now = messagebox.askyesno("Tải xuống hoàn tất", 
                                                             f"Đã tải phiên bản {latest_version} về:\n{temp_file}\n\nBạn có muốn cài đặt ngay bây giờ không?")
                                
                                if install_now:
                                    # Chạy tệp cài đặt và thoát ứng dụng hiện tại
                                    os.startfile(temp_file)
                                    if root:
                                        root.after(1000, root.destroy)
                                    else:
                                        # Nếu không có root, thoát chương trình sau khi cài đặt
                                        import sys
                                        def exit_app():
                                            sys.exit(0)
                                        threading.Timer(1.0, exit_app).start()
                                
                                if status_label:
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
                else:
                    # Khi tự động kiểm tra, cập nhật thanh trạng thái
                    if status_label:
                        status_label.config(text=f"Có phiên bản mới! {latest_version}", style="StatusWarning.TLabel")
                        if root:
                            root.update_idletasks()
                
                if file_path and status_label:
                    status_label.config(text=f"Có phiên bản mới ({latest_version})! Đã tải file: {os.path.basename(file_path)}", style="StatusWarning.TLabel")
                return True
            else:
                # Không có phiên bản mới
                if show_notification:
                    if file_path and status_label:
                        status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusGood.TLabel")
                    messagebox.showinfo("Cập nhật", f"Bạn đang sử dụng phiên bản mới nhất ({current_version}).")
                else:
                    # Khi tự động kiểm tra, cập nhật thanh trạng thái
                    if status_label:
                        status_label.config(text=f"Đang sử dụng phiên bản mới nhất ({current_version})", style="StatusGood.TLabel")
                        if root:
                            root.update_idletasks()
                return False
        else:
            # Lỗi kết nối
            print(f"Lỗi kết nối đến GitHub API, mã: {response.status_code}")
            if show_notification:
                if file_path and status_label:
                    status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusError.TLabel")
                messagebox.showwarning("Lỗi kết nối", f"Không thể kết nối đến máy chủ GitHub. Mã lỗi: {response.status_code}")
            else:
                # Khi tự động kiểm tra, cập nhật thanh trạng thái
                if status_label:
                    status_label.config(text=f"Không thể kết nối đến máy chủ GitHub. Mã: {response.status_code}", style="StatusError.TLabel")
                    if root:
                        root.update_idletasks()
            return False
    
    except requests.exceptions.ConnectionError as e:
        # Lỗi kết nối mạng cụ thể
        print(f"Lỗi kết nối: {str(e)}")
        if show_notification:
            if file_path and status_label:
                status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusGood.TLabel")
            messagebox.showwarning("Lỗi kết nối", "Không thể kết nối đến máy chủ GitHub. Vui lòng kiểm tra kết nối mạng của bạn.")
        else:
            # Khi chạy tự động, vẫn cập nhật status_label nếu có
            if status_label:
                status_label.config(text="Không thể kết nối đến máy chủ GitHub", style="StatusError.TLabel")
                if root:
                    root.update_idletasks()
        return False
    
    except requests.exceptions.Timeout as e:
        # Lỗi timeout
        print(f"Lỗi timeout: {str(e)}")
        if show_notification:
            if file_path and status_label:
                status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusGood.TLabel")
            messagebox.showwarning("Lỗi kết nối", "Máy chủ GitHub phản hồi quá chậm. Vui lòng thử lại sau.")
        else:
            # Khi chạy tự động, vẫn cập nhật status_label nếu có
            if status_label:
                status_label.config(text="Máy chủ GitHub phản hồi quá chậm", style="StatusError.TLabel")
                if root:
                    root.update_idletasks()
        return False
        
    except requests.RequestException as e:
        # Xử lý lỗi request khác
        print(f"Lỗi request: {str(e)}")
        if show_notification:
            if file_path and status_label:
                status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusGood.TLabel")
            messagebox.showwarning("Lỗi kết nối", f"Không thể kiểm tra phiên bản mới: {str(e)}")
        else:
            # Khi tự động kiểm tra, cập nhật thanh trạng thái
            if status_label:
                status_label.config(text="Lỗi khi kiểm tra cập nhật", style="StatusError.TLabel")
                if root:
                    root.update_idletasks()
        return False
    
    except json.JSONDecodeError as e:
        # Lỗi khi parse JSON
        print(f"Lỗi parse JSON: {str(e)}")
        if show_notification:
            if file_path and status_label:
                status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusGood.TLabel")
            messagebox.showwarning("Lỗi dữ liệu", "Không thể đọc thông tin phiên bản từ máy chủ. Định dạng dữ liệu không hợp lệ.")
        else:
            # Khi tự động kiểm tra, cập nhật thanh trạng thái
            if status_label:
                status_label.config(text="Lỗi định dạng dữ liệu từ máy chủ", style="StatusError.TLabel")
                if root:
                    root.update_idletasks()
        return False
    
    except Exception as e:
        # Xử lý các lỗi khác
        print(f"Lỗi không xác định: {str(e)}")
        if show_notification:
            if file_path and status_label:
                status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusGood.TLabel")
            messagebox.showwarning("Lỗi", f"Đã xảy ra lỗi khi kiểm tra phiên bản mới: {str(e)}")
        else:
            # Khi tự động kiểm tra, cập nhật thanh trạng thái
            if status_label:
                status_label.config(text="Lỗi khi kiểm tra cập nhật", style="StatusError.TLabel")
                if root:
                    root.update_idletasks()
        return False
    finally:
        # Nếu đã tạo root tạm thời, hủy nó
        if 'temp_root' in locals() and temp_root:
            temp_root.destroy()

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
    
    # Cập nhật UI trước khi bắt đầu kiểm tra
    if status_label:
        status_label.config(text="Đang kiểm tra cập nhật...", style="StatusInfo.TLabel")
        if root:
            root.update_idletasks()
    
    # Tạo một thread riêng để kiểm tra cập nhật
    def check_update_thread():
        try:
            # Thay đổi tham số show_notification thành True để hiển thị thông báo cập nhật
            result = check_for_updates(root, status_label, file_path, config, save_config, True)
            print(f"Update check completed, result: {result}")
        except Exception as e:
            print(f"Error in update check thread: {str(e)}")
            # Cập nhật UI nếu có lỗi
            if status_label:
                status_label.config(text="Lỗi khi kiểm tra cập nhật", style="StatusError.TLabel")
                if root:
                    root.update_idletasks()
    
    threading.Thread(
        target=check_update_thread, 
        daemon=True
    ).start()
    print("Update check thread started")

def download_file_multipart(url, target_path, num_workers=4, chunk_size=1024*1024, progress_callback=None):
    """
    Tải một file sử dụng phương pháp đa luồng để tăng tốc độ
    
    Args:
        url: URL của file cần tải
        target_path: Đường dẫn đích để lưu file
        num_workers: Số luồng tải xuống song song
        chunk_size: Kích thước của mỗi phần (bytes)
        progress_callback: Hàm callback để cập nhật tiến trình và kiểm tra hủy
    
    Returns:
        bool: True nếu tải xuống thành công, False nếu có lỗi hoặc bị hủy
    """
    try:
        # Kiểm tra kích thước của file và nếu máy chủ hỗ trợ Range requests
        resp = requests.head(url, allow_redirects=True)
        supports_range = 'accept-ranges' in resp.headers and resp.headers['accept-ranges'] == 'bytes'
        
        if not supports_range:
            print("Máy chủ không hỗ trợ Range requests, sử dụng tải xuống thông thường")
            return False
            
        file_size = int(resp.headers.get('content-length', 0))
        if file_size == 0:
            print("Không thể xác định kích thước file, sử dụng tải xuống thông thường")
            return False
            
        # Tính toán các phần (chunks) để tải xuống
        parts = []
        for i in range(0, file_size, chunk_size):
            part = {
                'start': i,
                'end': min(i + chunk_size - 1, file_size - 1),
                'downloaded': False,
                'path': f"{target_path}.part{i // chunk_size}"
            }
            parts.append(part)
            
        # Tạo thư mục chứa file nếu chưa tồn tại
        os.makedirs(os.path.dirname(os.path.abspath(target_path)), exist_ok=True)
        
        # Kiểm tra xem đã có tệp nào được tải trước đó chưa
        for part in parts:
            if os.path.exists(part['path']):
                part_size = os.path.getsize(part['path'])
                if part_size == part['end'] - part['start'] + 1:
                    part['downloaded'] = True
        
        # Biến để kiểm soát việc hủy tải xuống
        cancel_download = False
        
        # Hàm để tải một phần
        def download_part(part):
            try:
                # Kiểm tra nếu đã hủy tải xuống
                if cancel_download:
                    return False
                    
                if part['downloaded']:
                    return True
                
                headers = {'Range': f'bytes={part["start"]}-{part["end"]}'}
                response = requests.get(url, headers=headers, stream=True, timeout=30)
                
                if response.status_code in [200, 206]:  # 200 OK or 206 Partial Content
                    with open(part['path'], 'wb') as f:
                        for chunk in response.iter_content(chunk_size=32768):
                            # Kiểm tra nếu đã hủy tải xuống
                            if cancel_download:
                                return False
                                
                            if chunk:
                                f.write(chunk)
                    part['downloaded'] = True
                    return True
                else:
                    print(f"Lỗi khi tải phần {part['start']}-{part['end']}: HTTP {response.status_code}")
                    return False
            except Exception as e:
                print(f"Lỗi khi tải phần {part['start']}-{part['end']}: {str(e)}")
                return False
        
        # Tải các phần chưa tải bằng ThreadPoolExecutor
        not_downloaded_parts = [part for part in parts if not part['downloaded']]
        total_parts = len(parts)
        downloaded_parts = total_parts - len(not_downloaded_parts)
        
        # Cập nhật tiến trình ban đầu nếu có tệp phần đã tải
        if progress_callback and downloaded_parts > 0:
            progress = downloaded_parts / total_parts
            result = progress_callback(int(progress * 100), downloaded_parts * chunk_size, file_size)
            # Kiểm tra nếu callback trả về False (người dùng đã hủy)
            if result is False:
                cancel_download = True
                return False
        
        # Tạo một biến đếm cho số phần đã tải
        download_counter = downloaded_parts
        
        # Sử dụng ThreadPoolExecutor để tải các phần còn lại
        with concurrent.futures.ThreadPoolExecutor(max_workers=num_workers) as executor:
            future_to_part = {executor.submit(download_part, part): part for part in not_downloaded_parts}
            
            for future in concurrent.futures.as_completed(future_to_part):
                part = future_to_part[future]
                try:
                    # Kiểm tra nếu đã hủy tải xuống
                    if cancel_download:
                        # Hủy tất cả các tác vụ đang chạy
                        for f in future_to_part:
                            f.cancel()
                        return False
                        
                    success = future.result()
                    if not success:
                        print(f"Không thể tải phần {part['start']}-{part['end']}")
                        cancel_download = True
                        return False
                    
                    # Cập nhật tiến trình nếu có callback
                    download_counter += 1
                    if progress_callback:
                        progress = download_counter / total_parts
                        bytes_downloaded = min(download_counter * chunk_size, file_size)
                        result = progress_callback(int(progress * 100), bytes_downloaded, file_size)
                        # Kiểm tra nếu callback trả về False (người dùng đã hủy)
                        if result is False:
                            cancel_download = True
                            # Hủy tất cả các tác vụ đang chạy
                            for f in future_to_part:
                                f.cancel()
                            return False
                        
                except Exception as e:
                    print(f"Lỗi khi xử lý phần {part['start']}-{part['end']}: {str(e)}")
                    cancel_download = True
                    return False
        
        # Nếu đã hủy tải xuống, không ghép các phần lại
        if cancel_download:
            # Xóa các file tạm
            for part in parts:
                try:
                    os.remove(part['path'])
                except:
                    pass
            return False
            
        # Ghép các phần lại với nhau
        with open(target_path, 'wb') as target_file:
            for part in parts:
                with open(part['path'], 'rb') as part_file:
                    target_file.write(part_file.read())
                    
        # Xóa các file tạm
        for part in parts:
            try:
                os.remove(part['path'])
            except:
                pass
                
        return True
    except Exception as e:
        print(f"Lỗi khi tải xuống đa luồng: {str(e)}")
        return False

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
