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
import time
import stat
import sys
# Các thư viện khác sẽ được import khi cần

# Các hằng số cho kênh cập nhật
UPDATE_CHANNEL_STABLE = "stable"
UPDATE_CHANNEL_DEV = "dev"

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
    # Lấy kênh cập nhật từ cấu hình, mặc định là stable
    update_channel = config.get('update_channel', UPDATE_CHANNEL_STABLE)
    
    # Cập nhật UI nếu có status_label
    if status_label:
        status_label.config(text="Đang kiểm tra phiên bản mới...", style="StatusInfo.TLabel")
        if root:
            root.update_idletasks()
        
    print(f"Bắt đầu kiểm tra cập nhật (kênh: {update_channel})...")
    
    # URL của GitHub API để kiểm tra phiên bản mới nhất
    github_api_url = "https://api.github.com/repos/TranMC/Import-score/releases"
    
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
        releases = response.json()
        
        # Lọc release theo kênh
        if update_channel == UPDATE_CHANNEL_STABLE:
            # Chỉ lấy các bản release chính thức (không phải pre-release)
            filtered_releases = [r for r in releases if not r.get('prerelease', False)]
        else:
            # Lấy tất cả các bản release kể cả pre-release
            filtered_releases = releases
        
        if not filtered_releases:
            if show_notification:
                messagebox.showinfo("Không tìm thấy phiên bản", f"Không tìm thấy phiên bản nào trong kênh {update_channel}.")
            return False
            
        # Lấy phiên bản mới nhất từ danh sách đã lọc
        release_info = filtered_releases[0]
        
        # Lấy phiên bản mới nhất
        latest_version = release_info.get('tag_name', '').lstrip('v')
        # Lấy phiên bản hiện tại từ version.json
        import version_utils
        version_info = version_utils.load_version_info()
        current_version = version_info.get('version', '0.0.0')
        is_dev = version_info.get('is_dev', False)
        release_channel = version_info.get('release_channel', UPDATE_CHANNEL_STABLE)
        
        # Kiểm tra nếu phiên bản hiện tại cao hơn phiên bản mới nhất, đánh dấu là dev
        if current_version > latest_version and not is_dev:
            # Thông báo cho người dùng biết họ đang sử dụng phiên bản phát triển
            if show_notification:
                if status_label:
                    status_label.config(text=f"Đang sử dụng phiên bản phát triển ({current_version})", style="StatusWarning.TLabel")
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
                    f"Phiên bản hiện tại: {current_version}\n"
                    f"Kênh cập nhật: {update_channel}\n\n"
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
                            
                            # Tạo cửa sổ hiển thị tiến trình
                            progress_window = tk.Toplevel(root) if root else tk.Tk()
                            progress_window.title("Đang tải xuống")
                            progress_window.geometry("450x310")
                            if root:
                                progress_window.transient(root)  # Đặt là cửa sổ con của root
                            
                            # Đặt cửa sổ ở giữa màn hình
                            window_width = 450
                            window_height = 310
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
                            
                            # Biến để kiểm tra trạng thái hủy
                            progress_window.cancelled = False
                            
                            # Hàm xử lý khi nhấn nút hủy
                            def on_cancel_download():
                                # Đổi tiêu đề cửa sổ ngay lập tức
                                progress_window.title("Đang hủy tải xuống...")
                                # Hiển thị thông báo ngay trong UI
                                percent_label.config(text="Đang hủy...")
                                # Vô hiệu hóa nút hủy để tránh nhấn nhiều lần
                                cancel_button.config(state="disabled")
                                # Đặt cờ hủy
                                progress_window.cancelled = True
                                
                                # Cập nhật UI ngay lập tức
                                progress_window.update()
                            
                            # Thêm nút hủy tải xuống
                            cancel_button = ttk.Button(control_frame, text="Hủy tải xuống",
                                                     command=on_cancel_download,
                                                     style="Cancel.TButton")
                            cancel_button.pack(side="right", padx=10, pady=5)
                            
                            # Tạo nút hoàn thành (ẩn ban đầu)
                            complete_button = ttk.Button(control_frame, text="Đóng",
                                                       command=progress_window.destroy)
                            # Ban đầu ẩn nút hoàn thành
                            
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
                                    try:
                                        if current_speed > 0 and total_size > 0:
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
                                    try:
                                        progress_window.after(250, update_speed_and_time)
                                    except Exception as e:
                                        print(f"Lỗi khi lên lịch cập nhật tốc độ: {str(e)}")
                            
                            # Bắt đầu cập nhật tốc độ và thời gian
                            update_speed_and_time()
                            
                            # Luôn sử dụng tải đơn luồng
                            with requests.get(download_url, stream=True) as r:
                                r.raise_for_status()
                                total_size = int(r.headers.get('content-length', 0))
                                with open(temp_file, 'wb') as f:
                                    for chunk in r.iter_content(chunk_size=8192):
                                        # Kiểm tra nếu đã hủy tải xuống
                                        if progress_window.cancelled:
                                            print("Đã hủy tải xuống đơn luồng")
                                            break
                                                
                                        if chunk:
                                            f.write(chunk)
                                            downloaded += len(chunk)
                                            
                                            # Cập nhật tiến trình tải mỗi 32KB (4 chunks) để UI mượt hơn
                                            if downloaded % 32768 == 0 and total_size > 0:
                                                percent = int(100 * downloaded / total_size)
                                                try:
                                                    progress["value"] = percent
                                                    percent_label.config(text=f"{percent}%")
                                                    download_detail_label.config(text=f"({downloaded/1024/1024:.1f}/{total_size/1024/1024:.1f} MB)")
                                                except Exception as e:
                                                    print(f"Lỗi cập nhật UI tải xuống: {str(e)}")
                            
                            # Kiểm tra nếu đã hủy tải xuống
                            if progress_window.cancelled:
                                try:
                                    # Xóa file tạm nếu đã hủy tải xuống
                                    if os.path.exists(temp_file):
                                        # Xóa file triệt để (không vào thùng rác)
                                        force_delete_file(temp_file)
                                    
                                    # Đóng cửa sổ tiến trình
                                    progress_window.destroy()
                                    
                                    # Thông báo cho người dùng
                                    messagebox.showinfo("Đã hủy tải xuống", "Quá trình tải xuống đã bị hủy và file tạm đã được xóa.")
                                    
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
                print(f"Bạn đang sử dụng phiên bản mới nhất ({current_version}).")
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
            print(f"Lỗi kết nối đến GitHub API, mã: {response.status_code}")
        return False

def show_update_channel_dialog(root, config, save_config, callback=None):
    """
    Hiển thị hộp thoại để người dùng chọn kênh cập nhật
    
    Args:
        root: Cửa sổ chính của ứng dụng
        config: Dictionary chứa cấu hình ứng dụng
        save_config: Hàm lưu cấu hình
        callback: Hàm callback để gọi sau khi cập nhật cấu hình
        
    Returns:
        str: Kênh cập nhật đã chọn
    """
    dialog = tk.Toplevel(root)
    dialog.title("Chọn kênh cập nhật")
    dialog.geometry("400x320")
    dialog.transient(root)
    dialog.resizable(False, False)
    
    # Đặt cửa sổ ở giữa màn hình chính
    position_x = max(0, int(root.winfo_x() + (root.winfo_width() / 2) - 200))
    position_y = max(0, int(root.winfo_y() + (root.winfo_height() / 2) - 160))
    dialog.geometry(f"400x320+{position_x}+{position_y}")
    
    # Thiết lập frame chính
    main_frame = ttk.Frame(dialog, padding=15)
    main_frame.pack(fill="both", expand=True)
    
    # Tiêu đề
    ttk.Label(main_frame, 
              text="Chọn kênh cập nhật", 
              font=(config['ui']['font_family'], 12, 'bold'),
              foreground=config['ui']['theme']['primary']).pack(pady=(0, 10))
    
    # Giá trị kênh hiện tại
    current_channel = config.get('update_channel', UPDATE_CHANNEL_STABLE)
    channel_var = tk.StringVar(value=current_channel)
    
    # Frame cho các tùy chọn
    options_frame = ttk.Frame(main_frame)
    options_frame.pack(fill="x", pady=10)
    
    # Các tham số chung cho labels
    label_font = (config['ui']['font_family'], 9)
    label_fg = config['ui']['theme'].get('text_secondary', '#757575')
    
    # Tùy chọn kênh ổn định
    stable_frame = ttk.Frame(options_frame)
    stable_frame.pack(fill="x", pady=5)
    
    ttk.Radiobutton(stable_frame, 
                   text="Kênh ổn định", 
                   variable=channel_var, 
                   value=UPDATE_CHANNEL_STABLE).pack(side="left")
    
    ttk.Label(stable_frame,
             text="Chỉ nhận các bản cập nhật chính thức",
             font=label_font,
             foreground=label_fg).pack(side="left", padx=10)
    
    # Tùy chọn kênh phát triển
    dev_frame = ttk.Frame(options_frame)
    dev_frame.pack(fill="x", pady=5)
    
    ttk.Radiobutton(dev_frame, 
                    text="Kênh phát triển", 
                    variable=channel_var, 
                    value=UPDATE_CHANNEL_DEV).pack(side="left")
    
    ttk.Label(dev_frame,
             text="Nhận cả bản pre-release và các tính năng mới",
             font=label_font,
             foreground=label_fg).pack(side="left", padx=10)
    
    # Mô tả thêm
    description_text = (
        "Kênh ổn định cung cấp các phiên bản đã được kiểm tra kỹ lưỡng, phù hợp "
        "với người dùng thông thường.\n\n"
        "Kênh phát triển cung cấp các phiên bản mới nhất với tính năng mới nhất, "
        "nhưng có thể chứa lỗi. Phù hợp với người dùng muốn trải nghiệm sớm."
    )
    
    ttk.Label(main_frame,
             text=description_text,
             font=label_font,
             foreground=label_fg,
             wraplength=370,
             justify="left").pack(fill="x", pady=10)
    
    # Frame cho các nút
    button_frame = ttk.Frame(main_frame, height=50)
    button_frame.pack(side="bottom", fill="x", pady=10)
    button_frame.pack_propagate(False)  # Giữ kích thước cố định
    
    # Hàm lưu thay đổi
    def save_channel():
        selected_channel = channel_var.get()
        config['update_channel'] = selected_channel
        save_config()
        
        messagebox.showinfo("Thành công", f"Đã lưu cấu hình kênh cập nhật: {selected_channel}")
        dialog.destroy()
        
        # Gọi callback nếu có
        if callback:
            callback(selected_channel)
    
    # Nút lưu và hủy
    ttk.Button(button_frame, 
               text="Lưu thay đổi", 
               command=save_channel,
               width=15).pack(side="right", padx=5)
    
    ttk.Button(button_frame, 
               text="Hủy bỏ", 
               command=dialog.destroy,
               width=10).pack(side="right", padx=5)
    
    # Đặt focus và chế độ modal
    dialog.focus_set()
    dialog.grab_set()
    dialog.wait_window()
    
    return channel_var.get()

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
    # Cập nhật UI trước khi bắt đầu kiểm tra
    if status_label:
        status_label.config(text="Đang kiểm tra cập nhật...", style="StatusInfo.TLabel")
        root and root.update_idletasks()
    
    # Tạo một thread riêng để kiểm tra cập nhật
    def check_update_thread():
        try:
            result = check_for_updates(root, status_label, file_path, config, save_config, True)
            print(f"Kết quả kiểm tra cập nhật: {result}")
        except Exception as e:
            print(f"Lỗi trong thread kiểm tra cập nhật: {str(e)}")
            if status_label:
                status_label.config(text="Lỗi khi kiểm tra cập nhật", style="StatusError.TLabel")
                root and root.update_idletasks()
    
    # Tạo và khởi động thread
    update_thread = threading.Thread(
        target=check_update_thread,
        daemon=True
    )
    update_thread.start()

def download_file_multipart(url, target_path, num_workers=4, chunk_size=1024*1024, progress_callback=None):
    """
    Chức năng tải đa luồng đã bị loại bỏ - luôn trả về False để sử dụng tải đơn luồng
    """
    print("Chức năng tải đa luồng đã bị vô hiệu hóa, sử dụng tải đơn luồng để đảm bảo ổn định")
    return False

def force_delete_file(file_path):
    """
    Xóa file một cách triệt để, không vào thùng rác
    """
    try:
        # Đảm bảo file không ở chế độ chỉ đọc
        if os.path.exists(file_path):
            try:
                os.chmod(file_path, stat.S_IWRITE)
            except:
                pass
            
            # Xóa file trực tiếp
            os.remove(file_path)
            return True
    except Exception as e:
        print(f"Không thể xóa file {file_path}: {str(e)}")
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
