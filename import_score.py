import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import pandas as pd
from pandas import CategoricalDtype
import os
import sys
import json
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
from datetime import datetime
import traceback
import base64
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
import threading
import requests

# Import các module mới
import version_utils
import themes
import ui_utils

# Tạo class ToolTip để hiển thị gợi ý khi di chuột qua các nút
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)
        
    def show_tooltip(self, event=None):
        # Lấy vị trí của widget
        x = self.widget.winfo_rootx() + 25
        y = self.widget.winfo_rooty() + 25
        
        # Đóng tooltip cũ nếu có
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None
        
        # Tạo cửa sổ toplevel
        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        
        label = ttk.Label(self.tooltip, text=self.text, 
                         background="#ffffe0", relief="solid", borderwidth=1,
                         font=("Segoe UI", 9, "normal"))
        label.pack(ipadx=5, ipady=2)
        
    def hide_tooltip(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

# Undo/Redo Manager
class UndoManager:
    """Quản lý undo/redo cho các thao tác chỉnh sửa dữ liệu"""
    def __init__(self, max_history=50):
        self.undo_stack = []  # Stack lưu các state trước đó
        self.redo_stack = []  # Stack lưu các state đã undo
        self.max_history = max_history
    
    def push_state(self, state, action_name):
        """Lưu state hiện tại vào undo stack
        
        Args:
            state: DataFrame copy hoặc dict chứa thông tin state
            action_name: Tên hành động (VD: "Nhập điểm cho Nguyễn Văn A")
        """
        if state is not None:
            self.undo_stack.append({
                'state': state.copy() if hasattr(state, 'copy') else state,
                'action': action_name,
                'timestamp': datetime.now()
            })
            # Giới hạn kích thước stack
            if len(self.undo_stack) > self.max_history:
                self.undo_stack.pop(0)
            # Clear redo stack khi có thay đổi mới
            self.redo_stack.clear()
    
    def undo(self):
        """Hoàn tác thao tác cuối cùng
        
        Returns:
            tuple: (state, action_name) hoặc (None, None) nếu không thể undo
        """
        if self.can_undo():
            current = self.undo_stack.pop()
            self.redo_stack.append(current)
            return current['state'], current['action']
        return None, None
    
    def redo(self):
        """Làm lại thao tác đã hoàn tác
        
        Returns:
            tuple: (state, action_name) hoặc (None, None) nếu không thể redo
        """
        if self.can_redo():
            current = self.redo_stack.pop()
            self.undo_stack.append(current)
            return current['state'], current['action']
        return None, None
    
    def can_undo(self):
        """Kiểm tra có thể undo không"""
        return len(self.undo_stack) > 0
    
    def can_redo(self):
        """Kiểm tra có thể redo không"""
        return len(self.redo_stack) > 0
    
    def clear(self):
        """Xóa toàn bộ history"""
        self.undo_stack.clear()
        self.redo_stack.clear()
    
    def get_undo_info(self):
        """Lấy thông tin về thao tác có thể undo"""
        if self.can_undo():
            return self.undo_stack[-1]['action']
        return None
    
    def get_redo_info(self):
        """Lấy thông tin về thao tác có thể redo"""
        if self.can_redo():
            return self.redo_stack[-1]['action']
        return None

# Toast Notification System
class ToastNotification:
    """Hệ thống thông báo toast hiện đại"""
    active_toasts = []
    
    @staticmethod
    def show(message, toast_type="info", duration=3000):
        """
        Hiển thị toast notification
        
        Args:
            message (str): Nội dung thông báo
            toast_type (str): Loại thông báo: 'success', 'error', 'warning', 'info'
            duration (int): Thời gian hiển thị (ms), 0 = không tự động đóng
        """
        # Tạo cửa sổ toast
        toast = tk.Toplevel(root)
        toast.wm_overrideredirect(True)
        toast.attributes('-topmost', True)
        
        # Màu sắc theo loại
        colors = {
            'success': {'bg': '#10B981', 'fg': 'white', 'icon': '✓'},
            'error': {'bg': '#EF4444', 'fg': 'white', 'icon': '✕'},
            'warning': {'bg': '#F59E0B', 'fg': 'white', 'icon': '⚠'},
            'info': {'bg': '#3B82F6', 'fg': 'white', 'icon': 'ℹ'}
        }
        
        style_config = colors.get(toast_type, colors['info'])
        
        # Frame chính với padding
        main_frame = tk.Frame(toast, bg=style_config['bg'], padx=20, pady=12)
        main_frame.pack(fill='both', expand=True)
        
        # Icon + Message
        content_frame = tk.Frame(main_frame, bg=style_config['bg'])
        content_frame.pack(fill='both', expand=True)
        
        tk.Label(content_frame, 
                text=style_config['icon'], 
                font=('Segoe UI', 16),
                bg=style_config['bg'],
                fg=style_config['fg']).pack(side='left', padx=(0, 10))
        
        tk.Label(content_frame, 
                text=message,
                font=('Segoe UI', 10, 'bold'),
                bg=style_config['bg'],
                fg=style_config['fg'],
                wraplength=300,
                justify='left').pack(side='left', fill='both', expand=True)
        
        # Tính toán vị trí (góc phải dưới màn hình)
        toast.update_idletasks()
        toast_width = toast.winfo_width()
        toast_height = toast.winfo_height()
        
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        
        # Cleanup các toast đã bị destroyed
        ToastNotification.active_toasts = [t for t in ToastNotification.active_toasts if t.winfo_exists()]
        
        # Stack toasts từ dưới lên
        y_offset = 0
        for t in ToastNotification.active_toasts:
            try:
                if t.winfo_exists():
                    y_offset += t.winfo_height() + 10
            except:
                pass
        
        x = screen_width - toast_width - 20
        y = screen_height - toast_height - 60 - y_offset
        
        toast.geometry(f"+{x}+{y}")
        
        # Thêm vào danh sách active
        ToastNotification.active_toasts.append(toast)
        
        # Animation fade in
        toast.attributes('-alpha', 0.0)
        for i in range(1, 11):
            toast.attributes('-alpha', i / 10)
            toast.update()
            root.after(20)
        
        def close_toast():
            # Xóa khỏi danh sách trước
            if toast in ToastNotification.active_toasts:
                ToastNotification.active_toasts.remove(toast)
            
            # Animation fade out
            try:
                for i in range(10, 0, -1):
                    if toast.winfo_exists():
                        toast.attributes('-alpha', i / 10)
                        toast.update()
                        root.after(20)
                    else:
                        break
            except:
                pass
            
            # Destroy toast
            try:
                if toast.winfo_exists():
                    toast.destroy()
            except:
                pass
        
        # Tự động đóng sau duration nếu > 0
        if duration > 0:
            toast.after(duration, close_toast)
        
        # Click để đóng sớm
        toast.bind('<Button-1>', lambda e: close_toast())
        
        return toast

def get_config_path():
    """Lấy đường dẫn đến file config trong thư mục AppData của Windows"""
    # Lấy thư mục AppData/Local
    appdata_local = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'StudentScoreImport')
    
    # Tạo thư mục nếu chưa tồn tại
    if not os.path.exists(appdata_local):
        try:
            os.makedirs(appdata_local)
        except Exception as e:
            print(f"Không thể tạo thư mục config: {str(e)}")
            # Fallback về thư mục hiện tại nếu không tạo được AppData
            return os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_config.json")
    
    return os.path.join(appdata_local, 'app_config.json')

def load_bundled_config():
    """Tải config mẫu từ file được bundle với ứng dụng"""
    try:
        # Khi chạy từ PyInstaller, file được bundle nằm trong sys._MEIPASS
        if getattr(sys, 'frozen', False):
            bundle_dir = sys._MEIPASS
        else:
            bundle_dir = os.path.dirname(os.path.abspath(__file__))
        
        bundled_config_path = os.path.join(bundle_dir, 'app_config.json')
        
        if os.path.exists(bundled_config_path):
            with open(bundled_config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        print(f"Không thể đọc config mẫu: {str(e)}")
    
    return None

# Cấu hình cơ bản - tải từ file thay vì hardcode
def load_config():
    """Tải cấu hình từ file JSON"""
    try:
        config_path = get_config_path()
        print(f"Đường dẫn config: {config_path}")  # Debug
        
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print("File config chưa tồn tại, tạo config mặc định")
        
        # Thử tải config mẫu từ bundle trước
        bundled_config = load_bundled_config()
        if bundled_config:
            print("Sử dụng config mẫu từ bundle")
            default_config = bundled_config
        else:
            print("Sử dụng config mặc định hardcoded")
            # Trả về cấu hình mặc định và lưu luôn
            default_config = {
                'columns': {
                    'name': 'Tên Học Sinh',
                    'exam_code': 'Mã Đề',
                    'score': 'Điểm'
                },
                'max_questions': 40,
                'score_per_question': 0.25,
                'exam_codes': ['701', '702', '703', '704'],
                'shortcuts': {
                    'search': '<Control-f>',
                    'direct_score': '<Control-g>',
                    'undo': '<Control-z>'
                },
                'security': {
                    'encrypt_backups': False,
                    'encrypt_sensitive_data': False,
                    'password_protect_app': False,
                    'auto_lock_timeout_minutes': 0
                },
                'ui': {
                    'font_family': 'Segoe UI',
                    'font_size': {
                        'normal': 11,
                        'heading': 12,
                        'button': 11
                    },
                    'padding': {
                        'frame': 10,
                        'widget': 5
                    },
                    'min_width': {
                        'button': 120,
                        'entry': 150,
                        'combobox': 100
                    },
                    'theme': themes.LIGHT_THEME,
                    'rounded_corners': 4,
                    'dark_mode': False,
                    'responsive': {
                        'initial_width': 1400,
                        'initial_height': 800,
                        'min_width': 1280,
                        'min_height': 720
                    }
                }
            }
        
        # Lưu config mặc định ngay lập tức
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, ensure_ascii=False, indent=4)
            print(f"Đã tạo file config mới tại: {config_path}")
        except Exception as e:
            print(f"Không thể lưu config mặc định: {str(e)}")
        
        return default_config
    except Exception as e:
        print(f"Lỗi khi đọc cấu hình: {str(e)}")
        # Thử tải config mẫu từ bundle
        bundled_config = load_bundled_config()
        if bundled_config:
            return bundled_config
        
        # Trả về cấu hình mặc định
        return {
            'columns': {
                'name': 'Tên Học Sinh',
                'exam_code': 'Mã Đề',
                'score': 'Điểm'
            },
            'max_questions': 40,
            'score_per_question': 0.25,
            'exam_codes': ['701', '702', '703', '704'],
            'shortcuts': {
                'search': '<Control-f>',
                'direct_score': '<Control-g>',
                'undo': '<Control-z>'
            },
            'security': {
                'encrypt_backups': False,
                'encrypt_sensitive_data': False,
                'password_protect_app': False,
                'auto_lock_timeout_minutes': 0
            },
            'ui': {
                'font_family': 'Segoe UI',
                'font_size': {
                    'normal': 11,
                    'heading': 12,
                    'button': 11
                },
                'padding': {
                    'frame': 10,
                    'widget': 5
                },
                'min_width': {
                    'button': 120,
                    'entry': 150,
                    'combobox': 100
                },
                'theme': themes.LIGHT_THEME,
                'rounded_corners': 4,
                'dark_mode': False,
                'responsive': {
                    'initial_width': 1400,
                    'initial_height': 800,
                    'min_width': 1280,
                    'min_height': 720
                }
            }
        }

config = load_config()

def save_config():
    """Lưu cấu hình ra file JSON trong AppData"""
    try:
        config_path = get_config_path()
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
        print(f"Đã lưu config tại: {config_path}")  # Debug
    except Exception as e:
        print(f"Lỗi khi lưu cấu hình: {str(e)}")
        messagebox.showwarning("Cảnh báo", f"Không thể lưu cấu hình: {str(e)}")

def encrypt_data(data, password):
    """
    Mã hóa dữ liệu sử dụng Fernet với password-based key derivation
    
    Args:
        data (str): Dữ liệu cần mã hóa (thường là JSON)
        password (str): Mật khẩu để tạo khóa
        
    Returns:
        tuple: (encrypted_data, salt)
    """
    # Tạo salt ngẫu nhiên
    salt = os.urandom(16)
    
    # Tạo key từ password và salt
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=100000,
    )
    
    key = base64.urlsafe_b64encode(kdf.derive(password.encode()))
    
    # Mã hóa dữ liệu
    fernet = Fernet(key)
    encrypted_data = fernet.encrypt(data.encode())
    
    return encrypted_data, salt

def decrypt_data(encrypted_data, password, salt):
    """
    Giải mã dữ liệu đã được mã hóa bằng encrypt_data
    
    Args:
        encrypted_data (bytes): Dữ liệu đã mã hóa
        password (str): Mật khẩu dùng để tạo khóa
        salt (bytes): Salt đã dùng khi mã hóa
        
    Returns:
        str: Dữ liệu đã giải mã
    """
    # Tạo lại key từ password và salt
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=100000,
    )
    
    key = base64.urlsafe_b64encode(kdf.derive(password.encode()))
    
    # Giải mã dữ liệu
    fernet = Fernet(key)
    decrypted_data = fernet.decrypt(encrypted_data).decode('utf-8')
    
    return decrypted_data

# Biến toàn cục
root = tk.Tk()
root.title("Quản lí điểm học sinh")
df = None
file_path = None
undo_stack = []
undo_manager = UndoManager(max_history=50)  # Quản lý undo/redo
search_timer_id = None  # Thêm biến để theo dõi timer
search_index = None    # Thêm biến để lưu search index
last_activity_time = None  # Thêm biến để theo dõi thời gian hoạt động cuối cùng
lock_window = None  # Thêm biến để theo dõi cửa sổ khóa
lock_time = None  # Thêm biến để theo dõi thời gian khóa



def ensure_proper_dtypes(df_input):
    """
    Đảm bảo các kiểu dữ liệu đúng cho các cột quan trọng, đặc biệt là xử lý cột Categorical
    để tránh lỗi khi gán giá trị mới
    """
    if df_input is None:
        return pd.DataFrame()  # Trả về DataFrame rỗng thay vì None
        
    df_copy = df_input.copy()
    
    # Xử lý cột Điểm
    if 'Điểm' in df_copy.columns:
        # Chuyển đổi tất cả các giá trị sang dạng số, thay thế giá trị không hợp lệ bằng NaN
        try:
            df_copy['Điểm'] = pd.to_numeric(df_copy['Điểm'], errors='coerce')
        except Exception as e:
            print(f"Lỗi khi chuyển đổi cột Điểm: {str(e)}")
            df_copy['Điểm'] = df_copy['Điểm'].astype('object')
    
    # Xử lý cột Mã đề
    if 'Mã đề' in df_copy.columns:
        # Đảm bảo Mã đề là dạng chuỗi hoặc null
        try:
            # Chuyển từ các kiểu dữ liệu khác thành chuỗi
            df_copy['Mã đề'] = df_copy['Mã đề'].astype(str)
            # Thay thế 'nan' bằng chuỗi rỗng
            df_copy.loc[df_copy['Mã đề'] == 'nan', 'Mã đề'] = ''
        except Exception as e:
            print(f"Lỗi khi chuyển đổi cột Mã đề: {str(e)}")
            df_copy['Mã đề'] = df_copy['Mã đề'].astype('object')
    
    # Đảm bảo cột tên học sinh
    name_col = config['columns']['name']
    if name_col in df_copy.columns:
        try:
            # Chuyển sang kiểu chuỗi
            df_copy[name_col] = df_copy[name_col].fillna('')
            df_copy[name_col] = df_copy[name_col].astype(str)
        except Exception as e:
            print(f"Lỗi khi chuyển đổi cột {name_col}: {str(e)}")
    
    return df_copy



def lock_application():
    """Khóa ứng dụng và yêu cầu mật khẩu để mở khóa"""
    global lock_time, lock_window
    
    # Nếu không bật tính năng bảo vệ bằng mật khẩu, không làm gì cả
    if not config.get('security', {}).get('password_protect_app', False):
        return
    
    # Nếu chưa có mật khẩu, không thể khóa
    if not config.get('security', {}).get('password'):
        return
    
    # Nếu đã có cửa sổ khóa đang hiển thị, không tạo mới
    if 'lock_window' in globals() and lock_window is not None:
        try:
            # Kiểm tra xem cửa sổ còn tồn tại không
            lock_window.winfo_exists()
            return
        except:
            pass  # Nếu có lỗi, nghĩa là cửa sổ không còn tồn tại
    
    # Lưu thời gian khóa
    lock_time = datetime.now()
    
    # Vô hiệu hóa tạm thời các widget chính
    for widget in root.winfo_children():
        try:
            widget.configure(state="disabled")
        except:
            pass  # Bỏ qua lỗi nếu widget không hỗ trợ state
    
    # Tạo cửa sổ khóa
    lock_window = tk.Toplevel(root)
    lock_window.title("Ứng dụng đã bị khóa")
    lock_window.geometry("400x250")
    lock_window.transient(root)
    lock_window.grab_set()  # Đặt là modal
    lock_window.protocol("WM_DELETE_WINDOW", lambda: None)  # Ngăn chặn việc đóng cửa sổ
    lock_window.resizable(False, False)
    
    # Đặt cửa sổ ở giữa màn hình
    window_width = 400
    window_height = 250
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_x = int(screen_width/2 - window_width/2)
    position_y = int(screen_height/2 - window_height/2)
    lock_window.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")
    
    # Thêm biểu tượng khóa (có thể thay bằng hình ảnh)
    ttk.Label(lock_window, text="🔒", font=("Segoe UI", 48)).pack(pady=10)
    
    # Thông báo
    ttk.Label(lock_window, text="Ứng dụng đã bị khóa để bảo vệ dữ liệu", 
             font=(config['ui']['font_family'], 12, 'bold')).pack(pady=5)
    
    # Hiển thị thời gian khóa
    lock_time_str = lock_time.strftime("%H:%M:%S")
    ttk.Label(lock_window, text=f"Thời gian khóa: {lock_time_str}", 
             font=(config['ui']['font_family'], 10)).pack(pady=5)
    
    # Frame nhập mật khẩu
    password_frame = ttk.Frame(lock_window)
    password_frame.pack(pady=10)
    
    ttk.Label(password_frame, text="Mật khẩu:").pack(side="left", padx=5)
    password_entry = ttk.Entry(password_frame, show="*", width=20)
    password_entry.pack(side="left", padx=5)
    password_entry.focus_set()  # Đặt focus vào ô nhập mật khẩu
    
    def unlock():
        """Kiểm tra mật khẩu và mở khóa ứng dụng"""
        if password_entry.get() == config.get('security', {}).get('password', ''):
            # Kích hoạt lại các widget
            for widget in root.winfo_children():
                try:
                    widget.configure(state="normal")
                except:
                    pass  # Bỏ qua lỗi nếu widget không hỗ trợ state
            
            # Cập nhật thời gian hoạt động
            update_activity_time()
            
            # Đóng cửa sổ khóa
            lock_window.destroy()
        else:
            messagebox.showerror("Lỗi", "Mật khẩu không đúng")
            password_entry.delete(0, tk.END)
            password_entry.focus_set()
    
    # Thêm phím Enter để mở khóa
    password_entry.bind("<Return>", lambda event: unlock())
    
    # Nút mở khóa
    ttk.Button(lock_window, text="Mở khóa", command=unlock, width=15).pack(pady=10)

def check_auto_lock():
    """Kiểm tra và tự động khóa ứng dụng sau một khoảng thời gian không hoạt động"""
    # Nếu không bật tính năng bảo vệ bằng mật khẩu, không làm gì cả
    if not config.get('security', {}).get('password_protect_app', False):
        root.after(60000, check_auto_lock)  # Vẫn kiểm tra lại sau 1 phút
        return
    
    # Nếu chưa có mật khẩu, không thể khóa
    if not config.get('security', {}).get('password'):
        root.after(60000, check_auto_lock)  # Vẫn kiểm tra lại sau 1 phút
        return
    
    global last_activity_time
    
    # Nếu chưa có hoạt động nào, khởi tạo thời gian hoạt động cuối cùng
    if 'last_activity_time' not in globals() or last_activity_time is None:
        last_activity_time = datetime.now()
    
    # Nếu đã có cửa sổ khóa, không cần kiểm tra thêm
    if 'lock_window' in globals() and lock_window is not None:
        try:
            if lock_window.winfo_exists():
                root.after(60000, check_auto_lock)  # Vẫn kiểm tra lại sau 1 phút
                return
        except:
            pass  # Nếu có lỗi, nghĩa là cửa sổ không còn tồn tại
    
    # Tính thời gian không hoạt động (phút)
    idle_time = (datetime.now() - last_activity_time).total_seconds() / 60
    auto_lock_timeout = config.get('security', {}).get('auto_lock_timeout_minutes', 30)
    
    # Nếu thời gian không hoạt động vượt quá thời gian cài đặt, khóa ứng dụng
    if idle_time >= auto_lock_timeout:
        lock_application()
    
    # Đặt lịch kiểm tra lại sau 1 phút
    root.after(60000, check_auto_lock)

def update_activity_time(event=None):
    """Cập nhật thời gian hoạt động cuối cùng"""
    global last_activity_time
    last_activity_time = datetime.now()

def load_excel_lazily(file_path, chunk_size=1000, header_row=None):
    """
    Đọc file Excel với kỹ thuật lazy loading để xử lý dữ liệu lớn
    
    Args:
        file_path (str): Đường dẫn tới file Excel
        chunk_size (int): Kích thước mỗi chunk khi đọc dữ liệu
        header_row (int): Dòng chứa tiêu đề, None nếu dòng đầu tiên
        
    Returns:
        pd.DataFrame: DataFrame hoàn chỉnh sau khi đọc
    """
    from pandas.io.excel._openpyxl import OpenpyxlReader
    
    # Trạng thái tiến trình
    status_label.config(text=f"Đang lập kế hoạch đọc file lớn: {os.path.basename(file_path)}...", 
)
    root.update()
    
    try:
        # Đọc thông tin file để tìm hiểu số dòng
        excel_reader = pd.ExcelFile(file_path, engine='openpyxl')
        sheet_name = excel_reader.sheet_names[0]  # Lấy sheet đầu tiên
        
        # Chỉ đọc header để phân tích
        if header_row is None:
            # Đọc 20 dòng đầu tiên để phân tích header
            header_sample = pd.read_excel(excel_reader, sheet_name=sheet_name, nrows=20)
            
            # Tìm hàng chứa header
            header_row = 0
            for i in range(len(header_sample)):
                row_values = header_sample.iloc[i].astype(str)
                if any(name.lower() in ' '.join(row_values.str.lower()) 
                      for name in ['họ và tên', 'tên học sinh', 'học sinh']):
                    header_row = i
                    break
        
        # Cập nhật trạng thái
        status_label.config(text=f"Đang đọc dữ liệu theo chunk từ dòng {header_row+1}...", 
)
        root.update()
        
        # Đọc dữ liệu theo từng chunk
        chunks = []
        reader = pd.read_excel(
            excel_reader, 
            sheet_name=sheet_name,
            header=header_row,
            chunksize=chunk_size
        )
        
        total_rows = 0
        
        for i, chunk in enumerate(reader):
            chunks.append(chunk)
            
            total_rows += len(chunk)
            status_label.config(text=f"Đang đọc dữ liệu: {total_rows} dòng...", 
)
            root.update()
            
            # Nếu đã đọc quá nhiều, hiển thị cảnh báo
            if total_rows > 10000:
                status_label.config(text=f"File lớn: đã đọc {total_rows} dòng...", 
)            # Ghép các chunk lại
        
        if chunks:
            status_label.config(text=f"Đang ghép {len(chunks)} chunk dữ liệu...", 
)
            root.update()
        
            result = pd.concat(chunks, ignore_index=True)
        
            # Đảm bảo kiểu dữ liệu phù hợp cho các cột quan trọng
            result = ensure_proper_dtypes(result)
        
            status_label.config(text=f"Đã đọc xong {total_rows} dòng dữ liệu", 
)
            return result
        else:
            return pd.DataFrame()
            
    except Exception as e:
        status_label.config(text=f"Lỗi khi đọc file: {str(e)}", 
)
        ToastNotification.show(
            f"❌ Không thể đọc file Excel\n"
            f"📄 Lỗi: {str(e)}\n"
            f"💡 Kiểm tra:\n"
            f"  • File có đúng định dạng .xlsx?\n"
            f"  • File có đang mở ở ứng dụng khác?\n"
            f"  • File có bị hỏng không?", 
            "error")
        print(f"Chi tiết lỗi: {traceback.format_exc()}")
        return pd.DataFrame()  # Trả về DataFrame rỗng thay vì None

def add_to_recent_files(filepath):
    """Thêm file vào danh sách recent files"""
    if 'recent_files' not in config:
        config['recent_files'] = []
    
    # Xóa file cũ nếu đã tồn tại trong danh sách
    if filepath in config['recent_files']:
        config['recent_files'].remove(filepath)
    
    # Thêm vào đầu danh sách
    config['recent_files'].insert(0, filepath)
    
    # Giới hạn 5 file gần nhất
    config['recent_files'] = config['recent_files'][:5]
    
    # Lưu config
    save_config()
    
    # Cập nhật menu
    update_recent_files_menu()

def open_recent_file(filepath):
    """Mở file từ danh sách recent files"""
    global df, file_path
    
    # Kiểm tra file có tồn tại không
    if not os.path.exists(filepath):
        ToastNotification.show(f"File không tồn tại: {os.path.basename(filepath)}", "error")
        # Xóa khỏi danh sách recent
        if filepath in config.get('recent_files', []):
            config['recent_files'].remove(filepath)
            save_config()
            update_recent_files_menu()
        return
    
    file_path = filepath
    status_label.config(text=f"Đang đọc file: {os.path.basename(file_path)}...")
    root.update()
    
    try:
        df = read_excel_file(file_path)
        
        if df is not None and not df.empty:
            total_students = len(df)
            status_label.config(text=f"Đã đọc xong: {os.path.basename(file_path)} ({total_students} học sinh)")
            
            df = ensure_required_columns(df)
            refresh_ui()
            
            # Thêm vào recent files
            add_to_recent_files(file_path)
            
            ToastNotification.show(f"Đã mở file {os.path.basename(file_path)}", "success")
        else:
            status_label.config(text="Không có dữ liệu để hiển thị")
    except Exception as e:
        error_message = str(e)
        status_label.config(text=f"Lỗi: {error_message[:50]}")
        ToastNotification.show(f"Lỗi khi đọc file: {error_message[:100]}", "error")
        traceback.print_exc()

def update_recent_files_menu():
    """Cập nhật menu Recent Files"""
    # Tìm menu File hoặc tạo mới
    try:
        # Xóa menu cũ nếu tồn tại
        menubar = root.nametowidget(root.cget('menu'))
        
        # Tìm hoặc tạo menu File
        file_menu = None
        for i in range(menubar.index('end') + 1 if menubar.index('end') is not None else 0):
            try:
                label = menubar.entrycget(i, 'label')
                if 'File' in label or 'Tệp' in label:
                    file_menu = menubar.nametowidget(menubar.entrycget(i, 'menu'))
                    break
            except:
                pass
        
        # Nếu chưa có, tạo menu File mới
        if file_menu is None:
            file_menu = tk.Menu(menubar, tearoff=0,
                              background=config['ui']['theme']['card'],
                              foreground=config['ui']['theme']['text'],
                              activebackground=config['ui']['theme']['primary'],
                              activeforeground='white')
            menubar.insert_cascade(0, label="📁 File", menu=file_menu)
        
        # Xóa các mục recent files cũ
        try:
            last_index = file_menu.index('end')
            if last_index is not None:
                for i in range(last_index, -1, -1):
                    try:
                        label = file_menu.entrycget(i, 'label')
                        if 'Recent' in label or any(ext in label for ext in ['.xlsx', '.xls']):
                            file_menu.delete(i)
                    except:
                        pass
        except:
            pass
        
        # Thêm separator và recent files mới
        if config.get('recent_files'):
            file_menu.add_separator()
            file_menu.add_command(label="📌 File Gần Đây", state='disabled')
            
            for i, filepath in enumerate(config['recent_files'][:5]):
                if os.path.exists(filepath):
                    filename = os.path.basename(filepath)
                    file_menu.add_command(
                        label=f"  {i+1}. {filename}",
                        command=lambda fp=filepath: open_recent_file(fp)
                    )
    except Exception as e:
        print(f"Lỗi khi cập nhật recent files menu: {str(e)}")

def select_file():
    global df, file_path
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        # Hiển thị thông báo trạng thái khi bắt đầu đọc file
        status_label.config(text=f"Đang đọc file: {os.path.basename(file_path)}...", 
)
        root.update()  # Cập nhật giao diện ngay lập tức để hiển thị trạng thái
        
        try:
            # Đọc file Excel bình thường
            df = read_excel_file(file_path)
            
            if df is not None and not df.empty:
                # Hiển thị số lượng học sinh
                total_students = len(df)
                status_label.config(
                    text=f"Đã đọc xong: {os.path.basename(file_path)} ({total_students} học sinh)",

                )
                
                # Đảm bảo các cột cần thiết tồn tại
                df = ensure_required_columns(df)
                
                # Thêm vào recent files
                add_to_recent_files(file_path)
                
                # Cập nhật giao diện
                refresh_ui()
                
                # Tự động làm mới dữ liệu sau khi mở file
                update_stats()
                update_score_extremes()
                
                ToastNotification.show(f"Đã mở file {os.path.basename(file_path)}", "success")
            else:
                status_label.config(
                    text="Không có dữ liệu để hiển thị, vui lòng tải file Excel có dữ liệu",

                )
                
        except Exception as e:
            error_message = str(e)
            status_label.config(
                text=f"Lỗi: {error_message[:50] + '...' if len(error_message) > 50 else error_message}",

            )
            ToastNotification.show(f"Lỗi: {error_message[:100]}", "error")
            traceback.print_exc()  # In chi tiết lỗi ra console để debug

def read_excel_normally(file_path):
    """Đọc file Excel theo cách thông thường"""
    status_label.config(text=f"Đang đọc file Excel...")
    root.update()
    
    try:
        # Thử đọc với engine openpyxl trước
        headers_df = pd.read_excel(file_path, nrows=10, engine='openpyxl')
        
        # Tìm hàng chứa header thực sự
        header_row = find_header_row(headers_df)
        
        # Đọc lại với header đúng
        df_result = pd.read_excel(file_path, header=header_row, engine='openpyxl')
        return df_result
    except ImportError:
        # Nếu không có openpyxl, thử dùng engine mặc định
        print("Không tìm thấy openpyxl, sử dụng engine mặc định...")
        headers_df = pd.read_excel(file_path, nrows=10)
        
        # Tìm hàng chứa header thực sự
        header_row = find_header_row(headers_df)
        
        # Đọc lại với header đúng
        df_result = pd.read_excel(file_path, header=header_row)
        return df_result
    except Exception as e:
        print(f"Chi tiết lỗi đọc file Excel: {traceback.format_exc()}")
        raise Exception(f"Lỗi khi đọc file Excel: {str(e)}")

def save_state():
    """Lưu trạng thái hiện tại để hoàn tác"""
    if df is not None:
        undo_stack.append(df.copy())
        if len(undo_stack) > 10:  # Giới hạn 10 bước hoàn tác
            undo_stack.pop(0)

def undo(event=None):
    """Hoàn tác thay đổi gần nhất"""
    global df
    if undo_stack:
        df = undo_stack.pop()
        save_excel()
        search_student()

def save_excel():
    """Lưu file Excel"""
    if df is not None and file_path:
        # Đảm bảo kiểu dữ liệu phù hợp trước khi lưu
        df_to_save = ensure_proper_dtypes(df)
        df_to_save.to_excel(file_path, index=False)
        # Cập nhật thống kê sau khi lưu
        update_stats()
        update_score_extremes()

def find_matching_column(df, target_name):
    """Find column that best matches the target name"""
    target_name = target_name.lower()
    
    # Direct match first (case insensitive)
    for col in df.columns:
        if col.lower() == target_name:
            return col
            
    # Common variations
    name_variations = {
        'tên học sinh': ['họ và tên', 'họ tên', 'tên', 'học sinh'],
        'mã đề': ['mã', 'đề', 'số đề', 'mã số đề'],
        'điểm': ['điểm số', 'số điểm', 'point', 'score']
    }
    
    # Check variations
    for col in df.columns:
        col_lower = col.lower()
        for key, variations in name_variations.items():
            if target_name == key:
                if any(var in col_lower for var in variations):
                    return col
                    
    return None

def search_student(event=None):
    """Tìm kiếm học sinh"""
    global df, search_timer_id, search_index
    
    # Reset timer
    search_timer_id = None
    
    # Hiển thị thông báo xử lý nếu dữ liệu lớn
    if df is not None and len(df) > 1000:
        status_label.config(text="Đang tìm kiếm trong dữ liệu lớn...")
        root.update()  # Cập nhật giao diện trước khi thực hiện tìm kiếm
    
    # Clear existing items
    for item in tree.get_children():
        tree.delete(item)
        
    if df is None or df.empty:
        ToastNotification.show("ℹ️ Chưa có dữ liệu học sinh", "info")
        update_stats()
        return

    # Find matching columns
    column_mapping = {}
    for key, configured_name in config['columns'].items():
        if configured_name.strip():
            matched_col = find_matching_column(df, configured_name)
            if matched_col:
                column_mapping[key] = matched_col

    if not column_mapping.get('name'):
        ToastNotification.show(
            "❌ Không tìm thấy cột tên học sinh\n"
            f"💡 Kiểm tra file Excel có cột '{config['columns']['name']}'\n"
            "   hoặc vào File > Cài đặt để cấu hình tên cột", 
            "error")
        return

    ten_hoc_sinh = entry_student_name.get().strip().lower()  # Chuyển về chữ thường để tìm kiếm không phân biệt hoa thường
    name_col = column_mapping['name']
    
    # Tạo search index nếu là lần đầu tìm kiếm hoặc nếu df mới được load
    if 'search_index' not in globals() or search_index is None or len(search_index) != len(df):
        search_index = df[name_col].str.lower()

    # Get display values helper (giữ nguyên)
    def get_display_values(row):
        values = []
        for col_key in ['name', 'exam_code', 'score']:
            if col_key in column_mapping:
                value = row[column_mapping[col_key]]
                if col_key == 'score':
                    values.append(f"{value:.2f}" if pd.notna(value) else 'Chưa có điểm')
                else:
                    values.append(str(value) if pd.notna(value) else '')
        return values

    # Hàm áp dụng màu theo điểm
    def apply_color_by_score(item_id, score_value):
        """Tô màu dòng theo điểm số"""
        try:
            if pd.notna(score_value):
                score = float(score_value)
                if score < 5:
                    tree.item(item_id, tags=('low_score',))
                elif score < 7:
                    tree.item(item_id, tags=('medium_score',))
                else:
                    tree.item(item_id, tags=('high_score',))
            else:
                tree.item(item_id, tags=('no_score',))
        except:
            tree.item(item_id, tags=('no_score',))
    
    # Display results with improved performance
    if not ten_hoc_sinh:
        # Nếu số lượng học sinh lớn, giới hạn hiển thị ban đầu
        display_limit = 100 if len(df) > 100 else len(df)
        for _, row in df.head(display_limit).iterrows():
            values = get_display_values(row)
            if values:
                item_id = tree.insert('', 'end', values=values)
                # Áp dụng màu theo điểm
                if 'score' in column_mapping:
                    apply_color_by_score(item_id, row[column_mapping['score']])
        
        # Thông báo nếu chỉ hiển thị một phần
        if len(df) > display_limit:
            tree.insert('', 'end', values=(f"--- Hiển thị {display_limit}/{len(df)} học sinh. Nhập từ khóa để tìm kiếm ---", "", ""))
    else:
        # Tạo mask cho việc tìm kiếm
        mask = search_index.str.contains(ten_hoc_sinh, na=False)
        result = df[mask]
        
        if result.empty:
            tree.insert('', 'end', values=('Không tìm thấy học sinh',) * len(column_mapping))
        else:
            # Giới hạn kết quả nếu quá nhiều
            result_limit = 200 if len(result) > 200 else len(result)
            
            # Đưa kết quả vào treeview
            for _, row in result.head(result_limit).iterrows():
                values = get_display_values(row)
                if values:
                    item_id = tree.insert('', 'end', values=values)
                    # Áp dụng màu theo điểm
                    if 'score' in column_mapping:
                        apply_color_by_score(item_id, row[column_mapping['score']])
            
            # Thông báo nếu có quá nhiều kết quả
            if len(result) > result_limit:
                tree.insert('', 'end', values=(f"--- Hiển thị {result_limit}/{len(result)} kết quả. Thêm ký tự để lọc chi tiết hơn ---", "", ""))
                    
            # Nếu chỉ tìm thấy một học sinh, tự động chọn học sinh đó
            if len(result) == 1:
                first_item = tree.get_children()[0]
                tree.selection_set(first_item)
                tree.focus(first_item)
                tree.see(first_item)
    
    # Cập nhật trạng thái
    if df is not None and len(df) > 1000:
        status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}")
                
    update_stats()

def add_student():
    """Thêm học sinh mới"""
    global df, undo_manager
    if df is None:
        ToastNotification.show("❌ Chưa chọn file Excel\n💡 Vui lòng chọn File > Mở để chọn file dữ liệu", "error")
        return

    ten_hoc_sinh = entry_student_name.get().strip()
    if not ten_hoc_sinh:
        ToastNotification.show("ℹ️ Vui lòng nhập tên học sinh để thêm", "info")
        return
    
    # Lưu state trước khi thêm
    undo_manager.push_state(df, f"Thêm học sinh '{ten_hoc_sinh}'")

    if ten_hoc_sinh in df[config['columns']['name']].values:
        ToastNotification.show("⚠️ Học sinh đã tồn tại trong danh sách", "warning")
    else:
        save_state()
        new_row = pd.DataFrame({config['columns']['name']: [ten_hoc_sinh], 'Điểm': [None]})
        df = pd.concat([df, new_row], ignore_index=True)
        save_excel()
        search_student()
        
        # Cập nhật thống kê
        update_stats()
        update_score_extremes()
        update_undo_redo_buttons()  # Cập nhật trạng thái undo/redo

def calculate_score(event=None):
    """Tính điểm từ số câu đúng"""
    global df, undo_manager
    if df is None:
        ToastNotification.show("Đề xuất tải file Excel trước khi nhập điểm", "error")
        return
 
    selected_item = tree.selection()
    if not selected_item:
        ToastNotification.show("Vui lòng chọn học sinh để nhập điểm", "warning")
        return

    selected = tree.item(selected_item[0])['values'][0]
    
    # Lưu state trước khi thay đổi
    undo_manager.push_state(df, f"Tính điểm cho '{selected}'")
    
    try:
        so_cau_dung = int(entry_correct_count.get())
        ma_de = entry_exam_code.get().strip()
        
        # Xử lý mã đề
        if ma_de.lower() == 'x':  # Xóa mã đề
            ma_de = None
        elif ma_de:  # Nếu có nhập mã đề
            try:
                ma_de = int(ma_de)
            except ValueError:
                ToastNotification.show("Mã đề phải là số", "error")
                return
        # Nếu để trống thì giữ nguyên mã đề cũ
        else:
            ma_de = df.loc[df[config['columns']['name']] == selected, 'Mã đề'].iloc[0]
            
        if not (0 <= so_cau_dung <= config['max_questions']):
            ToastNotification.show(f"Số câu đúng phải từ 0 đến {config['max_questions']}", "error")
            return

        diem = round(so_cau_dung * config['score_per_question'], 2)
        
        if diem > 10:
            ToastNotification.show("Điểm tính được vượt quá 10", "error")
            return
            
        save_state()
        
        # Đảm bảo kiểu dữ liệu phù hợp cho DataFrame trước khi gán giá trị
        df_proper = ensure_proper_dtypes(df)
        
        # Gán điểm và mã đề
        df_proper.loc[df_proper[config['columns']['name']] == selected, 'Điểm'] = diem
        if ma_de != df_proper.loc[df_proper[config['columns']['name']] == selected, 'Mã đề'].iloc[0]:
            df_proper.loc[df_proper[config['columns']['name']] == selected, 'Mã đề'] = ma_de
        
        # Cập nhật df chính
        df = df_proper
        
        save_excel()
        
        # Lưu index của học sinh hiện tại
        current_index = tree.index(selected_item[0])
        
        search_student()
        
        # Cập nhật thống kê
        update_stats()
        update_score_extremes()
        update_undo_redo_buttons()  # Cập nhật trạng thái undo/redo
        
        # Hiển thị toast thành công
        ToastNotification.show(f"Đã lưu điểm {diem} cho {selected}", "success")
        
        # Xóa nội dung ô số câu đúng
        entry_correct_count.delete(0, tk.END)
        
        # Tự động chọn học sinh tiếp theo
        children = tree.get_children()
        if children and current_index + 1 < len(children):
            next_item = children[current_index + 1]
            tree.selection_set(next_item)
            tree.focus(next_item)
            tree.see(next_item)
            # Focus vào ô nhập số câu đúng để tiếp tục nhập
            entry_correct_count.focus_set()
    except ValueError:
        ToastNotification.show("Vui lòng nhập số câu đúng hợp lệ", "error")

def update_config(event=None):
    """Cập nhật cấu hình tính điểm"""
    global score_per_q_label
    try:
        max_q = int(entry_max_questions.get())
        
        if max_q <= 0:
            ToastNotification.show("❌ Số câu hỏi không hợp lệ\n💡 Vui lòng nhập số nguyên dương (> 0)", "error")
            return
            
        # Tự động tính điểm mỗi câu bằng cách chia 10 cho số câu hỏi
        score_per_q = round(10 / max_q, 2)
            
        config['max_questions'] = max_q
        config['score_per_question'] = score_per_q
        
        save_config()  # Lưu vào file
        
        # Cập nhật label thông tin ở header
        score_per_q_label.config(text=f"({score_per_q}đ/câu)")
        
        ToastNotification.show(f"✅ Đã cập nhật: {max_q} câu, mỗi câu {score_per_q} điểm", "success")
        
    except ValueError:
        ToastNotification.show("❌ Giá trị không hợp lệ\n💡 Vui lòng nhập số nguyên (VD: 20, 30, 40)", "error")

def focus_student_search(event=None):
    """Di chuyển con trỏ đến ô tìm kiếm học sinh"""
    entry_student_name.focus_set()
    entry_student_name.select_range(0, tk.END)

def perform_undo():
    """Thực hiện hoàn tác thao tác cuối"""
    global df, undo_manager
    
    if not undo_manager.can_undo():
        ToastNotification.show("ℹ️ Không có thao tác nào để hoàn tác", "info")
        return
    
    previous_state, action_name = undo_manager.undo()
    if previous_state is not None:
        df = previous_state.copy()
        refresh_ui()
        save_excel()
        ToastNotification.show(f"⏪ Đã hoàn tác: {action_name}", "success")
        update_undo_redo_buttons()

def perform_redo():
    """Thực hiện làm lại thao tác đã hoàn tác"""
    global df, undo_manager
    
    if not undo_manager.can_redo():
        ToastNotification.show("ℹ️ Không có thao tác nào để làm lại", "info")
        return
    
    next_state, action_name = undo_manager.redo()
    if next_state is not None:
        df = next_state.copy()
        refresh_ui()
        save_excel()
        ToastNotification.show(f"⏩ Đã làm lại: {action_name}", "success")
        update_undo_redo_buttons()

def update_undo_redo_buttons():
    """Cập nhật trạng thái enabled/disabled của nút undo/redo"""
    global undo_manager
    try:
        if undo_manager.can_undo():
            undo_button.config(state='normal')
        else:
            undo_button.config(state='disabled')
        
        if undo_manager.can_redo():
            redo_button.config(state='normal')
        else:
            redo_button.config(state='disabled')
    except:
        pass  # Buttons chưa được tạo

def calculate_score_direct(event=None):
    """Tính điểm trực tiếp từ điểm số nhập vào"""
    global df, entry_direct_score, undo_manager  # Thêm entry_direct_score vào global
    if df is None:
        ToastNotification.show("Đề xuất tải file Excel trước khi nhập điểm", "error")
        return

    selected_item = tree.selection()
    if not selected_item:
        ToastNotification.show("Vui lòng chọn học sinh để nhập điểm", "warning")
        return

    selected = tree.item(selected_item[0])['values'][0]
    
    # Lưu state trước khi thay đổi
    undo_manager.push_state(df, f"Nhập điểm trực tiếp cho '{selected}'")
    
    try:
        diem = float(entry_direct_score.get())
        ma_de = entry_exam_code.get().strip()
        
        # Xử lý mã đề
        if ma_de.lower() == 'x':  # Xóa mã đề
            ma_de = None
        elif ma_de:  # Nếu có nhập mã đề
            try:
                ma_de = int(ma_de)
            except ValueError:
                ToastNotification.show("Mã đề phải là số", "error")
                return
        # Nếu để trống thì giữ nguyên mã đề cũ
        else:
            ma_de = df.loc[df[config['columns']['name']] == selected, 'Mã đề'].iloc[0]
            
        if not (0 <= diem <= 10):
            ToastNotification.show("Điểm phải từ 0 đến 10", "error")
            return
            
        save_state()
        
        # Đảm bảo kiểu dữ liệu phù hợp cho DataFrame trước khi gán giá trị
        df_proper = ensure_proper_dtypes(df)
        
        # Gán điểm và mã đề
        df_proper.loc[df_proper[config['columns']['name']] == selected, 'Điểm'] = diem
        if ma_de != df_proper.loc[df_proper[config['columns']['name']] == selected, 'Mã đề'].iloc[0]:
            df_proper.loc[df_proper[config['columns']['name']] == selected, 'Mã đề'] = ma_de
        
        # Cập nhật df chính
        df = df_proper
        
        save_excel()
        
        # Lưu index của học sinh hiện tại
        current_index = tree.index(selected_item[0])
        
        search_student()
        
        # Cập nhật thống kê
        update_stats()
        update_score_extremes()
        update_undo_redo_buttons()  # Cập nhật trạng thái undo/redo
        
        # Hiển thị toast thành công
        ToastNotification.show(f"Đã lưu điểm {diem} cho {selected}", "success")
        
        # Xóa nội dung ô nhập điểm
        entry_direct_score.delete(0, tk.END)
        
        # Tự động chọn học sinh tiếp theo
        children = tree.get_children()
        if children and current_index + 1 < len(children):
            next_item = children[current_index + 1]
            tree.selection_set(next_item)
            tree.focus(next_item)
            tree.see(next_item)
            # Focus vào ô nhập điểm trực tiếp để tiếp tục nhập
            entry_direct_score.focus_set()
    except ValueError:
        ToastNotification.show("Vui lòng nhập điểm hợp lệ", "error")

def focus_direct_score(event=None):
    """Di chuyển con trỏ đến ô nhập điểm trực tiếp"""
    entry_direct_score.focus_set()
    entry_direct_score.select_range(0, tk.END)

def focus_correct_count(event=None):
    """Di chuyển con trỏ đến ô nhập số câu đúng"""
    entry_correct_count.focus_set()
    entry_correct_count.select_range(0, tk.END)

def customize_shortcuts():
    """Mở cửa sổ tùy chỉnh phím tắt"""
    shortcut_window = tk.Toplevel(root)
    shortcut_window.title("Tùy chỉnh phím tắt")
    shortcut_window.geometry("400x300")
    
    def save_shortcuts():
        config['shortcuts']['search'] = search_entry.get()
        config['shortcuts']['direct_score'] = score_entry.get()
        config['shortcuts']['undo'] = undo_entry.get()
        
        # Cập nhật bindings
        root.bind(config['shortcuts']['search'], focus_student_search)
        root.bind(config['shortcuts']['direct_score'], focus_direct_score)
        root.bind(config['shortcuts']['undo'], undo)
        
        save_config()  # Lưu vào file
        shortcut_window.destroy()
        ToastNotification.show("✅ Đã lưu cấu hình phím tắt", "success")
    
    ttk.Label(shortcut_window, text="Tìm kiếm:").pack(pady=5)
    search_entry = ttk.Entry(shortcut_window)
    search_entry.insert(0, config['shortcuts']['search'])
    search_entry.pack(pady=5)
    
    ttk.Label(shortcut_window, text="Nhập điểm:").pack(pady=5)
    score_entry = ttk.Entry(shortcut_window)
    score_entry.insert(0, config['shortcuts']['direct_score'])
    score_entry.pack(pady=5)
    
    ttk.Label(shortcut_window, text="Hoàn tác:").pack(pady=5)
    undo_entry = ttk.Entry(shortcut_window)
    undo_entry.insert(0, config['shortcuts']['undo'])
    undo_entry.pack(pady=5)
    
    ttk.Button(shortcut_window, text="Lưu", command=save_shortcuts).pack(pady=10)

def customize_exam_codes():
    """Mở cửa sổ tùy chỉnh mã đề"""
    code_window = tk.Toplevel(root)
    code_window.title("Tùy chỉnh mã đề")
    code_window.geometry("400x300")
    
    def save_codes():
        codes = code_text.get("1.0", tk.END).strip().split('\n')
        codes = [code.strip() for code in codes if code.strip()]
        config['exam_codes'] = codes
        entry_exam_code['values'] = codes
        save_config()  # Lưu vào file
        code_window.destroy()
        ToastNotification.show("✅ Đã lưu danh sách mã đề", "success")
    
    ttk.Label(code_window, text="Nhập mỗi mã đề trên một dòng:").pack(pady=5)
    code_text = tk.Text(code_window, height=10)
    code_text.pack(pady=5, padx=5, fill=tk.BOTH, expand=True)
    code_text.insert("1.0", '\n'.join(config['exam_codes']))
    
    ttk.Button(code_window, text="Lưu", command=save_codes).pack(pady=10)

def customize_columns():
    """Open window to customize column names"""
    column_window = tk.Toplevel(root)
    column_window.title("Tùy chỉnh tên cột")
    column_window.geometry("400x300")
    
    # Create frame for entries
    entry_frame = ttk.Frame(column_window)
    entry_frame.pack(pady=10, padx=10)
    
    entries = {}
    
    # Create entries for each column
    for i, (key, value) in enumerate(config['columns'].items()):
        ttk.Label(entry_frame, text=f"Tên cột {value}:").grid(row=i, column=0, pady=5, sticky='e')
        entry = ttk.Entry(entry_frame, width=30)
        entry.insert(0, value)
        entry.grid(row=i, column=1, pady=5, padx=5)
        entries[key] = entry

    def save_columns():
        # Save new column names
        for key, entry in entries.items():
            new_name = entry.get().strip()
            if new_name:  # Only update if not empty
                config['columns'][key] = new_name
        
        # Update column headers in treeview
        tree.heading('name', text=config['columns']['name'])
        tree.heading('exam_code', text=config['columns']['exam_code'])
        tree.heading('score', text=config['columns']['score'])
        
        save_config()  # Save to config file
        column_window.destroy()
        ToastNotification.show("✅ Đã lưu tên cột mới", "success")
        if 'df' in globals() and df is not None and not df.empty:
            search_student()  # Refresh display

    ttk.Button(column_window, text="Lưu", command=save_columns).pack(pady=10)

def customize_security():
    """Mở cửa sổ tùy chỉnh cài đặt bảo mật"""
    security_window = tk.Toplevel(root)
    security_window.title("Tùy chỉnh bảo mật")
    security_window.geometry("500x450")
    security_window.transient(root)  # Đặt là cửa sổ con của root
    security_window.grab_set()  # Đặt là modal
    
    # Frame chứa các tùy chọn bảo mật
    options_frame = ttk.LabelFrame(security_window, text="Tùy chọn bảo mật", padding=10)
    options_frame.pack(fill="both", expand=True, padx=15, pady=10)
    
    # Khởi tạo các biến để lưu trạng thái
    encrypt_backups_var = tk.BooleanVar(value=config.get('security', {}).get('encrypt_backups', False))
    encrypt_sensitive_var = tk.BooleanVar(value=config.get('security', {}).get('encrypt_sensitive_data', False))
    password_protect_var = tk.BooleanVar(value=config.get('security', {}).get('password_protect_app', False))
    
    # Mức độ bảo mật backup
    encryption_level_var = tk.StringVar(value=config.get('security', {}).get('backup_encryption_level', 'medium'))
    
    # Auto lock timeout (minutes)
    auto_lock_var = tk.IntVar(value=config.get('security', {}).get('auto_lock_timeout_minutes', 30))
    
    # Tùy chọn tự động sao lưu
    auto_backup_var = tk.BooleanVar(value=config.get('auto_backup', False))
    ttk.Checkbutton(options_frame, text="Tự động sao lưu khi đóng chương trình", 
                  variable=auto_backup_var).pack(anchor="w", pady=5)
    
    # Tùy chọn mã hóa backup
    ttk.Checkbutton(options_frame, text="Mã hóa file sao lưu tự động", 
                  variable=encrypt_backups_var).pack(anchor="w", pady=5)
    
    # Tùy chọn mã hóa dữ liệu nhạy cảm
    ttk.Checkbutton(options_frame, text="Mã hóa dữ liệu nhạy cảm (điểm số)", 
                  variable=encrypt_sensitive_var).pack(anchor="w", pady=5)
    
    # Tùy chọn bảo vệ ứng dụng bằng mật khẩu
    ttk.Checkbutton(options_frame, text="Bảo vệ ứng dụng bằng mật khẩu", 
                  variable=password_protect_var).pack(anchor="w", pady=5)
    
    # Frame cho mức độ bảo mật
    security_level_frame = ttk.Frame(options_frame)
    security_level_frame.pack(fill="x", pady=10)
    
    ttk.Label(security_level_frame, text="Mức độ mã hóa backup:").pack(side="left", padx=5)
    
    # Combobox cho mức độ mã hóa
    level_combo = ttk.Combobox(security_level_frame, 
                             values=["low", "medium", "high"], 
                             textvariable=encryption_level_var,
                             state="readonly",
                             width=10)
    level_combo.pack(side="left", padx=5)
    
    # Frame cho thời gian tự động khóa
    auto_lock_frame = ttk.Frame(options_frame)
    auto_lock_frame.pack(fill="x", pady=10)
    
    ttk.Label(auto_lock_frame, text="Tự động khóa sau (phút):").pack(side="left", padx=5)
    
    # Spinbox cho thời gian tự động khóa
    auto_lock_spinbox = ttk.Spinbox(auto_lock_frame, 
                                  from_=1, 
                                  to=240, 
                                  textvariable=auto_lock_var,
                                  width=5)
    auto_lock_spinbox.pack(side="left", padx=5)
    
    # Frame cho mật khẩu
    password_frame = ttk.LabelFrame(options_frame, text="Cài đặt mật khẩu", padding=10)
    password_frame.pack(fill="x", pady=10, padx=5)
    
    # Các entry cho mật khẩu
    password_var = tk.StringVar()
    confirm_password_var = tk.StringVar()
    
    # Label và Entry cho mật khẩu
    ttk.Label(password_frame, text="Mật khẩu mới:").grid(row=0, column=0, sticky="w", pady=5)
    password_entry = ttk.Entry(password_frame, show="*", textvariable=password_var, width=30)
    password_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
    
    # Label và Entry cho xác nhận mật khẩu
    ttk.Label(password_frame, text="Xác nhận mật khẩu:").grid(row=1, column=0, sticky="w", pady=5)
    confirm_password_entry = ttk.Entry(password_frame, show="*", textvariable=confirm_password_var, width=30)
    confirm_password_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
    
    # Thêm checkbox hiện/ẩn mật khẩu
    show_password_var = tk.BooleanVar(value=False)
    
    def toggle_password_visibility():
        """Hiện/ẩn mật khẩu"""
        if show_password_var.get():
            password_entry.config(show="")
            confirm_password_entry.config(show="")
        else:
            password_entry.config(show="*")
            confirm_password_entry.config(show="*")
    
    ttk.Checkbutton(password_frame, text="Hiện mật khẩu", 
                  variable=show_password_var,
                  command=toggle_password_visibility).grid(row=2, column=0, columnspan=2, sticky="w", pady=5)
    
    # Frame cho nút bấm
    button_frame = ttk.Frame(security_window)
    button_frame.pack(fill="x", padx=15, pady=15)
    
    def save_security_settings():
        """Lưu cài đặt bảo mật"""
        # Kiểm tra mật khẩu nếu bảo vệ ứng dụng được bật
        if password_protect_var.get():
            if not password_var.get():
                messagebox.showwarning("Cảnh báo", "Vui lòng nhập mật khẩu khi bật tính năng bảo vệ ứng dụng.")
                return
            
            if password_var.get() != confirm_password_var.get():
                messagebox.showwarning("Cảnh báo", "Mật khẩu xác nhận không khớp. Vui lòng nhập lại.")
                return
        
        # Lưu các cài đặt vào config
        if 'security' not in config:
            config['security'] = {}
        
        # Lưu cài đặt auto backup
        config['auto_backup'] = auto_backup_var.get()
            
        config['security']['encrypt_backups'] = encrypt_backups_var.get()
        config['security']['encrypt_sensitive_data'] = encrypt_sensitive_var.get()
        config['security']['password_protect_app'] = password_protect_var.get()
        config['security']['backup_encryption_level'] = encryption_level_var.get()
        config['security']['auto_lock_timeout_minutes'] = auto_lock_var.get()
        
        # Lưu mật khẩu nếu đã nhập
        if password_var.get():
            config['security']['password'] = password_var.get()
        
        # Lưu cấu hình
        save_config()
        
        # Khởi động tính năng tự động khóa nếu được bật
        if password_protect_var.get() and not hasattr(root, 'auto_lock_job'):
            check_auto_lock()
        
        # Đóng cửa sổ cài đặt
        security_window.destroy()
        
        ToastNotification.show("✅ Đã lưu cấu hình bảo mật", "success")
    
    # Nút lưu cài đặt
    ttk.Button(button_frame, text="Lưu cài đặt", 
             command=save_security_settings, 
             width=15).pack(side="left", padx=5)
    
    # Nút hủy
    ttk.Button(button_frame, text="Hủy", 
             command=security_window.destroy, 
             width=15).pack(side="right", padx=5)

def backup_data():
    """Sao lưu dữ liệu ra file"""
    global df, file_path
    if df is None or file_path is None:
        ToastNotification.show("❌ Chưa có dữ liệu để sao lưu\n💡 Vui lòng mở file Excel trước khi sao lưu", "error")
        return
    
    try:
        # Tạo tên file sao lưu
        backup_dir = os.path.join(os.path.dirname(file_path), "Backup")
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)
            
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.basename(file_path)
        name, ext = os.path.splitext(filename)
        
        # Thông tin về số lượng học sinh
        total_students = len(df)
        scored_students = df.notna()[config['columns']['score']].sum() if config['columns']['score'] in df.columns else 0
        
        # Thêm thông tin vào tên backup
        backup_filename = f"{name}_backup_{timestamp}_{total_students}hs{ext}"
        backup_path = os.path.join(backup_dir, backup_filename)
        
        # Tạo cửa sổ xác nhận với thông tin chi tiết
        confirm_window = tk.Toplevel(root)
        confirm_window.title("Xác nhận sao lưu")
        confirm_window.geometry("450x350")
        confirm_window.transient(root)  # Đặt là cửa sổ con của root
        confirm_window.grab_set()  # Đặt là modal
        
        # Thêm thông tin backup
        ttk.Label(confirm_window, text="Thông tin sao lưu", font=(config['ui']['font_family'], 12, 'bold')).pack(pady=10)
        
        info_frame = ttk.Frame(confirm_window)
        info_frame.pack(fill="both", padx=20, pady=5)
        
        info_labels = [
            f"File gốc: {filename}",
            f"Thời gian: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
            f"Tổng số học sinh: {total_students}",
            f"Học sinh có điểm: {scored_students}",
            f"Đường dẫn: {backup_path}"
        ]
        
        for idx, info in enumerate(info_labels):
            ttk.Label(info_frame, text=info, wraplength=400, justify="left").pack(anchor="w", pady=3)
        
        # Tùy chọn bảo mật
        security_frame = ttk.Frame(confirm_window)
        security_frame.pack(fill="x", padx=20, pady=5)
        
        encrypt_var = tk.BooleanVar(value=config.get('security', {}).get('encrypt_backups', False))
        ttk.Checkbutton(security_frame, text="Mã hóa file sao lưu", variable=encrypt_var).pack(anchor="w")
        
        # Frame cho mật khẩu
        password_frame = ttk.Frame(confirm_window)
        password_frame.pack(fill="x", padx=20, pady=5)
        
        ttk.Label(password_frame, text="Mật khẩu:").pack(side="left")
        password_entry = ttk.Entry(password_frame, show="*")
        password_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        # Dùng biến để theo dõi trạng thái
        confirm_window.encrypt_password = ""
        
        def toggle_password_entry(*args):
            # Kích hoạt/vô hiệu hóa ô mật khẩu dựa trên trạng thái của checkbox
            if encrypt_var.get():
                password_entry.config(state="normal")
            else:
                password_entry.delete(0, tk.END)
                password_entry.config(state="disabled")
        
        # Thiết lập callback khi thay đổi trạng thái encrypt
        encrypt_var.trace_add("write", toggle_password_entry)
        toggle_password_entry()  # Gọi lần đầu để thiết lập trạng thái ban đầu
        
        # Nút xác nhận
        btn_frame = ttk.Frame(confirm_window)
        btn_frame.pack(fill="x", padx=20, pady=15)
        
        def confirm_backup():
            # Kiểm tra mật khẩu khi mã hóa được chọn
            if encrypt_var.get():
                password = password_entry.get()
                if not password:
                    messagebox.showwarning("Cảnh báo", "Vui lòng nhập mật khẩu để mã hóa file sao lưu.")
                    return
                confirm_window.encrypt_password = password
            
            # Lưu cấu hình mã hóa
            # Đảm bảo config['security'] đã được khởi tạo
            if 'security' not in config:
                config['security'] = {}
            config['security']['encrypt_backups'] = encrypt_var.get()
            save_config()
            
            # Chuẩn bị thông tin metadata
            backup_info = {
                "original_file": filename,
                "backup_time": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                "total_students": int(total_students),
                "scored_students": int(scored_students),
                "app_version": config.get('version', 'unknown'),
                "encrypted": encrypt_var.get()
            }
            
            try:
                # Nếu chọn mã hóa
                if encrypt_var.get():
                    # Chuyển DataFrame thành chuỗi JSON
                    json_data = df.to_json(orient='records')
                    
                    # Mã hóa dữ liệu
                    encrypted_data, salt = encrypt_data(json_data, confirm_window.encrypt_password)
                    
                    # Lưu dữ liệu mã hóa
                    encrypted_path = os.path.splitext(backup_path)[0] + ".enc"
                    with open(encrypted_path, 'wb') as f:
                        f.write(encrypted_data)
                    
                    # Thêm thông tin về salt vào metadata
                    backup_info["salt"] = base64.b64encode(salt).decode('utf-8')
                    backup_info["encryption_method"] = "fernet_pbkdf2"
                    
                    # Cập nhật đường dẫn file để thông báo
                    backup_path_display = encrypted_path
                else:
                    # Lưu file thông thường nếu không mã hóa
                    df.to_excel(backup_path, index=False)
                    backup_path_display = backup_path
                
                # Lưu metadata vào file json riêng
                metadata_path = os.path.splitext(backup_path)[0] + "_info.json"
                with open(metadata_path, 'w', encoding='utf-8') as f:
                    json.dump(backup_info, f, ensure_ascii=False, indent=2)
                
                confirm_window.destroy()
                ToastNotification.show(f"✅ Đã sao lưu dữ liệu vào:\n{backup_path_display}", "success")
                
            except Exception as enc_error:
                ToastNotification.show(
                    f"❌ Không thể mã hóa dữ liệu\n"
                    f"📄 Lỗi: {str(enc_error)}\n"
                    f"💡 Thử lại hoặc sao lưu không mã hóa",
                    "error")
                traceback.print_exc()
        
        ttk.Button(btn_frame, text="Sao lưu", command=confirm_backup, width=15).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Hủy", command=confirm_window.destroy, width=15).pack(side="right", padx=10)
        
    except Exception as e:
        ToastNotification.show(
            f"❌ Không thể sao lưu dữ liệu\n"
            f"📄 Lỗi: {str(e)}\n"
            f"💡 Kiểm tra:\n"
            f"  • Có quyền ghi vào thư mục?\n"
            f"  • Đủ dung lượng ổ đĩa?",
            "error")
        traceback.print_exc()

def restore_backup():
    """Phục hồi dữ liệu từ bản sao lưu"""
    global df, file_path, search_index
    
    try:
        # Chọn file backup
        backup_path = filedialog.askopenfilename(
            title="Chọn file backup để phục hồi",
            filetypes=[
                ("Excel file", "*.xlsx *.xls"),
                ("Encrypted backup", "*.enc"),
                ("All files", "*.*")
            ]
        )
        
        if not backup_path:
            return
            
        # Kiểm tra xem có phải file mã hóa không
        is_encrypted = backup_path.lower().endswith('.enc')
        
        # Tìm file metadata nếu có
        metadata_path = os.path.splitext(backup_path)[0] + "_info.json"
        backup_info = {}
        
        if os.path.exists(metadata_path):
            try:
                with open(metadata_path, 'r', encoding='utf-8') as f:
                    backup_info = json.load(f)
            except:
                backup_info = {}  # Nếu không thể đọc metadata, bỏ qua
        
        # Tạo cửa sổ xác nhận
        confirm_window = tk.Toplevel(root)
        confirm_window.title("Xác nhận phục hồi dữ liệu")
        confirm_window.geometry("500x400")
        confirm_window.transient(root)
        confirm_window.grab_set()
        
        # Hiển thị thông tin backup
        ttk.Label(confirm_window, text="Thông tin bản sao lưu:", 
                 font=(config['ui']['font_family'], 12, 'bold')).pack(pady=10)
                 
        info_frame = ttk.Frame(confirm_window)
        info_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Hiển thị thông tin chi tiết từ metadata nếu có
        if backup_info:
            info_labels = [
                f"File gốc: {backup_info.get('original_file', os.path.basename(backup_path))}",
                f"Thời gian sao lưu: {backup_info.get('backup_time', 'Không xác định')}",
                f"Số học sinh: {backup_info.get('total_students', 'Không xác định')}",
                f"Học sinh có điểm: {backup_info.get('scored_students', 'Không xác định')}",
                f"Phiên bản ứng dụng: {backup_info.get('app_version', 'Không xác định')}",
                f"Loại backup: {'Tự động' if backup_info.get('auto_backup', False) else 'Thủ công'}",
                f"Mã hóa: {'Có' if backup_info.get('encrypted', False) else 'Không'}"
            ]
            
            if backup_info.get('encrypted', False):
                info_labels.append(f"Mức độ mã hóa: {backup_info.get('encryption_level', 'medium')}")
        else:
            # Nếu không có metadata, hiển thị thông tin cơ bản
            info_labels = [
                f"File: {os.path.basename(backup_path)}",
                f"Đường dẫn: {os.path.dirname(backup_path)}",
                f"Kích thước: {os.path.getsize(backup_path) / 1024:.1f} KB",
                f"Mã hóa: {'Có' if is_encrypted else 'Không'}"
            ]
        
        for info in info_labels:
            ttk.Label(info_frame, text=info, wraplength=460).pack(anchor="w", pady=2)
            
        # Cảnh báo
        warning_text = "CẢNH BÁO: Dữ liệu hiện tại sẽ bị ghi đè bởi dữ liệu từ bản sao lưu này!"
        ttk.Label(confirm_window, text=warning_text, foreground="red").pack(pady=10)
        
        # Ô nhập mật khẩu nếu file được mã hóa
        password_var = tk.StringVar()
        password_frame = ttk.Frame(confirm_window)
        
        if is_encrypted:
            password_frame.pack(pady=10, fill="x", padx=20)
            ttk.Label(password_frame, text="Mật khẩu giải mã:").pack(side="left", padx=5)
            password_entry = ttk.Entry(password_frame, show="*", textvariable=password_var)
            password_entry.pack(side="left", padx=5, expand=True, fill="x")
            password_entry.focus_set()
        
        # Frame chứa các nút
        btn_frame = ttk.Frame(confirm_window)
        btn_frame.pack(pady=20)
        
        def confirm_restore():
            """Thực hiện phục hồi dữ liệu"""
            global df, file_path, search_index
            
            try:
                if is_encrypted:
                    # Đọc dữ liệu mã hóa
                    with open(backup_path, 'rb') as f:
                        encrypted_data = f.read()
                    
                    # Lấy thông tin về salt từ metadata
                    salt = None
                    iterations = 100000  # Medium level by default
                    
                    if 'salt' in backup_info:
                        salt = base64.b64decode(backup_info['salt'])
                    
                    if 'iterations' in backup_info:
                        iterations = backup_info['iterations']
                    
                    # Nếu không có salt trong metadata, không thể giải mã
                    if salt is None:
                        raise ValueError("Không tìm thấy thông tin salt, không thể giải mã")
                    
                    # Giải mã dữ liệu
                    password = password_var.get()
                    if not password:
                        ToastNotification.show("❌ Chưa nhập mật khẩu\n💡 Vui lòng nhập mật khẩu để giải mã backup", "error")
                        return
                    
                    # Tạo khóa
                    kdf = PBKDF2HMAC(
                        algorithm=hashes.SHA256(),
                        length=32,
                        salt=salt,
                        iterations=iterations,
                    )
                    
                    key = base64.urlsafe_b64encode(kdf.derive(password.encode()))
                    fernet = Fernet(key)
                    
                    try:
                        # Giải mã
                        decrypted_data = fernet.decrypt(encrypted_data).decode('utf-8')
                        
                        # Parse dữ liệu JSON thành DataFrame
                        df = pd.read_json(decrypted_data, orient='records')
                    except Exception as decrypt_error:
                        ToastNotification.show(
                            f"❌ Không thể giải mã dữ liệu\n"
                            f"📄 Lỗi: {str(decrypt_error)}\n"
                            f"💡 Nguyên nhân có thể:\n"
                            f"  • Mật khẩu không đúng\n"
                            f"  • File backup bị hỏng\n"
                            f"  • File không phải backup mã hóa",
                            "error")
                        return
                else:
                    # Đọc file Excel bình thường
                    df = pd.read_excel(backup_path)
                
                # Cập nhật file_path nếu file backup có thông tin về file gốc
                if 'original_file' in backup_info and not file_path:
                    # Tìm đường dẫn gốc dựa trên tên file
                    original_file = backup_info['original_file']
                    original_dir = os.path.dirname(backup_path)  # Giả sử file gốc ở cùng thư mục với backup
                    file_path = os.path.join(original_dir, original_file)
                    
                    # Kiểm tra xem file gốc có tồn tại không
                    if not os.path.exists(file_path):
                        # Nếu không tồn tại, sử dụng đường dẫn backup (loại bỏ _backup phần)
                        base_name = os.path.basename(backup_path)
                        if '_backup_' in base_name:
                            original_name = base_name.split('_backup_')[0] + os.path.splitext(base_name)[1]
                            file_path = os.path.join(original_dir, original_name)
                
                # Đảm bảo kiểu dữ liệu phù hợp
                df = ensure_proper_dtypes(df)
                
                # Khởi tạo lại search index
                search_index = None
                
                # Cập nhật giao diện
                refresh_ui()
                
                confirm_window.destroy()
                ToastNotification.show("✅ Đã phục hồi dữ liệu từ bản sao lưu", "success")
            except Exception as e:
                ToastNotification.show(
                    f"❌ Không thể phục hồi dữ liệu\n"
                    f"📄 Lỗi: {str(e)}\n"
                    f"💡 Kiểm tra:\n"
                    f"  • File backup có còn tồn tại?\n"
                    f"  • File có bị hỏng không?\n"
                    f"  • Định dạng file có đúng?",
                    "error")
                traceback.print_exc()
                confirm_window.destroy()
        
        ttk.Button(btn_frame, text="Xác nhận", command=confirm_restore, width=15).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Hủy", command=confirm_window.destroy, width=15).pack(side="right", padx=5)
        
    except Exception as e:
        ToastNotification.show(
            f"❌ Không thể đọc file backup\n"
            f"📄 Lỗi: {str(e)}\n"
            f"💡 Vui lòng chọn file backup hợp lệ (.xlsx hoặc .bak)",
            "error")
        traceback.print_exc()

def auto_backup_on_exit():
    """Tự động sao lưu dữ liệu khi thoát ứng dụng nếu được bật"""
    global df, file_path

    # Kiểm tra cài đặt tự động sao lưu
    if df is not None and file_path is not None and config.get('auto_backup', False):
        try:
            # Tạo tên file sao lưu
            backup_dir = os.path.join(os.path.dirname(file_path), "Backup")
            if not os.path.exists(backup_dir):
                os.makedirs(backup_dir)
                
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = os.path.basename(file_path)
            name, ext = os.path.splitext(filename)
            
            # Thông tin về số lượng học sinh
            total_students = len(df)
            scored_students = df.notna()[config['columns']['score']].sum() if config['columns']['score'] in df.columns else 0
            
            # Kiểm tra nếu mã hóa được bật trong cấu hình
            encrypt_backups = config.get('security', {}).get('encrypt_backups', False)
            backup_path = ""
            backup_info = {
                "original_file": filename,
                "backup_time": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                "total_students": int(total_students),
                "scored_students": int(scored_students),
                "app_version": config.get('version', 'unknown'),
                "auto_backup": True,
                "encrypted": encrypt_backups,
                "encryption_level": config.get('security', {}).get('backup_encryption_level', 'medium')
            }
            
            # Nếu mã hóa được bật
            if encrypt_backups:
                # Lấy mật khẩu từ cấu hình
                backup_password = config.get('security', {}).get('password', '')
                
                if not backup_password:
                    # Không thể mã hóa mà không có mật khẩu, sao lưu không mã hóa thay thế
                    print("Cảnh báo: Không thể mã hóa do không có mật khẩu. Sao lưu không mã hóa.")
                    backup_info["encrypted"] = False
                    backup_filename = f"{name}_auto_backup_{timestamp}{ext}"
                    backup_path = os.path.join(backup_dir, backup_filename)
                    df.to_excel(backup_path, index=False)
                else:
                    # Tên file sao lưu với đuôi .enc
                    backup_filename = f"{name}_auto_backup_{timestamp}.enc"
                    backup_path = os.path.join(backup_dir, backup_filename)
                    
                    # Chuyển DataFrame thành chuỗi JSON
                    json_data = df.to_json(orient='records')
                    
                    # Mã hóa dữ liệu với độ bảo mật tương ứng
                    encryption_level = config.get('security', {}).get('backup_encryption_level', 'medium')
                    iterations = 100000  # Mặc định, medium level
                    
                    if encryption_level == 'low':
                        iterations = 50000
                    elif encryption_level == 'high':
                        iterations = 200000
                    
                    # Dùng iterations khác để encrypt
                    salt = os.urandom(16)
                    kdf = PBKDF2HMAC(
                        algorithm=hashes.SHA256(),
                        length=32,
                        salt=salt,
                        iterations=iterations,
                    )
                    
                    key = base64.urlsafe_b64encode(kdf.derive(backup_password.encode()))
                    fernet = Fernet(key)
                    encrypted_data = fernet.encrypt(json_data.encode())
                    
                    # Lưu dữ liệu mã hóa
                    with open(backup_path, 'wb') as f:
                        f.write(encrypted_data)
                    
                    # Thêm thông tin về salt vào metadata
                    backup_info["salt"] = base64.b64encode(salt).decode('utf-8')
                    backup_info["encryption_method"] = "fernet_pbkdf2"
                    backup_info["iterations"] = iterations
            else:
                # Sao lưu bình thường nếu không mã hóa
                backup_filename = f"{name}_auto_backup_{timestamp}{ext}"
                backup_path = os.path.join(backup_dir, backup_filename)
                
                # Lưu file
                df.to_excel(backup_path, index=False)
            
            # Lưu metadata vào file json riêng để dễ đọc
            metadata_path = os.path.splitext(backup_path)[0] + "_info.json"
            with open(metadata_path, 'w', encoding='utf-8') as f:
                json.dump(backup_info, f, ensure_ascii=False, indent=2)
                
            print(f"Đã tự động sao lưu dữ liệu vào: {backup_path}")
            
        except Exception as e:
            print(f"Lỗi khi tự động sao lưu: {str(e)}")
            traceback.print_exc()

def on_closing():
    """Xử lý khi đóng ứng dụng"""
    auto_backup_on_exit()
    root.destroy()

def refresh_ui():
    """Làm mới giao diện sau khi đọc dữ liệu"""
    global df, tree, status_label
    
    if df is not None and not df.empty:
        # Xóa dữ liệu cũ trong treeview
        for item in tree.get_children():
            tree.delete(item)
        
        # Hiển thị tối đa 100 học sinh đầu tiên để không làm chậm giao diện
        display_limit = 100 if len(df) > 100 else len(df)
        
        # Tìm các cột phù hợp
        name_col = find_matching_column(df, config['columns']['name'])
        exam_code_col = find_matching_column(df, config['columns']['exam_code']) or 'Mã đề'
        score_col = find_matching_column(df, config['columns']['score']) or 'Điểm'
        
        # Hàm áp dụng màu theo điểm
        def apply_color_by_score(item_id, score_value):
            """Tô màu dòng theo điểm số"""
            try:
                if pd.notna(score_value):
                    score = float(score_value)
                    if score < 5:
                        tree.item(item_id, tags=('low_score',))
                    elif score < 7:
                        tree.item(item_id, tags=('medium_score',))
                    else:
                        tree.item(item_id, tags=('high_score',))
                else:
                    tree.item(item_id, tags=('no_score',))
            except:
                tree.item(item_id, tags=('no_score',))
        
        for i, (_, row) in enumerate(df.head(display_limit).iterrows()):
            values = [
                row[name_col] if name_col in row else "",
                row[exam_code_col] if exam_code_col in row and pd.notna(row[exam_code_col]) else "",
                row[score_col] if score_col in row and pd.notna(row[score_col]) else ""
            ]
            item_id = tree.insert('', 'end', values=values)
            
            # Áp dụng màu theo điểm ngay khi hiển thị
            if score_col in row:
                apply_color_by_score(item_id, row[score_col])
            
        # Hiển thị thông báo nếu chỉ hiển thị một phần
        if len(df) > display_limit:
            tree.insert('', 'end', values=(f"--- Hiển thị {display_limit}/{len(df)} học sinh. Nhập từ khóa để tìm kiếm ---", "", ""))
        
        # Cập nhật status label nếu file_path tồn tại
        if 'file_path' in globals() and file_path:
            status_label.config(
                text=f"Dữ liệu đang hiển thị: {os.path.basename(file_path)} ({len(df)} học sinh)",

            )
            
        # Cập nhật thống kê
        update_stats()
        
        # Cập nhật điểm cao nhất/thấp nhất
        update_score_extremes()
        
        # Kiểm tra và xác minh các cột dữ liệu cần thiết
        verify_required_columns(df)
    else:
        # Nếu không có dữ liệu, hiển thị thông báo
        status_label.config(text="Chưa tải file Excel")
        for item in tree.get_children():
            tree.delete(item)
        tree.insert('', 'end', values=("Không có dữ liệu để hiển thị. Vui lòng tải file Excel.", "", ""))

def verify_required_columns(dataframe):
    """Kiểm tra xem dataframe có các cột cần thiết không"""
    # Kiểm tra dataframe có tồn tại không
    if dataframe is None:
        ToastNotification.show(
            "❌ Dữ liệu trống\n"
            "💡 Kiểm tra:\n"
            "  • File Excel có dữ liệu không?\n"
            "  • Sheet đầu tiên có đúng không?",
            "error")
        return False
        
    # Kiểm tra dataframe có rỗng không
    if dataframe.empty:
        ToastNotification.show(
            "❌ File Excel không có dữ liệu\n"
            "💡 Vui lòng thêm dữ liệu vào file rồi mở lại",
            "error")
        return False

    # Lấy tên các cột cần thiết từ config
    required_names = config['columns'].values()
    
    # Danh sách cột bị thiếu
    missing_columns = []
    
    # Kiểm tra từng cột
    for column_name in required_names:
        if column_name not in dataframe.columns:
            missing_columns.append(column_name)
    
    # Nếu có cột bị thiếu
    if missing_columns:
        ToastNotification.show(
            f"❌ Thiếu cột dữ liệu: {', '.join(missing_columns)}\n"
            f"💡 Giải pháp:\n"
            f"  • Kiểm tra tên cột trong file Excel\n"
            f"  • Hoặc vào File > Cài đặt để cấu hình lại tên cột",
            "error")
        return False
    
    return True

def ensure_required_columns(df):
    """
    Đảm bảo các cột cần thiết luôn tồn tại trong DataFrame
    
    Args:
        df (DataFrame): DataFrame cần kiểm tra
        
    Returns:
        DataFrame: DataFrame với các cột cần thiết đã thêm nếu cần
    """
    if df is None:
        # Tạo DataFrame rỗng với các cột cần thiết
        return pd.DataFrame({
            config['columns']['name']: [],
            'Điểm': [],
            'Mã đề': []
        })
    
    # Tạo bản sao để tránh thay đổi DataFrame gốc
    df_copy = df.copy()
    
    # Đảm bảo có cột tên học sinh
    if config['columns']['name'] not in df_copy.columns:
        name_col = find_matching_column(df_copy, 'tên học sinh')
        if name_col:
            # Nếu tìm thấy cột tương tự, đổi tên nó
            df_copy = df_copy.rename(columns={name_col: config['columns']['name']})
        else:
            # Nếu không, tạo mới
            df_copy[config['columns']['name']] = ["Học sinh " + str(i+1) for i in range(len(df_copy))]
    
    # Đảm bảo có cột điểm
    if 'Điểm' not in df_copy.columns:
        score_col = find_matching_column(df_copy, 'điểm')
        if score_col:
            df_copy = df_copy.rename(columns={score_col: 'Điểm'})
        else:
            df_copy['Điểm'] = None
    
    # Đảm bảo có cột mã đề
    if 'Mã đề' not in df_copy.columns:
        exam_code_col = find_matching_column(df_copy, 'mã đề')
        if exam_code_col:
            df_copy = df_copy.rename(columns={exam_code_col: 'Mã đề'})
        else:
            df_copy['Mã đề'] = None
        
    return df_copy

def create_ui():
    global status_label, entry_student_name, tree
    global entry_correct_count, entry_exam_code
    global entry_max_questions, score_per_q_label
    global score_info_label, entry_direct_score, stats_label
    global highest_score_label, lowest_score_label  # Thêm biến cho điểm cao/thấp
    global progress_bar, progress_label, high_score_count, medium_score_count, low_score_count, no_score_count  # Dashboard widgets
    global undo_button, redo_button  # Undo/Redo buttons
    
    # Khởi tạo style
    style = ttk.Style()
    
    # Áp dụng style theo cấu hình
    style = ui_utils.apply_styles(config, style, root)
    
    # Thiết lập responsive cho cửa sổ
    ui_utils.init_responsive_settings(root, config)
    
    # ========== HEADER HIỆN ĐẠI ==========
    header_frame = tk.Frame(root, bg=config['ui']['theme']['primary'], height=55)
    header_frame.pack(fill="x")
    header_frame.pack_propagate(False)
    
    header_content = tk.Frame(header_frame, bg=config['ui']['theme']['primary'])
    header_content.pack(fill="both", expand=True, padx=20, pady=12)
    
    # Title
    tk.Label(header_content, 
            text="📊 Quản Lý Điểm Học Sinh",
            font=(config['ui']['font_family'], 15, 'bold'),
            bg=config['ui']['theme']['primary'],
            fg='white').pack(side="left")
    
    # Config số câu ở giữa header
    config_container = tk.Frame(header_content, bg=config['ui']['theme']['primary'])
    config_container.pack(side="left", padx=30)
    
    tk.Label(config_container, text="⚙️ Số câu:",
            font=(config['ui']['font_family'], 10),
            bg=config['ui']['theme']['primary'],
            fg='white').pack(side="left", padx=(0, 5))
    
    entry_max_questions = tk.Entry(config_container, width=8,
                                  font=(config['ui']['font_family'], 10),
                                  justify='center',
                                  bg='#FFFFFF',
                                  fg='#1A202C',
                                  relief='flat',
                                  bd=0)
    entry_max_questions.insert(0, str(config['max_questions']))
    entry_max_questions.pack(side="left", padx=(0, 5))
    
    save_config_btn = tk.Button(config_container, text="💾",
                               font=(config['ui']['font_family'], 10),
                               command=update_config,
                               bg=config['ui']['theme']['success'],
                               fg='white',
                               relief='flat',
                               bd=0,
                               padx=8,
                               cursor='hand2')
    save_config_btn.pack(side="left")
    
    score_per_q_label = tk.Label(config_container, 
                                text=f"({config['score_per_question']}đ/câu)",
                                font=(config['ui']['font_family'], 8),
                                bg=config['ui']['theme']['primary'],
                                fg='#E0E0E0')
    score_per_q_label.pack(side="left", padx=(5, 0))
    
    # Right side
    right_container = tk.Frame(header_content, bg=config['ui']['theme']['primary'])
    right_container.pack(side="right")
    
    version_display = version_utils.get_version_display()
    tk.Label(right_container, text=version_display,
            font=(config['ui']['font_family'], 8),
            bg=config['ui']['theme']['primary'],
            fg='#E0E0E0').pack(side="left", padx=(0, 12))
    
    dark_mode_frame = ui_utils.create_dark_mode_switch(right_container, config, style, root, save_config)
    dark_mode_frame.pack(side="left")
    
    # ========== STATUSBAR Ở DƯỚI CÙNG (PACK TRƯỚC ĐỂ HIỆN) ==========
    statusbar = tk.Frame(root, bg=config['ui']['theme'].get('statusbar_bg', config['ui']['theme']['card']), 
                        height=28, relief='flat')
    statusbar.pack(side='bottom', fill='x')
    statusbar.pack_propagate(False)
    
    status_label = tk.Label(statusbar, text="📌 Sẵn sàng - Chưa tải file", 
                          bg=config['ui']['theme'].get('statusbar_bg', config['ui']['theme']['card']),
                          fg=config['ui']['theme'].get('statusbar_text', config['ui']['theme']['text']),
                          font=(config['ui']['font_family'], 9),
                          anchor='w',
                          padx=15)
    status_label.pack(side='left', fill='both', expand=True)
    
    # Version info ở bên phải statusbar
    version_display = version_utils.get_version_display()
    version_label = tk.Label(statusbar, text=f"v{version_display}", 
                            bg=config['ui']['theme'].get('statusbar_bg', config['ui']['theme']['card']),
                            fg=config['ui']['theme'].get('statusbar_text', '#A0AEC0'),
                            font=(config['ui']['font_family'], 8),
                            padx=15)
    version_label.pack(side='right')
    
    # Helper function cho refresh
    def refresh_data():
        global df, file_path
        if file_path and os.path.exists(file_path):
            try:
                df = pd.read_excel(file_path)
                search_student()
                update_stats()
                update_score_extremes()
                status_label.config(text=f"✅ Đã làm mới dữ liệu")
                ToastNotification.show("✅ Đã làm mới dữ liệu", "success")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể đọc file: {str(e)}")
        else:
            ToastNotification.show("ℹ️ Chưa chọn file", "info")
    
    # ========== MAIN CONTAINER - 2 CỘT ==========
    main_container = ttk.Frame(root, style='TFrame')
    main_container.pack(fill="both", expand=True, padx=8, pady=8)
    
    # Configure grid weights - 70% trái, 30% phải với minsize đảm bảo không bị khuất
    main_container.grid_columnconfigure(0, weight=7, minsize=800)
    main_container.grid_columnconfigure(1, weight=3, minsize=360)
    main_container.grid_rowconfigure(0, weight=1)
    
    # ========== CỘT TRÁI - DANH SÁCH HỌC SINH ==========
    left_column = ttk.Frame(main_container, style='TFrame')
    left_column.grid(row=0, column=0, sticky='nsew', padx=(0, 6))
    
    # Search bar
    search_frame = ttk.LabelFrame(left_column, text="🔍 Tìm Kiếm & Lọc", padding=8)
    search_frame.pack(fill="x", pady=(0, 6))
    
    search_input_frame = ttk.Frame(search_frame)
    search_input_frame.pack(fill="x")
    
    entry_student_name = ttk.Entry(search_input_frame, 
                                  font=(config['ui']['font_family'], 11))
    entry_student_name.pack(side="left", fill="x", expand=True, padx=(0, 6))
    entry_student_name.bind("<KeyRelease>", delayed_search)
    
    ttk.Button(search_input_frame, text="➕ Thêm", 
              command=add_student, style='Success.TButton', width=10).pack(side="left")
    
    # Undo/Redo buttons - thêm global để update được
    global undo_button, redo_button
    undo_redo_frame = ttk.Frame(search_frame)
    undo_redo_frame.pack(fill="x", pady=(4, 0))
    
    undo_button = ttk.Button(undo_redo_frame, text="⏪ Hoàn tác", 
                            command=perform_undo, width=12, state='disabled')
    undo_button.pack(side="left", padx=(0, 4))
    ToolTip(undo_button, "Hoàn tác thao tác cuối (Ctrl+Z)")
    
    redo_button = ttk.Button(undo_redo_frame, text="⏩ Làm lại", 
                            command=perform_redo, width=12, state='disabled')
    redo_button.pack(side="left")
    ToolTip(redo_button, "Làm lại thao tác đã hoàn tác (Ctrl+Y)")
    
    # Filter nâng cao
    filter_frame = ttk.Frame(search_frame)
    filter_frame.pack(fill="x", pady=(6, 0))
    
    ttk.Label(filter_frame, text="🎯 Lọc nhanh:", 
             font=(config['ui']['font_family'], 9)).pack(side="left", padx=(0, 5))
    
    def apply_quick_filter(filter_type):
        """Áp dụng filter nhanh"""
        global df
        if df is None or df.empty:
            return
        
        # Xóa treeview
        for item in tree.get_children():
            tree.delete(item)
        
        # Tìm cột
        name_col = find_matching_column(df, config['columns']['name'])
        score_col = 'Điểm'
        exam_col = 'Mã đề'
        
        # Áp dụng filter
        if filter_type == 'all':
            filtered_df = df
        elif filter_type == 'high':  # Điểm >= 7
            filtered_df = df[df[score_col] >= 7]
        elif filter_type == 'medium':  # 5 <= Điểm < 7
            filtered_df = df[(df[score_col] >= 5) & (df[score_col] < 7)]
        elif filter_type == 'low':  # Điểm < 5
            filtered_df = df[df[score_col] < 5]
        elif filter_type == 'no_score':  # Chưa có điểm
            filtered_df = df[df[score_col].isna()]
        else:
            filtered_df = df
        
        # Hiển thị kết quả
        display_limit = 200 if len(filtered_df) > 200 else len(filtered_df)
        
        def apply_color(item_id, score_val):
            try:
                if pd.notna(score_val):
                    score = float(score_val)
                    if score < 5:
                        tree.item(item_id, tags=('low_score',))
                    elif score < 7:
                        tree.item(item_id, tags=('medium_score',))
                    else:
                        tree.item(item_id, tags=('high_score',))
                else:
                    tree.item(item_id, tags=('no_score',))
            except:
                tree.item(item_id, tags=('no_score',))
        
        for _, row in filtered_df.head(display_limit).iterrows():
            values = [
                row[name_col] if name_col in row else "",
                row[exam_col] if exam_col in row and pd.notna(row[exam_col]) else "",
                f"{row[score_col]:.2f}" if score_col in row and pd.notna(row[score_col]) else "Chưa có điểm"
            ]
            item_id = tree.insert('', 'end', values=values)
            if score_col in row:
                apply_color(item_id, row[score_col])
        
        if len(filtered_df) > display_limit:
            tree.insert('', 'end', values=(f"--- Hiển thị {display_limit}/{len(filtered_df)} kết quả ---", "", ""))
        
        # Hiển thị toast
        ToastNotification.show(f"Đã lọc: {len(filtered_df)} học sinh", "info")
    
    ttk.Button(filter_frame, text="Tất cả", 
              command=lambda: apply_quick_filter('all'), width=8).pack(side="left", padx=2)
    ttk.Button(filter_frame, text="🟢 ≥7", 
              command=lambda: apply_quick_filter('high'), width=8).pack(side="left", padx=2)
    ttk.Button(filter_frame, text="🟡 5-7", 
              command=lambda: apply_quick_filter('medium'), width=8).pack(side="left", padx=2)
    ttk.Button(filter_frame, text="🔴 <5", 
              command=lambda: apply_quick_filter('low'), width=8).pack(side="left", padx=2)
    ttk.Button(filter_frame, text="⚪ Chưa", 
              command=lambda: apply_quick_filter('no_score'), width=8).pack(side="left", padx=2)
    
    ttk.Label(search_frame, text="💡 Ctrl+F để tìm nhanh | Click header để sắp xếp", 
             font=(config['ui']['font_family'], 8),
             foreground=config['ui']['theme']['text_secondary']).pack(pady=(4, 0))
    
    # Danh sách học sinh - CHÍNH
    list_frame = ttk.LabelFrame(left_column, text="👥 Danh Sách Học Sinh", padding=8)
    list_frame.pack(fill="both", expand=True)
    
    # Treeview
    columns = ('name', 'exam_code', 'score')
    tree = ttk.Treeview(list_frame, columns=columns, show='headings', style='Treeview')
    
    # Biến lưu trạng thái sort
    sort_column = {'col': None, 'reverse': False}
    
    def sort_treeview(col):
        """Sắp xếp treeview theo cột được click"""
        global df
        
        if df is None or df.empty:
            return
        
        # Đảo ngược thứ tự nếu click vào cùng cột
        if sort_column['col'] == col:
            sort_column['reverse'] = not sort_column['reverse']
        else:
            sort_column['col'] = col
            sort_column['reverse'] = False
        
        # Map column name to dataframe column
        col_map = {
            'name': find_matching_column(df, config['columns']['name']),
            'exam_code': 'Mã đề',
            'score': 'Điểm'
        }
        
        df_col = col_map.get(col)
        if df_col and df_col in df.columns:
            # Sort dataframe
            if col == 'score':
                # Sắp xếp điểm: đưa NaN xuống cuối
                df_sorted = df.sort_values(
                    by=df_col, 
                    ascending=not sort_column['reverse'],
                    na_position='last'
                )
            else:
                df_sorted = df.sort_values(
                    by=df_col, 
                    ascending=not sort_column['reverse']
                )
            
            # Xóa treeview hiện tại
            for item in tree.get_children():
                tree.delete(item)
            
            # Hiển thị lại với thứ tự mới
            display_limit = 100 if len(df_sorted) > 100 else len(df_sorted)
            
            # Hàm áp dụng màu theo điểm
            def apply_color(item_id, score_val):
                try:
                    if pd.notna(score_val):
                        score = float(score_val)
                        if score < 5:
                            tree.item(item_id, tags=('low_score',))
                        elif score < 7:
                            tree.item(item_id, tags=('medium_score',))
                        else:
                            tree.item(item_id, tags=('high_score',))
                    else:
                        tree.item(item_id, tags=('no_score',))
                except:
                    tree.item(item_id, tags=('no_score',))
            
            name_col = col_map['name']
            exam_col = col_map['exam_code']
            score_col = col_map['score']
            
            for _, row in df_sorted.head(display_limit).iterrows():
                values = [
                    row[name_col] if name_col in row else "",
                    row[exam_col] if exam_col in row and pd.notna(row[exam_col]) else "",
                    f"{row[score_col]:.2f}" if score_col in row and pd.notna(row[score_col]) else "Chưa có điểm"
                ]
                item_id = tree.insert('', 'end', values=values)
                
                # Áp dụng màu
                if score_col in row:
                    apply_color(item_id, row[score_col])
            
            if len(df_sorted) > display_limit:
                tree.insert('', 'end', values=(f"--- Hiển thị {display_limit}/{len(df_sorted)} học sinh ---", "", ""))
            
            # Cập nhật icon cho header
            arrow = " ▼" if sort_column['reverse'] else " ▲"
            tree.heading('name', text=f"👤 {config['columns']['name']}{arrow if col == 'name' else ''}")
            tree.heading('exam_code', text=f"📋 Mã Đề{arrow if col == 'exam_code' else ''}")
            tree.heading('score', text=f"📊 Điểm{arrow if col == 'score' else ''}")
    
    tree.heading('name', text=f"👤 {config['columns']['name']}", command=lambda: sort_treeview('name'))
    tree.heading('exam_code', text=f"📋 Mã Đề", command=lambda: sort_treeview('exam_code'))
    tree.heading('score', text=f"📊 Điểm", command=lambda: sort_treeview('score'))
    
    tree.column('name', width=750, minwidth=400)
    tree.column('exam_code', width=150, minwidth=100, anchor='center')
    tree.column('score', width=150, minwidth=100, anchor='center')
    
    # Định nghĩa màu cho các tags (tô màu theo điểm)
    tree.tag_configure('high_score', background='#D1FAE5', foreground='#065F46')  # Xanh lá nhạt
    tree.tag_configure('medium_score', background='#FEF3C7', foreground='#92400E')  # Vàng nhạt
    tree.tag_configure('low_score', background='#FEE2E2', foreground='#991B1B')  # Đỏ nhạt
    tree.tag_configure('no_score', background='#F3F4F6', foreground='#6B7280')  # Xám nhạt
    
    vsb = ttk.Scrollbar(list_frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(list_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    
    tree.grid(column=0, row=0, sticky='nsew', padx=2, pady=2)
    vsb.grid(column=1, row=0, sticky='ns', pady=2)
    hsb.grid(column=0, row=1, sticky='ew', padx=2)
    
    list_frame.grid_columnconfigure(0, weight=1)
    list_frame.grid_rowconfigure(0, weight=1)
    
    # ========== CỘT PHẢI - CONTROLS với Scrollable Canvas ==========
    right_column = ttk.Frame(main_container, style='TFrame')
    right_column.grid(row=0, column=1, sticky='nsew')
    
    # Tạo Canvas với scrollbar cho cột phải
    right_canvas = tk.Canvas(right_column, bg=config['ui']['theme']['background'], 
                            highlightthickness=0, width=360)
    right_scrollbar = ttk.Scrollbar(right_column, orient="vertical", command=right_canvas.yview)
    scrollable_right_frame = ttk.Frame(right_canvas, style='TFrame')
    
    scrollable_right_frame.bind(
        "<Configure>",
        lambda e: right_canvas.configure(scrollregion=right_canvas.bbox("all"))
    )
    
    canvas_window = right_canvas.create_window((0, 0), window=scrollable_right_frame, anchor="nw", width=360)
    right_canvas.configure(yscrollcommand=right_scrollbar.set)
    
    # Update canvas window width khi canvas resize
    def on_canvas_resize(event):
        right_canvas.itemconfig(canvas_window, width=event.width)
    right_canvas.bind("<Configure>", on_canvas_resize)
    
    right_canvas.pack(side="left", fill="both", expand=True)
    right_scrollbar.pack(side="right", fill="y")
    
    # Bind mousewheel cho scroll mượt - CHỈ KHI CHUỘT Ở TRONG VÙNG PHẢI
    def on_mousewheel(event):
        right_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def bind_mousewheel(event):
        right_canvas.bind_all("<MouseWheel>", on_mousewheel)
    
    def unbind_mousewheel(event):
        right_canvas.unbind_all("<MouseWheel>")
    
    # Chỉ scroll khi chuột vào vùng cột phải
    right_canvas.bind("<Enter>", bind_mousewheel)
    right_canvas.bind("<Leave>", unbind_mousewheel)
    scrollable_right_frame.bind("<Enter>", bind_mousewheel)
    scrollable_right_frame.bind("<Leave>", unbind_mousewheel)
    
    # Thống kê
    stats_card = ttk.LabelFrame(scrollable_right_frame, text="📊 Thống Kê Chi Tiết", padding=10)
    stats_card.pack(fill="x", pady=(0, 6), padx=4)
    
    # Progress bar cho tỷ lệ hoàn thành
    progress_frame = ttk.Frame(stats_card)
    progress_frame.pack(fill="x", pady=(0, 8))
    
    ttk.Label(progress_frame, text="Tiến độ nhập điểm:", 
             font=(config['ui']['font_family'], 9)).pack(anchor="w")
    
    progress_bar = ttk.Progressbar(progress_frame, mode='determinate', length=200)
    progress_bar.pack(fill="x", pady=(2, 2))
    
    progress_label = ttk.Label(progress_frame, text="0%", 
                              font=(config['ui']['font_family'], 9, 'bold'),
                              foreground=config['ui']['theme']['primary'])
    progress_label.pack(anchor="w")
    
    # Thống kê tổng quan
    stats_label = ttk.Label(stats_card, text="0/0 có điểm", 
                          font=(config['ui']['font_family'], 10, 'bold'))
    stats_label.pack(pady=4)
    
    # Phân loại theo điểm
    category_frame = ttk.Frame(stats_card)
    category_frame.pack(fill="x", pady=4)
    
    high_score_count = ttk.Label(category_frame, text="🟢 Giỏi (≥7): 0", 
                                 font=(config['ui']['font_family'], 9),
                                 foreground='#065F46')
    high_score_count.pack(anchor="w", pady=1)
    
    medium_score_count = ttk.Label(category_frame, text="🟡 Khá (5-7): 0", 
                                   font=(config['ui']['font_family'], 9),
                                   foreground='#92400E')
    medium_score_count.pack(anchor="w", pady=1)
    
    low_score_count = ttk.Label(category_frame, text="🔴 Yếu (<5): 0", 
                                font=(config['ui']['font_family'], 9),
                                foreground='#991B1B')
    low_score_count.pack(anchor="w", pady=1)
    
    no_score_count = ttk.Label(category_frame, text="⚪ Chưa có: 0", 
                               font=(config['ui']['font_family'], 9),
                               foreground='#6B7280')
    no_score_count.pack(anchor="w", pady=1)
    
    # Separator
    ttk.Separator(stats_card, orient='horizontal').pack(fill="x", pady=6)
    
    # Điểm cao/thấp nhất
    highest_score_label = ttk.Label(stats_card, text="🏆 Cao nhất: N/A", 
                                  font=(config['ui']['font_family'], 9),
                                  foreground=config['ui']['theme']['success'])
    highest_score_label.pack(pady=2)
    
    lowest_score_label = ttk.Label(stats_card, text="📉 Thấp nhất: N/A", 
                                 font=(config['ui']['font_family'], 9),
                                 foreground=config['ui']['theme']['warning'])
    lowest_score_label.pack(pady=2)
    
    # File Management
    file_card = ttk.LabelFrame(scrollable_right_frame, text="📁 File", padding=6)
    file_card.pack(fill="x", pady=(0, 6), padx=4)
    
    file_btn_frame = ttk.Frame(file_card)
    file_btn_frame.pack(fill="x")
    
    ttk.Button(file_btn_frame, text="📂 Mở", 
              command=select_file, width=12).pack(fill="x", pady=2)
    ttk.Button(file_btn_frame, text="🔄 Làm Mới", 
              command=refresh_data, width=12).pack(fill="x", pady=2)
    ttk.Button(file_btn_frame, text="💾 Sao Lưu", 
              command=backup_data, width=12).pack(fill="x", pady=2)
    ttk.Button(file_btn_frame, text="📄 Báo Cáo", 
              command=generate_report, style='Accent.TButton', width=12).pack(fill="x", pady=2)
    
    # Nhập điểm
    score_card = ttk.LabelFrame(scrollable_right_frame, text="✏️ Nhập Điểm", padding=6)
    score_card.pack(fill="x", padx=4)
    
    ttk.Label(score_card, text="🔢 Mã Đề:", 
             font=(config['ui']['font_family'], 9)).pack()
    entry_exam_code = ttk.Combobox(score_card, width=14, values=config['exam_codes'],
                                 font=(config['ui']['font_family'], 10))
    entry_exam_code.pack(pady=(2, 8))
    
    ttk.Label(score_card, text="✅ Số Câu Đúng:", 
             font=(config['ui']['font_family'], 9)).pack()
    entry_correct_count = ttk.Entry(score_card, width=14,
                                  font=(config['ui']['font_family'], 11),
                                  justify='center')
    entry_correct_count.pack(pady=2)
    entry_correct_count.bind("<Return>", calculate_score)
    
    ttk.Button(score_card, text="💯 Tính Điểm", 
              command=calculate_score, width=14).pack(fill="x", pady=(4, 8))
    
    ttk.Label(score_card, text="🎯 Hoặc Nhập Điểm:", 
             font=(config['ui']['font_family'], 9)).pack()
    entry_direct_score = ttk.Entry(score_card, width=14,
                               font=(config['ui']['font_family'], 11),
                               justify='center')
    entry_direct_score.pack(pady=2)
    entry_direct_score.bind("<Return>", calculate_score_direct)
    
    ttk.Button(score_card, text="✓ Lưu Điểm", 
              command=calculate_score_direct, style='Success.TButton', width=14).pack(fill="x", pady=(4, 0))
    
    ttk.Label(score_card, text="⌨️ Ctrl+G / Ctrl+D", 
             font=(config['ui']['font_family'], 8),
             foreground=config['ui']['theme']['text_secondary']).pack(pady=(4, 0))
    
    # Menu
    menubar = tk.Menu(root, 
                     background=config['ui']['theme']['card'],
                     foreground=config['ui']['theme']['text'],
                     activebackground=config['ui']['theme']['primary'],
                     activeforeground='white',
                     borderwidth=0)
    root.config(menu=menubar)
    
    settings_menu = tk.Menu(menubar, tearoff=0,
                           background=config['ui']['theme']['card'],
                           foreground=config['ui']['theme']['text'],
                           activebackground=config['ui']['theme']['primary'],
                           activeforeground='white')
    menubar.add_cascade(label="⚙️ Cài đặt", menu=settings_menu)
    settings_menu.add_command(label="⌨️ Phím tắt", command=customize_shortcuts)
    settings_menu.add_command(label="🔢 Mã đề", command=customize_exam_codes)
    settings_menu.add_command(label="📝 Tên cột", command=customize_columns)
    settings_menu.add_command(label="🔒 Bảo mật", command=customize_security)
    settings_menu.add_separator()
    settings_menu.add_command(label="📡 Kênh cập nhật", command=choose_update_channel)
    settings_menu.add_command(label="🌓 Chế độ tối/sáng", command=lambda: toggle_theme(style))
    
    function_menu = tk.Menu(menubar, tearoff=0,
                           background=config['ui']['theme']['card'],
                           foreground=config['ui']['theme']['text'],
                           activebackground=config['ui']['theme']['primary'],
                           activeforeground='white')
    menubar.add_cascade(label="🔧 Chức năng", menu=function_menu)
    function_menu.add_command(label="📊 Biểu đồ phân phối", command=show_score_distribution)
    function_menu.add_command(label="📄 Xuất báo cáo PDF", command=generate_report)
    function_menu.add_separator()
    function_menu.add_command(label="⬆️ Kiểm tra cập nhật", command=lambda: check_for_updates_wrapper(True))
    function_menu.add_command(label="ℹ️ Thông tin", command=show_about)

def toggle_theme(style):
    """Chuyển đổi giữa chế độ sáng và tối"""
    global config
    print(f"Trạng thái dark mode trước khi đổi: {config['ui']['dark_mode']}")
    
    # Đảo ngược trạng thái dark mode trước khi gọi hàm
    config['ui']['dark_mode'] = not config['ui']['dark_mode']
    print(f"Trạng thái dark mode sau khi đổi: {config['ui']['dark_mode']}")
    
    # Cập nhật giao diện với trạng thái mới
    config = ui_utils.toggle_dark_mode(config, style, root)
    
    # In thông tin cấu hình sau khi cập nhật
    print(f"Trạng thái dark mode sau khi toggle_dark_mode: {config['ui']['dark_mode']}")
    print(f"Theme hiện tại: {config['ui']['theme']['background']}")
    
    # Lưu cấu hình
    save_config()
    
    # Cập nhật màu sắc của menu
    menubar = root.option_get('menu', '')
    if menubar:
        menu = tk._default_root.nametowidget(menubar)
        if menu:
            menu.configure(background=config['ui']['theme']['background'], 
                         foreground=config['ui']['theme']['text'])
            for submenu in ['Cài đặt', 'Chức năng']:
                try:
                    submenu_widget = menu.nametowidget(f"{menubar}.{submenu}")
                    submenu_widget.configure(background=config['ui']['theme']['background'], 
                                          foreground=config['ui']['theme']['text'])
                except:
                    pass
    
    # Force update để đảm bảo thay đổi được áp dụng
    root.update_idletasks()
    
    # Cập nhật màu sắc của tooltip
    if hasattr(ToolTip, 'tooltip') and ToolTip.tooltip:
        for label in ToolTip.tooltip.winfo_children():
            if isinstance(label, ttk.Label):
                label.configure(background=config['ui']['theme']['card'], 
                              foreground=config['ui']['theme']['text'])

def show_about():
    """Hiển thị thông tin về phần mềm"""
    # Lấy thông tin phiên bản và changelog
    version_info = version_utils.load_version_info()
    version = version_info.get('version', '1.0.0')
    is_dev = version_info.get('is_dev', False)
    code_name = version_info.get('code_name', '')
    build_date = version_info.get('build_date', '')
    
    # Lấy changelog cho phiên bản hiện tại
    changes = version_utils.get_changelog_for_version(version)
    
    # Tạo chuỗi thông tin
    about_text = f"""Phần mềm Quản lý Điểm Học Sinh
Version: {version}{' dev' if is_dev else ''}
{code_name}
Ngày phát hành: {build_date}

Changelog v{version}:
""" + "\n".join(f"• {change}" for change in changes) + """

Xem thêm thông tin và cập nhật mới nhất tại GitHub.
    """
    
    about_window = tk.Toplevel(root)
    about_window.title("Thông tin")
    about_window.geometry("500x400")
    about_window.transient(root)
    about_window.grab_set()
    
    # Áp dụng theme hiện tại
    about_window.configure(bg=config['ui']['theme']['background'])
    
    # Thêm scrollable text widget
    text_frame = ttk.Frame(about_window)
    text_frame.pack(fill="both", expand=True, padx=15, pady=15)
    
    # Tạo Text widget với màu sắc phù hợp với theme
    about_text_widget = tk.Text(text_frame, wrap="word", 
                              font=(config['ui']['font_family'], 11),
                              bg=config['ui']['theme']['card'],
                              fg=config['ui']['theme']['text'])
    about_text_widget.pack(side="left", fill="both", expand=True)
    
    scroll = ttk.Scrollbar(text_frame, command=about_text_widget.yview)
    scroll.pack(side="right", fill="y")
    about_text_widget.config(yscrollcommand=scroll.set)
    
    # Chèn text
    about_text_widget.insert("1.0", about_text)
    about_text_widget.config(state="disabled")  # Không cho phép chỉnh sửa
    
    # Thêm nút kiểm tra cập nhật
    button_frame = ttk.Frame(about_window)
    button_frame.pack(fill="x", padx=15, pady=15)
    
    ttk.Button(button_frame, text="Kiểm tra cập nhật", 
              command=lambda: check_for_updates_wrapper(True)).pack(side="left", padx=5)
    
    ttk.Button(button_frame, text="Đóng", 
              command=about_window.destroy).pack(side="right", padx=5)

def check_for_updates_wrapper(show_notification=True):
    """
    Kiểm tra phiên bản mới trên GitHub
    
    Args:
        show_notification (bool): Có hiển thị thông báo khi không có phiên bản mới không
    
    Returns:
        bool: True nếu có phiên bản mới, False nếu không
    """
    # Sử dụng hàm được tách ra từ module check_for_updates để sửa lỗi
    try:
        from check_for_updates import check_for_updates
        return check_for_updates(root, status_label, file_path, config, save_config, show_notification)
    except ImportError:
        # Nếu không tìm thấy module fixed, dùng module gốc
        from check_for_updates import check_for_updates
        return check_for_updates(root, status_label, file_path, config, save_config, show_notification)

def check_updates_async():
    """Kiểm tra cập nhật trong luồng riêng biệt để không làm đóng băng giao diện"""
    # Sử dụng hàm từ module check_for_updates
    print("import_score: Gọi check_updates_async - kiểm tra cập nhật tự động khi khởi động")
    try:
        from check_for_updates import check_updates_async
        # Gọi hàm check_updates_async với các tham số cần thiết
        check_updates_async(root, status_label, file_path, config, save_config)
    except ImportError:
        # Nếu không tìm thấy module fixed, dùng module gốc
        from check_for_updates import check_updates_async
        check_updates_async(root, status_label, file_path, config, save_config)
    print("import_score: Đã gọi check_updates_async từ module")

def update_stats():
    """Cập nhật các thống kê cơ bản với progress bar và phân loại chi tiết"""
    global progress_bar, progress_label, high_score_count, medium_score_count, low_score_count, no_score_count
    
    if df is None or df.empty:
        stats_label.config(text="Không có dữ liệu")
        progress_bar['value'] = 0
        progress_label.config(text="0%")
        high_score_count.config(text="🟢 Giỏi (≥7): 0")
        medium_score_count.config(text="🟡 Khá (5-7): 0")
        low_score_count.config(text="🔴 Yếu (<5): 0")
        no_score_count.config(text="⚪ Chưa có: 0")
        return
        
    total_students = len(df)
    score_column = config['columns']['score']
    
    # Đếm số học sinh có điểm
    students_with_scores = df[score_column].notna().sum()
    students_no_scores = total_students - students_with_scores
    
    # Tính phần trăm
    percentage = (students_with_scores / total_students * 100) if total_students > 0 else 0
    
    # Cập nhật progress bar
    progress_bar['value'] = percentage
    progress_label.config(text=f"{percentage:.1f}% ({students_with_scores}/{total_students})")
    
    # Phân loại theo điểm
    df_with_scores = df[df[score_column].notna()]
    high_count = len(df_with_scores[df_with_scores[score_column] >= 7])
    medium_count = len(df_with_scores[(df_with_scores[score_column] >= 5) & (df_with_scores[score_column] < 7)])
    low_count = len(df_with_scores[df_with_scores[score_column] < 5])
    
    # Cập nhật labels
    stats_label.config(text=f"{students_with_scores}/{total_students} đã có điểm")
    high_score_count.config(text=f"🟢 Giỏi (≥7): {high_count} ({high_count/total_students*100:.1f}%)" if total_students > 0 else "🟢 Giỏi (≥7): 0")
    medium_score_count.config(text=f"🟡 Khá (5-7): {medium_count} ({medium_count/total_students*100:.1f}%)" if total_students > 0 else "🟡 Khá (5-7): 0")
    low_score_count.config(text=f"🔴 Yếu (<5): {low_count} ({low_count/total_students*100:.1f}%)" if total_students > 0 else "🔴 Yếu (<5): 0")
    no_score_count.config(text=f"⚪ Chưa có: {students_no_scores} ({students_no_scores/total_students*100:.1f}%)" if total_students > 0 else "⚪ Chưa có: 0")

def update_score_extremes():
    """Cập nhật điểm cao nhất và thấp nhất"""
    if df is None or df.empty:
        highest_score_label.config(text="🏆 Cao nhất: N/A")
        lowest_score_label.config(text="📉 Thấp nhất: N/A")
        return
        
    score_column = config['columns']['score']
    name_column = config['columns']['name']
    
    # Lọc các hàng có điểm không phải NaN
    df_with_scores = df[df[score_column].notna()]
    
    # Nếu không có ai có điểm
    if df_with_scores.empty:
        highest_score_label.config(text="🏆 Cao nhất: N/A")
        lowest_score_label.config(text="📉 Thấp nhất: N/A")
        return
    
    # Tìm điểm cao nhất
    max_score = df_with_scores[score_column].max()
    max_students = df_with_scores[df_with_scores[score_column] == max_score][name_column].tolist()
    max_students_text = ", ".join(max_students[:2])
    if len(max_students) > 2:
        max_students_text += f" +{len(max_students) - 2}"
    
    # Tìm điểm thấp nhất
    min_score = df_with_scores[score_column].min()
    min_students = df_with_scores[df_with_scores[score_column] == min_score][name_column].tolist()
    min_students_text = ", ".join(min_students[:2])
    if len(min_students) > 2:
        min_students_text += f" +{len(min_students) - 2}"
    
    # Cập nhật giao diện
    highest_score_label.config(text=f"🏆 Cao nhất: {max_score:.2f} ({max_students_text})")
    lowest_score_label.config(text=f"📉 Thấp nhất: {min_score:.2f} ({min_students_text})")

def delayed_search(event=None):
    """Tìm kiếm sau một khoảng thời gian để tránh quá nhiều tìm kiếm liên tiếp"""
    global search_timer_id
    
    # Hủy timer cũ nếu có
    if search_timer_id is not None:
        root.after_cancel(search_timer_id)
    
    # Thiết lập timer mới
    search_timer_id = root.after(300, search_student)  # 300ms delay

def generate_report():
    """Tạo báo cáo PDF với thống kê và biểu đồ phân phối điểm"""
    global df
    
    if df is None or df.empty:
        messagebox.showwarning("Thông báo", "Không có dữ liệu để tạo báo cáo")
        return
        
    if 'Điểm' not in df.columns:
        messagebox.showwarning("Thông báo", "Không tìm thấy cột điểm trong dữ liệu")
        return
        
    # Lọc chỉ lấy học sinh có điểm
    scores_df = df[df['Điểm'].notna()].copy()
    
    if len(scores_df) == 0:
        messagebox.showwarning("Thông báo", "Chưa có học sinh nào có điểm")
        return
    
    try:
        # Hỏi nơi lưu file báo cáo
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Lưu báo cáo PDF"
        )
        
        if not file_path:
            return  # Người dùng hủy
            
        # Tạo cửa sổ hiển thị tiến trình tạo báo cáo
        progress_window = tk.Toplevel(root)
        progress_window.title("Đang tạo báo cáo PDF")
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
        progress_frame = ttk.Frame(progress_window, padding=15)
        progress_frame.pack(fill="both", expand=True)
        
        # Label hiển thị thông tin
        ttk.Label(progress_frame, 
               text="Đang tạo báo cáo PDF...",
               font=(config['ui']['font_family'], 12, 'bold'),
               foreground=config['ui']['theme']['primary']).pack(pady=5)
        
        # Frame chứa thông tin chi tiết
        info_frame = ttk.Frame(progress_frame)
        info_frame.pack(fill="x", pady=5)
        
        # Label hiển thị tên file
        file_label = ttk.Label(info_frame, 
                            text=f"File: {os.path.basename(file_path)}",
                            font=(config['ui']['font_family'], 10))
        file_label.pack(anchor="w", pady=2)
        
        # Label hiển thị số lượng học sinh
        students_label = ttk.Label(info_frame, 
                               text=f"Tổng số học sinh: {len(df)}, Có điểm: {len(scores_df)}",
                               font=(config['ui']['font_family'], 10))
        students_label.pack(anchor="w", pady=2)
        
        # Tạo style cho thanh tiến trình
        style = ttk.Style()
        style.configure("PDF.Horizontal.TProgressbar", 
                      troughcolor=config['ui']['theme']['background'],
                      background=config['ui']['theme']['primary'],
                      thickness=15)
        
        # Thanh tiến trình
        progress = ttk.Progressbar(progress_frame, 
                                orient="horizontal", 
                                length=400, 
                                mode="determinate",
)
        progress.pack(pady=10, fill="x")
        
        # Label hiển thị trạng thái
        status_label_progress = ttk.Label(progress_frame, 
                                      text="Đang chuẩn bị dữ liệu...", 
                                      font=(config['ui']['font_family'], 10))
        status_label_progress.pack(pady=5)
        
        # Cập nhật giao diện trước khi bắt đầu
        progress_window.update()
        
        # Hàm cập nhật tiến trình
        def update_progress(value, message):
            progress["value"] = value
            status_label_progress.config(text=message)
            progress_window.update_idletasks()
        
        # Hiển thị thông báo đang tạo báo cáo
        status_label.config(text="Đang tạo báo cáo PDF...")
        root.update()
        
        # Tạo tệp PDF
        update_progress(10, "Đang chuẩn bị dữ liệu...")
        
        with PdfPages(file_path) as pdf:
            # Trang 1: Thông tin chung
            update_progress(20, "Đang tạo trang thống kê chung...")
            plt.figure(figsize=(10, 12))
            plt.axis('off')
            
            # Tiêu đề báo cáo
            title_text = f"BÁO CÁO THỐNG KÊ ĐIỂM SỐ"
            subtitle_text = f"Ngày tạo: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
            
            plt.text(0.5, 0.98, title_text, fontsize=16, ha='center', va='top', fontweight='bold')
            plt.text(0.5, 0.95, subtitle_text, fontsize=12, ha='center', va='top')
            
            # Thông tin cơ bản
            info_text = [
                f"Tổng số học sinh: {len(df)}",
                f"Học sinh có điểm: {len(scores_df)} ({len(scores_df)/len(df):.1%})",
                f"Điểm trung bình: {scores_df['Điểm'].mean():.2f}",
                f"Điểm cao nhất: {scores_df['Điểm'].max():.2f}",
                f"Điểm thấp nhất: {scores_df['Điểm'].min():.2f}",
                f"Độ lệch chuẩn: {scores_df['Điểm'].std():.2f}"
            ]
            
            plt.text(0.1, 0.9, "Tổng quan:", fontsize=14, va='top', fontweight='bold')
            y_pos = 0.85
            for info in info_text:
                plt.text(0.1, y_pos, info, fontsize=12, va='top')
                y_pos -= 0.03
            
            update_progress(30, "Đang vẽ biểu đồ phân phối điểm...")
            # Biểu đồ phân phối điểm
            ax1 = plt.axes([0.1, 0.45, 0.8, 0.3])
            scores_df['Điểm'].hist(ax=ax1, bins=10, alpha=0.7, color='blue', edgecolor='black')
            ax1.set_title('Phân phối điểm số')
            ax1.set_xlabel('Điểm')
            ax1.set_ylabel('Số học sinh')
            ax1.grid(True, alpha=0.3)
            
            # Thêm đường trung bình
            mean_score = scores_df['Điểm'].mean()
            ax1.axvline(mean_score, color='red', linestyle='dashed', linewidth=1)
            ax1.text(mean_score + 0.1, ax1.get_ylim()[1]*0.9, f'TB: {mean_score:.2f}', color='red')
            
            update_progress(40, "Đang vẽ biểu đồ tỷ lệ đạt/không đạt...")
            # Biểu đồ tròn tỷ lệ đạt/không đạt
            ax2 = plt.axes([0.1, 0.1, 0.35, 0.25])
            pass_threshold = 5.0
            passed = (scores_df['Điểm'] >= pass_threshold).sum()
            failed = len(scores_df) - passed
            
            labels = ['Đạt (≥ 5.0)', 'Không đạt (< 5.0)']
            sizes = [passed, failed]
            colors = ['#66b3ff', '#ff9999']
            
            ax2.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)
            ax2.axis('equal')
            ax2.set_title('Tỷ lệ học sinh đạt/không đạt')
            
            update_progress(50, "Đang tạo bảng học sinh điểm cao...")
            # Bảng học sinh điểm cao nhất
            ax3 = plt.axes([0.55, 0.05, 0.35, 0.3])
            ax3.axis('tight')
            ax3.axis('off')
            
            # Lấy 5 học sinh có điểm cao nhất
            name_col = config['columns']['name']
            top_students = scores_df.sort_values('Điểm', ascending=False).head(5)
            top_data = [[student, score] for student, score in 
                       zip(top_students[name_col], top_students['Điểm'])]
            
            top_table = ax3.table(cellText=top_data, colLabels=[name_col, 'Điểm'],
                              loc='center', cellLoc='center')
            top_table.auto_set_font_size(False)
            top_table.set_fontsize(10)
            top_table.scale(1, 1.5)
            
            ax3.set_title('Học sinh điểm cao nhất')
            
            update_progress(60, "Đang lưu trang thống kê...")
            pdf.savefig()
            plt.close()
            
            # Trang 2: Danh sách điểm
            update_progress(70, "Đang tạo danh sách điểm học sinh...")
            plt.figure(figsize=(10, 12))
            plt.axis('off')
            
            plt.text(0.5, 0.98, "DANH SÁCH ĐIỂM HỌC SINH", fontsize=16, ha='center', fontweight='bold')
            
            # Vẽ bảng điểm
            ax = plt.axes([0.05, 0.05, 0.9, 0.85])
            ax.axis('tight')
            ax.axis('off')
            
            # Số học sinh tối đa hiển thị trên một trang
            page_size = 40
            name_col = config['columns']['name']
            exam_code_col = 'Mã đề' if 'Mã đề' in scores_df.columns else None
            
            # Nếu có mã đề thì hiển thị
            if exam_code_col:
                sorted_df = scores_df.sort_values(name_col)
                table_data = [[student, exam_code, score] for student, exam_code, score in 
                             zip(sorted_df[name_col][:page_size], 
                                 sorted_df[exam_code_col][:page_size], 
                                 sorted_df['Điểm'][:page_size])]
                col_labels = [name_col, exam_code_col, 'Điểm']
            else:
                sorted_df = scores_df.sort_values(name_col)
                table_data = [[student, score] for student, score in 
                             zip(sorted_df[name_col][:page_size], 
                                 sorted_df['Điểm'][:page_size])]
                col_labels = [name_col, 'Điểm']
            
            table = ax.table(cellText=table_data, colLabels=col_labels,
                          loc='center', cellLoc='center')
            table.auto_set_font_size(False)
            table.set_fontsize(9)
            table.scale(1, 1.2)
            
            update_progress(80, "Đang lưu danh sách điểm...")
            pdf.savefig()
            plt.close()
            
            # Tạo thêm trang nếu số học sinh > page_size
            if len(scores_df) > page_size:
                pages_needed = (len(scores_df) - 1) // page_size + 1
                progress_per_page = 19 / (pages_needed - 1) if pages_needed > 1 else 0
                
                for page in range(1, pages_needed):
                    update_progress(80 + page * progress_per_page, f"Đang tạo trang danh sách {page+1}/{pages_needed}...")
                    plt.figure(figsize=(10, 12))
                    plt.axis('off')
                    
                    plt.text(0.5, 0.98, f"DANH SÁCH ĐIỂM HỌC SINH (trang {page+1}/{pages_needed})", 
                          fontsize=16, ha='center', fontweight='bold')
                    
                    ax = plt.axes([0.05, 0.05, 0.9, 0.85])
                    ax.axis('tight')
                    ax.axis('off')
                    
                    start_idx = page * page_size
                    end_idx = min((page + 1) * page_size, len(scores_df))
                    
                    if exam_code_col:
                        table_data = [[student, exam_code, score] for student, exam_code, score in 
                                     zip(sorted_df[name_col][start_idx:end_idx], 
                                         sorted_df[exam_code_col][start_idx:end_idx], 
                                         sorted_df['Điểm'][start_idx:end_idx])]
                    else:
                        table_data = [[student, score] for student, score in 
                                     zip(sorted_df[name_col][start_idx:end_idx], 
                                         sorted_df['Điểm'][start_idx:end_idx])]
                    
                    table = ax.table(cellText=table_data, colLabels=col_labels,
                                  loc='center', cellLoc='center')
                    table.auto_set_font_size(False)
                    table.set_fontsize(9)
                    table.scale(1, 1.2)
                    
                    pdf.savefig()
                    plt.close()
            
            update_progress(100, "Hoàn tất tạo báo cáo!")
        
        # Đóng cửa sổ tiến trình sau 1 giây
        progress_window.after(1000, progress_window.destroy)
        
        status_label.config(text=f"Đã tạo báo cáo PDF: {os.path.basename(file_path)}", 
)
        ToastNotification.show(f"✅ Đã tạo báo cáo PDF tại:\n{file_path}", "success")
        
    except Exception as e:
        status_label.config(text=f"Lỗi tạo báo cáo: {str(e)}", 
)
        messagebox.showerror("Lỗi", f"Không thể tạo báo cáo PDF: {str(e)}")
        traceback.print_exc()

def show_score_distribution():
    """Hiển thị biểu đồ phân phối điểm số"""
    global df
    if df is None or df.empty:
        ToastNotification.show("ℹ️ Không có dữ liệu để hiển thị thống kê", "info")
        return
        
    if 'Điểm' not in df.columns:
        ToastNotification.show("ℹ️ Không có cột điểm trong dữ liệu", "info")
        return
    
    # Lọc chỉ lấy học sinh có điểm
    scores = df[df['Điểm'].notna()]['Điểm']
    
    if len(scores) == 0:
        ToastNotification.show("ℹ️ Chưa có học sinh nào có điểm", "info")
        return
    
    try:
        # Tạo cửa sổ mới với style hiện đại
        chart_window = tk.Toplevel(root)
        chart_window.title("📊 Thống Kê Điểm Số Lớp")
        chart_window.geometry("1100x650")
        chart_window.transient(root)
        chart_window.configure(bg=config['ui']['theme']['background'])
        
        # Tạo frame chứa biểu đồ
        chart_frame = ttk.Frame(chart_window)
        chart_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # Tạo figure với dark background
        fig = Figure(figsize=(11, 6), dpi=100, facecolor=config['ui']['theme']['background'])
        
        # Tính toán thống kê
        mean_score = scores.mean()
        median_score = scores.median()
        pass_threshold = 5.0
        passed = (scores >= pass_threshold).sum()
        failed = len(scores) - passed
        
        # Phân loại học lực
        excellent = (scores >= 8.5).sum()  # Giỏi
        good = ((scores >= 7) & (scores < 8.5)).sum()  # Khá
        average = ((scores >= 5) & (scores < 7)).sum()  # TB
        weak = (scores < 5).sum()  # Yếu
        
        # Layout: 1 hàng 3 cột - đơn giản, dễ nhìn
        gs = fig.add_gridspec(1, 3, wspace=0.3)
        
        # 1. BIỂU ĐỒ CỘT - Phân loại học lực (CHÍNH)
        ax1 = fig.add_subplot(gs[0, 0:2])  # Chiếm 2/3 không gian
        
        categories = ['Giỏi\n(≥8.5)', 'Khá\n(7-8.5)', 'TB\n(5-7)', 'Yếu\n(<5)']
        values = [excellent, good, average, weak]
        colors = ['#667EEA', '#68D391', '#F6AD55', '#FC8181']
        
        bars = ax1.bar(categories, values, color=colors, alpha=0.8, edgecolor='white', linewidth=2)
        
        # Thêm số liệu lên đầu mỗi cột
        for bar, value in zip(bars, values):
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height,
                    f'{int(value)} HS\n({value/len(scores)*100:.1f}%)',
                    ha='center', va='bottom', color='white', fontsize=11, fontweight='bold')
        
        ax1.set_title('PHÂN LOẠI HỌC LỰC', fontsize=15, fontweight='bold', color='white', pad=15)
        ax1.set_ylabel('Số Học Sinh', fontsize=12, color='white')
        ax1.set_facecolor(config['ui']['theme']['card'])
        ax1.grid(True, alpha=0.2, axis='y', color='white', linestyle='--')
        ax1.tick_params(colors='white', labelsize=10)
        ax1.spines['top'].set_visible(False)
        ax1.spines['right'].set_visible(False)
        ax1.spines['left'].set_color('white')
        ax1.spines['bottom'].set_color('white')
        
        # 2. THỐNG KÊ TỔNG QUAN - Bảng số liệu
        ax2 = fig.add_subplot(gs[0, 2])
        ax2.axis('off')
        ax2.set_facecolor(config['ui']['theme']['card'])
        
        # Tạo bảng thông tin (không dùng emoji để tránh lỗi font)
        stats_data = [
            ['TỔNG SỐ', f'{len(scores)} HS'],
            ['', ''],
            ['Đạt (≥5)', f'{passed} HS'],
            ['Chưa đạt', f'{failed} HS'],
            ['Tỷ lệ đạt', f'{passed/len(scores)*100:.1f}%'],
            ['', ''],
            ['Điểm TB', f'{mean_score:.2f}'],
            ['Trung bình', f'{median_score:.2f}'],
            ['Cao nhất', f'{scores.max():.2f}'],
            ['Thấp nhất', f'{scores.min():.2f}'],
        ]
        
        table = ax2.table(cellText=stats_data, cellLoc='left',
                         bbox=[0, 0, 1, 1],
                         colWidths=[0.5, 0.5])
        
        table.auto_set_font_size(False)
        table.set_fontsize(11)
        
        # Style cho bảng
        for i, row in enumerate(stats_data):
            for j in range(2):
                cell = table[(i, j)]
                cell.set_facecolor(config['ui']['theme']['card'])
                cell.set_text_props(color='white', weight='bold' if j == 0 else 'normal')
                cell.set_edgecolor('white' if row[0] else config['ui']['theme']['card'])
                cell.set_linewidth(0.5 if row[0] else 0)
        
        ax2.set_title('THỐNG KÊ', fontsize=13, fontweight='bold', color='white', pad=10)
        
        fig.tight_layout(pad=2.0)
        
        # Tạo canvas để hiển thị biểu đồ
        canvas = FigureCanvasTkAgg(fig, chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        
        # Thêm nút để lưu biểu đồ
        button_frame = ttk.Frame(chart_window)
        button_frame.pack(fill="x", padx=10, pady=10)
        
        def save_chart():
            save_path = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[
                    ("PNG files", "*.png"),
                    ("PDF files", "*.pdf"),
                    ("SVG files", "*.svg"),
                    ("All files", "*.*")
                ],
                title="Lưu biểu đồ"
            )
            if save_path:
                fig.savefig(save_path, dpi=300, bbox_inches='tight', facecolor=config['ui']['theme']['background'])
                ToastNotification.show(f"✅ Đã lưu biểu đồ tại:\n{save_path}", "success")
        
        ttk.Button(button_frame, text="Lưu biểu đồ", command=save_chart).pack(side="left")
        ttk.Button(button_frame, text="Làm mới", command=lambda: (chart_window.destroy(), show_score_distribution())).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Đóng", command=chart_window.destroy).pack(side="right")
        
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể hiển thị biểu đồ: {str(e)}")
        traceback.print_exc()

def focus_correct_count(event=None):
    """Di chuyển con trỏ đến ô nhập số câu đúng"""
    entry_correct_count.focus_set()
    entry_correct_count.select_range(0, tk.END)

def find_header_row(headers_df):
    """Tìm hàng chứa header thực sự trong DataFrame."""
    header_row = None
    for i in range(len(headers_df)):
        row_values = headers_df.iloc[i].astype(str)
        if any(name.lower() in ' '.join(row_values.str.lower()) 
              for name in ['họ và tên', 'tên học sinh', 'học sinh']):
            header_row = i
            break
    
    # Nếu không tìm thấy header phù hợp, dùng hàng đầu tiên
    if header_row is None:
        header_row = 0
    
    return header_row

def read_excel_file(file_path):
    """Đọc file Excel theo cách thông thường"""
    status_label.config(text=f"Đang đọc file Excel...")
    root.update()
    
    try:
        # Thử đọc với engine openpyxl trước
        headers_df = pd.read_excel(file_path, nrows=10, engine='openpyxl')
        
        # Tìm hàng chứa header thực sự
        header_row = find_header_row(headers_df)
        
        # Đọc lại với header đúng
        df_result = pd.read_excel(file_path, header=header_row, engine='openpyxl')
        
        # Xử lý trường hợp DataFrame rỗng
        if df_result.empty:
            status_label.config(text="File Excel không có dữ liệu")
            return pd.DataFrame()  # Trả về DataFrame rỗng thay vì None
        
        # Đảm bảo các cột cần thiết tồn tại
        df_result = ensure_required_columns(df_result)
        
        # Đảm bảo kiểu dữ liệu phù hợp
        df_result = ensure_proper_dtypes(df_result)
        
        status_label.config(text=f"Đã đọc xong file Excel")
        return df_result
        
    except ImportError:
        # Nếu không có openpyxl, thử dùng engine mặc định
        print("Không tìm thấy openpyxl, sử dụng engine mặc định...")
        headers_df = pd.read_excel(file_path, nrows=10)
        
        # Tìm hàng chứa header thực sự
        header_row = find_header_row(headers_df)
        
        # Đọc lại với header đúng
        df_result = pd.read_excel(file_path, header=header_row)
        
        # Xử lý trường hợp DataFrame rỗng
        if df_result.empty:
            status_label.config(text="File Excel không có dữ liệu")
            return pd.DataFrame()  # Trả về DataFrame rỗng thay vì None
        
        # Đảm bảo các cột cần thiết tồn tại
        df_result = ensure_required_columns(df_result)
        
        # Đảm bảo kiểu dữ liệu phù hợp
        df_result = ensure_proper_dtypes(df_result)
        
        status_label.config(text=f"Đã đọc xong file Excel")
        return df_result
        
    except Exception as e:
        status_label.config(text=f"Lỗi: {str(e)}")
        messagebox.showerror("Lỗi", f"Không thể đọc file Excel: {str(e)}")
        traceback.print_exc()
        return pd.DataFrame()  # Trả về DataFrame rỗng khi có lỗi 
    
    
def auto_update_stats():
    """Tự động cập nhật thống kê theo chu kỳ"""
    # Cập nhật thống kê
    if df is not None and not df.empty:
        update_stats()
        update_score_extremes()
    
    # Thiết lập lịch cập nhật tiếp theo (5 giây)
    root.after(5000, auto_update_stats)

def choose_update_channel():
    """
    Hiển thị hộp thoại để người dùng chọn kênh cập nhật
    """
    try:
        from check_for_updates import show_update_channel_dialog
        
        def on_channel_changed(new_channel):
            # Cập nhật UI sau khi thay đổi kênh
            status_label.config(text=f"Đã chuyển sang kênh cập nhật: {new_channel}")
            
        show_update_channel_dialog(root, config, save_config, callback=on_channel_changed)
    except ImportError as e:
        print(f"Lỗi khi nhập module show_update_channel_dialog: {str(e)}")
        messagebox.showerror("Lỗi", "Không thể mở dialog chọn kênh cập nhật.")

# Khởi tạo giao diện
create_ui()

# Tạo biến style
style = ttk.Style()

# Tạo menu
menubar = tk.Menu(root)
root.config(menu=menubar)

# Menu Cài đặt
settings_menu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="Cài đặt", menu=settings_menu)
settings_menu.add_command(label="Tùy chỉnh phím tắt", command=customize_shortcuts)
settings_menu.add_command(label="Tùy chỉnh mã đề", command=customize_exam_codes)
settings_menu.add_command(label="Tùy chỉnh tên cột", command=customize_columns)
settings_menu.add_command(label="Bảo mật", command=customize_security)
settings_menu.add_separator()
settings_menu.add_command(label="Chọn kênh cập nhật", command=choose_update_channel)
settings_menu.add_command(label="Chế độ tối/sáng", command=lambda: toggle_theme(style))

# Menu Chức năng
function_menu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="Chức năng", menu=function_menu)
function_menu.add_command(label="Biểu đồ phân phối điểm", command=show_score_distribution)
function_menu.add_command(label="Xuất báo cáo PDF", command=generate_report)
function_menu.add_separator()
function_menu.add_command(label="Kiểm tra cập nhật", command=lambda: check_for_updates_wrapper(True))
function_menu.add_command(label="Thông tin", command=show_about)

# Thêm binding phím tắt
root.bind('<Control-z>', lambda e: perform_undo())  # Undo
root.bind('<Control-y>', lambda e: perform_redo())  # Redo
root.bind('<Control-f>', focus_student_search)
root.bind('<Control-g>', focus_direct_score)
root.bind('<Control-d>', focus_correct_count)
root.bind('<Control-s>', lambda e: save_excel() if df is not None else None)
root.bind('<Control-o>', lambda e: select_file())
root.bind('<Control-n>', lambda e: entry_student_name.focus_set())

# Cập nhật menu recent files sau khi tạo UI
update_recent_files_menu()

# Cập nhật thời gian hoạt động cuối cùng khi có sự kiện
root.bind_all("<Motion>", update_activity_time)
root.bind_all("<KeyPress>", update_activity_time)

# Kiểm tra tự động khóa ứng dụng
check_auto_lock()

# Kiểm tra cập nhật sau khi khởi động (tăng thời gian delay để đảm bảo giao diện đã được tạo hoàn chỉnh)
print("Lên lịch kiểm tra cập nhật tự động sau 5 giây...")
root.after(5000, check_updates_async)

# Hiện thông báo khi đã sẵn sàng
status_label.config(text="✅ Ứng dụng đã sẵn sàng! Đang kiểm tra cập nhật...")

# Bắt đầu cập nhật thống kê tự động
root.after(1000, auto_update_stats)

# Chạy ứng dụng
root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()
    
