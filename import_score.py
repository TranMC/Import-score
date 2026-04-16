import tkinter as tk
import customtkinter as ctk
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
import re
from collections import OrderedDict
from openpyxl import load_workbook

# Import các module mới
import version_utils
import themes
import ui_utils

# ========================================
# CACHING LAYER
# ========================================

class ConfigCache:
    """Cache cho config file để tránh đọc file liên tục"""
    _cache = None
    _last_modified = None
    
    @classmethod
    def get_config(cls, config_path):
        """Lấy config từ cache hoặc đọc file nếu có thay đổi"""
        try:
            if not os.path.exists(config_path):
                return None
                
            current_mtime = os.path.getmtime(config_path)
            
            if cls._cache is None or cls._last_modified != current_mtime:
                with open(config_path, 'r', encoding='utf-8') as f:
                    cls._cache = json.load(f)
                cls._last_modified = current_mtime
            
            return cls._cache
        except Exception as e:
            print(f"Error getting config from cache: {e}")
            return None
    
    @classmethod
    def invalidate(cls):
        """Xóa cache"""
        cls._cache = None
        cls._last_modified = None


class DataFrameCache:
    """Cache cho DataFrame từ Excel files"""
    def __init__(self, max_size=3):
        self.cache = OrderedDict()
        self.max_size = max_size
    
    def _get_file_key(self, file_path):
        """Tạo key duy nhất cho file dựa trên path và modification time"""
        try:
            if not os.path.exists(file_path):
                return None
            mtime = os.path.getmtime(file_path)
            return f"{file_path}_{mtime}"
        except:
            return None
    
    def get(self, file_path):
        """Lấy DataFrame từ cache"""
        key = self._get_file_key(file_path)
        if key and key in self.cache:
            # Di chuyển key lên đầu (most recently used)
            self.cache.move_to_end(key)
            return self.cache[key].copy()
        return None
    
    def set(self, file_path, df):
        """Lưu DataFrame vào cache"""
        key = self._get_file_key(file_path)
        if key is None:
            return
        
        # Xóa file cũ nếu vượt quá max_size
        if len(self.cache) >= self.max_size:
            self.cache.popitem(last=False)  # Xóa oldest
        
        self.cache[key] = df.copy()
    
    def clear(self):
        """Xóa toàn bộ cache"""
        self.cache.clear()


class SearchCache:
    """Cache cho kết quả search"""
    def __init__(self):
        self.cache = {}
        self.df_hash = None
    
    def _get_df_hash(self, df):
        """Tạo hash cho DataFrame để detect thay đổi"""
        try:
            return hash(f"{len(df)}_{df.shape[1]}_{df.iloc[0].sum() if len(df) > 0 else 0}")
        except:
            return None
    
    def clear_if_data_changed(self, df):
        """Xóa cache nếu DataFrame thay đổi"""
        current_hash = self._get_df_hash(df)
        if current_hash != self.df_hash:
            self.cache.clear()
            self.df_hash = current_hash
    
    def get(self, query):
        """Lấy kết quả search từ cache"""
        return self.cache.get(query.lower())
    
    def set(self, query, results):
        """Lưu kết quả search vào cache"""
        self.cache[query.lower()] = results
    
    def clear(self):
        """Xóa toàn bộ cache"""
        self.cache.clear()
        self.df_hash = None


class StatsCache:
    """Cache cho statistics calculations"""
    def __init__(self):
        self.stats = None
        self.df_version = None
    
    def _get_df_version(self, df):
        """Tạo version identifier cho DataFrame"""
        try:
            if df is None or len(df) == 0:
                return "empty"
            return f"{len(df)}_{df.shape[1]}_{df.iloc[0].sum() if len(df) > 0 else 0}"
        except:
            return None
    
    def get_stats(self, df, force_recalc=False):
        """Lấy stats từ cache hoặc tính lại nếu cần"""
        current_version = self._get_df_version(df)
        
        if force_recalc or self.stats is None or self.df_version != current_version:
            self.stats = self._calculate_stats(df)
            self.df_version = current_version
        
        return self.stats
    
    def _calculate_stats(self, df):
        """Tính toán statistics"""
        try:
            if df is None or len(df) == 0:
                return {
                    'total': 0,
                    'scored': 0,
                    'mean': 0,
                    'max': 0,
                    'min': 0
                }
            
            score_col = None
            for col in ['Điểm', 'score', 'Score']:
                if col in df.columns:
                    score_col = col
                    break
            
            if score_col is None:
                return {'total': len(df), 'scored': 0, 'mean': 0, 'max': 0, 'min': 0}
            
            scores = df[score_col].dropna()
            return {
                'total': len(df),
                'scored': len(scores),
                'mean': float(scores.mean()) if len(scores) > 0 else 0,
                'max': float(scores.max()) if len(scores) > 0 else 0,
                'min': float(scores.min()) if len(scores) > 0 else 0
            }
        except Exception as e:
            print(f"Error calculating stats: {e}")
            return {'total': 0, 'scored': 0, 'mean': 0, 'max': 0, 'min': 0}
    
    def clear(self):
        """Xóa cache"""
        self.stats = None
        self.df_version = None


# Khởi tạo global cache instances
_df_cache = DataFrameCache(max_size=3)
_search_cache = SearchCache()
_stats_cache = StatsCache()

# ========================================
# END CACHING LAYER
# ========================================

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
        self.tooltip = ctk.CTkToplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        
        label = ctk.CTkLabel(self.tooltip, text=self.text, 
                         fg_color="#ffffe0", text_color="black",
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
    """Hệ thống thông báo toast hiện đại bằng CustomTkinter Frame"""
    active_toasts = []
    
    @staticmethod
    def show(message, toast_type="info", duration=3000):
        # Màu sắc theo loại
        colors = {
            'success': {'bg': '#10B981', 'fg': 'white', 'icon': '✓'},
            'error': {'bg': '#EF4444', 'fg': 'white', 'icon': '✕'},
            'warning': {'bg': '#F59E0B', 'fg': 'white', 'icon': '⚠'},
            'info': {'bg': '#3B82F6', 'fg': 'white', 'icon': 'ℹ'}
        }
        
        style_config = colors.get(toast_type, colors['info'])
        
        # Tạo toast sử dụng frame đặt trực tiếp lên root
        toast_frame = ctk.CTkFrame(root, fg_color=style_config['bg'], corner_radius=8)
        
        # Icon
        ctk.CTkLabel(toast_frame, 
                text=style_config['icon'], 
                font=('Segoe UI', 16),
                text_color=style_config['fg']).pack(side='left', padx=(12, 8), pady=10)
        
        # Message
        ctk.CTkLabel(toast_frame, 
                text=message,
                font=('Segoe UI', 12, 'bold'),
                text_color=style_config['fg'],
                wraplength=250).pack(side='left', fill='both', expand=True, padx=(0, 15), pady=10)
        
        # Lọc các toast còn tồn tại
        ToastNotification.active_toasts = [t for t in ToastNotification.active_toasts if t.winfo_exists()]
        
        # Tính toán Y offset
        base_y = 20
        offset_y = len(ToastNotification.active_toasts) * 60
        y_pos = base_y + offset_y
        
        # Đặt frame ở góc dưới bên phải
        toast_frame.place(relx=1.0, rely=1.0, anchor="se", x=-20, y=-y_pos)
        
        # Thêm vào danh sách active
        ToastNotification.active_toasts.append(toast_frame)
        
        def close_toast(event=None):
            if toast_frame in ToastNotification.active_toasts:
                ToastNotification.active_toasts.remove(toast_frame)
            if toast_frame.winfo_exists():
                toast_frame.destroy()
            
            # Reposition remaining toasts
            for idx, t in enumerate([x for x in ToastNotification.active_toasts if x.winfo_exists()]):
                t.place_configure(y=-(20 + idx * 60))
                
        # Tự động đóng sau duration nếu > 0
        if duration > 0:
            root.after(duration, close_toast)
            
        # Click để đóng sớm
        toast_frame.bind('<Button-1>', close_toast)
        for child in toast_frame.winfo_children():
            child.bind('<Button-1>', close_toast)
            
        return toast_frame

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
config['ui']['dark_mode'] = True
# Áp dụng theme tối trực tiếp vào config
config = themes.apply_theme_to_config(config, True)

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
ctk.set_appearance_mode("Dark")
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
current_filter_type = 'all'  # Lọc mặc định là tất cả


def ensure_proper_dtypes(df_input):
    """
    Đảm bảo các kiểu dữ liệu đúng cho các cột quan trọng, đặc biệt là xử lý cột Categorical
    để tránh lỗi khi gán giá trị mới
    """
    if df_input is None:
        return pd.DataFrame()  # Trả về DataFrame rỗng thay vì None
        
    df_copy = df_input.copy()
    
    # Xử lý tất cả các cột điểm số (tìm theo pattern)
    score_patterns = ['điểm', 'đđg', 'đtb', 'score', 'point']
    for col in df_copy.columns:
        col_lower = str(col).lower()
        if any(pattern in col_lower for pattern in score_patterns):
            # Chuyển đổi sang dạng số
            try:
                df_copy[col] = pd.to_numeric(df_copy[col], errors='coerce')
            except Exception as e:
                print(f"Lỗi khi chuyển đổi cột {col}: {str(e)}")
    
    # Xử lý cột Điểm (legacy support)
    if 'Điểm' in df_copy.columns:
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
        header_row (int): Dòng chứa tiêu đề, None nếu cần tự động tìm
        
    Returns:
        pd.DataFrame: DataFrame hoàn chỉnh sau khi đọc
    """
    from pandas.io.excel._openpyxl import OpenpyxlReader
    global _df_cache
    
    # Trạng thái tiến trình
    status_label.configure(text=f"Đang lập kế hoạch đọc file lớn: {os.path.basename(file_path)}...", 
)
    root.update()
    
    try:
        # Check cache
        config = load_config()
        enable_caching = config.get('excel_reading', {}).get('performance', {}).get('enable_caching', True)
        
        if enable_caching:
            cached_df = _df_cache.get(file_path)
            if cached_df is not None:
                status_label.configure(text=f"Đã load file từ cache")
                return cached_df
        
        # Đọc thông tin file để tìm hiểu số dòng
        excel_reader = pd.ExcelFile(file_path, engine='openpyxl')
        sheet_name = excel_reader.sheet_names[0]  # Lấy sheet đầu tiên
        
        # Chỉ đọc header để phân tích
        if header_row is None:
            # Đọc nhiều dòng hơn để phân tích header
            max_rows = config.get('excel_reading', {}).get('header_detection', {}).get('max_search_rows', 50)
            header_sample = pd.read_excel(excel_reader, sheet_name=sheet_name, nrows=max_rows)
            
            # Tìm hàng chứa header với scoring
            header_row = find_header_row(header_sample, config)
        
        # Cập nhật trạng thái
        status_label.configure(text=f"Đang đọc dữ liệu theo chunk từ dòng {header_row+1}...", 
)
        root.update()
        
        # Check multi-level headers
        multi_level = config.get('excel_reading', {}).get('header_detection', {}).get('multi_level_support', True)
        header_rows_list = None
        
        if multi_level and header_row < max_rows - 1:
            # Có thể có multi-level, đọc thêm vài dòng
            header_rows_to_read = min(3, max_rows - header_row)
            if header_rows_to_read > 1:
                header_rows_list = list(range(header_row, header_row + header_rows_to_read))
        
        # Đọc dữ liệu theo từng chunk
        chunks = []
        if header_rows_list:
            reader = pd.read_excel(
                excel_reader, 
                sheet_name=sheet_name,
                header=header_rows_list,
                chunksize=chunk_size
            )
        else:
            reader = pd.read_excel(
                excel_reader, 
                sheet_name=sheet_name,
                header=header_row,
                chunksize=chunk_size
            )
        
        total_rows = 0
        
        for i, chunk in enumerate(reader):
            # Flatten MultiIndex columns nếu có
            if isinstance(chunk.columns, pd.MultiIndex):
                separator = config.get('excel_reading', {}).get('header_detection', {}).get('merge_separator', '_')
                chunk.columns = [separator.join(str(i) for i in col if str(i) != 'nan').strip(separator) 
                                for col in chunk.columns.values]
            
            chunks.append(chunk)
            
            total_rows += len(chunk)
            status_label.configure(text=f"Đang đọc dữ liệu: {total_rows} dòng...", 
)
            root.update()
            
            # Nếu đã đọc quá nhiều, hiển thị cảnh báo
            if total_rows > 10000:
                status_label.configure(text=f"File lớn: đã đọc {total_rows} dòng...", 
)            # Ghép các chunk lại
        
        if chunks:
            status_label.configure(text=f"Đang ghép {len(chunks)} chunk dữ liệu...", 
)
            root.update()
        
            result = pd.concat(chunks, ignore_index=True)
        
            # Đảm bảo kiểu dữ liệu phù hợp cho các cột quan trọng
            result = ensure_proper_dtypes(result)
            
            # Cache result
            if enable_caching:
                _df_cache.set(file_path, result)
        
            status_label.configure(text=f"Đã đọc xong {total_rows} dòng dữ liệu", 
)
            return result
        else:
            return pd.DataFrame()
            
    except Exception as e:
        status_label.configure(text=f"Lỗi khi đọc file: {str(e)}", 
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
    status_label.configure(text=f"Đang đọc file: {os.path.basename(file_path)}...")
    root.update()
    
    try:
        df = read_excel_file(file_path)
        
        if df is not None and not df.empty:
            total_students = len(df)
            status_label.configure(text=f"Đã đọc xong: {os.path.basename(file_path)} ({total_students} học sinh)")
            
            df = ensure_required_columns(df)
            refresh_ui()
            
            # Thêm vào recent files
            add_to_recent_files(file_path)
            
            ToastNotification.show(f"Đã mở file {os.path.basename(file_path)}", "success")
        else:
            status_label.configure(text="Không có dữ liệu để hiển thị")
    except Exception as e:
        error_message = str(e)
        status_label.configure(text=f"Lỗi: {error_message[:50]}")
        ToastNotification.show(f"Lỗi khi đọc file: {error_message[:100]}", "error")
        traceback.print_exc()

file_menu_widget = None

def update_recent_files_menu():
    """Cập nhật menu Recent Files một cách ổn định, tránh lặp lại"""
    global file_menu_widget
    try:
        # Lấy menubar chính
        menu_ptr = root.cget('menu')
        if not menu_ptr:
            return
        menubar = root.nametowidget(menu_ptr)
        
        # 1. Kiểm tra cache toàn cục
        if file_menu_widget is not None:
            try:
                if not file_menu_widget.winfo_exists():
                    file_menu_widget = None
            except:
                file_menu_widget = None
        
        # 2. Nếu chưa có trong cache, tìm trong menubar
        file_menu = file_menu_widget
        if file_menu is None:
            try:
                last_idx = menubar.index('end')
                if last_idx is not None:
                    for i in range(last_idx + 1):
                        try:
                            label = str(menubar.entrycget(i, 'label'))
                            # Tìm label có chứa "File" hoặc "Tệp" (không quan tâm emoji)
                            if "File" in label or "Tệp" in label:
                                menu_name = menubar.entrycget(i, 'menu')
                                if menu_name:
                                    file_menu = menubar.nametowidget(menu_name)
                                    file_menu_widget = file_menu
                                    break
                        except:
                            continue
            except:
                pass
        
        # 3. Nếu vẫn không thấy, tạo mới ở vị trí đầu tiên
        if file_menu is None:
            file_menu = tk.Menu(menubar, tearoff=0)
            menubar.insert_cascade(0, label="📁 File", menu=file_menu)
            file_menu_widget = file_menu
        
        # XOÁ TOÀN BỘ để xây dựng lại từ đầu, tránh bị chồng chất separator/items
        file_menu.delete(0, 'end')
        
        # 1. Các lệnh cơ bản
        file_menu.add_command(label="📂 Mở file Excel...", accelerator="Ctrl+O", 
                             command=select_file)
        file_menu.add_command(label="💾 Lưu file (Excel)", accelerator="Ctrl+S", 
                             command=lambda: save_excel() if df is not None else None)
        file_menu.add_command(label="📑 Xuất báo cáo (PDF)", 
                             command=generate_report)
        
        # 2. Danh sách file gần đây
        recent_files = config.get('recent_files', [])
        if recent_files:
            file_menu.add_separator()
            file_menu.add_command(label="📌 File Gần Đây", state='disabled')
            
            # Chỉ lấy các file thực sự tồn tại
            valid_recent = []
            for fp in recent_files:
                if os.path.exists(fp):
                    valid_recent.append(fp)
            
            # Cập nhật lại config nếu có file không tồn tại
            if len(valid_recent) != len(recent_files):
                config['recent_files'] = valid_recent
                # Tránh gọi save_config() ở đây để không tạo vòng lặp, 
                # chỉ cập nhật biến cục bộ hoặc lưu im lặng
                
            for i, filepath in enumerate(valid_recent[:5]):
                filename = os.path.basename(filepath)
                file_menu.add_command(
                    label=f"  {i+1}. {filename}",
                    command=lambda fp=filepath: open_recent_file(fp)
                )
        
        # Thêm phím thoát
        file_menu.add_separator()
        file_menu.add_command(label="❌ Thoát", command=on_closing)
        
    except Exception as e:
        print(f"Lỗi khi cập nhật recent files menu: {str(e)}")
        traceback.print_exc()

def select_file():
    global df, file_path
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        # Hiển thị thông báo trạng thái khi bắt đầu đọc file
        status_label.configure(text=f"Đang đọc file: {os.path.basename(file_path)}...", 
)
        root.update()  # Cập nhật giao diện ngay lập tức để hiển thị trạng thái
        
        try:
            # Đọc file Excel bình thường
            df = read_excel_file(file_path)
            
            if df is not None and not df.empty:
                # Hiển thị số lượng học sinh
                total_students = len(df)
                status_label.configure(
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
                status_label.configure(
                    text="Không có dữ liệu để hiển thị, vui lòng tải file Excel có dữ liệu",

                )
                
        except Exception as e:
            error_message = str(e)
            status_label.configure(
                text=f"Lỗi: {error_message[:50] + '...' if len(error_message) > 50 else error_message}",

            )
            ToastNotification.show(f"Lỗi: {error_message[:100]}", "error")
            traceback.print_exc()  # In chi tiết lỗi ra console để debug

def read_excel_normally(file_path):
    """Đọc file Excel theo cách thông thường"""
    status_label.configure(text=f"Đang đọc file Excel...")
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
    """Find column that best matches the target name using config patterns"""
    target_name_lower = target_name.lower().strip()
    
    # Direct match first (case insensitive)
    for col in df.columns:
        if str(col).lower().strip() == target_name_lower:
            return col
    
    # Try using new pattern matching system
    try:
        config = load_config()
        patterns_config = config.get('excel_reading', {}).get('column_patterns', {})
        
        # Check if target_name matches any known field
        for col in df.columns:
            col_str = str(col).strip()
            
            # Try matching with student_info patterns
            for field_name, field_config in patterns_config.get('student_info', {}).items():
                patterns = field_config.get('patterns', [])
                for pattern in patterns:
                    if pattern.lower() == target_name_lower:
                        # This is the field we're looking for
                        # Check if col matches this field
                        for col_pattern in patterns:
                            if col_pattern.lower() in col_str.lower() or col_str.lower() in col_pattern.lower():
                                return col
                        # Try regex
                        regex = field_config.get('regex', '')
                        if regex:
                            try:
                                if re.match(regex, col_str, re.IGNORECASE):
                                    return col
                            except:
                                pass
            
            # Try matching with score_columns patterns
            for score_type, score_config in patterns_config.get('score_columns', {}).items():
                patterns = score_config.get('patterns', [])
                for pattern in patterns:
                    if pattern.lower() == target_name_lower:
                        # This is the field we're looking for
                        # Check if col matches this field
                        for col_pattern in patterns:
                            if col_pattern.lower() in col_str.lower() or col_str.lower() in col_pattern.lower():
                                return col
                        # Try regex
                        regex = score_config.get('regex', '')
                        if regex:
                            try:
                                if re.match(regex, col_str, re.IGNORECASE):
                                    return col
                            except:
                                pass
    except Exception as e:
        print(f"Error using pattern matching: {e}")
    
    # Fallback to old variations
    name_variations = {
        'tên học sinh': ['họ và tên', 'họ tên', 'tên', 'học sinh'],
        'họ và tên': ['họ và tên', 'họ tên', 'tên', 'học sinh', 'tên học sinh'],
        'mã đề': ['mã', 'đề', 'số đề', 'mã số đề', 'exam code'],
        'điểm': ['điểm số', 'số điểm', 'point', 'score'],
        'đđgck': ['đđgck', 'điểm ck', 'điểm cuối kỳ', 'ck', 'cuối kỳ'],
        'đđggk': ['đđggk', 'điểm gk', 'điểm giữa kỳ', 'gk', 'giữa kỳ'],
        'đđgtx': ['đđgtx', 'điểm tx', 'điểm thường xuyên', 'tx', 'thường xuyên'],
        'đtbmhki': ['đtbmhki', 'đtb hk1', 'điểm tb', 'điểm trung bình']
    }
    
    # Check variations
    for col in df.columns:
        col_lower = str(col).lower().strip()
        for key, variations in name_variations.items():
            if target_name_lower == key or target_name_lower in variations:
                if any(var in col_lower for var in variations):
                    return col
                # Partial match
                if any(col_lower in var or var in col_lower for var in variations):
                    return col
    
    return None

def search_student(event=None):
    """Tìm kiếm học sinh với search caching"""
    global df, search_timer_id, search_index, _search_cache
    
    # Reset timer
    search_timer_id = None
    
    # Clear search cache if data changed
    if df is not None:
        _search_cache.clear_if_data_changed(df)
    
    # Hiển thị thông báo xử lý nếu dữ liệu lớn
    if df is not None and len(df) > 1000:
        status_label.configure(text="Đang tìm kiếm trong dữ liệu lớn...")
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
    
    # 1. Lấy dữ liệu cơ sở và áp dụng Quick Filter nếu có
    result = df.copy()
    if 'current_filter_type' in globals() and current_filter_type != 'all':
        score_col = column_mapping.get('score')
        if not score_col:
            score_col = find_matching_column(result, 'Điểm')
            
        if score_col:
            temp_scores = pd.to_numeric(result[score_col], errors='coerce')
            if current_filter_type == 'high':
                result = result[temp_scores >= 7]
            elif current_filter_type == 'medium':
                result = result[(temp_scores >= 5) & (temp_scores < 7)]
            elif current_filter_type == 'low':
                result = result[temp_scores < 5]
            elif current_filter_type == 'no_score':
                result = result[temp_scores.isna()]
                
    # 2. Áp dụng tìm kiếm (với kết quả đã lọc)
    if ten_hoc_sinh:
        mask = result[name_col].astype(str).str.lower().str.contains(ten_hoc_sinh, na=False)
        result = result[mask]

    # Get display values helper (giữ nguyên)
    def get_display_values(row):
        values = []
        for col_key in ['name', 'exam_code', 'score']:
            if col_key in column_mapping:
                value = row[column_mapping[col_key]]
                if col_key == 'score':
                    # Convert to float safely
                    try:
                        score_val = pd.to_numeric(value, errors='coerce')
                        values.append(f"{score_val:.2f}" if pd.notna(score_val) else 'Chưa có điểm')
                    except:
                        values.append('Chưa có điểm')
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
    if result.empty:
        tree.insert('', 'end', values=('Không tìm thấy kết quả phù hợp', '', ''))
    else:
        # Giới hạn kết quả hiển thị
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
            msg = f"--- Hiển thị {result_limit}/{len(result)} kết quả. Thêm ký tự để tìm kiếm chi tiết ---"
            tree.insert('', 'end', values=(msg, "", ""))
                
        # Nếu chỉ tìm thấy một học sinh, tự động chọn học sinh đó
        if len(result) == 1:
            first_item = tree.get_children()[0]
            tree.selection_set(first_item)
            tree.focus(first_item)
            tree.see(first_item)
    
    # Cập nhật trạng thái

    if df is not None and len(df) > 1000:
        status_label.configure(text=f"Đã tải file: {os.path.basename(file_path)}")
                
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
        score_per_q_label.configure(text=f"({score_per_q}đ/câu)")
        
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
            undo_button.configure(state='normal')
        else:
            undo_button.configure(state='disabled')
        
        if undo_manager.can_redo():
            redo_button.configure(state='normal')
        else:
            redo_button.configure(state='disabled')
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

def customize_excel_reading():
    """Mở cửa sổ cài đặt Excel Reading (header detection, patterns, caching)"""
    excel_window = tk.Toplevel(root)
    excel_window.title("⚙️ Cài đặt đọc file Excel")
    excel_window.geometry("600x550")
    excel_window.transient(root)
    excel_window.grab_set()
    
    # Notebook for tabs
    notebook = ttk.Notebook(excel_window)
    notebook.pack(fill='both', expand=True, padx=10, pady=10)
    
    # ===== TAB 1: Header Detection =====
    header_tab = ttk.Frame(notebook)
    notebook.add(header_tab, text="📍 Phát hiện Header")
    
    header_frame = ttk.LabelFrame(header_tab, text="Cài đặt phát hiện header", padding=10)
    header_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    excel_config = config.get('excel_reading', {})
    header_config = excel_config.get('header_detection', {})
    
    # Enabled
    header_enabled_var = tk.BooleanVar(value=header_config.get('enabled', True))
    ttk.Checkbutton(header_frame, text="Bật tự động phát hiện header", 
                   variable=header_enabled_var).pack(anchor="w", pady=5)
    
    # Max search rows
    max_rows_frame = ttk.Frame(header_frame)
    max_rows_frame.pack(fill="x", pady=5)
    ttk.Label(max_rows_frame, text="Số dòng quét tối đa:").pack(side="left", padx=5)
    max_rows_var = tk.IntVar(value=header_config.get('max_search_rows', 50))
    max_rows_spin = ttk.Spinbox(max_rows_frame, from_=10, to=100, textvariable=max_rows_var, width=10)
    max_rows_spin.pack(side="left", padx=5)
    ttk.Label(max_rows_frame, text="dòng", foreground="gray").pack(side="left")
    
    # Multi-level support
    multi_level_var = tk.BooleanVar(value=header_config.get('multi_level_support', True))
    ttk.Checkbutton(header_frame, text="Hỗ trợ header nhiều cấp (merged cells)", 
                   variable=multi_level_var).pack(anchor="w", pady=5)
    
    # Separator
    sep_frame = ttk.Frame(header_frame)
    sep_frame.pack(fill="x", pady=5)
    ttk.Label(sep_frame, text="Ký tự ngăn cách khi kết hợp header:").pack(side="left", padx=5)
    sep_var = tk.StringVar(value=header_config.get('merge_separator', '_'))
    sep_entry = ttk.Entry(sep_frame, textvariable=sep_var, width=5)
    sep_entry.pack(side="left", padx=5)
    ttk.Label(sep_frame, text='(ví dụ: "ĐĐGtx" + "1" → "ĐĐGtx_1")', foreground="gray").pack(side="left")
    
    # Min columns match
    min_cols_frame = ttk.Frame(header_frame)
    min_cols_frame.pack(fill="x", pady=5)
    ttk.Label(min_cols_frame, text="Số cột tối thiểu phải khớp:").pack(side="left", padx=5)
    min_cols_var = tk.IntVar(value=header_config.get('min_columns_match', 3))
    min_cols_spin = ttk.Spinbox(min_cols_frame, from_=1, to=10, textvariable=min_cols_var, width=10)
    min_cols_spin.pack(side="left", padx=5)
    
    # Validate data rows
    validate_var = tk.BooleanVar(value=header_config.get('validate_data_rows', True))
    ttk.Checkbutton(header_frame, text="Validate dòng data tiếp theo", 
                   variable=validate_var).pack(anchor="w", pady=5)
    
    # ===== TAB 2: Merged Cells =====
    merged_tab = ttk.Frame(notebook)
    notebook.add(merged_tab, text="🔗 Merged Cells")
    
    merged_frame = ttk.LabelFrame(merged_tab, text="Xử lý merged cells", padding=10)
    merged_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    merged_config = excel_config.get('merged_cell_handling', {})
    
    merged_enabled_var = tk.BooleanVar(value=merged_config.get('enabled', True))
    ttk.Checkbutton(merged_frame, text="Bật xử lý merged cells", 
                   variable=merged_enabled_var).pack(anchor="w", pady=5)
    
    forward_fill_var = tk.BooleanVar(value=merged_config.get('forward_fill', True))
    ttk.Checkbutton(merged_frame, text="Forward fill merged cells", 
                   variable=forward_fill_var).pack(anchor="w", pady=5)
    
    combine_var = tk.BooleanVar(value=merged_config.get('combine_with_subheader', True))
    ttk.Checkbutton(merged_frame, text="Kết hợp với sub-header", 
                   variable=combine_var).pack(anchor="w", pady=5)
    
    # Info label
    info_label = ttk.Label(merged_frame, 
                          text="ℹ️ Xử lý merged cells giúp nhận diện header phức tạp\n"
                               "như 'ĐĐGtx' merge với các cột '1, 2, 3, 4'",
                          foreground="gray", justify="left")
    info_label.pack(anchor="w", pady=10)
    
    # ===== TAB 3: Performance =====
    perf_tab = ttk.Frame(notebook)
    notebook.add(perf_tab, text="⚡ Performance")
    
    perf_frame = ttk.LabelFrame(perf_tab, text="Cài đặt hiệu năng", padding=10)
    perf_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    perf_config = excel_config.get('performance', {})
    
    cache_enabled_var = tk.BooleanVar(value=perf_config.get('enable_caching', True))
    ttk.Checkbutton(perf_frame, text="Bật caching (khuyến nghị)", 
                   variable=cache_enabled_var).pack(anchor="w", pady=5)
    
    # Cache max files
    cache_frame = ttk.Frame(perf_frame)
    cache_frame.pack(fill="x", pady=5)
    ttk.Label(cache_frame, text="Số file Excel cache tối đa:").pack(side="left", padx=5)
    cache_max_var = tk.IntVar(value=perf_config.get('cache_max_files', 3))
    cache_spin = ttk.Spinbox(cache_frame, from_=1, to=10, textvariable=cache_max_var, width=10)
    cache_spin.pack(side="left", padx=5)
    
    cache_search_var = tk.BooleanVar(value=perf_config.get('cache_search_results', True))
    ttk.Checkbutton(perf_frame, text="Cache kết quả tìm kiếm", 
                   variable=cache_search_var).pack(anchor="w", pady=5)
    
    cache_stats_var = tk.BooleanVar(value=perf_config.get('cache_stats', True))
    ttk.Checkbutton(perf_frame, text="Cache thống kê", 
                   variable=cache_stats_var).pack(anchor="w", pady=5)
    
    # Clear cache button
    def clear_all_caches():
        global _df_cache, _search_cache, _stats_cache
        _df_cache.clear()
        _search_cache.clear()
        _stats_cache.clear()
        ConfigCache.invalidate()
        ToastNotification.show("✅ Đã xóa toàn bộ cache", "success")
    
    ttk.Button(perf_frame, text="🗑️ Xóa cache", command=clear_all_caches).pack(anchor="w", pady=10)
    
    # Info
    info_label2 = ttk.Label(perf_frame, 
                           text="ℹ️ Caching giúp tăng tốc load file lần 2 và search.\n"
                                "Cache tự động clear khi file/data thay đổi.",
                           foreground="gray", justify="left")
    info_label2.pack(anchor="w", pady=10)
    
    # ===== SAVE BUTTON =====
    button_frame = ttk.Frame(excel_window)
    button_frame.pack(fill="x", padx=10, pady=10)
    
    def save_excel_settings():
        # Update config
        if 'excel_reading' not in config:
            config['excel_reading'] = {}
        
        # Header detection
        config['excel_reading']['header_detection'] = {
            'enabled': header_enabled_var.get(),
            'max_search_rows': max_rows_var.get(),
            'multi_level_support': multi_level_var.get(),
            'merge_separator': sep_var.get(),
            'min_columns_match': min_cols_var.get(),
            'validate_data_rows': validate_var.get()
        }
        
        # Merged cells
        config['excel_reading']['merged_cell_handling'] = {
            'enabled': merged_enabled_var.get(),
            'forward_fill': forward_fill_var.get(),
            'combine_with_subheader': combine_var.get(),
            'cache_merged_info': merged_config.get('cache_merged_info', True)
        }
        
        # Performance
        config['excel_reading']['performance'] = {
            'enable_caching': cache_enabled_var.get(),
            'cache_max_files': cache_max_var.get(),
            'cache_search_results': cache_search_var.get(),
            'cache_stats': cache_stats_var.get()
        }
        
        # Update cache max size
        global _df_cache
        _df_cache.max_size = cache_max_var.get()
        
        save_config()
        excel_window.destroy()
        ToastNotification.show("✅ Đã lưu cài đặt Excel Reading", "success")
    
    ttk.Button(button_frame, text="💾 Lưu", command=save_excel_settings).pack(side="right", padx=5)
    ttk.Button(button_frame, text="❌ Hủy", command=excel_window.destroy).pack(side="right")

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
            password_entry.configure(show="")
            confirm_password_entry.configure(show="")
        else:
            password_entry.configure(show="*")
            confirm_password_entry.configure(show="*")
    
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
        score_col = find_matching_column(df, config['columns']['score'])
        scored_students = df[score_col].notna().sum() if score_col else 0
        
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
                password_entry.configure(state="normal")
            else:
                password_entry.delete(0, tk.END)
                password_entry.configure(state="disabled")
        
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
            score_col = find_matching_column(df, config['columns']['score'])
            scored_students = df[score_col].notna().sum() if score_col else 0
            
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
            status_label.configure(
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
        status_label.configure(text="Chưa tải file Excel")
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
    
    # ========== HEADER ==========
    header_frame = ctk.CTkFrame(root, height=55, corner_radius=0, fg_color=config['ui']['theme']['primary'])
    header_frame.pack(fill="x", side="top")
    header_frame.pack_propagate(False)
    
    header_content = ctk.CTkFrame(header_frame, fg_color="transparent")
    header_content.pack(fill="both", expand=True, padx=20, pady=10)
    
    # Title — clean, no emoji overload
    ctk.CTkLabel(header_content, 
            text="Quản Lý Điểm Học Sinh",
            font=(config['ui']['font_family'], 16, 'bold'),
            text_color='white').pack(side="left")
    
    # Right side — Cấu hình
    right_header = ctk.CTkFrame(header_content, fg_color="transparent")
    right_header.pack(side="right")
    
    ctk.CTkLabel(right_header, text="Mã đề:", text_color='white', font=(config['ui']['font_family'], 12)).pack(side="left", padx=(0, 4))
    entry_exam_code = ctk.CTkComboBox(right_header, values=config['exam_codes'], width=100, height=28, font=(config['ui']['font_family'], 12))
    entry_exam_code.pack(side="left", padx=(0, 15))
    
    ctk.CTkLabel(right_header, text="Số câu:", text_color='white', font=(config['ui']['font_family'], 12)).pack(side="left", padx=(0, 4))
    entry_max_questions = ctk.CTkEntry(right_header, width=40, height=28, font=(config['ui']['font_family'], 12), justify='center')
    entry_max_questions.insert(0, str(config['max_questions']))
    entry_max_questions.pack(side="left")
    
    score_per_q_label = ctk.CTkLabel(right_header, text=f"({config['score_per_question']}đ)", text_color="#A0AEC0", font=(config['ui']['font_family'], 10))
    score_per_q_label.pack(side="left", padx=(4, 4))
    
    ctk.CTkButton(right_header, text="Lưu", command=update_config, width=40, height=28, fg_color="#4A5568", hover_color="#2D3748").pack(side="left", padx=(0, 20))
    
    # ========== STATUSBAR ==========
    statusbar = ctk.CTkFrame(root, height=35, corner_radius=0)
    statusbar.pack(side='bottom', fill='x')
    statusbar.pack_propagate(False)
    
    status_label = ctk.CTkLabel(statusbar, text="Sẵn sàng", 
                          font=(config['ui']['font_family'], 12))
    status_label.pack(side='left', padx=15)
    
    version_display = version_utils.get_version_display()
    version_label = ctk.CTkLabel(statusbar, text=f"v{version_display}", 
                            text_color="gray",
                            font=(config['ui']['font_family'], 11))
    version_label.pack(side='right', padx=15)
    
    # ========== MAIN CONTAINER ==========
    main_container = ctk.CTkFrame(root, fg_color="transparent")
    main_container.pack(fill="both", expand=True, padx=15, pady=15)
    
    # Tỷ lệ co giãn: 7:3, minsize thấp để scale tốt trên mọi màn hình
    main_container.grid_columnconfigure(0, weight=7, minsize=500)
    main_container.grid_columnconfigure(1, weight=3, minsize=280)
    main_container.grid_rowconfigure(0, weight=1)
    
    # ========== CỘT TRÁI ==========
    left_column = ctk.CTkFrame(main_container, fg_color="transparent")
    left_column.grid(row=0, column=0, sticky='nsew', padx=(0, 8))
    
    # Toolbar — tìm kiếm + lọc + undo/redo trên cùng 1 vùng gọn
    search_frame = ctk.CTkFrame(left_column, fg_color="transparent")
    search_frame.pack(fill="x", pady=(0, 6))
    
    # Hàng 1: Ô tìm kiếm + nút thêm + undo/redo
    search_row = ctk.CTkFrame(search_frame, fg_color="transparent")
    search_row.pack(fill="x")
    
    entry_student_name = ctk.CTkEntry(search_row, height=35, placeholder_text="Nhập tên để tìm kiếm...", font=(config['ui']['font_family'], 13))
    entry_student_name.pack(side="left", fill="x", expand=True, padx=(0, 4))
    entry_student_name.bind("<KeyRelease>", delayed_search)
    
    ctk.CTkButton(search_row, text="Thêm HS", height=35, command=add_student, fg_color=config['ui']['theme']['success'], hover_color="#276749", width=80).pack(side="left", padx=(0, 4))
    
    global undo_button, redo_button
    undo_button = ctk.CTkButton(search_row, text="↩", width=35, height=35, command=perform_undo, state='disabled', fg_color="#4A5568", hover_color="#2D3748")
    undo_button.pack(side="left", padx=(0, 2))
    ToolTip(undo_button, "Hoàn tác (Ctrl+Z)")
    
    redo_button = ctk.CTkButton(search_row, text="↪", width=35, height=35, command=perform_redo, state='disabled', fg_color="#4A5568", hover_color="#2D3748")
    redo_button.pack(side="left")
    ToolTip(redo_button, "Làm lại (Ctrl+Y)")
    
    # Hàng 2: Bộ lọc nhanh
    filter_frame = ctk.CTkFrame(search_frame, fg_color="transparent")
    filter_frame.pack(fill="x", pady=(6, 0))
    
    def apply_quick_filter(filter_type):
        """Áp dụng filter nhanh và gọi tìm kiếm"""
        global current_filter_type
        current_filter_type = filter_type
        for btn, type_key in quick_filter_btns:
            if type_key == filter_type:
                btn.configure(fg_color=config['ui']['theme']['primary'])
            else:
                btn.configure(fg_color="#4A5568")
        search_student()
        
    global quick_filter_btns
    quick_filter_btns = []
    
    filter_labels = [
        ("Tất cả", 'all'), ("≥7", 'high'), ("5-7", 'medium'),
        ("<5", 'low'), ("Chưa", 'no_score')
    ]
    for label, key in filter_labels:
        btn = ctk.CTkButton(filter_frame, text=label, command=lambda k=key: apply_quick_filter(k), width=60, height=28, fg_color="#4A5568", hover_color=config['ui']['theme']['primary_dark'])
        btn.pack(side="left", padx=(0, 4))
        if key == 'all':
            btn.configure(fg_color=config['ui']['theme']['primary'])
        quick_filter_btns.append((btn, key))
    
    # Danh sách học sinh (vẫn giữ ttk.Treeview vì ctk không có bảng mặc định tốt)
    list_frame = ctk.CTkFrame(left_column)
    list_frame.pack(fill="both", expand=True, pady=(5, 0))
    
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
            
            # Cập nhật header với mũi tên sắp xếp
            arrow = " ▼" if sort_column['reverse'] else " ▲"
            tree.heading('name', text=f"{config['columns']['name']}{arrow if col == 'name' else ''}")
            tree.heading('exam_code', text=f"Mã Đề{arrow if col == 'exam_code' else ''}")
            tree.heading('score', text=f"Điểm{arrow if col == 'score' else ''}")
    
    tree.heading('name', text=config['columns']['name'], command=lambda: sort_treeview('name'))
    tree.heading('exam_code', text="Mã Đề", command=lambda: sort_treeview('exam_code'))
    tree.heading('score', text="Điểm", command=lambda: sort_treeview('score'))
    
    tree.column('name', width=500, minwidth=200)
    tree.column('exam_code', width=120, minwidth=80, anchor='center')
    tree.column('score', width=120, minwidth=80, anchor='center')
    
    # Màu sắc dịu hơn (giảm độ chói của nền) nhưng chữ nổi bật hơn
    tree.tag_configure('high_score', background='#022c22', foreground='#34d399', font=(config['ui']['font_family'], 11, 'bold'))
    tree.tag_configure('medium_score', background='#451a03', foreground='#fbbf24', font=(config['ui']['font_family'], 11, 'bold'))
    tree.tag_configure('low_score', background='#450a0a', foreground='#f87171', font=(config['ui']['font_family'], 11, 'bold'))
    tree.tag_configure('no_score', background='#1f2937', foreground='#ffffff', font=(config['ui']['font_family'], 11))
    
    vsb = ttk.Scrollbar(list_frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(list_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    
    tree.grid(column=0, row=0, sticky='nsew', padx=2, pady=2)
    vsb.grid(column=1, row=0, sticky='ns', pady=2)
    hsb.grid(column=0, row=1, sticky='ew', padx=2)
    
    list_frame.grid_columnconfigure(0, weight=1)
    list_frame.grid_rowconfigure(0, weight=1)
    
    # ========== CỘT PHẢI ==========
    scrollable_right_frame = ctk.CTkScrollableFrame(main_container, fg_color="transparent")
    scrollable_right_frame.grid(row=0, column=1, sticky='nsew', padx=(10, 0))
    
    # ── Cấu hình số câu ──
    config_card = ctk.CTkFrame(scrollable_right_frame, fg_color=config['ui']['theme']['card'], corner_radius=8)
    # Nhập Điểm Frame retains just score input now
    # ── Nhập điểm ──
    score_card = ctk.CTkFrame(scrollable_right_frame, fg_color=config['ui']['theme']['card'], corner_radius=8)
    score_card.pack(fill="x", pady=(0, 10))
    
    ctk.CTkLabel(score_card, text="Nhập Điểm", font=(config['ui']['font_family'], 14, "bold"), text_color=config['ui']['theme']['primary']).pack(anchor="w", padx=15, pady=(10, 5))
    
    # Số câu đúng + tính điểm
    ctk.CTkLabel(score_card, text="Số câu đúng", font=(config['ui']['font_family'], 12)).pack(anchor="w", padx=15)
    entry_correct_count = ctk.CTkEntry(score_card, height=36,
                                  font=(config['ui']['font_family'], 16, "bold"),
                                  justify='center')
    entry_correct_count.pack(fill="x", padx=15, pady=(2, 6))
    entry_correct_count.bind("<Return>", calculate_score)
    
    ctk.CTkButton(score_card, text="Tính Điểm", height=36,
              command=calculate_score).pack(fill="x", padx=15, pady=(4, 15))
    
    # Nhập điểm trực tiếp
    ctk.CTkLabel(score_card, text="Nhập điểm trực tiếp", font=(config['ui']['font_family'], 12)).pack(anchor="w", padx=15)
    entry_direct_score = ctk.CTkEntry(score_card, height=36,
                               font=(config['ui']['font_family'], 16, "bold"),
                               justify='center')
    entry_direct_score.pack(fill="x", padx=15, pady=(2, 6))
    entry_direct_score.bind("<Return>", calculate_score_direct)
    
    ctk.CTkButton(score_card, text="Lưu Điểm", height=36, fg_color=config['ui']['theme']['success'], hover_color="#276749",
              command=calculate_score_direct).pack(fill="x", padx=15, pady=(4, 15))
    
    # ── Thống kê ──
    stats_card = ctk.CTkFrame(scrollable_right_frame, fg_color=config['ui']['theme']['card'], corner_radius=8)
    stats_card.pack(fill="x", pady=(0, 10))
    
    ctk.CTkLabel(stats_card, text="Thống Kê", font=(config['ui']['font_family'], 14, "bold"), text_color=config['ui']['theme']['primary']).pack(anchor="w", padx=15, pady=(10, 5))
    
    progress_frame = ctk.CTkFrame(stats_card, fg_color="transparent")
    progress_frame.pack(fill="x", padx=15, pady=(0, 6))
    
    ctk.CTkLabel(progress_frame, text="Tiến độ nhập điểm", 
             font=(config['ui']['font_family'], 12)).pack(anchor="w")
    progress_bar = ctk.CTkProgressBar(progress_frame)
    progress_bar.pack(fill="x", pady=(4, 4))
    progress_bar.set(0)
    progress_label = ctk.CTkLabel(progress_frame, text="0%", 
                              font=(config['ui']['font_family'], 12, 'bold'),
                              text_color=config['ui']['theme']['primary'])
    progress_label.pack(anchor="w")
    
    stats_label = ctk.CTkLabel(stats_card, text="0/0 có điểm", 
                          font=(config['ui']['font_family'], 14, 'bold'))
    stats_label.pack(pady=4)
    
    category_frame = ctk.CTkFrame(stats_card, fg_color="transparent")
    category_frame.pack(fill="x", padx=15, pady=(4, 15))
    
    high_score_count = ctk.CTkLabel(category_frame, text=" Giỏi (≥7): 0 ", font=(config['ui']['font_family'], 12, "bold"), fg_color="#022c22", text_color="#34d399", corner_radius=6)
    high_score_count.pack(anchor="w", pady=2)
    
    medium_score_count = ctk.CTkLabel(category_frame, text=" Khá (5-7): 0 ", font=(config['ui']['font_family'], 12, "bold"), fg_color="#451a03", text_color="#fbbf24", corner_radius=6)
    medium_score_count.pack(anchor="w", pady=2)
    
    low_score_count = ctk.CTkLabel(category_frame, text=" Yếu (<5): 0 ", font=(config['ui']['font_family'], 12, "bold"), fg_color="#450a0a", text_color="#f87171", corner_radius=6)
    low_score_count.pack(anchor="w", pady=2)
    
    no_score_count = ctk.CTkLabel(category_frame, text=" Chưa có điểm: 0 ", font=(config['ui']['font_family'], 12), fg_color="#1f2937", text_color="#ffffff", corner_radius=6)
    no_score_count.pack(anchor="w", pady=2)
    
    highest_score_label = ctk.CTkLabel(stats_card, text="Cao nhất: N/A", font=(config['ui']['font_family'], 12))
    highest_score_label.pack(anchor="w", padx=15, pady=1)
    
    lowest_score_label = ctk.CTkLabel(stats_card, text="Thấp nhất: N/A", font=(config['ui']['font_family'], 12))
    lowest_score_label.pack(anchor="w", padx=15, pady=(1, 15))
    
    # ── Quản lý File ──
    file_card = ctk.CTkFrame(scrollable_right_frame, fg_color=config['ui']['theme']['card'], corner_radius=8)
    file_card.pack(fill="x", pady=(0, 10))
    
    ctk.CTkLabel(file_card, text="Quản lý File", font=(config['ui']['font_family'], 14, "bold"), text_color=config['ui']['theme']['primary']).pack(anchor="w", padx=15, pady=(10, 5))
    
    ctk.CTkButton(file_card, text="Mở File", height=32, command=select_file).pack(fill="x", padx=15, pady=(2, 6))
    ctk.CTkButton(file_card, text="Sao Lưu", height=32, command=backup_data).pack(fill="x", padx=15, pady=(2, 6))
    ctk.CTkButton(file_card, text="Xuất Báo Cáo", height=32, command=generate_report, fg_color="#4A5568", hover_color="#2D3748").pack(fill="x", padx=15, pady=(2, 15))
    
    # Menu
    menubar = tk.Menu(root, 
                     background=config['ui']['theme']['card'],
                     foreground=config['ui']['theme']['text'],
                     activebackground=config['ui']['theme']['primary'],
                     activeforeground='white',
                     borderwidth=0)
    root.configure(menu=menubar)
    
    settings_menu = tk.Menu(menubar, tearoff=0,
                           background=config['ui']['theme']['card'],
                           foreground=config['ui']['theme']['text'],
                           activebackground=config['ui']['theme']['primary'],
                           activeforeground='white')
    menubar.add_cascade(label="⚙️ Cài đặt", menu=settings_menu)
    settings_menu.add_command(label="⌨️ Phím tắt", command=customize_shortcuts)
    settings_menu.add_command(label="🔢 Mã đề", command=customize_exam_codes)
    settings_menu.add_command(label="📝 Tên cột", command=customize_columns)
    settings_menu.add_command(label="📊 Đọc Excel", command=customize_excel_reading)
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
    about_text_widget.configure(yscrollcommand=scroll.set)
    
    # Chèn text
    about_text_widget.insert("1.0", about_text)
    about_text_widget.configure(state="disabled")  # Không cho phép chỉnh sửa
    
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
    """Cập nhật các thống kê cơ bản với progress bar và phân loại chi tiết, sử dụng cache"""
    global progress_bar, progress_label, high_score_count, medium_score_count, low_score_count, no_score_count, _stats_cache
    
    if df is None or df.empty:
        stats_label.configure(text="Không có dữ liệu")
        progress_bar['value'] = 0
        progress_label.configure(text="0%")
        high_score_count.configure(text=" Giỏi (≥7): 0 ")
        medium_score_count.configure(text=" Khá (5-7): 0 ")
        low_score_count.configure(text=" Yếu (<5): 0 ")
        no_score_count.configure(text=" Chưa có điểm: 0 ")
        return
        
    total_students = len(df)
    score_column_config = config['columns']['score']
    
    # Tìm cột điểm thực tế trong DataFrame
    score_column = find_matching_column(df, score_column_config)
    
    # Kiểm tra cột điểm có tồn tại không
    if not score_column:
        stats_label.configure(text=f"Không tìm thấy cột '{score_column_config}'")
        progress_bar['value'] = 0
        progress_label.configure(text="0%")
        high_score_count.configure(text=" Giỏi (≥7): 0 ")
        medium_score_count.configure(text=" Khá (5-7): 0 ")
        low_score_count.configure(text=" Yếu (<5): 0 ")
        no_score_count.configure(text=" Chưa có điểm: 0 ")
        return
    
    # Get basic stats from cache
    cached_stats = _stats_cache.get_stats(df)
    students_with_scores = cached_stats['scored']
    students_no_scores = total_students - students_with_scores
    
    # Tính phần trăm
    percentage = (students_with_scores / total_students * 100) if total_students > 0 else 0
    
    # Cập nhật progress bar
    progress_bar['value'] = percentage
    progress_label.configure(text=f"{percentage:.1f}% ({students_with_scores}/{total_students})")
    
    # Phân loại theo điểm
    df_with_scores = df[df[score_column].notna()].copy()
    # Chuyển đổi cột điểm sang dạng số, các giá trị không hợp lệ sẽ thành NaN
    df_with_scores[score_column] = pd.to_numeric(df_with_scores[score_column], errors='coerce')
    # Loại bỏ các giá trị NaN sau khi chuyển đổi
    df_with_scores = df_with_scores[df_with_scores[score_column].notna()]
    
    high_count = len(df_with_scores[df_with_scores[score_column] >= 7])
    medium_count = len(df_with_scores[(df_with_scores[score_column] >= 5) & (df_with_scores[score_column] < 7)])
    low_count = len(df_with_scores[df_with_scores[score_column] < 5])
    
    # Cập nhật labels
    stats_label.configure(text=f"{students_with_scores}/{total_students} đã có điểm")
    high_score_count.configure(text=f" Giỏi (≥7): {high_count} ({high_count/total_students*100:.1f}%)" if total_students > 0 else " Giỏi (≥7): 0 ")
    medium_score_count.configure(text=f" Khá (5-7): {medium_count} ({medium_count/total_students*100:.1f}%)" if total_students > 0 else " Khá (5-7): 0 ")
    low_score_count.configure(text=f" Yếu (<5): {low_count} ({low_count/total_students*100:.1f}%)" if total_students > 0 else " Yếu (<5): 0 ")
    no_score_count.configure(text=f" Chưa có điểm: {students_no_scores} ({students_no_scores/total_students*100:.1f}%)" if total_students > 0 else " Chưa có điểm: 0 ")

def update_score_extremes():
    """Cập nhật điểm cao nhất và thấp nhất"""
    if df is None or df.empty:
        highest_score_label.configure(text="🏆 Cao nhất: N/A")
        lowest_score_label.configure(text="📉 Thấp nhất: N/A")
        return
        
    score_column = find_matching_column(df, config['columns']['score'])
    name_column = find_matching_column(df, config['columns']['name'])
    
    # Kiểm tra cột điểm có tồn tại không
    if not score_column or not name_column:
        highest_score_label.configure(text="🏆 Cao nhất: N/A")
        lowest_score_label.configure(text="📉 Thấp nhất: N/A")
        return
    
    # Lọc các hàng có điểm không phải NaN
    df_with_scores = df[df[score_column].notna()].copy()
    # Chuyển đổi cột điểm sang dạng số
    df_with_scores[score_column] = pd.to_numeric(df_with_scores[score_column], errors='coerce')
    # Loại bỏ các giá trị NaN sau khi chuyển đổi
    df_with_scores = df_with_scores[df_with_scores[score_column].notna()]
    
    # Nếu không có ai có điểm
    if df_with_scores.empty:
        highest_score_label.configure(text="🏆 Cao nhất: N/A")
        lowest_score_label.configure(text="📉 Thấp nhất: N/A")
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
    highest_score_label.configure(text=f"🏆 Cao nhất: {max_score:.2f} ({max_students_text})")
    lowest_score_label.configure(text=f"📉 Thấp nhất: {min_score:.2f} ({min_students_text})")

def delayed_search(event=None):
    """Tìm kiếm sau một khoảng thời gian để tránh quá nhiều tìm kiếm liên tiếp"""
    global search_timer_id
    
    # Hủy timer cũ nếu có
    if search_timer_id is not None:
        root.after_cancel(search_timer_id)
    
    # Thiết lập timer mới
    search_timer_id = root.after(300, search_student)  # 300ms delay

def generate_report():
    """Tạo báo cáo PDF chuyên nghiệp với thiết kế hiện đại và căn chỉnh đồng bộ"""
    global df
    
    if df is None or df.empty:
        messagebox.showwarning("Thông báo", "Không có dữ liệu để tạo báo cáo")
        return
        
    if 'Điểm' not in df.columns:
        messagebox.showwarning("Thông báo", "Không tìm thấy cột điểm trong dữ liệu")
        return
        
    # Lấy thông tin thống kê giống UI
    scores_df = df[df['Điểm'].notna()].copy()
    if len(scores_df) == 0:
        messagebox.showwarning("Thông báo", "Chưa có học sinh nào có điểm để làm báo cáo")
        return
    
    try:
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Lưu báo cáo PDF Chuyên Nghiệp"
        )
        if not file_path: return

        # Giao diện tiến trình
        progress_window = tk.Toplevel(root)
        progress_window.title("Hệ thống báo cáo")
        progress_window.geometry("450x220")
        ui_utils.center_window(progress_window, 450, 220)
        progress_window.transient(root)
        
        progress_frame = ttk.Frame(progress_window, padding=20)
        progress_frame.pack(fill="both", expand=True)
        
        ttk.Label(progress_frame, text="ĐANG KHỞI TẠO BÁO CÁO", 
                 font=(config['ui']['font_family'], 12, 'bold')).pack(pady=(0, 10))
                 
        progress = ttk.Progressbar(progress_frame, orient="horizontal", length=400, mode="determinate")
        progress.pack(pady=10, fill="x")
        
        status_label_progress = ttk.Label(progress_frame, text="Đang chuẩn bị...")
        status_label_progress.pack()

        def update_p(val, msg):
            progress["value"] = val
            status_label_progress.configure(text=msg)
            progress_window.update()

        # THIẾT LẬP MÀU SẮC (Indigo Theme)
        COLOR_PRIMARY = "#5A67D8"  # Indigo
        COLOR_SUCCESS = "#38A169"  # Green
        COLOR_WARNING = "#DD6B20"  # Orange
        COLOR_DANGER = "#E53E3E"   # Red
        COLOR_TEXT = "#2D3748"
        COLOR_BG_LIGHT = "#F7FAFC"

        # TẾP PDF CHÍNH
        with PdfPages(file_path) as pdf:
            # --- TRANG 1: THỐNG KÊ TỔNG QUAN ---
            update_p(20, "Đang thiết kế trang tổng quan...")
            fig = plt.figure(figsize=(8.27, 11.69)) # A4 size
            
            # 1. Header Bar
            plt.fill_between([0, 1], 0.92, 1.0, color=COLOR_PRIMARY, transform=fig.transFigure)
            plt.text(0.5, 0.96, "BÁO CÁO KẾT QUẢ HỌC TẬP", color='white', 
                    fontsize=18, fontweight='bold', ha='center', transform=fig.transFigure)
            plt.text(0.5, 0.935, f"Ngày xuất bản: {datetime.now().strftime('%d/%m/%Y %H:%M')}", 
                    color='white', fontsize=10, ha='center', transform=fig.transFigure)

            # 2. Metrics Cards (Metric Boxes)
            avg_score = scores_df['Điểm'].mean()
            max_score = scores_df['Điểm'].max()
            min_score = scores_df['Điểm'].min()
            
            # Vẽ các ô tóm tắt
            metrics = [
                ("TRUNG BÌNH", f"{avg_score:.2f}", COLOR_PRIMARY),
                ("CAO NHẤT", f"{max_score:.2f}", COLOR_SUCCESS),
                ("THẤP NHẤT", f"{min_score:.2f}", COLOR_DANGER)
            ]
            
            for i, (label, val, color) in enumerate(metrics):
                x_pos = 0.15 + i * 0.28
                rect = plt.Rectangle((x_pos, 0.82), 0.22, 0.07, color=COLOR_BG_LIGHT, 
                                     transform=fig.transFigure, zorder=0)
                fig.patches.append(rect)
                plt.text(x_pos + 0.11, 0.865, label, fontsize=9, fontweight='bold', 
                        color=color, ha='center', transform=fig.transFigure)
                plt.text(x_pos + 0.11, 0.835, val, fontsize=16, fontweight='bold', 
                        color=COLOR_TEXT, ha='center', transform=fig.transFigure)

            # 3. Biểu đồ phân phối (Histogram + KDE)
            update_p(40, "Đang vẽ biểu đồ thống kê...")
            ax1 = fig.add_axes([0.1, 0.52, 0.8, 0.25])
            n, bins, patches = ax1.hist(scores_df['Điểm'], bins=10, range=(0,10), 
                                       color=COLOR_PRIMARY, alpha=0.6, rwidth=0.85)
            ax1.set_title('Phân phối điểm số toàn lớp', fontsize=12, fontweight='bold', pad=15)
            ax1.set_xlabel('Thang điểm', fontsize=10)
            ax1.set_ylabel('Số lượng học sinh', fontsize=10)
            ax1.spines['top'].set_visible(False)
            ax1.spines['right'].set_visible(False)
            ax1.grid(axis='y', linestyle='--', alpha=0.4)
            
            # 4. Biểu đồ tròn Phân loại
            update_p(60, "Đang tính toán tỷ lệ xếp loại...")
            ax2 = fig.add_axes([0.1, 0.15, 0.35, 0.25])
            
            gioi = len(scores_df[scores_df['Điểm'] >= 7])
            kha = len(scores_df[(scores_df['Điểm'] >= 5) & (scores_df['Điểm'] < 7)])
            yeu = len(scores_df[scores_df['Điểm'] < 5])
            
            labels = [f'Giỏi (≥7)\n{gioi} HS', f'Khá (5-7)\n{kha} HS', f'Yếu (<5)\n{yeu} HS']
            sizes = [gioi, kha, yeu]
            colors = [COLOR_SUCCESS, COLOR_WARNING, COLOR_DANGER]
            
            if sum(sizes) > 0:
                ax2.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', 
                       startangle=140, pctdistance=0.75, wedgeprops={'edgecolor': 'white', 'linewidth': 2})
                ax2.set_title('Tỷ lệ xếp loại học sinh', fontsize=11, fontweight='bold')
            
            # 5. Top 5 Vinh Danh
            ax3 = fig.add_axes([0.55, 0.15, 0.35, 0.25])
            ax3.axis('off')
            ax3.set_title('Top 5 Học sinh xuất sắc', fontsize=11, fontweight='bold', pad=10)
            
            name_col = config['columns']['name']
            top_5 = scores_df.sort_values('Điểm', ascending=False).head(5)
            top_data = [[f"{i+1}. {r[name_col]}", f"{r['Điểm']:.2f}"] for i, r in enumerate(top_5.to_dict('records'))]
            
            table_top = ax3.table(cellText=top_data, colLabels=["Họ và Tên", "Điểm"], 
                                 loc='center', cellLoc='left')
            table_top.auto_set_font_size(False)
            table_top.set_fontsize(9)
            table_top.scale(1, 1.8)
            # Style table top
            for (row, col), cell in table_top.get_celld().items():
                cell.set_edgecolor('#E2E8F0')
                if row == 0: 
                    cell.set_facecolor('#EDF2F7')
                    cell.set_text_props(weight='bold')

            # Footer Page 1
            plt.text(0.5, 0.05, "Trang 1", fontsize=9, color='gray', ha='center', transform=fig.transFigure)
            
            pdf.savefig(fig)
            plt.close(fig)

            # --- TRANG 2+: DANH SÁCH CHI TIẾT ---
            update_p(80, "Đang tạo danh sách chi tiết...")
            sorted_all = df.sort_values(config['columns']['name'])
            rows_per_page = 32
            total_pages = (len(sorted_all)-1)//rows_per_page + 1
            
            for p in range(total_pages):
                update_p(80 + (p/total_pages)*20, f"Đang xuất danh sách trang {p+1}/{total_pages}...")
                fig_list = plt.figure(figsize=(8.27, 11.69))
                plt.axis('off')
                
                # Header trang phụ
                plt.fill_between([0, 1], 0.95, 1.0, color=COLOR_PRIMARY, transform=fig_list.transFigure)
                plt.text(0.1, 0.97, "DANH SÁCH CHI TIẾT ĐIỂM SỐ", color='white', 
                        fontsize=12, fontweight='bold', transform=fig_list.transFigure)
                
                ax_list = fig_list.add_axes([0.05, 0.08, 0.9, 0.85])
                ax_list.axis('off')
                
                start_r = p * rows_per_page
                end_r = min((p+1) * rows_per_page, len(sorted_all))
                page_df = sorted_all.iloc[start_r:end_r]
                
                # Chuẩn bị data cho bảng
                col_name = config['columns']['name']
                col_exam = 'Mã đề' if 'Mã đề' in page_df.columns else None
                
                headers = ["STT", "Họ và Tên", "Mã đề", "Điểm số"] if col_exam else ["STT", "Họ và Tên", "Điểm số"]
                data_rows = []
                for i, (_, row) in enumerate(page_df.iterrows()):
                    score_val = f"{row['Điểm']:.2f}" if pd.notna(row['Điểm']) else "Chưa có"
                    row_data = [start_r + i + 1, row[col_name]]
                    if col_exam: row_data.append(row['Mã đề'] if pd.notna(row['Mã đề']) else "-")
                    row_data.append(score_val)
                    data_rows.append(row_data)
                
                table_list = ax_list.table(cellText=data_rows, colLabels=headers, loc='upper center', cellLoc='center')
                table_list.auto_set_font_size(False)
                table_list.set_fontsize(9)
                table_list.scale(1, 1.8)
                
                # Định dạng bảng: Zebra stripes và Header đậm
                for (row, col), cell in table_list.get_celld().items():
                    cell.set_edgecolor('#CBD5E0')
                    cell.set_linewidth(0.5)
                    if row == 0:
                        cell.set_facecolor(COLOR_PRIMARY)
                        cell.set_text_props(color='white', weight='bold')
                    elif row % 2 == 0:
                        cell.set_facecolor('#F7FAFC')
                    
                    # Căn trái cho cột tên
                    if col == 1: cell.set_text_props(ha='left')

                plt.text(0.5, 0.04, f"Trang {p+2} / {total_pages+1}", fontsize=9, color='gray', ha='center', transform=fig_list.transFigure)
                
                pdf.savefig(fig_list)
                plt.close(fig_list)
        
        # Kết thúc
        update_p(100, "HOÀN TẤT!")
        progress_window.after(1000, progress_window.destroy)
        
        status_label.configure(text=f"Báo cáo sẵn sàng: {os.path.basename(file_path)}")
        ToastNotification.show(f"✅ Báo cáo chuyên nghiệp đã được xuất!\nLưu tại: {file_path}", "success")
        
    except Exception as e:
        status_label.configure(text="Lỗi xuất báo cáo!")
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

def detect_merged_cells_structure(file_path, config):
    """
    Phát hiện cấu trúc merged cells từ Excel file
    Returns: dict với thông tin về merged cells và multi-level headers
    """
    try:
        if not config.get('excel_reading', {}).get('merged_cell_handling', {}).get('enabled', True):
            return None
        
        # Không dùng read_only mode để có thể access merged_cells
        wb = load_workbook(file_path, data_only=True)
        ws = wb.active
        
        merged_ranges = list(ws.merged_cells.ranges)
        
        if not merged_ranges:
            wb.close()
            return None
        
        # Parse merged cell info
        merged_info = {}
        for merged_range in merged_ranges:
            min_col, min_row, max_col, max_row = merged_range.bounds
            # Lưu thông tin: cell nào được merge với cells nào
            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    merged_info[(row, col)] = {
                        'parent': (min_row, min_col),
                        'is_parent': (row == min_row and col == min_col),
                        'bounds': (min_row, min_col, max_row, max_col)
                    }
        
        wb.close()
        return merged_info
        
    except Exception as e:
        print(f"Error detecting merged cells: {e}")
        return None


def combine_multi_level_headers(df_headers, merged_info, config):
    """
    Kết hợp multi-level headers thành tên cột đầy đủ
    df_headers: DataFrame chứa các dòng header
    merged_info: dict thông tin về merged cells
    """
    try:
        separator = config.get('excel_reading', {}).get('header_detection', {}).get('merge_separator', '_')
        
        if df_headers.shape[0] == 1:
            # Chỉ có 1 dòng header, không cần combine
            return list(df_headers.iloc[0])
        
        combined_headers = []
        num_cols = df_headers.shape[1]
        
        for col_idx in range(num_cols):
            col_parts = []
            
            for row_idx in range(df_headers.shape[0]):
                value = df_headers.iloc[row_idx, col_idx]
                
                # Skip NaN và empty values
                if pd.notna(value) and str(value).strip():
                    col_parts.append(str(value).strip())
            
            # Kết hợp các phần với separator
            if col_parts:
                combined_name = separator.join(col_parts)
                combined_headers.append(combined_name)
            else:
                combined_headers.append(f"Column_{col_idx}")
        
        return combined_headers
        
    except Exception as e:
        print(f"Error combining headers: {e}")
        return list(df_headers.iloc[0]) if len(df_headers) > 0 else []


def match_column_pattern(column_name, config):
    """
    Match column name với patterns trong config
    Returns: dict với thông tin về pattern matched hoặc None
    """
    try:
        patterns_config = config.get('excel_reading', {}).get('column_patterns', {})
        column_name_lower = str(column_name).lower().strip()
        
        # Check student info patterns
        for field_name, field_config in patterns_config.get('student_info', {}).items():
            # Try exact patterns first
            patterns = field_config.get('patterns', [])
            for pattern in patterns:
                if pattern.lower() in column_name_lower or column_name_lower in pattern.lower():
                    return {
                        'type': 'student_info',
                        'field': field_name,
                        'matched_pattern': pattern,
                        'required': field_config.get('required', False)
                    }
            
            # Try regex
            regex = field_config.get('regex', '')
            if regex:
                try:
                    if re.match(regex, column_name, re.IGNORECASE):
                        return {
                            'type': 'student_info',
                            'field': field_name,
                            'matched_pattern': regex,
                            'required': field_config.get('required', False)
                        }
                except:
                    pass
        
        # Check score columns patterns
        for score_type, score_config in patterns_config.get('score_columns', {}).items():
            patterns = score_config.get('patterns', [])
            for pattern in patterns:
                if pattern.lower() in column_name_lower or column_name_lower in pattern.lower():
                    return {
                        'type': 'score',
                        'score_type': score_type,
                        'matched_pattern': pattern,
                        'has_sub_columns': score_config.get('has_sub_columns', False),
                        'description': score_config.get('description', '')
                    }
            
            # Try regex
            regex = score_config.get('regex', '')
            if regex:
                try:
                    match = re.match(regex, column_name, re.IGNORECASE)
                    if match:
                        result = {
                            'type': 'score',
                            'score_type': score_type,
                            'matched_pattern': regex,
                            'has_sub_columns': score_config.get('has_sub_columns', False),
                            'description': score_config.get('description', '')
                        }
                        # Extract sub-column number if exists
                        if match.groups():
                            result['sub_column'] = match.group(len(match.groups()))
                        return result
                except:
                    pass
        
        return None
        
    except Exception as e:
        print(f"Error matching column pattern: {e}")
        return None


def score_header_row(row_data, config):
    """
    Tính điểm cho một dòng dựa trên số lượng cột khớp với patterns
    Returns: float score (cao hơn = nhiều khả năng là header)
    """
    try:
        score = 0
        patterns_config = config.get('excel_reading', {}).get('column_patterns', {})
        
        # Keywords để tìm trong header row
        student_keywords = ['họ và tên', 'tên học sinh', 'tên', 'họ tên', 'học sinh', 
                           'mã số', 'mssv', 'mã định danh', 'stt', 'số tt']
        score_keywords = ['điểm', 'đđgtx', 'đđggk', 'đđgck', 'đtb', 'score', 'tx', 'gk', 'ck']
        other_keywords = ['ngày sinh', 'giới tính', 'phái', 'lớp', 'khối']
        
        row_str = ' '.join([str(val).lower() for val in row_data if pd.notna(val)])
        
        # Điểm cho student info keywords
        for keyword in student_keywords:
            if keyword in row_str:
                score += 3
        
        # Điểm cho score keywords
        for keyword in score_keywords:
            if keyword in row_str:
                score += 2
        
        # Điểm cho other keywords
        for keyword in other_keywords:
            if keyword in row_str:
                score += 1
        
        # Bonus nếu có nhiều cột không trống
        non_empty_cols = sum(1 for val in row_data if pd.notna(val) and str(val).strip())
        score += non_empty_cols * 0.5
        
        return score
        
    except Exception as e:
        print(f"Error scoring header row: {e}")
        return 0


def find_header_row(headers_df, config=None):
    """Tìm hàng chứa header thực sự trong DataFrame với scoring mechanism."""
    try:
        if config is None:
            # Fallback to old simple logic
            header_row = None
            for i in range(len(headers_df)):
                row_values = headers_df.iloc[i].astype(str)
                if any(name.lower() in ' '.join(row_values.str.lower()) 
                      for name in ['họ và tên', 'tên học sinh', 'học sinh']):
                    header_row = i
                    break
            
            if header_row is None:
                header_row = 0
            
            return header_row
        
        # New scoring-based logic
        max_score = 0
        best_header_row = 0
        min_columns_match = config.get('excel_reading', {}).get('header_detection', {}).get('min_columns_match', 3)
        
        for i in range(len(headers_df)):
            row_data = headers_df.iloc[i]
            score = score_header_row(row_data, config)
            
            if score > max_score:
                max_score = score
                best_header_row = i
        
        # Validate: dòng tiếp theo phải có data hợp lệ
        validate = config.get('excel_reading', {}).get('header_detection', {}).get('validate_data_rows', True)
        if validate and best_header_row < len(headers_df) - 1:
            next_row = headers_df.iloc[best_header_row + 1]
            # Check xem dòng tiếp theo có phải là data không
            non_empty = sum(1 for val in next_row if pd.notna(val) and str(val).strip())
            if non_empty < 2:  # Quá ít data, có thể không phải header đúng
                # Try next row
                if best_header_row + 1 < len(headers_df):
                    best_header_row += 1
        
        return best_header_row
        
    except Exception as e:
        print(f"Error finding header row: {e}")
        return 0


def read_excel_file(file_path):
    """Đọc file Excel theo cách thông thường với caching và merged cells support"""
    global _df_cache
    status_label.configure(text=f"Đang đọc file Excel...")
    root.update()
    
    try:
        # Check cache trước
        config = load_config()
        enable_caching = config.get('excel_reading', {}).get('performance', {}).get('enable_caching', True)
        
        if enable_caching:
            cached_df = _df_cache.get(file_path)
            if cached_df is not None:
                status_label.configure(text=f"Đã load file từ cache")
                return cached_df
        
        # Đọc header với max_search_rows từ config
        max_rows = config.get('excel_reading', {}).get('header_detection', {}).get('max_search_rows', 50)
        
        # Detect merged cells structure nếu enabled
        merged_info = None
        if config.get('excel_reading', {}).get('merged_cell_handling', {}).get('enabled', True):
            merged_info = detect_merged_cells_structure(file_path, config)
        
        # Thử đọc với engine openpyxl trước
        headers_df = pd.read_excel(file_path, nrows=max_rows, engine='openpyxl')
        
        # Tìm hàng chứa header thực sự với scoring
        header_row = find_header_row(headers_df, config)
        
        # Check nếu có multi-level headers
        multi_level = config.get('excel_reading', {}).get('header_detection', {}).get('multi_level_support', True)
        if multi_level and merged_info and header_row < len(headers_df) - 1:
            # Đọc thêm vài dòng sau header row để check multi-level
            header_rows_to_read = min(3, len(headers_df) - header_row)
            if header_rows_to_read > 1:
                # Đọc với header là list các rows
                df_result = pd.read_excel(file_path, header=list(range(header_row, header_row + header_rows_to_read)), engine='openpyxl')
                # Combine multi-level headers
                if isinstance(df_result.columns, pd.MultiIndex):
                    # Flatten MultiIndex columns
                    separator = config.get('excel_reading', {}).get('header_detection', {}).get('merge_separator', '_')
                    df_result.columns = [separator.join(str(i) for i in col if str(i) != 'nan').strip(separator) 
                                        for col in df_result.columns.values]
            else:
                # Đọc bình thường
                df_result = pd.read_excel(file_path, header=header_row, engine='openpyxl')
        else:
            # Đọc lại với header đúng
            df_result = pd.read_excel(file_path, header=header_row, engine='openpyxl')
        
        # Xử lý trường hợp DataFrame rỗng
        if df_result.empty:
            status_label.configure(text="File Excel không có dữ liệu")
            return pd.DataFrame()
        
        # Đảm bảo các cột cần thiết tồn tại
        df_result = ensure_required_columns(df_result)
        
        # Đảm bảo kiểu dữ liệu phù hợp
        df_result = ensure_proper_dtypes(df_result)
        
        # Cache result
        if enable_caching:
            _df_cache.set(file_path, df_result)
        
        status_label.configure(text=f"Đã đọc xong file Excel")
        return df_result
        
    except ImportError:
        # Nếu không có openpyxl, thử dùng engine mặc định
        print("Không tìm thấy openpyxl, sử dụng engine mặc định...")
        config = load_config()
        max_rows = config.get('excel_reading', {}).get('header_detection', {}).get('max_search_rows', 50)
        headers_df = pd.read_excel(file_path, nrows=max_rows)
        
        # Tìm hàng chứa header thực sự
        header_row = find_header_row(headers_df, config)
        
        # Đọc lại với header đúng
        df_result = pd.read_excel(file_path, header=header_row)
        
        # Xử lý trường hợp DataFrame rỗng
        if df_result.empty:
            status_label.configure(text="File Excel không có dữ liệu")
            return pd.DataFrame()
        
        # Đảm bảo các cột cần thiết tồn tại
        df_result = ensure_required_columns(df_result)
        
        # Đảm bảo kiểu dữ liệu phù hợp
        df_result = ensure_proper_dtypes(df_result)
        
        status_label.configure(text=f"Đã đọc xong file Excel")
        return df_result
        
    except Exception as e:
        status_label.configure(text=f"Lỗi: {str(e)}")
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
            status_label.configure(text=f"Đã chuyển sang kênh cập nhật: {new_channel}")
            
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
root.configure(menu=menubar)

# Menu Cài đặt
settings_menu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="Cài đặt", menu=settings_menu)
settings_menu.add_command(label="Tùy chỉnh phím tắt", command=customize_shortcuts)
settings_menu.add_command(label="Tùy chỉnh mã đề", command=customize_exam_codes)
settings_menu.add_command(label="Tùy chỉnh tên cột", command=customize_columns)
settings_menu.add_command(label="Đọc Excel", command=customize_excel_reading)
settings_menu.add_command(label="Bảo mật", command=customize_security)
settings_menu.add_separator()
settings_menu.add_command(label="Chọn kênh cập nhật", command=choose_update_channel)
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
status_label.configure(text="✅ Ứng dụng đã sẵn sàng! Đang kiểm tra cập nhật...")

# Bắt đầu cập nhật thống kê tự động
root.after(1000, auto_update_stats)

# Chạy ứng dụng
root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()
    

