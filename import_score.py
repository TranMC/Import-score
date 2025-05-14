import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import pandas as pd
from pandas import CategoricalDtype
import os
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

# Tạo class ToolTip để hiển thị gợi ý khi di chuột qua các nút
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)
        
    def show_tooltip(self, event=None):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        
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

# Cấu hình cơ bản
config = {
    'columns': {
        'name': 'Tên Học Sinh',  # Default column name
        'exam_code': 'Mã Đề',    # Default exam code column 
        'score': 'Điểm'          # Default score column
    },    'max_questions': 40,
    'score_per_question': 0.25,
    'version': '1.7.1',
    'exam_codes': ['701', '702', '703', '704'],  # Thêm danh sách mã đề
    'shortcuts': {
        'search': '<Control-f>',
        'direct_score': '<Control-g>',
        'undo': '<Control-z>'    },
    'security': {
        'encrypt_backups': False,
        'encrypt_sensitive_data': False,
        'password_protect_app': False,
        'auto_lock_timeout_minutes': 0
    },
    'changelog': {
        '1.7.1':[
            'Sửa lại một số lỗi trước đó'
        ],
        '1.6.0': [
            'Thêm tính năng xuất báo cáo PDF',
            'Thêm biểu đồ phân phối điểm số',
            'Thêm thống kê tỷ lệ đạt/không đạt',
            'Cải thiện giao diện thống kê',
            'Sửa các lỗi liên quan đến pandas mới'
        ],
        '1.5.5': [
            'Thêm tính năng hiển thị điểm cao nhất và thấp nhất',
            'Tự động chọn học sinh khi chỉ có một kết quả tìm kiếm',
            'Cải thiện trải nghiệm tìm kiếm với delay thông minh',
            'Thêm thống kê số học sinh có điểm/tổng số học sinh'
        ],
        '1.5.4': [
            'Thêm tính năng hoàn tác (Ctrl+Z)',
            'Thêm tính năng tùy chỉnh phím tắt',
            'Thêm tính năng tùy chỉnh mã đề',
            'Cải thiện UX cho tìm kiếm học sinh',
            'Thêm tính năng lưu cấu hình',
            'Chỉnh lại vài lỗi đã biết (không tự động tìm kiếm, lỗi khi đổi mục tìm)'
        ],
        '1.4.0': [
            'Cải thiện giao diện người dùng',
            'Thêm cấu hình số câu hỏi và điểm mỗi câu',
            'Tối ưu hóa kích thước cửa sổ',
            'Thêm danh sách mã đề có sẵn',
            'Tự động lưu sau mỗi thay đổi'
        ],
        '1.3.0': [
            'Thêm tính năng hoàn tác (Ctrl+Z)',
            'Cải thiện hiệu suất tìm kiếm',
            'Sửa lỗi nhập điểm'
        ],
        '1.2.0': [
            'Thêm tính năng import từ Excel khác',
            'Cải thiện hiển thị danh sách học sinh',
            'Thêm validation dữ liệu nhập vào'
        ],
        '1.1.0': [
            'Thêm tính năng tìm kiếm học sinh',
            'Thêm tính năng tự động tính điểm',
            'Cải thiện UI/UX'
        ],
        '1.0.0': [
            'Phát hành phiên bản đầu tiên',
            'Tính năng cơ bản: nhập điểm, quản lý học sinh',
            'Lưu trữ dữ liệu bằng Excel'
        ]
    },
    'ui': {
        'font_family': 'Segoe UI',  # Font chữ phổ biến, dễ đọc
        'font_size': {
            'normal': 11,           # Cỡ chữ thường
            'heading': 12,          # Cỡ chữ tiêu đề
            'button': 11           # Cỡ chữ nút
        },
        'padding': {
            'frame': 10,           # Padding cho frame
            'widget': 5            # Padding cho widget
        },
        'min_width': {
            'button': 120,         # Độ rộng tối thiểu cho nút
            'entry': 150,          # Độ rộng tối thiểu cho ô nhập
            'combobox': 100        # Độ rộng tối thiểu cho combobox
        }
    }
}

def save_config():
    """Lưu cấu hình ra file JSON"""
    try:
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_config.json")
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"Lỗi khi lưu cấu hình: {str(e)}")

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
search_timer_id = None  # Thêm biến để theo dõi timer
search_index = None    # Thêm biến để lưu search index
last_activity_time = None  # Thêm biến để theo dõi thời gian hoạt động cuối cùng
lock_window = None  # Thêm biến để theo dõi cửa sổ khóa
lock_time = None  # Thêm biến để theo dõi thời gian khóa

def optimize_memory_usage(df):
    """Tối ưu hóa sử dụng bộ nhớ cho DataFrame lớn"""
    import gc
    import sys
    
    # Nếu DataFrame không có thì không cần tối ưu
    if df is None:
        return pd.DataFrame()  # Trả về DataFrame rỗng thay vì None
    
    # Nếu DataFrame rỗng, trả về ngay
    if df.empty:
        return df
        
    # Ghi nhận thông tin bộ nhớ trước khi tối ưu
    initial_memory = df.memory_usage(deep=True).sum()
    initial_memory_mb = initial_memory / (1024 * 1024)
    
    # Tối ưu các kiểu dữ liệu để giảm bộ nhớ
    for col in df.columns:
        try:
            # Chuyển cột số nguyên về kiểu nhỏ hơn nếu có thể
            if pd.api.types.is_integer_dtype(df[col]):
                # Bỏ qua các cột có giá trị null
                if df[col].isna().any():
                    continue
                    
                # Không cần kiểm tra min/max cho các cột nhỏ hơn
                col_min = df[col].min()
                col_max = df[col].max()
                
                if col_min >= -128 and col_max <= 127:
                    df[col] = df[col].astype('int8')
                elif col_min >= -32768 and col_max <= 32767:
                    df[col] = df[col].astype('int16')
                elif col_min >= -2147483648 and col_max <= 2147483647:
                    df[col] = df[col].astype('int32')
            
            # Chuyển cột số thực về float32 nếu độ chính xác không quan trọng
            elif pd.api.types.is_float_dtype(df[col]):
                df[col] = df[col].astype('float32')
                
            # Tối ưu cột chuỗi bằng category nếu số giá trị khác nhau thấp
            elif pd.api.types.is_object_dtype(df[col]):
                # Không chuyển các cột quan trọng sang category để tránh vấn đề khi gán giá trị
                if col not in ['Điểm', 'Mã đề']:
                    num_unique = df[col].nunique()
                    num_total = len(df)
                    # Nếu số giá trị khác nhau nhỏ hơn 50% số hàng, dùng category
                    if num_unique < num_total * 0.5:
                        df[col] = df[col].astype('category')
                    # Thử chuyển về chuỗi UTF-8 (tiết kiệm hơn object)
                    elif df[col].map(lambda x: isinstance(x, str)).all():
                        df[col] = df[col].astype('string')
                # Nếu là cột 'Điểm', chỉ chuyển sang float nếu có thể
                elif col == 'Điểm':
                    df[col] = pd.to_numeric(df[col], errors='coerce')
        except Exception as e:
            print(f"Lỗi khi tối ưu cột {col}: {str(e)}")
    
    # Tính toán bộ nhớ đã tiết kiệm
    final_memory = df.memory_usage(deep=True).sum()
    final_memory_mb = final_memory / (1024 * 1024)
    memory_saved_mb = initial_memory_mb - final_memory_mb
    memory_saved_percent = (1 - final_memory/initial_memory) * 100 if initial_memory > 0 else 0
    
    # Cưỡng chế thu gom rác để giải phóng bộ nhớ
    gc.collect()
    
    # Lưu thông tin về tiết kiệm bộ nhớ vào một biến thay vì gán trực tiếp cho DataFrame
    memory_optimization_info = {
        'initial_size_mb': initial_memory_mb,
        'optimized_size_mb': final_memory_mb,
        'saved_mb': memory_saved_mb,
        'saved_percent': memory_saved_percent
    }
    
    # In thông tin về bộ nhớ tiết kiệm được
    if memory_saved_mb > 0.1:  # Chỉ hiển thị nếu tiết kiệm được đáng kể
        print(f"Đã tối ưu bộ nhớ: {initial_memory_mb:.2f} MB → {final_memory_mb:.2f} MB (Tiết kiệm {memory_saved_percent:.1f}%)")
    
    # Luôn trả về DataFrame đã tối ưu
    return df

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

def monitor_memory_usage():
    """Giám sát và báo cáo việc sử dụng bộ nhớ của ứng dụng"""
    import gc
    import os
    import psutil
    
    try:
        # Cưỡng chế thu gom rác để có số liệu chính xác
        gc.collect()
        
        # Lấy thông tin bộ nhớ
        process = psutil.Process(os.getpid())
        memory_info = process.memory_info()
        
        # Tính toán sử dụng bộ nhớ
        memory_usage_mb = memory_info.rss / (1024 * 1024)
        
        # Thông tin về DataFrame
        df_memory = 0
        if df is not None:
            df_memory = df.memory_usage(deep=True).sum() / (1024 * 1024)
        
        memory_info = {
            "total_memory_mb": memory_usage_mb,
            "dataframe_memory_mb": df_memory,
            "other_memory_mb": memory_usage_mb - df_memory,
            "process_info": {
                "pid": os.getpid(),
                "cpu_percent": process.cpu_percent(interval=0.1),
                "memory_percent": process.memory_percent()
            }
        }
        
        return memory_info
        
    except ImportError:
        # Nếu không có psutil, chỉ báo cáo thông tin cơ bản
        memory_info = {
            "dataframe_memory_mb": df.memory_usage(deep=True).sum() / (1024 * 1024) if df is not None else 0,
        }
        return memory_info
    except Exception as e:
        print(f"Lỗi theo dõi bộ nhớ: {str(e)}")
        return {"error": str(e)}

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
    import gc
    from pandas.io.excel._openpyxl import OpenpyxlReader
    
    # Trạng thái tiến trình
    status_label.config(text=f"Đang lập kế hoạch đọc file lớn: {os.path.basename(file_path)}...", 
                       style="StatusWarning.TLabel")
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
                         style="StatusWarning.TLabel")
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
            # Tối ưu bộ nhớ ngay lập tức cho mỗi chunk
            chunk = optimize_memory_usage(chunk)
            chunks.append(chunk)
            
            total_rows += len(chunk)
            status_label.config(text=f"Đang đọc dữ liệu: {total_rows} dòng...", 
                             style="StatusWarning.TLabel")
            root.update()
            
            # Nếu đã đọc quá nhiều, hiển thị cảnh báo
            if total_rows > 10000:
                status_label.config(text=f"File lớn: đã đọc {total_rows} dòng...", 
                                 style="StatusCritical.TLabel")            # Ghép các chunk lại
        
        if chunks:
            status_label.config(text=f"Đang ghép {len(chunks)} chunk dữ liệu...", 
                         style="StatusWarning.TLabel")
            root.update()
        
            result = pd.concat(chunks, ignore_index=True)
        
            # Giải phóng bộ nhớ chunks
            chunks.clear()
            gc.collect()
        
            # Tối ưu hóa lần cuối
            result = optimize_memory_usage(result)
            
            # Đảm bảo kiểu dữ liệu phù hợp cho các cột quan trọng
            result = ensure_proper_dtypes(result)
        
            status_label.config(text=f"Đã đọc xong {total_rows} dòng dữ liệu", 
                         style="StatusSuccess.TLabel")
            return result
        else:
            return pd.DataFrame()
            
    except Exception as e:
        status_label.config(text=f"Lỗi khi đọc file: {str(e)}", 
                         style="StatusCritical.TLabel")
        messagebox.showerror("Lỗi", f"Không thể đọc file Excel: {str(e)}")
        print(f"Chi tiết lỗi: {traceback.format_exc()}")
        return pd.DataFrame()  # Trả về DataFrame rỗng thay vì None

def select_file():
    global df, file_path
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        # Hiển thị thông báo trạng thái khi bắt đầu đọc file
        status_label.config(text=f"Đang đọc file: {os.path.basename(file_path)}...", 
                          style="StatusWarning.TLabel")
        root.update()  # Cập nhật giao diện ngay lập tức để hiển thị trạng thái
        
        try:
            # Kiểm tra kích thước file
            file_size = os.path.getsize(file_path) / (1024 * 1024)  # Kích thước file (MB)
            is_large_file = file_size > config.get('memory_optimization', {}).get('large_file_threshold_mb', 5)
            
            if is_large_file:
                status_label.config(text=f"Đang phân tích file lớn ({file_size:.1f} MB)...", 
                                  style="StatusWarning.TLabel")
                root.update()
                
                # Kiểm tra xem có bật lazy loading hay không
                if config.get('memory_optimization', {}).get('lazy_loading', True):
                    chunk_size = config.get('memory_optimization', {}).get('chunk_size', 1000)
                    df = load_excel_lazily(file_path, chunk_size=chunk_size)
                else:
                    # Đọc theo cách thông thường nếu không sử dụng lazy loading
                    df = read_excel_normally(file_path)
            else:
                # Đọc file nhỏ theo cách thông thường
                df = read_excel_normally(file_path)
            
            if df is not None and not df.empty:
                # Hiển thị số lượng học sinh
                total_students = len(df)
                status_label.config(
                    text=f"Đã đọc xong: {os.path.basename(file_path)} ({total_students} học sinh)",
                    style="StatusSuccess.TLabel"
                )
                
                # Tối ưu hóa bộ nhớ nếu được cấu hình
                if config.get('memory_optimization', {}).get('auto_optimize', True):
                    df = optimize_memory_usage(df)
                
                # Đảm bảo các cột cần thiết tồn tại
                df = ensure_required_columns(df)
                
                # Cập nhật giao diện
                refresh_ui()
                
                # Báo cáo sử dụng bộ nhớ
                memory_info = monitor_memory_usage()
                if 'dataframe_memory_mb' in memory_info:
                    df_memory = memory_info['dataframe_memory_mb']
                    if df_memory > 100:  # Hiển thị cảnh báo nếu dữ liệu lớn
                        print(f"Cảnh báo: DataFrame đang sử dụng {df_memory:.1f} MB bộ nhớ")
            else:
                status_label.config(
                    text="Không có dữ liệu để hiển thị, vui lòng tải file Excel có dữ liệu",
                    style="StatusCritical.TLabel"
                )
                
        except Exception as e:
            error_message = str(e)
            status_label.config(
                text=f"Lỗi: {error_message[:50] + '...' if len(error_message) > 50 else error_message}",
                style="StatusCritical.TLabel"
            )
            messagebox.showerror("Lỗi", f"Không thể đọc file Excel:\n{error_message}")
            traceback.print_exc()  # In chi tiết lỗi ra console để debug

def read_excel_normally(file_path):
    """Đọc file Excel theo cách thông thường"""
    status_label.config(text=f"Đang đọc file Excel...", style="StatusWarning.TLabel")
    root.update()
    
    try:
        # Đọc 10 dòng đầu để phân tích header với engine openpyxl
        headers_df = pd.read_excel(file_path, nrows=10, engine='openpyxl')
        
        # Tìm hàng chứa header
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
            
        # Đọc file với header đúng và engine openpyxl
        df_result = pd.read_excel(file_path, header=header_row, engine='openpyxl')
        
        # Xử lý trường hợp DataFrame rỗng
        if df_result.empty:
            status_label.config(text="File Excel không có dữ liệu", style="StatusCritical.TLabel")
            return pd.DataFrame()  # Trả về DataFrame rỗng thay vì None
        
        # Đảm bảo các cột cần thiết tồn tại
        df_result = ensure_required_columns(df_result)
        
        # Đảm bảo kiểu dữ liệu phù hợp
        df_result = ensure_proper_dtypes(df_result)
        
        status_label.config(text=f"Đã đọc xong file Excel", style="StatusSuccess.TLabel")
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
        status_label.config(text="Đang tìm kiếm trong dữ liệu lớn...", style="StatusWarning.TLabel")
        root.update()  # Cập nhật giao diện trước khi thực hiện tìm kiếm
    
    # Clear existing items
    for item in tree.get_children():
        tree.delete(item)
        
    if df is None or df.empty:
        messagebox.showinfo("Thông báo", "Chưa có dữ liệu học sinh")
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
        messagebox.showerror("Lỗi", "Không tìm thấy cột tên học sinh trong file Excel")
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

    # Display results with improved performance
    if not ten_hoc_sinh:
        # Nếu số lượng học sinh lớn, giới hạn hiển thị ban đầu
        display_limit = 100 if len(df) > 100 else len(df)
        for _, row in df.head(display_limit).iterrows():
            values = get_display_values(row)
            if values:
                tree.insert('', 'end', values=values)
        
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
                    tree.insert('', 'end', values=values)
            
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
        status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusGood.TLabel")
                
    update_stats()

def add_student():
    """Thêm học sinh mới"""
    global df
    if df is None:
        messagebox.showerror("Lỗi", "Chưa chọn file Excel.")
        return

    ten_hoc_sinh = entry_student_name.get().strip()
    if not ten_hoc_sinh:
        messagebox.showinfo("Thông báo", "Vui lòng nhập tên học sinh để thêm.")
        return

    if ten_hoc_sinh in df[config['columns']['name']].values:
        messagebox.showinfo("Thông báo", "Học sinh đã tồn tại trong danh sách.")
    else:
        save_state()
        new_row = pd.DataFrame({config['columns']['name']: [ten_hoc_sinh], 'Điểm': [None]})
        df = pd.concat([df, new_row], ignore_index=True)
        save_excel()
        search_student()

def calculate_score(event=None):
    """Tính điểm từ số câu đúng"""
    global df
    if df is None:
        messagebox.showerror("Lỗi", "Chưa chọn file Excel.")
        return
 
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showerror("Lỗi", "Vui lòng chọn học sinh.")
        return

    selected = tree.item(selected_item[0])['values'][0]
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
                messagebox.showerror("Lỗi", "Mã đề phải là số")
                return
        # Nếu để trống thì giữ nguyên mã đề cũ
        else:
            ma_de = df.loc[df[config['columns']['name']] == selected, 'Mã đề'].iloc[0]
            
        if not (0 <= so_cau_dung <= config['max_questions']):
            messagebox.showerror("Lỗi", 
                               f"Số câu đúng phải từ 0 đến {config['max_questions']}.")
            return

        diem = round(so_cau_dung * config['score_per_question'], 2)
        
        if diem > 10:
            messagebox.showerror("Lỗi", "Điểm tính được vượt quá 10.")
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
        search_student()
        
        # Chỉ xóa nội dung ô số câu đúng
        entry_correct_count.delete(0, tk.END)
    except ValueError:
        messagebox.showerror("Lỗi", "Vui lòng nhập số câu đúng hợp lệ.")

def update_config(event=None):
    """Cập nhật cấu hình tính điểm"""
    try:
        max_q = int(entry_max_questions.get())
        
        if max_q <= 0:
            messagebox.showerror("Lỗi", "Số câu hỏi phải lớn hơn 0")
            return
            
        # Tự động tính điểm mỗi câu bằng cách chia 10 cho số câu hỏi
        score_per_q = round(10 / max_q, 2)
            
        config['max_questions'] = max_q
        config['score_per_question'] = score_per_q
        
        save_config()  # Lưu vào file
        
        # Cập nhật label thông tin
        score_info_label.config(
            text=f"(Mỗi câu = {score_per_q} điểm, tối đa {max_q} câu, tổng điểm = 10)"
        )
        messagebox.showinfo("Thành công", "Đã cập nhật cấu hình tính điểm")
        
    except ValueError:
        messagebox.showerror("Lỗi", "Vui lòng nhập số hợp lệ")

def focus_student_search(event=None):
    """Di chuyển con trỏ đến ô tìm kiếm học sinh"""
    entry_student_name.focus_set()
    entry_student_name.select_range(0, tk.END)

def calculate_score_direct(event=None):
    """Tính điểm trực tiếp từ điểm số nhập vào"""
    global df, entry_direct_score  # Thêm entry_direct_score vào global
    if df is None:
        messagebox.showerror("Lỗi", "Chưa chọn file Excel.")
        return

    selected_item = tree.selection()
    if not selected_item:
        messagebox.showerror("Lỗi", "Vui lòng chọn học sinh.")
        return

    selected = tree.item(selected_item[0])['values'][0]
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
                messagebox.showerror("Lỗi", "Mã đề phải là số")
                return
        # Nếu để trống thì giữ nguyên mã đề cũ
        else:
            ma_de = df.loc[df[config['columns']['name']] == selected, 'Mã đề'].iloc[0]
            
        if not (0 <= diem <= 10):
            messagebox.showerror("Lỗi", "Điểm phải từ 0 đến 10.")
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
        search_student()
        
        # Chỉ xóa nội dung ô nhập điểm
        entry_direct_score.delete(0, tk.END)
    except ValueError:
        messagebox.showerror("Lỗi", "Vui lòng nhập điểm hợp lệ.")

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
        messagebox.showinfo("Thành công", "Đã lưu cấu hình phím tắt")
    
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
        messagebox.showinfo("Thành công", "Đã lưu danh sách mã đề")
    
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
        messagebox.showinfo("Thành công", "Đã lưu tên cột mới")
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
        
        messagebox.showinfo("Thành công", "Đã lưu cấu hình bảo mật.")
    
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
        messagebox.showerror("Lỗi", "Chưa có dữ liệu để sao lưu.")
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
        
        # Tùy chọn backup tự động
        options_frame = ttk.Frame(confirm_window)
        options_frame.pack(fill="x", padx=20, pady=5)
        
        auto_backup_var = tk.BooleanVar(value=config.get('auto_backup', False))
        ttk.Checkbutton(options_frame, text="Bật sao lưu tự động", variable=auto_backup_var).pack(side="left")
        ttk.Label(options_frame, text="(Tự động sao lưu khi đóng chương trình)", 
                foreground="gray", font=(config['ui']['font_family'], 9)).pack(side="left", padx=5)
        
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
            
            # Lưu cấu hình
            config['auto_backup'] = auto_backup_var.get()
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
                "app_version": config['version'],
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
                messagebox.showinfo("Thành công", f"Đã sao lưu dữ liệu vào:\n{backup_path_display}")
                
            except Exception as enc_error:
                messagebox.showerror("Lỗi mã hóa", f"Không thể mã hóa dữ liệu: {str(enc_error)}")
                traceback.print_exc()
        
        ttk.Button(btn_frame, text="Sao lưu", command=confirm_backup, width=15).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Hủy", command=confirm_window.destroy, width=15).pack(side="right", padx=10)
        
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể sao lưu dữ liệu: {str(e)}")
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
                        messagebox.showerror("Lỗi", "Vui lòng nhập mật khẩu để giải mã")
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
                        messagebox.showerror("Lỗi giải mã", 
                                          f"Không thể giải mã dữ liệu. Mật khẩu có thể không đúng.\n\nLỗi: {str(decrypt_error)}")
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
                messagebox.showinfo("Thành công", "Đã phục hồi dữ liệu từ bản sao lưu.")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể phục hồi dữ liệu: {str(e)}")
                traceback.print_exc()
                confirm_window.destroy()
        
        ttk.Button(btn_frame, text="Xác nhận", command=confirm_restore, width=15).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Hủy", command=confirm_window.destroy, width=15).pack(side="right", padx=5)
        
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể đọc file backup: {str(e)}")
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
                "app_version": config['version'],
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
        
        for i, (_, row) in enumerate(df.head(display_limit).iterrows()):
            values = [
                row[name_col] if name_col in row else "",
                row[exam_code_col] if exam_code_col in row and pd.notna(row[exam_code_col]) else "",
                row[score_col] if score_col in row and pd.notna(row[score_col]) else ""
            ]
            tree.insert('', 'end', values=values)
            
        # Hiển thị thông báo nếu chỉ hiển thị một phần
        if len(df) > display_limit:
            tree.insert('', 'end', values=(f"--- Hiển thị {display_limit}/{len(df)} học sinh. Nhập từ khóa để tìm kiếm ---", "", ""))
        
        # Cập nhật status label nếu file_path tồn tại
        if 'file_path' in globals() and file_path:
            status_label.config(
                text=f"Dữ liệu đang hiển thị: {os.path.basename(file_path)} ({len(df)} học sinh)",
                style="StatusGood.TLabel"
            )
            
        # Cập nhật thống kê
        update_stats()
        
        # Cập nhật điểm cao nhất/thấp nhất
        update_score_extremes()
        
        # Kiểm tra và xác minh các cột dữ liệu cần thiết
        verify_required_columns(df)
    else:
        # Nếu không có dữ liệu, hiển thị thông báo
        status_label.config(text="Chưa tải file Excel", style="StatusWarning.TLabel")
        for item in tree.get_children():
            tree.delete(item)
        tree.insert('', 'end', values=("Không có dữ liệu để hiển thị. Vui lòng tải file Excel.", "", ""))

def verify_required_columns(dataframe):
    """Kiểm tra xem dataframe có các cột cần thiết không"""
    # Kiểm tra dataframe có tồn tại không
    if dataframe is None:
        messagebox.showerror("Lỗi", "Dữ liệu trống, vui lòng kiểm tra lại file Excel.")
        return False
        
    # Kiểm tra dataframe có rỗng không
    if dataframe.empty:
        messagebox.showerror("Lỗi", "Dữ liệu trống, vui lòng kiểm tra lại file Excel.")
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
        messagebox.showerror("Thiếu cột dữ liệu", 
                            f"File Excel thiếu các cột sau: {', '.join(missing_columns)}.\n\n"
                            f"Vui lòng kiểm tra lại tên cột hoặc cấu hình tên cột trong ứng dụng.")
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
    global entry_max_questions
    global score_info_label, entry_direct_score, stats_label
    global highest_score_label, lowest_score_label  # Thêm biến cho điểm cao/thấp    # Tạo style cho các widget
    style = ttk.Style()
    default_font = (config['ui']['font_family'], config['ui']['font_size']['normal'])
    heading_font = (config['ui']['font_family'], config['ui']['font_size']['heading'])
    button_font = (config['ui']['font_family'], config['ui']['font_size']['button'])
    
    style.configure('TLabel', font=default_font)
    style.configure('TButton', font=button_font)
    style.configure('TEntry', font=default_font)
    style.configure('Heading.TLabel', font=heading_font)
    style.configure('Treeview', font=default_font, rowheight=25)
    style.configure('Treeview.Heading', font=heading_font)
    
    # Định nghĩa thêm các styles mới
    style.configure('Status.TLabel', font=default_font, background="#f0f0f0")
    style.configure('StatusGood.TLabel', font=default_font, foreground='green')
    style.configure('StatusWarning.TLabel', font=default_font, foreground='orange')
    style.configure('StatusError.TLabel', font=default_font, foreground='red')
    
    # File Frame với layout cải tiến
    file_frame = ttk.LabelFrame(root, text="Quản lý File", padding=config['ui']['padding']['frame'])
    file_frame.pack(fill="x", padx=10, pady=5)
      # Tạo frame con để chứa status và thông tin phiên bản
    status_frame = ttk.Frame(file_frame)
    status_frame.pack(fill="x", pady=config['ui']['padding']['widget'])
    
    status_label = ttk.Label(status_frame, text="Chưa tải file Excel", 
                          style="StatusError.TLabel")
    status_label.pack(side="left", pady=config['ui']['padding']['widget'])
    
    # Thêm thông tin phiên bản vào cùng dòng với status
    version_info = ttk.Label(status_frame, text=f"v{config['version']}", 
                           font=(config['ui']['font_family'], 9),
                           foreground="gray")
    version_info.pack(side="right", pady=config['ui']['padding']['widget'])
    
    # Thêm frame chứa các nút để căn chỉnh tốt hơn
    buttons_frame = ttk.Frame(file_frame)
    buttons_frame.pack(fill="x", pady=config['ui']['padding']['widget'])
    
    # Nút chọn file Excel với tooltip
    open_button = ttk.Button(buttons_frame, text="Chọn File Excel", 
                           command=select_file,
                           width=config['ui']['min_width']['button'])
    open_button.pack(side="left", padx=config['ui']['padding']['widget'])
    ToolTip(open_button, "Mở file Excel chứa danh sách học sinh và điểm số")

    # Nút sao lưu dữ liệu với tooltip
    backup_button = ttk.Button(buttons_frame, text="Sao lưu dữ liệu", 
                             command=backup_data,
                             width=config['ui']['min_width']['button'])
    backup_button.pack(side="left", padx=config['ui']['padding']['widget'])
    ToolTip(backup_button, "Tạo bản sao lưu dữ liệu hiện tại (Khuyến nghị trước khi thay đổi lớn)")

    # Nút phục hồi dữ liệu với tooltip
    restore_button = ttk.Button(buttons_frame, text="Phục hồi dữ liệu", 
                              command=restore_backup,
                              width=config['ui']['min_width']['button'])
    restore_button.pack(side="left", padx=config['ui']['padding']['widget'])
    ToolTip(restore_button, "Khôi phục dữ liệu từ một bản sao lưu trước đó")
      # Nút xuất báo cáo vào file frame
    report_button = ttk.Button(buttons_frame, text="Xuất báo cáo PDF", 
                             command=generate_report,
                             width=config['ui']['min_width']['button'])
    report_button.pack(side="left", padx=config['ui']['padding']['widget'])
    ToolTip(report_button, "Tạo báo cáo PDF với thống kê và biểu đồ phân phối điểm")
    
    # Thêm nút làm mới (refresh) dữ liệu
    def refresh_data():
        """Làm mới dữ liệu từ file Excel hiện tại"""
        global df, file_path
        if file_path and os.path.exists(file_path):
            try:
                # Đọc lại file Excel và cập nhật giao diện
                df = pd.read_excel(file_path)
                search_student()
                update_stats()
                status_label.config(text=f"Đã làm mới dữ liệu từ: {os.path.basename(file_path)}", 
                                  style="StatusGood.TLabel")
                messagebox.showinfo("Thành công", "Đã làm mới dữ liệu từ file Excel")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể đọc file Excel: {str(e)}")
        else:
            messagebox.showinfo("Thông báo", "Chưa chọn file Excel hoặc file không tồn tại")
    
    refresh_button = ttk.Button(buttons_frame, text="Làm mới dữ liệu", 
                              command=refresh_data,
                              width=config['ui']['min_width']['button'])
    refresh_button.pack(side="left", padx=config['ui']['padding']['widget'])
    ToolTip(refresh_button, "Đọc lại dữ liệu từ file Excel hiện tại")

    # Thêm nút kiểm tra cập nhật
    update_button = ttk.Button(buttons_frame, text="Kiểm tra cập nhật", 
                             command=lambda: check_for_updates(True),
                             width=config['ui']['min_width']['button'])
    update_button.pack(side="left", padx=config['ui']['padding']['widget'])
    ToolTip(update_button, "Kiểm tra phiên bản mới trên GitLab")

    # Config Frame
    config_frame = ttk.LabelFrame(root, text="Cấu hình tính điểm", 
                                padding=config['ui']['padding']['frame'])
    config_frame.pack(fill="x", padx=10, pady=5)

    # Số câu hỏi
    ttk.Label(config_frame, text="Số câu hỏi:", 
             font=default_font).pack(side="left", padx=config['ui']['padding']['widget'])
    entry_max_questions = ttk.Entry(config_frame, width=10, font=default_font)
    entry_max_questions.insert(0, str(config['max_questions']))
    entry_max_questions.pack(side="left", padx=config['ui']['padding']['widget'])
    
    # Thêm chú thích về cách tính điểm
    ttk.Label(config_frame, text="(Điểm mỗi câu = 10/tổng số câu)", 
             font=(config['ui']['font_family'], 9),
             foreground="gray").pack(side="left", padx=5)

    ttk.Button(config_frame, text="Cập nhật", 
              command=update_config,
              width=15).pack(side="left", padx=config['ui']['padding']['widget'])

    score_info_label = ttk.Label(config_frame, 
        text=f"(Mỗi câu = {config['score_per_question']} điểm, tối đa {config['max_questions']} câu, tổng điểm = 10)",
        font=default_font)
    score_info_label.pack(side="left", padx=20)

    # Search Frame
    search_frame = ttk.LabelFrame(root, text="Tìm kiếm học sinh", 
                                padding=config['ui']['padding']['frame'])
    search_frame.pack(fill="x", padx=10, pady=5)

    # Tạo frame con để chứa label, entry và chú thích
    search_input_frame = ttk.Frame(search_frame)
    search_input_frame.pack(side="left", padx=config['ui']['padding']['widget'])

    ttk.Label(search_input_frame, text="Tên học sinh:", 
             font=default_font).pack(side="left", padx=config['ui']['padding']['widget'])
    entry_student_name = ttk.Entry(search_input_frame, width=30, font=default_font)
    entry_student_name.pack(side="left", padx=config['ui']['padding']['widget'])
    # Thêm sự kiện <KeyRelease> cho ô nhập
    entry_student_name.bind("<KeyRelease>", delayed_search)  
    ttk.Label(search_input_frame, text="(Ctrl+F để tìm nhanh)", 
             font=(config['ui']['font_family'], 9),
             foreground="gray").pack(side="left", padx=5)

    ttk.Button(search_frame, text="Thêm học sinh", 
              command=add_student,
              width=15).pack(side="left", padx=config['ui']['padding']['widget'])

    # Thêm frame thống kê riêng
    stats_frame = ttk.LabelFrame(root, text="Thống kê lớp học", padding=config['ui']['padding']['frame'])
    stats_frame.pack(fill="x", padx=10, pady=5)

    # Tạo grid layout cho các thông tin
    stats_label = ttk.Label(stats_frame, text="0/0 học sinh có điểm", font=default_font)
    stats_label.grid(row=0, column=0, padx=20, sticky="w")

    highest_score_label = ttk.Label(stats_frame, text="Cao nhất: N/A", font=default_font)
    highest_score_label.grid(row=0, column=1, padx=20, sticky="w")

    lowest_score_label = ttk.Label(stats_frame, text="Thấp nhất: N/A", font=default_font)
    lowest_score_label.grid(row=0, column=2, padx=20, sticky="w")

    # Result Frame với Treeview
    result_frame = ttk.LabelFrame(root, text="Danh sách học sinh", 
                                padding=config['ui']['padding']['frame'])
    result_frame.pack(fill="both", expand=True, padx=10, pady=5)

    # Treeview với font size mới
    columns = ('name', 'exam_code', 'score')
    tree = ttk.Treeview(result_frame, columns=columns, show='headings', 
                       style='Treeview')
    
    tree.heading('name', text=config['columns']['name'])
    tree.heading('exam_code', text=config['columns']['exam_code'])
    tree.heading('score', text=config['columns']['score'])
    
    tree.column('name', width=350)  # Giảm độ rộng cột tên
    tree.column('exam_code', width=100, anchor='center')
    tree.column('score', width=100, anchor='center')

    vsb = ttk.Scrollbar(result_frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(result_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    
    tree.grid(column=0, row=0, sticky='nsew')
    vsb.grid(column=1, row=0, sticky='ns')
    hsb.grid(column=0, row=1, sticky='ew')
    
    result_frame.grid_columnconfigure(0, weight=1)
    result_frame.grid_rowconfigure(0, weight=1)

    # Score Frame
    score_frame = ttk.LabelFrame(root, text="Nhập điểm", 
                               padding=config['ui']['padding']['frame'])
    score_frame.pack(fill="x", padx=10, pady=5)

    # Mã đề với chú thích
    ma_de_frame = ttk.Frame(score_frame)
    ma_de_frame.pack(side="left", padx=config['ui']['padding']['widget'])
    
    ttk.Label(ma_de_frame, text="Mã đề:", 
             font=default_font).pack(side="left", padx=config['ui']['padding']['widget'])
    entry_exam_code = ttk.Combobox(ma_de_frame, 
                                 width=10, 
                                 values=config['exam_codes'],
                                 font=default_font)
    entry_exam_code.pack(side="left", padx=config['ui']['padding']['widget'])
    ttk.Label(ma_de_frame, text="(nhập x để xóa)", 
             font=(config['ui']['font_family'], 9),
             foreground="gray").pack(side="left")

    # Frame cho nhập điểm qua số câu đúng
    correct_frame = ttk.LabelFrame(score_frame, text="Nhập số câu đúng")
    correct_frame.pack(side="left", padx=10)

    ttk.Label(correct_frame, text="Số câu đúng:", font=default_font).pack(
        side="left", padx=config['ui']['padding']['widget'])
    entry_correct_count = ttk.Entry(correct_frame, width=10, font=default_font)
    entry_correct_count.pack(side="left", padx=config['ui']['padding']['widget'])
    entry_correct_count.bind("<Return>", calculate_score)
    # Thêm thông báo phím tắt
    ttk.Label(correct_frame, text="(Ctrl+D)", font=(config['ui']['font_family'], 9),
            foreground="gray").pack(side="left", padx=5)


    ttk.Button(correct_frame, text="Tính điểm", 
              command=calculate_score,
              width=15).pack(side="left", padx=config['ui']['padding']['widget'])

    # Frame cho nhập điểm trực tiếp với chú thích
    direct_frame = ttk.LabelFrame(score_frame, text="Nhập điểm trực tiếp")
    direct_frame.pack(side="left", padx=10)

    ttk.Label(direct_frame, text="Điểm số:", 
             font=default_font).pack(side="left", padx=config['ui']['padding']['widget'])
    entry_direct_score = ttk.Entry(direct_frame, width=10, font=default_font)
    entry_direct_score.pack(side="left", padx=config['ui']['padding']['widget'])
    entry_direct_score.bind("<Return>", calculate_score_direct)
    ttk.Label(direct_frame, text="(Ctrl+G)", 
             font=(config['ui']['font_family'], 9),
             foreground="gray").pack(side="left")

    # Điều chỉnh padding và khoảng cách
    for frame in [file_frame, config_frame, search_frame, stats_frame, result_frame, score_frame]:
        frame.pack(fill="x", padx=5, pady=2)  # Giảm padding

    # Status bar với version
    status_frame = ttk.Frame(root)
    status_frame.pack(side='bottom', fill='x')
    
    # Version label
    version_label = ttk.Label(status_frame, 
                            text=f"v{config['version']}", 
                            anchor='e')
    version_label.pack(side='right', padx=5, pady=2)

    # Thêm menu
    menubar = tk.Menu(root)
    root.config(menu=menubar)
    
    settings_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Cài đặt", menu=settings_menu)
    settings_menu.add_command(label="Tùy chỉnh phím tắt", command=customize_shortcuts)
    settings_menu.add_command(label="Tùy chỉnh mã đề", command=customize_exam_codes)
    settings_menu.add_command(label="Tùy chỉnh tên cột", command=customize_columns)  # Add new menu item
    settings_menu.add_command(label="Tùy chỉnh bảo mật", command=customize_security)  # Add new menu item
    settings_menu.add_separator()
    settings_menu.add_command(label="Kiểm tra cập nhật", command=lambda: check_for_updates(True))
    settings_menu.add_command(label="About", command=show_about)
    
    report_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Báo cáo", menu=report_menu)
    report_menu.add_command(label="Tạo báo cáo PDF", command=generate_report)
    report_menu.add_command(label="Hiển thị biểu đồ phân phối điểm", command=show_score_distribution)

def show_about():
    """Hiển thị thông tin về phần mềm"""
    about_text = f"""Phần mềm Quản lý Điểm Học Sinh
Version: {config['version']}

Changelog v{config['version']}:
""" + "\n".join(f"• {change}" for change in config['changelog'][config['version']]) + """

Xem thêm thông tin và cập nhật mới nhất tại GitLab.
    """
    
    about_window = tk.Toplevel(root)
    about_window.title("Thông tin")
    about_window.geometry("500x400")
    about_window.transient(root)
    about_window.grab_set()
    
    # Thêm scrollable text widget
    text_frame = ttk.Frame(about_window)
    text_frame.pack(fill="both", expand=True, padx=15, pady=15)
    
    about_text_widget = tk.Text(text_frame, wrap="word", font=(config['ui']['font_family'], 11))
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
              command=lambda: check_for_updates(True)).pack(side="left", padx=5)
    
    ttk.Button(button_frame, text="Đóng", 
              command=about_window.destroy).pack(side="right", padx=5)

def check_for_updates(show_notification=True):
    """
    Kiểm tra phiên bản mới trên GitHub
    
    Args:
        show_notification (bool): Có hiển thị thông báo khi không có phiên bản mới không
    
    Returns:
        bool: True nếu có phiên bản mới, False nếu không
    """
    try:
        # URL của GitHub API để kiểm tra phiên bản mới nhất
        github_api_url = "https://api.github.com/repos/TranMC/Import-score/releases/latest"
        
        # Hiển thị thông báo đang kiểm tra
        if show_notification:
            status_label.config(text="Đang kiểm tra phiên bản mới...", style="StatusWarning.TLabel")
            root.update()
        
        # Gọi API để lấy thông tin phiên bản mới nhất
        response = requests.get(github_api_url, timeout=5)
        
        # Kiểm tra kết quả
        if response.status_code == 200:
            release_info = response.json()
            
            # Lấy phiên bản mới nhất
            latest_version = release_info.get('tag_name', '').lstrip('v')
            current_version = config.get('version', '0.0.0')
            
            # So sánh phiên bản
            if latest_version > current_version:
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
                        # Mở trang download
                        download_url = ""
                        for asset in release_info.get('assets', []):
                            if asset.get('name', '').endswith('.exe'):
                                download_url = asset.get('browser_download_url', '')
                                break
                                
                        if download_url:
                            import webbrowser
                            webbrowser.open(download_url)
                        else:
                            webbrowser.open(release_info.get('html_url', ''))
                
                if 'file_path' in globals() and file_path:
                    status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusGood.TLabel")
                return True
            else:
                # Không có phiên bản mới
                if show_notification:
                    if 'file_path' in globals() and file_path:
                        status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusGood.TLabel")
                    messagebox.showinfo("Cập nhật", f"Bạn đang sử dụng phiên bản mới nhất ({current_version}).")
                return False
        else:
            # Lỗi kết nối
            if show_notification:
                if 'file_path' in globals() and file_path:
                    status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusGood.TLabel")
                messagebox.showwarning("Lỗi kết nối", f"Không thể kết nối đến máy chủ GitHub. Mã lỗi: {response.status_code}")
            return False
    
    except requests.RequestException as e:
        # Xử lý lỗi request
        if show_notification:
            if 'file_path' in globals() and file_path:
                status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusGood.TLabel")
            messagebox.showwarning("Lỗi kết nối", f"Không thể kiểm tra phiên bản mới: {str(e)}")
        return False
    
    except Exception as e:
        # Xử lý các lỗi khác
        if show_notification:
            if 'file_path' in globals() and file_path:
                status_label.config(text=f"Đã tải file: {os.path.basename(file_path)}", style="StatusGood.TLabel")
            messagebox.showwarning("Lỗi", f"Đã xảy ra lỗi khi kiểm tra phiên bản mới: {str(e)}")
        return False

def check_updates_async():
    """Kiểm tra cập nhật trong luồng riêng biệt để không làm đóng băng giao diện"""
    threading.Thread(target=check_for_updates, daemon=True).start()

def update_stats():
    """Cập nhật các thống kê cơ bản"""
    if df is None or df.empty:
        stats_label.config(text="Không có dữ liệu để hiển thị thống kê")
        return
        
    total_students = len(df)
    score_column = config['columns']['score']
    
    # Đếm số học sinh có điểm
    students_with_scores = df[score_column].notna().sum()
    
    # Hiển thị thông tin
    stats_text = f"Tổng số học sinh: {total_students} | Đã có điểm: {students_with_scores} ({students_with_scores/total_students:.1%})"
    stats_label.config(text=stats_text)

def update_score_extremes():
    """Cập nhật điểm cao nhất và thấp nhất"""
    if df is None or df.empty:
        highest_score_label.config(text="Điểm cao nhất: N/A")
        lowest_score_label.config(text="Điểm thấp nhất: N/A")
        return
        
    score_column = config['columns']['score']
    name_column = config['columns']['name']
    
    # Lọc các hàng có điểm không phải NaN
    df_with_scores = df[df[score_column].notna()]
    
    # Nếu không có ai có điểm
    if df_with_scores.empty:
        highest_score_label.config(text="Điểm cao nhất: N/A")
        lowest_score_label.config(text="Điểm thấp nhất: N/A")
        return
    
    # Tìm điểm cao nhất
    max_score = df_with_scores[score_column].max()
    max_students = df_with_scores[df_with_scores[score_column] == max_score][name_column].tolist()
    max_students_text = ", ".join(max_students[:3])
    if len(max_students) > 3:
        max_students_text += f" và {len(max_students) - 3} học sinh khác"
    
    # Tìm điểm thấp nhất
    min_score = df_with_scores[score_column].min()
    min_students = df_with_scores[df_with_scores[score_column] == min_score][name_column].tolist()
    min_students_text = ", ".join(min_students[:3])
    if len(min_students) > 3:
        min_students_text += f" và {len(min_students) - 3} học sinh khác"
    
    # Cập nhật giao diện
    highest_score_label.config(text=f"Điểm cao nhất: {max_score} ({max_students_text})")
    lowest_score_label.config(text=f"Điểm thấp nhất: {min_score} ({min_students_text})")

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
        messagebox.showwarning("Warning", "No data available for report generation")
        return
        
    if 'Điểm' not in df.columns:
        messagebox.showwarning("Warning", "Score column not found in data")
        return
        
    # Lọc chỉ lấy học sinh có điểm
    scores_df = df[df['Điểm'].notna()].copy()
    
    if len(scores_df) == 0:
        messagebox.showwarning("Warning", "No students with scores found")
        return
    
    try:
        # Hỏi nơi lưu file báo cáo
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save PDF Report"
        )
        
        if not file_path:
            return  # Người dùng hủy
            
        # Hiển thị thông báo đang tạo báo cáo
        status_label.config(text="Creating PDF report...", style="StatusWarning.TLabel")
        root.update()
        
        # Tạo tệp PDF
        with PdfPages(file_path) as pdf:
            # Trang 1: Thông tin chung
            plt.figure(figsize=(10, 12))
            plt.axis('off')
            
            # Tiêu đề báo cáo
            title_text = f"SCORE STATISTICS REPORT"
            subtitle_text = f"Created: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
            
            plt.text(0.5, 0.98, title_text, fontsize=16, ha='center', va='top', fontweight='bold')
            plt.text(0.5, 0.95, subtitle_text, fontsize=12, ha='center', va='top')
            
            # Thông tin cơ bản
            info_text = [
                f"Total students: {len(df)}",
                f"Students with scores: {len(scores_df)} ({len(scores_df)/len(df):.1%})",
                f"Average score: {scores_df['Điểm'].mean():.2f}",
                f"Highest score: {scores_df['Điểm'].max():.2f}",
                f"Lowest score: {scores_df['Điểm'].min():.2f}",
                f"Standard deviation: {scores_df['Điểm'].std():.2f}"
            ]
            
            plt.text(0.1, 0.9, "Overview:", fontsize=14, va='top', fontweight='bold')
            y_pos = 0.85
            for info in info_text:
                plt.text(0.1, y_pos, info, fontsize=12, va='top')
                y_pos -= 0.03
            
            # Biểu đồ phân phối điểm
            ax1 = plt.axes([0.1, 0.45, 0.8, 0.3])
            scores_df['Điểm'].hist(ax=ax1, bins=10, alpha=0.7, color='blue', edgecolor='black')
            ax1.set_title('Score Distribution')
            ax1.set_xlabel('Score')
            ax1.set_ylabel('Number of Students')
            ax1.grid(True, alpha=0.3)
            
            # Thêm đường trung bình
            mean_score = scores_df['Điểm'].mean()
            ax1.axvline(mean_score, color='red', linestyle='dashed', linewidth=1)
            ax1.text(mean_score + 0.1, ax1.get_ylim()[1]*0.9, f'Avg: {mean_score:.2f}', color='red')
            
            # Biểu đồ tròn tỷ lệ đạt/không đạt
            ax2 = plt.axes([0.1, 0.1, 0.35, 0.25])
            pass_threshold = 5.0
            passed = (scores_df['Điểm'] >= pass_threshold).sum()
            failed = len(scores_df) - passed
            
            labels = ['Pass (≥ 5.0)', 'Fail (< 5.0)']
            sizes = [passed, failed]
            colors = ['#66b3ff', '#ff9999']
            
            ax2.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)
            ax2.axis('equal')
            ax2.set_title('Pass/Fail Ratio')
            
            # Bảng học sinh điểm cao nhất
            ax3 = plt.axes([0.55, 0.05, 0.35, 0.3])
            ax3.axis('tight')
            ax3.axis('off')
            
            # Lấy 5 học sinh có điểm cao nhất
            name_col = config['columns']['name']
            top_students = scores_df.sort_values('Điểm', ascending=False).head(5)
            top_data = [[student, score] for student, score in 
                       zip(top_students[name_col], top_students['Điểm'])]
            
            top_table = ax3.table(cellText=top_data, colLabels=[name_col, 'Score'],
                              loc='center', cellLoc='center')
            top_table.auto_set_font_size(False)
            top_table.set_fontsize(10)
            top_table.scale(1, 1.5)
            
            ax3.set_title('Top Students')
            
            pdf.savefig()
            plt.close()
            
            # Trang 2: Danh sách điểm
            plt.figure(figsize=(10, 12))
            plt.axis('off')
            
            plt.text(0.5, 0.98, "STUDENT SCORE LIST", fontsize=16, ha='center', fontweight='bold')
            
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
                col_labels = [name_col, exam_code_col, 'Score']
            else:
                sorted_df = scores_df.sort_values(name_col)
                table_data = [[student, score] for student, score in 
                             zip(sorted_df[name_col][:page_size], 
                                 sorted_df['Điểm'][:page_size])]
                col_labels = [name_col, 'Score']
            
            table = ax.table(cellText=table_data, colLabels=col_labels,
                          loc='center', cellLoc='center')
            table.auto_set_font_size(False)
            table.set_fontsize(9)
            table.scale(1, 1.2)
            
            pdf.savefig()
            plt.close()
            
            # Tạo thêm trang nếu số học sinh > page_size
            if len(scores_df) > page_size:
                pages_needed = (len(scores_df) - 1) // page_size + 1
                for page in range(1, pages_needed):
                    plt.figure(figsize=(10, 12))
                    plt.axis('off')
                    
                    plt.text(0.5, 0.98, f"STUDENT SCORE LIST (page {page+1}/{pages_needed})", 
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
        
        status_label.config(text=f"PDF report created: {os.path.basename(file_path)}", 
                          style="StatusSuccess.TLabel")
        messagebox.showinfo("Success", f"PDF report created at:\n{file_path}")
        
    except Exception as e:
        status_label.config(text=f"Error creating report: {str(e)}", 
                          style="StatusCritical.TLabel")
        messagebox.showerror("Error", f"Cannot create PDF report: {str(e)}")
        traceback.print_exc()

def show_score_distribution():
    """Hiển thị biểu đồ phân phối điểm số"""
    global df
    if df is None or df.empty:
        messagebox.showinfo("Thông báo", "Không có dữ liệu để hiển thị thống kê")
        return
        
    if 'Điểm' not in df.columns:
        messagebox.showinfo("Thông báo", "Không có cột điểm trong dữ liệu")
        return
    
    # Lọc chỉ lấy học sinh có điểm
    scores = df[df['Điểm'].notna()]['Điểm']
    
    if len(scores) == 0:
        messagebox.showinfo("Thông báo", "Chưa có học sinh nào có điểm")
        return
    
    try:
        # Tạo cửa sổ mới
        chart_window = tk.Toplevel(root)
        chart_window.title("Biểu đồ phân phối điểm số")
        chart_window.geometry("800x600")
        chart_window.transient(root)
        
        
        # Tạo frame chứa biểu đồ
        chart_frame = ttk.Frame(chart_window)
        chart_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Tạo figure và subplot
        fig = Figure(figsize=(8, 6), dpi=100)
        
        # Biểu đồ histogram
        ax1 = fig.add_subplot(211)
        ax1.hist(scores, bins=10, alpha=0.7, color='blue', edgecolor='black')
        ax1.set_title('Phân phối điểm số')
        ax1.set_xlabel('Điểm')
        ax1.set_ylabel('Số học sinh')
        ax1.grid(True, alpha=0.3)
        
        # Thêm đường trung bình
        mean_score = scores.mean()
        ax1.axvline(mean_score, color='red', linestyle='dashed', linewidth=1)
        ax1.text(mean_score + 0.1, ax1.get_ylim()[1]*0.9, f'TB: {mean_score:.2f}', color='red')
        
        # Biểu đồ tròn tỷ lệ đạt/không đạt
        ax2 = fig.add_subplot(212)
        pass_threshold = 5.0
        passed = (scores >= pass_threshold).sum()
        failed = len(scores) - passed
        
        labels = ['Đạt (≥ 5.0)', 'Chưa đạt (< 5.0)']
        sizes = [passed, failed]
        colors = ['#66b3ff', '#ff9999']
        
        ax2.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)
        ax2.axis('equal')
        ax2.set_title('Tỷ lệ học sinh đạt/không đạt')
        
        # Thêm thông tin thống kê
        stats_text = f"""
        Tổng số: {len(scores)} học sinh
        Điểm trung bình: {mean_score:.2f}
        Điểm cao nhất: {scores.max():.2f}
        Điểm thấp nhất: {scores.min():.2f}
        Độ lệch chuẩn: {scores.std():.2f}
        Đạt (≥5): {passed} ({passed/len(scores)*100:.1f}%)
        Chưa đạt (<5): {failed} ({failed/len(scores)*100:.1f}%)
        """
        
        fig.text(0.02, 0.02, stats_text, fontsize=9)
        
        # Điều chỉnh khoảng cách giữa các subplot
        fig.tight_layout(rect=[0, 0.1, 1, 0.95])
        
        # Tạo canvas để hiển thị biểu đồ
        canvas = FigureCanvasTkAgg(fig, chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        
        # Thêm thanh công cụ (toolbar)
        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        toolbar = NavigationToolbar2Tk(canvas, chart_frame)
        toolbar.update()
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
                fig.savefig(save_path, dpi=300, bbox_inches='tight')
                messagebox.showinfo("Thành công", f"Đã lưu biểu đồ tại:\n{save_path}")
        
        ttk.Button(button_frame, text="Lưu biểu đồ", command=save_chart).pack(side="right")
        ttk.Button(button_frame, text="Đóng", command=chart_window.destroy).pack(side="right", padx=5)
        
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể hiển thị biểu đồ: {str(e)}")
        traceback.print_exc()

def focus_correct_count(event=None):
    """Di chuyển con trỏ đến ô nhập số câu đúng"""
    entry_correct_count.focus_set()
    entry_correct_count.select_range(0, tk.END)

# Khởi tạo giao diện
create_ui()

# Thêm binding phím tắt
root.bind('<Control-z>', undo)
root.bind('<Control-f>', focus_student_search)
root.bind('<Control-g>', focus_direct_score)
root.bind('<Control-d>', focus_correct_count)

# Cập nhật thời gian hoạt động cuối cùng khi có sự kiện
root.bind_all("<Motion>", update_activity_time)
root.bind_all("<KeyPress>", update_activity_time)

# Kiểm tra tự động khóa ứng dụng
check_auto_lock()

# Kiểm tra cập nhật sau khi khởi động
root.after(2000, check_updates_async)

# Chạy ứng dụng
root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()
    