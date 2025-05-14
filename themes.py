"""
Module để quản lý các theme của ứng dụng (light/dark mode)
"""

# Theme mặc định cho chế độ sáng (light mode)
LIGHT_THEME = {
    "primary": "#1976D2",          # Xanh dương đẹp nhưng ít chói hơn
    "secondary": "#7CB342",         # Xanh lá đậm hơn, dễ nhìn hơn
    "warning": "#FB8C00",           # Cam đẹp nhưng ít chói hơn
    "error": "#E53935",             # Đỏ đậm, dễ nhìn
    "success": "#43A047",           # Xanh lá đậm hơn, dễ nhìn
    "text": "#263238",              # Đen nhưng nhẹ hơn, dễ đọc
    "text_secondary": "#546E7A",    # Xám đậm, dễ đọc
    "background": "#ECEFF1",        # Nền trắng xám nhẹ, dễ chịu cho mắt
    "card": "#FFFFFF",              # Nền card màu trắng
    "border": "#CFD8DC",            # Viền nhẹ, không quá tương phản
    "hover": "#E1F5FE",             # Màu khi hover - xanh nhạt
    "active": "#0D47A1",            # Màu khi active - xanh đậm
    "selected": "#BBDEFB",          # Màu khi select - xanh nhạt
    "treeview_bg": "#FFFFFF",       # Nền treeview
    "treeview_selected": "#1976D2", # Màu khi chọn dòng trong treeview
    "disabled": "#B0BEC5"           # Màu cho các widget bị vô hiệu hóa
}

# Theme cho chế độ tối (dark mode)
DARK_THEME = {
    "primary": "#2196F3",           # Xanh dương sáng hơn trên nền tối
    "secondary": "#8BC34A",         # Xanh lá sáng hơn trên nền tối
    "warning": "#FFA726",           # Cam sáng hơn trên nền tối
    "error": "#EF5350",             # Đỏ sáng hơn trên nền tối
    "success": "#66BB6A",           # Xanh lá sáng hơn trên nền tối
    "text": "#FFFFFF",              # Trắng hoàn toàn để tăng độ tương phản
    "text_secondary": "#E0E0E0",    # Xám sáng hơn, gần với màu trắng
    "background": "#0A0A0A",        # Nền tối đen hơn - điều chỉnh để tăng tương phản
    "card": "#1A1A1A",              # Card màu tối hơn nền một chút, gần với màu đen
    "border": "#333333",            # Viền tối, dễ nhìn hơn trên nền tối
    "hover": "#1E88E5",             # Màu khi hover - xanh đậm, rõ hơn
    "active": "#90CAF9",            # Màu khi active - xanh sáng, dễ nhận biết
    "selected": "#1976D2",          # Màu khi select - xanh đậm
    "treeview_bg": "#1A1A1A",       # Nền treeview tối hơn
    "treeview_selected": "#2196F3", # Màu khi chọn dòng trong treeview - tăng độ tương phản
    "disabled": "#757575"           # Màu cho các widget bị vô hiệu hóa, sáng hơn
}

def get_theme(dark_mode=False):
    """
    Trả về theme phù hợp tùy theo chế độ tối/sáng
    
    Args:
        dark_mode (bool): True nếu muốn sử dụng dark mode
        
    Returns:
        dict: Theme được chọn
    """
    return DARK_THEME if dark_mode else LIGHT_THEME

def apply_theme_to_config(config, dark_mode=False):
    """
    Áp dụng theme vào cấu hình
    
    Args:
        config (dict): Cấu hình hiện tại
        dark_mode (bool): True nếu muốn sử dụng dark mode
        
    Returns:
        dict: Cấu hình đã được cập nhật
    """
    theme = get_theme(dark_mode)
    config['ui']['theme'] = theme
    config['ui']['dark_mode'] = dark_mode
    return config 