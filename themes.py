"""
Module để quản lý các theme của ứng dụng (light/dark mode)
"""

# Theme mặc định cho chế độ sáng (light mode) - Modern & Stylish
LIGHT_THEME = {
    "primary": "#4F46E5",           # Indigo mềm mại, chuyên nghiệp
    "primary_dark": "#3730A3",      # Indigo đậm khi hover/active
    "secondary": "#10B981",         # Emerald tươi sáng
    "accent": "#EC4899",            # Pink để tạo điểm nhấn
    "warning": "#F59E0B",           # Amber ấm
    "error": "#EF4444",             # Đỏ nhẹ
    "success": "#10B981",           # Xanh lá
    "info": "#3B82F6",              # Xanh dương sáng
    "text": "#1F2937",              # Xám đen (nhẹ hơn đen tuyền)
    "text_secondary": "#6B7280",    # Xám trung tính
    "background": "#F3F4F6",        # Nền xám rất nhạt (Gray 100)
    "card": "#FFFFFF",              # Nền card trắng tinh
    "card_hover": "#F9FAFB",        # Hover mảng trắng
    "border": "#E5E7EB",            # Viền siêu nhạt (Gray 200)
    "hover": "#E0E7FF",             # Hover tím siêu nhạt (Indigo 100)
    "active": "#4338CA",            # Nhấn đậm
    "selected": "#EEF2FF",          # Selected nhạt (Indigo 50)
    "treeview_bg": "#FFFFFF",       
    "treeview_selected": "#4F46E5",
    "treeview_hover": "#F9FAFB",    
    "disabled": "#9CA3AF",          # Disabled xám
    "gradient_start": "#4F46E5",    
    "gradient_end": "#7C3AED",      
    "shadow": "#00000010"           # Bóng mờ nhạt
}

# Theme cho chế độ tối (dark mode) - Modern & Stylish
DARK_THEME = {
    "primary": "#818CF8",           # Indigo sáng cho nền tối
    "primary_dark": "#6366F1",      # Indigo cơ bản
    "secondary": "#34D399",         # Emerald neon
    "accent": "#F472B6",            # Pink neon
    "warning": "#FBBF24",           
    "error": "#F87171",             
    "success": "#34D399",           
    "info": "#60A5FA",              
    "text": "#FFFFFF",              # Trắng tinh khiết cho nổi bật nhất
    "text_secondary": "#D1D5DB",    # Xám nhạt (Gray 300)
    "background": "#111827",        # Nền đen sâu (Gray 900)
    "card": "#1F2937",              # Card nổi (Gray 800)
    "card_hover": "#374151",        # Hover (Gray 700)
    "border": "#374151",            # Viền tối
    "hover": "#312E81",             # Hover tím đen (Indigo 900)
    "active": "#4F46E5",            
    "selected": "#3730A3",          # Selected chìm (Indigo 800)
    "treeview_bg": "#1F2937",       
    "treeview_selected": "#6366F1", 
    "treeview_hover": "#374151",    
    "disabled": "#4B5563",          # Disabled tối màu
    "gradient_start": "#818CF8",    
    "gradient_end": "#F472B6",      
    "shadow": "#00000060"           # Bóng mờ đậm
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
    # Ensure backwards compatibility - add missing keys if needed
    theme = ensure_theme_compatibility(theme)
    config['ui']['theme'] = theme
    config['ui']['dark_mode'] = dark_mode
    return config

def ensure_theme_compatibility(theme):
    """
    Đảm bảo theme có tất cả các keys cần thiết cho backwards compatibility
    
    Args:
        theme (dict): Theme hiện tại
        
    Returns:
        dict: Theme đã được bổ sung đầy đủ keys
    """
    # Default values for new keys
    default_keys = {
        'primary_dark': theme.get('primary', '#5A67D8'),
        'accent': theme.get('secondary', '#F687B3'),
        'info': theme.get('primary', '#4299E1'),
        'card_hover': theme.get('hover', '#FAFBFC'),
        'treeview_hover': theme.get('hover', '#F7FAFC'),
        'gradient_start': theme.get('primary', '#667EEA'),
        'gradient_end': theme.get('primary', '#764BA2'),
        'shadow': '#00000015'
    }
    
    # Add missing keys
    for key, default_value in default_keys.items():
        if key not in theme:
            theme[key] = default_value
    
    return theme 