"""
Module để quản lý các theme của ứng dụng (light/dark mode)
"""

# Theme mặc định cho chế độ sáng (light mode) - Modern & Stylish
LIGHT_THEME = {
    "primary": "#667EEA",           # Purple gradient start - hiện đại
    "primary_dark": "#5A67D8",      # Purple đậm hơn
    "secondary": "#48BB78",         # Green tươi sáng
    "accent": "#F687B3",            # Pink accent cho điểm nhấn
    "warning": "#F6AD55",           # Cam ấm áp
    "error": "#FC8181",             # Đỏ san hô
    "success": "#68D391",           # Xanh lá tươi
    "info": "#4299E1",              # Xanh dương thông tin
    "text": "#2D3748",              # Xám đen mềm
    "text_secondary": "#718096",    # Xám trung tính
    "background": "#F7FAFC",        # Nền trắng xanh nhẹ
    "card": "#FFFFFF",              # Card trắng tinh
    "card_hover": "#FAFBFC",        # Card hover nhẹ
    "border": "#E2E8F0",            # Viền mềm mại
    "hover": "#EDF2F7",             # Hover nhẹ nhàng
    "active": "#5A67D8",            # Active tím đậm
    "selected": "#D6BCFA",          # Selected tím nhạt
    "treeview_bg": "#FFFFFF",       # Nền treeview
    "treeview_selected": "#667EEA", # Tím gradient
    "treeview_hover": "#F7FAFC",    # Hover nhẹ
    "disabled": "#CBD5E0",          # Disabled xám nhạt
    "gradient_start": "#667EEA",    # Gradient tím
    "gradient_end": "#764BA2",      # Gradient tím đậm
    "shadow": "#00000015"           # Shadow nhẹ
}

# Theme cho chế độ tối (dark mode) - Modern & Stylish
DARK_THEME = {
    "primary": "#7C3AED",           # Purple vibrant
    "primary_dark": "#6D28D9",      # Purple đậm
    "secondary": "#10B981",         # Green emerald
    "accent": "#EC4899",            # Pink neon
    "warning": "#FBBF24",           # Vàng ấm
    "error": "#EF4444",             # Đỏ neon
    "success": "#34D399",           # Xanh lá neon
    "info": "#3B82F6",              # Xanh dương sáng
    "text": "#F9FAFB",              # Trắng gần tinh
    "text_secondary": "#D1D5DB",    # Xám sáng
    "background": "#111827",        # Nền xanh đen
    "card": "#1F2937",              # Card xanh đậm
    "card_hover": "#374151",        # Card hover
    "border": "#374151",            # Viền xanh xám
    "hover": "#374151",             # Hover xanh xám
    "active": "#6D28D9",            # Active tím
    "selected": "#A78BFA",          # Selected tím sáng
    "treeview_bg": "#1F2937",       # Nền treeview
    "treeview_selected": "#7C3AED", # Tím vibrant
    "treeview_hover": "#374151",    # Hover xanh xám
    "disabled": "#6B7280",          # Disabled xám
    "gradient_start": "#7C3AED",    # Gradient tím
    "gradient_end": "#EC4899",      # Gradient pink
    "shadow": "#00000040"           # Shadow đậm hơn
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