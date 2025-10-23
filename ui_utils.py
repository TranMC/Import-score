"""
Module chứa các hàm tiện ích cho giao diện người dùng (UI utilities)
"""

import tkinter as tk
from tkinter import ttk
import themes

def center_window(window, width=None, height=None):
    """
    Căn giữa cửa sổ trên màn hình
    
    Args:
        window: Cửa sổ cần căn giữa
        width (int, optional): Chiều rộng cửa sổ
        height (int, optional): Chiều cao cửa sổ
    """
    if width is None:
        width = window.winfo_width()
    if height is None:
        height = window.winfo_height()
    
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    
    window.geometry(f"{width}x{height}+{x}+{y}")
    
    # Đảm bảo cửa sổ được cập nhật
    window.update_idletasks()

def apply_styles(config, style, root):
    """
    Áp dụng style hiện đại cho ttk widgets dựa trên config
    
    Args:
        config (dict): Cấu hình ứng dụng
        style: ttk.Style object
        root: Cửa sổ chính
    """
    # Lấy các font và theme từ config
    font_family = config['ui']['font_family']
    font_sizes = config['ui']['font_size']
    theme = config['ui']['theme']
    dark_mode = config['ui'].get('dark_mode', False)
    
    # Ensure theme compatibility - add missing keys if needed
    theme = themes.ensure_theme_compatibility(theme)
    config['ui']['theme'] = theme
    
    # Font mặc định với anti-aliasing
    default_font = (font_family, font_sizes['normal'])
    heading_font = (font_family, font_sizes['heading'], 'bold')
    button_font = (font_family, font_sizes['button'], 'bold')
    
    # Sử dụng theme cơ sở dễ tùy chỉnh nhất
    style.theme_use('clam')
    
    # ===== MODERN FRAME STYLES =====
    style.configure('TFrame', 
                   background=theme['background'])
    
    style.configure('Card.TFrame', 
                   background=theme['card'],
                   relief='flat',
                   borderwidth=0)
    
    style.configure('TLabelframe', 
                   background=theme['card'],
                   bordercolor=theme['border'],
                   relief='solid',
                   borderwidth=1)
    
    style.configure('TLabelframe.Label', 
                   font=heading_font,
                   foreground=theme['primary'],
                   background=theme['card'])
    
    # ===== MODERN LABEL STYLES =====
    style.configure('TLabel', 
                   font=default_font,
                   background=theme['background'],
                   foreground=theme['text'])
    
    # ===== MODERN BUTTON STYLES =====
    style.configure('TButton', 
                   font=button_font,
                   background=theme['primary'],
                   foreground='white',
                   borderwidth=0,
                   focuscolor='none',
                   padding=(20, 10),
                   relief='flat')
    
    style.map('TButton',
             foreground=[('pressed', 'white'), ('active', 'white'), ('disabled', theme['text_secondary'])],
             background=[('pressed', theme['active']), 
                        ('active', theme['primary_dark']),
                        ('disabled', theme['disabled'])],
             relief=[('pressed', 'flat'), ('active', 'flat')])
    
    # Accent Button Style
    style.configure('Accent.TButton',
                   font=button_font,
                   background=theme['accent'],
                   foreground='white',
                   borderwidth=0,
                   focuscolor='none',
                   padding=(20, 10),
                   relief='flat')
    
    style.map('Accent.TButton',
             foreground=[('pressed', 'white'), ('active', 'white')],
             background=[('pressed', theme['accent']), ('active', theme['accent'])],
             relief=[('pressed', 'flat'), ('active', 'flat')])
    
    # Success Button
    style.configure('Success.TButton',
                   font=button_font,
                   background=theme['success'],
                   foreground='white',
                   borderwidth=0,
                   focuscolor='none',
                   padding=(20, 10),
                   relief='flat')
    
    # Warning Button
    style.configure('Warning.TButton',
                   font=button_font,
                   background=theme['warning'],
                   foreground='white',
                   borderwidth=0,
                   focuscolor='none',
                   padding=(20, 10),
                   relief='flat')
    
    # ===== MODERN ENTRY STYLES =====
    style.configure('TEntry', 
                   font=default_font,
                   fieldbackground=theme['card'],
                   foreground=theme['text'],
                   bordercolor=theme['border'],
                   lightcolor=theme['primary'],
                   darkcolor=theme['primary'],
                   borderwidth=2,
                   relief='solid',
                   padding=8)
    
    style.map('TEntry',
             fieldbackground=[('focus', theme['card'])],
             bordercolor=[('focus', theme['primary']), ('!focus', theme['border'])],
             lightcolor=[('focus', theme['primary'])],
             darkcolor=[('focus', theme['primary'])])
    
    # ===== MODERN COMBOBOX STYLES =====
    style.configure('TCombobox', 
                   font=default_font,
                   fieldbackground=theme['card'],
                   foreground=theme['text'],
                   bordercolor=theme['border'],
                   arrowcolor=theme['primary'],
                   borderwidth=2,
                   relief='solid',
                   padding=8)
    
    style.map('TCombobox',
             fieldbackground=[('readonly', theme['card']), 
                            ('disabled', theme['background']),
                            ('focus', theme['card'])],
             selectbackground=[('readonly', theme['selected'])],
             selectforeground=[('readonly', 'white')],
             bordercolor=[('focus', theme['primary']), ('!focus', theme['border'])],
             arrowcolor=[('disabled', theme['disabled']), ('!disabled', theme['primary'])])
    
    # Combobox dropdown styling
    if root:
        try:
            root.option_add('*TCombobox*Listbox.background', theme['card'])
            root.option_add('*TCombobox*Listbox.foreground', theme['text'])
            root.option_add('*TCombobox*Listbox.selectBackground', theme['primary'])
            root.option_add('*TCombobox*Listbox.selectForeground', 'white')
            root.option_add('*TCombobox*Listbox.font', default_font)
            
            for widget in root.winfo_children():
                if isinstance(widget, ttk.Combobox):
                    widget.configure(style='TCombobox')
        except Exception as e:
            print(f"Lỗi khi cấu hình combobox dropdown: {str(e)}")
    
    # ===== MODERN TREEVIEW STYLES =====
    # Font lớn hơn cho Treeview
    treeview_font = (font_family, font_sizes['normal'] + 1)  # Tăng 1pt
    treeview_heading_font = (font_family, font_sizes['heading'], 'bold')
    
    style.configure('Treeview', 
                   font=treeview_font,
                   rowheight=40,  # Tăng từ 35 lên 40
                   background=theme['treeview_bg'],
                   fieldbackground=theme['treeview_bg'],
                   foreground=theme['text'],
                   borderwidth=0,
                   relief='flat')
    
    style.configure('Treeview.Heading', 
                   font=treeview_heading_font,
                   background=theme['primary'],
                   foreground='white',
                   borderwidth=0,
                   relief='flat',
                   padding=12)  # Tăng padding
    
    style.map('Treeview',
             background=[('selected', theme['treeview_selected'])],
             foreground=[('selected', 'white')])
    
    style.map('Treeview.Heading',
             background=[('active', theme['primary_dark'])],
             foreground=[('active', 'white')])
    
    # ===== MODERN SCROLLBAR STYLES =====
    scrollbar_bg = theme['background'] if not dark_mode else theme['card']
    scrollbar_trough = theme['card'] if not dark_mode else theme['background']
    
    style.configure('TScrollbar',
                   background=theme['primary'],
                   troughcolor=scrollbar_trough,
                   bordercolor=scrollbar_trough,
                   arrowcolor='white',
                   borderwidth=0,
                   relief='flat',
                   width=12)
    
    style.map('TScrollbar',
             background=[('active', theme['primary_dark']), 
                        ('pressed', theme['active']),
                        ('disabled', theme['disabled'])])
    
    # ===== MODERN PROGRESSBAR =====
    style.configure('TProgressbar',
                   background=theme['primary'],
                   troughcolor=theme['background'],
                   bordercolor=theme['border'],
                   lightcolor=theme['primary'],
                   darkcolor=theme['primary'],
                   borderwidth=0,
                   thickness=8)
    
    # ===== MODERN CHECKBUTTON =====
    style.configure('TCheckbutton',
                   font=default_font,
                   background=theme['background'],
                   foreground=theme['text'])
    
    style.map('TCheckbutton',
             background=[('active', theme['hover'])],
             foreground=[('disabled', theme['disabled'])])
    
    # ===== STATUS LABEL STYLES =====
    style.configure('Status.TLabel', 
                   font=default_font,
                   background=theme['card'],
                   foreground=theme['text'],
                   padding=5)
    
    style.configure('StatusGood.TLabel',
                   font=default_font,
                   background=theme['card'],
                   foreground=theme['success'])
    
    style.configure('StatusSuccess.TLabel',
                   font=default_font,
                   background=theme['card'],
                   foreground=theme['success'])
    
    style.configure('StatusWarning.TLabel',
                   font=default_font,
                   background=theme['card'],
                   foreground=theme['warning'])
    
    style.configure('StatusCritical.TLabel',
                   font=default_font,
                   background=theme['card'],
                   foreground=theme['error'])
    
    style.configure('StatusError.TLabel',
                   font=default_font,
                   background=theme['card'],
                   foreground=theme['error'])
    
    style.configure('StatusInfo.TLabel',
                   font=default_font,
                   background=theme['card'],
                   foreground=theme['info'])
    
    # ===== HEADING LABEL =====
    style.configure('Heading.TLabel',
                   font=heading_font,
                   foreground=theme['primary'],
                   background=theme['background'])
    
    # ===== ROOT WINDOW STYLING =====
    if root:
        root.configure(bg=theme['background'])
        
        root.tk_setPalette(
            background=theme['background'],
            foreground=theme['text'],
            activeBackground=theme['selected'],
            activeForeground='white'
        )
        
        # Menu styling
        try:
            menu_name = root.cget('menu')
            if menu_name:
                menu = root.nametowidget(menu_name)
                if menu:
                    menu.configure(
                        background=theme['card'],
                        foreground=theme['text'],
                        activebackground=theme['primary'],
                        activeforeground='white',
                        borderwidth=0,
                        relief='flat'
                    )
                    
                    for i in range(menu.index('end') + 1 if menu.index('end') is not None else 0):
                        try:
                            submenu = menu.nametowidget(menu.entrycget(i, 'menu'))
                            if submenu:
                                submenu.configure(
                                    background=theme['card'],
                                    foreground=theme['text'],
                                    activebackground=theme['primary'],
                                    activeforeground='white',
                                    borderwidth=0,
                                    relief='flat'
                                )
                        except:
                            pass
        except Exception as e:
            print(f"Lỗi khi cập nhật menu: {str(e)}")
    
    if root:
        root.update_idletasks()
    
    return style

def toggle_dark_mode(config, style, root):
    """
    Chuyển đổi giữa chế độ sáng và tối với cập nhật mượt mà
    
    Args:
        config (dict): Cấu hình ứng dụng
        style: ttk.Style object
        root: Cửa sổ chính
        
    Returns:
        dict: Cấu hình đã được cập nhật
    """
    dark_mode = config['ui'].get('dark_mode', False)
    print(f"UI Utils - Đang áp dụng chế độ: {'Tối' if dark_mode else 'Sáng'}")
    
    # Áp dụng theme mới
    config = themes.apply_theme_to_config(config, dark_mode)
    theme = config['ui']['theme']
    
    # Áp dụng style mới
    apply_styles(config, style, root)
    
    # Hàm cập nhật widget đệ quy
    def update_widget_recursively(widget):
        try:
            widget_class = widget.winfo_class()
            
            # Cập nhật TTK widgets
            if isinstance(widget, ttk.Widget):
                if isinstance(widget, ttk.Treeview):
                    widget.configure(style='Treeview')
                elif isinstance(widget, ttk.Combobox):
                    widget.configure(style='TCombobox')
                elif isinstance(widget, ttk.Entry):
                    widget.configure(style='TEntry')
                elif isinstance(widget, ttk.Button):
                    # Giữ nguyên style riêng của button
                    current_style = str(widget.cget('style'))
                    if current_style:
                        widget.configure(style=current_style)
                elif isinstance(widget, ttk.Label):
                    current_style = str(widget.cget('style'))
                    if current_style:
                        widget.configure(style=current_style)
                elif isinstance(widget, ttk.Frame):
                    current_style = str(widget.cget('style'))
                    if current_style:
                        widget.configure(style=current_style)
                    
            # Cập nhật Tkinter thuần widgets
            elif isinstance(widget, tk.Widget):
                if widget_class == 'Text':
                    widget.configure(bg=theme['card'], fg=theme['text'], 
                                   insertbackground=theme['text'])
                elif widget_class == 'Entry':
                    widget.configure(bg=theme['card'], fg=theme['text'], 
                                   insertbackground=theme['text'])
                elif widget_class == 'Label':
                    # Chỉ cập nhật nếu không có bg riêng (không phải header)
                    current_bg = widget.cget('bg')
                    if current_bg == theme.get('old_background', theme['background']) or \
                       current_bg in ['SystemButtonFace', '#f0f0f0', 'white']:
                        widget.configure(bg=theme['background'], fg=theme['text'])
                elif widget_class == 'Frame':
                    current_bg = widget.cget('bg')
                    # Giữ nguyên primary color cho header
                    if current_bg not in [theme.get('old_primary', ''), theme['primary']]:
                        widget.configure(bg=theme['background'])
                elif widget_class == 'LabelFrame':
                    widget.configure(bg=theme['background'], fg=theme['text'])
                elif widget_class == 'Menu':
                    widget.configure(
                        background=theme['card'],
                        foreground=theme['text'],
                        activebackground=theme['primary'],
                        activeforeground='white',
                        borderwidth=0
                    )
            
            # Đệ quy cập nhật con
            for child in widget.winfo_children():
                update_widget_recursively(child)
                
        except Exception as e:
            # Bỏ qua lỗi để tiếp tục với widgets khác
            pass
    
    # Lưu old values để so sánh
    theme['old_background'] = theme['background']
    theme['old_primary'] = theme['primary']
    
    # Cập nhật root window
    root.configure(bg=theme['background'])
    
    # Cập nhật toàn bộ cây widget
    update_widget_recursively(root)
    
    # Cập nhật combobox dropdown
    try:
        root.option_add('*TCombobox*Listbox.background', theme['card'])
        root.option_add('*TCombobox*Listbox.foreground', theme['text'])
        root.option_add('*TCombobox*Listbox.selectBackground', theme['primary'])
        root.option_add('*TCombobox*Listbox.selectForeground', 'white')
    except:
        pass
    
    # Cập nhật menu bar
    try:
        menubar = root.nametowidget(root.cget('menu'))
        if menubar:
            menubar.configure(
                background=theme['card'],
                foreground=theme['text'],
                activebackground=theme['primary'],
                activeforeground='white',
                borderwidth=0
            )
            
            # Cập nhật submenus
            for i in range(menubar.index('end') + 1 if menubar.index('end') is not None else 0):
                try:
                    submenu = menubar.nametowidget(menubar.entrycget(i, 'menu'))
                    if submenu:
                        submenu.configure(
                            background=theme['card'],
                            foreground=theme['text'],
                            activebackground=theme['primary'],
                            activeforeground='white',
                            borderwidth=0
                        )
                except:
                    pass
    except:
        pass
    
    # Force update
    root.update_idletasks()
    
    return config

def init_responsive_settings(root, config):
    """
    Khởi tạo cài đặt responsive cho cửa sổ
    
    Args:
        root: Cửa sổ chính
        config (dict): Cấu hình ứng dụng
    """
    responsive_config = config['ui'].get('responsive', {})
    
    # Lấy kích thước mặc định
    initial_width = responsive_config.get('initial_width', 900)
    initial_height = responsive_config.get('initial_height', 650)
    
    # Đặt kích thước tối thiểu
    min_width = responsive_config.get('min_width', 800)
    min_height = responsive_config.get('min_height', 600)
    root.minsize(min_width, min_height)
    
    # Đặt kích thước tối đa (nếu được cấu hình)
    max_width = responsive_config.get('max_width', 0)
    max_height = responsive_config.get('max_height', 0)
    if max_width > 0 and max_height > 0:
        root.maxsize(max_width, max_height)
    
    # Đặt kích thước ban đầu và căn giữa
    center_window(root, initial_width, initial_height)

def create_dark_mode_switch(parent, config, style, root, save_config):
    """
    Tạo switch để chuyển đổi giữa chế độ sáng và tối với update mượt mà
    
    Args:
        parent: Widget cha chứa switch
        config (dict): Cấu hình ứng dụng
        style: ttk.Style object
        root: Cửa sổ chính
        save_config: Hàm lưu cấu hình
        
    Returns:
        tk.Frame: Frame chứa switch
    """
    # Sử dụng Frame tk thuần để dễ cập nhật màu
    switch_frame = tk.Frame(parent, bg=config['ui']['theme']['primary'])
    
    # Tạo variable để theo dõi trạng thái
    is_dark_mode = config['ui'].get('dark_mode', False)
    
    # Thêm emoji mặt trời/mặt trăng để dễ nhận biết hơn
    mode_text = "☀️ Sáng" if is_dark_mode else "🌙 Tối"
    
    # Tạo button với tk thuần để dễ kiểm soát màu
    switch = tk.Button(switch_frame, 
                      text=mode_text,
                      font=(config['ui']['font_family'], 9, 'bold'),
                      bg='white' if is_dark_mode else '#2D3748',
                      fg='#2D3748' if is_dark_mode else 'white',
                      activebackground='#E2E8F0',
                      activeforeground='#2D3748',
                      relief='flat',
                      bd=0,
                      padx=12,
                      pady=6,
                      cursor='hand2')
    
    def toggle_theme():
        # Đảo ngược trạng thái dark mode
        config['ui']['dark_mode'] = not config['ui']['dark_mode']
        is_now_dark = config['ui']['dark_mode']
        
        # Cập nhật cấu hình theme
        updated_config = toggle_dark_mode(config, style, root)
        
        # Cập nhật config từ kết quả trả về
        for key in updated_config:
            config[key] = updated_config[key]
        
        # Cập nhật button
        new_text = "☀️ Sáng" if is_now_dark else "🌙 Tối"
        new_bg = 'white' if is_now_dark else '#2D3748'
        new_fg = '#2D3748' if is_now_dark else 'white'
        
        switch.config(text=new_text, bg=new_bg, fg=new_fg)
        switch_frame.config(bg=config['ui']['theme']['primary'])
        
        # Lưu cấu hình
        save_config()
        
        # Force update toàn bộ giao diện
        root.update()
    
    switch.config(command=toggle_theme)
    switch.pack(side="right")
    
    return switch_frame