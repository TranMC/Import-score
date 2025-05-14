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
    Áp dụng style cho ttk widgets dựa trên config
    
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
    
    # Font mặc định
    default_font = (font_family, font_sizes['normal'])
    heading_font = (font_family, font_sizes['heading'], 'bold')
    button_font = (font_family, font_sizes['button'])
    
    # Sử dụng theme cơ sở dễ tùy chỉnh nhất
    style.theme_use('clam')
    
    # Cấu hình theme
    # Frame
    style.configure('TFrame', background=theme['background'])
    style.configure('TLabelframe', background=theme['background'], bordercolor=theme['border'])
    style.configure('TLabelframe.Label', font=heading_font, foreground=theme['primary'], background=theme['background'])
    
    # Label
    style.configure('TLabel', font=default_font, background=theme['background'], foreground=theme['text'])
    
    # Button
    style.configure('TButton', font=button_font, background=theme['primary'], foreground='white', borderwidth=0)
    style.map('TButton',
             foreground=[('pressed', 'white'), ('active', 'white')],
             background=[('pressed', theme['active']), ('active', theme['hover'])])
    
    # Entry
    style.configure('TEntry', font=default_font, fieldbackground=theme['card'], 
                  foreground=theme['text'], bordercolor=theme['border'])
    
    # Combobox
    style.configure('TCombobox', font=default_font, 
                  fieldbackground=theme['card'], 
                  foreground=theme['text'],
                  bordercolor=theme['border'],
                  arrowcolor=theme['primary'])
    style.map('TCombobox',
            fieldbackground=[('readonly', theme['card']), ('disabled', theme['background'])],
            selectbackground=[('readonly', theme['selected'])],
            selectforeground=[('readonly', 'white')])
    
    # Combobox dropdown (popup) - cải thiện để áp dụng hiệu quả hơn
    if root:
        try:
            root.option_add('*TCombobox*Listbox.background', theme['card'])
            root.option_add('*TCombobox*Listbox.foreground', theme['text'])
            root.option_add('*TCombobox*Listbox.selectBackground', theme['selected'])
            root.option_add('*TCombobox*Listbox.selectForeground', 'white')
            
            # Áp dụng ngay cho tất cả combobox hiện có
            for widget in root.winfo_children():
                if isinstance(widget, ttk.Combobox):
                    widget.configure(style='TCombobox')
        except Exception as e:
            print(f"Lỗi khi cấu hình combobox dropdown: {str(e)}")
    
    # Heading
    style.configure('Heading.TLabel', font=heading_font, foreground=theme['primary'])
    
    # Treeview
    style.configure('Treeview', font=default_font, rowheight=28, 
                   background=theme['treeview_bg'], fieldbackground=theme['treeview_bg'],
                   foreground=theme['text'])
    style.configure('Treeview.Heading', font=heading_font, 
                   background=theme['primary'], foreground='white')
    style.map('Treeview',
             background=[('selected', theme['treeview_selected'])],
             foreground=[('selected', 'white')])
    
    # Scrollbar - cải thiện tính nhìn thấy của scrollbar trong dark mode
    scrollbar_bg = theme['background'] if not dark_mode else theme['card']
    scrollbar_trough = theme['card'] if not dark_mode else theme['background']
    
    style.configure('TScrollbar', background=scrollbar_bg, 
                  troughcolor=scrollbar_trough,
                  bordercolor=theme['border'],
                  arrowcolor=theme['text'])
    style.map('TScrollbar',
            background=[('active', theme['primary']), ('disabled', theme['disabled'])])
    
    # Progressbar
    style.configure('TProgressbar', 
                  background=theme['primary'],
                  troughcolor=theme['background'],
                  bordercolor=theme['border'])
    
    # Checkbutton
    style.configure('TCheckbutton', 
                  font=default_font,
                  background=theme['background'],
                  foreground=theme['text'])
    
    # Status Labels
    style.configure('Status.TLabel', font=default_font, background=theme['card'], foreground=theme['text'])
    style.configure('StatusGood.TLabel', font=default_font, background=theme['card'], foreground=theme['success'])
    style.configure('StatusWarning.TLabel', font=default_font, background=theme['card'], foreground=theme['warning'])
    style.configure('StatusCritical.TLabel', font=default_font, background=theme['card'], foreground=theme['error'])
    style.configure('StatusError.TLabel', font=default_font, background=theme['card'], foreground=theme['error'])
    style.configure('StatusInfo.TLabel', font=default_font, background=theme['card'], foreground=theme['primary'])
    
    # Card Frame
    style.configure('Card.TFrame', background=theme['card'])
    
    # Thiết lập màu nền cho cửa sổ root (không phải ttk)
    if root:
        # Cập nhật màu nền chính cho cửa sổ
        root.configure(bg=theme['background'])
        
        # Đặc biệt xử lý combobox để đảm bảo cập nhật đúng
        root.tk_setPalette(
            background=theme['background'],
            foreground=theme['text'],
            activeBackground=theme['selected'],
            activeForeground='white'
        )
        
        # Đặt màu cho menu (nếu có)
        try:
            menu_name = root.cget('menu')
            if menu_name:
                menu = root.nametowidget(menu_name)
                if menu:
                    menu.configure(background=theme['card'], foreground=theme['text'],
                                activebackground=theme['selected'], activeforeground='white')
                    
                    # Cập nhật các submenu
                    for i in range(menu.index('end') + 1 if menu.index('end') is not None else 0):
                        try:
                            submenu = menu.nametowidget(menu.entrycget(i, 'menu'))
                            if submenu:
                                submenu.configure(
                                    background=theme['card'],
                                    foreground=theme['text'],
                                    activebackground=theme['selected'],
                                    activeforeground='white'
                                )
                        except:
                            pass
        except Exception as e:
            print(f"Lỗi khi cập nhật menu: {str(e)}")
    
    # Cập nhật cửa sổ
    if root:
        root.update_idletasks()
    
    return style

def toggle_dark_mode(config, style, root):
    """
    Chuyển đổi giữa chế độ sáng và tối
    
    Args:
        config (dict): Cấu hình ứng dụng
        style: ttk.Style object
        root: Cửa sổ chính
        
    Returns:
        dict: Cấu hình đã được cập nhật
    """
    # Lấy trạng thái dark mode hiện tại - đã được cập nhật bởi hàm gọi toggle_theme
    dark_mode = config['ui'].get('dark_mode', False)
    print(f"UI Utils - Đang áp dụng chế độ: {'Tối' if dark_mode else 'Sáng'}")
    
    # Áp dụng theme mới
    config = themes.apply_theme_to_config(config, dark_mode)
    
    # Áp dụng style mới
    apply_styles(config, style, root)
    
    # Cập nhật màu nền cho tất cả widget
    theme = config['ui']['theme']
    
    # Cấu hình lại tất cả widget con để đảm bảo chúng được cập nhật
    def update_widget_styles(widget):
        try:
            # Cập nhật các widget là ttk
            if isinstance(widget, ttk.Widget):
                # Đặc biệt xử lý một số loại widget cụ thể
                if isinstance(widget, ttk.Treeview):
                    # Cập nhật style cho Treeview
                    widget.configure(style='Treeview')
                    for child_id in widget.get_children():
                        widget.item(child_id, tags=())
                elif isinstance(widget, ttk.Combobox):
                    # Đặc biệt cập nhật style cho Combobox
                    widget.configure(style='TCombobox')
                    # Đảm bảo dropdown list được cập nhật
                    if hasattr(widget, 'tk'):
                        widget.tk.eval(f"""
                        option add *TCombobox*Listbox.background {theme['card']} widgetDefault
                        option add *TCombobox*Listbox.foreground {theme['text']} widgetDefault
                        option add *TCombobox*Listbox.selectBackground {theme['selected']} widgetDefault
                        option add *TCombobox*Listbox.selectForeground white widgetDefault
                        """)
                elif isinstance(widget, ttk.Entry):
                    # Cập nhật style cho Entry
                    widget.configure(style='TEntry')
                
            # Cập nhật cho widget tkinter thuần
            elif isinstance(widget, tk.Widget):
                if isinstance(widget, tk.Text):
                    widget.configure(bg=theme['card'], fg=theme['text'])
                elif isinstance(widget, tk.Entry):
                    widget.configure(bg=theme['card'], fg=theme['text'])
                elif isinstance(widget, tk.Label):
                    widget.configure(bg=theme['background'], fg=theme['text'])
                elif isinstance(widget, tk.Frame) or isinstance(widget, tk.LabelFrame):
                    widget.configure(bg=theme['background'])
                    
                # Cập nhật menu
                if isinstance(widget, tk.Menu):
                    widget.configure(
                        background=theme['card'],
                        foreground=theme['text'],
                        activebackground=theme['selected'],
                        activeforeground='white'
                    )
            
            # Đệ quy qua các widget con
            children = widget.winfo_children()
            for child in children:
                update_widget_styles(child)
        except Exception as e:
            print(f"Lỗi khi cập nhật widget: {str(e)}")
            pass

    # Cập nhật toàn bộ cây widget
    update_widget_styles(root)
    
    # Cập nhật menu
    try:
        menubar = root.nametowidget(root.cget('menu'))
        if menubar:
            menubar.configure(
                background=theme['card'],
                foreground=theme['text'],
                activebackground=theme['selected'],
                activeforeground='white'
            )
            # Cập nhật tất cả menu con
            for i in range(menubar.index('end') + 1):
                try:
                    submenu = menubar.nametowidget(menubar.entrycget(i, 'menu'))
                    if submenu:
                        submenu.configure(
                            background=theme['card'],
                            foreground=theme['text'],
                            activebackground=theme['selected'],
                            activeforeground='white'
                        )
                except:
                    pass
    except Exception as e:
        print(f"Lỗi khi cập nhật menu: {str(e)}")
    
    # Force update giao diện
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
    Tạo switch để chuyển đổi giữa chế độ sáng và tối
    
    Args:
        parent: Widget cha chứa switch
        config (dict): Cấu hình ứng dụng
        style: ttk.Style object
        root: Cửa sổ chính
        save_config: Hàm lưu cấu hình
        
    Returns:
        ttk.Frame: Frame chứa switch
    """
    switch_frame = ttk.Frame(parent)
    
    # Tạo variable để theo dõi trạng thái
    is_dark_mode = config['ui'].get('dark_mode', False)
    dark_mode_var = tk.BooleanVar(value=is_dark_mode)
    
    # Cấu hình style cho checkbutton
    theme = config['ui']['theme']
    
    # Tạo style cho button điều khiển dark mode rõ ràng hơn
    style.configure('Switch.TCheckbutton', 
                  background=theme['background'],
                  foreground=theme['text'],
                  font=(config['ui']['font_family'], config['ui']['font_size']['normal']))
    
    # Thêm emoji mặt trời/mặt trăng để dễ nhận biết hơn
    mode_text = "☀️ Chế độ sáng" if is_dark_mode else "🌙 Chế độ tối"
    
    # Tạo checkbutton làm switch
    def toggle_theme():
        # Đảo ngược trạng thái dark mode trực tiếp trong config
        config['ui']['dark_mode'] = not config['ui']['dark_mode']
        
        # Cập nhật cấu hình
        updated_config = toggle_dark_mode(config, style, root)
        
        # Cập nhật config từ kết quả trả về
        for key in updated_config:
            config[key] = updated_config[key]
        
        # Cập nhật text theo trạng thái mới
        is_now_dark = config['ui']['dark_mode']
        new_text = "☀️ Chế độ sáng" if is_now_dark else "🌙 Chế độ tối"
        switch.config(text=new_text)
        
        # Đảm bảo trạng thái biến khớp với config
        dark_mode_var.set(is_now_dark)
        
        # Lưu cấu hình
        save_config()
        
        # Force update toàn bộ giao diện
        root.update()
    
    # Tạo nút chuyển đổi (không sử dụng Checkbutton vì có thể gây nhầm lẫn)
    switch = ttk.Button(switch_frame, 
                      text=mode_text,
                      command=toggle_theme,
                      style='Toggle.TButton',
                      width=15)
    
    # Tạo style đặc biệt cho nút toggle
    style.configure('Toggle.TButton', 
                  font=(config['ui']['font_family'], config['ui']['font_size']['normal']),
                  background=theme['primary'],
                  foreground='white')
    
    switch.pack(side="right", padx=5)
    
    return switch_frame