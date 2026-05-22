"""苹果系统亮色主题：配色 / 字体 / 控件样式预设 / Banner 提示。

使用方式：
    import customtkinter as ctk
    from apple_theme import apply_apple_theme, font_ui, BUTTON_PRIMARY

    root = ctk.CTk()
    apply_apple_theme(root)
    ctk.CTkButton(root, text='转换', font=font_ui(13, 'bold'), **BUTTON_PRIMARY).pack()
"""
from tkinter import font as tkfont, ttk

import customtkinter as ctk


# ============ 配色（Apple HIG，仅亮色） ============

BLUE = '#007AFF'
BLUE_HOVER = '#0066D6'
BLUE_TINT_BG = '#E5F1FF'
BLUE_TINT_HOVER = '#CCE0FF'
RED = '#FF3B30'
RED_TINT_BG = '#FFE4E1'
GREEN = '#34C759'
GREEN_TINT_BG = '#DFFFE2'
ORANGE = '#FF9500'
ORANGE_TINT_BG = '#FFF4D6'

WINDOW_BG = '#F2F2F7'
CARD_BG = '#FFFFFF'
INPUT_BG = '#F2F2F7'
HOVER_BG = '#E5E5EA'
DIVIDER = '#D1D1D6'

TEXT_PRIMARY = '#1C1C1E'
TEXT_SECONDARY = '#6C6C70'
TEXT_TERTIARY = '#A1A1A6'

# Banner 配色（图标用 Unicode 符号，避免引入图标字体）
BANNER_STYLES = {
    'info':    {'bg': BLUE_TINT_BG,   'fg': BLUE,      'icon': 'ℹ'},
    'success': {'bg': GREEN_TINT_BG,  'fg': '#1E7E34', 'icon': '✓'},
    'warning': {'bg': ORANGE_TINT_BG, 'fg': '#A86200', 'icon': '⚠'},
    'error':   {'bg': RED_TINT_BG,    'fg': '#C42820', 'icon': '✕'},
}


# ============ 控件样式预设（spread 到 CTkXxx 构造器） ============

CARD_STYLE = {
    'fg_color': CARD_BG,
    'corner_radius': 12,
    'border_width': 0,
}

BUTTON_PRIMARY = {
    'fg_color': BLUE,
    'hover_color': BLUE_HOVER,
    'text_color': '#FFFFFF',
    'corner_radius': 8,
    'height': 40,
}

BUTTON_SECONDARY = {
    'fg_color': BLUE_TINT_BG,
    'hover_color': BLUE_TINT_HOVER,
    'text_color': BLUE,
    'corner_radius': 8,
    'height': 32,
}

BUTTON_PLAIN = {
    'fg_color': 'transparent',
    'hover_color': HOVER_BG,
    'text_color': BLUE,
    'corner_radius': 6,
    'height': 28,
}

BUTTON_DANGER = {
    'fg_color': 'transparent',
    'hover_color': RED_TINT_BG,
    'text_color': RED,
    'corner_radius': 6,
    'height': 28,
}

ENTRY_STYLE = {
    'fg_color': INPUT_BG,
    'text_color': TEXT_PRIMARY,
    'corner_radius': 6,
    'border_width': 0,
    'height': 32,
}

TEXTBOX_STYLE = {
    'fg_color': '#FAFAFA',
    'text_color': TEXT_PRIMARY,
    'corner_radius': 8,
    'border_width': 0,
}


# ============ 字体（Windows 落地：YaHei UI / Segoe UI） ============

_FONT_UI_CANDIDATES = ['Microsoft YaHei UI', 'Segoe UI', 'PingFang SC', 'Microsoft YaHei']
_FONT_MONO_CANDIDATES = ['Cascadia Code', 'JetBrains Mono', 'Consolas', 'Courier New']

_font_ui_family = 'Microsoft YaHei UI'
_font_mono_family = 'Consolas'


def _pick_font(candidates):
    available = set(tkfont.families())
    for name in candidates:
        if name in available:
            return name
    return candidates[-1]


def font_ui(size=12, weight='normal'):
    """主字体（用于 ctk 控件）。必须在 apply_apple_theme 之后调用。"""
    return ctk.CTkFont(family=_font_ui_family, size=size, weight=weight)


def font_title(size=15):
    return ctk.CTkFont(family=_font_ui_family, size=size, weight='bold')


def font_mono(size=11):
    return ctk.CTkFont(family=_font_mono_family, size=size)


def font_tuple(size=12, weight='normal', mono=False):
    """字体元组（用于 ttk style / Treeview / tk.Label 等不接受 CTkFont 的场景）。"""
    family = _font_mono_family if mono else _font_ui_family
    return (family, size, weight)


# ============ 纯布局容器 helper ============

def transparent_frame(parent, **kwargs):
    """纯布局用 CTkFrame：fg_color=transparent + width/height 默认 0。

    CTkFrame 默认 width=height=200，对纯布局容器会撑出 200px 空白。
    显式传 height/width 可覆盖（如固定 height 的标题栏）。
    """
    kwargs.setdefault('fg_color', 'transparent')
    kwargs.setdefault('width', 0)
    kwargs.setdefault('height', 0)
    return ctk.CTkFrame(parent, **kwargs)


# ============ 主题应用 ============

def apply_apple_theme(root):
    """初始化 ctk 模式、字体检测、ttk 控件样式。

    必须在 root = ctk.CTk() 之后、添加任何子控件之前调用。
    """
    global _font_ui_family, _font_mono_family

    ctk.set_appearance_mode('light')
    ctk.set_default_color_theme('blue')

    _font_ui_family = _pick_font(_FONT_UI_CANDIDATES)
    _font_mono_family = _pick_font(_FONT_MONO_CANDIDATES)

    try:
        root.configure(fg_color=WINDOW_BG)
    except Exception:
        pass

    style = ttk.Style()
    try:
        style.theme_use('clam')
    except Exception:
        pass

    # ---- Treeview 全局：白底、行高 28、选中蓝 ----
    style.configure('Treeview',
                    background=CARD_BG,
                    fieldbackground=CARD_BG,
                    foreground=TEXT_PRIMARY,
                    borderwidth=0,
                    rowheight=28,
                    font=(_font_ui_family, 11))
    style.configure('Treeview.Heading',
                    background=WINDOW_BG,
                    foreground=TEXT_SECONDARY,
                    relief='flat',
                    borderwidth=0,
                    padding=(8, 6),
                    font=(_font_ui_family, 11, 'bold'))
    style.map('Treeview',
              background=[('selected', BLUE)],
              foreground=[('selected', '#FFFFFF')])
    style.map('Treeview.Heading',
              background=[('active', HOVER_BG)])

    # ---- Combobox 全局 ----
    style.configure('TCombobox',
                    fieldbackground=INPUT_BG,
                    background=INPUT_BG,
                    foreground=TEXT_PRIMARY,
                    arrowcolor=TEXT_SECONDARY,
                    borderwidth=0,
                    padding=6,
                    relief='flat')
    style.map('TCombobox',
              fieldbackground=[('readonly', INPUT_BG), ('focus', '#FFFFFF')],
              bordercolor=[('focus', BLUE)])

    # ---- Scrollbar 细化 ----
    for orient in ('Vertical', 'Horizontal'):
        style.configure(f'{orient}.TScrollbar',
                        background=HOVER_BG,
                        troughcolor=WINDOW_BG,
                        borderwidth=0,
                        arrowsize=12)
        style.map(f'{orient}.TScrollbar',
                  background=[('active', TEXT_TERTIARY)])

    return style


# ============ Banner（顶部弹条提示，替代 messagebox） ============

def show_banner(banner_area, message, level='info', duration=3500):
    """在 banner_area 中显示彩色提示条。同一个 area 同一时刻只保留一个。

    banner_area: 预留在窗口顶部的空 CTkFrame
    level: 'info' | 'success' | 'warning' | 'error'
    duration: 自动消失毫秒数；0 = 不自动消失
    """
    style = BANNER_STYLES.get(level, BANNER_STYLES['info'])

    for child in banner_area.winfo_children():
        try:
            child.destroy()
        except Exception:
            pass

    banner = ctk.CTkFrame(banner_area, fg_color=style['bg'], corner_radius=8)
    banner.pack(fill='x', padx=20, pady=(8, 0))

    ctk.CTkLabel(banner, text=style['icon'], text_color=style['fg'],
                 font=font_ui(14, 'bold'), width=20
                 ).pack(side='left', padx=(12, 4), pady=8)

    ctk.CTkLabel(banner, text=message, text_color=style['fg'],
                 font=font_ui(12), anchor='w', justify='left',
                 wraplength=500
                 ).pack(side='left', fill='x', expand=True, pady=8)

    ctk.CTkButton(banner, text='✕', width=22, height=22,
                  fg_color='transparent', text_color=style['fg'],
                  hover_color=style['bg'],
                  font=font_ui(11),
                  command=banner.destroy
                  ).pack(side='right', padx=(4, 8), pady=8)

    def _safe_destroy():
        try:
            if banner.winfo_exists():
                banner.destroy()
        except Exception:
            pass

    if duration > 0:
        banner.after(duration, _safe_destroy)

    return banner


# ============ Modal 二选一对话框（替代 messagebox.askyesno） ============

def ask_yes_no(parent, title, message, yes_text='确定', no_text='取消',
               yes_style=None, width=380, height=160):
    """阻塞式二选一对话框。返回 True / False。

    parent: 父窗口（用于居中定位）
    yes_style: 'primary' (默认) / 'danger'
    """
    yes_btn_style = BUTTON_DANGER if yes_style == 'danger' else BUTTON_PRIMARY

    dialog = ctk.CTkToplevel(parent)
    dialog.title(title)
    dialog.geometry(f'{width}x{height}')
    dialog.resizable(False, False)
    dialog.configure(fg_color=WINDOW_BG)
    dialog.transient(parent)
    dialog.grab_set()

    parent.update_idletasks()
    px = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
    py = parent.winfo_rooty() + (parent.winfo_height() - height) // 2
    dialog.geometry(f'{width}x{height}+{max(0, px)}+{max(0, py)}')

    result = {'value': False}

    ctk.CTkLabel(dialog, text=message, font=font_ui(12),
                 text_color=TEXT_PRIMARY, wraplength=width - 40,
                 justify='left'
                 ).pack(padx=20, pady=(24, 16), fill='x')

    btn_row = transparent_frame(dialog)
    btn_row.pack(pady=(0, 16))

    def _no():
        dialog.destroy()

    def _yes():
        result['value'] = True
        dialog.destroy()

    ctk.CTkButton(btn_row, text=no_text, command=_no,
                  width=100, font=font_ui(12), **BUTTON_PLAIN
                  ).pack(side='left', padx=6)
    ctk.CTkButton(btn_row, text=yes_text, command=_yes,
                  width=140, font=font_ui(13, 'bold'), **yes_btn_style
                  ).pack(side='left', padx=6)

    parent.wait_window(dialog)
    return result['value']
