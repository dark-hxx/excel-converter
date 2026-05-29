"""通用工具函数：窗口居中、跨平台打开目录等。"""
import os
import sys
import tkinter as tk
from tkinter import ttk


INVALID_FILENAME_CHARS = '<>:"/\\|?*'


def center_window(win, width, height):
    """将窗口居中到屏幕中央并设置尺寸"""
    win.update_idletasks()
    screen_w = win.winfo_screenwidth()
    screen_h = win.winfo_screenheight()
    x = (screen_w - width) // 2
    y = (screen_h - height) // 2
    win.geometry(f'{width}x{height}+{x}+{y}')


def open_folder(path):
    """跨平台打开文件夹"""
    try:
        if sys.platform == 'win32':
            os.startfile(path)
        elif sys.platform == 'darwin':
            import subprocess
            subprocess.run(['open', path], check=False)
        else:
            import subprocess
            subprocess.run(['xdg-open', path], check=False)
    except Exception:
        pass


def sanitize_filename_part(value):
    filename = ''.join(
        '_' if char in INVALID_FILENAME_CHARS else char
        for char in str(value)
    ).strip()
    return filename or '未命名'


def build_unique_save_path(directory, filename):
    save_path = os.path.join(directory, f'{filename}.xlsx')
    if not os.path.exists(save_path):
        return save_path

    index = 1
    while True:
        save_path = os.path.join(directory, f'{filename}_{index}.xlsx')
        if not os.path.exists(save_path):
            return save_path
        index += 1


def make_searchable_combobox(parent, values, **kwargs):
    """创建可搜索的下拉框：键入时按子串模糊过滤候选项。

    返回 (combo, var) tuple。kwargs 透传给 ttk.Combobox。
    注意：不要传 state='readonly'，那样无法输入搜索文本。
    """
    var = kwargs.pop('textvariable', None) or tk.StringVar()
    kwargs['textvariable'] = var
    kwargs.setdefault('values', list(values))
    combo = ttk.Combobox(parent, **kwargs)
    full_values = list(values)

    def on_keyrelease(event):
        if event.keysym in ('Up', 'Down', 'Return', 'Escape', 'Tab', 'Left', 'Right'):
            return
        typed = var.get().strip().lower()
        combo['values'] = [v for v in full_values if typed in v.lower()] or full_values

    combo.bind('<KeyRelease>', on_keyrelease)
    return combo, var
