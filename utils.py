"""通用工具函数：窗口居中、跨平台打开目录等。"""
import os
import sys


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
