# -*- coding: utf-8 -*-
"""安装程序 / 卸载程序共用的图标工具（打包时经 --add-data 把 app_icon.ico 放进包内）。

与主程序 invoice_qr_tool.py 里的 setup_app_id / apply_window_icon 行为一致：
- setup_app_id()      设置任务栏身份，须在 tk.Tk() 之前调用
- apply_window_icon() 用 Win32 LoadImage 按 32/16px 显式设置窗口图标
  （Tk 的 iconbitmap 在 Windows 上会取 ICO 内最小档再放大，导致图标发虚）
"""
import os
import sys


def _icon_path():
    cands = []
    base = getattr(sys, "_MEIPASS", None)
    if base:
        cands.append(os.path.join(base, "app_icon.ico"))
    here = os.path.dirname(os.path.abspath(__file__))
    cands.append(os.path.join(here, "app_icon.ico"))
    cands.append(os.path.join(os.path.dirname(here), "app_icon.ico"))
    for c in cands:
        if os.path.isfile(c):
            return c
    return None


def setup_app_id(app_id="jianRY.InvoiceQrTool"):
    """设置 Windows 任务栏身份（AppUserModelID），必须在 tk.Tk() 之前调用。"""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass


def apply_window_icon(win):
    """给窗口设置应用图标（主窗口与每个 Toplevel 都要调用）。"""
    ico = _icon_path()
    if not ico:
        return
    try:
        win.iconbitmap(ico)          # 非 Windows 平台靠它；Windows 上仅作保底
    except Exception:
        pass
    if sys.platform != "win32":
        return
    try:
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        LR_LOADFROMFILE, IMAGE_ICON, WM_SETICON = 0x0010, 1, 0x0080
        user32.LoadImageW.argtypes = [wintypes.HINSTANCE, wintypes.LPCWSTR, wintypes.UINT,
                                      ctypes.c_int, ctypes.c_int, wintypes.UINT]
        user32.LoadImageW.restype = ctypes.c_void_p
        user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT,
                                        ctypes.c_void_p, ctypes.c_void_p]
        user32.SendMessageW.restype = ctypes.c_void_p
        win.update_idletasks()
        hwnd = user32.GetAncestor(win.winfo_id(), 2)   # GA_ROOT = 2
        if not hwnd:
            return
        big = user32.LoadImageW(None, ico, IMAGE_ICON, 32, 32, LR_LOADFROMFILE)
        small = user32.LoadImageW(None, ico, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
        if big:
            user32.SendMessageW(hwnd, WM_SETICON, ctypes.c_void_p(1), ctypes.c_void_p(big))
        if small:
            user32.SendMessageW(hwnd, WM_SETICON, ctypes.c_void_p(0), ctypes.c_void_p(small))
    except Exception:
        pass
