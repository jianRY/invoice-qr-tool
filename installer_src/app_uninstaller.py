# -*- coding: utf-8 -*-
"""发票二维码工具 —— 卸载程序（PyInstaller onefile 打包，需管理员权限写注册表）。

逻辑：
- 从注册表读取安装目录；读不到则回退到本程序所在目录。
- 删除程序目录、开始菜单/桌面快捷方式、卸载注册表项。
- 纯 tkinter 简单 GUI，确认后执行卸载。
"""
import os
import sys
import shutil
import tkinter as tk
from tkinter import messagebox

from appicon import setup_app_id, apply_window_icon

APP_NAME = "发票二维码工具"
INSTALL_REG = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\{}".format(APP_NAME)


def find_install_dir():
    try:
        import winreg
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                k = winreg.OpenKey(root, INSTALL_REG)
                p = winreg.QueryValueEx(k, "InstallLocation")[0]
                winreg.CloseKey(k)
                if p and os.path.isdir(p):
                    return p
            except Exception:
                pass
    except Exception:
        pass
    return os.path.dirname(os.path.abspath(sys.argv[0]))


def remove_shortcuts():
    prog = os.environ.get("ProgramData", r"C:\ProgramData")
    cu = os.environ.get("APPDATA", "")
    user = os.environ.get("USERPROFILE", "")
    candidates = [
        os.path.join(prog, "Microsoft", "Windows", "Start Menu", "Programs", APP_NAME + ".lnk"),
        os.path.join(cu, "Microsoft", "Windows", "Start Menu", "Programs", APP_NAME + ".lnk"),
        os.path.join(user, "Desktop", APP_NAME + ".lnk"),
    ]
    for p in candidates:
        try:
            if os.path.exists(p):
                os.remove(p)
        except Exception:
            pass


def remove_registry():
    try:
        import winreg
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                winreg.DeleteKey(root, INSTALL_REG)
            except Exception:
                pass
    except Exception:
        pass


def do_uninstall():
    d = find_install_dir()
    if d and os.path.isdir(d):
        try:
            shutil.rmtree(d, ignore_errors=True)
        except Exception as e:
            messagebox.showerror("卸载失败", "删除目录出错：\n" + str(e))
            return
    remove_shortcuts()
    remove_registry()
    messagebox.showinfo("卸载完成", "{} 已卸载。".format(APP_NAME))
    sys.exit(0)


def main():
    setup_app_id()          # 必须在 tk.Tk() 之前
    root = tk.Tk()
    root.title("卸载 " + APP_NAME)
    root.geometry("380x170")
    root.resizable(False, False)
    apply_window_icon(root)
    tk.Label(
        root,
        text="确定要卸载 {} 吗？\n该操作将删除程序目录及快捷方式。".format(APP_NAME),
        wraplength=320,
        justify="left",
    ).pack(padx=20, pady=18)
    f = tk.Frame(root)
    f.pack(pady=6)
    tk.Button(f, text="卸载", width=12, command=do_uninstall).pack(side="left", padx=12)
    tk.Button(f, text="取消", width=12, command=root.destroy).pack(side="left", padx=12)
    root.mainloop()


if __name__ == "__main__":
    main()
