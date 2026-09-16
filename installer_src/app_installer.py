# -*- coding: utf-8 -*-
r"""发票二维码工具 —— 安装程序（PyInstaller onefile 打包，--uac-admin 提权）。

打包时通过 --add-data 把以下两个文件嵌进来：
- 便携版主程序：app_payload/发票二维码工具.exe   （即「单文件运行版」）
- 卸载程序：    uninstaller.exe

运行后：选安装目录（默认 C:\Program Files\发票二维码工具）→ 复制主程序+卸载程序
→ 建开始菜单/桌面快捷方式 → 写注册表卸载项（HKLM，失败回退 HKCU）→ 可选立即运行。
"""
import os
import sys
import shutil
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox

from appicon import setup_app_id, apply_window_icon

APP_NAME = "发票二维码工具"
APP_EXE = "发票二维码工具.exe"
UNINST_EXE = "uninstaller.exe"
PROJECT_URL = "https://github.com/jianRY/invoice-qr-tool"


def resource_path(rel):
    """PyInstaller 冻结后资源在 _MEIPASS，否则用脚本所在目录。"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def default_target():
    pf = os.environ.get("ProgramFiles", r"C:\Program Files")
    return os.path.join(pf, APP_NAME)


def ps_run(script):
    """把脚本写成 UTF-8-SIG 的 .ps1 再跑，避免中文路径编码问题。"""
    import tempfile
    fd, path = tempfile.mkstemp(suffix=".ps1")
    os.close(fd)
    try:
        with open(path, "w", encoding="utf-8-sig") as f:
            f.write(script)
        r = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-File", path],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
        )
        return r.returncode == 0
    finally:
        try:
            os.remove(path)
        except Exception:
            pass


def create_shortcut(target_exe, link_path):
    script = (
        "$ws = New-Object -ComObject WScript.Shell\n"
        "$lnk = $ws.CreateShortcut('{link}')\n"
        "$lnk.TargetPath = '{tgt}'\n"
        "$lnk.WorkingDirectory = '{wd}'\n"
        "$lnk.Description = '{desc}'\n"
        "$lnk.Save()\n"
    ).format(link=link_path, tgt=target_exe, wd=os.path.dirname(target_exe), desc=APP_NAME)
    return ps_run(script)


def write_registry(target_dir, uninstall_exe):
    script = (
        "$err = $null\n"
        "try {{\n"
        "  $key = 'HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{name}'\n"
        "  New-Item -Path $key -Force | Out-Null\n"
        "  Set-ItemProperty -Path $key -Name 'DisplayName'    -Value '{name}'\n"
        "  Set-ItemProperty -Path $key -Name 'UninstallString'-Value '\"{un}\"'\n"
        "  Set-ItemProperty -Path $key -Name 'InstallLocation'-Value '{dir}'\n"
        "  Set-ItemProperty -Path $key -Name 'DisplayIcon'    -Value '{icon}'\n"
        "  Set-ItemProperty -Path $key -Name 'Publisher'      -Value 'jianRY'\n"
        "  Set-ItemProperty -Path $key -Name 'URLInfoAbout'   -Value '{url}'\n"
        "  Set-ItemProperty -Path $key -Name 'NoModify'       -Value 1\n"
        "  Set-ItemProperty -Path $key -Name 'NoRepair'       -Value 1\n"
        "}} catch {{ $err = $_\n"
        "  try {{\n"
        "    $key = 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{name}'\n"
        "    New-Item -Path $key -Force | Out-Null\n"
        "    Set-ItemProperty -Path $key -Name 'DisplayName'    -Value '{name}'\n"
        "    Set-ItemProperty -Path $key -Name 'UninstallString'-Value '\"{un}\"'\n"
        "    Set-ItemProperty -Path $key -Name 'InstallLocation'-Value '{dir}'\n"
        "  }} catch {{ }}\n"
        "}}\n"
    ).format(name=APP_NAME, un=uninstall_exe, dir=target_dir,
             icon=os.path.join(target_dir, APP_EXE), url=PROJECT_URL)
    return ps_run(script)


def do_install(target, make_desktop, run_after):
    try:
        os.makedirs(target, exist_ok=True)
        src_app = resource_path(os.path.join("app_payload", APP_EXE))
        src_un = resource_path(UNINST_EXE)
        if not os.path.exists(src_app):
            messagebox.showerror("错误", "找不到主程序文件，安装包可能已损坏。")
            return
        dst_app = os.path.join(target, APP_EXE)
        dst_un = os.path.join(target, UNINST_EXE)
        shutil.copy2(src_app, dst_app)
        if os.path.exists(src_un):
            shutil.copy2(src_un, dst_un)
        # 开始菜单快捷方式
        sm_dir = os.path.join(
            os.environ.get("ProgramData", r"C:\ProgramData"),
            "Microsoft", "Windows", "Start Menu", "Programs",
        )
        os.makedirs(sm_dir, exist_ok=True)
        create_shortcut(dst_app, os.path.join(sm_dir, APP_NAME + ".lnk"))
        # 桌面快捷方式
        if make_desktop:
            dt = os.path.join(os.environ.get("USERPROFILE", ""), "Desktop")
            if dt:
                create_shortcut(dst_app, os.path.join(dt, APP_NAME + ".lnk"))
        # 注册表卸载项
        write_registry(target, dst_un)
        messagebox.showinfo("安装完成", "{} 已安装到：\n{}".format(APP_NAME, target))
        if run_after:
            subprocess.Popen([dst_app], shell=False)
        sys.exit(0)
    except Exception as e:
        messagebox.showerror("安装失败", str(e))


def main():
    setup_app_id()          # 必须在 tk.Tk() 之前
    root = tk.Tk()
    root.title("安装 " + APP_NAME)
    root.geometry("480x220")
    root.resizable(False, False)
    apply_window_icon(root)

    tk.Label(root, text="选择安装位置：", anchor="w").pack(fill="x", padx=18, pady=(14, 2))
    frm = tk.Frame(root)
    frm.pack(fill="x", padx=18)
    path_var = tk.StringVar(value=default_target())
    ent = tk.Entry(frm, textvariable=path_var)
    ent.pack(side="left", fill="x", expand=True)
    tk.Button(frm, text="浏览...", command=lambda: _browse(path_var)).pack(side="left", padx=6)

    desk_var = tk.BooleanVar(value=True)
    tk.Checkbutton(root, text="创建桌面快捷方式", variable=desk_var).pack(
        anchor="w", padx=18, pady=(8, 2))

    def _go():
        t = path_var.get().strip()
        if not t:
            messagebox.showwarning("提示", "请选择安装目录。")
            return
        do_install(t, desk_var.get(), True)

    tk.Button(root, text="安装", width=14, height=1, command=_go).pack(pady=14)

    root.mainloop()


def _browse(var):
    d = filedialog.askdirectory(initialdir=var.get() or default_target())
    if d:
        var.set(d)


if __name__ == "__main__":
    main()
