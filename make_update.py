import io, os, sys

path = r"C:\Users\toxuj\WorkBuddy\发票处理\invoice_qr_tool.py"
with open(path, "r", encoding="utf-8") as f:
    src = f.read()

start = src.index("def perform_update")
end = src.index("class UpdateProgressDialog")

NEW_CODE = '''def _derive_base_name(exe_path: str) -> str:
    """从当前 exe 文件名推导「基础名」，去掉 .exe 以及末尾已有的 _vX.Y 版本段。

    例：发票二维码工具.exe -> 发票二维码工具；发票二维码工具_v3.4.exe -> 发票二维码工具。
    """
    name = os.path.basename(exe_path)
    if name.lower().endswith(".exe"):
        name = name[:-4]
    # 去掉末尾可能的 _v 版本段（如 _v3.4 / _v3.4.1）
    name = re.sub(r"_v\\d+(?:\\.\\d+)*$", "", name, flags=re.IGNORECASE)
    return name or "发票二维码工具"


def _version_str(version: tuple) -> str:
    """把版本元组格式化成字符串，去掉末尾多余的 0。 (3,4,0)->'3.4'；(3,5,1)->'3.5.1'。"""
    parts = [str(x) for x in version]
    while len(parts) > 1 and parts[-1] == "0":
        parts.pop()
    return ".".join(parts)


def _resource_path(relative: str) -> str:
    """获取打包后 / 开发时资源文件的绝对路径（用于图标等）。"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)


def send_to_recycle_bin(path: str) -> bool:
    """把文件移动到回收站（而非彻底删除），便于误删后恢复。返回是否成功。

    使用 Windows Shell 的 SHFileOperation 并带 FOF_ALLOWUNDO 标志，即“删除到回收站”。
    非 Windows 平台或操作失败均返回 False（调用方据此决定是否保留旧文件）。
    """
    if not os.path.exists(path):
        return True
    if sys.platform != "win32":
        try:
            os.remove(path)
            return True
        except Exception:
            return False
    try:
        import ctypes
        from ctypes import wintypes

        FO_DELETE = 0x0003
        FOF_ALLOWUNDO = 0x0040
        FOF_NOCONFIRMATION = 0x0010
        FOF_SILENT = 0x0004
        FOF_NOERRORUI = 0x0400

        class SHFILEOPSTRUCTW(ctypes.Structure):
            _fields_ = [
                ("hwnd", wintypes.HWND),
                ("wFunc", wintypes.UINT),
                ("pFrom", wintypes.LPCWSTR),
                ("pTo", wintypes.LPCWSTR),
                ("fFlags", wintypes.UINT),
                ("fAnyOperationsAborted", wintypes.BOOL),
                ("hNameMappings", wintypes.LPVOID),
                ("lpszProgressTitle", wintypes.LPCWSTR),
            ]

        from_buf = ctypes.create_unicode_buffer(path + "\\0\\0")
        op = SHFILEOPSTRUCTW()
        op.hwnd = 0
        op.wFunc = FO_DELETE
        op.pFrom = ctypes.cast(from_buf, wintypes.LPCWSTR)
        op.pTo = None
        op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT | FOF_NOERRORUI
        op.fAnyOperationsAborted = 0
        op.hNameMappings = None
        op.lpszProgressTitle = None
        res = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
        return res == 0
    except Exception:
        return False


def perform_update(download_url: str, latest_version: tuple, on_event=None) -> bool:
    """下载 GitHub 最新版本的 EXE，保存到当前程序同一目录（文件名带版本号），
    并把旧的 EXE 移动到回收站。

    流程：
      1. 下载最新 exe 到临时 .part 文件；
      2. 原子移动到「同目录/基础名_v版本.exe」；
      3. 以 --recycle-old "<当前exe路径>" 启动新版本，再由新进程回收正在运行的旧 exe
         （避免直接删除正在运行的自身文件被系统锁定）。

    on_event: 更新进度框回调，事件格式同前（stage/progress/detail/done）。
    返回是否成功触发替换。
    """
    def emit(ev):
        if on_event:
            on_event(ev)

    emit({"type": "stage", "text": "正在下载新版本…"})
    tmp_dir = tempfile.gettempdir()
    part = os.path.join(tmp_dir, "InvoiceQRDownloader_update.part")
    try:
        if os.path.exists(part):
            os.remove(part)
    except Exception:
        pass

    # 计算下载速度（基于相邻两次进度回调的时间差）
    _last_t = [time.time()]
    _last_w = [0]

    def _prog(written, total):
        now = time.time()
        dt = now - _last_t[0]
        if dt <= 0:
            dt = 0.001
        speed = (written - _last_w[0]) / dt
        _last_t[0] = now
        _last_w[0] = written
        emit({"type": "progress", "written": written, "total": total, "speed": speed})

    if not _download_file(download_url, part, progress_cb=_prog):
        emit({"type": "detail", "text": "下载失败：无法获取更新文件，请稍后重试或手动更新。"})
        emit({"type": "done", "ok": False})
        return False

    size_mb = os.path.getsize(part) / 1048576
    emit({"type": "detail", "text": f"下载完成（{size_mb:.1f} MB），正在保存到原目录…"})
    emit({"type": "stage", "text": "下载完成，正在保存新版本…"})

    try:
        current_exe = sys.executable  # 当前 EXE 自身路径（打包后）
        target_dir = os.path.dirname(current_exe)
        base = _derive_base_name(current_exe)
        ver = _version_str(latest_version)
        new_exe = os.path.join(target_dir, f"{base}_v{ver}.exe")

        # 若同名新文件已存在（理论上应为更高版本），先移除再落入
        if os.path.exists(new_exe):
            try:
                os.remove(new_exe)
            except Exception:
                pass

        # 同盘原子移动（比跨盘 copy 快且不会残留半截文件）；跨盘则回退到 move
        try:
            os.replace(part, new_exe)
        except Exception:
            import shutil
            shutil.move(part, new_exe)

        emit({"type": "detail",
              "text": f"新版本已保存为：{os.path.basename(new_exe)}"})
        emit({"type": "detail",
              "text": "旧版本将被移入回收站；软件即将切换到新版本…"})
        emit({"type": "done", "ok": True})

        # 启动新版本并把“当前（旧）exe 路径”交给它回收；随后退出旧进程，
        # 交由新进程在旧进程释放文件后把旧 exe 移入回收站。
        try:
            subprocess.Popen(
                [new_exe, "--recycle-old", current_exe],
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            return True
        except Exception as e:
            emit({"type": "detail",
                  "text": f"启动新版本失败：{e}，旧版本仍保留，可手动打开新文件。"})
            emit({"type": "done", "ok": False})
            return False
    except Exception as e:
        emit({"type": "detail", "text": f"保存新版本失败：{e}"})
        emit({"type": "done", "ok": False})
        return False


'''

new_src = src[:start] + NEW_CODE + src[end:]

# 清理遗留的回滚相关辅助函数（如仍存在）
# （本段已整体替换 perform_update 与 _post_update_self_check）

with open(path, "w", encoding="utf-8") as f:
    f.write(new_src)

print("Replaced perform_update block. New length:", len(new_src))
print("Contains _post_update_self_check:", "_post_update_self_check" in new_src)
print("Contains send_to_recycle_bin:", "send_to_recycle_bin" in new_src)
print("Contains _derive_base_name:", "_derive_base_name" in new_src)
