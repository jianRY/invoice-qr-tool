#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""发票二维码工具 · 自动更新

查询 GitHub 最新 Release、下载新版本 exe、把旧版本移入回收站，并启动新版本接管。

拆出来的目的：主程序专注界面与流程编排。更新进度对话框（界面）留在主程序，
本模块只做「版本比较 → 下载 → 落盘 → 启动新版本」这条纯逻辑链。
PyInstaller 打包会沿 import 自动收集，spec 无需改动。
"""

import os
import re
import shutil
import subprocess
import sys
import tempfile
import time

import requests

from iqr_net import _UpdateCancelled, _download_file


GITHUB_REPO_OWNER = "jianRY"


GITHUB_REPO_NAME = "invoice-qr-tool"


GITHUB_LATEST_RELEASE_URL = (
    f"https://api.github.com/repos/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/releases/latest"
)

# 自有下载站（阿里云 47.116.64.26，见「下载服务器」项目）。
# 2026-09-22 起：检查更新优先读它（国内快），GitHub 只作兜底。
SITE_URL = "http://47.116.64.26:8888"
SERVER_UPDATE_JSON = SITE_URL + "/updates/qr.json"
SERVER_FILES = SITE_URL + "/files"


def _parse_version(tag: str) -> tuple:
    """把 'v3.0' / '3.0.1' / 'V3' 这类版本号解析成可比较的元组。

    只取前 3 段数字；缺失段补 0。非数字字符全部跳过。"""
    digits = re.findall(r"\d+", tag or "")
    nums = [int(x) for x in digits[:3]]
    while len(nums) < 3:
        nums.append(0)
    return tuple(nums)


def _get_json(url: str, timeout: int = 8):
    """取一份 JSON；失败一律返回 None（检查更新不能因为某个源挂了就把程序搞卡）。"""
    headers = {"User-Agent": "InvoiceQRDownloader", "Accept": "application/json"}
    resp = requests.get(url, headers=headers, timeout=timeout)
    if resp.status_code != 200:
        return None
    return resp.json()


def _parse_update_json(data):
    """把站点上的 update.json 解析成 (version, [下载源...], notes)。"""
    if not isinstance(data, dict):
        return None
    tag = str(data.get("version") or "").strip()
    urls = [u for u in (data.get("url"), data.get("fallback_url")) if u]
    if not tag or not urls:
        return None
    return _parse_version(tag), urls, (data.get("notes") or "")


def _from_github_api():
    """兜底源：GitHub API。只有它能在 update.json 拿不到时给出更新说明。"""
    headers = {"User-Agent": "InvoiceQRDownloader", "Accept": "application/vnd.github+json"}
    try:
        resp = requests.get(GITHUB_LATEST_RELEASE_URL, headers=headers, timeout=15)
    except Exception:
        return None
    if resp.status_code != 200:
        return None

    data = resp.json()
    tag = data.get("tag_name", "")
    version = _parse_version(tag)
    notes = data.get("body", "") or ""
    download_url = None

    for asset in data.get("assets", []):
        name = asset.get("name", "")
        # 只匹配 EXE 主程序附件；按版本号匹配，避免误抓旧版
        if name.lower().endswith(".exe") and re.sub(r"\W", "", name).lower().startswith(
            "invoiceqrdownloader"
        ):
            if _parse_version(name) == version:
                download_url = asset.get("browser_download_url")
                break
    if not download_url:
        # 兜底：取第一个 exe 附件
        for asset in data.get("assets", []):
            if asset.get("name", "").lower().endswith(".exe"):
                download_url = asset.get("browser_download_url")
                break

    if not download_url:
        return None
    return version, [download_url], notes


def get_latest_release():
    """查询最新版本，返回 (version, download_urls, notes) 或 None。

    download_urls 是**候选下载源列表**，按顺序尝试：
        [0] 自有服务器直链（国内快，支持 Range）
        [1] GitHub Release 直链（服务器没同步到 / 不可达时兜底）
    调用方把整个列表交给 perform_update()，由它在下载层逐个试。

    检查顺序（2026-09-22 改为双源，起因：用户反馈 GitHub 拉包慢、易超时）：
        ① 自有服务器 /updates/qr.json —— 秒回，国内直连
        ② GitHub API releases/latest —— 兜底
    发布附件仍须命名为 InvoiceQRDownloader_<版本>.exe（ASCII，避免中文名被剥离）。
    """
    try:
        got = _parse_update_json(_get_json(SERVER_UPDATE_JSON))
        if got:
            return got
    except Exception:
        pass
    return _from_github_api()


def _derive_base_name(exe_path: str) -> str:
    """从当前 exe 文件名推导「基础名」，去掉 .exe 以及末尾已有的 _vX.Y 版本段。

    例：发票二维码工具.exe -> 发票二维码工具；发票二维码工具_v3.4.exe -> 发票二维码工具。
    """
    name = os.path.basename(exe_path)
    if name.lower().endswith(".exe"):
        name = name[:-4]
    # 去掉末尾可能的 _v 版本段（如 _v3.4 / _v3.4.1）
    name = re.sub(r"_v\d+(?:\.\d+)*$", "", name, flags=re.IGNORECASE)
    return name or "发票二维码工具"


def _version_str(version: tuple) -> str:
    """把版本元组格式化成三段式字符串。 (4,7,0)->'4.7.0'；(4,6,1)->'4.6.1'。"""
    parts = [str(x) for x in version][:3]
    while len(parts) < 3:
        parts.append("0")
    return ".".join(parts)


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

        from_buf = ctypes.create_unicode_buffer(path + "\0")
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
        # 某些环境下 SHFileOperation 会返回非 0（如 fAnyOperationsAborted），
        # 但文件实际已被移入回收站 / 删除。以“原路径是否已不存在”作为最终成功判据。
        return res == 0 or not os.path.exists(path)
    except Exception:
        return False


def perform_update(download_urls, latest_version: tuple, on_event=None,
                   cancel=None) -> bool:
    """下载最新版本的 EXE，保存到当前程序同一目录（文件名带版本号），
    并把旧的 EXE 移动到回收站。

    download_urls: 候选下载源列表（取自 get_latest_release），**逐个尝试直到成功**——
        首选自有服务器直链，失败自动切 GitHub 直链；单个源网络抖动不影响整体更新。
        为兼容旧调用，也可以直接传一个字符串 URL。

    流程（cancel 只在能干净收尾的阶段生效）：
      ① 下载 → 可取消：立刻断开、删除半截 .part，等于什么都没发生；
      ② 下载完成后的落盘 / 原子移动（约 1~2 秒内）→ **不可取消**，
         此时打断可能留下损坏的 exe，故只提示「正在保存，请稍候」；
      ③ 新版本已启动 → 既成事实，无法回退。

    参数：
      cancel: threading.Event；置位表示用户点了「取消更新」。
    返回：是否成功触发替换（取消 / 失败均返回 False）。
    """
    def emit(ev):
        if on_event:
            on_event(ev)

    def _cancelled() -> bool:
        return cancel is not None and cancel.is_set()

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

    def _cleanup_part():
        """删除半截 .part（取消 / 下载失败共用；删除失败不影响主流程）。"""
        try:
            if os.path.exists(part):
                os.remove(part)
                return True
        except Exception:
            pass
        return False

    def _emit_cancelled(extra):
        _cleanup_part()
        emit({"type": "detail", "text": extra})
        emit({"type": "detail", "text": "当前版本保持不变，软件可继续正常使用。"})
        emit({"type": "cancelled"})
        emit({"type": "done", "ok": False})

    # ---- ① 下载（可取消）：多源依次尝试 ----
    urls = download_urls if isinstance(download_urls, (list, tuple)) else [download_urls]
    urls = [u for u in urls if u]
    if not urls:
        _cleanup_part()
        emit({"type": "detail", "text": "下载失败：没有可用的下载源。"})
        emit({"type": "done", "ok": False})
        return False

    for idx, url in enumerate(urls, 1):
        if _cancelled():
            _emit_cancelled("已取消更新：临时文件已清理。")
            return False
        if idx > 1:
            emit({"type": "stage", "text": f"正在切换备用下载源（{idx}/{len(urls)}）…"})
            emit({"type": "detail", "text": f"上一个源不可用，改从备用源下载：{url[:64]}"})
            _last_w[0] = 0          # 换源后进度从 0 重新计数，速度采样一并重置
            _last_t[0] = time.time()
        try:
            if _download_file(url, part, progress_cb=_prog, cancel=cancel,
                              expect_magic=b"MZ"):
                break
        except _UpdateCancelled:
            _emit_cancelled("已取消更新：下载已中断，临时文件已清理。")
            return False
        except Exception as e:
            # 单个源出错不致命，继续试下一个；全都失败才报错
            emit({"type": "detail", "text": f"下载源出错：{e}"})
    else:
        _cleanup_part()
        emit({"type": "detail",
              "text": "下载失败：所有下载源都不可用，请稍后重试或手动更新。"})
        emit({"type": "done", "ok": False})
        return False

    if _cancelled():
        # 极端情形：刚好在下载结束时点取消，同样干净收尾
        _emit_cancelled("已取消更新：临时文件已清理。")
        return False

    # ---- ② 落盘（临界区，不可取消）----
    try:
        size_mb = os.path.getsize(part) / 1048576
    except OSError:
        # 文件被安全软件 / 清理工具顺走时不要直接抛出去：抛了就没人再报 done，
        # 进度框会永远停在「下载完成」且关不掉。
        size_mb = 0.0
    emit({"type": "detail", "text": f"下载完成（{size_mb:.1f} MB），正在保存到原目录…"})
    emit({"type": "stage", "text": "下载完成，正在保存新版本（此步请稍候，无法取消）…"})
    emit({"type": "lock"})   # 通知 UI 禁用「取消更新」

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
            emit({"type": "detail",
                  "text": f"新版本文件已保存在：{new_exe}"})
            emit({"type": "done", "ok": False})
            return False
    except Exception as e:
        emit({"type": "detail", "text": f"保存新版本失败：{e}"})
        emit({"type": "done", "ok": False})
        return False


def _cleanup_legacy_update_artifacts():
    """清理早期版本（bat / --pending / 备份机制）遗留的临时文件，避免堆积。

    另外清理临时目录里的半截更新包（.part）—— 之前若更新被强杀 / 进程卡死，
    会留下上百 MB 的无用文件且再没人管它。
    """
    try:
        work_dir = os.path.dirname(sys.executable)
        for name in ("_pending.exe", "_backup.exe", "_replace_in_progress",
                     "_update_rolled_back", "invoiceqrdl_update.bat"):
            p = os.path.join(work_dir, name)
            if os.path.isfile(p):
                try:
                    os.remove(p)
                except Exception:
                    pass
    except Exception:
        pass
    # 半截更新包：正被占用（说明有更新在进行）时删不掉，静默跳过即可
    try:
        part = os.path.join(tempfile.gettempdir(),
                            "InvoiceQRDownloader_update.part")
        if os.path.isfile(part):
            os.remove(part)
    except Exception:
        pass
