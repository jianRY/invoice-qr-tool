#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""发票二维码工具 · 自动更新

查询 GitHub 最新 Release、下载新版本 exe、把旧版本移入回收站，并启动新版本接管。

拆出来的目的：主程序专注界面与流程编排。更新进度对话框（界面）留在主程序，
本模块只做「版本比较 → 下载 → 落盘 → 启动新版本」这条纯逻辑链。
PyInstaller 打包会沿 import 自动收集，spec 无需改动。
"""

import hashlib
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time

import requests

from iqr_net import _UpdateCancelled, _download_file


GITHUB_REPO_OWNER = "jianRY"


GITHUB_REPO_NAME = "invoice-qr-tool"


GITHUB_LATEST_RELEASE_URL = (
    f"https://api.github.com/repos/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/releases/latest"
)

# Release 附件里的 update.json（发版脚本上传时固定叫这个名字）——
# 带 sha256/size，是完整性校验的唯一依据，故优先于 GitHub API。
RELEASE_UPDATE_JSON = (
    "https://github.com/%s/%s/releases/latest/download/update.json"
    % (GITHUB_REPO_OWNER, GITHUB_REPO_NAME)
)

# 自有下载站（阿里云 47.116.64.26，见「下载服务器」项目）。
# 2026-09-24 定位：**最末兜底**。官网的手动下载按钮直接指向它（国内直连快），
# 自动更新则把它放在「加速镜像 → GitHub 原站」之后兜底 —— 万一镜像全挂、
# GitHub 也连不上，还能从这台服务器把更新包拉下来。
SITE_URL = "http://47.116.64.26:8888"
SERVER_FILES = SITE_URL + "/files"
SERVER_UPDATE_JSON = SITE_URL + "/updates/qr.json"

# 更新链路优先级（2026-09-24 定稿）：
#   下载：加速镜像（实测择优）→ GitHub 原站直链 → 官网自有服务器
#   查询：加速镜像上的 update.json → GitHub 原站 update.json → 自有服务器 → GitHub API
# 即「镜像优先，GitHub 与官网下载双双兜底」。

# ---------------- GitHub 加速镜像（2026-09-22 新增） ----------------
# 起因：实测裸网直连 GitHub 只有 3.7 KB/s（基本等于不可用），必须走公共加速镜像；
# 而镜像之间也差近 10 倍 —— 同一次实测：gh-proxy.com 439.6 KB/s、
# ghfast.top 222.8 KB/s、ghproxy.net 45.6 KB/s。所以下载前并发探测量一小段、
# 按实测速度择优，而不是死按固定顺序硬试（硬试会一头撞在最慢的源上干等超时）。
#
# ⚠️ 三条硬约束（改这里之前先读）：
#   ① 镜像属第三方服务，随时可能失效 —— 实测 9 个常见候选里 6 个已经死了
#      （ghproxy.cc / hub.gitmirror.com / gh.llkk.cc / github.moeyy.xyz /
#        ghproxy.cfd / hk.gh-proxy.com 全部拿不到连接）。
#      所以列表**硬编码在客户端、靠发版换源**，不写进 update.json。
#   ② 探测失败的源一律**不丢弃**，只排到最后继续尝试（探测失败 ≠ 不能下载）。
#   ③ GitHub 原站与自有服务器直链永远保留在候选里兜底（原站在前、服务器垫底）。
MIRROR_PREFIXES = (
    "https://gh-proxy.com/",
    "https://ghfast.top/",
    "https://ghproxy.net/",
)

# 能代理 **API 查询** 的镜像（与下载用的不是同一批，别混用）。
#
# ⚠️ 实测（2026-09-24）「下载」与「API 查询」的可用站点并不重合：
#     · 下载：gh-proxy.com / ghfast.top / ghproxy.net 三个都行
#     · API ：只有 gh-proxy.com 能代理 api.github.com；
#             ghfast.top 与 ghproxy.net 对 API 一律返回 403
#   好在 api.github.com 直连在国内本来就能通（约 0.8s），所以这里只放一个兜底，
#   顺序是「直连优先 → 再试镜像」（见 _api_url_candidates）。
API_MIRROR_PREFIXES = (
    "https://gh-proxy.com/",
)

PROBE_BYTES = 256 * 1024        # 探测时最多读取的字节数
PROBE_TIMEOUT = 4               # 单源探测最长等待秒数
MIN_USEFUL_SPEED = 50 * 1024    # B/s：低于此速度视为「探不到」，排最后但仍会尝试

_UA = "InvoiceQRDownloader"


def _is_github_url(url) -> bool:
    """是否为 GitHub 上的地址（只有这类才值得拼加速镜像前缀）。"""
    u = str(url or "")
    return "github.com/" in u or "githubusercontent.com/" in u


def _mirror_variants(url) -> list:
    """给一个 GitHub 直链生成全部镜像版本；非 GitHub 链接返回空列表。

    镜像用法就是「把原始完整 URL 直接拼在前缀后面」，
    raw.githubusercontent.com 与 github.com/releases/... 实测都支持。
    """
    if not _is_github_url(url):
        return []
    return [p + url for p in MIRROR_PREFIXES]


def _source_label(url) -> str:
    """给人看的源名（进度框与日志里显示）。"""
    u = str(url or "")
    for p in MIRROR_PREFIXES:
        if u.startswith(p):
            return "加速镜像 %s" % p.split("//")[1].strip("/")
    if _is_github_url(u):
        return "GitHub 原站"
    try:
        return u.split("//")[1].split("/")[0]
    except IndexError:
        return u[:30]


def _probe_speed(url, nbytes: int = PROBE_BYTES, timeout: int = PROBE_TIMEOUT) -> float:
    """拉一小段（Range）测速，返回 KB/s；失败或超时返回 0.0 —— 绝不抛异常。

    只读 256KB 就断开，几十 MB 的包不会因为测速被白下。
    """
    headers = {"User-Agent": _UA, "Range": "bytes=0-%d" % (nbytes - 1)}
    try:
        t0 = time.monotonic()
        got = 0
        resp = requests.get(url, headers=headers, stream=True, timeout=timeout)
        try:
            if resp.status_code not in (200, 206):
                return 0.0
            for chunk in resp.iter_content(chunk_size=65536):
                if not chunk:
                    break
                got += len(chunk)
                if got >= nbytes or time.monotonic() - t0 > timeout:
                    break
        finally:
            try:
                resp.close()
            except Exception:
                pass
        dt = time.monotonic() - t0
        if got <= 0 or dt <= 0:
            return 0.0
        return got / 1024.0 / dt
    except Exception:
        return 0.0


def order_download_urls(urls, timeout: int = PROBE_TIMEOUT, on_event=None) -> list:
    """把候选下载地址展开成「实测最快优先」的有序列表。

    展开规则：
      · GitHub 直链 → 生成全部镜像版本作为**加速候选**放前面，原链留作兜底
      · 非 GitHub 地址（自有服务器/官网直链）→ 一律排**最后**，只作兜底
        （2026-09-24 定稿：镜像优先 → GitHub 原站 → 官网下载兜底）
      · 去重保序

    返回顺序 = 实际尝试顺序。下载前并发探测各加速候选的速度并按快慢重排；
    探测失败的源排在最后但**不会丢弃**（详见 MIRROR_PREFIXES 上方注释②）。
    """
    def emit(ev):
        if on_event:
            on_event(ev)

    fast, gh_tail, other = [], [], []

    def _add(bucket, u):
        u = (u or "").strip()
        if u and u not in fast and u not in gh_tail and u not in other:
            bucket.append(u)

    for u in (urls or []):
        u = (u or "").strip()
        if not u:
            continue
        if _is_github_url(u):
            for m in _mirror_variants(u):
                _add(fast, m)
            _add(gh_tail, u)       # GitHub 原站：排在镜像之后、自有服务器之前
        else:
            _add(other, u)         # 自有服务器/官网直链：**永远最后兜底**

    if not fast:
        return gh_tail + other     # 没有可加速的源，直接用原顺序（不白等一次探测）

    emit({"type": "stage", "text": "正在选择最快的下载源…"})
    speeds = {}
    lock = threading.Lock()

    def _one(u):
        s = _probe_speed(u, timeout=timeout)
        with lock:
            speeds[u] = s

    threads = [threading.Thread(target=_one, args=(u,), daemon=True) for u in fast]
    for t in threads:
        t.start()
    for t in threads:
        t.join(timeout + 2)

    emit({"type": "detail", "text": "加速源测速：" + "、".join(
        "%s %.0f KB/s" % (_source_label(u), speeds.get(u, 0.0)) for u in fast)})

    usable = [u for u in fast if speeds.get(u, 0.0) >= MIN_USEFUL_SPEED / 1024.0]
    unusable = [u for u in fast if u not in usable]
    usable.sort(key=lambda u: -speeds.get(u, 0.0))
    return usable + unusable + gh_tail + other


def _meta_candidates():
    """检查更新的元数据候选源，按「GitHub 加速镜像 → GitHub 原站 → 自有服务器」排序。

    为什么 update.json 必须排在 GitHub API 之前：
        只有 update.json 带 sha256，是完整性校验的唯一依据；
        API 不返回该字段，优先走 API 等于每次都把校验跳过。

    ⚠️ update.json 是 **Release 附件**（不是仓库文件），所以它同样能用
       `releases/latest/download/` 这个稳定地址取，且能被加速镜像代理。
    ⚠️ 自有服务器的 /updates/qr.json 排**最后**兜底（2026-09-24 定稿）：
       万一镜像与原站全挂，还能从官网那台服务器拿元数据。
    """
    out, seen = [], set()

    def _add(u, label):
        u = (u or "").strip()
        if u and u not in seen:
            seen.add(u)
            out.append((u, label))

    for m in _mirror_variants(RELEASE_UPDATE_JSON):
        _add(m, _source_label(m))
    _add(RELEASE_UPDATE_JSON, _source_label(RELEASE_UPDATE_JSON))
    _add(SERVER_UPDATE_JSON, "自有服务器")
    return out


def _parse_version(tag: str) -> tuple:
    """把 'v3.0' / '3.0.1' / 'V3' 这类版本号解析成可比较的元组。

    只取前 3 段数字；缺失段补 0。非数字字符全部跳过。"""
    digits = re.findall(r"\d+", tag or "")
    nums = [int(x) for x in digits[:3]]
    while len(nums) < 3:
        nums.append(0)
    return tuple(nums)


def _get_json(url: str, timeout: int = 5):
    """取一份 JSON；失败一律返回 None（检查更新不能因为某个源挂了就把程序搞卡）。"""
    headers = {"User-Agent": "InvoiceQRDownloader", "Accept": "application/json"}
    resp = requests.get(url, headers=headers, timeout=timeout)
    if resp.status_code != 200:
        return None
    return resp.json()


def _parse_update_json(data):
    """把站点上的 update.json 解析成 (version, [下载源...], notes, sha256)。

    sha256 由发版脚本算好写进 update.json，供 perform_update 下载后校验；
    拿不到就返回空串（老版本元数据没有该字段，此时跳过校验、不影响更新）。
    """
    if not isinstance(data, dict):
        return None
    tag = str(data.get("version") or "").strip()
    urls = [u for u in (data.get("url"), data.get("fallback_url")) if u]
    if not tag or not urls:
        return None
    sha = str(data.get("sha256") or "").strip().lower()
    return _parse_version(tag), urls, (data.get("notes") or ""), sha


def _file_sha256(path: str) -> str:
    """算文件 SHA256（小写十六进制）。用于校验下载到的更新包完整。"""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


# Release 正文里写 sha256 的约定格式，由发版脚本自动追加，例如：
#     SHA256: 90f1ed78078ddbd1a1b2c3...
# 客户端解析出来作为额外校验来源（update.json 读不到时的保险）。
_SHA256_IN_BODY_RE = re.compile(r"SHA256[:\s]+([0-9a-fA-F]{64})")


def _sha256_from_notes(notes: str) -> str:
    """从 Release 正文里抠出 SHA256（小写）；没有则返回空串。"""
    m = _SHA256_IN_BODY_RE.search(notes or "")
    return m.group(1).lower() if m else ""


def _api_url_candidates():
    """API 查询候选：直连优先，再试能代理 API 的镜像。

    实测（2026-09-24）：api.github.com 直连在国内能通（约 0.8s），
    所以直连排第一；镜像只作兜底 —— 多数 gh-proxy 系镜像**不代理 API**，
    这里只放实测确认可用的那一个（见 API_MIRROR_PREFIXES）。
    """
    out, seen = [], set()
    for u in [GITHUB_LATEST_RELEASE_URL] + [p + GITHUB_LATEST_RELEASE_URL
                                            for p in API_MIRROR_PREFIXES]:
        if u not in seen:
            seen.add(u)
            out.append(u)
    return out


def _from_github_api():
    """兜底源：GitHub API。只有它能在 update.json 拿不到时给出更新说明。

    直连不通时会自动改走能代理 API 的加速镜像（见 _api_url_candidates）。
    另外会尝试从 Release 正文里解析 `SHA256: <64位>` —— 发版脚本会写进去，
    这样即便 update.json 拿不到、走的是 API 路径，也依然能校验完整性。
    """
    headers = {"User-Agent": "InvoiceQRDownloader", "Accept": "application/vnd.github+json"}
    data = None
    for u in _api_url_candidates():
        try:
            resp = requests.get(u, headers=headers, timeout=15)
            if resp.status_code == 200:
                data = resp.json()
                break
        except Exception:
            continue
    if not isinstance(data, dict):
        return None

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
    # API 本身不提供 SHA256，改为从 Release 正文里解析（发版脚本会写）；
    # 解析不到就返回空串，客户端跳过校验（不影响更新）。
    return version, [download_url], notes, _sha256_from_notes(notes)


def get_latest_release():
    """查询最新版本，返回 (version, download_urls, notes, sha256) 或 None。

    download_urls 是**候选下载源列表**，交给 perform_update() 逐个尝试；
    perform_update 下载前还会把它展开成镜像并通过测速择优（见 order_download_urls）。
    已同步的 sha256 也一并传下去做完整性校验。

    检查顺序（2026-09-24 定稿：镜像优先，GitHub 与官网下载双双兜底）：
        ① 加速镜像 + Release 附件 update.json  —— 最快，且带 sha256
        ② GitHub 原站 Release 附件 update.json  —— 镜像全挂时兜底，同样带 sha256
        ③ 自有服务器 /updates/qr.json           —— 官网那台，最后兜底
        ④ GitHub API releases/latest            —— 连服务器也挂时用（sha256 取正文）
    发布附件仍须命名为 InvoiceQRDownloader_<版本>.exe（ASCII，避免中文名被剥离）。
    """
    for url, source in _meta_candidates():
        try:
            got = _parse_update_json(_get_json(url))
            if got:
                return got
        except Exception:
            continue
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
                   cancel=None, sha256: str = "") -> bool:
    """下载最新版本的 EXE，保存到当前程序同一目录（文件名带版本号），
    并把旧的 EXE 移动到回收站。

    download_urls: 候选下载源列表（取自 get_latest_release），**逐个尝试直到成功**。
        进入下载前会展开 GitHub 加速镜像并并发测速择优，按实测速度从快到慢尝试；
        GitHub 原站与自有服务器直链留在最后兜底。
        为兼容旧调用，也可以直接传一个字符串 URL。
    sha256: 更新包期望的 SHA256（取自 update.json，小写十六进制）。非空时下载完先校验，
        不一致就当次更新失败并丢弃半截包——网络中途断流 / 缓存坏包不会再被装上。
        为空（GitHub API 兜底源、老元数据）则跳过校验。

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
    # 展开加速镜像 + 并发测速择优；择优本身出错就退回原顺序，绝不因此中断更新
    try:
        urls = order_download_urls(urls, on_event=on_event)
    except Exception as e:
        emit({"type": "detail", "text": f"选择下载源时出错，改用原顺序重试：{e}"})

    for idx, url in enumerate(urls, 1):
        if _cancelled():
            _emit_cancelled("已取消更新：临时文件已清理。")
            return False
        if idx > 1:
            emit({"type": "stage", "text": f"正在切换备用下载源（{idx}/{len(urls)}）…"})
            emit({"type": "detail",
                  "text": f"上一个源不可用，改从备用源下载：{_source_label(url)}"})
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

    # ---- ①.5 完整性校验（update.json 带 sha256 才做）----
    # 放在落盘之前：校验不过就丢弃，绝不让损坏的 exe 进入程序目录。
    if sha256:
        emit({"type": "stage", "text": "正在校验更新包完整性…"})
        try:
            actual = _file_sha256(part)
        except Exception as e:
            actual = ""
            emit({"type": "detail", "text": f"读取更新包准备校验时出错：{e}"})
        # 读不到（actual 为空）时放行，避免因偶发 IO 问题把正常更新挡掉
        if actual and actual != sha256:
            _cleanup_part()
            emit({"type": "detail",
                  "text": "更新包校验失败（SHA256 不一致），已丢弃该文件。"})
            emit({"type": "detail",
                  "text": "可能是下载中途断流或缓存了损坏的文件，请稍后重试。"})
            emit({"type": "done", "ok": False})
            return False
        if actual:
            emit({"type": "detail", "text": "更新包校验通过（SHA256 一致）。"})

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
