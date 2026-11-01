#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""iqr_update.py — 发票二维码工具 · 自动更新（共享内核适配层）

=====================================================================
⚠️ 本文件是**适配层**，不是实现。真正的逻辑在共享内核里：

    updater_core.py   纯逻辑（版本比较 / 源排序 / 元数据 / 校验 / 落位 / 清理 / 启动）

该文件由 `_shared/updater/sync_updater.py` 从唯一权威源分发而来，
**逐字节相同**。改逻辑请改权威源再同步，不要直接改这里的副本。

本文件只做三件事：
    ① 声明本项目的专属配置（仓库 / 自有站 / UA / 应用名…）
    ② 把共享内核的接口**原样转发**出去，保持旧调用方一行不改
    ③ 保留本项目独有的历史接口（用户设置、旧版回收协议…）

对外接口（全部与旧版一致，调用方无需改动）：
    GITHUB_REPO_OWNER / GITHUB_REPO_NAME / GITHUB_LATEST_RELEASE_URL
    RELEASE_UPDATE_JSON / SERVER_UPDATE_JSON / SERVER_FILES / SITE_URL
    MIRROR_PREFIXES / API_MIRROR_PREFIXES
    get_latest_release / perform_update / send_to_recycle_bin
    get_delete_old / set_delete_old / _remove_legacy_settings
    _cleanup_legacy_update_artifacts / _derive_base_name / _version_str
    _parse_version / order_download_urls / _meta_candidates …

架构说明（本项目属「库族」）：
    主程序 invoice_qr_tool.py 自带更新界面，本模块只提供纯逻辑与事件回调，
    不 import tkinter。界面与本模块之间靠事件字典通信：
        stage    阶段提示      {"type":"stage","text":...}
        detail   细节说明      {"type":"detail","text":...}
        progress 下载进度      {"type":"progress","written":..,"total":..,"speed":..}
        lock     进入不可取消的临界区（落盘）
        cancelled 用户取消
        done     结束          {"type":"done","ok":bool}
"""

import os
import re
import subprocess          # noqa: F401  —— 旧版在此 import；测试用 U.subprocess.Popen 打桩
import sys
import tempfile
import time

import updater_core as _core

# 本项目的下载器（requests 版，支持进度回调与取消）。
# ⚠️ 必须以「模块属性」形式被 perform_update 调用：测试/探针会 monkeypatch
#    本模块的 _download_file，直接 `from iqr_net import _download_file`
#    后在别处调用则会绕过补丁。
from iqr_net import _UpdateCancelled, _download_file


# ================================================================ 本项目配置

GITHUB_REPO_OWNER = "jianRY"
GITHUB_REPO_NAME = "invoice-qr-tool"

GITHUB_LATEST_RELEASE_URL = (
    "https://api.github.com/repos/%s/%s/releases/latest"
    % (GITHUB_REPO_OWNER, GITHUB_REPO_NAME)
)

# Release 附件里的 update.json（发版脚本上传时固定叫这个名字）——
# 带 sha256/size，是完整性校验的唯一依据，故优先于 GitHub API。
RELEASE_UPDATE_JSON = (
    "https://github.com/%s/%s/releases/latest/download/update.json"
    % (GITHUB_REPO_OWNER, GITHUB_REPO_NAME)
)

# 自有下载站（官网手动下载与自动更新的最末兜底源）。
# ⚠️ 客户端必须知道该地址才能下载，无法从公开源码中隐藏；
#    服务器安全靠自身加固（面板访问限制 / SSH 端口 / 防火墙），不依赖此处保密。
try:
    from _endpoints import SITE_URL as _RAW_SITE_URL   # 本机私有，.gitignore 已排除
except Exception:
    _RAW_SITE_URL = ""
# 末尾斜杠归一化：下面要拼 "/files" "/updates/qr.json"，手写带斜杠会拼出 "//"。
SITE_URL = (_RAW_SITE_URL or "").rstrip("/")

SERVER_FILES = (SITE_URL + "/files") if SITE_URL else ""
SERVER_UPDATE_JSON = (SITE_URL + "/updates/qr.json") if SITE_URL else ""

APP_NAME = "发票二维码工具"
APP_KEY = "qr"

# 共享内核配置对象。四个项目的差异全部收敛在这一处。
CFG = _core.UpdaterConfig(
    site_url=SITE_URL,
    site_meta_path="updates/%s.json" % APP_KEY,
    site_file_path="files",
    repo="%s/%s" % (GITHUB_REPO_OWNER, GITHUB_REPO_NAME),
    repo_to_app={("%s/%s" % (GITHUB_REPO_OWNER, GITHUB_REPO_NAME)).lower(): APP_KEY},
    app_name=APP_NAME,
    asset_prefix="InvoiceQRDownloader",
    user_agent="InvoiceQRDownloader",
    trust_site_meta=True,
    require_sha256=True,
    # settle_after_update 的清理是否丢后台线程（各项目历史行为不同，见内核说明）
    settle_background=False,
)


# ================================================================ 常量转发

UA = _core.UA
MIRROR_PREFIXES = _core.MIRROR_PREFIXES
API_MIRROR_PREFIXES = _core.API_MIRROR_PREFIXES
PROBE_BYTES = _core.PROBE_BYTES
PROBE_TIMEOUT = _core.PROBE_TIMEOUT
MIN_USEFUL_SPEED = _core.MIN_USEFUL_SPEED
SWITCH_SPEED = _core.SWITCH_SPEED
TRUSTED_META_HOSTS = _core.TRUSTED_META_HOSTS
INSTALLER_HINTS = _core.INSTALLER_HINTS
# 旧版模块级常量名（带下划线前缀），保留以兼容外部引用。
_INSTALLER_HINTS = INSTALLER_HINTS
# 旧版模块级正则，保留以兼容外部引用。
_SHA256_IN_BODY_RE = _core._SHA256_IN_BODY_RE
VERSION_PARTS = _core.VERSION_PARTS
CORE_VERSION = _core.CORE_VERSION
_UA = CFG.user_agent


# ================================================================ 基础工具转发

def _parse_version(tag=None):
    """把 'v3.0' / '3.0.1' / 'V3' 这类版本号解析成可比较的元组。"""
    return _core.parse_version(tag)


def parse_version(tag=None):
    """parse_version 的统一别名（与另外三个软件同名同义）。"""
    return _core.parse_version(tag)


def version_greater(remote=None, local=None):
    """远端版本是否比本地新。"""
    return _core.version_greater(remote, local)


def _strip_md(text=None, limit=4000):
    """极简 markdown 清理（弹窗显示用）。"""
    return _core.strip_md(text, limit)


def _version_str(version=None):
    """把版本元组格式化成三段式字符串。 (4,7,0)->'4.7.0'；(4,6,1)->'4.6.1'。"""
    parts = [str(x) for x in (version or ())][:3]
    while len(parts) < 3:
        parts.append("0")
    return ".".join(parts)


def _is_github_url(url=None):
    return _core.is_github_url(url)


def _mirror_variants(url=None):
    return _core.mirror_variants(url, CFG.mirrors)


def _source_label(url=None):
    return _core.source_label(url, CFG)


def _src_label(url=None):
    """_source_label 的统一别名（与另外三个软件同名同义）。"""
    return _core.source_label(url, CFG)


def _is_mirror(url=None):
    return _core.is_mirror(url, CFG.mirrors)


def _src_rank(url=None):
    return _core.src_rank(url, CFG)


def _file_sha256(path=None):
    return _core.file_sha256(path)


def _sha256_from_notes(notes=None):
    return _core.sha256_from_notes(notes)


def _probe_speed(url=None, nbytes=PROBE_BYTES, timeout=PROBE_TIMEOUT):
    return _core.probe_speed(url, nbytes=nbytes, timeout=timeout, cfg=CFG)


def _is_trusted_meta_url(url=None):
    return _core.is_trusted_meta_url(url, CFG)


def _is_downloadable(url=None):
    return _core.is_downloadable(url)


def _looks_like_exe(path=None):
    return _core.looks_like_exe(path)


def _get_json(url=None, timeout=5):
    """取一份 JSON；失败一律返回 None（检查更新不能因为某个源挂了就把程序搞卡）。"""
    try:
        return _core.http_json(url, timeout, CFG)
    except Exception:  # noqa: BLE001
        return None


def _http_json(url=None, timeout=8):
    return _get_json(url, timeout)


# ================================================================ 下载源排序


def order_download_urls(urls=None, timeout=PROBE_TIMEOUT, on_event=None):
    """把候选下载地址展开成「实测最快优先」的有序列表。

    展开规则（与另外三个软件**完全一致**）：
      · GitHub 直链 → 生成全部镜像版本作为**加速候选**放前面，原链留作兜底
      · 非 GitHub 地址（自有服务器）→ 一律排**最后**，只作兜底
      · 去重保序

    返回顺序 = 实际尝试顺序。下载前并发探测各加速候选的速度并按快慢重排；
    探测失败的源排在最后但**不会丢弃**。

    ⚠️ 事件回调沿用旧版的措辞与顺序（主程序与测试都依赖）：
       先「正在选择最快的下载源…」，再逐个列出各镜像的实测速度。
    """
    def emit(ev):
        if on_event:
            try:
                on_event(ev)
            except Exception:  # noqa: BLE001
                pass

    expanded = _core.build_download_sources(urls, CFG)
    mirrors = [u for u in expanded if _core.is_mirror(u, CFG.mirrors)]
    if not mirrors:
        # 没有可加速的源 → 直接用原顺序，不白等一次探测
        return expanded

    emit({"type": "stage", "text": "正在选择最快的下载源…"})

    def _done(speeds):
        emit({"type": "detail", "text": "加速源测速：" + "、".join(
            "%s %.0f KB/s" % (_source_label(u), speeds.get(u, 0.0)) for u in mirrors)})

    return _core.rank_sources(expanded, timeout=timeout,
                              on_probe_done=_done, cfg=CFG)


def build_download_sources(urls=None):
    return _core.build_download_sources(urls, CFG)


def probe_speed(url=None, nbytes=PROBE_BYTES, timeout=PROBE_TIMEOUT):
    return _core.probe_speed(url, nbytes=nbytes, timeout=timeout, cfg=CFG)


def rank_sources(urls=None, timeout=PROBE_TIMEOUT, on_probe_done=None):
    return _core.rank_sources(urls, timeout=timeout,
                              on_probe_done=on_probe_done, cfg=CFG)


# ================================================================ 元数据


def _meta_candidates():
    """检查更新的元数据候选源，按「可信源优先」排序。

    ⚠️⚠️ 与旧版的差异（2026-11 统一时改的，**有意为之**）：
       旧版把「加速镜像上的 update.json」排在最前，本版**不再走镜像**。
       原因：镜像同时提供 exe 地址与 sha256，镜像一旦被控就能两者一起换，
       校验等于没有 —— 另外三个软件（OCR / 法院 / 图片转表格）本来就是
       「元数据不走镜像」的策略，这里统一到同一条基线。

       镜像并没有被浪费：exe **字节**仍然从镜像下载（最快），
       只是「元数据（版本号 + sha256）」只信 GitHub 原站与自有站。

    顺序：
      ① Release 附件里的 update.json（与版本严格同步，自带 sha256/size）
      ② 自有站上的 update.json（自有站上了 HTTPS，可信）
      ③ GitHub API releases/latest（兜底；sha256 从 Release 正文抠）
    """
    return _core.meta_candidates(CFG)


def _parse_update_json(data=None):
    """把 update.json 解析成 (version, [下载源...], notes, sha256)。

    形状与旧版**完全一致**（version 是三元组），调用方无需改动。
    """
    if not isinstance(data, dict):
        return None
    ver = str(data.get("version") or "").strip()
    urls = [u for u in (data.get("url"), data.get("fallback_url")) if u]
    if not ver or not urls:
        return None
    sha = str(data.get("sha256") or "").strip().lower()
    # 自有站文件兜底：update.json 是**随 Release 公开的资产**，不能把站内地址
    # 写进去，所以线上那份文件里没有 fallback_url。只看 url，一旦 GitHub 与
    # 所有加速镜像同时不通，就会「拿到了版本号却下不到文件」—— 自有站白搭。
    extra = CFG.site_file_url(data.get("asset"))
    if extra and extra not in urls:
        urls.append(extra)
    return _parse_version(ver), urls, (data.get("notes") or ""), sha


def _api_url_candidates():
    """API 查询候选：直连优先，再试能代理 API 的镜像。"""
    return _core.api_url_candidates(CFG)


def _release_meta_url(api_url=None):
    """Release 附件里的 update.json 地址（与另外三个软件同名同义）。"""
    return CFG.release_meta_url


def _server_meta_url(api_url=None):
    """自有站上的 update.json 地址（与另外三个软件同名同义）。"""
    return CFG.site_meta_url


def _looks_like_installer(name=None):
    """安装包绝不能被当成主程序自动更新（与另外三个软件同名同义）。"""
    return _core._looks_like_installer(name)


def _from_github_api():
    """兜底源：GitHub API。只有它能在 update.json 拿不到时给出更新说明。

    sha256 从 Release 正文里解析（发版脚本会写 `SHA256: <64位>`），
    这样即便 update.json 拿不到、走的是 API 路径，也依然能校验完整性。
    """
    try:
        info = _core.fetch_latest_release(CFG, timeout=15)
    except Exception:  # noqa: BLE001
        return None
    if not info:
        return None
    return (info.parsed_version, list(info.download_urls),
            info.notes, info.sha256)


def get_latest_release():
    """查询最新版本，返回 (version, download_urls, notes, sha256) 或 None。

    download_urls 是**候选下载源列表**，交给 perform_update() 逐个尝试；
    perform_update 下载前还会把它展开成镜像并通过测速择优（见 order_download_urls）。
    已同步的 sha256 也一并传下去做完整性校验。

    检查顺序（与另外三个软件统一）：
        ① GitHub 原站 Release 附件 update.json  —— 带 sha256，与版本严格同步
        ② 自有服务器 /updates/qr.json           —— 官网那台，兜底
        ③ GitHub API releases/latest            —— 连服务器也挂时用（sha256 取正文）
    发布附件仍须命名为 InvoiceQRDownloader_<版本>.exe（ASCII，避免中文名被剥离）。
    """
    for url, _source in _meta_candidates():
        try:
            got = _parse_update_json(_get_json(url))
            if got:
                return got
        except Exception:  # noqa: BLE001
            continue
    return _from_github_api()


def fetch_latest_release():
    """统一接口：返回 dict（tag / notes / download_urls / sha256 …）或 None。

    ⚠️ 与本项目历史的 `get_latest_release()` **形状不同**：
       get_latest_release() 返回四元组 (version, urls, notes, sha256)（主程序在用，
       不能改）；本函数按另外三个软件的统一形状返回 dict。
    """
    got = get_latest_release()
    if not got:
        return None
    version, urls, notes, sha = got
    return {
        "tag": _version_str(version),
        "version": _version_str(version),
        "notes": notes,
        "download_url": urls[0] if urls else None,
        "download_urls": list(urls),
        "sha256": sha,
        "html_url": "",
        "source": "qr",
    }


def fetch_update_info(timeout=5, log_fn=None):
    """统一接口：取更新信息，返回 dict 或 None（与另外三个软件同形）。"""
    got = get_latest_release()
    if not got:
        return None
    version, urls, notes, sha = got
    if log_fn:
        try:
            log_fn("已获取版本信息：v%s" % _version_str(version))
        except Exception:  # noqa: BLE001
            pass
    return {
        "tag": _version_str(version),
        "version": _version_str(version),
        "notes": notes,
        "download_url": urls[0] if urls else None,
        "download_urls": list(urls),
        "sha256": sha,
        "html_url": "",
        "source": "qr",
    }


# ================================================================ 文件命名


def _derive_base_name(exe_path=None):
    """从当前 exe 文件名推导「基础名」，去掉 .exe 以及末尾已有的 _vX.Y 版本段。

    例：发票二维码工具.exe -> 发票二维码工具；发票二维码工具_v3.4.exe -> 发票二维码工具。
    """
    name = os.path.basename(exe_path or "")
    if name.lower().endswith(".exe"):
        name = name[:-4]
    # 去掉末尾可能的 _v 版本段（如 _v3.4 / _v3.4.1）
    name = re.sub(r"_v\d+(?:\.\d+)*$", "", name, flags=re.IGNORECASE)
    return name or APP_NAME


# ================================================================ 用户设置
# 历史：v5.4.0 曾把「更新后是否自动删除旧版本文件」持久化到程序同目录的
# settings.json。现已**彻底停用持久化**：
#   · 该选项每次都取默认值（DEFAULT_DELETE_OLD），不再读写任何配置文件；
#   · 启动时自动清理旧版本留下的 settings.json（见 _remove_legacy_settings）。
# 理由：这个选项只有两个取值，持久化的收益抵不上「程序目录多一个文件」的代价；
# 而且装到 Program Files 时该目录只读，写入本来就会静默失败。
#
# ⚠️ 路径**不能**用 os.path.dirname(sys.executable)：源码直接运行时那是 python.exe 的
#    安装目录（通常只读）。打包后 PyInstaller 会置 sys.frozen。
SETTINGS_FILE = "settings.json"
DEFAULT_DELETE_OLD = True      # 默认勾选 = 维持既有行为（更新完自动把旧版送回收站）

# 源码模式下是否允许真的走完「落盘 → 启动新版本」。
# 正常恒为 False：源码运行时 sys.executable 是 python.exe，落盘会写出
# 「python_vX.Y.Z.exe」到 Python 安装目录（2026-10-31 审计 #11）。
# 只有回归测试（outputs/_test_cancel*.py、_bug_probe5.py）会临时置 True。
ALLOW_SOURCE_UPDATE = False


def _app_dir():
    """程序所在目录：打包后是 exe 目录，源码运行时是本文件所在目录。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def settings_path():
    """历史配置文件的绝对路径。**仅用于清理，不再读写。**"""
    return os.path.join(_app_dir(), SETTINGS_FILE)


def _remove_legacy_settings():
    """删除历史版本留下的 settings.json，返回是否真的删掉了。

    静默失败：装在 Program Files 等只读目录时删不掉属正常，绝不能因此阻断启动
    或更新；也只动这一个文件，不碰同目录的其它内容。
    """
    try:
        p = settings_path()
        if os.path.isfile(p):
            os.remove(p)
            return True
    except Exception:  # noqa: BLE001
        pass
    return False


def get_delete_old(default=DEFAULT_DELETE_OLD):
    """「更新后是否自动删除旧版本文件」的当前选择。

    ⚠️ 已不再读取任何配置文件 —— 恒返回默认值（勾选）。
    """
    return bool(default)


def set_delete_old(value):
    """记录该项选择。⚠️ 已不再落盘：只清理历史遗留的 settings.json。

    保留函数签名以兼容既有调用方；选择仅在本次更新内生效，恒返回 True。
    """
    _remove_legacy_settings()
    return True


# ================================================================ 回收站


def send_to_recycle_bin(path=None):
    """把文件移动到回收站（而非彻底删除），便于误删后恢复。返回是否成功。"""
    return _core.send_to_recycle_bin(path)


# ================================================================ 配置（auto_update / cleanup_old 开关）
# ⚠️ 本项目历史上**不落盘**这两个开关（「更新后是否删除旧版本」只在本次更新内生效，
#    见 get_delete_old / set_delete_old）。下面这四个函数是**统一接口**，
#    行为与另外三个软件一致：读写 APPDATA 下的配置文件，读不到时按默认走。
#    本项目主流程不调用它们，保留只为让四个软件的接口面完全一致。

_APP_DIR = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"),
                        "InvoiceQRDownloader")
CONFIG_FILE = os.path.join(_APP_DIR, "config.json")


def _load_auto_update(config_file=None):
    return _core.load_auto_update(config_file or CONFIG_FILE)


def _save_auto_update(config_file=None, enabled=None):
    return _core.save_auto_update(config_file or CONFIG_FILE, enabled)


def _load_cleanup_old(config_file=None):
    return _core.load_cleanup_old(config_file or CONFIG_FILE)


def _save_cleanup_old(config_file=None, enabled=None):
    return _core.save_cleanup_old(config_file or CONFIG_FILE, enabled)


# ================================================================ 就地更新（统一接口）
# ⚠️ 本项目主流程走的是「--recycle-old」协议（见 perform_update / launch_replaced_version），
#    **不**用下面这套「中间名 + 标记文件 + 接管原文件名」的协议。
#    但四个软件的接口面必须一致，所以照样暴露这些名字（供上层统一调度、
#    契约测试与将来的工具化使用）。

def _is_frozen():
    return _core.is_frozen()


def _base_stem(stem=None):
    return _core.base_stem(stem)


def _old_version_pattern(base_stem=None, ext=None):
    return _core.old_version_pattern(base_stem, ext)


def _old_version_pattern_multi(prefixes=None, ext=None):
    return _core.old_version_pattern_multi(prefixes, ext)


def desired_base_name(app_name=None, version=None, fallback=None):
    return _core.desired_base_name(app_name, version, fallback)


def _purge(directory=None, current=None, log_fn=None, max_wait=8.0,
           prefixes=None, cleanup_old=True):
    return _core.purge(directory=directory, current=current, log_fn=log_fn,
                       max_wait=max_wait, prefixes=prefixes,
                       cleanup_old=cleanup_old)


def cleanup_old_versions(log_fn=None, max_wait=8.0, app_name=None, cleanup_old=None):
    # ⚠️ 把本模块的 _purge / _is_frozen 原样传进内核：宿主与测试会 monkeypatch
    #    这两个名字，内核若用自己那份，补丁就被绕过了。
    return _core.cleanup_old_versions(log_fn=log_fn, max_wait=max_wait,
                                       app_name=app_name, cleanup_old=cleanup_old,
                                       purge_fn=_purge, is_frozen_fn=_is_frozen)


def write_cleanup_flag(directory=None, cleanup_old=None):
    return _core.write_cleanup_flag(directory, cleanup_old)


def read_and_consume_cleanup_flag(directory=None):
    return _core.read_and_consume_cleanup_flag(directory)


def _park_file(directory=None, path=None, ext=None):
    return _core.park_file(directory, path, ext)


def apply_pending_rename(log_fn=None):
    return _core.apply_pending_rename(log_fn)


# 待改名状态住在共享内核里 ⇒ 状态代理，让外部读写的是同一份。
# ⚠️ 本项目走「协议 B」（_vX.Y.exe + --recycle-old），不用改名协议，
#    但保持四个适配层的接口面一致。
_self = sys.modules.get(__name__)
if _self is not None:
    _core.install_state_proxy(_self)
else:
    # 少数自定义 loader 不把模块登记进 sys.modules ⇒ 退化成「只读转发」
    def __getattr__(name):  # noqa: N807  (PEP 562)
        if name in _core.STATE_PROXY_NAMES:
            return getattr(_core, name)
        raise AttributeError(name)


def settle_after_update(log_fn=None, app_name=None, version=None, cleanup_old=None):
    """程序启动时调用：登记「退出时要改的文件名」，并清掉更新留下的旧版文件。"""
    # ⚠️ 钩子注入同 cleanup_old_versions：让宿主的 monkeypatch 能穿透到内核。
    return _core.settle_after_update(log_fn=log_fn, app_name=app_name,
                                     version=version, cleanup_old=cleanup_old,
                                     background=CFG.settle_background,
                                     purge_fn=_purge, is_frozen_fn=_is_frozen)


def windows_replace_and_restart(new_exe_path=None, log_fn=None, cleanup_old=True):
    return _core.windows_replace_and_restart(new_exe_path, log_fn=log_fn,
                                             cleanup_old=cleanup_old,
                                             app_name=APP_NAME)


# ================================================================ 更新主流程


def perform_update(download_urls, latest_version, on_event=None,
                   cancel=None, sha256="",
                   delete_old=DEFAULT_DELETE_OLD):
    """下载最新版本的 EXE，保存到当前程序同一目录（文件名带版本号），
    并把旧的 EXE 移动到回收站。

    download_urls: 候选下载源列表（取自 get_latest_release），**逐个尝试直到成功**。
        进入下载前会展开 GitHub 加速镜像并并发测速择优，按实测速度从快到慢尝试；
        GitHub 原站与自有服务器直链留在最后兜底。
        为兼容旧调用，也可以直接传一个字符串 URL。
    latest_version: 目标版本（三元组）。
    sha256: 更新包期望的 SHA256（取自 update.json，小写十六进制）。非空时下载完先校验，
        不一致就当次更新失败并丢弃半截包——网络中途断流 / 缓存坏包不会再被装上。
        ⚠️ 与另外三个软件统一：自动安装路径上**强制要求**校验值，
           拿不到就不装（宁可不更新，也不装来路不明的包）。
    delete_old: 更新成功后是否自动删除旧版本文件。True（默认，= 既有行为）把旧 exe
        送进回收站；False 则原样留在程序目录，由用户自行处理。

        实现上只是「给新进程带不带 --recycle-old 参数」—— 新进程没收到该参数就不会
        回收任何文件，因此不需要新增任何删除逻辑。老调用方不传该参数即维持原行为。

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

    def _cancelled():
        return cancel is not None and cancel.is_set()

    # 这一行挪到「源码模式」检查之后：源码运行时不该先闪一句「正在下载新版本…」
    tmp_dir = tempfile.gettempdir()
    part = os.path.join(tmp_dir, "InvoiceQRDownloader_update.part")
    try:
        if os.path.exists(part):
            os.remove(part)
    except Exception:  # noqa: BLE001
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
        except Exception:  # noqa: BLE001
            pass
        return False

    def _emit_cancelled(extra):
        _cleanup_part()
        emit({"type": "detail", "text": extra})
        emit({"type": "detail", "text": "当前版本保持不变，软件可继续正常使用。"})
        emit({"type": "cancelled"})
        emit({"type": "done", "ok": False})

    # ---- ⓿ 源码模式保护 ----
    # 未打包运行时 sys.executable 是 python.exe：继续走下面的落盘逻辑，会把更新包
    # 写成「python_v5.9.0.exe」塞进 Python 安装目录并试图启动它 —— 纯属污染。
    # 回归测试要用完整更新流程，就把 ALLOW_SOURCE_UPDATE 置 True（默认关闭）。
    if not getattr(sys, "frozen", False) and not ALLOW_SOURCE_UPDATE:
        _cleanup_part()
        emit({"type": "detail",
              "text": "当前是源码运行模式，不支持自动更新（请使用打包后的 exe）。"})
        emit({"type": "done", "ok": False})
        return False

    # ---- ⓿.5 完整性校验值缺失 → 拒绝自动更新 ----
    # 与另外三个软件统一的安全基线：自动安装这条路上，sha256 是唯一的完整性依据，
    # 缺了它就等于「装一个不知道有没有被改过的 exe」，宁可不更新。
    if CFG.require_sha256 and not str(sha256 or "").strip():
        emit({"type": "detail",
              "text": "本次更新信息缺少完整性校验值（SHA256），出于安全考虑不会自动更新。"})
        emit({"type": "detail",
              "text": "请前往 GitHub Releases 手动下载最新版本。"})
        emit({"type": "done", "ok": False})
        return False

    emit({"type": "stage", "text": "正在下载新版本…"})

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
    except Exception as e:  # noqa: BLE001
        emit({"type": "detail", "text": "选择下载源时出错，改用原顺序重试：%s" % e})

    for idx, url in enumerate(urls, 1):
        if _cancelled():
            _emit_cancelled("已取消更新：临时文件已清理。")
            return False
        if idx > 1:
            emit({"type": "stage", "text": "正在切换备用下载源（%d/%d）…" % (idx, len(urls))})
            emit({"type": "detail",
                  "text": "上一个源不可用，改从备用源下载：%s" % _source_label(url)})
            _last_w[0] = 0          # 换源后进度从 0 重新计数，速度采样一并重置
            _last_t[0] = time.time()
        try:
            if _download_file(url, part, progress_cb=_prog, cancel=cancel,
                              expect_magic=b"MZ"):
                break
        except _UpdateCancelled:
            _emit_cancelled("已取消更新：下载已中断，临时文件已清理。")
            return False
        except Exception as e:  # noqa: BLE001
            # 单个源出错不致命，继续试下一个；全都失败才报错
            emit({"type": "detail", "text": "下载源出错：%s" % e})
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

    # ---- ①.5 完整性校验 ----
    # 放在落盘之前：校验不过就丢弃，绝不让损坏的 exe 进入程序目录。
    if sha256:
        emit({"type": "stage", "text": "正在校验更新包完整性…"})
        try:
            actual = _file_sha256(part)
        except Exception as e:  # noqa: BLE001
            actual = ""
            emit({"type": "detail", "text": "读取更新包准备校验时出错：%s" % e})
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
    emit({"type": "detail", "text": "下载完成（%.1f MB），正在保存到原目录…" % size_mb})
    emit({"type": "stage", "text": "下载完成，正在保存新版本（此步请稍候，无法取消）…"})
    emit({"type": "lock"})   # 通知 UI 禁用「取消更新」

    try:
        current_exe = sys.executable  # 当前 EXE 自身路径（打包后）
        target_dir = os.path.dirname(current_exe)
        base = _derive_base_name(current_exe)
        ver = _version_str(latest_version)
        new_exe = os.path.join(target_dir, "%s_v%s.exe" % (base, ver))

        # 若同名新文件已存在（理论上应为更高版本），先移除再落入
        if os.path.exists(new_exe):
            try:
                os.remove(new_exe)
            except Exception:  # noqa: BLE001
                pass

        # 同盘原子移动（比跨盘 copy 快且不会残留半截文件）；跨盘则回退到 move
        try:
            os.replace(part, new_exe)
        except Exception:  # noqa: BLE001
            import shutil
            shutil.move(part, new_exe)

        emit({"type": "detail",
              "text": "新版本已保存为：%s" % os.path.basename(new_exe)})
        if delete_old:
            emit({"type": "detail",
                  "text": "旧版本将被移入回收站；软件即将切换到新版本…"})
        else:
            emit({"type": "detail",
                  "text": "按当前设置保留旧版本文件（未删除）：%s"
                          % os.path.basename(current_exe)})
        # ⚠️ done(ok=True) 必须等新进程**真的起来之后**再发：界面收到它就会排定
        #    「1 秒后关窗」，若启动抛错，用户看到的是「更新完软件直接消失」，
        #    而新版本其实一个都没启动（2026-10-31 审计 #2，已复现）。

        # 启动新版本并把「当前（旧）exe 路径」交给它回收；随后退出旧进程，
        # 交由新进程在旧进程释放文件后把旧 exe 移入回收站。
        # ⚠️ 启动细节（PYINSTALLER_RESET_ENVIRONMENT / 清 _MEIPASS 系变量 / cwd）
        #    已收进共享内核 launch_replaced_version()，四个软件同一份实现。
        try:
            _core.launch_replaced_version(new_exe, current_exe=current_exe,
                                          cleanup_old=delete_old)
        except Exception as e:  # noqa: BLE001
            emit({"type": "detail",
                  "text": "启动新版本失败：%s，旧版本仍保留，可手动打开新文件。" % e})
            emit({"type": "detail",
                  "text": "新版本文件已保存在：%s" % new_exe})
            emit({"type": "done", "ok": False})
            return False
        # 新进程已经拉起来了 → 到这一步才通知界面「可以关窗切新版本」
        emit({"type": "done", "ok": True})
        return True
    except Exception as e:  # noqa: BLE001
        emit({"type": "detail", "text": "保存新版本失败：%s" % e})
        emit({"type": "done", "ok": False})
        return False


# ================================================================ 遗留清理


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
                except Exception:  # noqa: BLE001
                    pass
    except Exception:  # noqa: BLE001
        pass
    # 半截更新包：正被占用（说明有更新在进行）时删不掉，静默跳过即可
    try:
        part = os.path.join(tempfile.gettempdir(),
                            "InvoiceQRDownloader_update.part")
        if os.path.isfile(part):
            os.remove(part)
    except Exception:  # noqa: BLE001
        pass
