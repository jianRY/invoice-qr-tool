#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""发票二维码工具 · 网络层

线程局部的 requests 会话、PDF 下载（网络类失败自动重试）、更新包下载
（带进度回调、可中途取消、校验完整性）。

拆出来的目的：主程序专注界面与流程编排。PyInstaller 打包会沿 import 自动
收集同目录模块，spec 无需改动。
"""

import os
import re
import threading
import time
from urllib.parse import urljoin

import requests



# 并发设置
#   每张图片是一个独立任务，由线程池并发执行。绝大多数耗时都在“等网络下载 PDF”上，
#   这部分 CPU 全程空转，并发收益最大（实测 16 张：串行 2.01s → 6 路 0.40s，约 5×）。
#   CPU 部分（二维码识别）也能提速（60 张 4.82s → 1.62s，约 3×），因为 zxing/OpenCV
#   是 C++ 扩展、会释放 GIL。
#   注意：「汇总发票」阶段的 pdfplumber 解析是纯 Python、被 GIL 锁死，实测并发无收益（1.0×），
#   所以那里保持串行，不要改成线程池。
DEFAULT_WORKERS = 6        # 并发路数（对同一发票平台是 6 次并发请求，过低提不上速、过高易被限流）


DOWNLOAD_RETRIES = 2       # 单张 PDF 下载失败后的重试次数（不含首次），仅对网络类错误重试


RETRY_BACKOFF = 0.6        # 重试退避基数（秒）：0.6s、1.2s 递增


class _UpdateCancelled(Exception):
    """用户主动取消更新。

    单独用一个异常类型，是为了让它能穿过 `_download_file` 的兜底 except
    （那里会把网络类异常统一当作「下载失败」返回 False，取消不能被混为一谈）。"""


def _download_file(url: str, dest: str, progress_cb=None, cancel=None,
                   expect_magic: bytes | None = None) -> bool:
    """带进度回调的下载；返回是否成功。

    cancel: threading.Event；置位则立刻断开连接，并抛出 _UpdateCancelled。
            半截文件的清理由调用方在最外层 finally 统一负责（见 perform_update），
            这样无论内部走哪条异常路径，都不会把 .part 留在磁盘上。
    expect_magic: 期望的文件头字节（如更新包传 b"MZ"）。首字节不符视为失败 ——
            防止把错误页 / 拦截页（HTTP 200 的 HTML）当更新包存下来。
    """
    headers = {"User-Agent": "InvoiceQRDownloader"}
    resp = None
    abort = False
    try:
        resp = requests.get(url, headers=headers, stream=True, timeout=60)
        resp.raise_for_status()
        total = int(resp.headers.get("Content-Length", 0)) or 0
        written = 0
        f = open(dest, "wb")
        try:
            for chunk in resp.iter_content(chunk_size=256 * 1024):
                if cancel is not None and cancel.is_set():
                    abort = True
                    raise _UpdateCancelled()
                if not chunk:
                    continue
                f.write(chunk)
                written += len(chunk)
                if progress_cb and total:
                    progress_cb(written, total)
        finally:
            try:
                f.close()
            except Exception:
                pass
        # ⚠️ 服务器提前断流时 iter_content 会「正常」结束、不抛任何异常，
        #    只看有没有异常会把半截文件当成功 —— 更新包尤其致命（装上去就废）。
        if total and written != total:
            abort = True
            return False
        if expect_magic:
            with open(dest, "rb") as fh:
                if fh.read(len(expect_magic)) != expect_magic:
                    abort = True
                    return False
        return True
    except _UpdateCancelled:
        abort = True
        raise
    except Exception:
        # 网络类异常统一视为「下载失败」，交由调用方清理 .part
        abort = True
        return False
    finally:
        if abort:
            _close_quietly(resp, abort=True)   # 立刻断开，不等连接池回收


# ---------------------------------------------------------------- 并发基础设施


class CancelToken:
    """跨线程的「停止」标志。

    主线程点「停止」时调用 cancel()；各工作线程在开始处理下一张图片前检查 cancelled，
    尚未开始的任务会被快速取消。已经发出的网络请求无法中途打断，会自然收尾（≤ 超时时间）。
    """

    def __init__(self):
        self._event = threading.Event()

    def cancel(self):
        self._event.set()

    @property
    def cancelled(self) -> bool:
        return self._event.is_set()


class PdfNotAvailable(RuntimeError):
    """网址可达，但确实拿不到 PDF（业务性失败，重试没有意义）。"""


class DownloadNetworkError(RuntimeError):
    """网络类失败（超时、连接被重置等），值得重试。"""


# 每个线程复用自己的 requests.Session（连接池复用，省掉每张发票重复的 TLS 握手）。
# requests.Session 并非线程安全，所以按线程隔离，而不是全局共享同一个。
_thread_local = threading.local()


_sessions_lock = threading.Lock()


_all_sessions: list = []


def _get_session() -> requests.Session:
    session = getattr(_thread_local, "session", None)
    if session is None:
        session = requests.Session()
        adapter = requests.adapters.HTTPAdapter(
            pool_connections=DEFAULT_WORKERS * 2,
            pool_maxsize=DEFAULT_WORKERS * 2,
            max_retries=0,  # 重试由 download_pdf 自己控制，避免双重重试放大请求量
        )
        session.mount("https://", adapter)
        session.mount("http://", adapter)
        _thread_local.session = session
        with _sessions_lock:
            _all_sessions.append(session)
    return session


def close_all_sessions():
    """任务结束后关闭各线程的 Session，释放连接。

    （关闭后的 Session 仍可继续使用，requests 会在需要时重建连接池，
    所以这里不需要额外清理线程局部变量。）
    """
    with _sessions_lock:
        sessions = list(_all_sessions)
        _all_sessions.clear()
    for s in sessions:
        try:
            s.close()
        except Exception:
            pass


def _close_quietly(resp, abort: bool = False) -> None:
    """关闭响应（连带释放底层连接），任何异常都吞掉。

    只为「把连接还给连接池」用，失败与否都不影响主流程；正常路径走 ``abort=False``。
    ``abort=True`` 用于「用户点了取消」：必须**先关原始 socket** —— requests 在响应体
    未读尽时调 close()，会试图先把剩余数据读完来「补全」连接，那一步会阻塞住，表现为
    「点了取消却迟迟不返回」；先 raw.close() 就绕开了这条补全路径。
    """
    if resp is None:
        return
    if abort:
        try:
            if getattr(resp, "raw", None) is not None:
                resp.raw.close()
        except Exception:
            pass
    try:
        resp.close()
    except Exception:
        pass


def _looks_like_pdf(response: requests.Response) -> bool:
    """判断响应体是不是 PDF。

    优先看 Content-Type，兜底再看开头 4 个字节。
    ⚠️ 这里**不能**用 `response.content`：stream=True 的响应一旦读 `.content`，
    会把整个响应体一次性拉进内存，白耗内存、也失去了流式下载的意义。
    `raw.peek()` 只取缓冲区里已有的前几个字节，不会消费 body。
    """
    if "pdf" in response.headers.get("Content-Type", "").lower():
        return True
    try:
        head = response.raw.peek(5)
    except Exception:
        return False
    return head.lstrip()[:4] == b"%PDF"


def _save_response(response: requests.Response, save_path: str) -> None:
    """流式写入目标文件。

    先写 `<目标>.part` 再 `os.replace` 原子替换：这样下载中途失败时不会在「PDF」目录里
    留下半截的损坏 PDF（并发时尤其重要，半截文件很容易被误当成功结果）。
    """
    tmp_path = save_path + ".part"
    try:
        with open(tmp_path, "wb") as f:
            for chunk in response.iter_content(chunk_size=64 * 1024):
                if chunk:
                    f.write(chunk)
        os.replace(tmp_path, save_path)
    except Exception:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass
        raise


def _download_once(url: str, save_path: str, timeout: int, session: requests.Session) -> None:
    """单次下载尝试。

    失败时区分两类，交给上层决定要不要重试：
    - DownloadNetworkError：网络类问题（超时 / 连接被重置等）→ 值得重试；
    - PdfNotAvailable   ：网址能打开但确实拿不到 PDF → 重试无意义。
    """
    from urllib.parse import urljoin

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )
    }

    resp = None
    try:
        resp = session.get(url, headers=headers, timeout=timeout, stream=True, allow_redirects=True)
        resp.raise_for_status()
    except requests.RequestException as e:
        if resp is not None:
            _close_quietly(resp)
        raise DownloadNetworkError(f"访问网址失败：{e}") from e

    # ⚠️ 用 finally 兜底 close：失败分支（网址返回的既不是 PDF、也不是带下载表单的 HTML）
    #    原先既不消费 body 也不关闭响应，连接不会回连接池 —— 6 路并发重试几轮就够把
    #    pool_maxsize 耗光，之后每个请求都要重建连接。
    try:
        if _looks_like_pdf(resp):
            _save_response(resp, save_path)
            return

        # 若是 HTML 页面，尝试找隐藏表单域并提交下载
        ct = resp.headers.get("Content-Type", "").lower()
        if "html" in ct:
            try:
                html = resp.text
            except requests.RequestException as e:
                raise DownloadNetworkError(f"读取页面内容失败：{e}") from e

            hidden_inputs = re.findall(r"<input[^>]+type=[\"']hidden[\"'][^>]*>", html, flags=re.IGNORECASE)
            fields = {}
            for tag in hidden_inputs:
                name_match = re.search("name=['\"]([^'\"]+)['\"]", tag, flags=re.IGNORECASE)
                value_match = re.search("value=['\"]([^'\"]*)['\"]", tag, flags=re.IGNORECASE)
                if name_match:
                    fields[name_match.group(1)] = value_match.group(1) if value_match else ""

            if "idBase" in fields:
                download_url = urljoin(resp.url, "/download")
                try:
                    dl_resp = session.post(
                        download_url, data=fields, headers=headers, timeout=timeout, stream=True
                    )
                    dl_resp.raise_for_status()
                except requests.RequestException as e:
                    raise DownloadNetworkError(f"提交下载接口失败：{e}") from e

                try:
                    if _looks_like_pdf(dl_resp):
                        _save_response(dl_resp, save_path)
                        return
                    body = dl_resp.text[:200]
                finally:
                    _close_quietly(dl_resp)
                raise PdfNotAvailable(f"下载接口返回的不是 PDF：{body}")

        raise PdfNotAvailable("该网址没有直接返回 PDF，也未找到可下载的隐藏表单")
    finally:
        _close_quietly(resp)


def download_pdf(
    url: str,
    save_path: str,
    timeout: int = 60,
    session: requests.Session | None = None,
    retries: int = DOWNLOAD_RETRIES,
) -> None:
    """下载 URL 指向的内容并保存为 PDF。

    支持两种常见情况：
    1. URL 直接返回 PDF 流；
    2. URL 返回发票展示页，页面里包含 name='idBase' 等隐藏域，
       此时自动提取并 POST 到 /download 获取 PDF。

    网络类失败会按 RETRY_BACKOFF 指数退避重试（默认 2 次）：一次网络抖动不该让发票
    被误判成「未下载」而要求人工重跑。「网址可达但没有 PDF」属业务性失败，不重试。
    """
    sess = session if session is not None else _get_session()
    last_err: Exception | None = None
    for attempt in range(retries + 1):
        try:
            _download_once(url, save_path, timeout, sess)
            return
        except DownloadNetworkError as e:
            last_err = e
            if attempt < retries:
                time.sleep(RETRY_BACKOFF * (2 ** attempt))
    raise last_err if last_err is not None else PdfNotAvailable("下载失败")
