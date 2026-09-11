#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
发票二维码识别下载工具
功能：
1. 识别指定文件夹内图片中的二维码（全部图片均识别，不再按网址去重）；
2. 若二维码为网址，则下载对应 PDF 到同目录 PDF 文件夹下，文件名与原始图片相同；
3. 识别结果的归类与文件处理：
   - 成功识别并下载到 PDF：原图片文件名保持不变；
   - 识别到网址但无可下载的 PDF：复制一份到「未识别」文件夹，文件名前加“未下载-”；
   - 未识别到任何二维码：复制一份到「未识别」文件夹，文件名前加“未识别-”；
   - 其它情况（识别到二维码但非网址等）：复制一份到「未识别」文件夹，文件名前加“其它-”；
4. 可选任务完成后打开文件夹；
5. 可选将下载的 PDF 转换为 JPG 图片（长边 2000px，短边自适应）；
6. 可选任务完成后汇总发票（对 PDF 文件夹内发票 PDF 提取字段并生成 Excel，
   金额/统筹为纯数字、无千分位，含「是否重复」列与右侧统计区汇总）。

界面：整行扁平「▶ 开始处理」主按钮 + 自绘扁平进度条（右侧显示「进度：x / y 份」）；
「使用说明 / 检查更新」位于顶部菜单栏「帮助」。
"""

import os
import re
import sys
import glob
import time
import queue
import shutil
import subprocess
import tempfile
import datetime
import threading
from collections import Counter
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import requests

# 注意：cv2 / numpy / pymupdf / pdfplumber / openpyxl / zxingcpp 等重型依赖
# 在启动时并不需要（GUI 与更新检查仅用到 tkinter + requests）。它们改为在
# 对应功能被真正调用时才“懒加载”，可大幅缩短启动时间（窗口更快出现）。
# 懒加载由各功能函数内部 import 完成，Python 会缓存已导入模块，重复调用无额外开销。

URL_RE = re.compile(r"https?://[^\s<>\"{}|\\^`\[\]]+", re.IGNORECASE)

SUPPORTED_IMAGE_EXTS = (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp")

# 状态前缀：把“未下载 / 未识别 / 其它”的图片复制到「未识别」子文件夹时加在文件名前
PREFIX_UNRECOGNIZED = "未识别-"  # 未识别到任何二维码
PREFIX_NOT_DOWNLOADED = "未下载-"  # 识别到网址但无可下载的 PDF
PREFIX_OTHER = "其它-"           # 其它情况（识别到二维码但内容非网址等）

# 软件自身版本与 GitHub 更新源（公开仓库，更新检查无需鉴权）
__VERSION__ = "3.7.2"
GITHUB_REPO_OWNER = "jianRY"
GITHUB_REPO_NAME = "invoice-qr-tool"
GITHUB_LATEST_RELEASE_URL = (
    f"https://api.github.com/repos/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/releases/latest"
)


USAGE_TEXT = """发票二维码识别下载工具 · 使用说明
================================

【功能】
1. 识别指定文件夹内图片中的二维码（全部图片均识别，不再按网址去重）。
2. 二维码为网址：自动下载对应 PDF 到「PDF」子文件夹，文件名与原始图片相同。
3. 识别结果归类与文件处理（后三类仅做“复制”，原图片始终保留在原始文件夹中不动）：
   - 成功识别并下载到 PDF：原图片文件名保持不变。
   - 识别到网址但无可下载的 PDF：复制一份到「未识别」文件夹，文件名前加「未下载-」。
   - 未识别到任何二维码：复制一份到「未识别」文件夹，文件名前加「未识别-」。
   - 其它情况（识别到二维码但内容非网址等）：复制一份到「未识别」文件夹，文件名前加「其它-」。
4. 可选：处理完成后自动打开文件夹。
5. 可选：将下载的 PDF 转为 JPG 图片（长边 2000px，短边自适应），保存到「PDF/图片」。
6. 可选：处理完成后汇总发票（生成 Excel）。对「PDF」文件夹内所有发票 PDF，
   提取「交款人 / 票据号码 / 开票日期 / 金额合计（小写）/ 医保统筹基金支付」，
   汇总为「PDF/发票汇总_YYYYMMDD_HHMMSS.xlsx」；金额与统筹列为数字格式（纯数字、无千分位，
   可直接求和与二次计算），并新增「是否重复」列（同一票据号码出现≥2次则标记“是”）；
   表格右侧统计区新增：票据张数 / 合计总金额 / 合计总统筹金额 / 可赔付金额 /
   重复票据金额合计 / 重复票据统筹合计。
7. 处理结束后，弹出「识别结果统计」：总计识别图片数、成功下载 PDF、识别但未下载、
   未识别、其它各多少张，方便核对处理结果。
8. 日志区右上角提供「清空日志」按钮，一键清空历史日志，便于开始下一个任务。

【使用步骤】
1. 把待处理的图片放在同一个文件夹里。
2. 打开本软件，点「浏览…」选择该文件夹。
3. 按需勾选：
   - 处理完成后打开文件夹
   - 将下载的 PDF 转换为图片（JPG，长边 2000px）
   - 处理完成后汇总发票（生成 Excel）
4. 点「▶ 开始处理」（主界面整行的大按钮），在日志区查看进度；按钮下方进度条右侧
   实时显示「进度：x / y 份」。
   日志区右上角有「清空日志」，可随时清空历史日志、开始下一个任务。
   「使用说明 / 更新记录」「检查更新」「清空日志」也可在顶部菜单栏「帮助」中找到。
5. 处理完成后，PDF 在「PDF」文件夹，转换图片在「PDF/图片」，问题图片的副本在「未识别」文件夹，
   汇总表在「PDF/发票汇总_*.xlsx」。

【自动更新】
- 软件启动后会静默检查 GitHub 上的最新版本；发现新版本时弹窗提示，点「是」即自动
  下载并安装（无需手动去网页下载）。
- 也可随时通过顶部菜单栏「帮助 → 检查更新」手动检查。
- 更新源为公开仓库 jianRY/invoice-qr-tool 的 Release，无需任何账号或令牌。
- 更新过程有「更新进度」提示框：展示阶段、进度条、下载速度、已下载大小与详细日志。
- 下载完成后，新版本直接保存到当前程序同一目录下，文件名自动带版本号
  （如「发票二维码工具_v3.6.exe」）；旧的 EXE 会被移入回收站（随时可还原），
  随后自动切换到新版本，安全且无损，不再需要复杂的覆盖/备份/回滚机制。

【输出规则速查】
- 网址 + 下载成功      → PDF/<原名>.pdf，原图片文件名保持不变
- 网址 + 无 PDF 可下载 → 复制一份到「未识别/未下载-<原名>」
- 未识别到二维码       → 复制一份到「未识别/未识别-<原名>」
- 识别到二维码但非网址 → 复制一份到「未识别/其它-<原名>」
- 勾选转图             → PDF/图片/<原名>_第N页.jpg（JPG 格式，长边 2000px）
- 勾选汇总             → PDF/发票汇总_YYYYMMDD_HHMMSS.xlsx（金额/统筹为纯数字、无千分位，
                         含「是否重复」列与右侧统计区汇总）

【说明】
- 二维码识别使用 zxing-cpp，对截图 / 小二维码会自动多尺度放大，比 OpenCV 自带更稳。
- 部分发票平台（如 jsczt.cn）打开后是一个展示页，软件会自动提取页面隐藏参数并提交下载接口获取 PDF。
- 单文件 EXE，无需安装，双击即用。
"""

CHANGELOG_TEXT = """发票二维码识别下载工具 · 更新记录
================================

2026-09-11  v3.7.2
- 主按钮图标由「⬇」改为「▶」（实心右三角）：更贴合「开始处理」的执行语义，
  实心色块的视觉重量也与按钮粗体文字更协调，整体不再头重脚轻。

2026-09-11  v3.7.1
- 界面布局重新设计（按用户反馈修正 v3.7 的按钮排版）：
  ① 「开始处理」改为**整行扁平主按钮**——浅色底、细边框、居中「⬇ 开始处理」粗体文字，
     占据整个窗口宽度，点击区域更大、视觉更清爽；
  ② 进度条改为**自绘扁平样式**（浅灰轨道 + 细边框 + 蓝色填充），替掉系统默认的绿色渐变，
     与整体浅色扁平风格统一；进度条右侧新增「进度：x / y 份」实时计数；
  ③ 「使用说明 / 检查更新」从主界面移除，移入**顶部菜单栏「帮助」**，
     主界面更简洁，把注意力集中在「开始处理」上（「清空日志」在菜单栏与日志区均可使用）。

2026-09-11  v3.7
- 主界面按钮重排：「开始处理」改为醒目的大号主按钮（蓝色底、加粗字体、更大点击区域），
  位于左侧；「使用说明 / 检查更新」改为右侧统一尺寸的次要小按钮；进度条独占一行、
  横贯窗口，操作更直观、更容易点。
- PDF 转图片的格式由 PNG 改为 JPG（质量 92）：文件体积明显更小，便于上传、分享与打印；
  输出位置不变（PDF/图片/<原名>_第N页.jpg）。
- 日志区新增「清空日志」按钮（位于“处理日志”右侧），一键清空历史日志，
  便于开始下一个任务前保持界面清爽。
- 汇总 Excel 的金额去掉千分位：金额合计（小写）、医保统筹基金支付及统计区数值
  由 “1,234.56” 改为 “1234.56”，避免复制到其它表格或做二次计算时出错。
- 修复汇总 Excel 两处公式错误：①合计行的「金额合计 / 统筹合计」原先写到了相邻一列
  （分别落到统筹列与「是否重复」列），现更正为对应「金额合计（小写）」「医保统筹基金支付」列；
  ②右侧「可赔付金额」原先引用的是文字标签列（会显示 #VALUE! 错误），
  现改为引用数值列（=合计总金额 − 合计总统筹金额），计算恢复正常。

2026-08-19  v3.6
- 修复金额识别 bug：金额中的千位分隔符（逗号，如 “1,234.56”）此前会被当作结束符，
  只识别到 “1” 导致金额残缺；现已支持中英文逗号，完整识别到小数点后两位。
  另新增「金额合计（大写）」兜底：小写金额解析失败时用大写金额补回，避免漏识别。
- 重构图片识别与文件处理：去除「重复网址」判定，全部图片均识别；
  成功并下载 PDF 的图片文件名保持不变；未下载 / 未识别 / 其它三类仅“复制”一份到
  新建的「未识别」子文件夹，并分别加前缀「未下载- / 未识别- / 其它-」，原图片保留不动。
- 汇总 Excel 增强：新增「是否重复」列（同一票据号码出现≥2次标记“是”）；
  右侧统计区新增「重复票据金额合计 / 重复票据统筹合计」（仅对标记为“是”的票据求和）。

2026-08-17  v3.5.1
- 修复汇总 Excel 的 bug：金额合计（小写）、医保统筹基金支付两列之前以文本格式写入，
  导致表格下方的 SUM 合计无法计算（求和为 0）。现已统一转为数字格式（保留两位小数），
  合计行与统计区均可正常求和。
- 新增右侧统计区：在表格右侧列出四项汇总指标——
  票据张数（=COUNTA 统计发票总数）、合计总金额（=SUM 金额列）、合计总统筹金额（=SUM 统筹列）、
  可赔付金额（=合计总金额 − 合计总统筹金额），便于一眼核对整体数据。

2026-08-14  v3.5
- 更新下载逻辑重构：直接下载 GitHub 最新 Release 的 EXE，保存到当前程序同一目录，
  文件名自动带版本号（如「发票二维码工具_v3.5.exe」）；旧版本 EXE 移到回收站（可还原），
  随后自动切换到新版本。不再依赖 bat 覆盖 / 备份 / 回滚那套复杂机制，更稳更简单。
- 更换软件图标：内置养眼的发票 + 二维码主题图标（已嵌入 EXE 与窗口标题）。

2026-08-14  v3.4
- 新增「识别结果统计」：处理结束后弹出统计框，并写入日志，分项统计
  总计识别图片、成功下载 PDF、识别但未下载 PDF、纯数字忽略、识别失败（未识别）各多少张，
  重复网址额外单列。
- 成功识别并下载到 PDF 的图片，文件名前自动加「1」前缀（如「1发票001.jpg」），
  与「未下载 / 未识别- / 重复」等状态前缀一致，便于在文件夹中一眼区分已完成项。

2026-08-12  v3.3
- 修复「更新后未能真正替换原文件」的致命 bug：更新脚本 bat 之前以 UTF-8（无 BOM）写入，
  含中文的 exe 路径被 cmd 按 GBK 误读，导致 del/copy/start 全部失败、文件从未被替换；
  同时成功弹窗阻塞导致旧进程未退出就删除文件引发竞态。现已改用 utf-8-sig（带 BOM）写入
  bat，并以「轮询删除直到文件可删」替代固定等待。
- 新增「更新进度」提示框：实时展示更新阶段、进度条、下载速度、已下载大小与详细日志，
  下载完成后自动替换并重启，不再有阻塞式弹窗。

2026-08-12  v3.2
- 启动速度优化：将 cv2 / numpy / pymupdf / pdfplumber / openpyxl / zxingcpp 等重型依赖
  改为「懒加载」（仅在实际使用相关功能时才导入），GUI 与更新检查不再在启动时初始化它们，
  窗口出现更快。
- 更新回滚增加用户可见提示：若自动更新因新版本启动校验失败被回滚，重启后弹出
  「更新已回滚」警告框，说明已还原到可用版本，不再静默。
- 更新校验时机后移：在 GUI 成功构建之后才确认更新成功，进一步避免把界面初始化
  失败的版本误判为成功。

2026-08-12  v3.1
- 自动更新新增「失败自动回滚」：替换前备份当前程序为 _backup.exe；
  新版本以 --pending 自启并完成启动校验，校验失败则由更新脚本自动还原旧版并重启，
  避免更新把软件弄成打不开的状态。

2026-08-12  v3.0
- 新增「依托 GitHub 的自动更新」功能：软件启动后静默检查最新版本，
  发现新版本可一键下载并自动替换重启（也可点「检查更新」手动触发）。
- 更新源为公开仓库 jianRY/invoice-qr-tool 的 Release，全程无需账号或令牌。
- 主界面标题显示当前版本号，并新增「检查更新」按钮。

2026-08-11  v2.0
- 集成「汇总发票」功能（来自“汇总票据”会话的 invoice_summary.py）：
  处理完成后可将「PDF」文件夹内发票 PDF 提取关键字段并汇总成 Excel。
- 主页面新增开关：「处理完成后汇总发票（生成 Excel）」。
- 提取字段：文件名 / 交款人 / 票据号码 / 开票日期 / 金额合计（小写）/ 医保统筹基金支付；
  输出「PDF/发票汇总_YYYYMMDD_HHMMSS.xlsx」，含金额与医保统筹合计行。
- 新增依赖 pdfplumber（PDF 文字抽取）+ openpyxl（Excel 写入），已纳入单文件 EXE。

2026-08-11  v1.1
- 新增：二维码为网址但无可下载 PDF 时，图片重命名为「未下载<原文件名>」。
- 新增：检测到重复网址时，对应图片文件名前加「重复」标识（可与「未下载」叠加）。
- 新增：内置「使用说明」窗口与「更新记录」窗口，
  并随软件附带 使用说明.txt、更新日志.txt。

2026-08-11  v1.0
- 初版发布：二维码识别、网址下载 PDF、纯数字忽略、未识别重命名、
  可选打开文件夹、可选 PDF 转图片。
"""



def _parse_version(tag: str) -> tuple:
    """把 'v3.0' / '3.0.1' / 'V3' 这类版本号解析成可比较的元组。

    只取前 3 段数字；缺失段补 0。非数字字符全部跳过。"""
    digits = re.findall(r"\d+", tag or "")
    nums = [int(x) for x in digits[:3]]
    while len(nums) < 3:
        nums.append(0)
    return tuple(nums)


def get_latest_release():
    """查询 GitHub 最新 Release，返回 (version, download_url, notes) 或 None。

    仅读取公开仓库的 Release 列表，无需任何鉴权 token。
    发布附件必须命名为 InvoiceQRDownloader_<版本>.exe（ASCII，避免中文名被剥离）。
    """
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
    return version, download_url, notes


def _download_file(url: str, dest: str, progress_cb=None) -> bool:
    """带进度回调的下载；返回是否成功。"""
    headers = {"User-Agent": "InvoiceQRDownloader"}
    try:
        resp = requests.get(url, headers=headers, stream=True, timeout=60)
        resp.raise_for_status()
        total = int(resp.headers.get("Content-Length", 0)) or 0
        written = 0
        with open(dest, "wb") as f:
            for chunk in resp.iter_content(chunk_size=256 * 1024):
                if not chunk:
                    continue
                f.write(chunk)
                written += len(chunk)
                if progress_cb and total:
                    progress_cb(written, total)
        return True
    except Exception:
        return False


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


class UpdateProgressDialog:
    """更新进度框：展示阶段、进度条、下载速度、已下载大小与详细日志。"""

    def __init__(self, parent):
        self.parent = parent
        self.queue = queue.Queue()
        self._closed = False
        self.win = tk.Toplevel(parent)
        self.win.title("更新进度")
        self.win.geometry("480x380")
        self.win.resizable(False, False)
        try:
            self.win.transient(parent)
            self.win.grab_set()
            # 更新进行中禁止手动关闭，避免中断替换
            self.win.protocol("WM_DELETE_WINDOW", lambda: None)
        except Exception:
            pass
        self._build_widgets()
        self._poll()

    def _build_widgets(self):
        pad = {"padx": 12, "pady": 6}
        self.stage_var = tk.StringVar(value="准备中…")
        ttk.Label(
            self.win, textvariable=self.stage_var,
            font=("Microsoft YaHei", 11, "bold"),
        ).pack(anchor=tk.W, **pad)

        self.bar = ttk.Progressbar(self.win, mode="determinate", maximum=100)
        self.bar.pack(fill=tk.X, padx=12, pady=(0, 6))

        row = ttk.Frame(self.win)
        row.pack(fill=tk.X, padx=12, pady=(0, 6))
        self.pct_var = tk.StringVar(value="0%")
        self.size_var = tk.StringVar(value="0.0 / 0.0 MB")
        self.speed_var = tk.StringVar(value="— KB/s")
        ttk.Label(row, textvariable=self.pct_var, width=10).pack(side=tk.LEFT)
        ttk.Label(row, textvariable=self.size_var, width=22).pack(side=tk.LEFT)
        ttk.Label(row, textvariable=self.speed_var, width=16).pack(side=tk.LEFT)

        ttk.Label(self.win, text="详细进度：").pack(anchor=tk.W, padx=12)
        self.txt = scrolledtext.ScrolledText(
            self.win, wrap=tk.WORD, state=tk.DISABLED, height=11
        )
        self.txt.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 8))

        self.btn_close = ttk.Button(
            self.win, text="关闭", command=self._close, state=tk.DISABLED
        )
        self.btn_close.pack(pady=(0, 10))

    def emit(self, **kw):
        self.queue.put(kw)

    def _close(self):
        if self._closed:
            return
        self._closed = True
        try:
            self.win.destroy()
        except Exception:
            pass

    def _poll(self):
        try:
            while True:
                ev = self.queue.get_nowait()
                self._apply(ev)
        except queue.Empty:
            pass
        if not self._closed:
            self.win.after(80, self._poll)

    def _apply(self, ev):
        t = ev.get("type")
        if t == "stage":
            self.stage_var.set(ev.get("text", ""))
            self._append(ev.get("text", ""))
        elif t == "progress":
            w, tot = ev.get("written", 0), ev.get("total", 0)
            pct = (w / tot * 100) if tot else 0
            self.bar["value"] = pct
            self.pct_var.set(f"{pct:.1f}%")
            self.size_var.set(f"{w / 1048576:.1f} / {tot / 1048576:.1f} MB")
            sp = ev.get("speed", 0) or 0
            if sp >= 1048576:
                self.speed_var.set(f"{sp / 1048576:.2f} MB/s")
            else:
                self.speed_var.set(f"{sp / 1024:.1f} KB/s")
        elif t == "detail":
            self._append(ev.get("text", ""))
        elif t == "done":
            ok = ev.get("ok", False)
            self.stage_var.set("更新完成，正在切换到新版本…" if ok else "更新失败")
            self._append("更新完成，即将切换到新版本。" if ok else "更新失败，请重试或手动更新。")
            self.btn_close.configure(state=tk.NORMAL)
            self.win.protocol("WM_DELETE_WINDOW", self._close)

    def _append(self, text):
        self.txt.configure(state=tk.NORMAL)
        self.txt.insert(tk.END, text + "\n")
        self.txt.see(tk.END)
        self.txt.configure(state=tk.DISABLED)


def _start_update_flow(root: tk.Tk, download_url: str, version: tuple):
    """在进度框中执行更新；成功后短暂展示“更新完成”再关闭主程序，由新版本接管。"""
    dlg = UpdateProgressDialog(root)

    def _worker():
        def on_event(ev):
            dlg.emit(**ev)
        ok = perform_update(download_url, version, on_event=on_event)
        if ok:
            # 让进度框显示“更新完成”约 1 秒，再关闭主程序交给新版本接管
            root.after(1000, lambda: (dlg._close(), root.destroy()))

    threading.Thread(target=_worker, daemon=True).start()


def check_and_prompt_update(root: tk.Tk):
    """后台检查更新，若有新版本则弹窗询问是否更新。供 UI 按钮调用。"""
    result = get_latest_release()
    if not result:
        root.after(0, lambda: messagebox.showinfo(
            "检查更新", "暂时无法连接到更新服务器（或当前已是最新）。"))
        return
    version, download_url, notes = result
    if version <= _parse_version(__VERSION__):
        root.after(0, lambda: messagebox.showinfo(
            "检查更新", f"当前已是最新版本 v{__VERSION__}。"))
        return

    note_text = notes.strip() or "（无更新说明）"

    def _ask():
        ask = messagebox.askyesno(
            "发现新版本",
            f"发现新版本 v{'.'.join(map(str, version))}，当前为 v{__VERSION__}。\n\n"
            f"更新内容：\n{note_text[:600]}\n\n是否立即下载并更新？",
        )
        if ask:
            _start_update_flow(root, download_url, version)

    root.after(0, _ask)


def _silent_startup_check(root: tk.Tk):
    """启动后静默检查更新；发现新版本且用户确认则更新。"""
    try:
        result = get_latest_release()
        if not result:
            return
        version, download_url, notes = result
        if version <= _parse_version(__VERSION__):
            return

        def _ask():
            note_text = notes.strip() or "（无更新说明）"
            ok = messagebox.askyesno(
                "发现新版本",
                f"发现新版本 v{'.'.join(map(str, version))}，当前为 v{__VERSION__}。\n\n"
                f"更新内容：\n{note_text[:600]}\n\n是否立即下载并更新？",
            )
            if ok:
                _start_update_flow(root, download_url, version)

        root.after(0, _ask)
    except Exception:
        pass


def detect_qr_codes(image_path: str):
    """返回图片中识别到的所有二维码文本列表（去重）"""
    # 懒加载：二维码识别相关的重型库仅在真正识别时才导入，缩短启动时间
    import cv2
    import numpy as np

    try:
        import zxingcpp
    except Exception:
        zxingcpp = None

    # 使用 numpy 读字节再 imdecode，避免 OpenCV 在 Windows 上处理中文路径的编码问题
    img = cv2.imdecode(np.fromfile(image_path, dtype=np.uint8), cv2.IMREAD_COLOR)
    if img is None:
        return []

    codes = []

    # 1. zxing-cpp 识别能力更强，先尝试；对小二维码会自动多尺度放大重试
    if zxingcpp is not None:
        h, w = img.shape[:2]
        max_dim = max(h, w)
        if max_dim < 800:
            scales = [1, 2, 3, 4]
        elif max_dim < 1600:
            scales = [1, 2, 3]
        else:
            scales = [1, 2]

        for scale in scales:
            if scale == 1:
                scaled = img
            else:
                new_w = int(w * scale)
                new_h = int(h * scale)
                scaled = cv2.resize(img, (new_w, new_h), interpolation=cv2.INTER_CUBIC)
            try:
                results = zxingcpp.read_barcodes(scaled)
                for r in results:
                    # 只保留 QR 码（zxing-cpp 不同版本枚举名可能不同，这里同时兼容两种写法）
                    qr_format = getattr(zxingcpp.BarcodeFormat, "QRCode", None) or getattr(
                        zxingcpp.BarcodeFormat, "QR_CODE", None
                    )
                    if qr_format and r.format == qr_format:
                        text = r.text.strip()
                        if text and text not in codes:
                            codes.append(text)
            except Exception:
                pass
            if codes:
                break

    # 2. OpenCV 自带 QRCodeDetector 兜底
    if not codes:
        try:
            detector = cv2.QRCodeDetector()
            result = detector.detectAndDecodeMulti(img)
            if result and result[0] and result[1]:
                for info in result[1]:
                    if info and info.strip() not in codes:
                        codes.append(info.strip())
        except Exception:
            pass

    if not codes:
        try:
            detector = cv2.QRCodeDetector()
            data, _, _ = detector.detectAndDecode(img)
            if data and data.strip() not in codes:
                codes.append(data.strip())
        except Exception:
            pass

    return codes


def is_url(text: str) -> str | None:
    """若文本包含 URL，返回提取到的完整 URL；否则返回 None"""
    match = URL_RE.search(text)
    return match.group(0) if match else None


def _looks_like_pdf(response: requests.Response) -> bool:
    ct = response.headers.get("Content-Type", "").lower()
    return "pdf" in ct or response.content.startswith(b"%PDF")


def _save_response(response: requests.Response, save_path: str) -> None:
    with open(save_path, "wb") as f:
        for chunk in response.iter_content(chunk_size=64 * 1024):
            if chunk:
                f.write(chunk)


def download_pdf(url: str, save_path: str, timeout: int = 60) -> None:
    """下载 URL 指向的内容并保存为 PDF。

    支持两种常见情况：
    1. URL 直接返回 PDF 流；
    2. URL 返回发票展示页，页面里包含 name='idBase' 等隐藏域，
       此时自动提取并 POST 到 /download 获取 PDF。
    """
    from urllib.parse import urljoin

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )
    }

    session = requests.Session()
    resp = session.get(url, headers=headers, timeout=timeout, stream=True, allow_redirects=True)
    resp.raise_for_status()

    if _looks_like_pdf(resp):
        _save_response(resp, save_path)
        return

    # 若是 HTML 页面，尝试找隐藏表单域并提交下载
    ct = resp.headers.get("Content-Type", "").lower()
    if "html" in ct:
        hidden_inputs = re.findall(r"<input[^>]+type=[\"']hidden[\"'][^>]*>", resp.text, flags=re.IGNORECASE)
        fields = {}
        for tag in hidden_inputs:
            name_match = re.search("name=['\"]([^'\"]+)['\"]", tag, flags=re.IGNORECASE)
            value_match = re.search("value=['\"]([^'\"]*)['\"]", tag, flags=re.IGNORECASE)
            if name_match:
                fields[name_match.group(1)] = value_match.group(1) if value_match else ""

        if "idBase" in fields:
            download_url = urljoin(resp.url, "/download")
            dl_resp = session.post(download_url, data=fields, headers=headers, timeout=timeout, stream=True)
            dl_resp.raise_for_status()
            if _looks_like_pdf(dl_resp):
                _save_response(dl_resp, save_path)
                return
            raise RuntimeError(f"下载接口返回的不是 PDF：{dl_resp.text[:200]}")

    raise RuntimeError("该网址没有直接返回 PDF，也未找到可下载的隐藏表单")


def convert_pdf_to_images(pdf_path: str, out_dir: str, base_name: str) -> list[str]:
    """将 PDF 每一页渲染为 JPG 图片，长边 2000px，短边自适应；返回生成的文件路径列表"""
    import pymupdf  # 懒加载：仅在转图时才需要

    generated = []
    doc = pymupdf.open(pdf_path)
    try:
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            rect = page.rect
            w, h = rect.width, rect.height
            scale = 2000.0 / max(w, h)
            mat = pymupdf.Matrix(scale, scale)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            out_name = f"{base_name}_第{page_num + 1}页.jpg"
            out_path = os.path.join(out_dir, out_name)
            # 以 JPEG 输出（质量 92）：体积远小于 PNG，便于上传与分享。
            # alpha=False 已保证无透明通道，符合 JPEG 要求。
            try:
                pix.save(out_path, jpg_quality=92)
            except TypeError:
                # 兼容不支持 jpg_quality 参数的旧版 PyMuPDF
                with open(out_path, "wb") as fh:
                    fh.write(pix.tobytes("jpeg"))
            generated.append(out_path)
    finally:
        doc.close()
    return generated


# =====================================================================
# 发票汇总模块（来自“汇总票据”会话的 invoice_summary.py，集成到此工具）
# 针对江苏省医疗门诊收费票据 / 电子票据，提取关键字段并汇总到 Excel。
# =====================================================================

def _norm_colon(s: str) -> str:
    """把全角冒号统一成半角，便于匹配标签。"""
    return s.replace("：", ":")


def _extract_field(words, label, gather_line=False):
    """在单页 words 中查找包含 label 的词，返回其后的取值。"""
    labeln = _norm_colon(label)
    for w in words:
        tn = _norm_colon(w["text"])
        idx = tn.find(labeln)
        if idx != -1:
            val = tn[idx + len(labeln):].strip()
            if gather_line:
                line_words = [x for x in words
                              if abs(x["top"] - w["top"]) < 3 and x["x0"] > w["x0"]]
                line_words.sort(key=lambda x: x["x0"])
                val = (val + "".join(x["text"] for x in line_words)).strip()
            return val
    return None


def _field_search(pages_words, candidates):
    """跨页查找字段。candidates: [(label, gather_line), ...]。"""
    for words in pages_words:
        for label, gather in candidates:
            v = _extract_field(words, label, gather)
            if v:
                return v
    return None


def _normalize_date(s):
    if not s:
        return ""
    s = s.strip()
    m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return s


def _extract_number(s):
    if not s:
        return None
    txt = str(s).replace(",", "").replace("，", "")  # 去除千分位逗号（中英文）
    m = re.search(r"-?\d+(?:\.\d+)?", txt)
    return m.group(0) if m else None


def _to_num(s):
    """把金额/统筹字符串转成 float（写入 Excel 后即为数字，可被 SUM 计算）。

    支持千位分隔符（英文逗号与中文逗号），例如 "1,234.56" / "1，234.56"
    会正确解析为 1234.56（之前遇到逗号会把金额截断，只识别到 "1"）。
    返回 float 或 None（缺失/无法解析时）。数字一律非负，便于求和统计。
    """
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return float(s)
    txt = str(s).replace(",", "").replace("，", "")  # 去除千分位逗号（中英文）
    m = re.search(r"-?\d+(?:\.\d+)?", txt)
    if not m:
        return None
    try:
        return float(m.group(0))
    except ValueError:
        return None


# ---- 中文大写金额解析（小写解析失败时的兜底/校验）----
_CN_DIGITS = {
    "零": 0, "〇": 0, "一": 1, "壹": 1, "二": 2, "贰": 2, "两": 2,
    "三": 3, "叁": 3, "四": 4, "肆": 4, "五": 5, "伍": 5, "六": 6,
    "陆": 6, "七": 7, "柒": 7, "八": 8, "捌": 8, "九": 9, "玖": 9,
}
_CN_UNITS = {
    "十": 10, "拾": 10, "百": 100, "佰": 100, "千": 1000, "仟": 1000,
    "万": 10000, "萬": 10000, "亿": 100000000,
}


def _cn_section_to_num(s: str) -> int:
    """解析不含「万/亿」的中文数字段为整数。"""
    total = 0
    section = 0
    number = 0
    for ch in s:
        if ch in _CN_DIGITS:
            number = _CN_DIGITS[ch]
        elif ch in _CN_UNITS:
            unit = _CN_UNITS[ch]
            if unit >= 10000:
                section = (section + number) * unit
                total += section
                section = 0
                number = 0
            else:
                section += number * unit
                number = 0
    total += section + number
    return total


def _cn_capital_to_num(text):
    """把「金额合计（大写）」的中文大写金额解析为 float。
    仅用作小写金额解析失败时的兜底；无法解析返回 None。"""
    if not text:
        return None
    t = text.replace(" ", "").replace("整", "").replace("正", "")
    int_part = t
    dec_part = ""
    m = re.search(r"[元圆]", t)
    if m:
        int_part = t[:m.start()]
        dec_part = t[m.end():]
    dec_val = 0.0
    jiao = re.search(r"([零壹贰叁肆伍陆柒捌玖一二三四五六七八九])角", dec_part)
    if jiao:
        dec_val += _CN_DIGITS.get(jiao.group(1), 0) * 0.1
    fen = re.search(r"([零壹贰叁肆伍陆柒捌玖一二三四五六七八九])分", dec_part)
    if fen:
        dec_val += _CN_DIGITS.get(fen.group(1), 0) * 0.01
    if dec_val == 0.0 and dec_part:
        dm = re.search(r"\d+", dec_part)
        if dm:
            dec_val = float(dm.group(0)) / 100.0
    int_val = _cn_section_to_num(int_part)
    if int_val == 0 and dec_val == 0.0:
        return None
    return int_val + dec_val


def parse_invoice(pdf_path):
    """解析单个 PDF，返回 (字段字典, 错误信息或 None)。"""
    import pdfplumber  # 懒加载：仅在汇总时才需要

    fields = {
        "交款人": None,
        "票据号码": None,
        "开票日期": None,
        "金额合计（小写）": None,
        "医保统筹基金支付": None,
    }
    try:
        with pdfplumber.open(pdf_path) as pdf:
            pages_words = [page.extract_words() for page in pdf.pages]
    except Exception as e:
        return fields, f"打开失败: {e}"

    fields["交款人"] = _field_search(pages_words, [("交款人：", False)])
    fields["票据号码"] = _field_search(
        pages_words, [("票据号码：", False), ("所属电子票据号码:", False)])
    raw_date = _field_search(pages_words, [("开票日期：", True)])
    fields["开票日期"] = _normalize_date(raw_date) if raw_date else None
    raw_amount = _field_search(pages_words, [("小写", False)])
    lower = _to_num(raw_amount) if raw_amount else None
    # 小写解析失败时，用「金额合计（大写）」兜底，避免金额漏识别
    if lower is None:
        raw_upper = _field_search(pages_words, [("大写", False)])
        upper = _cn_capital_to_num(raw_upper) if raw_upper else None
        if upper is not None:
            lower = upper
    fields["金额合计（小写）"] = lower
    fields["医保统筹基金支付"] = _to_num(
        _field_search(pages_words, [("医保统筹基金支付：", False)]))

    return fields, None


_SUMMARY_HEADERS = ["文件名", "交款人", "票据号码", "开票日期",
                    "金额合计（小写）", "医保统筹基金支付", "是否重复"]


def _write_summary_excel(rows, out_path):
    # 懒加载 openpyxl：仅在真正写 Excel 时才导入，缩短启动时间
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

    _SUMMARY_FILL = PatternFill("solid", fgColor="1F4E78")
    _SUMMARY_FONT = Font(bold=True, color="FFFFFF", size=11)
    _SUMMARY_THIN = Side(style="thin", color="BFBFBF")
    _SUMMARY_BORDER = Border(left=_SUMMARY_THIN, right=_SUMMARY_THIN,
                              top=_SUMMARY_THIN, bottom=_SUMMARY_THIN)
    _SUMMARY_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    # 金额显示格式：不用千分位分隔（避免外部工具解析 / 二次计算出错），仅保留两位小数
    _MONEY_FMT = "0.00"

    wb = Workbook()
    ws = wb.active
    ws.title = "发票汇总"

    ws.append(_SUMMARY_HEADERS)
    for c in range(1, len(_SUMMARY_HEADERS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = _SUMMARY_FILL
        cell.font = _SUMMARY_FONT
        cell.alignment = _SUMMARY_CENTER
        cell.border = _SUMMARY_BORDER

    # 先按「票据号码」(第 3 列, 索引 2) 统计重复：非空且出现 >=2 次标记为“是”
    _TICKET_IDX = 2
    _ticket_counts = Counter()
    for r in rows:
        tn = r[_TICKET_IDX]
        if tn:
            _ticket_counts[tn] += 1

    for r in rows:
        tn = r[_TICKET_IDX]
        r.append("是" if (tn and _ticket_counts[tn] >= 2) else "")
        ws.append(r)
        row_idx = ws.max_row
        for c in range(1, len(_SUMMARY_HEADERS) + 1):
            cell = ws.cell(row=row_idx, column=c)
            cell.border = _SUMMARY_BORDER
            cell.alignment = _SUMMARY_CENTER
        # 金额列（E）、统筹列（F）写入的是数字，套两位小数金额格式；
        # 空字符串（解析失败/缺失）保持为空，不影响求和。
        for c in (5, 6):
            v = ws.cell(row=row_idx, column=c).value
            if isinstance(v, (int, float)):
                ws.cell(row=row_idx, column=c).number_format = _MONEY_FMT

    last = ws.max_row
    if last >= 2:
        # 合计行：金额合计写在 E 列、统筹合计写在 F 列（与表头列一一对应）
        ws.append(["合计", "", "", "",
                   f"=SUM(E2:E{last})", f"=SUM(F2:F{last})"])
        for c in range(1, len(_SUMMARY_HEADERS) + 1):
            cell = ws.cell(row=ws.max_row, column=c)
            cell.font = Font(bold=True)
            cell.border = _SUMMARY_BORDER
            cell.alignment = _SUMMARY_CENTER
            cell.fill = PatternFill("solid", fgColor="DDEBF7")
        for c in (5, 6):
            ws.cell(row=ws.max_row, column=c).number_format = _MONEY_FMT

    # ---- 右侧统计区（单独一组，竖排列更醒目）----
    # 列：I=标签, J=数值（G 列已用作「是否重复」数据列，H 为间隔列）
    stat_label_col, stat_val_col = 9, 10
    stat_title_row = 1
    stat_first_row = 2
    stat_items = [
        ("票据张数", f"=COUNTA(A2:A{last})"),
        ("合计总金额", f"=SUM(E2:E{last})"),
        ("合计总统筹金额", f"=SUM(F2:F{last})"),
        ("可赔付金额", "=J3-J4"),  # 合计总金额(J3) - 合计总统筹金额(J4)
        ("重复票据金额合计", f'=SUMIF(G2:G{last},"是",E2:E{last})'),
        ("重复票据统筹合计", f'=SUMIF(G2:G{last},"是",F2:F{last})'),
    ]
    # 标题
    tcell = ws.cell(row=stat_title_row, column=stat_label_col, value="统计")
    tcell.fill = _SUMMARY_FILL
    tcell.font = _SUMMARY_FONT
    tcell.alignment = _SUMMARY_CENTER
    tcell.border = _SUMMARY_BORDER
    tcell2 = ws.cell(row=stat_title_row, column=stat_val_col, value="数值")
    tcell2.fill = _SUMMARY_FILL
    tcell2.font = _SUMMARY_FONT
    tcell2.alignment = _SUMMARY_CENTER
    tcell2.border = _SUMMARY_BORDER
    # 数据行
    for i, (label, formula) in enumerate(stat_items):
        rrow = stat_first_row + i
        lc = ws.cell(row=rrow, column=stat_label_col, value=label)
        lc.font = Font(bold=True, color="1F4E78")
        lc.alignment = Alignment(horizontal="left", vertical="center")
        lc.border = _SUMMARY_BORDER
        vc = ws.cell(row=rrow, column=stat_val_col, value=formula)
        vc.font = Font(bold=True)
        vc.alignment = _SUMMARY_CENTER
        vc.border = _SUMMARY_BORDER
        vc.number_format = _MONEY_FMT if label != "票据张数" else "0"

    widths = [26, 14, 18, 14, 16, 18, 12, 3, 16, 18]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    ws.freeze_panes = "A2"
    wb.save(out_path)


def summarize_invoices(pdf_folder: str, log=print):
    """汇总 pdf_folder 内（含子目录）的所有 PDF 发票，输出 Excel。

    返回生成的 Excel 路径；无 PDF 或出错返回 None。
    """
    pdf_files = sorted(
        glob.glob(os.path.join(pdf_folder, "**", "*.pdf"), recursive=True),
        key=str.lower,
    )
    pdf_files = [f for f in pdf_files
                 if not os.path.basename(f).startswith("发票汇总_")]

    if not pdf_files:
        log("未找到可汇总的 PDF 文件，跳过汇总。")
        return None

    log(f"开始汇总：找到 {len(pdf_files)} 个 PDF 发票。")
    rows = []
    ok = 0
    for f in pdf_files:
        name = os.path.basename(f)
        fields, err = parse_invoice(f)
        if err:
            log(f"  [跳过] {name}：{err}")
            rows.append([name, "解析失败", "", "", "", ""])
            continue
        rows.append([
            name,
            fields["交款人"] or "",
            fields["票据号码"] or "",
            fields["开票日期"] or "",
            fields["金额合计（小写）"] if fields["金额合计（小写）"] is not None else "",
            fields["医保统筹基金支付"] if fields["医保统筹基金支付"] is not None else "",
        ])
        ok += 1
        log(f"  [OK] {name}  交款人={fields['交款人']}  "
            f"票据号={fields['票据号码']}  金额={fields['金额合计（小写）']}")

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(pdf_folder, f"发票汇总_{ts}.xlsx")
    _write_summary_excel(rows, out_path)
    log(f"汇总完成：成功 {ok} / 共 {len(pdf_files)} 个。")
    log(f"已导出：{out_path}")
    return out_path


def process_folder(
    folder: str,
    open_after: bool,
    convert_pdf: bool,
    summarize: bool,
    log_queue: queue.Queue,
):
    def log(msg: str):
        log_queue.put(("log", msg))

    def progress(current: int, total: int):
        log_queue.put(("progress", current, total))

    if not folder or not os.path.isdir(folder):
        log("错误：请选择一个有效的文件夹路径。")
        log_queue.put(("done",))
        return

    pdf_dir = os.path.join(folder, "PDF")
    img_dir = os.path.join(pdf_dir, "图片") if convert_pdf else None

    os.makedirs(pdf_dir, exist_ok=True)
    if img_dir:
        os.makedirs(img_dir, exist_ok=True)

    files = [
        f
        for f in os.listdir(folder)
        if f.lower().endswith(SUPPORTED_IMAGE_EXTS)
        and os.path.isfile(os.path.join(folder, f))
    ]
    files.sort(key=lambda x: x.lower())
    total = len(files)
    log(f"共发现 {total} 张待处理图片。")

    # 统计各项识别结果（已去除「重复网址」逻辑：所有图片均识别，不再按网址去重）
    stats = {
        "total": total,
        "success": 0,       # 成功识别并下载 PDF（原文件名不动）
        "no_pdf": 0,        # 识别到网址但未下载到 PDF（复制到「未识别」，前缀 未下载-）
        "unrecognized": 0,  # 未识别到任何二维码（复制到「未识别」，前缀 未识别-）
        "other": 0,         # 其它情况：识别到二维码但内容非网址等（复制到「未识别」，前缀 其它-）
    }

    def _copy_to_unrecognized(fpath, fname, prefix):
        """把问题图片复制一份到 folder/未识别/ 下，文件名前加 prefix；原图片保留不动。
        返回新路径或 None。"""
        sub = os.path.join(folder, "未识别")
        try:
            os.makedirs(sub, exist_ok=True)
        except Exception as e:
            log(f"  -> 创建「未识别」文件夹失败：{e}")
            return None
        new_name = prefix + fname
        dest = os.path.join(sub, new_name)
        try:
            shutil.copy2(fpath, dest)
            log(f"  -> 已复制到「未识别」：{new_name}")
        except Exception as e:
            log(f"  -> 复制失败：{e}")
            return None
        return dest

    for idx, fname in enumerate(files, start=1):
        progress(idx, total)
        fpath = os.path.join(folder, fname)
        base_name, _ = os.path.splitext(fname)
        log(f"[{idx}/{total}] 正在处理：{fname}")

        try:
            codes = detect_qr_codes(fpath)
        except Exception as e:
            log(f"  -> 识别过程出错：{e}")
            codes = []

        if not codes:
            # 未识别到任何二维码
            stats["unrecognized"] += 1
            _copy_to_unrecognized(fpath, fname, PREFIX_UNRECOGNIZED)
            continue

        # 在识别到的二维码中优先找一个网址
        url = None
        for code in codes:
            candidate = is_url(code)
            if candidate:
                url = candidate
                break

        if url:
            pdf_path = os.path.join(pdf_dir, f"{base_name}.pdf")
            try:
                download_pdf(url, pdf_path)
                log(f"  -> 已下载 PDF：{base_name}.pdf")
                stats["success"] += 1
                # 成功：原文件名不动，不重命名
                if convert_pdf and img_dir:
                    try:
                        imgs = convert_pdf_to_images(pdf_path, img_dir, base_name)
                        for img_path in imgs:
                            log(f"  -> 已生成图片：{os.path.basename(img_path)}")
                    except Exception as e:
                        log(f"  -> PDF 转图片失败：{e}")
            except Exception as e:
                log(f"  -> 网址无可下载的 PDF（{e}）")
                stats["no_pdf"] += 1
                _copy_to_unrecognized(fpath, fname, PREFIX_NOT_DOWNLOADED)
        else:
            # 识别到二维码但不是网址（含纯数字等）：归入「其它」
            log(f"  -> 识别到二维码但非网址，归入「其它」")
            stats["other"] += 1
            _copy_to_unrecognized(fpath, fname, PREFIX_OTHER)

    log("全部处理完成。")

    # 输出识别结果统计
    summary = [
        "=== 识别结果统计 ===",
        f"总计识别图片：{stats['total']} 张",
        f"✓ 成功识别并下载 PDF（原文件名不变）：{stats['success']} 张",
        f"⚠ 识别到网址但未下载 PDF（复制到未识别/，未下载-）：{stats['no_pdf']} 张",
        f"✗ 未识别到二维码（复制到未识别/，未识别-）：{stats['unrecognized']} 张",
        f"· 其它情况（复制到未识别/，其它-）：{stats['other']} 张",
    ]
    for line in summary:
        log(line)

    # 把结构化统计传给 GUI（用于结束弹窗）
    log_queue.put(("stats", dict(stats)))

    if summarize:
        log("--- 开始汇总发票 ---")
        try:
            out = summarize_invoices(pdf_dir, log=log)
            if out:
                log(f"发票汇总已生成：{os.path.basename(out)}")
        except Exception as e:
            log(f"汇总发票出错：{e}")

    if open_after:
        try:
            os.startfile(folder)
            log("已打开目标文件夹。")
        except Exception as e:
            log(f"打开文件夹失败：{e}")

    log_queue.put(("done",))


class _FlatProgressBar(tk.Canvas):
    """扁平进度条：浅灰轨道 + 细边框 + 蓝色填充，与整体界面风格保持一致。

    系统默认进度条在 Windows 上是绿色渐变，跟这套扁平浅色界面不搭，故自绘。"""

    TRACK = "#EDEFF2"
    BORDER = "#D6DBE1"
    FILL = "#2563EB"

    def __init__(self, master, height: int = 20, **kw):
        super().__init__(
            master, height=height, bg=self.TRACK, highlightthickness=0, bd=0, **kw
        )
        self._value = 0.0
        self.bind("<Configure>", lambda _e: self._redraw())

    def set_value(self, pct: float):
        try:
            pct = float(pct)
        except (TypeError, ValueError):
            pct = 0.0
        self._value = max(0.0, min(100.0, pct))
        self._redraw()

    def _redraw(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w <= 2 or h <= 2:
            return
        self.create_rectangle(
            0, 0, w - 1, h - 1, outline=self.BORDER, fill=self.TRACK
        )
        fill_w = (w - 2) * self._value / 100.0
        if fill_w >= 1:
            self.create_rectangle(
                1, 1, 1 + fill_w, h - 2, outline="", fill=self.FILL
            )


class InvoiceQrToolApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"发票二维码识别下载工具 v{__VERSION__}")
        self.root.geometry("800x600")
        self.root.minsize(700, 450)
        try:
            self.root.iconbitmap(_resource_path("app_icon.ico"))
        except Exception:
            pass

        self.folder_var = tk.StringVar()
        self.open_after_var = tk.BooleanVar(value=True)
        self.convert_pdf_var = tk.BooleanVar(value=False)
        self.summarize_var = tk.BooleanVar(value=False)

        self.log_queue: queue.Queue = queue.Queue()
        self.worker_thread: threading.Thread | None = None
        self.last_stats: dict | None = None

        self._build_ui()
        self._poll_log()
        # 启动后静默检查更新（仅发现新版本时弹窗，无网络/无更新时不打扰）
        threading.Thread(
            target=_silent_startup_check, args=(self.root,), daemon=True
        ).start()

    def _build_menubar(self):
        """顶部菜单栏：承载「使用说明 / 检查更新 / 清空日志」等次要功能。"""
        self.menubar = tk.Menu(self.root)

        self.menu_help = tk.Menu(self.menubar, tearoff=0)
        self.menu_help.add_command(
            label="使用说明 / 更新记录", command=self._show_help
        )
        self.menu_help.add_command(label="检查更新", command=self._check_update)
        self.menu_help.add_separator()
        self.menu_help.add_command(label="清空日志", command=self._clear_log)
        self.menubar.add_cascade(label="帮助", menu=self.menu_help)

        self.root.config(menu=self.menubar)

    def _build_ui(self):
        pad = {"padx": 10, "pady": 8}

        self._build_menubar()

        # 文件夹选择
        frame_path = ttk.Frame(self.root)
        frame_path.pack(fill=tk.X, **pad)
        ttk.Label(frame_path, text="目标文件夹：").pack(side=tk.LEFT)
        ent_path = ttk.Entry(frame_path, textvariable=self.folder_var)
        ent_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(frame_path, text="浏览…", command=self._browse_folder).pack(
            side=tk.LEFT
        )

        # 选项
        frame_opts = ttk.Frame(self.root)
        frame_opts.pack(fill=tk.X, **pad)
        ttk.Checkbutton(
            frame_opts,
            text="处理完成后打开文件夹",
            variable=self.open_after_var,
        ).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(
            frame_opts,
            text="将下载的 PDF 转换为图片（JPG，长边 2000px）",
            variable=self.convert_pdf_var,
        ).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(
            frame_opts,
            text="处理完成后汇总发票（生成 Excel）",
            variable=self.summarize_var,
        ).pack(side=tk.LEFT, padx=5)

        # 操作按钮区：整行扁平主按钮（浅底 + 细边框 + 居中文字），点击区域大、易点；
        # 「使用说明 / 检查更新」等次要功能已移入顶部菜单栏。
        frame_btn_border = tk.Frame(self.root, bg="#D6DBE1")  # 外层充当 1px 细边框
        frame_btn_border.pack(fill=tk.X, padx=10, pady=(12, 6))

        self.btn_start = tk.Button(
            frame_btn_border,
            text="▶  开始处理",
            command=self._start_processing,
            font=("Microsoft YaHei", 12, "bold"),
            bg="#FAFAFB",
            fg="#1F2937",
            activebackground="#EEF1F4",
            activeforeground="#1F2937",
            disabledforeground="#9CA3AF",
            relief=tk.FLAT,
            bd=0,
            highlightthickness=0,
            cursor="hand2",
            pady=11,
        )
        self.btn_start.pack(fill=tk.X, padx=1, pady=1)

        # 进度条 + 右侧「进度：x / y 份」计数，同一行展示
        frame_progress = ttk.Frame(self.root)
        frame_progress.pack(fill=tk.X, padx=10, pady=(0, 8))
        self.progress = _FlatProgressBar(frame_progress, height=20)
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.progress_var = tk.StringVar(value="进度：0 / 0 份")
        ttk.Label(
            frame_progress, textvariable=self.progress_var, width=14, anchor=tk.E
        ).pack(side=tk.LEFT, padx=(10, 2))

        # 日志区（标题行右侧提供「清空日志」，便于开始下一个任务前清空）
        frame_log_head = ttk.Frame(self.root)
        frame_log_head.pack(fill=tk.X, padx=10)
        ttk.Label(frame_log_head, text="处理日志：").pack(side=tk.LEFT)
        ttk.Button(
            frame_log_head, text="清空日志", command=self._clear_log, width=10
        ).pack(side=tk.RIGHT)
        self.txt_log = scrolledtext.ScrolledText(
            self.root, wrap=tk.WORD, state=tk.DISABLED, height=20
        )
        self.txt_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=(4, 10))

    def _browse_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folder_var.set(path)

    def _check_update(self):
        self._log("正在检查更新…")
        threading.Thread(
            target=check_and_prompt_update, args=(self.root,), daemon=True
        ).start()

    def _show_help(self):
        win = tk.Toplevel(self.root)
        win.title("使用说明 / 更新记录")
        win.geometry("660x540")
        try:
            win.transient(self.root)
            win.grab_set()
        except Exception:
            pass
        txt = scrolledtext.ScrolledText(win, wrap=tk.WORD, state=tk.NORMAL)
        txt.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        txt.insert(tk.END, USAGE_TEXT + "\n\n" + CHANGELOG_TEXT)
        txt.configure(state=tk.DISABLED)
        ttk.Button(win, text="关闭", command=win.destroy).pack(pady=6)

    def _clear_log(self):
        """清空日志区内容（不影响正在运行的任务），便于开始下一个任务。"""
        self.txt_log.configure(state=tk.NORMAL)
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.configure(state=tk.DISABLED)

    def _log(self, msg: str):
        now = time.strftime("%H:%M:%S")
        self.txt_log.configure(state=tk.NORMAL)
        self.txt_log.insert(tk.END, f"[{now}] {msg}\n")
        self.txt_log.see(tk.END)
        self.txt_log.configure(state=tk.DISABLED)

    def _poll_log(self):
        try:
            while True:
                item = self.log_queue.get_nowait()
                if item[0] == "log":
                    self._log(item[1])
                elif item[0] == "progress":
                    _, current, total = item
                    if total > 0:
                        self.progress.set_value((current / total) * 100)
                        self.progress_var.set(f"进度：{current} / {total} 份")
                    else:
                        self.progress.set_value(100)
                        self.progress_var.set("进度：0 / 0 份")
                elif item[0] == "stats":
                    self.last_stats = item[1]
                elif item[0] == "done":
                    self.progress.set_value(100)
                    self.btn_start.configure(state=tk.NORMAL, text="▶  开始处理")
                    self._log("--- 任务结束 ---")
                    s = getattr(self, "last_stats", None)
                    if s:
                        self._show_stats_popup(s)
                        self.last_stats = None
        except queue.Empty:
            pass
        self.root.after(100, self._poll_log)

    def _show_stats_popup(self, s: dict):
        msg = (
            "本次识别结果统计：\n\n"
            f"总计识别图片：{s['total']} 张\n"
            f"成功下载 PDF（原文件名不变）：{s['success']} 张\n"
            f"识别但未下载 PDF（未下载-）：{s['no_pdf']} 张\n"
            f"未识别到二维码（未识别-）：{s['unrecognized']} 张\n"
            f"其它情况（其它-）：{s['other']} 张"
        )
        messagebox.showinfo("识别结果统计", msg)

    def _start_processing(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("路径错误", "请选择一个有效的文件夹。")
            return

        self.btn_start.configure(state=tk.DISABLED, text="处理中…")
        self.progress.set_value(0)
        self.progress_var.set("进度：0 / 0 份")
        self._log("=== 开始处理 ===")

        self.worker_thread = threading.Thread(
            target=process_folder,
            args=(
                folder,
                self.open_after_var.get(),
                self.convert_pdf_var.get(),
                self.summarize_var.get(),
                self.log_queue,
            ),
            daemon=True,
        )
        self.worker_thread.start()


def _cli_test(folder: str, summarize: bool = False) -> None:
    """无窗口测试入口：验证 EXE 内各依赖是否能正常工作。"""
    log_path = os.path.join(os.path.dirname(sys.argv[0]), "test_run.log")
    q: queue.Queue = queue.Queue()
    process_folder(
        folder,
        open_after=False,
        convert_pdf=True,
        summarize=summarize,
        log_queue=q,
    )
    with open(log_path, "w", encoding="utf-8") as f:
        while True:
            item = q.get()
            if item[0] == "log":
                f.write(item[1] + "\n")
            elif item[0] == "done":
                f.write("--- 测试完成 ---\n")
                break


def _cleanup_legacy_update_artifacts():
    """清理早期版本（bat / --pending / 备份机制）遗留的临时文件，避免堆积。"""
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


def main():
    args = sys.argv[1:]

    # “回收旧版本”模式：由上一版本以 --recycle-old "<旧exe>" 启动本进程，
    # 等旧进程退出后再把旧 exe 移入回收站（避免删除正在运行的自身文件被系统锁定）。
    if "--recycle-old" in args:
        i = args.index("--recycle-old")
        old_path = args[i + 1] if i + 1 < len(args) else None
        if old_path and old_path.lower().endswith(".exe"):
            for _ in range(40):  # 最多约 20 秒，等旧进程释放文件
                if not os.path.exists(old_path):
                    break
                if send_to_recycle_bin(old_path):
                    break
                time.sleep(0.5)

    # 清理早期更新机制遗留的临时文件
    _cleanup_legacy_update_artifacts()

    if "--test" in args:
        i = args.index("--test")
        folder = args[i + 1] if i + 1 < len(args) else None
        if not folder:
            print("用法：发票二维码工具.exe --test <文件夹> [--summary]")
            return
        summarize = "--summary" in args
        _cli_test(folder, summarize=summarize)
        return

    root = tk.Tk()
    app = InvoiceQrToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
