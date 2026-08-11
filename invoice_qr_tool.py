#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
发票二维码识别下载工具
功能：
1. 识别指定文件夹内图片中的二维码；
2. 若二维码为网址，则下载对应 PDF 到同目录 PDF 文件夹下，文件名与原始图片相同；
3. 若二维码为纯数字，则忽略；
4. 若未识别到二维码 / 无法识别，则图片名前加“未识别-”；
5. 若二维码为网址但无可下载的 PDF，则图片名前加“未下载”；
6. 若同一网址在多处出现（重复），则对应图片名前加“重复”；
   “未下载”与“重复”可叠加，例如“未下载重复原文件名”；
7. 可选任务完成后打开文件夹；
8. 可选将下载的 PDF 转换为图片（长边 2000px，短边自适应）；
9. 可选任务完成后汇总发票（对 PDF 文件夹内发票 PDF 提取字段并生成 Excel）。
"""

import os
import re
import sys
import glob
import time
import queue
import datetime
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import cv2
import numpy as np
import requests
import pymupdf  # PyMuPDF（新版推荐入口，避免 fitz 弃用警告）
import pdfplumber  # 汇总发票：PDF 文本/词抽取
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# zxing-cpp 识别能力比 OpenCV 自带 QR 检测更强，对截图/小二维码更鲁棒
ZXING_AVAILABLE = False
try:
    import zxingcpp

    ZXING_AVAILABLE = True
except Exception:
    pass

URL_RE = re.compile(r"https?://[^\s<>\"{}|\\^`\[\]]+", re.IGNORECASE)

SUPPORTED_IMAGE_EXTS = (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp")

# 状态前缀：用于“未识别/未下载/重复”重命名，并在重复处理时避免叠加前缀
PREFIX_UNRECOGNIZED = "未识别-"  # 注意带连字符，保持与原规则一致
PREFIX_NOT_DOWNLOADED = "未下载"
PREFIX_DUPLICATE = "重复"


USAGE_TEXT = """发票二维码识别下载工具 · 使用说明
================================

【功能】
1. 识别指定文件夹内图片中的二维码。
2. 二维码为网址：自动下载对应 PDF 到「PDF」子文件夹，文件名与原始图片相同。
3. 二维码为纯数字：忽略，不做任何处理。
4. 未识别到二维码 / 二维码无法识别：图片重命名为「未识别-原文件名」。
5. 二维码为网址但无可下载的 PDF：图片重命名为「未下载原文件名」。
6. 同一网址在多处出现（重复）：对应图片文件名前加「重复」。
7. 可选：处理完成后自动打开文件夹。
8. 可选：将下载的 PDF 转为图片（长边 2000px，短边自适应），保存到「PDF/图片」。
9. 可选：处理完成后汇总发票（生成 Excel）。对「PDF」文件夹内所有发票 PDF，
   提取「交款人 / 票据号码 / 开票日期 / 金额合计（小写）/ 医保统筹基金支付」，
   汇总为「PDF/发票汇总_YYYYMMDD_HHMMSS.xlsx」，并带合计行（金额、医保统筹自动求和）。

【使用步骤】
1. 把待处理的图片放在同一个文件夹里。
2. 打开本软件，点「浏览…」选择该文件夹。
3. 按需勾选：
   - 处理完成后打开文件夹
   - 将下载的 PDF 转换为图片（长边 2000px）
   - 处理完成后汇总发票（生成 Excel）
4. 点「开始处理」，在日志区查看进度。
5. 处理完成后，PDF 在「PDF」文件夹，转换图片在「PDF/图片」，汇总表在「PDF/发票汇总_*.xlsx」。

【输出规则速查】
- 网址 + 下载成功      → PDF/<原名>.pdf
- 网址 + 无 PDF 可下载 → 图片改名「未下载<原名>」
- 纯数字二维码         → 忽略
- 无/无法识别二维码     → 图片改名「未识别-<原名>」
- 重复网址             → 图片加「重复」前缀
- 上述情况可叠加，例如「未下载重复<原名>」
- 勾选汇总             → PDF/发票汇总_YYYYMMDD_HHMMSS.xlsx（含合计行）

【说明】
- 二维码识别使用 zxing-cpp，对截图 / 小二维码会自动多尺度放大，比 OpenCV 自带更稳。
- 部分发票平台（如 jsczt.cn）打开后是一个展示页，软件会自动提取页面隐藏参数并提交下载接口获取 PDF。
- 单文件 EXE，无需安装，双击即用。
"""

CHANGELOG_TEXT = """发票二维码识别下载工具 · 更新记录
================================

2026-08-11  v1.0
- 初版发布：二维码识别、网址下载 PDF、纯数字忽略、未识别重命名、
  可选打开文件夹、可选 PDF 转图片。

2026-08-11  v1.1
- 新增：二维码为网址但无可下载 PDF 时，图片重命名为「未下载<原文件名>」。
- 新增：检测到重复网址时，对应图片文件名前加「重复」标识（可与「未下载」叠加）。
- 新增：内置「使用说明」窗口与「更新记录」窗口，
  并随软件附带 使用说明.txt、更新日志.txt。

2026-08-11  v2.0
- 集成「汇总发票」功能（来自“汇总票据”会话的 invoice_summary.py）：
  处理完成后可将「PDF」文件夹内发票 PDF 提取关键字段并汇总成 Excel。
- 主页面新增开关：「处理完成后汇总发票（生成 Excel）」。
- 提取字段：文件名 / 交款人 / 票据号码 / 开票日期 / 金额合计（小写）/ 医保统筹基金支付；
  输出「PDF/发票汇总_YYYYMMDD_HHMMSS.xlsx」，含金额与医保统筹合计行。
- 新增依赖 pdfplumber（PDF 文字抽取）+ openpyxl（Excel 写入），已纳入单文件 EXE。
"""



def detect_qr_codes(image_path: str):
    """返回图片中识别到的所有二维码文本列表（去重）"""
    # 使用 numpy 读字节再 imdecode，避免 OpenCV 在 Windows 上处理中文路径的编码问题
    img = cv2.imdecode(np.fromfile(image_path, dtype=np.uint8), cv2.IMREAD_COLOR)
    if img is None:
        return []

    codes = []

    # 1. zxing-cpp 识别能力更强，先尝试；对小二维码会自动多尺度放大重试
    if ZXING_AVAILABLE:
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
    """将 PDF 每一页渲染为图片，长边 2000px，短边自适应；返回生成的文件路径列表"""
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
            out_name = f"{base_name}_第{page_num + 1}页.png"
            out_path = os.path.join(out_dir, out_name)
            pix.save(out_path)
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
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return m.group(0) if m else None


def parse_invoice(pdf_path):
    """解析单个 PDF，返回 (字段字典, 错误信息或 None)。"""
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
    fields["金额合计（小写）"] = _extract_number(raw_amount) if raw_amount else None
    fields["医保统筹基金支付"] = _extract_number(
        _field_search(pages_words, [("医保统筹基金支付：", False)]))

    return fields, None


_SUMMARY_HEADERS = ["文件名", "交款人", "票据号码", "开票日期",
                    "金额合计（小写）", "医保统筹基金支付"]
_SUMMARY_FILL = PatternFill("solid", fgColor="1F4E78")
_SUMMARY_FONT = Font(bold=True, color="FFFFFF", size=11)
_SUMMARY_THIN = Side(style="thin", color="BFBFBF")
_SUMMARY_BORDER = Border(left=_SUMMARY_THIN, right=_SUMMARY_THIN,
                          top=_SUMMARY_THIN, bottom=_SUMMARY_THIN)
_SUMMARY_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _write_summary_excel(rows, out_path):
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

    for r in rows:
        ws.append(r)
        row_idx = ws.max_row
        for c in range(1, len(_SUMMARY_HEADERS) + 1):
            cell = ws.cell(row=row_idx, column=c)
            cell.border = _SUMMARY_BORDER
            cell.alignment = _SUMMARY_CENTER

    last = ws.max_row
    if last >= 2:
        ws.append(["合计", "", "", "",
                   f"=SUM(E2:E{last})", f"=SUM(F2:F{last})"])
        for c in range(1, len(_SUMMARY_HEADERS) + 1):
            cell = ws.cell(row=ws.max_row, column=c)
            cell.font = Font(bold=True)
            cell.border = _SUMMARY_BORDER
            cell.alignment = _SUMMARY_CENTER
            cell.fill = PatternFill("solid", fgColor="DDEBF7")

    widths = [26, 14, 18, 14, 16, 18]
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
            fields["金额合计（小写）"] or "",
            fields["医保统筹基金支付"] or "",
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

    # 记录本批次已出现过的网址，用于重复检测
    seen_urls: set[str] = set()

    def _rename_with_prefix(fpath, fname, prefix):
        """若尚未加过该前缀，则重命名；返回新的完整路径"""
        if fname.startswith(prefix):
            return fpath
        new_name = prefix + fname
        new_path = os.path.join(folder, new_name)
        try:
            os.rename(fpath, new_path)
            log(f"  -> 已重命名为：{new_name}")
        except Exception as e:
            log(f"  -> 重命名失败：{e}")
        return new_path

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
            _rename_with_prefix(fpath, fname, PREFIX_UNRECOGNIZED)
            continue

        url = None
        has_digit_only = False
        for code in codes:
            candidate = is_url(code)
            if candidate:
                url = candidate
                break
            if code.isdigit():
                has_digit_only = True

        if url:
            # 重复检测：同一网址再次出现则标记“重复”
            is_duplicate = url in seen_urls
            seen_urls.add(url)

            pdf_path = os.path.join(pdf_dir, f"{base_name}.pdf")
            download_ok = False
            try:
                download_pdf(url, pdf_path)
                download_ok = True
                log(f"  -> 已下载 PDF：{base_name}.pdf")

                if convert_pdf and img_dir:
                    try:
                        imgs = convert_pdf_to_images(pdf_path, img_dir, base_name)
                        for img_path in imgs:
                            log(f"  -> 已生成图片：{os.path.basename(img_path)}")
                    except Exception as e:
                        log(f"  -> PDF 转图片失败：{e}")
            except Exception as e:
                log(f"  -> 网址无可下载的 PDF（{e}）")

            # 根据结果叠加前缀
            if not download_ok:
                fpath = _rename_with_prefix(fpath, fname, PREFIX_NOT_DOWNLOADED)
                fname = os.path.basename(fpath)
            if is_duplicate:
                fpath = _rename_with_prefix(fpath, fname, PREFIX_DUPLICATE)
                fname = os.path.basename(fpath)
        elif has_digit_only:
            log(f"  -> 识别到纯数字二维码，已忽略。")
        else:
            # 识别到二维码但内容不是网址，也视为无效
            _rename_with_prefix(fpath, fname, PREFIX_UNRECOGNIZED)

    log("全部处理完成。")

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


class InvoiceQrToolApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("发票二维码识别下载工具")
        self.root.geometry("800x600")
        self.root.minsize(700, 450)

        self.folder_var = tk.StringVar()
        self.open_after_var = tk.BooleanVar(value=True)
        self.convert_pdf_var = tk.BooleanVar(value=False)
        self.summarize_var = tk.BooleanVar(value=False)

        self.log_queue: queue.Queue = queue.Queue()
        self.worker_thread: threading.Thread | None = None

        self._build_ui()
        self._poll_log()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 8}

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
            text="将下载的 PDF 转换为图片（长边 2000px）",
            variable=self.convert_pdf_var,
        ).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(
            frame_opts,
            text="处理完成后汇总发票（生成 Excel）",
            variable=self.summarize_var,
        ).pack(side=tk.LEFT, padx=5)

        # 操作按钮 + 进度条
        frame_action = ttk.Frame(self.root)
        frame_action.pack(fill=tk.X, **pad)
        self.btn_start = ttk.Button(
            frame_action, text="开始处理", command=self._start_processing
        )
        self.btn_start.pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_action, text="使用说明", command=self._show_help).pack(
            side=tk.RIGHT, padx=5
        )
        self.progress = ttk.Progressbar(
            frame_action, mode="determinate", maximum=100
        )
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)

        # 日志区
        ttk.Label(self.root, text="处理日志：").pack(anchor=tk.W, padx=10)
        self.txt_log = scrolledtext.ScrolledText(
            self.root, wrap=tk.WORD, state=tk.DISABLED, height=20
        )
        self.txt_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    def _browse_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folder_var.set(path)

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
                        self.progress["value"] = (current / total) * 100
                    else:
                        self.progress["value"] = 100
                elif item[0] == "done":
                    self.progress["value"] = 100
                    self.btn_start.configure(state=tk.NORMAL, text="开始处理")
                    self._log("--- 任务结束 ---")
        except queue.Empty:
            pass
        self.root.after(100, self._poll_log)

    def _start_processing(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("路径错误", "请选择一个有效的文件夹。")
            return

        self.btn_start.configure(state=tk.DISABLED, text="处理中…")
        self.progress["value"] = 0
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


def main():
    args = sys.argv[1:]
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
