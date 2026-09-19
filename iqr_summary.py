#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""发票二维码工具 · 发票解析与汇总

从票据 PDF 提取字段（交款人 / 票据号码 / 开票日期 / 金额合计 / 各支付类），
并把整批票据汇总成 Excel（「统计（剔重后）」+「重复票据」两块，含公式缓存值注入）。

拆出来的目的：主程序专注界面与流程编排。pdfplumber / openpyxl 仍是函数内懒加载，
import 本模块不会拖慢启动。PyInstaller 打包会沿 import 自动收集，spec 无需改动。
"""

import os
import re
from collections import Counter
# ⚠️ 必须是 `import datetime`（模块），不是 `from datetime import datetime`：
#    下面调用写的是 datetime.datetime.now()。拆模块时这行被写成了导入类，
#    结果汇总发票在生成文件名这一步直接 AttributeError 崩掉（v4.10.0 线上 bug）。
import datetime




# =====================================================================
# 发票汇总模块（来自“汇总票据”会话的 invoice_summary.py，集成到此工具）
# 针对江苏省医疗门诊收费票据 / 电子票据，提取关键字段并汇总到 Excel。
# =====================================================================

def _norm_colon(s: str) -> str:
    """把全角冒号统一成半角，便于匹配标签。"""
    return s.replace("：", ":")


class _Pages:
    """页词表的查询视图：每页「全角冒号规范化」后的文本按需生成、只做一次。

    parse_invoice 会按 8 个标签轮流查同一批页；若每次查询都现场规范化整页文本，
    同一页就要重复处理 8 遍。这里把结果缓存下来复用，而且**只在该页真的被查到
    时才生成** —— 靠前就能命中的字段不必为整页付代价。
    """

    __slots__ = ("_pages", "_norm")

    def __init__(self, pages_words):
        self._pages = list(pages_words)
        self._norm: list = [None] * len(self._pages)

    def __iter__(self):
        for i, words in enumerate(self._pages):
            norm = self._norm[i]
            if norm is None:
                norm = [_norm_colon(w["text"]) for w in words]
                self._norm[i] = norm
            yield words, norm


def _extract_field(words, label, gather_line=False, norm=None):
    """在单页 words 中查找包含 label 的词，返回其后的取值。

    ``norm`` 是该页 words 的规范化文本表（见 ``_Pages``）；传入可省掉重复规范化。
    """
    if norm is None:
        norm = [_norm_colon(w["text"]) for w in words]
    labeln = _norm_colon(label)
    for i, tn in enumerate(norm):
        idx = tn.find(labeln)
        if idx == -1:
            continue
        val = tn[idx + len(labeln):].strip()
        if gather_line:
            # 「同一行」= |Δtop| < 3 且位于该标签右侧，按 x0 从左到右拼接
            w = words[i]
            line_words = [x for x in words
                          if abs(x["top"] - w["top"]) < 3 and x["x0"] > w["x0"]]
            line_words.sort(key=lambda x: x["x0"])
            val = (val + "".join(x["text"] for x in line_words)).strip()
        return val
    return None


def _field_search(pages, candidates):
    """跨页查找字段。candidates: [(label, gather_line), ...]。

    ``pages`` 推荐传 ``_Pages(page.extract_words() 列表)``，让每页的文本规范化
    在整份 PDF 内只做一次；传普通 words 列表也兼容（会自动包装）。
    """
    if not isinstance(pages, _Pages):
        pages = _Pages(pages)
    for words, norm in pages:
        for label, gather in candidates:
            v = _extract_field(words, label, gather, norm=norm)
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
        "大病保险支付": None,
        "医疗救助支付": None,
    }
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # 用 _Pages 包一层：8 个字段轮流查询，每页文本只规范化一次
            pages_words = _Pages([page.extract_words() for page in pdf.pages])
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
    # 大病保险支付：与「医保统筹基金支付」同属票据右下角的基金支付区，
    # 版式与取值规则一致，故用同样的方式识别（全角/半角冒号由 _norm_colon 统一）。
    # 识别到该栏（含 0.00）即记为数值；整份票据没有这一栏时为 None，
    # 汇总表据此决定是否输出该列 —— 没识别到就整体忽略，不留空列。
    fields["大病保险支付"] = _to_num(
        _field_search(pages_words, [("大病保险支付：", False)]))
    # 医疗救助支付：同区块、同版式，同样处理（识别到 0.00 也算"有这一类目"）
    fields["医疗救助支付"] = _to_num(
        _field_search(pages_words, [("医疗救助支付：", False)]))

    return fields, None


# ---- 汇总表列定义 ----
# 固定列：无论票据有无该类目都会输出
_SUMMARY_BASE_FIELDS = ["文件名", "交款人", "票据号码", "开票日期",
                        "金额合计（小写）", "医保统筹基金支付"]


# 条件列：**识别到该类目才输出**（本批票据一张都没识别到 → 该列连同其汇总项整体省略）。
# 顺序即列序。将来再加新类目：这里加个名字 + parse_invoice 里取值即可，其余全自动。
_SUMMARY_COND_FIELDS = ["大病保险支付", "医疗救助支付"]


# 参与「可赔付金额」扣减的支付类字段：凡出现在表里的都要从合计总金额中减去
# （这些钱由医保基金 / 商业保险 / 医疗救助支付，患者并未实际支出）
_DEDUCT_FIELDS = ["医保统筹基金支付", "大病保险支付", "医疗救助支付"]


_SUMMARY_ALL_FIELDS = _SUMMARY_BASE_FIELDS + _SUMMARY_COND_FIELDS


_AMOUNT_FIELD = "金额合计（小写）"


_POOL_FIELD = "医保统筹基金支付"


_TICKET_FIELD = "票据号码"


# 序号列：**不来自解析结果**，由写表时按明细顺序生成（1、2、3…），
# 放在最左侧。它同时也是「是否重复」互指文案里的编号基准
# （文案写「与序号 4、6 重复」，用户照着序号列就能直接找到那一行，
#   避免用 Excel 行号时与序号列差 2 造成找不到）。
_SEQ_FIELD = "序号"


SEQ_COL_WIDTH = 6


_AUX_TAG = "重复份"        # 辅助列标记值：同一票据号第 2 张及以后


_AUX_FIRST = "首份"        # 辅助列标记值：该票号的首张（空票号也算首份，不算重复）


_AUX_HEADER = "首份/重复份标记（辅助列，不参与展示）"


def _total_label(field):
    """统计区「合计 XX」的显示名（统筹沿用历史叫法「合计总统筹金额」）。"""
    return "合计总统筹金额" if field == _POOL_FIELD else "合计" + field


def _dup_label(field):
    """统计区「重复票据」类项目的显示名。"""
    return "重复票据统筹合计" if field == _POOL_FIELD else "重复票据" + field + "合计"


def _inject_formula_cache(xlsx_path, values):
    """把公式的计算结果写回单元格缓存值 ``<v>``（**公式本身保留**）。

    背景：openpyxl 只写公式、不计算结果，单元格里留的是空 ``<v />``。Excel / WPS 打开会
    重算，所以本机看着正常；但**不重算公式的查看器**（腾讯文档在线预览、部分网盘预览、
    ``pandas.read_excel`` / ``openpyxl(data_only=True)``）读到的是空值 —— 统计栏显示 0
    或空白，看着像算错了。这里按公式语义把结果写回，任何查看器都能直接看到数值。

    ``values``：``{单元格引用（如 "L2"）: (值, 是否文本)}``；返回命中的单元格数。
    """
    import zipfile
    from xml.sax.saxutils import escape

    if not values:
        return 0

    sheet = "xl/worksheets/sheet1.xml"
    # ⚠️ 匹配一个完整单元格时，必须兼顾自闭合空单元格：若写成
    #    '<c [^>]*r="..."[^>]*>.*?</c>'，遇到 <c r="D18" s="9"/> 时 `>` 会匹配到它自己的
    #    `>`，随后 `.*?</c>` 一路吞掉**紧跟其后的那个单元格**，导致后者漏注入。
    cell_re = re.compile(r"<c\b[^>]*?(?:/>|>.*?</c>)", re.S)
    ref_re = re.compile(r'\br="([A-Z]+)(\d+)"')
    empty_v = re.compile(r"<v\s*/>|<v>\s*</v>")

    def _num_str(v):
        if isinstance(v, bool):
            return "1" if v else "0"
        if isinstance(v, int):
            return str(v)
        s = ("%.10f" % float(v)).rstrip("0").rstrip(".")
        return s or "0"

    def _patch(m):
        cell = m.group(0)
        rm = ref_re.search(cell)
        if not rm:
            return cell
        ref = rm.group(1) + rm.group(2)
        if ref not in values:
            return cell
        val, is_text = values[ref]
        text = str(val) if is_text else _num_str(val)
        head, sep, body = cell.partition(">")
        if is_text and ' t="' not in head:
            head = head.replace("<c ", '<c t="str" ', 1)
        body, n = empty_v.subn("<v>%s</v>" % escape(text), body)
        if n == 0:                       # 本来就没有 <v>，补在 </c> 之前
            body = body.replace("</c>", "<v>%s</v></c>" % escape(text))
        return head + sep + body

    zin = zipfile.ZipFile(xlsx_path, "r")
    try:
        items = [(it, zin.read(it.filename)) for it in zin.infolist()]
    finally:
        zin.close()

    hit = 0
    out = []
    for it, data in items:
        if it.filename == sheet:
            xml = data.decode("utf-8")
            hit = sum(1 for ref in values if ('r="%s"' % ref) in xml)
            data = cell_re.sub(_patch, xml).encode("utf-8")
        out.append((it, data))

    # ⚠️ 直接覆写目标文件。不要用「临时文件 + os.replace」：目标被 Office / 预览器打开时，
    #    os.replace 需要目标文件的 DELETE 权限 → PermissionError(WinError 5)；
    #    而普通写入（截断重写）是放行的（openpyxl 的 wb.save 就是这么保存的）。
    zout = zipfile.ZipFile(xlsx_path, "w", zipfile.ZIP_DEFLATED)
    try:
        for it, data in out:
            zout.writestr(it, data)
    finally:
        zout.close()
    return hit


def _write_summary_excel(rows, out_path, include_cond=()):
    """把解析结果写成汇总 Excel。

    rows：``[{"字段名": 值, ...}, ...]``，值为 None 表示该票据没有这一栏。
    include_cond：需要输出的条件列集合（未识别到的条件列整体省略）。
    """
    # 懒加载 openpyxl：仅在真正写 Excel 时才导入，缩短启动时间
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter

    _SUMMARY_FILL = PatternFill("solid", fgColor="1F4E78")
    _SUMMARY_FONT = Font(bold=True, color="FFFFFF", size=11)
    _SUMMARY_THIN = Side(style="thin", color="BFBFBF")
    _SUMMARY_BORDER = Border(left=_SUMMARY_THIN, right=_SUMMARY_THIN,
                              top=_SUMMARY_THIN, bottom=_SUMMARY_THIN)
    _SUMMARY_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    # 金额显示格式：不用千分位分隔（避免外部工具解析 / 二次计算出错），仅保留两位小数
    _MONEY_FMT = "0.00"

    # ---- 列布局（全部按列名定位，增删列都不需要改硬编码列号）----
    #   序号 | 明细列 … | 是否重复 | 间隔 | 左块「统计（剔重后）」标签/数值 | 间隔
    #        | 右块「重复票据」标签/数值 | 辅助列（隐藏）
    cond_in_table = [f for f in _SUMMARY_COND_FIELDS if f in include_cond]
    headers = [_SEQ_FIELD] + list(_SUMMARY_BASE_FIELDS) + cond_in_table + ["是否重复"]
    data_fields = headers[:-1]                      # 明细列（不含「是否重复」）
    ncols = len(headers)
    col_of = {name: i + 1 for i, name in enumerate(headers)}
    col_letter = {name: get_column_letter(c) for name, c in col_of.items()}
    # 金额类列（金额合计 + 各支付列）：写数字、套两位小数格式、合计行求和。
    # 序号是整数、其余是文本，都不套金额格式。
    _TEXT_FIELDS = (_SEQ_FIELD, "文件名", "交款人", "票据号码", "开票日期")
    money_fields = [f for f in data_fields if f not in _TEXT_FIELDS]
    deduct_fields = [f for f in _DEDUCT_FIELDS if f in col_of]
    # 两个统计块 + 隐藏在它们右侧的辅助列，列号全部由数据列数推算
    stat_label_col, stat_val_col = ncols + 2, ncols + 3      # 左块「统计（剔重后）」
    dup_label_col, dup_val_col = ncols + 5, ncols + 6        # 右块「重复票据」
    aux_col_idx = ncols + 7                                  # 辅助列（隐藏）
    aux_col = get_column_letter(aux_col_idx)

    wb = Workbook()
    ws = wb.active
    ws.title = "发票汇总"

    ws.append(headers)
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = _SUMMARY_FILL
        cell.font = _SUMMARY_FONT
        cell.alignment = _SUMMARY_CENTER
        cell.border = _SUMMARY_BORDER

    # ---- 同票号分组与去重判定（**全部在程序侧算好**，不交给公式）----
    # ⚠️ 千万不要退回「=IF(AND(票号<>"",COUNTIF(票号累计区间,票号)>1),"重复份","首份")」：
    #    票号通常是 19 位纯数字，而 Excel / WPS / 腾讯文档的 SUMIF / COUNTIF 会把
    #    「长得像数字的文本」内部转成数字再比较，双精度浮点只有 15~16 位有效数字 ——
    #    同前缀的票号会全部塌陷成同一个值、被当成同一张票（实测 16 行 19 位票号里
    #    15 行被判「重复份」，「票据张数」由 13 错成 1）。放在这里用 Counter 判定，
    #    既准确，又天然与「是否重复」列口径一致。
    _ticket_counts = Counter()
    for r in rows:
        tn = r.get(_TICKET_FIELD)
        if tn:
            _ticket_counts[tn] += 1

    # 同票号各自的**明细序号**（就是「序号」列里那个 1、2、3…），用于互指文案
    _dup_group_rows = {}
    for i, r in enumerate(rows):
        tn = r.get(_TICKET_FIELD)
        if tn and _ticket_counts[tn] >= 2:
            _dup_group_rows.setdefault(tn, []).append(i + 1)

    # 每行的「首份 / 重复份」：同票号第 2 张及以后算「重复份」；
    # 空票号（识别失败）不算重复，记「首份」，保证 首份数 + 重复份数 = 明细总行数。
    _seen_ticket = Counter()
    aux_flags = []
    for r in rows:
        tn = r.get(_TICKET_FIELD)
        _seen_ticket[tn] += 1
        aux_flags.append(_AUX_TAG if (tn and _seen_ticket[tn] > 1) else _AUX_FIRST)

    for i, r in enumerate(rows):
        tn = r.get(_TICKET_FIELD)
        # 「是否重复」列：不再只标「是」，改为**互指序号**（序号列里的编号），
        # 一眼能看出与哪几行是同一张票；用序号而不是 Excel 行号，免得跟序号列差 2 对不上。
        if tn and _ticket_counts[tn] >= 2:
            others = [x for x in _dup_group_rows[tn] if x != i + 1]
            dup = "与序号" + "、".join(str(x) for x in others) + "重复"
        else:
            dup = ""
        values = []
        for f in data_fields:
            if f == _SEQ_FIELD:                     # 序号：按明细顺序 1、2、3…
                values.append(i + 1)
                continue
            v = r.get(f)
            values.append(v if v is not None else "")
        ws.append(values + [dup])
        row_idx = ws.max_row
        for c in range(1, ncols + 1):
            cell = ws.cell(row=row_idx, column=c)
            cell.border = _SUMMARY_BORDER
            cell.alignment = _SUMMARY_CENTER
        # 序号列写的是整数，套整数格式（不显示小数）
        ws.cell(row=row_idx, column=col_of[_SEQ_FIELD]).number_format = "0"
        # 金额列写入的是数字，套两位小数金额格式；
        # 空字符串（解析失败/缺失）保持为空，不影响求和。
        for f in money_fields:
            v = ws.cell(row=row_idx, column=col_of[f]).value
            if isinstance(v, (int, float)):
                ws.cell(row=row_idx, column=col_of[f]).number_format = _MONEY_FMT
        # 辅助列：写**文本字面量**（不是公式），供两个统计块的 SUMIF / COUNTIF 使用。
        # 多张同号票据里只有「重复份」计入「重复票据…合计」；
        # 首张已计入「统计（剔重后）」各项，明细行本身全部保留、不做删改。
        ws.cell(row=row_idx, column=aux_col_idx, value=aux_flags[i])

    last = ws.max_row
    if last >= 2:
        # 合计行：各金额列分别求和
        total_row = [""] * ncols
        # 「合计」写在「文件名」列（序号列留空，免得跟序号混在一起）
        total_row[col_of["文件名"] - 1] = "合计"
        for f in money_fields:
            L = col_letter[f]
            total_row[col_of[f] - 1] = f"=SUM({L}2:{L}{last})"
        ws.append(total_row)
        for c in range(1, ncols + 1):
            cell = ws.cell(row=ws.max_row, column=c)
            cell.font = Font(bold=True)
            cell.border = _SUMMARY_BORDER
            cell.alignment = _SUMMARY_CENTER
            cell.fill = PatternFill("solid", fgColor="DDEBF7")
        for f in money_fields:
            ws.cell(row=ws.max_row, column=col_of[f]).number_format = _MONEY_FMT

    # ---- 两个统计块（列号由数据列数推算；行号先按标签顺序算好，再拼公式）----
    stat_vcol = get_column_letter(stat_val_col)
    dup_vcol = get_column_letter(dup_val_col)
    stat_first_row = 2                                # 第 1 行是块标题，数据从第 2 行起
    amount_col = col_letter[_AMOUNT_FIELD]
    aux_rng = f"{aux_col}2:{aux_col}{last}"           # 辅助列（仅数据行区间）
    first_crit = f'"{_AUX_FIRST}"'
    tag_crit = f'"{_AUX_TAG}"'

    # 左块「统计（剔重后）」：**每一项都只统计「首份」** —— 即剔除重复票据之后的口径。
    left_labels = (["票据张数", "合计总金额"]
                   + [_total_label(f) for f in deduct_fields]
                   + ["可赔付金额"])
    left_row = {lab: stat_first_row + i for i, lab in enumerate(left_labels)}
    # 可赔付金额 = 合计总金额 − 各支付类合计（统筹 / 大病保险 / 医疗救助 …
    # 这些钱由基金 / 保险 / 救助支付，患者并未实际支出）
    deduct_cells = [f"{stat_vcol}{left_row[_total_label(f)]}" for f in deduct_fields]
    left_items = [
        ("票据张数", f"=COUNTIF({aux_rng},{first_crit})"),
        ("合计总金额",
         f"=SUMIF({aux_rng},{first_crit},{amount_col}2:{amount_col}{last})"),
    ]
    for f in deduct_fields:
        L = col_letter[f]
        left_items.append(
            (_total_label(f), f"=SUMIF({aux_rng},{first_crit},{L}2:{L}{last})"))
    left_items.append(
        ("可赔付金额",
         f"={stat_vcol}{left_row['合计总金额']}-" + "-".join(deduct_cells)))

    # 右块「重复票据」：只累计「重复份」（同票号第 2 张及以后），首张不重复计入
    right_items = [
        ("重复票据张数", f"=COUNTIF({aux_rng},{tag_crit})"),
        ("重复票据金额合计",
         f"=SUMIF({aux_rng},{tag_crit},{amount_col}2:{amount_col}{last})"),
    ]
    for f in deduct_fields:
        L = col_letter[f]
        right_items.append(
            (_dup_label(f), f"=SUMIF({aux_rng},{tag_crit},{L}2:{L}{last})"))

    count_labels = ("票据张数", "重复票据张数")
    for lcol, vcol, title, items in (
            (stat_label_col, stat_val_col, "统计（剔重后）", left_items),
            (dup_label_col, dup_val_col, "重复票据", right_items)):
        for col, text in ((lcol, title), (vcol, "数值")):
            tc = ws.cell(row=1, column=col, value=text)
            tc.fill = _SUMMARY_FILL
            tc.font = _SUMMARY_FONT
            tc.alignment = _SUMMARY_CENTER
            tc.border = _SUMMARY_BORDER
        for i, (label, formula) in enumerate(items):
            rrow = stat_first_row + i
            lc = ws.cell(row=rrow, column=lcol, value=label)
            lc.font = Font(bold=True, color="1F4E78")
            lc.alignment = Alignment(horizontal="left", vertical="center")
            lc.border = _SUMMARY_BORDER
            vc = ws.cell(row=rrow, column=vcol, value=formula)
            vc.font = Font(bold=True)
            vc.alignment = _SUMMARY_CENTER
            vc.border = _SUMMARY_BORDER
            vc.number_format = "0" if label in count_labels else _MONEY_FMT

    # 辅助列表头 + 隐藏（只为两块统计的 SUMIF / COUNTIF 服务，不影响阅读）
    ac = ws.cell(row=1, column=aux_col_idx, value=_AUX_HEADER)
    ac.font = Font(size=9, color="808080")

    widths = [SEQ_COL_WIDTH, 26, 14, 18, 14, 16, 18]   # 序号 + 基础列
    widths += [18] * len(cond_in_table)                # 条件列
    # 是否重复 | 间隔 | 左块标签 | 左块数值 | 间隔 | 右块标签 | 右块数值 | 辅助列
    widths += [18, 3, 30, 18, 3, 26, 18, 10]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.column_dimensions[aux_col].hidden = True
    ws.freeze_panes = "A2"

    # ---- 把公式结果写回缓存值（openpyxl 不算公式，详见 _inject_formula_cache）----
    def _col_sum(rs, field):
        """按列求和：跳过空串 / None（解析失败或该票没有这一栏）。"""
        return sum(v for v in (r.get(field) for r in rs)
                   if isinstance(v, (int, float)))

    first_rows = [r for r, fl in zip(rows, aux_flags) if fl == _AUX_FIRST]
    tag_rows = [r for r, fl in zip(rows, aux_flags) if fl == _AUX_TAG]
    first_amount = _col_sum(first_rows, _AMOUNT_FIELD)
    first_deduct = {f: _col_sum(first_rows, f) for f in deduct_fields}
    tag_amount = _col_sum(tag_rows, _AMOUNT_FIELD)
    tag_deduct = {f: _col_sum(tag_rows, f) for f in deduct_fields}

    cache = {}
    left_values = ([len(first_rows), first_amount]
                   + [first_deduct[f] for f in deduct_fields]
                   + [first_amount - sum(first_deduct.values())])
    for i, v in enumerate(left_values):
        cache[f"{stat_vcol}{stat_first_row + i}"] = (
            round(v, 2) if isinstance(v, float) else v, False)
    right_values = [len(tag_rows), tag_amount] + [tag_deduct[f] for f in deduct_fields]
    for i, v in enumerate(right_values):
        cache[f"{dup_vcol}{stat_first_row + i}"] = (
            round(v, 2) if isinstance(v, float) else v, False)
    if last >= 2:                                     # 合计行照旧：明细列全部求和（不剔重）
        total_row_idx = last + 1
        cache[f"{amount_col}{total_row_idx}"] = (
            round(_col_sum(rows, _AMOUNT_FIELD), 2), False)
        for f in deduct_fields:
            cache[f"{col_letter[f]}{total_row_idx}"] = (round(_col_sum(rows, f), 2), False)

    wb.save(out_path)
    _inject_formula_cache(out_path, cache)


def summarize_invoices(pdf_folder: str, log=print):
    """汇总 pdf_folder 内（含子目录）的所有 PDF 发票，输出 Excel。

    返回生成的 Excel 路径；无 PDF 或出错返回 None。
    """
    # ⚠️ 用 os.walk 而不是 glob("**/*.pdf")：glob 会把路径里的 [ ] 当字符类，
    #    文件夹名带方括号（如「2026 发票[1]」）时会静默返回空、汇总被跳过。
    pdf_files = []
    for root_dir, _sub, fnames in os.walk(pdf_folder):
        for fn in fnames:
            if fn.lower().endswith(".pdf") and not fn.startswith("发票汇总_"):
                pdf_files.append(os.path.join(root_dir, fn))
    pdf_files.sort(key=str.lower)

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
            rows.append({"文件名": name, "交款人": "解析失败"})
            continue
        row = {"文件名": name}
        for fld in _SUMMARY_BASE_FIELDS[1:] + list(_SUMMARY_COND_FIELDS):
            v = fields.get(fld)
            row[fld] = v if v is not None else ""
        rows.append(row)
        ok += 1
        log(f"  [OK] {name}  交款人={fields['交款人']}  "
            f"票据号={fields['票据号码']}  金额={fields['金额合计（小写）']}")

    # 条件列：本批票据至少有一张识别到该类目才输出，一张都没有则整体忽略
    include_cond = tuple(fld for fld in _SUMMARY_COND_FIELDS
                         if any(r.get(fld) not in (None, "") for r in rows))

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(pdf_folder, f"发票汇总_{ts}.xlsx")
    _write_summary_excel(rows, out_path, include_cond=include_cond)
    log(f"汇总完成：成功 {ok} / 共 {len(pdf_files)} 个。")
    if include_cond:
        for fld in include_cond:
            n = sum(1 for r in rows if r.get(fld) not in (None, ""))
            log(f"识别到「{fld}」的票据 {n} / {ok} 张，已输出该列及其汇总。")
    else:
        log("本批票据未识别到任何条件类支付栏目，汇总表不输出相应列。")
    log(f"已导出：{out_path}")
    return out_path
