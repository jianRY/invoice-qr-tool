#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从 invoice_qr_tool.py 抽取 USAGE_TEXT / CHANGELOG_TEXT（用 ast，不执行模块），
同步写入两处：
  1) outputs/使用说明.txt、outputs/更新日志.txt        —— 本地中文版（给本地用户看）
  2) outputs/release_assets/InvoiceQR_Usage.txt、
     outputs/release_assets/InvoiceQR_Changelog.txt     —— 发布用 ASCII 版（GitHub 附件）
用 AST 抽取可保证「源码/本地/Release」三处说明永远一致，避免手动改漏。"""
import ast
import os

ROOT = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(ROOT, "invoice_qr_tool.py")
ASSET_DIR = os.path.join(ROOT, "outputs", "release_assets")
os.makedirs(ASSET_DIR, exist_ok=True)

tree = ast.parse(open(SRC, encoding="utf-8").read())


def _str_value(node, ns):
    """把「纯字符串常量 / f-string / 相邻字符串相加」求成文本，不执行任何代码。

    USAGE_TEXT 现为 f-string（顶部插入仓库地址时引用了 GITHUB_REPO_* 常量），
    旧的 ast.Constant 判断取不到值，故在此补上 JoinedStr / BinOp 两种形态。
    """
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        return node.value
    if isinstance(node, ast.JoinedStr):
        parts = []
        for v in node.values:
            if isinstance(v, ast.Constant):
                parts.append(str(v.value))
            elif isinstance(v, ast.FormattedValue) and isinstance(v.value, ast.Name):
                parts.append(str(ns.get(v.value.id, "")))
            else:
                return None
        return "".join(parts)
    if isinstance(node, ast.BinOp) and isinstance(node.op, ast.Add):
        left = _str_value(node.left, ns)
        right = _str_value(node.right, ns)
        if left is not None and right is not None:
            return left + right
    return None


# 第一遍：收集模块级字符串常量，供 f-string 插值使用（如 GITHUB_REPO_OWNER）
ns = {}
for node in tree.body:
    if isinstance(node, ast.Assign):
        for t in node.targets:
            if isinstance(t, ast.Name):
                val = _str_value(node.value, ns)
                if val is not None:
                    ns[t.id] = val

# 第二遍：抽取目标文案
out = {}
for node in tree.body:
    if isinstance(node, ast.Assign):
        for t in node.targets:
            if isinstance(t, ast.Name) and t.id in ("USAGE_TEXT", "CHANGELOG_TEXT"):
                val = _str_value(node.value, ns)
                if val is not None:
                    out[t.id] = val

for key in ("USAGE_TEXT", "CHANGELOG_TEXT"):
    if key not in out or not out[key].strip():
        raise SystemExit(f"抽取失败：源码中未找到可解析的 {key}（请检查其赋值形式）")

written = []
# 本地中文版：根目录（仓库内文档）与 outputs/ 都写，避免两处长期漂移
for key, fname in (("USAGE_TEXT", "使用说明.txt"),
                   ("CHANGELOG_TEXT", "更新日志.txt")):
    for base in (ROOT, os.path.join(ROOT, "outputs")):
        path = os.path.join(base, fname)
        open(path, "w", encoding="utf-8").write(out[key])
        written.append((path, len(out[key])))
# 发布 ASCII 版
for key, fname in (("USAGE_TEXT", "InvoiceQR_Usage.txt"),
                   ("CHANGELOG_TEXT", "InvoiceQR_Changelog.txt")):
    path = os.path.join(ASSET_DIR, fname)
    open(path, "w", encoding="utf-8").write(out[key])
    written.append((path, len(out[key])))

for path, n in written:
    print(f"写入 {path} ({n} 字符)")

