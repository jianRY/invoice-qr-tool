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
out = {}
for node in ast.walk(tree):
    if isinstance(node, ast.Assign):
        for t in node.targets:
            if isinstance(t, ast.Name) and t.id in ("USAGE_TEXT", "CHANGELOG_TEXT"):
                if isinstance(node.value, ast.Constant) and isinstance(node.value.value, str):
                    out[t.id] = node.value.value

written = []
# 本地中文版
for key, fname in (("USAGE_TEXT", "使用说明.txt"),
                   ("CHANGELOG_TEXT", "更新日志.txt")):
    path = os.path.join(ROOT, "outputs", fname)
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

