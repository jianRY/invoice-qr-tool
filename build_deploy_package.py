#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 website/ 内容 + deploy/deploy.sh + deploy/DEPLOY.md 打包成可一键部署的 zip。

zip 顶层目录为 invoice-qr-tool-website/，结构：
  invoice-qr-tool-website/
    index.html  download.html  assets/  downloads/  deploy.sh  DEPLOY.md
注意：deploy.sh / DEPLOY.md 属内部文件，放在项目根 deploy/ 目录，不进 website/（公开目录）。
"""
import os
import zipfile

ROOT = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(ROOT, "website")
DEPLOY_SRC = os.path.join(ROOT, "deploy")
OUT = os.path.join(ROOT, "invoice-qr-tool-website.zip")
TOP = "invoice-qr-tool-website"

# 文本类用压缩，二进制 exe 用存储（已是压缩格式，再压无意义且慢）
COMPRESS_EXT = {".html", ".css", ".md", ".sh", ".txt", ".png"}

def walk(root):
    for dirpath, dirnames, filenames in os.walk(root):
        # 跳过隐藏目录（如 .workbuddy）
        dirnames[:] = [d for d in dirnames if not d.startswith(".")]
        for fn in filenames:
            yield os.path.join(dirpath, fn)

def add(z, full, arcname):
    ext = os.path.splitext(os.path.basename(full))[1].lower()
    z.write(full, arcname,
            compress_type=zipfile.ZIP_DEFLATED if ext in COMPRESS_EXT else zipfile.ZIP_STORED)

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    count = 0
    total = 0
    for full in walk(SRC):
        rel = os.path.relpath(full, SRC)
        arcname = f"{TOP}/{rel.replace(os.sep, '/')}"
        add(z, full, arcname)
        count += 1
        total += os.path.getsize(full)
    # 内部部署文件（放在 zip 顶层）
    for fn in ("deploy.sh", "DEPLOY.md"):
        full = os.path.join(DEPLOY_SRC, fn)
        if os.path.exists(full):
            z.write(full, f"{TOP}/{fn}")
            count += 1
            total += os.path.getsize(full)
    # 报告
    print(f"已打包 {count} 个文件，总大小 {total/1024/1024:.1f} MB -> {OUT}")

# 列出 zip 顶层结构
with zipfile.ZipFile(OUT) as z:
    names = z.namelist()
    print("zip 内顶层条目：")
    seen = set()
    for n in names:
        first = n[len(TOP)+1:].split("/")[0] if n.startswith(TOP + "/") else n
        if first and first not in seen:
            seen.add(first)
            print("  ", first)
    print(f"zip 总条目数：{len(names)}")
