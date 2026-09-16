#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把发票二维码工具发布到 GitHub：
1. 将仓库 jianRY/invoice-qr-tool 设为公开（自动更新需要匿名读取 Release）；
2. 创建对应版本的 Release；
3. 上传三个 ASCII 命名的附件（避免中文名被剥离为 default）。

用法：
  python publish_v3.py <PAT>

PAT 需要 repo 权限（public_repo 或完整 repo scope）。
"""

import sys
import os
import json
import base64
import mimetypes
import urllib.request
import urllib.error

OWNER = "jianRY"
REPO = "invoice-qr-tool"
TAG = "v3.6"
RELEASE_NAME = "v3.6 - 修复金额识别(千位符/大写兜底) + 重构图片归类 + 汇总Excel重复列"
ASSET_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "outputs", "release_assets")

ASSETS = [
    ("InvoiceQRDownloader_v3.6.exe", "application/octet-stream"),
    ("InvoiceQR_Usage.txt", "text/plain; charset=utf-8"),
    ("InvoiceQR_Changelog.txt", "text/plain; charset=utf-8"),
]

RELEASE_BODY = """# 发票二维码识别下载工具 v3.6

## 修复：金额识别 bug
- 「金额合计（小写）」若含千位分隔符（逗号，如 `1,234.56`），此前会被当作结束符，
  只识别到 `1` 导致金额残缺。现已支持中英文逗号，完整识别到小数点后两位。
- 新增「金额合计（大写）」兜底：小写金额解析失败时用大写金额补回，避免漏识别。

## 重构：图片识别与文件归类
- 去除「重复网址」判定，**全部图片均参与识别**。
- 成功识别并下载到 PDF 的图片：**原文件名保持不变**。
- 未下载 / 未识别 / 其它三类：仅「复制」一份到自动新建的 `未识别` 子文件夹，
  分别加前缀 `未下载-` / `未识别-` / `其它-`，**原图片始终保留不动**。

## 增强：汇总 Excel
- 新增「是否重复」列：同一票据号码出现 ≥2 次则标记「是」。
- 右侧统计区新增「重复票据金额合计 / 重复票据统筹合计」
  （仅对标记为「是」的票据求和）。

## 沿用（v3.5 / v3.5.1）
- 自动更新：下载最新 EXE 保存到同目录（带版本号），旧版移入回收站，自动切换。
- 右侧四项统计：票据张数 / 合计总金额 / 合计总统筹金额 / 可赔付金额。
- 金额、统筹列为数字格式，可直接求和。

## 下载说明
- `InvoiceQRDownloader_v3.6.exe` 为主程序（英文文件名因 GitHub 附件不支持中文；
  软件界面与本地文件仍为中文名 `发票二维码工具.exe`）。
- 旧版用户可直接下载新版覆盖，或通过软件内「检查更新」自动升级。
"""


def api(method, url, token, data=None, content_type=None, raw_body=None):
    req = urllib.request.Request(url, method=method)
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Accept", "application/vnd.github+json")
    req.add_header("X-GitHub-Api-Version", "2022-11-28")
    if raw_body is not None:
        req.add_header("Content-Type", content_type or "application/octet-stream")
        req.data = raw_body
    elif data is not None:
        req.add_header("Content-Type", "application/json")
        req.data = json.dumps(data).encode("utf-8")
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            body = resp.read().decode("utf-8", "replace")
            return resp.status, json.loads(body) if body else {}
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", "replace")
        raise RuntimeError(f"HTTP {e.code} {method} {url}\n{detail}")


def make_repo_public(token):
    url = f"https://api.github.com/repos/{OWNER}/{REPO}"
    status, data = api("PATCH", url, token, data={"private": False})
    print(f"[1/3] 仓库可见性：{'private' if data.get('private') else 'public'} (HTTP {status})")


def create_release(token):
    url = f"https://api.github.com/repos/{OWNER}/{REPO}/releases"
    status, data = api(
        "POST",
        url,
        token,
        data={
            "tag_name": TAG,
            "name": RELEASE_NAME,
            "body": RELEASE_BODY,
            "draft": False,
            "prerelease": False,
        },
    )
    upload_url = data.get("upload_url", "").split("{")[0]
    print(f"[2/3] 创建 Release {TAG} (HTTP {status}) upload_url={upload_url}")
    return upload_url


def upload_asset(token, upload_url, filename, content_type):
    path = os.path.join(ASSET_DIR, filename)
    if not os.path.isfile(path):
        raise RuntimeError(f"附件不存在：{path}")
    with open(path, "rb") as f:
        raw = f.read()
    # GitHub 要求 URL 中 name 做百分号编码
    from urllib.parse import quote

    url = f"{upload_url}?name={quote(filename)}&label={quote(filename)}"
    status, data = api(
        "POST", url, token, raw_body=raw, content_type=content_type
    )
    print(f"      上传 {filename} ({len(raw)//1024}KB) -> HTTP {status}")


def main():
    if len(sys.argv) < 2:
        print("用法：python publish_v3.py <PAT>")
        sys.exit(1)
    token = sys.argv[1].strip()
    make_repo_public(token)
    upload_url = create_release(token)
    for name, ctype in ASSETS:
        upload_asset(token, upload_url, name, ctype)
    print("[3/3] 完成。Release 地址：")
    print(f"      https://github.com/{OWNER}/{REPO}/releases/tag/{TAG}")


if __name__ == "__main__":
    main()
