#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""仅把附件上传到已存在的 Release（幂等、可重试）。
用于 Release 已创建但附件上传被网络重置的情况。
用法：python upload_only.py <PAT>
"""
import sys
import os
import time
import importlib.util
import urllib.request
import urllib.error
from urllib.parse import quote

spec = importlib.util.spec_from_file_location("pub", os.path.join(os.path.dirname(os.path.abspath(__file__)), "publish_v3.py"))
pub = importlib.util.module_from_spec(spec)
spec.loader.exec_module(pub)


def upload_asset_retry(token, upload_url, filename, content_type, retries=4):
    path = os.path.join(pub.ASSET_DIR, filename)
    if not os.path.isfile(path):
        raise RuntimeError(f"附件不存在：{path}")
    with open(path, "rb") as f:
        raw = f.read()
    for attempt in range(1, retries + 1):
        url = f"{upload_url}?name={quote(filename)}&label={quote(filename)}"
        req = urllib.request.Request(url, method="POST", data=raw)
        req.add_header("Authorization", f"Bearer {token}")
        req.add_header("Accept", "application/vnd.github+json")
        req.add_header("X-GitHub-Api-Version", "2022-11-28")
        req.add_header("Content-Type", content_type)
        try:
            with urllib.request.urlopen(req, timeout=600) as resp:
                print(f"  上传 {filename} ({len(raw)//1024}KB) -> HTTP {resp.status}")
                return True
        except Exception as e:
            print(f"  上传 {filename} 尝试 {attempt}/{retries} 失败：{e}")
            time.sleep(8)
    raise RuntimeError(f"上传 {filename} 最终失败")


def main():
    if len(sys.argv) < 2:
        print("用法：python upload_only.py <PAT>")
        sys.exit(1)
    token = sys.argv[1].strip()

    # 1) 取已存在 Release 的 upload_url 与现有 assets
    url = f"https://api.github.com/repos/{pub.OWNER}/{pub.REPO}/releases/tags/{pub.TAG}"
    status, data = pub.api("GET", url, token)
    if status != 200:
        raise RuntimeError(f"获取 Release 失败 HTTP {status}：{data}")
    upload_url = data.get("upload_url", "").split("{")[0]
    existing = {a["name"]: a["id"] for a in data.get("assets", [])}
    print(f"[Release] {pub.TAG} id={data.get('id')} 已存在附件={list(existing)}")

    # 2) 逐个上传（若同名已存在则先删除，保证幂等）
    for name, ctype in pub.ASSETS:
        if name in existing:
            del_url = (f"https://api.github.com/repos/{pub.OWNER}/{pub.REPO}"
                       f"/releases/assets/{existing[name]}")
            try:
                pub.api("DELETE", del_url, token)
                print(f"  先删除已存在同名附件 {name}")
            except Exception as e:
                print(f"  删除 {name} 警告：{e}")
        print(f"[上传] {name}")
        upload_asset_retry(token, upload_url, name, ctype)
    print("[完成] 附件已全部上传。Release 地址：")
    print(f"      https://github.com/{pub.OWNER}/{pub.REPO}/releases/tag/{pub.TAG}")


if __name__ == "__main__":
    main()
