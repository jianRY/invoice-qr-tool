import os, sys, json
import urllib.request
import urllib.parse

OWNER = "jianRY"
REPO = "invoice-qr-tool"
TAG = "v3.5.1"
API = "https://api.github.com"
ASSET_NAME = "InvoiceQR_Changelog.txt"
LOCAL = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                     "outputs", "release_assets", ASSET_NAME)


def req(method, url, token, data=None, headers=None, retries=3):
    last_err = None
    for i in range(retries):
        try:
            h = {"Authorization": f"Bearer {token}",
                 "Accept": "application/vnd.github+json",
                 "User-Agent": "invoice-qr-publisher"}
            if headers:
                h.update(headers)
            body = data if data is not None else None
            r = urllib.request.Request(url, data=body, headers=h, method=method)
            with urllib.request.urlopen(r, timeout=600) as resp:
                code = resp.getcode()
                txt = resp.read().decode("utf-8", "replace")
                return code, txt
        except urllib.error.HTTPError as e:
            last_err = (e.code, e.read().decode("utf-8", "replace"))
            if e.code in (401, 403):
                return last_err  # auth/perm error, no point retrying
            print(f"  [retry {i+1}] HTTP {e.code}")
        except Exception as e:  # network blip (10054 etc.)
            last_err = str(e)
            print(f"  [retry {i+1}] {e}")
    return None, last_err


def main():
    token = sys.argv[1]
    # 1) verify token
    code, txt = req("GET", f"{API}/user", token)
    if code != 200:
        print(f"PAT 校验失败：HTTP {code} -> {txt[:200]}")
        sys.exit(2)
    login = json.loads(txt).get("login")
    print(f"PAT 有效，登录用户：{login}")

    # 2) get release by tag
    code, txt = req("GET", f"{API}/repos/{OWNER}/{REPO}/releases/tags/{TAG}", token)
    if code != 200:
        print(f"获取 Release {TAG} 失败：HTTP {code} -> {txt[:200]}")
        sys.exit(3)
    rel = json.loads(txt)
    rel_id = rel["id"]
    print(f"找到 Release id={rel_id} ({rel['tag_name']})")

    # 3) delete existing changelog asset if present
    for a in rel.get("assets", []):
        if a["name"] == ASSET_NAME:
            code, _ = req("DELETE", a["url"], token)
            print(f"删除旧资产 {ASSET_NAME}: HTTP {code}")

    # 4) upload new changelog
    # 注意：GitHub 资产上传必须用 uploads.github.com 域名（api.github.com 会 404）
    up_url = (f"https://uploads.github.com/repos/{OWNER}/{REPO}/releases/{rel_id}/assets"
              f"?name={urllib.parse.quote(ASSET_NAME)}")
    with open(LOCAL, "rb") as f:
        data = f.read()
    code, txt = req("POST", up_url, token,
                    data=data,
                    headers={"Content-Type": "text/plain; charset=utf-8"})
    if code == 201:
        nm = json.loads(txt).get("name")
        print(f"上传成功：{nm} (HTTP 201)")
    else:
        print(f"上传失败：HTTP {code} -> {txt[:300]}")
        sys.exit(4)

    # 5) anonymous verify
    import subprocess
    print("=== 匿名校验 ===")
    curl = ('curl -s "https://api.github.com/repos/%s/%s/releases/tags/%s" '
            '| python -c "import sys,json;d=json.load(sys.stdin);'
            'print(\'assets:\',[a[\'name\'] for a in d.get(\'assets\',[])])"' %
            (OWNER, REPO, TAG))
    os.system(curl)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法：python upload_changelog.py <PAT>")
        sys.exit(1)
    main()
