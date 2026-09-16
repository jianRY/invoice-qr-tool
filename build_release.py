# -*- coding: utf-8 -*-
"""
发票二维码工具 —— 构建 / 签名 / 发布 一体化脚本
================================================

职责：
  1. 构建「单文件运行版」exe（onefile，复用发票二维码工具.spec 的 Analysis 配置）
  2. 构建「卸载程序」exe（onefile）
  3. 构建「可安装版」exe（onefile，内嵌单文件版主程序 + 卸载程序）
  4. 三个 exe 全部走 sign.py 自签名（SHA256 + RFC3161 时间戳）
  5. 生成 Usage / Changelog 文本
  6. 同步官网 website/ 的版本号 / 下载直链 / 日期 / 体积（宝塔脚本拉取后即展示新版网页）
  7. （--publish 时）git push（走代理）→ 打 tag → 在 GitHub 创建 Release 并上传双 exe

安全护栏：
  - 本地目标版本必须 > GitHub 线上最新版本，否则拒绝发布（防止用旧代码覆盖新版）。
  - 仅在检测到「相对上次发布有新的提交」时才发布。
  - 不打印任何令牌 / 密码；PAT 运行时从 wincred 读取。

用法：
  python build_release.py                # 仅构建 + 签名 + 生成资源（本地验证用）
  python build_release.py --publish      # 构建 + 签名 + 推送 GitHub + 发 Release
  python build_release.py --publish --version 3.9   # 指定版本号
  python build_release.py --publish --force         # 即使版本号不高于线上也发布
  python build_release.py --skip-build   # 跳过 PyInstaller，直接对已有 dist/ 签名+发布
  python build_release.py --publish-only --version 4.0  # 已构建+签名，仅提交/打tag/发Release（不重签）
"""
import os
import re
import sys
import json
import shutil
import subprocess
import urllib.request
import urllib.error
import base64
import uuid
import time

# ---------------- 路径常量 ----------------
ROOT = os.path.dirname(os.path.abspath(__file__))
VENV_PY = os.path.join(ROOT, "envs", "default", "Scripts", "python.exe")
GIT = r"C:\Users\toxuj\.workbuddy\binaries\PortableGit\versions\1.2.0\mingw64\bin\git.exe"
WCRED = r"C:\Users\toxuj\.workbuddy\binaries\PortableGit\versions\1.2.0\mingw64\bin\git-credential-wincred.exe"
SIGN_PY = r"D:\workbuddy\诉讼案件网站\.pybuild_cache\signing\sign.py"
APP_NAME = "发票二维码工具"
PROJECT_URL = "https://github.com/jianRY/invoice-qr-tool"
OWNER = "jianRY"
REPO = "invoice-qr-tool"
APP_EXE = "发票二维码工具.exe"
PORTABLE_OUT = os.path.join(ROOT, "dist", APP_EXE)
UNINST_OUT = os.path.join(ROOT, "dist", "uninstaller.exe")
INSTALLER_OUT = os.path.join(ROOT, "dist", APP_NAME + "_安装程序.exe")
ASSET_DIR = os.path.join(ROOT, "outputs", "release_assets")
WEBSITE_DIR = os.path.join(ROOT, "website")   # 官网静态页（宝塔脚本 git pull 后自动部署）
VERSION_FILE = os.path.join(ROOT, "VERSION")
LAST_RELEASE_COMMIT = os.path.join(ROOT, ".last_release_commit")
PROXY = "http://127.0.0.1:10808"

# 代理全局生效（urllib / git 都用）
os.environ.setdefault("HTTPS_PROXY", PROXY)
os.environ.setdefault("HTTP_PROXY", PROXY)
os.environ.setdefault("https_proxy", PROXY)
os.environ.setdefault("http_proxy", PROXY)


# ---------------- 版本工具 ----------------
def parse_ver(tag):
    if not tag:
        return (0, 0, 0)
    m = re.search(r"(\d+)\.(\d+)(?:\.(\d+))?", str(tag))
    if not m:
        return (0, 0, 0)
    g = m.groups()
    return (int(g[0]), int(g[1]), int(g[2]) if g[2] else 0)


def fmt_ver(v):
    return "v{}.{}".format(v[0], v[1]) + ("" if v[2] == 0 else ".{}".format(v[2]))


def bump_minor(v):
    return (v[0], v[1] + 1, 0)


def read_version():
    if os.path.exists(VERSION_FILE):
        return open(VERSION_FILE, encoding="utf-8").read().strip()
    return "3.6"


def write_version(v):
    with open(VERSION_FILE, "w", encoding="utf-8") as f:
        f.write(v + "\n")


# ---------------- 凭据 ----------------
def get_pat():
    # 1) wincred（与 git 同一凭据，无需落盘）
    try:
        r = subprocess.run(
            [WCRED, "get"],
            input="protocol=https\nhost=github.com\n\n",
            capture_output=True, text=True, timeout=30,
            encoding="utf-8", errors="replace",
        )
        for line in r.stdout.splitlines():
            if line.startswith("password="):
                return line.split("=", 1)[1].strip()
    except Exception:
        pass
    # 2) 回退：本地 .github_token（gitignore）
    f = os.path.join(ROOT, ".github_token")
    if os.path.exists(f):
        return open(f, encoding="utf-8").read().strip()
    return None


# ---------------- Git ----------------
def git(*args, check=True, capture=True):
    env = os.environ.copy()
    # 关键：本环境 PATH 常失效，git 会找不到 git-remote-https（报
    # 'remote-https' is not a git command）。显式指定 GIT_EXEC_PATH 兜底。
    env["GIT_EXEC_PATH"] = os.path.dirname(GIT)
    r = subprocess.run(
        [GIT] + list(args), capture_output=capture, text=True,
        encoding="utf-8", errors="replace", env=env,
    )
    if check and r.returncode != 0:
        raise RuntimeError("git {} 失败:\n{}".format(" ".join(args), r.stderr))
    return r


def current_branch():
    return git("rev-parse", "--abbrev-ref", "HEAD").stdout.strip() or "main"


def head_commit():
    return git("rev-parse", "HEAD").stdout.strip()


# ---------------- GitHub API ----------------
def api(method, url, token=None, json_data=None, raw=None, ctype=None):
    """调用 GitHub REST API，**不因 4xx/5xx 抛异常**，统一返回 (status, body)。

    这一点很关键：发布流程要先 `GET /releases/tags/<tag>` 判断 Release 是否已存在，
    而「不存在」时 GitHub 返回的就是 404。如果这里直接把 HTTPError 抛出去，就永远
    走不到「创建 Release」的分支（v4.1 首次发布即因此中断）。调用方一律用返回的
    status 自行判断。
    """
    req = urllib.request.Request(url, method=method)
    req.add_header("Accept", "application/vnd.github+json")
    req.add_header("X-GitHub-Api-Version", "2022-11-28")
    if token:
        req.add_header("Authorization", "Bearer " + token)
    if raw is not None:
        req.add_header("Content-Type", ctype or "application/octet-stream")
        req.data = raw
    elif json_data is not None:
        req.add_header("Content-Type", "application/json")
        req.data = json.dumps(json_data).encode("utf-8")
    try:
        with urllib.request.urlopen(req, timeout=300) as resp:
            body = resp.read().decode("utf-8", "replace")
            return resp.status, (json.loads(body) if body else {})
    except urllib.error.HTTPError as e:
        try:
            body = e.read().decode("utf-8", "replace")
            data = json.loads(body) if body else {}
        except Exception:
            data = {}
        return e.code, data


def latest_release_tag(token):
    url = "https://api.github.com/repos/{}/{}/releases/latest".format(OWNER, REPO)
    try:
        status, data = api("GET", url, token)
        if status == 200:
            return data.get("tag_name")
    except Exception:
        pass
    return None


def repo_private(token):
    try:
        status, data = api("GET", "https://api.github.com/repos/{}/{}".format(OWNER, REPO), token)
        if status == 200:
            return bool(data.get("private"))
    except Exception:
        pass
    return False


# ---------------- 构建 ----------------
def run_pyinstaller(args):
    # 每次构建用全新的 workpath / distpath，避免 PyInstaller 清空旧 build/ 触发
    # 沙箱的安全删除拦截（批量删除会被阻断）。产物随后用 os.replace 落到 dist/，
    # 整条流水线不执行任何删除操作。
    work = _tmp_build_dir("wp")
    out = _tmp_build_dir("dist")
    cmd = [VENV_PY, "-m", "PyInstaller", "--noconfirm",
           "--workpath", work, "--distpath", out] + args
    print(">>>", " ".join(cmd))
    r = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    if r.returncode != 0:
        print(r.stdout[-2000:])
        print(r.stderr[-2000:])
        raise RuntimeError("PyInstaller 失败：" + " ".join(args[:3]))
    return out


def _tmp_build_dir(sub):
    base = os.path.join(ROOT, "outputs", "_tmp_build")
    p = os.path.join(base, sub + "_" + uuid.uuid4().hex[:8])
    os.makedirs(p, exist_ok=True)
    return p


def sign(exe_path):
    if not os.path.exists(exe_path):
        raise RuntimeError("待签名文件不存在：" + exe_path)
    print(">>> 签名:", exe_path)
    r = subprocess.run(
        [VENV_PY, SIGN_PY, exe_path, APP_NAME, PROJECT_URL],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    if r.returncode != 0:
        print(r.stdout[-1500:]); print(r.stderr[-1500:])
        raise RuntimeError("签名失败：" + exe_path)
    print("    签名输出:", (r.stdout or r.stderr).strip().splitlines()[-1] if (r.stdout or r.stderr).strip() else "ok")


def build_portable():
    """单文件运行版（onefile，复用既有 spec；spec 本身即单文件构建）。"""
    spec = os.path.join(ROOT, "发票二维码工具.spec")
    out = run_pyinstaller([spec])
    os.replace(os.path.join(out, APP_EXE), PORTABLE_OUT)
    sign(PORTABLE_OUT)


def build_uninstaller():
    src = os.path.join(ROOT, "installer_src", "app_uninstaller.py")
    out = run_pyinstaller([
        src, "--onefile", "--name", "uninstaller", "--uac-admin",
    ])
    os.replace(os.path.join(out, "uninstaller.exe"), UNINST_OUT)
    sign(UNINST_OUT)


def build_installer():
    """可安装版：内嵌单文件版主程序 + 卸载程序。"""
    src = os.path.join(ROOT, "installer_src", "app_installer.py")
    add_app = "{}{}app_payload".format(PORTABLE_OUT, os.pathsep)
    add_un = "{}{}.".format(UNINST_OUT, os.pathsep)
    out = run_pyinstaller([
        src, "--onefile", "--name", APP_NAME + "_安装程序",
        "--uac-admin",
        "--add-data", add_app,
        "--add-data", add_un,
    ])
    os.replace(os.path.join(out, APP_NAME + "_安装程序.exe"), INSTALLER_OUT)
    sign(INSTALLER_OUT)


# ---------------- 资源生成 ----------------
def make_assets(new_tag):
    os.makedirs(ASSET_DIR, exist_ok=True)
    # 复制双 exe（ASCII 文件名，GitHub 附件不支持中文）
    portable_name = "InvoiceQRDownloader_{}.exe".format(new_tag.lstrip("v"))
    installer_name = "InvoiceQRInstaller_{}.exe".format(new_tag.lstrip("v"))
    shutil.copy2(PORTABLE_OUT, os.path.join(ASSET_DIR, portable_name))
    shutil.copy2(INSTALLER_OUT, os.path.join(ASSET_DIR, installer_name))

    # Changelog：相对上次发布提交的 git log
    last = ""
    if os.path.exists(LAST_RELEASE_COMMIT):
        last = open(LAST_RELEASE_COMMIT, encoding="utf-8").read().strip()
    if last:
        rng = "{}..HEAD".format(last)
    else:
        rng = "-15"
    log = git("log", "--oneline", rng, check=False).stdout.strip()
    if not log:
        log = git("log", "--oneline", "-15").stdout.strip()
    lines = log.splitlines()
    changelog = "# 发票二维码识别下载工具 {}\n\n".format(new_tag)
    changelog += "## 更新内容\n"
    for ln in lines:
        changelog += "- " + ln + "\n"
    changelog += "\n## 下载说明\n"
    changelog += "- `{}`：单文件运行版（双击即用，无需安装）。\n".format(portable_name)
    changelog += "- `{}`：可安装版（装到 Program Files，开始菜单/桌面快捷方式，含卸载程序）。\n".format(installer_name)
    changelog += "- 软件界面与本地文件名仍为中文 `{}`。\n".format(APP_EXE)
    with open(os.path.join(ASSET_DIR, "InvoiceQR_Changelog.txt"), "w", encoding="utf-8") as f:
        f.write(changelog)

    usage = (
        "发票二维码识别下载工具 {tag}\n"
        "================================\n\n"
        "【单文件运行版】InvoiceQRDownloader_{tag}.exe\n"
        "  直接双击运行，无需安装。程序会在同目录读写配置与输出。\n\n"
        "【可安装版】InvoiceQRInstaller_{tag}.exe\n"
        "  右键「以管理员身份运行」→ 选择安装目录（默认 C:\\Program Files\\发票二维码工具）\n"
        "  → 自动创建开始菜单 / 桌面快捷方式，并写入「应用和功能」卸载项。\n"
        "  卸载：设置 → 应用 → 发票二维码工具 → 卸载，或控制面板。\n\n"
        "两版功能完全一致，按使用场景选择即可。\n"
        "项目主页：{url}\n"
    ).format(tag=new_tag.lstrip("v"), url=PROJECT_URL)
    with open(os.path.join(ASSET_DIR, "InvoiceQR_Usage.txt"), "w", encoding="utf-8") as f:
        f.write(usage)
    print("资源已生成:", ASSET_DIR)
    return portable_name, installer_name


# ---------------- 官网同步 ----------------
def update_website(new_tag):
    """把官网页面（website/）里的版本号 / 下载直链 / 发布日期 / 体积
    同步成本次发布的数据，保证宝塔自动脚本 git pull 后展示的新版网页一致。

    只改「版本相关」的字段，不动文案与版式，幂等可重跑。
    """
    ver = new_tag.lstrip("v")                    # 4.0
    today = time.strftime("%Y-%m-%d")            # 2026-09-16
    try:
        size_mb = round(os.path.getsize(PORTABLE_OUT) / 1_000_000)   # 十进制 MB
    except OSError:
        size_mb = 0

    def rewrite(path, pairs):
        if not os.path.exists(path):
            print("    跳过（缺失）:", os.path.basename(path)); return
        s = open(path, encoding="utf-8").read()
        for pat, rep in pairs:
            s = re.sub(pat, rep, s)
        with open(path, "w", encoding="utf-8") as f:
            f.write(s)

    # 首页：徽章「当前版本 vX.Y」、数据卡「vX.Y」、结尾按钮「免费下载 vX.Y」
    rewrite(os.path.join(WEBSITE_DIR, "index.html"), [
        (r"当前版本 v[\d.]+", "当前版本 " + new_tag),
        (r'(class="v"[^>]*>)\s*v[\d.]+', r"\g<1>" + new_tag),
        (r"免费下载 v[\d.]+", "免费下载 " + new_tag),
    ])
    # 下载页：meta 描述、版本 chip、日期 chip、体积 chip、主下载直链（含无 v 前缀的真实附件名）
    rewrite(os.path.join(WEBSITE_DIR, "download.html"), [
        (r"v[\d.]+（Windows", new_tag + "（Windows"),
        (r'(<span class="chip">)v[\d.]+', r"\g<1>" + new_tag),
        (r'(<span class="chip gray">)\d{4}-\d{2}-\d{2}', r"\g<1>" + today),
        (r'(<span class="chip gray">≈ )\d+( MB)', r"\g<1>" + str(size_mb) + r"\g<2>"),
        (r"releases/download/v[\d.]+/InvoiceQRDownloader[_v]*[\d.]+\.exe",
         "releases/download/{}/InvoiceQRDownloader_{}.exe".format(new_tag, ver)),
        (r"releases/download/v[\d.]+/InvoiceQRInstaller[_v]*[\d.]+\.exe",
         "releases/download/{}/InvoiceQRInstaller_{}.exe".format(new_tag, ver)),
    ])
    print("[网页] 官网版本数据已同步 -> {} | 日期 {} | {} MB".format(new_tag, today, size_mb))


# ---------------- 发布 ----------------
def publish(new_tag, token):
    # 1) 仓库公开（匿名读取 Release 需要）
    if repo_private(token):
        api("PATCH", "https://api.github.com/repos/{}/{}".format(OWNER, REPO), token,
            json_data={"private": False})
        print("[发布] 仓库已设为公开")
    # 2) 推送
    #    临时把令牌内联进 remote URL（避免弹凭据窗），用完**立刻**改回干净地址，
    #    绝不让 PAT 留在 .git/config 里。后续 push 走 wincred，无需内联。
    branch = current_branch()
    pat = get_pat()
    clean_url = "https://github.com/{}/{}.git".format(OWNER, REPO)
    remote_url = "https://{}@github.com/{}/{}.git".format(pat, OWNER, REPO)
    try:
        git("remote", "set-url", "origin", remote_url, check=False)
        git("push", "origin", branch, check=False)
        git("push", "origin", new_tag, check=False)
    finally:
        git("remote", "set-url", "origin", clean_url, check=False)
    print("[发布] 已推送分支 {} 与标签 {}".format(branch, new_tag))
    # 3) 获取或创建 Release（幂等：已存在则复用，便于补传附件 / 重跑）
    body = open(os.path.join(ASSET_DIR, "InvoiceQR_Changelog.txt"), encoding="utf-8").read()
    rel_url = "https://api.github.com/repos/{}/{}/releases".format(OWNER, REPO)
    status, data = api("GET", rel_url + "/tags/" + new_tag, token)
    if status == 200 and data.get("id"):
        release_id = data["id"]
        upload_url = data.get("upload_url", "").split("{")[0]
        print("[发布] 复用已有 Release {} (HTTP {})".format(new_tag, status))
    else:
        status, data = api(
            "POST", rel_url, token,
            json_data={"tag_name": new_tag, "name": "发票二维码识别下载工具 " + new_tag,
                       "body": body, "draft": False, "prerelease": False},
        )
        release_id = data.get("id")
        upload_url = data.get("upload_url", "").split("{")[0]
        print("[发布] 创建 Release {} (HTTP {})".format(new_tag, status))
    if not release_id:
        raise RuntimeError(
            "创建/复用 Release {} 失败：HTTP {}，响应 {}".format(new_tag, status, str(data)[:400])
        )
    # 4) 上传附件（先删同名旧附件，保证幂等可重跑）
    from urllib.parse import quote
    _, assets_data = api("GET", "https://api.github.com/repos/{}/{}/releases/{}/assets".format(OWNER, REPO, release_id), token)
    existing = {a.get("name"): a.get("id") for a in (assets_data or [])}
    for fn, ctype in [
        ("InvoiceQRDownloader_{}.exe".format(new_tag.lstrip("v")), "application/octet-stream"),
        ("InvoiceQRInstaller_{}.exe".format(new_tag.lstrip("v")), "application/octet-stream"),
        ("InvoiceQR_Usage.txt", "text/plain; charset=utf-8"),
        ("InvoiceQR_Changelog.txt", "text/plain; charset=utf-8"),
    ]:
        p = os.path.join(ASSET_DIR, fn)
        if not os.path.exists(p):
            print("    跳过（缺失）:", fn); continue
        if fn in existing:
            api("DELETE", "https://api.github.com/repos/{}/{}/releases/assets/{}".format(OWNER, REPO, existing[fn]), token)
            print("    删除旧附件:", fn)
        with open(p, "rb") as fh:
            raw = fh.read()
        url = "{}?name={}&label={}".format(upload_url, quote(fn), quote(fn))
        st, _ = api("POST", url, token, raw=raw, ctype=ctype)
        print("    上传 {} ({}KB) -> HTTP {}".format(fn, len(raw) // 1024, st))
        if st not in (200, 201):
            raise RuntimeError("上传附件 {} 失败：HTTP {}".format(fn, st))
    # 5) 记录已发布提交
    with open(LAST_RELEASE_COMMIT, "w", encoding="utf-8") as f:
        f.write(head_commit() + "\n")
    git("add", "-A", check=False)
    git("commit", "-m", "chore: 记录 {} 发布提交".format(new_tag), check=False)
    git("push", "origin", branch, check=False)
    print("[发布] 完成。地址：https://github.com/{}/{}/releases/tag/{}".format(OWNER, REPO, new_tag))


# ---------------- 主流程 ----------------
def main():
    args = sys.argv[1:]
    do_publish = ("--publish" in args) or ("--publish-only" in args)
    force = "--force" in args
    skip_build = "--skip-build" in args
    publish_only = "--publish-only" in args
    ver_override = None
    if "--version" in args:
        i = args.index("--version")
        if i + 1 < len(args):
            ver_override = args[i + 1]

    token = get_pat() if do_publish else None
    latest = latest_release_tag(token) if do_publish else None
    latest_v = parse_ver(latest)
    local_v = parse_ver(read_version())

    # 目标版本：显式指定 > 自动在线上最新版基础上 +1（满足“版本号按实际情况新增”）
    if ver_override:
        new_v = parse_ver(ver_override)
    elif do_publish and latest_v > (0, 0, 0):
        new_v = bump_minor(latest_v)
    else:
        new_v = local_v if local_v > (0, 0, 0) else (3, 7, 0)
    new_tag = fmt_ver(new_v)
    print("目标版本:", new_tag, "| 线上最新:", latest or "(未知)")

    # 自动化效率：相对上次发布无新提交则直接跳过（不构建）
    if do_publish:
        last = (open(LAST_RELEASE_COMMIT, encoding="utf-8").read().strip()
                if os.path.exists(LAST_RELEASE_COMMIT) else "")
        if last and head_commit() == last and not force:
            print("[护栏] 相对上次发布无新提交，跳过（无需构建/发布）。")
            return
        # 安全护栏：目标版本不高于线上最新则拒绝发布（防止旧代码覆盖新版）
        if not force and new_v <= latest_v:
            print("[护栏] 目标版本 {} 不高于线上最新 {}，跳过发布。"
                  "如需覆盖请用 --force 或先 --version 指定更高版本。".format(new_tag, latest))
            return

    # 构建
    if publish_only:
        print("[publish-only] 跳过构建与签名，直接使用现有 dist/ 已签名 exe")
    elif not skip_build:
        print("=== 1/3 构建单文件运行版 ===")
        build_portable()
        time.sleep(3)   # 给时间戳服务器留出余量，避免连续签名被限流
        print("=== 2/3 构建卸载程序 ===")
        build_uninstaller()
        time.sleep(3)
        print("=== 3/3 构建可安装版（内嵌主程序+卸载程序）===")
        build_installer()
    else:
        sign(PORTABLE_OUT); sign(UNINST_OUT); sign(INSTALLER_OUT)

    for p in (PORTABLE_OUT, UNINST_OUT, INSTALLER_OUT):
        if not os.path.exists(p):
            raise RuntimeError("构建产物缺失：" + p)
        print("产物:", p, "{}KB".format(os.path.getsize(p) // 1024))

    make_assets(new_tag)
    update_website(new_tag)   # 同步官网版本数据（宝塔脚本拉取后展示的新版网页）

    if do_publish:
        # 仅当有新提交才发布
        last = open(LAST_RELEASE_COMMIT, encoding="utf-8").read().strip() if os.path.exists(LAST_RELEASE_COMMIT) else ""
        if last and head_commit() == last and not force:
            print("[护栏] 相对上次发布无新提交，跳过发布。")
            return
        if token is None:
            print("[错误] 未能获取 GitHub PAT，无法发布。请确认 wincred 中 git:https://github.com 凭据。")
            return
        write_version(new_tag.lstrip("v"))
        git("add", "-A", check=False)
        git("commit", "-m", "release: {}".format(new_tag), check=False)
        git("tag", new_tag, check=False)
        publish(new_tag, token)
    else:
        print("\n[本地验证完成] 未执行发布（需要 --publish）。产物与签名均已就绪。")


if __name__ == "__main__":
    main()
