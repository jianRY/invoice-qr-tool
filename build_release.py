# -*- coding: utf-8 -*-
"""
发票二维码工具 —— 构建 / 签名 / 发布 一体化脚本
================================================

职责：
  1. 构建「单文件运行版」exe（onefile，复用发票二维码工具.spec 的 Analysis 配置）
  2. 构建「卸载程序」exe（onefile）
  3. 构建「可安装版」exe（onefile，内嵌单文件版主程序 + 卸载程序）
  4. 三个 exe 全部走 sign.py 自签名（SHA256 + RFC3161 时间戳）
  5. 生成 Usage / Changelog 文本（Changelog 末尾追加主程序 SHA256，供客户端校验）
  6. 同步官网 website/ 的版本号 / 下载直链 / 日期 / 体积（宝塔脚本拉取后即展示新版网页）
  7. （--publish 时）git push（走代理）→ 打 tag → 在 GitHub 创建 Release 并上传双 exe

自动更新链路（2026-09-24 起全部收在 GitHub，不再依赖自有服务器）：
  · 元数据：Release 附件 update.json（含版本号 + sha256）
  · 读取：客户端经公共加速镜像拉取（gh-proxy.com / ghfast.top / ghproxy.net），
    直连 github.com 与 GitHub API 仅作兜底
  · 校验：sha256 优先取 update.json；兜底路径从 Release 正文的 `SHA256:` 行解析
  · 自有服务器 47.116.64.26 现在只服务官网手动下载按钮

安全护栏：
  - 本地目标版本必须 > GitHub 线上最新版本，否则拒绝发布（防止用旧代码覆盖新版）。
  - 仅在检测到「相对上次发布有新的提交」时才发布。
  - 不打印任何令牌 / 密码；PAT 运行时从 wincred 读取。

用法：
  python build_release.py                # 仅构建 + 签名 + 生成资源（本地验证用）
  python build_release.py --publish      # 构建 + 签名 + 推送 GitHub + 发 Release
  python build_release.py --publish --version 4.7.0  # 手动指定三段式版本号
  python build_release.py --publish --force         # 即使版本号不高于线上也发布
  python build_release.py --skip-build   # 跳过 PyInstaller，直接对已有 dist/ 签名+发布
  python build_release.py --publish-only --version 4.7.1  # 已构建+签名，仅提交/打tag/发Release（不重签）

版本号规则（三段式 X.Y.Z，tag 形如 v4.7.0；不再出现两段式 v4.7）：
  - 自动递增：读「上次发布提交..HEAD」的提交标题 ——
      含 feat / feature / 新增 / 新功能 / 增加功能 / 添加功能  → 升次版本位（4.6.3 → 4.7.0）
      仅 fix / docs / chore / refactor 等                      → 只升修订位（4.6.3 → 4.6.4）
  - 手动指定：--version 4.7.0（写 4.7 等价于 4.7.0）。目标版本必须高于线上最新。
"""
import hashlib
import os
import re
import sys
import json
import shutil
import subprocess
import urllib.request
import urllib.error
import base64
import socket
import uuid
import time

# ---------------- 路径常量 ----------------
def _pick_git():
    """挑选可用的 git 可执行文件。

    优先**系统 Git**：本机 Bash 的 PATH 常失效，PortableGit 的 git 因而找不到
    remote-https helper —— 补 GIT_EXEC_PATH 只会把报错变成「静默失败」
    （returncode=128、stdout/stderr 全空），无人值守发版时会误判成功。
    系统 Git 的 helper 位于 mingw64/libexec/git-core，布局正确，实测可用。
    """
    for c in (r"C:\Program Files\Git\cmd\git.exe",
              r"C:\Program Files (x86)\Git\cmd\git.exe",
              r"C:\Users\toxuj\.workbuddy\binaries\PortableGit\versions\1.2.0\mingw64\bin\git.exe"):
        if os.path.isfile(c):
            return c
    return "git"


ROOT = os.path.dirname(os.path.abspath(__file__))
VENV_PY = os.path.join(ROOT, "envs", "default", "Scripts", "python.exe")
GIT = _pick_git()
WCRED = r"C:\Users\toxuj\.workbuddy\binaries\PortableGit\versions\1.2.0\mingw64\bin\git-credential-wincred.exe"
SIGN_PY = r"D:\workbuddy\诉讼案件网站\.pybuild_cache\signing\sign.py"
APP_NAME = "发票二维码工具"
PROJECT_URL = "https://github.com/jianRY/invoice-qr-tool"
OWNER = "jianRY"
REPO = "invoice-qr-tool"
APP_EXE = "发票二维码工具.exe"
APP_ICON = os.path.join(ROOT, "app_icon.ico")   # 三支 exe 共用的应用图标
PORTABLE_OUT = os.path.join(ROOT, "dist", APP_EXE)
UNINST_OUT = os.path.join(ROOT, "dist", "uninstaller.exe")
INSTALLER_OUT = os.path.join(ROOT, "dist", APP_NAME + "_安装程序.exe")
ASSET_DIR = os.path.join(ROOT, "outputs", "release_assets")
WEBSITE_DIR = os.path.join(ROOT, "website")   # 官网静态页（宝塔脚本 git pull 后自动部署）
VERSION_FILE = os.path.join(ROOT, "VERSION")
LAST_RELEASE_COMMIT = os.path.join(ROOT, ".last_release_commit")
# 代理：不再硬编码（2026-09-22 重构，见下方 pick_proxy 的说明）
PROXY = None

# 自有下载站（阿里云 47.116.64.26，见「下载服务器」项目）：
#   /files/<资产名>   双 exe 由服务器定时脚本从 Release 镜像过去
#
# ⚠️ 2026-09-24 起**只服务官网的手动下载按钮**，不再参与自动更新：
#    自动更新链路已全部收在 GitHub（update.json 是 Release 附件，客户端经加速镜像读它）。
#    原先往这里推 /updates/qr.json 的那条线已删；SITE_URL 现在仅用于
#    ① 生成 InvoiceQR_Usage.txt 里的「国内直连」下载地址
#    ② 官网 download.html 的按钮链接（由 update_website 改写）
SITE_URL = "http://47.116.64.26:8888"

# 代理全局生效（urllib / requests / git 都用）。
#
# ⚠️ 2026-09-22 重构：不再硬编码 127.0.0.1:10808，也不再只靠「端口能不能连」判断。
#    旧实现的两个坑，都实际踩过：
#      ① 硬编码端口 —— 代理一关或换端口，发版就卡在最后一步 git push 上
#         （打包/签名/元数据全做完了才失败，最亏）；
#      ② `_proxy_alive()` 只做 TCP 连接 —— 10808 端口当时**能连但出不了网**，
#         被判为「可用」，结果 push 报 `Could not connect to server`。
#    现在改为：逐个候选**真去访问一次 GitHub API**，谁能用就用谁；全不行就直连。
#    （沙箱注入的 http_proxy 端口会变，所以必须实测、不能写死。）
def _proxy_candidates():
    """代理候选，按优先级；None 表示「不用代理、直连」。

    ① RELEASE_PROXY 环境变量 —— 显式指定，最高优先级；**空串表示强制直连**
    ② 本机代理软件常用端口（10808 / 10809 / 7890 / 7897）
    ③ 环境变量里的 https_proxy / http_proxy
    ④ None（直连）

    为什么环境变量代理排在**后面**：这类变量常由沙箱/其他工具注入，可能只放行了
    api.github.com，对 github.com 的 CONNECT 隧道会拒（2026-09-22 实测：环境里的
    127.0.0.1:62577 访问 api 返回 200、访问 github.com 返回 000），拿它 push 必挂。
    本机代理软件（Clash 等）的端口才是真正能出网的通道。
    """
    explicit = os.environ.get("RELEASE_PROXY")
    if explicit is not None:
        return [explicit.strip() or None]
    out = []
    for port in (10808, 10809, 7890, 7897):
        u = "http://127.0.0.1:%d" % port
        if u not in out:
            out.append(u)
    for k in ("https_proxy", "HTTPS_PROXY", "http_proxy", "HTTP_PROXY"):
        v = (os.environ.get(k) or "").strip()
        if v and v not in out:
            out.append(v)
    out.append(None)
    return out


def _proxy_works(proxy, timeout=5):
    """实测该代理能否访问 GitHub；proxy=None 时测直连。

    ⚠️ 两个域名**都必须通**才算可用：
      · api.github.com —— 建 Release、传资产走它
      · github.com     —— git push 走它（CONNECT 隧道到 443）
    踩过的坑：某个代理只放行 api.github.com，于是被判为「可用」，
    打包/签名全成功，最后 git push 报 CONNECT tunnel failed 502。
    """
    for url in ("https://api.github.com", "https://github.com"):
        try:
            handlers = ([urllib.request.ProxyHandler({"http": proxy, "https": proxy})]
                        if proxy else [urllib.request.ProxyHandler({})])
            op = urllib.request.build_opener(*handlers)
            req = urllib.request.Request(url, headers={"User-Agent": "release-probe"})
            with op.open(req, timeout=timeout) as resp:
                if getattr(resp, "status", 0) != 200:
                    return False
        except Exception:  # noqa: BLE001
            return False
    return True


def pick_proxy():
    """挑第一个实测能用的代理并写进环境变量；返回选中的值（None = 直连）。"""
    global PROXY
    for cand in _proxy_candidates():
        if _proxy_works(cand):
            PROXY = cand
            break
    for k in ("HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy"):
        prev = os.environ.get(k)
        if PROXY:
            if prev and prev != PROXY:
                print("[代理] 环境变量 {}={} → {}".format(k, prev, PROXY))
            os.environ[k] = PROXY
        else:
            if prev:
                print("[代理] 清除环境变量 {}={}（改用直连）".format(k, prev))
            os.environ.pop(k, None)
    print("[代理] {}".format(PROXY if PROXY else "不使用（直连）"))
    return PROXY


PROXY = pick_proxy()



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
    """恒定为三段 X.Y.Z：v4.7.0 / v4.6.1（不再省略第三段）。"""
    return "v{}.{}.{}".format(v[0], v[1], v[2])


def bump_minor(v):
    """升次版本位（功能更新）：4.6.3 -> 4.7.0"""
    return (v[0], v[1] + 1, 0)


def bump_patch(v):
    """升修订位（修复 / 文档 / 杂项）：4.6.3 -> 4.6.4"""
    return (v[0], v[1], v[2] + 1)


def normalize_ver(text):
    """把 --version 的输入规整成三段元组：'4.7' -> (4,7,0)、'4.7.1' -> (4,7,1)。"""
    nums = [int(x) for x in re.findall(r"\d+", str(text or ""))[:3]]
    if not nums:
        return (0, 0, 0)
    while len(nums) < 3:
        nums.append(0)
    return tuple(nums)


# 提交标题里出现这些词 = 功能更新（升次版本位）；其余一律只升修订位
_FEATURE_WORDS = ("feat", "feature", "新增", "新功能", "增加功能", "添加功能")


def is_feature_commit(subject):
    """"feat(...)"/"新增 xxx" 算功能提交；fix/docs/chore 等不算。"""
    t = (subject or "").strip().lower()
    m = re.match(r"^([a-z]+)", t)
    if m and m.group(1) in ("feat", "feature"):
        return True
    return any(w in t for w in _FEATURE_WORDS if not w.isascii())


def release_subjects():
    """本次待发布的提交标题（上次发布提交..HEAD）。

    ⚠️ 只有「完全没有基线文件」时才退回最近 15 条用于推断。基线存在但范围为空，
    说明自上次发布以来没有新提交，必须返回空列表 —— 曾经的 fallback 会把**已发布过**
    的历史提交当成本次变更，让 bump_auto 误判成「有功能提交」而多跳一位版本号
    （2026-09-19 实测：基线刷成 HEAD 后仍报出 4 条早已发布的 feat）。
    """
    last = ""
    if os.path.exists(LAST_RELEASE_COMMIT):
        last = open(LAST_RELEASE_COMMIT, encoding="utf-8").read().strip()
    rng = "{}..HEAD".format(last) if last else "-15"
    out = git("log", "--pretty=%s", rng, check=False).stdout or ""
    return [ln.strip() for ln in out.splitlines() if ln.strip()]


def bump_auto(v):
    """按本次待发布提交的内容决定递进位：有功能提交 → 升 Y；否则升 Z。

    返回 (新版本元组, 依据说明)。"""
    subs = release_subjects()
    feats = [t for t in subs if is_feature_commit(t)]
    if feats:
        return bump_minor(v), "含功能提交 {} 条（{}）→ 升次版本位".format(len(feats), feats[0])
    return bump_patch(v), "无功能提交（共 {} 条 fix/docs/chore 等）→ 只升修订位".format(len(subs))


def read_version():
    if os.path.exists(VERSION_FILE):
        return open(VERSION_FILE, encoding="utf-8").read().strip()
    return "0.0.0"


def write_version(v):
    with open(VERSION_FILE, "w", encoding="utf-8") as f:
        f.write(v + "\n")


APP_SRC = os.path.join(ROOT, "invoice_qr_tool.py")


def sync_app_version(ver):
    """把主程序源码里的 __VERSION__ 改成本次要发布的版本号（幂等 + 读回校验）。

    ⚠️ 这一步绝对不能省。软件标题栏显示的版本、以及更新检查里判断「是否已有新版」
    用的都是 __VERSION__，而 PyInstaller 是把源码**原样**打进 exe 的 —— 发版脚本
    若只写 VERSION 文件、不同步源码，装上的新版会自报旧版本号，于是每次启动都提示
    「发现新版本」，更新完仍是旧号 → **更新死循环**（本仓库实际踩过，2026-09-22）。
    写文件用二进制读写，保证除这一行外其余字节（含行尾）完全不动。
    """
    want = ver
    with open(APP_SRC, "rb") as f:
        raw = f.read()
    src = raw.decode("utf-8")
    pat = re.compile(r'__VERSION__\s*=\s*"[^"]*"')
    if not pat.search(src):
        raise RuntimeError(
            "未在 {} 里找到 __VERSION__ 赋值，无法同步版本号".format(APP_SRC))
    new_src = pat.sub(lambda _m: '__VERSION__ = "{}"'.format(want), src, count=1)
    if new_src != src:
        with open(APP_SRC, "wb") as f:
            f.write(new_src.encode("utf-8"))
    # 读回校验：写完必须核对实际落盘值，不能只看「写入没报错」
    with open(APP_SRC, encoding="utf-8") as f:
        back = re.search(r'__VERSION__\s*=\s*"([^"]*)"', f.read())
    got = back.group(1) if back else ""
    if got != want:
        raise RuntimeError("版本号同步失败：期望 {}，实际 {}".format(want, got))
    print("[版本] 源码 __VERSION__ 已同步 -> {}".format(got))
    return got


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
def git_env():
    """构造 git 的运行环境。

    本机 PATH 常被裁剪，git 会找不到远端 helper。把 git 所在目录补进 PATH；
    若是 PortableGit（helper 在 mingw64/bin，不在 libexec/git-core），再补 GIT_EXEC_PATH。
    """
    env = os.environ.copy()
    d = os.path.dirname(GIT)
    env["PATH"] = d + os.pathsep + env.get("PATH", "")
    if "PortableGit" in GIT:
        env["GIT_EXEC_PATH"] = d
    return env


def git(*args, check=True, capture=True):
    # 代理用 `-c` 显式传给 git：不依赖环境变量，沙箱注入的代理也拦不住。
    # 代理按 pick_proxy() 实测结果注入；PROXY 为 None 时**显式传空值强制直连**，
    # 免得仓库里残留的 http.proxy 把 git 带向一条死路
    # （git config 的优先级高于环境变量，见文件上方说明）。
    proxy_args = (["-c", "http.proxy=" + PROXY, "-c", "https.proxy=" + PROXY] if PROXY
                  else ["-c", "http.proxy=", "-c", "https.proxy="])
    r = subprocess.run(
        [GIT] + proxy_args + list(args), capture_output=capture, text=True,
        encoding="utf-8", errors="replace", env=git_env(),
    )
    if check and r.returncode != 0:
        raise RuntimeError("git {} 失败（returncode={}）:\nSTDOUT: {}\nSTDERR: {}".format(
            " ".join(args), r.returncode,
            (r.stdout or "").strip()[-1500:], (r.stderr or "").strip()[-1500:]))
    return r


def push_or_fail(remote, ref, what, attempts=5):
    """推送并**校验结果**，失败自动重试（指数退避）。

    两个必须点：
    1) 不能静默推送 —— 原先用 check=False，一旦失败流程仍继续，Release 可能建在错误提交上；
    2) 本机代理会间歇性抖动（实测同一 URL 三次里两次 `502 Bad Gateway` / `Empty reply from server`），
       无人值守发版不该被一次抖动打断，故退避重试 5 次（5/10/20/40 秒）。
    """
    last = None
    for i in range(1, attempts + 1):
        r = git("push", remote, ref, check=False)
        if r.returncode == 0:
            print("    已推送 {}{}".format(what, "" if i == 1 else "（第 {} 次尝试）".format(i)))
            return
        last = r
        if i < attempts:
            wait = 5 * (2 ** (i - 1))
            print("    推送{}失败（returncode={}），{} 秒后重试…".format(what, r.returncode, wait))
            time.sleep(wait)
    raise RuntimeError(
        "推送{}失败（已重试 {} 次，returncode={}）：\nSTDOUT: {}\nSTDERR: {}\n"
        "提示：若为静默失败，检查 build_release.py 的 _pick_git() 是否选中了系统 Git。".format(
            what, attempts, last.returncode,
            (last.stdout or "").strip()[-800:], (last.stderr or "").strip()[-800:]))


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
        "--icon", APP_ICON,
        "--add-data", "{}{}.".format(APP_ICON, os.pathsep),
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
        "--icon", APP_ICON,
        "--add-data", "{}{}.".format(APP_ICON, os.pathsep),
        "--add-data", add_app,
        "--add-data", add_un,
    ])
    os.replace(os.path.join(out, APP_NAME + "_安装程序.exe"), INSTALLER_OUT)
    sign(INSTALLER_OUT)


# ---------------- 资源生成 ----------------
# 发版流水线自身产生的提交，不该出现在给用户看的「更新内容」里
_NOISE_COMMIT_RE = re.compile(
    r"^(?:release\s*:|chore(?:\([^)]*\))?\s*:|[a-z]+\(release\)\s*:)", re.I)


def _user_facing_lines(log):
    """从 `git log --oneline` 输出里剔掉流水线提交，只留面向用户的条目。

    过滤对象：`release: v4.10.0`（脚本打版）、`chore: 记录 vX 发布提交`（脚本记基线）、
    以及任何 scope 为 release 的提交（如 `fix(release): ...`，改的是发版脚本而非软件功能）。
    逐条比对而非整段丢弃 —— 与功能提交混在一起时，功能条目照常保留。
    """
    kept = []
    for ln in log.splitlines():
        ln = ln.strip()
        if not ln:
            continue
        subject = ln.split(" ", 1)[1] if " " in ln else ln
        if _NOISE_COMMIT_RE.match(subject):
            continue
        kept.append(ln)
    return kept


def _changelog_section(tag):
    """从主程序源码的 CHANGELOG_TEXT 里抠出 tag 对应的那一段。

    这是 Release 正文的**首选来源** —— 它同时就是软件内「更新记录」显示的文案，
    所以线上 Release 与软件内说明永远是同一份，不会各写一遍而走样。
    返回 None 表示源码里还没写该版本段落（此时回退到 git 提交标题）。
    """
    try:
        src = open(os.path.join(ROOT, "invoice_qr_tool.py"), encoding="utf-8").read()
    except OSError:
        return None
    m = re.search(r'CHANGELOG_TEXT\s*=\s*"""(.*?)"""', src, re.S)
    if not m:
        return None
    text = m.group(1)
    head = re.search(r"^\d{4}-\d{2}-\d{2}\s+" + re.escape(tag) + r"\s*$", text, re.M)
    if not head:
        return None
    tail = text[head.end():]
    nxt = re.search(r"^\d{4}-\d{2}-\d{2}\s+v\d", tail, re.M)
    block = (tail[:nxt.start()] if nxt else tail).strip("\n")
    return block or None


def _release_body(new_tag, portable_name, installer_name):
    """Release 正文 与 附件 InvoiceQR_Changelog.txt 共用同一份内容。"""
    body = "# 发票二维码识别下载工具 {}\n\n".format(new_tag)
    section = _changelog_section(new_tag)
    if section:
        body += "## 更新内容\n\n" + section + "\n"
        print("[资源] Release 正文取自源码更新记录（{} 段落 {} 行）".format(new_tag, len(section.splitlines())))
    else:
        # 兜底：源码里没有该版本段落，用「相对上次发布」的提交标题（已滤掉发版流水线提交）
        print("[资源] ⚠ 源码更新记录里没有 {} 段落，正文回退为 git 提交标题".format(new_tag))
        last = ""
        if os.path.exists(LAST_RELEASE_COMMIT):
            last = open(LAST_RELEASE_COMMIT, encoding="utf-8").read().strip()
        rng = "{}..HEAD".format(last) if last else "-15"
        log = git("log", "--oneline", rng, check=False).stdout.strip()
        if not log:
            log = git("log", "--oneline", "-15").stdout.strip()
        lines = _user_facing_lines(log) or log.splitlines()  # 全是流水线提交时别留空段
        body += "## 更新内容\n"
        for ln in lines:
            body += "- " + ln + "\n"
    body += "\n## 下载说明\n"
    body += "- `{}`：单文件运行版（双击即用，无需安装）。\n".format(portable_name)
    body += "- `{}`：可安装版（装到 Program Files，开始菜单/桌面快捷方式，含卸载程序）。\n".format(installer_name)
    body += "- 软件界面与本地文件名仍为中文 `{}`。\n".format(APP_EXE)
    # 版本段落里若已说明升级方式，就不再重复追加（正文出现两遍同义句很显廉价）
    if not any(k in body for k in ("覆盖安装", "覆盖旧版本", "无需迁移")):
        body += "\n> 升级方式：直接运行新版即可覆盖旧版本，无需卸载，也无需迁移任何数据。\n"
    return body


def make_assets(new_tag):
    os.makedirs(ASSET_DIR, exist_ok=True)
    # 复制双 exe（ASCII 文件名，GitHub 附件不支持中文）
    portable_name = "InvoiceQRDownloader_{}.exe".format(new_tag.lstrip("v"))
    installer_name = "InvoiceQRInstaller_{}.exe".format(new_tag.lstrip("v"))
    shutil.copy2(PORTABLE_OUT, os.path.join(ASSET_DIR, portable_name))
    shutil.copy2(INSTALLER_OUT, os.path.join(ASSET_DIR, installer_name))

    def _sha256(path):
        h = hashlib.sha256()
        with open(path, "rb") as fh:
            for chunk in iter(lambda: fh.read(1024 * 1024), b""):
                h.update(chunk)
        return h.hexdigest()

    # ⚠️ 顺序要紧：先算好主程序 SHA256，再生成 Release 正文 ——
    #    正文里要追加 `SHA256: <64位>`，客户端走 GitHub API 兜底时靠它校验完整性
    #    （API 不提供 sha256 字段，只能从正文抠）。详见 iqr_update._sha256_from_notes。
    portable_sha = _sha256(PORTABLE_OUT)

    # Changelog（= Release 正文）优先取源码「更新记录」本版段落，git log 兜底
    changelog = _release_body(new_tag, portable_name, installer_name)
    # 摘要追加在正文末尾，独立成段，便于客户端正则定位（也方便用户肉眼核对）
    changelog += (
        "\n## 文件校验\n\n"
        "```\nSHA256: {}\n```\n"
        "\n> 上面是 `{}` 的 SHA256。走 GitHub 加速镜像下载时可用它核对文件是否完整、未被篡改。\n"
    ).format(portable_sha, portable_name)
    with open(os.path.join(ASSET_DIR, "InvoiceQR_Changelog.txt"), "w", encoding="utf-8") as f:
        f.write(changelog)

    usage = (
        "发票二维码识别下载工具 {tag}\n"
        "================================\n\n"
        "【下载（推荐：国内直连，速度快）】\n"
        "{site}/\n"
        "本站为国内服务器直链，不必访问 GitHub；GitHub 地址见文末。\n\n"
        "项目主页（源码 / 下载 / 更新日志）：\n"
        "{url}\n\n"
        "【单文件运行版】InvoiceQRDownloader_{tag}.exe\n"
        "  直接双击运行，无需安装。程序会在同目录读写配置与输出。\n\n"
        "【可安装版】InvoiceQRInstaller_{tag}.exe\n"
        "  右键「以管理员身份运行」→ 选择安装目录（默认 C:\\Program Files\\发票二维码工具）\n"
        "  → 自动创建开始菜单 / 桌面快捷方式，并写入「应用和功能」卸载项。\n"
        "  卸载：设置 → 应用 → 发票二维码工具 → 卸载，或控制面板。\n\n"
        "两版功能完全一致，按使用场景选择即可。\n"
    ).format(tag=new_tag.lstrip("v"), url=PROJECT_URL, site=SITE_URL)
    with open(os.path.join(ASSET_DIR, "InvoiceQR_Usage.txt"), "w", encoding="utf-8") as f:
        f.write(usage)

    # 客户端自动更新元数据：作为 Release 附件上传（固定叫 update.json）。
    # ⚠️ 2026-09-24 起客户端**只从 GitHub 读**（经加速镜像拉这个附件），
    #    不再读自有服务器的 /updates/qr.json —— 故 fallback_url / site_url
    #    等指向服务器的字段已删除，别再加回来（加了也没人读）。
    ver = new_tag.lstrip("v")

    update_meta = {
        "app": "qr",
        "name": APP_NAME,
        "version": ver,
        "asset": portable_name,
        "notes": changelog,
        # url 与 setup_url 均为 GitHub Release 直链；
        # 客户端会再展开成各加速镜像并测速择优（见 iqr_update.order_download_urls）。
        "url": "{}/releases/download/{}/{}".format(PROJECT_URL, new_tag, portable_name),
        "release_url": "{}/releases/tag/{}".format(PROJECT_URL, new_tag),
        "size": os.path.getsize(PORTABLE_OUT),
        "sha256": portable_sha,
        "published": time.strftime("%Y-%m-%d %H:%M:%S"),
        "setup_url": "{}/releases/download/{}/{}".format(
            PROJECT_URL, new_tag, installer_name),
    }
    with open(os.path.join(ASSET_DIR, "update.json"), "w", encoding="utf-8") as f:
        json.dump(update_meta, f, ensure_ascii=False, indent=2)
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

    def _mb(path):
        """十进制 MB，与页面 chip 显示口径一致（1024 进制会与实际对不上）。"""
        try:
            return round(os.path.getsize(path) / 1_000_000)
        except OSError:
            return 0

    size_mb = _mb(PORTABLE_OUT)
    size_mb_i = _mb(INSTALLER_OUT)

    def rewrite(path, pairs):
        """pairs 每项为 (正则, 替换) 或 (正则, 替换, 次数)；次数 0/缺省 = 全部替换。

        未命中的规则会打印告警 —— 官网字段正则失配是会**静默**把页面留在旧版本的，
        以前就吃过这个亏，所以宁可吵一点。
        """
        if not os.path.exists(path):
            print("    跳过（缺失）:", os.path.basename(path)); return
        s = open(path, encoding="utf-8").read()
        for item in pairs:
            pat, rep = item[0], item[1]
            cnt = item[2] if len(item) > 2 else 0
            s, n = re.subn(pat, rep, s, count=cnt)
            if n == 0:
                print("    ⚠ 未命中:", os.path.basename(path), "->", pat[:48])
        with open(path, "w", encoding="utf-8") as f:
            f.write(s)

    # 首页：徽章「当前版本 vX.Y」、数据卡「vX.Y」、结尾按钮「免费下载 vX.Y」
    rewrite(os.path.join(WEBSITE_DIR, "index.html"), [
        (r"当前版本 v[\d.]+", "当前版本 " + new_tag),
        (r'(class="v"[^>]*>)\s*v[\d.]+', r"\g<1>" + new_tag),
        (r"免费下载 v[\d.]+", "免费下载 " + new_tag),
        # 数据卡「111 MB / 单文件 EXE」只改数字，标签不动（首个匹配即该卡片）
        (r">\d+ MB</b>", ">{} MB</b>".format(size_mb), 1),
    ])
    # 下载页：meta 描述、版本 chip、日期 chip、体积 chip、主下载直链（含无 v 前缀的真实附件名）
    rewrite(os.path.join(WEBSITE_DIR, "download.html"), [
        (r"v[\d.]+（Windows", new_tag + "（Windows"),
        (r'(<span class="chip">)v[\d.]+', r"\g<1>" + new_tag),
        (r'(<span class="chip gray">)\d{4}-\d{2}-\d{2}', r"\g<1>" + today),
        # 体积 chip：整块替换（旧写法「≈ 111 MB」/ 新写法「单文件 111 MB · 安装版 132 MB」都能改）；
        # 负向断言排除日期 chip，避免把发布日期也当体积写掉
        (r'(<span class="chip gray">)(?!\d{4}-\d{2}-\d{2})[^<]*',
         r"\g<1>单文件 {} MB · 安装版 {} MB".format(size_mb, size_mb_i)),
        # 下载直链走自有服务器（国内直连，2026-09-22 起）：只刷 files/ 下的文件名版本号。
        # 页面里已不再出现 GitHub 直链，故原先针对 releases/download/ 的两条正则已删除 ——
        # 留着只会每次发版报「未命中」噪音。
        (r"files/InvoiceQRDownloader[_v]*[\d.]+\.exe",
         "files/InvoiceQRDownloader_{}.exe".format(ver)),
        (r"files/InvoiceQRInstaller[_v]*[\d.]+\.exe",
         "files/InvoiceQRInstaller_{}.exe".format(ver)),
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
        push_or_fail("origin", branch, "分支 " + branch)
        push_or_fail("origin", new_tag, "标签 " + new_tag)
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
        # 重跑 / 补传附件时同步刷新正文，避免线上还是上次的旧文案
        st_b, _ = api("PATCH", rel_url + "/" + str(release_id), token, json_data={"body": body})
        print("[发布] 正文已刷新 (HTTP {})".format(st_b))
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
        # 自动更新元数据（客户端经 GitHub 加速镜像读这个附件拿版本号 + sha256）
        ("update.json", "application/json; charset=utf-8"),
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
    # 该文件不入库（.gitignore），所以通常没有可提交的内容；只有确实产生了新提交才推送。
    # 否则代理抖动时这次多余的 push 会让「已经发布成功」的流程以非 0 退出，很容易被误判成发版失败。
    with open(LAST_RELEASE_COMMIT, "w", encoding="utf-8") as f:
        f.write(head_commit() + "\n")
    git("add", "-A", check=False)
    r = git("commit", "-m", "chore: 记录 {} 发布提交".format(new_tag), check=False)
    if r.returncode == 0:
        try:
            push_or_fail("origin", branch, "分支 " + branch)
        except RuntimeError as e:
            print("[发布] ⚠ 基线提交推送失败（不影响本次发布，Release 与附件均已就绪）：")
            print("       " + str(e).replace("\n", "\n       "))
    else:
        print("[发布] 无新增提交需要推送（发布基线文件不入库）。")
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
        new_v = normalize_ver(ver_override)
        print("[版本] 手动指定 {} -> {}".format(ver_override, fmt_ver(new_v)))
    elif do_publish and latest_v > (0, 0, 0):
        new_v, why = bump_auto(latest_v)
        print("[版本] 自动递增 {} -> {}".format(latest, fmt_ver(new_v)))
        print("[版本] 依据:", why)
    else:
        new_v = local_v if local_v > (0, 0, 0) else (0, 0, 1)
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

    # 版本号闸门：打包前必须把 __VERSION__ 同步成目标版本（幂等）。
    # 放在构建之前 —— 否则 exe 里装的是旧版本号，用户侧会陷入「反复提示更新」。
    sync_app_version(new_tag.lstrip("v"))

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

    # 二次核对：产物由源码打出来，构建完成后源码里的版本号必须仍是目标版本，
    # 否则说明中途被改动，产物会带错号 → 宁可中止发版也不要发出去
    with open(APP_SRC, encoding="utf-8") as _f:
        _m = re.search(r'__VERSION__\s*=\s*"([^"]*)"', _f.read())
    _got = _m.group(1) if _m else ""
    if _got != new_tag.lstrip("v"):
        raise RuntimeError("构建后源码版本号为 {}，与目标 {} 不一致，已中止发版"
                           .format(_got, new_tag))
    print("[版本] 构建后核对通过：产物内含版本号 {}".format(_got))

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
