#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
发票二维码识别下载工具
功能：
1. 识别指定文件夹内图片中的二维码（全部图片均识别，不再按网址去重）；
2. 若二维码为网址，则下载对应 PDF 到同目录 PDF 文件夹下，文件名与原始图片相同；
3. 识别结果的归类与文件处理：
   - 成功识别并下载到 PDF：原图片文件名保持不变；
   - 识别到网址但无可下载的 PDF：复制一份到「未识别」文件夹，文件名前加“未下载-”；
   - 未识别到任何二维码：复制一份到「未识别」文件夹，文件名前加“未识别-”；
   - 其它情况（识别到二维码但非网址等）：复制一份到「未识别」文件夹，文件名前加“其它-”；
4. 可选任务完成后打开文件夹；
5. 可选将下载的 PDF 转换为 JPG 图片（长边 2000px，短边自适应）；
6. 可选任务完成后汇总发票（对 PDF 文件夹内发票 PDF 提取字段并生成 Excel，
   金额/统筹为纯数字、无千分位，含「是否重复」列与右侧统计区汇总）。

并发模型：1~3 步由线程池并发执行（默认 6 路，见 DEFAULT_WORKERS）；
第 6 步的汇总保持串行（pdfplumber 是纯 Python 解析，并发无收益）。
处理中可通过 CancelToken 请求停止。

界面：整行扁平「▶ 开始处理」主按钮 + 右侧「■ 停止」+ 自绘扁平进度条
（右侧显示「进度：x / y 份 · 6 路并发」）；「使用说明 / 检查更新」位于顶部菜单栏「帮助」。
"""

import os
import re
import sys
import glob
import time
import queue
import shutil
import subprocess
import tempfile
import datetime
import threading
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed, CancelledError
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import requests

# 注意：cv2 / numpy / pymupdf / pdfplumber / openpyxl / zxingcpp 等重型依赖
# 在启动时并不需要（GUI 与更新检查仅用到 tkinter + requests）。它们改为在
# 对应功能被真正调用时才“懒加载”，可大幅缩短启动时间（窗口更快出现）。
# 懒加载由各功能函数内部 import 完成，Python 会缓存已导入模块，重复调用无额外开销。

URL_RE = re.compile(r"https?://[^\s<>\"{}|\\^`\[\]]+", re.IGNORECASE)

SUPPORTED_IMAGE_EXTS = (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp")
SUPPORTED_PDF_EXTS = (".pdf",)

# 「PDF 前置转换」子文件夹名：只要目标文件夹里出现 PDF，就把 PDF 转出的图片与原有图片
# 一起收进这个子文件夹，再在其中按原逻辑处理（输出 PDF/ 未识别/ 汇总表也都落在它里面）。
POSTPROCESS_DIR_NAME = "处理后"

# 状态前缀：把“未下载 / 未识别 / 其它”的图片复制到「未识别」子文件夹时加在文件名前
PREFIX_UNRECOGNIZED = "未识别-"  # 未识别到任何二维码
PREFIX_NOT_DOWNLOADED = "未下载-"  # 识别到网址但无可下载的 PDF
PREFIX_OTHER = "其它-"           # 其它情况（识别到二维码但内容非网址等）

# 并发设置
#   每张图片是一个独立任务，由线程池并发执行。绝大多数耗时都在“等网络下载 PDF”上，
#   这部分 CPU 全程空转，并发收益最大（实测 16 张：串行 2.01s → 6 路 0.40s，约 5×）。
#   CPU 部分（二维码识别）也能提速（60 张 4.82s → 1.62s，约 3×），因为 zxing/OpenCV
#   是 C++ 扩展、会释放 GIL。
#   注意：「汇总发票」阶段的 pdfplumber 解析是纯 Python、被 GIL 锁死，实测并发无收益（1.0×），
#   所以那里保持串行，不要改成线程池。
DEFAULT_WORKERS = 6        # 并发路数（对同一发票平台是 6 次并发请求，过低提不上速、过高易被限流）
DOWNLOAD_RETRIES = 2       # 单张 PDF 下载失败后的重试次数（不含首次），仅对网络类错误重试
RETRY_BACKOFF = 0.6        # 重试退避基数（秒）：0.6s、1.2s 递增

# 软件自身版本与 GitHub 更新源（公开仓库，更新检查无需鉴权）
__VERSION__ = "4.8.0"
GITHUB_REPO_OWNER = "jianRY"
GITHUB_REPO_NAME = "invoice-qr-tool"
GITHUB_LATEST_RELEASE_URL = (
    f"https://api.github.com/repos/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/releases/latest"
)


USAGE_TEXT = f"""发票二维码识别下载工具 · 使用说明
================================

项目地址（源码 / 下载 / 更新日志）：
    https://github.com/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}
最新版本与安装包请见仓库 Releases 页；软件启动后也会自动检查更新。

【功能】
1. 识别指定文件夹内图片中的二维码（全部图片均识别，不再按网址去重）。
2. 二维码为网址：自动下载对应 PDF 到「PDF」子文件夹，文件名与原始图片相同。
3. 识别结果归类与文件处理（后三类仅做“复制”，原图片始终保留在原始文件夹中不动）：
   - 成功识别并下载到 PDF：原图片文件名保持不变。
   - 识别到网址但无可下载的 PDF：复制一份到「未识别」文件夹，文件名前加「未下载-」。
   - 未识别到任何二维码：复制一份到「未识别」文件夹，文件名前加「未识别-」。
   - 其它情况（识别到二维码但内容非网址等）：复制一份到「未识别」文件夹，文件名前加「其它-」。
4. 可选：处理完成后自动打开文件夹。
5. 可选：将下载的 PDF 转为 JPG 图片（长边 2000px，短边自适应），保存到「PDF/图片」。
6. 可选：处理完成后汇总发票（生成 Excel）。对「PDF」文件夹内所有发票 PDF，
   提取「交款人 / 票据号码 / 开票日期 / 金额合计（小写）/ 医保统筹基金支付」，
   以及「大病保险支付」「医疗救助支付」两类支付栏目（识别到才输出），
   汇总为「PDF/发票汇总_YYYYMMDD_HHMMSS.xlsx」；金额类列均为纯数字（无千分位、
   可直接求和与二次计算），并含「是否重复」列（同一票据号码出现≥2次则标记“是”）；
   表格右侧统计区：票据张数 / 合计总金额 / 合计总统筹金额 /〔合计大病保险支付〕/
   〔合计医疗救助支付〕/ 可赔付金额 / 重复票据金额合计 / 重复票据统筹合计 /
   〔重复票据大病保险支付合计〕/〔重复票据医疗救助支付合计〕。
   说明：
   - 加〔〕的项目与对应支付列一样，**仅在识别到该类目时才出现**（识别到 0.00 也算“有”该类目）；
     本批票据若一张都没有，该列连同其汇总项整体省略，输出与旧版完全一致。
   - 可赔付金额 = 合计总金额 − 各支付类合计之和（统筹 / 大病保险 / 医疗救助）。
   - 「重复票据…合计」只统计**多出来的份**：同一票据号码有多张时，第 1 张已计入
     「合计总金额」，不再计入重复合计；只有第 2 张及以后的金额才计入，避免同号票被重复累加。
     「是否重复」列仍对同号的每一张都标“是”（判定逻辑不变），明细行也全部保留。
7. 处理结束后，弹出「识别结果统计」：总计识别图片数、成功下载 PDF、识别但未下载、
   未识别、其它各多少张，并附本次耗时、平均每张耗时与并发路数，方便核对处理结果。
8. 日志区右上角提供「清空日志」按钮，一键清空历史日志，便于开始下一个任务。
9. 多张图片**并发处理**（默认 6 路）：识别、下载、转图并行推进。绝大多数时间都花在
   等网络下载 PDF 上，并发后批量处理明显更快（下载阶段实测约 5 倍、识别约 3 倍）。
10. 单张 PDF 下载失败会**自动重试**（最多 2 次、逐次退避）：偶发网络抖动不会再被
    误判成「未下载」而需要人工重跑。
11. 处理过程中可点「■ 停止」随时中止：尚未开始的图片会被直接跳过，已下载的 PDF 全部保留，
    原图始终不动；停止后不再执行汇总与打开文件夹。
12. **PDF 前置转换（自动，无需勾选）**：目标文件夹里只要有 PDF，软件会先自动把 PDF 逐页
    转成 JPG，再往下走识别/下载流程。分三种情况：
    - 全是图片 → 不新建任何文件夹，完全按原逻辑处理（与旧版一致）；
    - 全是 PDF → 在目标文件夹内新建「处理后」，PDF 转出的图片放入其中，然后处理「处理后」；
    - PDF + 图片混合 → 同样新建「处理后」，PDF 转出的图片 **加上原有图片的副本** 一起放入，
      再处理「处理后」（原始图片文件保留不动）。
    转出的页面命名为「原文件名_1.jpg / 原文件名_2.jpg …」（按页码递增）。
    「处理后」每次处理都会**清空重建**，不会带入上一次的遗留结果；
    命名冲突会自动加后缀区分；单个 PDF 损坏 / 加密时只跳过该份，不中断整体。

【使用步骤】
1. 把待处理的图片（和/或 PDF）放在同一个文件夹里。
2. 打开本软件，点「浏览…」选择该文件夹。
3. 按需勾选：
   - 处理完成后打开文件夹
   - 将下载的 PDF 转换为图片（JPG，长边 2000px）
   - 处理完成后汇总发票（生成 Excel）
4. 点「▶ 开始处理」（主界面整行的大按钮），在日志区查看进度；按钮下方进度条右侧
   实时显示「进度：x / y 份 · 6 路并发」，日志区每 3 秒报一次「已完成 x / y 张」，
   卡住也能一眼看出还在跑。
   要中止就点「开始处理」右侧的「■ 停止」；日志区右上角有「清空日志」。
   「使用说明 / 更新记录」「检查更新」「清空日志」也可在顶部菜单栏「帮助」中找到。
5. 处理完成后，PDF 在「PDF」文件夹，转换图片在「PDF/图片」，问题图片的副本在「未识别」文件夹，
   汇总表在「PDF/发票汇总_*.xlsx」。

【自动更新】
- 软件启动后会静默检查 GitHub 上的最新版本；发现新版本时弹窗提示，点「是」即自动
  下载并安装（无需手动去网页下载）。
- 也可随时通过顶部菜单栏「帮助 → 检查更新」手动检查。
- 更新源为公开仓库 jianRY/invoice-qr-tool 的 Release，无需任何账号或令牌。
- 更新过程有「更新进度」提示框：展示阶段、进度条、下载速度、已下载大小与详细日志。
- **下载过程中可随时取消**：进度框里有「✖ 取消更新」按钮（点窗口 ✕ 同样生效）。
  取消后立即断开下载并清理临时文件，当前版本不变、软件照常可用。
  注意：下载完成后「保存新版本」的一两秒内不可取消（避免留下损坏的文件）；
  新版本已启动后也无法回退。
- 下载完成后，新版本直接保存到当前程序同一目录下，文件名自动带版本号
  （如「发票二维码工具_v3.8.exe」）；旧的 EXE 会被移入回收站（随时可还原），
  随后自动切换到新版本，安全且无损，不再需要复杂的覆盖/备份/回滚机制。

【输出规则速查】
（注：只要文件夹里含 PDF，以下输出全部落在「处理后」子文件夹内 —— 即 目标文件夹\处理后\…）
- 含 PDF（自动）         → 新建「处理后」，PDF 转「原名_1.jpg / 原名_2.jpg …」+ 复制原有图片进
- 网址 + 下载成功      → PDF/<原名>.pdf，原图片文件名保持不变
- 网址 + 无 PDF 可下载 → 复制一份到「未识别/未下载-<原名>」
- 未识别到二维码       → 复制一份到「未识别/未识别-<原名>」
- 识别到二维码但非网址 → 复制一份到「未识别/其它-<原名>」
- 勾选转图             → PDF/图片/<原名>_第N页.jpg（JPG 格式，长边 2000px）
- 勾选汇总             → PDF/发票汇总_YYYYMMDD_HHMMSS.xlsx（金额类列为纯数字、无千分位，
                         含「是否重复」列与右侧统计区汇总；重复项只计同票号第 2 张及以后）

【独立工具：只转 PDF】
命令行方式可只做前置转换（不识别、不下载、不汇总）：
    发票二维码工具.exe --pdf2img "<文件夹路径>"
效果：在该文件夹内新建（或清空重建）「处理后」，把 PDF 转成「原名_1.jpg…」并与原有图片
一起放进去，随后可再打开软件对该「处理后」文件夹走完整流程。

【说明】
- 二维码识别使用 zxing-cpp，对截图 / 小二维码会自动多尺度放大，比 OpenCV 自带更稳。
- 部分发票平台（如 jsczt.cn）打开后是一个展示页，软件会自动提取页面隐藏参数并提交下载接口获取 PDF。
- 默认 6 路并发。同一平台短时间内并发请求过多可能被限流，若出现较多「未下载」，
  等几分钟再次点「开始处理」即可（已下载的会自动跳过同名文件、未识别的原图也还在）。
- 「汇总发票」这一步保持单线程：它用 pdfplumber 解析 PDF，属纯 Python 计算，
  并发实测没有收益（1.0×）。
- 单文件 EXE，无需安装，双击即用。
"""

CHANGELOG_TEXT = """发票二维码识别下载工具 · 更新记录
================================

2026-09-17  v4.8.0
- **更新过程支持中途取消**：下载新版本时进度框新增「✖ 取消更新」按钮，点窗口右上角
  ✕ 也等同取消。取消后立刻断开下载、清理临时文件，**当前版本完全不受影响**，
  软件可继续正常使用（原先一旦开始就只能等它下载完，无法中止）。
- 取消的生效范围做了分阶段处理：**下载中**可随时取消；**下载完成后保存新文件的
  那一两秒**属临界区不可取消（此时打断可能留下损坏的 exe），按钮会变为「保存中…」
  并提示「此步骤无法取消」；**新版本已启动**后则无法回退。
- 取消后进度框停在可查看状态（日志保留「已取消更新：下载已中断，临时文件已清理」），
  点「关闭」即可回到主界面，不必重启软件。
- 顺带修复：取消时 `requests` 会尝试读完剩余响应体才返回（表现为「点了取消却卡住」），
  改为先切断原始 socket 再关闭，取消即时生效。
- 顺带修复：更新失败或取消时，半截的 `.part` 更新包有时没被清掉（曾实测残留 100MB+）；
  现在统一在最外层清理，且软件启动时也会扫一遍临时目录，清掉历史遗留的半截更新包。

2026-09-16  v4.7.0
- **版本号改为三段式 X.Y.Z**（tag 形如 v4.7.0，不再出现两段式 v4.6）：
  主界面标题、更新日志、官网与更新包文件名（`发票二维码工具_v4.7.0.exe`）统一三段式。
- **发版时的递增规则**：读「上次发布提交..HEAD」的提交标题 —— 含 feat / 新增 等功能类提交
  则升次版本位（4.6.3 → 4.7.0）；仅有 fix / docs / chore 等则只升修订位（4.6.3 → 4.6.4）。
  也可用 `--version 4.7.1` 手动指定（写 4.7 等价于 4.7.0）。
- 自动更新的版本比较本就按三段元组（缺失段补 0），旧版客户端仍可正确识别新版本号。

2026-09-16  v4.6
- **Release 附件的安装指引（InvoiceQR_Usage.txt）同样把项目地址置顶**，下载后第一眼
  即可看到源码 / 下载 / 更新日志入口。
- 明确说明文件生成职责：outputs/release_assets/ 改由 build_release.py **独占生成**
  （安装指引 + 按本次提交生成的更新日志）；extract_release_assets.py 只写根目录与
  outputs/ 的本地说明，不再写该目录 —— 原先两个脚本写同名文件会互相覆盖。

2026-09-16  v4.5
- **使用说明置顶项目地址**：说明文本开头即列出「源码 / 下载 / 更新日志」入口
  https://github.com/jianRY/invoice-qr-tool，软件「帮助 → 使用说明」窗口与随附的
  使用说明.txt（根目录 / outputs 两处）内容一致。
- 该地址由源码常量 GITHUB_REPO_OWNER / GITHUB_REPO_NAME 拼接生成，仓库改名只需改一处。
- 修复说明文件抽取脚本：extract_release_assets.py 升级为支持 f-string 常量
  （原先只识别普通字符串，USAGE_TEXT 改为 f-string 后会抽取失败）。

2026-09-16  v4.4
- **汇总表新增「医疗救助支付」列**（与「大病保险支付」同一机制）：识别方式同为票据右下角
  基金支付区的固定版式，列位置紧随大病保险列之后，同样参与「可赔付金额」的扣减。
- 「可赔付金额」= 合计总金额 − 合计总统筹金额 − 合计大病保险支付 − 合计医疗救助支付
  （按实际输出的支付列依次扣减）。
- **修正「重复票据…合计」的重复计算**：「是否重复」列的判定逻辑保持不变（同一票据号码出现
  ≥2 次即标记“是”，同号的每一张都标），但金额 / 统筹 / 大病保险 / 医疗救助的重复合计
  **改为只统计「重复份」** —— 即同一票号第 2 张及以后，首张不再计入（它已包含在「合计总金额」里）。
  举例：同一票号出现 3 张、每张 100 元，原口径会把 300 元全算进重复合计，现口径只算 200 元。
  实现方式：明细里新增一个**隐藏辅助列**标记「重复份」，重复合计用 SUMIF 引用它，明细数据不变。
- 条件列改为可扩展列表（`_SUMMARY_COND_FIELDS`）：以后再加类目，只需加个名字 + 在
  `parse_invoice` 里取值，明细列、合计行、统计区会自动跟上，不必再改公式与列号。
- 统计区标签列加宽，「重复票据医疗救助支付合计」这类长标签可完整显示。

2026-09-16  v4.3
- **汇总表新增「大病保险支付」列**：识别方式与「医保统筹基金支付」完全一致（票据右下角
  基金支付区的同一版式、同样的取值规则），列位置紧随统筹列之后，写入纯数字可求和。
  例：某张住院票据 医保统筹 111,710.09、大病保险 4,868.33。
- 右侧统计区同步新增「合计大病保险支付」「重复票据大病保险合计」两项
  （后者只对标记为“是”的重复票据求和）。
- **该列为条件输出**：仅当本批票据中至少有一张识别到「大病保险支付」时才生成该列及其
  统计项；若一张都没识别到（票据本身无此栏），该列连同其汇总项整体省略 —— 没这数值就忽略，
  输出与旧版完全一致，不会多出空列。
- **口径调整**：「可赔付金额」由「合计总金额 − 合计总统筹金额」改为在此基础上**再减去
  合计大病保险支付**。大病保险同属已由基金 / 保险支付、患者并未实际支出的部分，
  不扣除会虚高可赔付额。未识别到大病保险列时公式不变。
- 汇总表列号不再硬编码：改为按列名动态推算（含列宽、金额格式、合计行与统计区公式），
  后续再加提取字段不会再牵动其它列。
2026-09-16  v4.2
- **重新设计应用图标**：由「AI 生图」改为几何绘制，新图标为「取景框 + 二维码」——
  深蓝渐变圆角底、中央二维码、四角青色识别框；内含 16 / 24 / 32 / 48 / 64 / 128 / 256
  共 7 档尺寸，任务栏与资源管理器里的小图标不再糊。旧图标右下角残留的「AI生成」水印一并去除。
- **修复窗口图标发虚**：Windows 上 Tk 的 iconbitmap 会取 ICO 内最小的一档（16px）再放大，
  导致标题栏与任务栏图标模糊；现改用 Win32 方式显式加载 32 / 16px 两档后设置，
  实测读回窗口图标的像素与图标文件完全一致。主窗口与「更新进度」「使用说明」两个弹窗均已应用。
- **任务栏身份修正**：设置 AppUserModelID，任务栏不再把本程序归到 Python 名下、显示 Python 图标。
- 安装程序 / 卸载程序同步换用新图标（窗口图标与可执行文件图标），三支 exe 图标统一。
- 官网 logo 与浏览器标签页图标同步更新（体积由约 1 MB 降至 59 KB）。

2026-09-16  v4.1
- **新增「PDF 前置转换」步骤（自动执行，无需勾选）**：目标文件夹里只要有 PDF，软件就先
  自动把 PDF 逐页转成 JPG，再往下走原有的识别 / 下载 / 汇总流程。三种情况：
  ① 全是图片 → 不新建任何文件夹，完全按原逻辑处理（与旧版一致）；
  ② 全是 PDF → 在目标文件夹内新建「处理后」，PDF 转出的图片放入其中，随后处理「处理后」；
  ③ PDF + 图片混合 → 同样新建「处理后」，PDF 转出的图片 + **原有图片的副本** 一起放入，
     再处理「处理后」（原始图片保留不动）。
- 转出的页面命名「原文件名_1.jpg / 原文件名_2.jpg …」（按页码递增，取代此前的 `_第N页`）。
- 「处理后」文件夹每次处理都会**清空重建**，避免上一次的遗留结果干扰本次处理；
  命名冲突（如转出的 X_1.jpg 与原文件夹的 X_1.jpg 撞名）自动加后缀区分，不覆盖；
  单个 PDF 损坏 / 加密 / 无权限时只记日志跳过，不中断整体流程。
- **新增独立工具入口**：`发票二维码工具.exe --pdf2img "<文件夹>"` 可只做这一步转换
  （不识别、不下载、不汇总），转换完可直接对生成的「处理后」文件夹跑完整流程。

2026-09-16  v4.0
- **合并发布：在 GitHub v3.8 基础上整合本地开发成果发布 4.0**：保留 v3.7/v3.8 全部能力（并发 6 路处理、可随时「■ 停止」、PDF 转 JPG、汇总 Excel 去千分位与公式修正等），并新增：
- **双 exe 分发**：除原有单文件运行版外，新增「可安装版」（`InvoiceQRInstaller_4.0.exe`），以管理员权限安装到 Program Files、创建开始菜单 / 桌面快捷方式、写入「应用和功能」卸载项（含卸载程序）。
- **发布流水线自签名**：所有 exe 走 SHA256 + RFC3161 时间戳自签名（签名者 CN=jianRY），消除 SmartScreen / 未知发布者警告。
- 构建与发布一体化：本地 `build_release.py` 一键完成「构建双 exe → 逐个签名 → 生成说明 → 推送 GitHub → 打 tag → 发 Release 上传双 exe」。

2026-09-11  v3.8
- **并发处理**：每张图片作为一个独立任务，由线程池并发执行（默认 6 路），
  识别 / 下载 / 转图不再一张张排队。实测：下载阶段 16 张由 2.01 秒降到 0.40 秒（约 5 倍），
  二维码识别 60 张由 4.82 秒降到 1.62 秒（约 3 倍）。
  说明：「汇总发票」这一步仍为单线程——pdfplumber 是纯 Python 解析，并发实测无收益。
- **新增「■ 停止」按钮**：处理过程中可随时中止。尚未开始的图片会被直接取消，
  已发出的网络请求自然收尾；已下载的 PDF 全部保留，原图不动，停止后不再执行汇总与打开文件夹。
- **进度语义修正**：旧版是在「开始处理第 N 张之前」就上报进度，导致进度条先跳一格再干活、
  且卡在下载时长时间不动，看着像死机。现在改为「已完成张数」计数，并显示「· 6 路并发」。
- **日志可读性**：每张图片的日志整块连续输出（标题行 + 明细行），并发交错也不会再出现
  「-> 已下载」看不出属于哪张图的问题；新增每 3 秒一次的「已完成 x / y 张」心跳。
- **下载自动重试**：网络类失败（超时 / 连接被重置）最多重试 2 次、逐次退避，
  偶发抖动不再被误判成「未下载」；「网址确实没有 PDF」属业务性失败，不重试。
- 修复并发隐患：`a.jpg` 与 `a.png` 去扩展名后基名相同，会指向同一个 `a.pdf`——
  串行时是后覆盖前，并发时会两个线程同时写坏文件。现在自动给冲突的名字加后缀区分。
- PDF 改为**先写临时文件再原子替换**：下载中途失败不再在「PDF」目录留下半截文件。
- 性能与资源：请求改为按线程复用 Session 连接池（省掉每张发票重复的 TLS 握手），
  任务结束后统一关闭；判断响应是否为 PDF 时不再把整个响应体读进内存。
- 其他：处理中不再弹出更新提示（避免打断任务）；启动静默检查与手动检查不再可能同时弹两个对话框；
  处理中关闭窗口会先确认，避免误关导致任务半途而废。

2026-09-11  v3.7.2
- 主按钮图标由「⬇」改为「▶」（实心右三角）：更贴合「开始处理」的执行语义，
  实心色块的视觉重量也与按钮粗体文字更协调，整体不再头重脚轻。

2026-09-11  v3.7.1
- 界面布局重新设计（按用户反馈修正 v3.7 的按钮排版）：
  ① 「开始处理」改为**整行扁平主按钮**——浅色底、细边框、居中「⬇ 开始处理」粗体文字，
     占据整个窗口宽度，点击区域更大、视觉更清爽；
  ② 进度条改为**自绘扁平样式**（浅灰轨道 + 细边框 + 蓝色填充），替掉系统默认的绿色渐变，
     与整体浅色扁平风格统一；进度条右侧新增「进度：x / y 份」实时计数；
  ③ 「使用说明 / 检查更新」从主界面移除，移入**顶部菜单栏「帮助」**，
     主界面更简洁，把注意力集中在「开始处理」上（「清空日志」在菜单栏与日志区均可使用）。

2026-09-11  v3.7
- 主界面按钮重排：「开始处理」改为醒目的大号主按钮（蓝色底、加粗字体、更大点击区域），
  位于左侧；「使用说明 / 检查更新」改为右侧统一尺寸的次要小按钮；进度条独占一行、
  横贯窗口，操作更直观、更容易点。
- PDF 转图片的格式由 PNG 改为 JPG（质量 92）：文件体积明显更小，便于上传、分享与打印；
  输出位置不变（PDF/图片/<原名>_第N页.jpg）。
- 日志区新增「清空日志」按钮（位于“处理日志”右侧），一键清空历史日志，
  便于开始下一个任务前保持界面清爽。
- 汇总 Excel 的金额去掉千分位：金额合计（小写）、医保统筹基金支付及统计区数值
  由 “1,234.56” 改为 “1234.56”，避免复制到其它表格或做二次计算时出错。
- 修复汇总 Excel 两处公式错误：①合计行的「金额合计 / 统筹合计」原先写到了相邻一列
  （分别落到统筹列与「是否重复」列），现更正为对应「金额合计（小写）」「医保统筹基金支付」列；
  ②右侧「可赔付金额」原先引用的是文字标签列（会显示 #VALUE! 错误），
  现改为引用数值列（=合计总金额 − 合计总统筹金额），计算恢复正常。

2026-08-19  v3.6
- 修复金额识别 bug：金额中的千位分隔符（逗号，如 “1,234.56”）此前会被当作结束符，
  只识别到 “1” 导致金额残缺；现已支持中英文逗号，完整识别到小数点后两位。
  另新增「金额合计（大写）」兜底：小写金额解析失败时用大写金额补回，避免漏识别。
- 重构图片识别与文件处理：去除「重复网址」判定，全部图片均识别；
  成功并下载 PDF 的图片文件名保持不变；未下载 / 未识别 / 其它三类仅“复制”一份到
  新建的「未识别」子文件夹，并分别加前缀「未下载- / 未识别- / 其它-」，原图片保留不动。
- 汇总 Excel 增强：新增「是否重复」列（同一票据号码出现≥2次标记“是”）；
  右侧统计区新增「重复票据金额合计 / 重复票据统筹合计」（仅对标记为“是”的票据求和）。

2026-08-17  v3.5.1
- 修复汇总 Excel 的 bug：金额合计（小写）、医保统筹基金支付两列之前以文本格式写入，
  导致表格下方的 SUM 合计无法计算（求和为 0）。现已统一转为数字格式（保留两位小数），
  合计行与统计区均可正常求和。
- 新增右侧统计区：在表格右侧列出四项汇总指标——
  票据张数（=COUNTA 统计发票总数）、合计总金额（=SUM 金额列）、合计总统筹金额（=SUM 统筹列）、
  可赔付金额（=合计总金额 − 合计总统筹金额），便于一眼核对整体数据。

2026-08-14  v3.5
- 更新下载逻辑重构：直接下载 GitHub 最新 Release 的 EXE，保存到当前程序同一目录，
  文件名自动带版本号（如「发票二维码工具_v3.5.exe」）；旧版本 EXE 移到回收站（可还原），
  随后自动切换到新版本。不再依赖 bat 覆盖 / 备份 / 回滚那套复杂机制，更稳更简单。
- 更换软件图标：内置养眼的发票 + 二维码主题图标（已嵌入 EXE 与窗口标题）。

2026-08-14  v3.4
- 新增「识别结果统计」：处理结束后弹出统计框，并写入日志，分项统计
  总计识别图片、成功下载 PDF、识别但未下载 PDF、纯数字忽略、识别失败（未识别）各多少张，
  重复网址额外单列。
- 成功识别并下载到 PDF 的图片，文件名前自动加「1」前缀（如「1发票001.jpg」），
  与「未下载 / 未识别- / 重复」等状态前缀一致，便于在文件夹中一眼区分已完成项。

2026-08-12  v3.3
- 修复「更新后未能真正替换原文件」的致命 bug：更新脚本 bat 之前以 UTF-8（无 BOM）写入，
  含中文的 exe 路径被 cmd 按 GBK 误读，导致 del/copy/start 全部失败、文件从未被替换；
  同时成功弹窗阻塞导致旧进程未退出就删除文件引发竞态。现已改用 utf-8-sig（带 BOM）写入
  bat，并以「轮询删除直到文件可删」替代固定等待。
- 新增「更新进度」提示框：实时展示更新阶段、进度条、下载速度、已下载大小与详细日志，
  下载完成后自动替换并重启，不再有阻塞式弹窗。

2026-08-12  v3.2
- 启动速度优化：将 cv2 / numpy / pymupdf / pdfplumber / openpyxl / zxingcpp 等重型依赖
  改为「懒加载」（仅在实际使用相关功能时才导入），GUI 与更新检查不再在启动时初始化它们，
  窗口出现更快。
- 更新回滚增加用户可见提示：若自动更新因新版本启动校验失败被回滚，重启后弹出
  「更新已回滚」警告框，说明已还原到可用版本，不再静默。
- 更新校验时机后移：在 GUI 成功构建之后才确认更新成功，进一步避免把界面初始化
  失败的版本误判为成功。

2026-08-12  v3.1
- 自动更新新增「失败自动回滚」：替换前备份当前程序为 _backup.exe；
  新版本以 --pending 自启并完成启动校验，校验失败则由更新脚本自动还原旧版并重启，
  避免更新把软件弄成打不开的状态。

2026-08-12  v3.0
- 新增「依托 GitHub 的自动更新」功能：软件启动后静默检查最新版本，
  发现新版本可一键下载并自动替换重启（也可点「检查更新」手动触发）。
- 更新源为公开仓库 jianRY/invoice-qr-tool 的 Release，全程无需账号或令牌。
- 主界面标题显示当前版本号，并新增「检查更新」按钮。

2026-08-11  v2.0
- 集成「汇总发票」功能（来自“汇总票据”会话的 invoice_summary.py）：
  处理完成后可将「PDF」文件夹内发票 PDF 提取关键字段并汇总成 Excel。
- 主页面新增开关：「处理完成后汇总发票（生成 Excel）」。
- 提取字段：文件名 / 交款人 / 票据号码 / 开票日期 / 金额合计（小写）/ 医保统筹基金支付；
  输出「PDF/发票汇总_YYYYMMDD_HHMMSS.xlsx」，含金额与医保统筹合计行。
- 新增依赖 pdfplumber（PDF 文字抽取）+ openpyxl（Excel 写入），已纳入单文件 EXE。

2026-08-11  v1.1
- 新增：二维码为网址但无可下载 PDF 时，图片重命名为「未下载<原文件名>」。
- 新增：检测到重复网址时，对应图片文件名前加「重复」标识（可与「未下载」叠加）。
- 新增：内置「使用说明」窗口与「更新记录」窗口，
  并随软件附带 使用说明.txt、更新日志.txt。

2026-08-11  v1.0
- 初版发布：二维码识别、网址下载 PDF、纯数字忽略、未识别重命名、
  可选打开文件夹、可选 PDF 转图片。
"""



def _parse_version(tag: str) -> tuple:
    """把 'v3.0' / '3.0.1' / 'V3' 这类版本号解析成可比较的元组。

    只取前 3 段数字；缺失段补 0。非数字字符全部跳过。"""
    digits = re.findall(r"\d+", tag or "")
    nums = [int(x) for x in digits[:3]]
    while len(nums) < 3:
        nums.append(0)
    return tuple(nums)


def get_latest_release():
    """查询 GitHub 最新 Release，返回 (version, download_url, notes) 或 None。

    仅读取公开仓库的 Release 列表，无需任何鉴权 token。
    发布附件必须命名为 InvoiceQRDownloader_<版本>.exe（ASCII，避免中文名被剥离）。
    """
    headers = {"User-Agent": "InvoiceQRDownloader", "Accept": "application/vnd.github+json"}
    try:
        resp = requests.get(GITHUB_LATEST_RELEASE_URL, headers=headers, timeout=15)
    except Exception:
        return None
    if resp.status_code != 200:
        return None

    data = resp.json()
    tag = data.get("tag_name", "")
    version = _parse_version(tag)
    notes = data.get("body", "") or ""
    download_url = None

    for asset in data.get("assets", []):
        name = asset.get("name", "")
        # 只匹配 EXE 主程序附件；按版本号匹配，避免误抓旧版
        if name.lower().endswith(".exe") and re.sub(r"\W", "", name).lower().startswith(
            "invoiceqrdownloader"
        ):
            if _parse_version(name) == version:
                download_url = asset.get("browser_download_url")
                break
    if not download_url:
        # 兜底：取第一个 exe 附件
        for asset in data.get("assets", []):
            if asset.get("name", "").lower().endswith(".exe"):
                download_url = asset.get("browser_download_url")
                break

    if not download_url:
        return None
    return version, download_url, notes


class _UpdateCancelled(Exception):
    """用户主动取消更新。

    单独用一个异常类型，是为了让它能穿过 `_download_file` 的兜底 except
    （那里会把网络类异常统一当作「下载失败」返回 False，取消不能被混为一谈）。"""


def _abort_response(resp):
    """彻底切断一个流式响应。

    ⚠️ 必须**先关原始 socket** 再 close()：requests 在响应体未读尽时调用
    close()，会尝试先从 socket 读取剩余数据来“补全”，那一步会阻塞住，
    表现为「点了取消却迟迟不返回」。先 raw.close() 就绕开了这条补全路径。
    """
    if resp is None:
        return
    try:
        if getattr(resp, "raw", None) is not None:
            resp.raw.close()
    except Exception:
        pass
    try:
        resp.close()
    except Exception:
        pass


def _download_file(url: str, dest: str, progress_cb=None, cancel=None) -> bool:
    """带进度回调的下载；返回是否成功。

    cancel: threading.Event；置位则立刻断开连接，并抛出 _UpdateCancelled。
            半截文件的清理由调用方在最外层 finally 统一负责（见 perform_update），
            这样无论内部走哪条异常路径，都不会把 .part 留在磁盘上。
    """
    headers = {"User-Agent": "InvoiceQRDownloader"}
    resp = None
    abort = False
    try:
        resp = requests.get(url, headers=headers, stream=True, timeout=60)
        resp.raise_for_status()
        total = int(resp.headers.get("Content-Length", 0)) or 0
        written = 0
        f = open(dest, "wb")
        try:
            for chunk in resp.iter_content(chunk_size=256 * 1024):
                if cancel is not None and cancel.is_set():
                    abort = True
                    raise _UpdateCancelled()
                if not chunk:
                    continue
                f.write(chunk)
                written += len(chunk)
                if progress_cb and total:
                    progress_cb(written, total)
        finally:
            try:
                f.close()
            except Exception:
                pass
        return True
    except _UpdateCancelled:
        abort = True
        raise
    except Exception:
        # 网络类异常统一视为「下载失败」，交由调用方清理 .part
        abort = True
        return False
    finally:
        if abort:
            _abort_response(resp)   # 立刻断开，不等连接池回收


def _derive_base_name(exe_path: str) -> str:
    """从当前 exe 文件名推导「基础名」，去掉 .exe 以及末尾已有的 _vX.Y 版本段。

    例：发票二维码工具.exe -> 发票二维码工具；发票二维码工具_v3.4.exe -> 发票二维码工具。
    """
    name = os.path.basename(exe_path)
    if name.lower().endswith(".exe"):
        name = name[:-4]
    # 去掉末尾可能的 _v 版本段（如 _v3.4 / _v3.4.1）
    name = re.sub(r"_v\d+(?:\.\d+)*$", "", name, flags=re.IGNORECASE)
    return name or "发票二维码工具"


def _version_str(version: tuple) -> str:
    """把版本元组格式化成三段式字符串。 (4,7,0)->'4.7.0'；(4,6,1)->'4.6.1'。"""
    parts = [str(x) for x in version][:3]
    while len(parts) < 3:
        parts.append("0")
    return ".".join(parts)


def _resource_path(relative: str) -> str:
    """获取打包后 / 开发时资源文件的绝对路径（用于图标等）。"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)


def setup_app_id() -> None:
    """设置 Windows 任务栏身份（AppUserModelID），必须在 tk.Tk() 之前调用。

    不设置的话，任务栏会把本程序归到 Python 名下、显示 Python 的图标。
    """
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "jianRY.InvoiceQrTool")
    except Exception:
        pass


def apply_window_icon(win) -> None:
    """给窗口设置应用图标（主窗口与每个 Toplevel 都要调用）。

    ⚠️ Windows 上 Tk 的 iconbitmap 会取 ICO 里最小的一档（16px）再放大，
    导致标题栏与任务栏图标发虚。这里额外用 Win32 LoadImage 按 32/16px 显式
    加载后经 WM_SETICON 设置，两个尺寸都清晰。
    """
    ico = _resource_path("app_icon.ico")
    if not os.path.isfile(ico):
        return
    try:
        win.iconbitmap(ico)          # 非 Windows 平台靠它；Windows 上仅作保底
    except Exception:
        pass
    if sys.platform != "win32":
        return
    try:
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        LR_LOADFROMFILE, IMAGE_ICON, WM_SETICON = 0x0010, 1, 0x0080
        user32.LoadImageW.argtypes = [wintypes.HINSTANCE, wintypes.LPCWSTR, wintypes.UINT,
                                      ctypes.c_int, ctypes.c_int, wintypes.UINT]
        user32.LoadImageW.restype = ctypes.c_void_p
        user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT,
                                        ctypes.c_void_p, ctypes.c_void_p]
        user32.SendMessageW.restype = ctypes.c_void_p
        win.update_idletasks()
        hwnd = user32.GetAncestor(win.winfo_id(), 2)   # GA_ROOT = 2
        if not hwnd:
            return
        big = user32.LoadImageW(None, ico, IMAGE_ICON, 32, 32, LR_LOADFROMFILE)
        small = user32.LoadImageW(None, ico, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
        if big:
            user32.SendMessageW(hwnd, WM_SETICON, ctypes.c_void_p(1), ctypes.c_void_p(big))
        if small:
            user32.SendMessageW(hwnd, WM_SETICON, ctypes.c_void_p(0), ctypes.c_void_p(small))
    except Exception:
        pass


def send_to_recycle_bin(path: str) -> bool:
    """把文件移动到回收站（而非彻底删除），便于误删后恢复。返回是否成功。

    使用 Windows Shell 的 SHFileOperation 并带 FOF_ALLOWUNDO 标志，即“删除到回收站”。
    非 Windows 平台或操作失败均返回 False（调用方据此决定是否保留旧文件）。
    """
    if not os.path.exists(path):
        return True
    if sys.platform != "win32":
        try:
            os.remove(path)
            return True
        except Exception:
            return False
    try:
        import ctypes
        from ctypes import wintypes

        FO_DELETE = 0x0003
        FOF_ALLOWUNDO = 0x0040
        FOF_NOCONFIRMATION = 0x0010
        FOF_SILENT = 0x0004
        FOF_NOERRORUI = 0x0400

        class SHFILEOPSTRUCTW(ctypes.Structure):
            _fields_ = [
                ("hwnd", wintypes.HWND),
                ("wFunc", wintypes.UINT),
                ("pFrom", wintypes.LPCWSTR),
                ("pTo", wintypes.LPCWSTR),
                ("fFlags", wintypes.UINT),
                ("fAnyOperationsAborted", wintypes.BOOL),
                ("hNameMappings", wintypes.LPVOID),
                ("lpszProgressTitle", wintypes.LPCWSTR),
            ]

        from_buf = ctypes.create_unicode_buffer(path + "\0")
        op = SHFILEOPSTRUCTW()
        op.hwnd = 0
        op.wFunc = FO_DELETE
        op.pFrom = ctypes.cast(from_buf, wintypes.LPCWSTR)
        op.pTo = None
        op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT | FOF_NOERRORUI
        op.fAnyOperationsAborted = 0
        op.hNameMappings = None
        op.lpszProgressTitle = None
        res = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
        # 某些环境下 SHFileOperation 会返回非 0（如 fAnyOperationsAborted），
        # 但文件实际已被移入回收站 / 删除。以“原路径是否已不存在”作为最终成功判据。
        return res == 0 or not os.path.exists(path)
    except Exception:
        return False


def perform_update(download_url: str, latest_version: tuple, on_event=None,
                   cancel=None) -> bool:
    """下载 GitHub 最新版本的 EXE，保存到当前程序同一目录（文件名带版本号），
    并把旧的 EXE 移动到回收站。

    流程（cancel 只在能干净收尾的阶段生效）：
      ① 下载 → 可取消：立刻断开、删除半截 .part，等于什么都没发生；
      ② 下载完成后的落盘 / 原子移动（约 1~2 秒内）→ **不可取消**，
         此时打断可能留下损坏的 exe，故只提示「正在保存，请稍候」；
      ③ 新版本已启动 → 既成事实，无法回退。

    参数：
      cancel: threading.Event；置位表示用户点了「取消更新」。
    返回：是否成功触发替换（取消 / 失败均返回 False）。
    """
    def emit(ev):
        if on_event:
            on_event(ev)

    def _cancelled() -> bool:
        return cancel is not None and cancel.is_set()

    emit({"type": "stage", "text": "正在下载新版本…"})
    tmp_dir = tempfile.gettempdir()
    part = os.path.join(tmp_dir, "InvoiceQRDownloader_update.part")
    try:
        if os.path.exists(part):
            os.remove(part)
    except Exception:
        pass

    # 计算下载速度（基于相邻两次进度回调的时间差）
    _last_t = [time.time()]
    _last_w = [0]

    def _prog(written, total):
        now = time.time()
        dt = now - _last_t[0]
        if dt <= 0:
            dt = 0.001
        speed = (written - _last_w[0]) / dt
        _last_t[0] = now
        _last_w[0] = written
        emit({"type": "progress", "written": written, "total": total, "speed": speed})

    def _cleanup_part():
        """删除半截 .part（取消 / 下载失败共用；删除失败不影响主流程）。"""
        try:
            if os.path.exists(part):
                os.remove(part)
                return True
        except Exception:
            pass
        return False

    def _emit_cancelled(extra):
        _cleanup_part()
        emit({"type": "detail", "text": extra})
        emit({"type": "detail", "text": "当前版本保持不变，软件可继续正常使用。"})
        emit({"type": "cancelled"})
        emit({"type": "done", "ok": False})

    # ---- ① 下载（可取消）----
    try:
        if not _download_file(download_url, part, progress_cb=_prog, cancel=cancel):
            _cleanup_part()
            emit({"type": "detail", "text": "下载失败：无法获取更新文件，请稍后重试或手动更新。"})
            emit({"type": "done", "ok": False})
            return False
    except _UpdateCancelled:
        _emit_cancelled("已取消更新：下载已中断，临时文件已清理。")
        return False

    if _cancelled():
        # 极端情形：刚好在下载结束时点取消，同样干净收尾
        _emit_cancelled("已取消更新：临时文件已清理。")
        return False

    # ---- ② 落盘（临界区，不可取消）----
    size_mb = os.path.getsize(part) / 1048576
    emit({"type": "detail", "text": f"下载完成（{size_mb:.1f} MB），正在保存到原目录…"})
    emit({"type": "stage", "text": "下载完成，正在保存新版本（此步请稍候，无法取消）…"})
    emit({"type": "lock"})   # 通知 UI 禁用「取消更新」

    try:
        current_exe = sys.executable  # 当前 EXE 自身路径（打包后）
        target_dir = os.path.dirname(current_exe)
        base = _derive_base_name(current_exe)
        ver = _version_str(latest_version)
        new_exe = os.path.join(target_dir, f"{base}_v{ver}.exe")

        # 若同名新文件已存在（理论上应为更高版本），先移除再落入
        if os.path.exists(new_exe):
            try:
                os.remove(new_exe)
            except Exception:
                pass

        # 同盘原子移动（比跨盘 copy 快且不会残留半截文件）；跨盘则回退到 move
        try:
            os.replace(part, new_exe)
        except Exception:
            import shutil
            shutil.move(part, new_exe)

        emit({"type": "detail",
              "text": f"新版本已保存为：{os.path.basename(new_exe)}"})
        emit({"type": "detail",
              "text": "旧版本将被移入回收站；软件即将切换到新版本…"})
        emit({"type": "done", "ok": True})

        # 启动新版本并把“当前（旧）exe 路径”交给它回收；随后退出旧进程，
        # 交由新进程在旧进程释放文件后把旧 exe 移入回收站。
        try:
            subprocess.Popen(
                [new_exe, "--recycle-old", current_exe],
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            return True
        except Exception as e:
            emit({"type": "detail",
                  "text": f"启动新版本失败：{e}，旧版本仍保留，可手动打开新文件。"})
            emit({"type": "detail",
                  "text": f"新版本文件已保存在：{new_exe}"})
            emit({"type": "done", "ok": False})
            return False
    except Exception as e:
        emit({"type": "detail", "text": f"保存新版本失败：{e}"})
        emit({"type": "done", "ok": False})
        return False


class UpdateProgressDialog:
    """更新进度框：展示阶段、进度条、下载速度、已下载大小与详细日志。

    支持中途取消：下载阶段点「取消更新」（或点窗口 ✕）即中断下载并清理临时文件，
    当前版本不受影响；进入「保存新版本」的临界区后不可取消（约 1~2 秒）。
    """

    def __init__(self, parent):
        self.parent = parent
        self.queue = queue.Queue()
        self._closed = False
        self.cancel = threading.Event()   # 置位 = 用户请求取消
        self._locked = False              # True = 已进入不可取消的落盘阶段
        self._cancelling = False          # 已点过取消，正在等下载线程收尾
        self.win = tk.Toplevel(parent)
        self.win.title("更新进度")
        self.win.geometry("480x380")
        self.win.resizable(False, False)
        try:
            self.win.transient(parent)
            self.win.grab_set()
            # ✕ 等同于「取消更新」：下载阶段可中断；临界区/已结束时按实际状态处理
            self.win.protocol("WM_DELETE_WINDOW", self._on_window_close)
        except Exception:
            pass
        apply_window_icon(self.win)
        self._build_widgets()
        self._poll()

    def _build_widgets(self):
        pad = {"padx": 12, "pady": 6}
        self.stage_var = tk.StringVar(value="准备中…")
        ttk.Label(
            self.win, textvariable=self.stage_var,
            font=("Microsoft YaHei", 11, "bold"),
        ).pack(anchor=tk.W, **pad)

        self.bar = ttk.Progressbar(self.win, mode="determinate", maximum=100)
        self.bar.pack(fill=tk.X, padx=12, pady=(0, 6))

        row = ttk.Frame(self.win)
        row.pack(fill=tk.X, padx=12, pady=(0, 6))
        self.pct_var = tk.StringVar(value="0%")
        self.size_var = tk.StringVar(value="0.0 / 0.0 MB")
        self.speed_var = tk.StringVar(value="— KB/s")
        ttk.Label(row, textvariable=self.pct_var, width=10).pack(side=tk.LEFT)
        ttk.Label(row, textvariable=self.size_var, width=22).pack(side=tk.LEFT)
        ttk.Label(row, textvariable=self.speed_var, width=16).pack(side=tk.LEFT)

        ttk.Label(self.win, text="详细进度：").pack(anchor=tk.W, padx=12)
        self.txt = scrolledtext.ScrolledText(
            self.win, wrap=tk.WORD, state=tk.DISABLED, height=11
        )
        self.txt.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 8))

        btn_row = ttk.Frame(self.win)
        btn_row.pack(pady=(0, 10))
        self.btn_cancel = ttk.Button(
            btn_row, text="✖ 取消更新", command=self.request_cancel
        )
        self.btn_cancel.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_close = ttk.Button(
            btn_row, text="关闭", command=self._close, state=tk.DISABLED
        )
        self.btn_close.pack(side=tk.LEFT)

    def emit(self, **kw):
        self.queue.put(kw)

    # ---- 取消 ----
    def request_cancel(self):
        """用户请求取消更新。

        - 下载阶段：对请求置位并禁用按钮，等下载线程清理完临时文件后切成「关闭」。
        - 落盘阶段 / 已结束：不做中断（前者会留半截 exe，后者无事可取消）。
        """
        if self._closed or self._cancelling or self._locked or self.cancel.is_set():
            return
        if str(self.btn_cancel["state"]) == "disabled":
            return
        self._cancelling = True
        self.cancel.set()
        # 按钮置灰并提示，避免重复点击；真正的收尾由工作线程回报后完成
        try:
            self.btn_cancel.configure(state=tk.DISABLED, text="正在取消…")
        except Exception:
            pass
        self.stage_var.set("正在取消更新…")
        self._append("收到取消请求，正在断开下载并清理临时文件…")

    def _on_window_close(self):
        """点 ✕：能取消就取消（不关窗，等收尾），否则按状态处理。"""
        if self.btn_close["state"] != "disabled":
            self._close()
            return
        if self._locked:
            # 落盘临界区：明确告知不打断，避免留下损坏的 exe
            self._append("正在保存新版本，此步骤无法取消，请稍候…")
            return
        if not self._cancelling:
            self.request_cancel()

    def _finish_cancelled(self):
        """工作线程确认取消已生效：定稿 UI（保留日志供查看，按钮切成「关闭」）。"""
        self._cancelling = False
        self._locked = True          # 关闭一切取消入口
        self.stage_var.set("已取消更新")
        try:
            self.btn_cancel.pack_forget()   # 取消入口消失，只留「关闭」
        except Exception:
            pass
        try:
            self.btn_close.configure(state=tk.NORMAL)
        except Exception:
            pass

    def _close(self):
        # 下载还没断开时，不允许直接关窗（先取消、等收尾）
        if self._cancelling and not self._locked:
            self._append("正在取消，请稍候…")
            return
        if self._closed:
            return
        self._closed = True
        try:
            self.win.destroy()
        except Exception:
            pass

    def _poll(self):
        try:
            while True:
                ev = self.queue.get_nowait()
                self._apply(ev)
        except queue.Empty:
            pass
        if not self._closed:
            self.win.after(80, self._poll)

    def _apply(self, ev):
        t = ev.get("type")
        if t == "stage":
            self.stage_var.set(ev.get("text", ""))
            self._append(ev.get("text", ""))
        elif t == "progress":
            w, tot = ev.get("written", 0), ev.get("total", 0)
            pct = (w / tot * 100) if tot else 0
            self.bar["value"] = pct
            self.pct_var.set(f"{pct:.1f}%")
            self.size_var.set(f"{w / 1048576:.1f} / {tot / 1048576:.1f} MB")
            sp = ev.get("speed", 0) or 0
            if sp >= 1048576:
                self.speed_var.set(f"{sp / 1048576:.2f} MB/s")
            else:
                self.speed_var.set(f"{sp / 1024:.1f} KB/s")
        elif t == "detail":
            self._append(ev.get("text", ""))
        elif t == "lock":
            # 进入不可取消的落盘阶段
            self._locked = True
            try:
                self.btn_cancel.configure(state=tk.DISABLED, text="保存中…")
            except Exception:
                pass
        elif t == "cancelled":
            self._finish_cancelled()
        elif t == "done":
            ok = ev.get("ok", False)
            cancelled = self.cancel.is_set()
            self._locked = True          # 收尾阶段不再允许取消
            if ok:
                self.stage_var.set("更新完成，正在切换到新版本…")
                self._append("更新完成，即将切换到新版本。")
            elif cancelled:
                self._finish_cancelled()
            else:
                self.stage_var.set("更新失败")
                self._append("更新失败，请重试或手动更新。")
            try:
                self.btn_cancel.configure(state=tk.DISABLED)
            except Exception:
                pass
            self.btn_close.configure(state=tk.NORMAL)

    def _append(self, text):
        self.txt.configure(state=tk.NORMAL)
        self.txt.insert(tk.END, text + "\n")
        self.txt.see(tk.END)
        self.txt.configure(state=tk.DISABLED)


def _start_update_flow(root: tk.Tk, download_url: str, version: tuple):
    """在进度框中执行更新；成功后短暂展示“更新完成”再关闭主程序，由新版本接管。

    取消 / 失败时主程序**照常继续运行**（不 destroy），进度框停在可关闭状态。
    """
    dlg = UpdateProgressDialog(root)

    def _worker():
        def on_event(ev):
            dlg.emit(**ev)
        ok = perform_update(download_url, version, on_event=on_event,
                            cancel=dlg.cancel)
        if ok:
            # 让进度框显示“更新完成”约 1 秒，再关闭主程序交给新版本接管
            root.after(1000, lambda: (dlg._close(), root.destroy()))
        # 取消 / 失败：什么都不做，主程序与进度框保持可用

    threading.Thread(target=_worker, daemon=True).start()


# 更新检查的互斥守卫：避免「启动静默检查」与手动点「检查更新」同时弹出两个对话框
_update_dialog_open = threading.Event()
# 处理任务是否正在进行（进行中就不再弹更新提示，免得打断用户）
_PROCESSING = threading.Event()


def _begin_update_check() -> bool:
    """尝试占用更新检查权。返回 False 表示已有一次检查在进行中，调用方应直接返回。"""
    if _update_dialog_open.is_set():
        return False
    _update_dialog_open.set()
    return True


def _end_update_check():
    _update_dialog_open.clear()


def check_and_prompt_update(root: tk.Tk):
    """后台检查更新，若有新版本则弹窗询问是否更新。供 UI 按钮调用。"""
    if not _begin_update_check():
        return
    scheduled = False
    try:
        result = get_latest_release()
        if not result:
            root.after(0, lambda: messagebox.showinfo(
                "检查更新", "暂时无法连接到更新服务器（或当前已是最新）。"))
            return
        version, download_url, notes = result
        if version <= _parse_version(__VERSION__):
            root.after(0, lambda: messagebox.showinfo(
                "检查更新", f"当前已是最新版本 v{__VERSION__}。"))
            return

        note_text = notes.strip() or "（无更新说明）"

        def _ask():
            try:
                ask = messagebox.askyesno(
                    "发现新版本",
                    f"发现新版本 v{'.'.join(map(str, version))}，当前为 v{__VERSION__}。\n\n"
                    f"更新内容：\n{note_text[:600]}\n\n是否立即下载并更新？",
                )
                if ask:
                    _start_update_flow(root, download_url, version)
            finally:
                _end_update_check()

        root.after(0, _ask)
        scheduled = True
    finally:
        # 只有确实排定了弹窗时才由 _ask 负责释放，否则这里立刻释放
        if not scheduled:
            _end_update_check()


def _silent_startup_check(root: tk.Tk):
    """启动后静默检查更新；发现新版本且用户确认则更新。"""
    if _PROCESSING.is_set():
        return
    if not _begin_update_check():
        return
    scheduled = False
    try:
        result = get_latest_release()
        if not result:
            return
        version, download_url, notes = result
        if version <= _parse_version(__VERSION__):
            return

        def _ask():
            try:
                note_text = notes.strip() or "（无更新说明）"
                ok = messagebox.askyesno(
                    "发现新版本",
                    f"发现新版本 v{'.'.join(map(str, version))}，当前为 v{__VERSION__}。\n\n"
                    f"更新内容：\n{note_text[:600]}\n\n是否立即下载并更新？",
                )
                if ok:
                    _start_update_flow(root, download_url, version)
            finally:
                _end_update_check()

        root.after(0, _ask)
        scheduled = True
    except Exception:
        pass
    finally:
        if not scheduled:
            _end_update_check()


def detect_qr_codes(image_path: str):
    """返回图片中识别到的所有二维码文本列表（去重）"""
    # 懒加载：二维码识别相关的重型库仅在真正识别时才导入，缩短启动时间
    import cv2
    import numpy as np

    try:
        import zxingcpp
    except Exception:
        zxingcpp = None

    # 使用 numpy 读字节再 imdecode，避免 OpenCV 在 Windows 上处理中文路径的编码问题
    img = cv2.imdecode(np.fromfile(image_path, dtype=np.uint8), cv2.IMREAD_COLOR)
    if img is None:
        return []

    codes = []

    # 1. zxing-cpp 识别能力更强，先尝试；对小二维码会自动多尺度放大重试
    if zxingcpp is not None:
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


# ---------------------------------------------------------------- 并发基础设施


class CancelToken:
    """跨线程的「停止」标志。

    主线程点「停止」时调用 cancel()；各工作线程在开始处理下一张图片前检查 cancelled，
    尚未开始的任务会被快速取消。已经发出的网络请求无法中途打断，会自然收尾（≤ 超时时间）。
    """

    def __init__(self):
        self._event = threading.Event()

    def cancel(self):
        self._event.set()

    @property
    def cancelled(self) -> bool:
        return self._event.is_set()


class PdfNotAvailable(RuntimeError):
    """网址可达，但确实拿不到 PDF（业务性失败，重试没有意义）。"""


class DownloadNetworkError(RuntimeError):
    """网络类失败（超时、连接被重置等），值得重试。"""


class _TaskCtx:
    """并发任务共享的**只读**上下文。刻意不含任何会被写入的共享状态，
    这样多个线程并发处理时就不存在统计竞态。"""

    __slots__ = ("folder", "pdf_dir", "img_dir", "convert_pdf", "cancel", "log_queue")

    def __init__(self, folder, pdf_dir, img_dir, convert_pdf, cancel):
        self.folder = folder
        self.pdf_dir = pdf_dir
        self.img_dir = img_dir
        self.convert_pdf = convert_pdf
        self.cancel = cancel if cancel is not None else CancelToken()
        self.log_queue = None  # 需要即时输出（如「开始处理」心跳）时由 process_folder 注入


# 每个线程复用自己的 requests.Session（连接池复用，省掉每张发票重复的 TLS 握手）。
# requests.Session 并非线程安全，所以按线程隔离，而不是全局共享同一个。
_thread_local = threading.local()
_sessions_lock = threading.Lock()
_all_sessions: list = []


def _get_session() -> requests.Session:
    session = getattr(_thread_local, "session", None)
    if session is None:
        session = requests.Session()
        adapter = requests.adapters.HTTPAdapter(
            pool_connections=DEFAULT_WORKERS * 2,
            pool_maxsize=DEFAULT_WORKERS * 2,
            max_retries=0,  # 重试由 download_pdf 自己控制，避免双重重试放大请求量
        )
        session.mount("https://", adapter)
        session.mount("http://", adapter)
        _thread_local.session = session
        with _sessions_lock:
            _all_sessions.append(session)
    return session


def close_all_sessions():
    """任务结束后关闭各线程的 Session，释放连接。

    （关闭后的 Session 仍可继续使用，requests 会在需要时重建连接池，
    所以这里不需要额外清理线程局部变量。）
    """
    with _sessions_lock:
        sessions = list(_all_sessions)
        _all_sessions.clear()
    for s in sessions:
        try:
            s.close()
        except Exception:
            pass


def _unique_out_bases(files: list) -> dict:
    """为每张图片分配唯一的输出基名（不带扩展名）。

    为什么必须做：`a.jpg` 和 `a.png` 去掉扩展名后基名都是 `a`，会指向同一个
    `PDF/a.pdf`。串行处理时是"后覆盖前"，并发时会变成**两个线程同时写同一个文件 → 文件损坏**。
    这里预先给冲突的名字追加扩展名（再冲突就加序号）来消解。
    """
    used = set()
    mapping = {}
    for fname in files:
        stem, ext = os.path.splitext(fname)
        candidate = stem
        if candidate.lower() in used:
            candidate = f"{stem}_{ext.lstrip('.').lower()}"
        n = 2
        while candidate.lower() in used:
            candidate = f"{stem}_{ext.lstrip('.').lower()}_{n}"
            n += 1
        used.add(candidate.lower())
        mapping[fname] = candidate
    return mapping


def _looks_like_pdf(response: requests.Response) -> bool:
    """判断响应体是不是 PDF。

    优先看 Content-Type，兜底再看开头 4 个字节。
    ⚠️ 这里**不能**用 `response.content`：stream=True 的响应一旦读 `.content`，
    会把整个响应体一次性拉进内存，白耗内存、也失去了流式下载的意义。
    `raw.peek()` 只取缓冲区里已有的前几个字节，不会消费 body。
    """
    if "pdf" in response.headers.get("Content-Type", "").lower():
        return True
    try:
        head = response.raw.peek(5)
    except Exception:
        return False
    return head.lstrip()[:4] == b"%PDF"


def _save_response(response: requests.Response, save_path: str) -> None:
    """流式写入目标文件。

    先写 `<目标>.part` 再 `os.replace` 原子替换：这样下载中途失败时不会在「PDF」目录里
    留下半截的损坏 PDF（并发时尤其重要，半截文件很容易被误当成功结果）。
    """
    tmp_path = save_path + ".part"
    try:
        with open(tmp_path, "wb") as f:
            for chunk in response.iter_content(chunk_size=64 * 1024):
                if chunk:
                    f.write(chunk)
        os.replace(tmp_path, save_path)
    except Exception:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass
        raise


def _download_once(url: str, save_path: str, timeout: int, session: requests.Session) -> None:
    """单次下载尝试。

    失败时区分两类，交给上层决定要不要重试：
    - DownloadNetworkError：网络类问题（超时 / 连接被重置等）→ 值得重试；
    - PdfNotAvailable   ：网址能打开但确实拿不到 PDF → 重试无意义。
    """
    from urllib.parse import urljoin

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )
    }

    try:
        resp = session.get(url, headers=headers, timeout=timeout, stream=True, allow_redirects=True)
        resp.raise_for_status()
    except requests.RequestException as e:
        raise DownloadNetworkError(f"访问网址失败：{e}") from e

    if _looks_like_pdf(resp):
        _save_response(resp, save_path)
        return

    # 若是 HTML 页面，尝试找隐藏表单域并提交下载
    ct = resp.headers.get("Content-Type", "").lower()
    if "html" in ct:
        try:
            html = resp.text
        except requests.RequestException as e:
            raise DownloadNetworkError(f"读取页面内容失败：{e}") from e

        hidden_inputs = re.findall(r"<input[^>]+type=[\"']hidden[\"'][^>]*>", html, flags=re.IGNORECASE)
        fields = {}
        for tag in hidden_inputs:
            name_match = re.search("name=['\"]([^'\"]+)['\"]", tag, flags=re.IGNORECASE)
            value_match = re.search("value=['\"]([^'\"]*)['\"]", tag, flags=re.IGNORECASE)
            if name_match:
                fields[name_match.group(1)] = value_match.group(1) if value_match else ""

        if "idBase" in fields:
            download_url = urljoin(resp.url, "/download")
            try:
                dl_resp = session.post(
                    download_url, data=fields, headers=headers, timeout=timeout, stream=True
                )
                dl_resp.raise_for_status()
            except requests.RequestException as e:
                raise DownloadNetworkError(f"提交下载接口失败：{e}") from e

            if _looks_like_pdf(dl_resp):
                _save_response(dl_resp, save_path)
                return
            raise PdfNotAvailable(f"下载接口返回的不是 PDF：{dl_resp.text[:200]}")

    raise PdfNotAvailable("该网址没有直接返回 PDF，也未找到可下载的隐藏表单")


def download_pdf(
    url: str,
    save_path: str,
    timeout: int = 60,
    session: requests.Session | None = None,
    retries: int = DOWNLOAD_RETRIES,
) -> None:
    """下载 URL 指向的内容并保存为 PDF。

    支持两种常见情况：
    1. URL 直接返回 PDF 流；
    2. URL 返回发票展示页，页面里包含 name='idBase' 等隐藏域，
       此时自动提取并 POST 到 /download 获取 PDF。

    网络类失败会按 RETRY_BACKOFF 指数退避重试（默认 2 次）：一次网络抖动不该让发票
    被误判成「未下载」而要求人工重跑。「网址可达但没有 PDF」属业务性失败，不重试。
    """
    sess = session if session is not None else _get_session()
    last_err: Exception | None = None
    for attempt in range(retries + 1):
        try:
            _download_once(url, save_path, timeout, sess)
            return
        except DownloadNetworkError as e:
            last_err = e
            if attempt < retries:
                time.sleep(RETRY_BACKOFF * (2 ** attempt))
    raise last_err if last_err is not None else PdfNotAvailable("下载失败")


def convert_pdf_to_images(
    pdf_path: str,
    out_dir: str,
    base_name: str,
    page_fmt: str = "_第{n}页",
    name_fn=None,
) -> list[str]:
    """将 PDF 每一页渲染为 JPG 图片，长边 2000px，短边自适应；返回生成的文件路径列表

    page_fmt：默认命名模板（与旧版一致「原名_第N页.jpg」）。
    name_fn(base_name, page_no) -> 文件名：需要自定义命名 / 去重时传入，优先生效，
              返回 None 表示跳过该页。
    """
    import pymupdf  # 懒加载：仅在转图时才需要

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
            if name_fn is not None:
                out_name = name_fn(base_name, page_num + 1)
                if not out_name:
                    continue
            else:
                out_name = f"{base_name}{page_fmt.format(n=page_num + 1)}.jpg"
            out_path = os.path.join(out_dir, out_name)
            # 以 JPEG 输出（质量 92）：体积远小于 PNG，便于上传与分享。
            # alpha=False 已保证无透明通道，符合 JPEG 要求。
            try:
                pix.save(out_path, jpg_quality=92)
            except TypeError:
                # 兼容不支持 jpg_quality 参数的旧版 PyMuPDF
                with open(out_path, "wb") as fh:
                    fh.write(pix.tobytes("jpeg"))
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
    txt = str(s).replace(",", "").replace("，", "")  # 去除千分位逗号（中英文）
    m = re.search(r"-?\d+(?:\.\d+)?", txt)
    return m.group(0) if m else None


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
            pages_words = [page.extract_words() for page in pdf.pages]
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
_AUX_TAG = "重复份"        # 辅助列标记值：同一票据号第 2 张及以后
_AUX_HEADER = "重复份标记（辅助列，不参与展示）"


def _total_label(field):
    """统计区「合计 XX」的显示名（统筹沿用历史叫法「合计总统筹金额」）。"""
    return "合计总统筹金额" if field == _POOL_FIELD else "合计" + field


def _dup_label(field):
    """统计区「重复票据」类项目的显示名。"""
    return "重复票据统筹合计" if field == _POOL_FIELD else "重复票据" + field + "合计"


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
    headers = list(_SUMMARY_BASE_FIELDS)
    headers += [f for f in _SUMMARY_COND_FIELDS if f in include_cond]
    headers.append("是否重复")
    data_fields = headers[:-1]                      # 明细列（不含「是否重复」）
    ncols = len(headers)
    col_of = {name: i + 1 for i, name in enumerate(headers)}
    col_letter = {name: get_column_letter(c) for name, c in col_of.items()}
    dup_col = col_letter["是否重复"]
    ticket_col = col_letter[_TICKET_FIELD]
    # 金额类列（金额合计 + 各支付列）：写数字、套两位小数格式、合计行求和
    _TEXT_FIELDS = ("文件名", "交款人", "票据号码", "开票日期")
    money_fields = [f for f in data_fields if f not in _TEXT_FIELDS]
    deduct_fields = [f for f in _DEDUCT_FIELDS if f in col_of]
    # 辅助列：藏在统计区右侧，用来给「重复票据…合计」去重（判断某行是否重复份）
    aux_col_idx = ncols + 4
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

    # 「是否重复」判定逻辑保持不变：同一票据号码出现 >= 2 次即标记“是”
    _ticket_counts = Counter()
    for r in rows:
        tn = r.get(_TICKET_FIELD)
        if tn:
            _ticket_counts[tn] += 1

    for r in rows:
        tn = r.get(_TICKET_FIELD)
        dup = "是" if (tn and _ticket_counts[tn] >= 2) else ""
        values = []
        for f in data_fields:
            v = r.get(f)
            values.append(v if v is not None else "")
        ws.append(values + [dup])
        row_idx = ws.max_row
        for c in range(1, ncols + 1):
            cell = ws.cell(row=row_idx, column=c)
            cell.border = _SUMMARY_BORDER
            cell.alignment = _SUMMARY_CENTER
        # 金额列写入的是数字，套两位小数金额格式；
        # 空字符串（解析失败/缺失）保持为空，不影响求和。
        for f in money_fields:
            v = ws.cell(row=row_idx, column=col_of[f]).value
            if isinstance(v, (int, float)):
                ws.cell(row=row_idx, column=col_of[f]).number_format = _MONEY_FMT
        # 辅助列：判断本行是否为「同一票据号的重复份」（第 2 张及以后）。
        # COUNTIF 用**累计区间**（$C$2:$C本行），计数 > 1 说明前面已有同号票据。
        # 多张同号票据里，只有多出来的那些份计入「重复票据…合计」，
        # 首张的数据仍完整保留在明细里、也已计入「合计总金额」，不再重复累加。
        ws.cell(row=row_idx, column=aux_col_idx,
                value=('=IF(AND(${tc}{r}<>"",COUNTIF(${tc}$2:${tc}{r},${tc}{r})>1),'
                       '"{tag}","")'.format(tc=ticket_col, r=row_idx, tag=_AUX_TAG)))

    last = ws.max_row
    if last >= 2:
        # 合计行：各金额列分别求和
        total_row = [""] * ncols
        total_row[0] = "合计"
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

    # ---- 右侧统计区（标签列 / 数值列由数据列数推算，紧跟数据列空一列）----
    stat_label_col = ncols + 2
    stat_val_col = ncols + 3
    stat_vcol = get_column_letter(stat_val_col)
    stat_title_row = 1
    stat_first_row = 2

    amount_col = col_letter[_AMOUNT_FIELD]
    # 先按显示顺序排好标签，据此算出各统计项行号，再用行号拼公式，
    # 避免「可赔付金额」这类跨行引用随列数 / 项数变化而写错。
    stat_labels = (["票据张数", "合计总金额"]
                   + [_total_label(f) for f in deduct_fields]
                   + ["可赔付金额", "重复票据金额合计"]
                   + [_dup_label(f) for f in deduct_fields])
    row_of = {lab: stat_first_row + i for i, lab in enumerate(stat_labels)}

    # 可赔付金额 = 合计总金额 − 各支付类合计（统筹 / 大病保险 / 医疗救助 …
    # 这些都属于已由基金、保险或救助支付、患者并未实际支出的部分）
    deduct_cells = [f"{stat_vcol}{row_of[_total_label(f)]}" for f in deduct_fields]

    stat_items = [
        ("票据张数", f"=COUNTA(A2:A{last})"),
        ("合计总金额", f"=SUM({amount_col}2:{amount_col}{last})"),
    ]
    for f in deduct_fields:
        L = col_letter[f]
        stat_items.append((_total_label(f), f"=SUM({L}2:{L}{last})"))
    stat_items.append(
        ("可赔付金额",
         f"={stat_vcol}{row_of['合计总金额']}-" + "-".join(deduct_cells)))
    # 重复票据各项：只累计「重复份」（同票号第 2 张及以后），首张不重复计入
    stat_items.append(
        ("重复票据金额合计",
         f'=SUMIF({aux_col}2:{aux_col}{last},"{_AUX_TAG}",'
         f'{amount_col}2:{amount_col}{last})'))
    for f in deduct_fields:
        L = col_letter[f]
        stat_items.append(
            (_dup_label(f),
             f'=SUMIF({aux_col}2:{aux_col}{last},"{_AUX_TAG}",{L}2:{L}{last})'))
    # 标题
    tcell = ws.cell(row=stat_title_row, column=stat_label_col, value="统计")
    tcell.fill = _SUMMARY_FILL
    tcell.font = _SUMMARY_FONT
    tcell.alignment = _SUMMARY_CENTER
    tcell.border = _SUMMARY_BORDER
    tcell2 = ws.cell(row=stat_title_row, column=stat_val_col, value="数值")
    tcell2.fill = _SUMMARY_FILL
    tcell2.font = _SUMMARY_FONT
    tcell2.alignment = _SUMMARY_CENTER
    tcell2.border = _SUMMARY_BORDER
    # 数据行
    for i, (label, formula) in enumerate(stat_items):
        rrow = stat_first_row + i
        lc = ws.cell(row=rrow, column=stat_label_col, value=label)
        lc.font = Font(bold=True, color="1F4E78")
        lc.alignment = Alignment(horizontal="left", vertical="center")
        lc.border = _SUMMARY_BORDER
        vc = ws.cell(row=rrow, column=stat_val_col, value=formula)
        vc.font = Font(bold=True)
        vc.alignment = _SUMMARY_CENTER
        vc.border = _SUMMARY_BORDER
        vc.number_format = _MONEY_FMT if label != "票据张数" else "0"

    # 辅助列表头 + 隐藏（只为重复票据去重服务，不影响阅读）
    ac = ws.cell(row=1, column=aux_col_idx, value=_AUX_HEADER)
    ac.font = Font(size=9, color="808080")

    widths = [26, 14, 18, 14, 16, 18]
    widths += [18] * (len(headers) - len(_SUMMARY_BASE_FIELDS) - 1)  # 条件列
    widths += [12, 3, 26, 18, 10]   # 是否重复 | 间隔列 | 统计标签 | 统计数值 | 辅助列
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.column_dimensions[aux_col].hidden = True

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



_KIND_LABEL = {
    "success": "✓ 已下载 PDF",
    "no_pdf": "⚠ 识别到网址但未下载",
    "unrecognized": "✗ 未识别到二维码",
    "other": "· 其它情况（二维码非网址）",
    "error": "✗ 处理出错",
}


def _copy_to_unrecognized(folder: str, fpath: str, fname: str, prefix: str):
    """把问题图片复制一份到 folder/未识别/ 下，文件名前加 prefix；原图片保留不动。

    返回 (目标路径或 None, 日志行列表)。
    并发安全：目标名固定为「前缀 + 原文件名」，同一张图片只会被处理一次，
    不存在两个线程写同一个目标文件的情况。
    """
    sub = os.path.join(folder, "未识别")
    try:
        os.makedirs(sub, exist_ok=True)
    except Exception as e:
        return None, [f"  -> 创建「未识别」文件夹失败：{e}"]
    new_name = prefix + fname
    dest = os.path.join(sub, new_name)
    try:
        shutil.copy2(fpath, dest)
    except Exception as e:
        return None, [f"  -> 复制失败：{e}"]
    return dest, [f"  -> 已复制到「未识别」：{new_name}"]


def _process_one(ctx: _TaskCtx, fname: str, out_base: str) -> dict:
    """处理单张图片 —— 并发任务的最小单元。

    关键设计：**不写任何共享状态**。统计、进度、日志全部通过返回值交回主线程汇总，
    所以无论多少路并发都不会出现计数竞态，也不需要加锁。

    返回 {"fname", "kind", "elapsed", "lines"}，
    kind ∈ success / no_pdf / unrecognized / other / error / cancelled。
    """
    t0 = time.perf_counter()
    fpath = os.path.join(ctx.folder, fname)
    lines: list = []

    def finish(kind: str) -> dict:
        return {
            "fname": fname,
            "kind": kind,
            "elapsed": time.perf_counter() - t0,
            "lines": lines,
        }

    if ctx.cancel.cancelled:
        return finish("cancelled")

    # 1) 识别二维码
    try:
        codes = detect_qr_codes(fpath)
    except Exception as e:
        lines.append(f"  -> 识别过程出错：{e}（按「未识别」处理）")
        codes = []

    if not codes:
        _, msgs = _copy_to_unrecognized(ctx.folder, fpath, fname, PREFIX_UNRECOGNIZED)
        lines.extend(msgs)
        return finish("unrecognized")

    # 2) 在识别到的二维码中优先找一个网址
    url = None
    for code in codes:
        candidate = is_url(code)
        if candidate:
            url = candidate
            break

    if not url:
        lines.append("  -> 识别到二维码但非网址，归入「其它」")
        _, msgs = _copy_to_unrecognized(ctx.folder, fpath, fname, PREFIX_OTHER)
        lines.extend(msgs)
        return finish("other")

    # 3) 下载 PDF（内部带网络重试；「网址确实没有 PDF」属业务性失败，不重试）
    pdf_path = os.path.join(ctx.pdf_dir, f"{out_base}.pdf")
    try:
        download_pdf(url, pdf_path)
    except PdfNotAvailable as e:
        lines.append(f"  -> 网址无可下载的 PDF（{e}）")
        _, msgs = _copy_to_unrecognized(ctx.folder, fpath, fname, PREFIX_NOT_DOWNLOADED)
        lines.extend(msgs)
        return finish("no_pdf")
    except DownloadNetworkError as e:
        lines.append(f"  -> 网络失败，重试后仍未成功（{e}）")
        _, msgs = _copy_to_unrecognized(ctx.folder, fpath, fname, PREFIX_NOT_DOWNLOADED)
        lines.extend(msgs)
        return finish("no_pdf")
    except Exception as e:
        lines.append(f"  -> 下载失败（{e}）")
        _, msgs = _copy_to_unrecognized(ctx.folder, fpath, fname, PREFIX_NOT_DOWNLOADED)
        lines.extend(msgs)
        return finish("no_pdf")

    lines.append(f"  -> 已下载 PDF：{os.path.basename(pdf_path)}")

    # 4) 可选：PDF 转 JPG
    if ctx.convert_pdf and ctx.img_dir:
        try:
            for img_path in convert_pdf_to_images(pdf_path, ctx.img_dir, out_base):
                lines.append(f"  -> 已生成图片：{os.path.basename(img_path)}")
        except Exception as e:
            lines.append(f"  -> PDF 转图片失败：{e}")

    return finish("success")


# =====================================================================
# PDF 前置转换（独立工具）
#   只要目标文件夹里出现 PDF，就把 PDF 逐页转成 JPG，与原有图片一起收进
#   「处理后」子文件夹，再把后续识别/下载/汇总全部放到该子文件夹里进行。
#   全为图片时完全跳过，保持原有行为不变。
# =====================================================================

def _unique_path(dest_dir: str, name: str, used: set) -> str:
    """在 dest_dir 下为 name 找一个不与 used 冲突的文件名（大小写不敏感）。

    used 为「已占用的文件名小写集合」，命中则在名字尾部追加 _2、_3…
    （如 原文件名_1.jpg 已被原图片占用 → 原文件名_1_2.jpg）。
    """
    stem, ext = os.path.splitext(name)
    cand = name
    i = 1
    while cand.lower() in used:
        i += 1
        cand = f"{stem}_{i}{ext}"
    used.add(cand.lower())
    return cand


def prepare_target_folder(
    folder: str,
    log=print,
    cancel: "CancelToken | None" = None,
    sub_name: str = POSTPROCESS_DIR_NAME,
):
    """PDF 前置转换（独立步骤，可单独复用）：返回后续流程应处理的文件夹路径。

    三种场景：
      ① 全为图片           → 不新建任何文件夹，原路径原样返回（行为与旧版完全一致）；
      ② 全为 PDF           → 新建/重建「处理后」，PDF 逐页转成 JPG 放进去，返回「处理后」；
      ③ PDF + 图片混合     → 新建/重建「处理后」，PDF 转图 + 复制原有图片一起放进去，返回「处理后」。

    约定：
      - 转出的页面命名「原文件名_1.jpg / 原文件名_2.jpg …」；
      - 原有图片是**复制**，原始文件夹里的文件一律保留不动；
      - 「处理后」每次都**清空重建**，避免上一次的遗留结果干扰本次处理；
      - 命名冲突自动加后缀区分，不覆盖已有文件；
      - 单个 PDF 转换失败（损坏 / 加密 / 无权限）只记日志跳过，不中断整体。

    返回 (目标文件夹, info)；info 字段：
      has_pdf / converted_pages / converted_pdfs / copied_images / failed / skipped / target。
    """
    cancel = cancel if cancel is not None else CancelToken()
    info = {
        "has_pdf": 0, "converted_pages": 0, "converted_pdfs": 0,
        "copied_images": 0, "failed": 0, "skipped": False, "target": folder,
    }

    try:
        entries = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
    except Exception as e:
        log(f"前置转换：读取文件夹失败（{e}），按原文件夹继续处理。")
        return folder, info

    pdfs = sorted(
        [f for f in entries if f.lower().endswith(SUPPORTED_PDF_EXTS)], key=lambda x: x.lower()
    )
    images = sorted(
        [f for f in entries if f.lower().endswith(SUPPORTED_IMAGE_EXTS)], key=lambda x: x.lower()
    )

    # ① 没有 PDF：跳过前置转换，原逻辑照旧
    if not pdfs:
        if images:
            log(f"前置转换：未发现 PDF（{len(images)} 个图片文件），跳过转换，直接处理原文件夹。")
        return folder, info

    info["has_pdf"] = len(pdfs)
    log(f"前置转换：发现 {len(pdfs)} 个 PDF、{len(images)} 个图片文件 → 先转图片，"
        f"并统一放入「{sub_name}」子文件夹后继续处理。")

    target = os.path.join(folder, sub_name)

    # 「处理后」清空重建（要求：每次重建，不带上次遗留）
    try:
        if os.path.isdir(target):
            shutil.rmtree(target)
            log(f"前置转换：已清空旧的「{sub_name}」文件夹。")
        elif os.path.exists(target):
            os.remove(target)
        os.makedirs(target, exist_ok=True)
    except Exception as e:
        log(f"前置转换：建立「{sub_name}」失败（{e}），按原文件夹继续处理。")
        return folder, info

    info["target"] = target
    used: set = set()

    # ② 先复制原有图片（原文件保留不动），占住文件名避免与转出的页面重名
    for name in images:
        if cancel.cancelled:
            info["skipped"] = True
            log("前置转换：已请求停止，剩余图片未复制。")
            break
        dest_name = _unique_path(target, name, used)
        try:
            shutil.copy2(os.path.join(folder, name), os.path.join(target, dest_name))
            info["copied_images"] += 1
        except Exception as e:
            info["failed"] += 1
            log(f"  -> 复制原图片失败：{name}（{e}）")

    # ③ 逐份 PDF 转 JPG，命名「原文件名_1.jpg / _2.jpg …」
    for name in pdfs:
        if cancel.cancelled:
            info["skipped"] = True
            log("前置转换：已请求停止，剩余 PDF 未转换。")
            break
        pdf_path = os.path.join(folder, name)
        base = os.path.splitext(name)[0]

        def _name_fn(b, n, _t=target, _u=used):
            # 用「原名_页码.jpg」，与已有文件冲突时自动加后缀
            return _unique_path(_t, f"{b}_{n}.jpg", _u)

        try:
            made = convert_pdf_to_images(pdf_path, target, base, name_fn=_name_fn)
            info["converted_pages"] += len(made)
            info["converted_pdfs"] += 1
            log(f"  -> {name}：转换 {len(made)} 页 → "
                f"{os.path.basename(made[0]) if made else '（无内容）'}"
                + (f" … {os.path.basename(made[-1])}" if len(made) > 1 else ""))
        except Exception as e:
            info["failed"] += 1
            log(f"  -> PDF 转换失败：{name}（{e}）")

    log(f"前置转换完成：PDF {info['converted_pdfs']}/{info['has_pdf']} 份、"
        f"共 {info['converted_pages']} 页转成图片，复制原图片 {info['copied_images']} 个"
        + (f"，失败 {info['failed']} 个" if info["failed"] else "")
        + f"；后续在「{sub_name}」内处理。")

    return target, info


def process_folder(
    folder: str,
    open_after: bool,
    convert_pdf: bool,
    summarize: bool,
    log_queue: queue.Queue,
    cancel: "CancelToken | None" = None,
    workers: int = DEFAULT_WORKERS,
    preconvert: bool = True,
):
    """处理整个文件夹：每张图片一个任务并发执行，结果由本线程统一汇总。

    前置步骤（preconvert=True 时）：目标文件夹里只要有 PDF，就先做「PDF → JPG」转换，
    并把转换结果与原有图片一起收进「处理后」子文件夹，再在该子文件夹内执行后续全部流程；
    全为图片时跳过该步骤，行为与旧版完全一致。

    并发只用在「识别 + 下载 + 转图」这条流水线上，最后的「汇总发票」保持串行
    —— pdfplumber 是纯 Python 解析、被 GIL 锁死，实测并发 1.0× 无收益（甚至略慢）。
    """
    def log(msg: str):
        log_queue.put(("log", msg))

    def progress(current: int, total: int):
        # 带上并发路数，界面可直接显示「进度：x / y 份 · 6 路并发」
        log_queue.put(("progress", current, total, workers))

    cancel = cancel if cancel is not None else CancelToken()
    workers = max(1, int(workers))

    if not folder or not os.path.isdir(folder):
        log("错误：请选择一个有效的文件夹路径。")
        log_queue.put(("done",))
        return

    # ★ 前置：PDF → JPG 转换（有 PDF 时才动作；全图片直接跳过）
    if preconvert:
        folder, _pc_info = prepare_target_folder(folder, log=log, cancel=cancel)
        if cancel.cancelled:
            log("已在「前置转换」阶段停止，本次未开始识别。")
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

    if total == 0:
        progress(0, 0)
        log("没有需要处理的图片。")
        log_queue.put(("stats", {
            "total": 0, "success": 0, "no_pdf": 0, "unrecognized": 0,
            "other": 0, "error": 0, "cancelled": 0,
            "workers": workers, "elapsed": 0.0,
        }))
        log_queue.put(("done",))
        return

    # 输出基名唯一化：避免 a.jpg 与 a.png 同时写同一个 a.pdf（并发下会写坏文件）
    out_bases = _unique_out_bases(files)
    ctx = _TaskCtx(folder, pdf_dir, img_dir, convert_pdf, cancel)
    ctx.log_queue = log_queue

    log(f"开始处理：{total} 张图片，{workers} 路并发。")
    # 起始进度归零：旧版是在「开始处理第 N 张之前」就上报 N，
    # 于是进度条先跳一格再干活，且卡在下载时会长时间不动，看着像死机。
    progress(0, total)

    results: list = []
    done = 0
    cancel_swept = False
    t_start = time.perf_counter()

    # 心跳：每 3 秒报一次进度，大批量时日志不会长时间空白
    heartbeat_stop = threading.Event()

    def _heartbeat():
        while not heartbeat_stop.wait(3.0):
            if cancel.cancelled:
                return
            left = total - done
            if left > 0:
                log(f"  …处理中：已完成 {done}/{total} 张，剩余 {left} 张（{workers} 路并发）")

    threading.Thread(target=_heartbeat, daemon=True).start()

    try:
        with ThreadPoolExecutor(max_workers=workers) as executor:
            future_map = {
                executor.submit(_process_one, ctx, fname, out_bases[fname]): fname
                for fname in files
            }
            for fut in as_completed(future_map):
                fname = future_map[fut]
                try:
                    res = fut.result()
                except CancelledError:
                    # 被「停止」直接取消、根本没开始的任务（不是错误）
                    res = {"fname": fname, "kind": "cancelled", "elapsed": 0.0, "lines": []}
                except Exception as e:      # 兜底：_process_one 内部已尽量不抛异常
                    res = {
                        "fname": fname, "kind": "error", "elapsed": 0.0,
                        "lines": [f"  -> 处理失败：{e}"],
                    }
                results.append(res)

                if res["kind"] != "cancelled":
                    done += 1
                    label = _KIND_LABEL.get(res["kind"], res["kind"])
                    # 每张图片的日志整块连续输出（标题行 + 明细行），
                    # 这样并发交错也不会出现「-> 已下载」看不出属于哪张图的问题
                    log(f"[{done}/{total}] {fname}（{res['elapsed']:.2f}s）{label}")
                    for line in res["lines"]:
                        log(line)
                    progress(done, total)

                if cancel.cancelled and not cancel_swept:
                    cancel_swept = True
                    for f in future_map:
                        if not f.done():
                            f.cancel()      # 尚未开始的任务直接取消，不等它跑完
    finally:
        heartbeat_stop.set()
        close_all_sessions()

    elapsed = time.perf_counter() - t_start
    kind_count = Counter(r["kind"] for r in results)
    # 统计在并发结束后统一汇总，避免在多个线程里做 `stats[k] += 1`（非原子，会丢计数）
    stats = {
        "total": total,
        "success": kind_count["success"],
        "no_pdf": kind_count["no_pdf"],
        "unrecognized": kind_count["unrecognized"],
        "other": kind_count["other"],
        "error": kind_count["error"],
        "cancelled": kind_count["cancelled"],
        "workers": workers,
        "elapsed": elapsed,
    }

    if cancel.cancelled:
        log(f"已停止：本次完成 {done}/{total} 张，其余 {total - done} 张未处理（原图保持不动）。")
    else:
        log("全部处理完成。")

    # 输出识别结果统计
    summary = [
        "=== 识别结果统计 ===",
        f"总计识别图片：{stats['total']} 张",
        f"✓ 成功识别并下载 PDF（原文件名不变）：{stats['success']} 张",
        f"⚠ 识别到网址但未下载 PDF（复制到未识别/，未下载-）：{stats['no_pdf']} 张",
        f"✗ 未识别到二维码（复制到未识别/，未识别-）：{stats['unrecognized']} 张",
        f"· 其它情况（复制到未识别/，其它-）：{stats['other']} 张",
    ]
    if stats["error"]:
        summary.append(f"✗ 处理出错（详见上方日志）：{stats['error']} 张")
    if stats["cancelled"]:
        summary.append(f"■ 因「停止」未处理：{stats['cancelled']} 张")
    summary.append(
        f"耗时：{elapsed:.1f} 秒（{workers} 路并发，平均 {elapsed / max(1, done):.2f} 秒/张）"
    )
    for line in summary:
        log(line)

    # 把结构化统计传给 GUI（用于结束弹窗）
    log_queue.put(("stats", dict(stats)))

    if cancel.cancelled:
        log("已停止：跳过汇总与打开文件夹。再次点「开始处理」会接着处理剩余图片，"
            "汇总时会扫描 PDF 文件夹里的全部发票，已下载的结果不会丢失。")
    else:
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


class _FlatProgressBar(tk.Canvas):
    """扁平进度条：浅灰轨道 + 细边框 + 蓝色填充，与整体界面风格保持一致。

    系统默认进度条在 Windows 上是绿色渐变，跟这套扁平浅色界面不搭，故自绘。"""

    TRACK = "#EDEFF2"
    BORDER = "#D6DBE1"
    FILL = "#2563EB"

    def __init__(self, master, height: int = 20, **kw):
        super().__init__(
            master, height=height, bg=self.TRACK, highlightthickness=0, bd=0, **kw
        )
        self._value = 0.0
        self.bind("<Configure>", lambda _e: self._redraw())

    def set_value(self, pct: float):
        try:
            pct = float(pct)
        except (TypeError, ValueError):
            pct = 0.0
        self._value = max(0.0, min(100.0, pct))
        self._redraw()

    def _redraw(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w <= 2 or h <= 2:
            return
        self.create_rectangle(
            0, 0, w - 1, h - 1, outline=self.BORDER, fill=self.TRACK
        )
        fill_w = (w - 2) * self._value / 100.0
        if fill_w >= 1:
            self.create_rectangle(
                1, 1, 1 + fill_w, h - 2, outline="", fill=self.FILL
            )


class InvoiceQrToolApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"发票二维码识别下载工具 v{__VERSION__}")
        self.root.geometry("800x600")
        self.root.minsize(700, 450)
        apply_window_icon(self.root)

        self.folder_var = tk.StringVar()
        self.open_after_var = tk.BooleanVar(value=True)
        self.convert_pdf_var = tk.BooleanVar(value=False)
        self.summarize_var = tk.BooleanVar(value=False)

        self.log_queue: queue.Queue = queue.Queue()
        self.worker_thread: threading.Thread | None = None
        self.last_stats: dict | None = None
        self.cancel_token: CancelToken | None = None
        self._busy = False

        self._build_ui()
        self._poll_log()
        # 关闭窗口时若任务还在跑，先让用户确认（避免误关导致半途而废）
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        # 启动后静默检查更新（仅发现新版本时弹窗，无网络/无更新时不打扰）
        threading.Thread(
            target=_silent_startup_check, args=(self.root,), daemon=True
        ).start()

    def _build_menubar(self):
        """顶部菜单栏：承载「使用说明 / 检查更新 / 清空日志」等次要功能。"""
        self.menubar = tk.Menu(self.root)

        self.menu_help = tk.Menu(self.menubar, tearoff=0)
        self.menu_help.add_command(
            label="使用说明 / 更新记录", command=self._show_help
        )
        self.menu_help.add_command(label="检查更新", command=self._check_update)
        self.menu_help.add_separator()
        self.menu_help.add_command(label="清空日志", command=self._clear_log)
        self.menubar.add_cascade(label="帮助", menu=self.menu_help)

        self.root.config(menu=self.menubar)

    def _build_ui(self):
        pad = {"padx": 10, "pady": 8}

        self._build_menubar()

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
            text="将下载的 PDF 转换为图片（JPG，长边 2000px）",
            variable=self.convert_pdf_var,
        ).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(
            frame_opts,
            text="处理完成后汇总发票（生成 Excel）",
            variable=self.summarize_var,
        ).pack(side=tk.LEFT, padx=5)

        # 操作按钮区：整行扁平主按钮（浅底 + 细边框 + 居中文字），点击区域大、易点；
        # 右侧配「■ 停止」，处理过程中可随时中止（尚未开始的任务会被直接取消）。
        # 「使用说明 / 检查更新」等次要功能已移入顶部菜单栏。
        frame_btn_row = ttk.Frame(self.root)
        frame_btn_row.pack(fill=tk.X, padx=10, pady=(12, 6))

        frame_btn_border = tk.Frame(frame_btn_row, bg="#D6DBE1")  # 外层充当 1px 细边框
        frame_btn_border.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.btn_start = tk.Button(
            frame_btn_border,
            text="▶  开始处理",
            command=self._start_processing,
            font=("Microsoft YaHei", 12, "bold"),
            bg="#FAFAFB",
            fg="#1F2937",
            activebackground="#EEF1F4",
            activeforeground="#1F2937",
            disabledforeground="#9CA3AF",
            relief=tk.FLAT,
            bd=0,
            highlightthickness=0,
            cursor="hand2",
            pady=11,
        )
        self.btn_start.pack(fill=tk.X, padx=1, pady=1)

        # 停止按钮：与主按钮同高、同样扁平细边框风格，但更窄
        frame_stop_border = tk.Frame(frame_btn_row, bg="#D6DBE1")
        frame_stop_border.pack(side=tk.LEFT, padx=(6, 0), fill=tk.Y)
        self.btn_stop = tk.Button(
            frame_stop_border,
            text="■  停止",
            command=self._stop_processing,
            font=("Microsoft YaHei", 11, "bold"),
            bg="#FAFAFB",
            fg="#B91C1C",
            activebackground="#FEF2F2",
            activeforeground="#B91C1C",
            disabledforeground="#C7CBD1",
            relief=tk.FLAT,
            bd=0,
            highlightthickness=0,
            cursor="hand2",
            state=tk.DISABLED,
            width=9,
        )
        self.btn_stop.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        # 进度条 + 右侧「进度：x / y 份」计数，同一行展示
        frame_progress = ttk.Frame(self.root)
        frame_progress.pack(fill=tk.X, padx=10, pady=(0, 8))
        self.progress = _FlatProgressBar(frame_progress, height=20)
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.progress_var = tk.StringVar(
            value=f"进度：0 / 0 份 · {DEFAULT_WORKERS} 路并发"
        )
        ttk.Label(
            frame_progress, textvariable=self.progress_var, width=24, anchor=tk.E
        ).pack(side=tk.LEFT, padx=(10, 2))

        # 日志区（标题行右侧提供「清空日志」，便于开始下一个任务前清空）
        frame_log_head = ttk.Frame(self.root)
        frame_log_head.pack(fill=tk.X, padx=10)
        ttk.Label(frame_log_head, text="处理日志：").pack(side=tk.LEFT)
        ttk.Button(
            frame_log_head, text="清空日志", command=self._clear_log, width=10
        ).pack(side=tk.RIGHT)
        self.txt_log = scrolledtext.ScrolledText(
            self.root, wrap=tk.WORD, state=tk.DISABLED, height=20
        )
        self.txt_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=(4, 10))

    def _browse_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folder_var.set(path)

    def _check_update(self):
        if self._busy:
            self._log("处理任务进行中，暂不检查更新（避免打断当前任务）。")
            return
        self._log("正在检查更新…")
        threading.Thread(
            target=check_and_prompt_update, args=(self.root,), daemon=True
        ).start()

    def _show_help(self):
        win = tk.Toplevel(self.root)
        win.title("使用说明 / 更新记录")
        win.geometry("660x540")
        try:
            win.transient(self.root)
            win.grab_set()
        except Exception:
            pass
        apply_window_icon(win)
        txt = scrolledtext.ScrolledText(win, wrap=tk.WORD, state=tk.NORMAL)
        txt.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        txt.insert(tk.END, USAGE_TEXT + "\n\n" + CHANGELOG_TEXT)
        txt.configure(state=tk.DISABLED)
        ttk.Button(win, text="关闭", command=win.destroy).pack(pady=6)

    def _clear_log(self):
        """清空日志区内容（不影响正在运行的任务），便于开始下一个任务。"""
        self.txt_log.configure(state=tk.NORMAL)
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.configure(state=tk.DISABLED)

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
                    current, total = item[1], item[2]
                    workers = item[3] if len(item) > 3 else DEFAULT_WORKERS
                    if total > 0:
                        self.progress.set_value((current / total) * 100)
                        self.progress_var.set(
                            f"进度：{current} / {total} 份 · {workers} 路并发"
                        )
                    else:
                        self.progress.set_value(100)
                        self.progress_var.set(
                            f"进度：0 / 0 份 · {DEFAULT_WORKERS} 路并发"
                        )
                elif item[0] == "stats":
                    self.last_stats = item[1]
                elif item[0] == "done":
                    self._finish_task()
        except queue.Empty:
            pass
        except tk.TclError:
            return          # 窗口已销毁，停止轮询
        try:
            self.root.after(100, self._poll_log)
        except tk.TclError:
            pass

    def _finish_task(self):
        """任务收尾：复位按钮与状态，弹出统计。"""
        self._busy = False
        self.cancel_token = None
        _PROCESSING.clear()
        self.progress.set_value(100)
        self.btn_start.configure(state=tk.NORMAL, text="▶  开始处理")
        self.btn_stop.configure(state=tk.DISABLED, text="■  停止")
        self._log("--- 任务结束 ---")
        s = self.last_stats
        if s:
            self._show_stats_popup(s)
            self.last_stats = None

    def _show_stats_popup(self, s: dict):
        msg = (
            "本次识别结果统计：\n\n"
            f"总计识别图片：{s['total']} 张\n"
            f"成功下载 PDF（原文件名不变）：{s['success']} 张\n"
            f"识别但未下载 PDF（未下载-）：{s['no_pdf']} 张\n"
            f"未识别到二维码（未识别-）：{s['unrecognized']} 张\n"
            f"其它情况（其它-）：{s['other']} 张"
        )
        if s.get("error"):
            msg += f"\n处理出错（详见日志）：{s['error']} 张"
        cancelled = s.get("cancelled") or 0
        if cancelled:
            msg += f"\n因「停止」未处理：{cancelled} 张"
        elapsed = s.get("elapsed") or 0.0
        if elapsed:
            processed = max(1, s["total"] - cancelled)
            msg += (
                f"\n\n耗时：{elapsed:.1f} 秒"
                f"（{s.get('workers', DEFAULT_WORKERS)} 路并发，"
                f"平均 {elapsed / processed:.2f} 秒/张）"
            )
        messagebox.showinfo("识别结果统计", msg)

    def _start_processing(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("路径错误", "请选择一个有效的文件夹。")
            return
        if self._busy:
            return

        self._busy = True
        _PROCESSING.set()
        self.cancel_token = CancelToken()

        self.btn_start.configure(state=tk.DISABLED, text="处理中…")
        self.btn_stop.configure(state=tk.NORMAL, text="■  停止")
        self.progress.set_value(0)
        self.progress_var.set(f"进度：0 / 0 份 · {DEFAULT_WORKERS} 路并发")
        self._log(f"=== 开始处理（{DEFAULT_WORKERS} 路并发）===")

        self.worker_thread = threading.Thread(
            target=process_folder,
            args=(
                folder,
                self.open_after_var.get(),
                self.convert_pdf_var.get(),
                self.summarize_var.get(),
                self.log_queue,
            ),
            kwargs={"cancel": self.cancel_token, "workers": DEFAULT_WORKERS},
            daemon=True,
        )
        self.worker_thread.start()

    def _stop_processing(self):
        """请求中止：尚未开始的任务会被直接取消，已发出的网络请求会自然收尾。"""
        if not self._busy or self.cancel_token is None:
            return
        self.btn_stop.configure(state=tk.DISABLED, text="正在停止…")
        self.cancel_token.cancel()
        self._log("--- 已请求停止：等待已开始的任务收尾（已发出的网络请求无法中途打断）---")

    def _on_close(self):
        """关闭窗口：处理中先确认，避免误关导致任务半途而废。"""
        if self._busy:
            ok = messagebox.askyesno(
                "确认退出",
                "任务正在处理中，现在退出会中断本次处理（已下载的 PDF 会保留）。\n\n"
                "确定要退出吗？",
            )
            if not ok:
                return
            if self.cancel_token is not None:
                self.cancel_token.cancel()
        self.root.destroy()


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
    close_all_sessions()
    with open(log_path, "w", encoding="utf-8") as f:
        while True:
            item = q.get()
            if item[0] == "log":
                f.write(item[1] + "\n")
            elif item[0] == "stats":
                s = item[1]
                f.write(
                    f"--- 统计：总 {s['total']} / 成功 {s['success']} / 未下载 {s['no_pdf']} / "
                    f"未识别 {s['unrecognized']} / 其它 {s['other']} / 出错 {s.get('error', 0)} / "
                    f"并发 {s.get('workers', 0)} / 耗时 {s.get('elapsed', 0):.2f}s ---\n"
                )
            elif item[0] == "done":
                f.write("--- 测试完成 ---\n")
                break


def _cleanup_legacy_update_artifacts():
    """清理早期版本（bat / --pending / 备份机制）遗留的临时文件，避免堆积。

    另外清理临时目录里的半截更新包（.part）—— 之前若更新被强杀 / 进程卡死，
    会留下上百 MB 的无用文件且再没人管它。
    """
    try:
        work_dir = os.path.dirname(sys.executable)
        for name in ("_pending.exe", "_backup.exe", "_replace_in_progress",
                     "_update_rolled_back", "invoiceqrdl_update.bat"):
            p = os.path.join(work_dir, name)
            if os.path.isfile(p):
                try:
                    os.remove(p)
                except Exception:
                    pass
    except Exception:
        pass
    # 半截更新包：正被占用（说明有更新在进行）时删不掉，静默跳过即可
    try:
        part = os.path.join(tempfile.gettempdir(),
                            "InvoiceQRDownloader_update.part")
        if os.path.isfile(part):
            os.remove(part)
    except Exception:
        pass


def _cli_pdf2img(folder: str) -> None:
    """独立工具入口：只做「PDF → JPG」前置转换，不识别、不下载、不汇总。

    效果与软件内置的前置步骤一致：在目标文件夹内新建（或清空重建）「处理后」，
    把 PDF 逐页转成「原文件名_1.jpg …」并与原有图片一起放进去。
    """
    print(f"PDF 前置转换：{folder}")
    target, info = prepare_target_folder(folder, log=print)
    print(f"完成：PDF {info['converted_pdfs']}/{info['has_pdf']} 份、"
          f"{info['converted_pages']} 页 → 图片；复制原图片 {info['copied_images']} 个。")
    print(f"输出目录：{target}")


def main():
    args = sys.argv[1:]

    # “回收旧版本”模式：由上一版本以 --recycle-old "<旧exe>" 启动本进程，
    # 等旧进程退出后再把旧 exe 移入回收站（避免删除正在运行的自身文件被系统锁定）。
    if "--recycle-old" in args:
        i = args.index("--recycle-old")
        old_path = args[i + 1] if i + 1 < len(args) else None
        if old_path and old_path.lower().endswith(".exe"):
            for _ in range(40):  # 最多约 20 秒，等旧进程释放文件
                if not os.path.exists(old_path):
                    break
                if send_to_recycle_bin(old_path):
                    break
                time.sleep(0.5)

    # 清理早期更新机制遗留的临时文件
    _cleanup_legacy_update_artifacts()

    # 独立工具模式：仅把文件夹里的 PDF 转成图片（含「处理后」整理），到此为止
    if "--pdf2img" in args:
        i = args.index("--pdf2img")
        folder = args[i + 1] if i + 1 < len(args) else None
        if not folder:
            print("用法：发票二维码工具.exe --pdf2img <文件夹>")
            return
        _cli_pdf2img(folder)
        return

    if "--test" in args:
        i = args.index("--test")
        folder = args[i + 1] if i + 1 < len(args) else None
        if not folder:
            print("用法：发票二维码工具.exe --test <文件夹> [--summary]")
            return
        summarize = "--summary" in args
        _cli_test(folder, summarize=summarize)
        return

    setup_app_id()          # 必须在 tk.Tk() 之前，任务栏才会认本程序的图标与身份
    root = tk.Tk()
    app = InvoiceQrToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
