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
   金额/统筹为纯数字、无千分位；「是否重复」列互指重复行号，右侧分
   「统计（剔重后）」与「重复票据」两块汇总）。

并发模型：1~3 步由线程池并发执行（默认 6 路，见 DEFAULT_WORKERS）；
第 6 步的汇总保持串行（pdfplumber 是纯 Python 解析，并发无收益）。
处理中可通过 CancelToken 请求停止。

界面：卡片化浅色界面（顶栏 + 圆角白卡 + 蓝色主色），整行「▶ 开始处理」主按钮
+ 右侧「停止」+ 自绘圆角进度条；「使用说明 / 检查更新」位于顶栏。
界面尺寸与配色集中在 ui_kit.py 的 SKIN（= PRESETS["标准"]），改外观只改那一处。
"""

import os
import re
import sys
import time
import queue
import shutil
import threading
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed, CancelledError
# ⚠️ ui_kit 必须早于 import tkinter：它 import 时即设置进程 DPI 感知，
#    顺序反了会让高分屏（125%/150%）下的界面被位图拉伸、文字发虚。
import ui_kit as K
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# 本地模块（原为本文件的一部分，见各模块 docstring）：
#   iqr_net     —— 会话管理与下载
#   iqr_summary —— 发票解析与汇总表写出
#   iqr_update  —— 自动更新
from iqr_net import (
    DEFAULT_WORKERS,
    CancelToken, PdfNotAvailable, DownloadNetworkError,
    close_all_sessions, download_pdf,
)
from iqr_summary import (
    parse_invoice, summarize_invoices, _write_summary_excel,
    # 汇总表的「契约」：列名、标签、工具函数。保留在主程序命名空间里，
    # 外部脚本与自测一直通过 invoice_qr_tool 访问它们。
    _SUMMARY_BASE_FIELDS, _SUMMARY_COND_FIELDS, _DEDUCT_FIELDS,
    _AMOUNT_FIELD, _POOL_FIELD, _TICKET_FIELD, _SEQ_FIELD, SEQ_COL_WIDTH,
    _AUX_TAG, _AUX_FIRST, _norm_colon, _Pages, _field_search,
    _normalize_date, _to_num, _inject_formula_cache,
)
from iqr_update import (
    GITHUB_REPO_OWNER, GITHUB_REPO_NAME,
    _parse_version, _version_str, _derive_base_name,
    get_latest_release, perform_update, send_to_recycle_bin,
    _cleanup_legacy_update_artifacts,
)


# 注意：cv2 / numpy / pymupdf / pdfplumber / openpyxl / zxingcpp 等重型依赖
# 在启动时并不需要（GUI 与更新检查仅用到 tkinter + requests）。它们改为在
# 对应功能被真正调用时才“懒加载”，可大幅缩短启动时间（窗口更快出现）。
# 懒加载由各功能函数内部 import 完成，Python 会缓存已导入模块，重复调用无额外开销。

URL_RE = re.compile(r"https?://[^\s<>\"{}|\\^`\[\]]+", re.IGNORECASE)
# 二维码文本里 URL 后面常紧跟中文标点（如「请访问 https://xxx。」）：这些字符不可能
# 属于合法 URL，却会被上面那条正则一并吞进去，导致下载失败，故统一削掉尾随标点。
URL_TAIL_JUNK = "。，、；：！？）】》」”’…"

SUPPORTED_IMAGE_EXTS = (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp")
SUPPORTED_PDF_EXTS = (".pdf",)

# 「PDF 前置转换」子文件夹名：只要目标文件夹里出现 PDF，就把 PDF 转出的图片与原有图片
# 一起收进这个子文件夹，再在其中按原逻辑处理（输出 PDF/ 未识别/ 汇总表也都落在它里面）。
POSTPROCESS_DIR_NAME = "处理后"

# 状态前缀：把“未下载 / 未识别 / 其它”的图片复制到「未识别」子文件夹时加在文件名前
PREFIX_UNRECOGNIZED = "未识别-"  # 未识别到任何二维码
PREFIX_NOT_DOWNLOADED = "未下载-"  # 识别到网址但无可下载的 PDF
PREFIX_OTHER = "其它-"           # 其它情况（识别到二维码但内容非网址等）

# 二维码识别的内存护栏。zxing-cpp 识别小二维码靠「逐级放大重试」，而放大后的位图是
# 实体内存（宽 × 高 × 3 字节）：4000×3000 再放大 2× 就是 8000×6000 ≈ 144MB / 张，
# 6 路并发叠起来接近 1GB，大批量处理时容易把机器拖垮。下面两个上限把单张占用封住：
#   QR_MAX_SOURCE_PIXELS —— 原图自身超预算就先等比降采样（超大照片 / 扫描件）；
#   QR_MAX_SCALE_PIXELS  —— 放大后位图的像素预算，据此推导最多放大几倍。
# 按此规则，小图仍会放大到 4×（与原分级几乎一致），大图则不再做无谓放大。
QR_MAX_SOURCE_PIXELS = 20_000_000    # 原图像素上限（约 60MB/张）
QR_MAX_SCALE_PIXELS = 12_000_000     # 放大后位图像素预算（约 36MB/张）

# 软件自身版本与 GitHub 更新源（公开仓库，更新检查无需鉴权）
__VERSION__ = "5.0.0"


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
   汇总为「PDF/发票汇总_YYYYMMDD_HHMMSS.xlsx」；最左侧为「序号」列（1、2、3…，
   配合文件名一眼看出是第几份），金额类列均为纯数字（无千分位、可直接求和与二次计算）。
   「是否重复」列不再只标「是」，而是**互指序号**（如「与序号3、5重复」），
   同一票号各有几张、分别是第几份一眼可见（编号与最左「序号」列一致）。
   表格右侧是两块汇总：
   - 左块「统计（剔重后）」：票据张数 / 合计总金额 / 合计总统筹金额 /
     〔合计大病保险支付〕/〔合计医疗救助支付〕/ 可赔付金额
     —— **全部按“剔除重复票据”之后的口径统计**（同一票号只算第 1 张）。
   - 右块「重复票据」：重复票据张数 / 重复票据金额合计 / 重复票据统筹合计 /
     〔重复票据大病保险支付合计〕/〔重复票据医疗救助支付合计〕
     —— 只统计**多出来的份**（同一票号第 2 张及以后），首张不重复计入。
   说明：
   - 加〔〕的项目与对应支付列一样，**仅在识别到该类目时才出现**（识别到 0.00 也算“有”该类目）；
     本批票据若一张都没有，该列连同其汇总项整体省略，输出与旧版一致。
   - 可赔付金额 = 合计总金额 − 各支付类合计之和（统筹 / 大病保险 / 医疗救助）。
   - 明细底部的「合计」行照旧对**全部**明细行求和（不剔重），方便与原始票据逐张核对；
     左块与右块的张数之和 = 明细行数，金额之和 = 明细合计行金额。
   - 重复判定在**软件内部**完成，不依赖 Excel 公式比较票号：票据号码常有 15 位以上，
     而表格软件的 COUNTIF / SUMIF 会把「数字样文本」按数字比较、只保留 15 位有效数字，
     会导致重复判定整体失效（详见更新记录 v4.9.0）。
7. 处理结束后，弹出「识别结果统计」：总计识别图片数、成功下载 PDF、识别但未下载、
   未识别、其它各多少张，并附本次耗时与平均每张耗时，方便核对处理结果。
8. 日志区右上角提供「清空日志」按钮，一键清空历史日志，便于开始下一个任务。
9. 多张图片**并发处理**（默认 6 路）：识别、下载、转图并行推进。绝大多数时间都花在
   等网络下载 PDF 上，并发后批量处理明显更快（下载阶段实测约 5 倍、识别约 3 倍）。
10. 单张 PDF 下载失败会**自动重试**（最多 2 次、逐次退避）：偶发网络抖动不会再被
    误判成「未下载」而需要人工重跑。
11. 处理过程中可点「停止」随时中止：尚未开始的图片会被直接跳过，已下载的 PDF 全部保留，
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
4. 点「▶ 开始处理」（主界面整行的大按钮），进度卡与日志区实时显示处理进度：
   进度条上方左侧是「正在处理 x / y 份」、右侧是「已用 mm:ss」，日志区每 3 秒
   报一次「已完成 x / y 张」，卡住也能一眼看出还在跑。
   要中止就点「开始处理」右侧的「停止」；日志卡右上角有「清空日志」。
   「使用说明 / 更新记录」「检查更新」在顶栏右侧。
5. 处理完成后，PDF 在「PDF」文件夹，转换图片在「PDF/图片」，问题图片的副本在「未识别」文件夹，
   汇总表在「PDF/发票汇总_*.xlsx」。

【自动更新】
- 软件启动后会静默检查 GitHub 上的最新版本；发现新版本时弹窗提示，点「是」即自动
  下载并安装（无需手动去网页下载）。
- 也可随时点顶栏右侧的「检查更新」手动检查。
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
- 勾选汇总             → PDF/发票汇总_YYYYMMDD_HHMMSS.xlsx（金额类列为纯数字、无千分位；
                         「是否重复」列互指重复行号；右侧「统计（剔重后）」按剔重口径统计，
                         「重复票据」块只计同票号第 2 张及以后）

【独立工具：只转 PDF】
命令行方式可只做前置转换（不识别、不下载、不汇总）：
    发票二维码工具.exe --pdf2img "<文件夹路径>"
效果：在该文件夹内新建（或清空重建）「处理后」，把 PDF 转成「原名_1.jpg…」并与原有图片
一起放进去，随后可再打开软件对该「处理后」文件夹走完整流程。

【说明】
- 二维码识别使用 zxing-cpp，对截图 / 小二维码会自动多尺度放大，比 OpenCV 自带更稳。
- 部分发票平台（如 jsczt.cn）打开后是一个展示页，软件会自动提取页面隐藏参数并提交下载接口获取 PDF。
- 默认 6 路并发。同一平台短时间内并发请求过多可能被限流，若出现较多「未下载」，
  等几分钟再次点「开始处理」即可：**已下载且完好的 PDF 会直接复用、不再重新下载**
  （中途「停止」后再点「开始处理」也是接着跑，不必从头再下一遍），未识别的原图也还在。
- 「汇总发票」这一步保持单线程：它用 pdfplumber 解析 PDF，属纯 Python 计算，
  并发实测没有收益（1.0×）。
- 单文件 EXE，无需安装，双击即用。
"""

CHANGELOG_TEXT = """发票二维码识别下载工具 · 更新记录
================================

2026-09-19  v5.0.0
- **全新界面（卡片化）**：整体换成「浅灰底 + 圆角白卡 + 蓝色主色」的现代布局，
  按钮、复选框、进度条、链接全部自绘，风格统一；并按 DPI 感知缩放，高分屏
  （125% / 150% / 200%）下文字不再发虚，子窗口尺寸也跟着缩放。
  主界面自上而下为「目标文件夹 → 处理选项 → 开始处理/停止 → 进度 → 处理日志」，
  「使用说明 / 更新记录」与「检查更新」移到顶栏右侧。
- **进度区更直观**：进度条上方左侧是「正在处理 x / y 份」，右侧新增「已用 mm:ss」计时；
  中途停止时进度条停在真实进度并转为琥珀色（不再跳满 100%，避免误以为整批跑完）；
  文件夹里没有可处理的图片时提示「没有需要处理的图片」。
- **日志分级着色**：✓ 绿色、⚠ 琥珀、✗ 红色、=== 蓝色，长任务里一眼扫出问题行。
- 界面上不再显示「6 路并发」等内部字样（并发能力不变，仍为默认 6 路）。
- **修复：勾选「处理完成后汇总发票」后程序必然报错** —— 生成汇总表文件名时把 datetime
  模块当成类调用，走到「出文件名」那一步就抛异常，导致汇总功能实际不可用。
  该问题自 v4.10.0 引入，勾选汇总的用户请务必更新到本版。
- **修复：关闭窗口后偶发「invalid command name」报错** —— 日志轮询、界面刷新、耗时计时
  三个内部定时器在关窗时未取消，窗口销毁后仍会被唤起。现在关闭时统一清理，
  无论走关闭按钮还是其他方式退出都不会再出现。
- 修复：使用说明窗口、更新进度框在高分屏上尺寸偏小（未按缩放换算，被压成窄条）。
- 说明：本次为主要版本更新（界面整体重构），使用方式与输出文件规则与 v4.10.0 完全一致，
  直接覆盖安装即可，无需迁移任何数据。

2026-09-18  v4.10.0
- **汇总表新增「序号」列**（放在最左侧）：1、2、3… 与明细行一一对应，配合「文件名」
  一眼看出是第几份票据。底部「合计」行写在「文件名」列，序号列留空。
- **「是否重复」列改为互指序号**：如「与序号3、5重复」，编号与最左「序号」列一致
  （v4.9.0 用的是 Excel 行号，与新的序号列差 2，容易找错行，故一并对齐）。
- **修复：重跑会全量重新下载** —— 使用说明里一直写着「已下载的会自动跳过同名文件」，
  但代码里其实每次都会重下一遍。现在真的会复用：文件存在、且确认是完好的 PDF（校验过
  文件头）才跳过；已经转好的图片也不再重复渲染。结束统计里会多一行「已存在，跳过下载」。
  中途点过「停止」再点「开始处理」，就能接着处理剩下的，不必再等一遍下载。
- **修复：处理任务遇上异常会让界面永久卡在「处理中」** —— 文件夹不可读、磁盘满、
  网络盘掉线等异常原先会让工作线程静默死亡：不复位按钮、不报错，只能重启软件。
  现在无论发生什么都会正常收尾，并在日志里说明原因。
- **修复：检查更新存在跨线程操作界面的隐患** —— 检查更新跑在后台线程里，却在其中直接
  操作窗口（Tkinter 只允许主线程调用），偶发卡死。现改为「后台线程只取版本号，
  界面动作统一交回主线程执行」，并顺带避免「检查途中用户已开始处理任务还弹更新框」。
- **加固更新下载**：校验实际下载字节数与 Content-Length 一致、且文件头为 MZ
  （可执行文件标志），避免把半截文件或拦截页当成新版本装上去；更新线程加兜底，
  出错也能正常收尾、不会留下关不掉的进度框。
- 修复：个别票据下载失败时未关闭响应，连接不回连接池，并发重试几轮后请求明显变慢。
- 修复：文件夹名带方括号（如「2026 发票[1]」）时，汇总会静默跳过（通配符把 [ ] 当字符类）。
- **修复：点「停止」后进度条会跳满 100%** —— 明明只跑了一半，看起来像整批都处理完了。
  现在停在真正处理到的位置，并用琥珀色与「已完成」的蓝色区分开；文件夹里没有可处理的图片时，
  进度条也不再显示成跑满，而是显示「没有需要处理的图片」。
- **大批量处理更省内存**：二维码识别对小图会逐级放大重试，而 4000×3000 的照片再放大 2 倍，
  单张位图就要 144MB，6 路并发叠起来接近 1GB。现在按「放大后位图的像素预算」决定放大到几倍
  —— 小图仍放大到 4 倍（识别能力不变），大图不再做无谓放大，单张占用封顶约 36MB；
  超大扫描件则先等比降采样再识别，内存峰值不再随图片尺寸失控。
- 内部重构（不影响功能与输出）：把发票汇总、网络下载、自动更新三块逻辑从主程序里拆成
  独立模块（iqr_summary / iqr_net / iqr_update），主程序只留界面与流程编排；
  重复的「响应关闭」与「重名去重」逻辑各合并为一处；发票字段查找改为按页缓存，
  同一页文本不再被 8 个字段各扫一遍（规范化次数约降为原来的 1/8）。
- 清理：移除未被调用的死代码与重复导入。

2026-09-18  v4.9.0
- **汇总表右侧统计区重构为左右两块**（原先是一竖列到底）：
  · 左块「统计（剔重后）」：票据张数 / 合计总金额 / 合计总统筹金额 /〔合计大病保险支付〕/
    〔合计医疗救助支付〕/ 可赔付金额 —— **全部改为“剔除重复票据”之后的口径**，
    即同一票号只算第 1 张，重复份不再参与这一块的计算。
  · 右块「重复票据」：新增**重复票据张数**，其后为重复票据金额合计 / 重复票据统筹合计 /
    〔重复票据大病保险支付合计〕/〔重复票据医疗救助支付合计〕—— 只统计多出来的份
    （同一票号第 2 张及以后），首张不重复计入。
  · 两块互补：左块张数 + 右块张数 = 明细行数；左块金额 + 右块金额 = 明细合计行金额。
  · 明细底部的「合计」行照旧对**全部**明细行求和（不剔重），方便与原始票据逐张核对。
- **「是否重复」列改为互指行号**：不再只标一个「是」，而是写成「与第3、5行重复」，
  同一票号各有几张、分别落在哪几行，一眼可见。
- **修复（重要）：票据号码超过 15 位时，重复判定会整体失效** —— 影响 v4.4 ~ v4.8 的
  「重复票据…合计」以及所有按票号去重的统计。原因是原先用 `COUNTIF(票号累计区间, 票号)`
  交给 Excel 判断重复，而 Excel / WPS / 在线表格的 COUNTIF / SUMIF 会把「长得像数字的文本」
  内部转成数字再比较，双精度浮点只能存 15~16 位有效数字 —— 19 位票号（如 3205002026000120032）
  会与同前缀的票号全部塌陷成同一个值，被判成同一张票。实测 16 行票据里有 15 行被误判为重复，
  「票据张数」由 13 错算成 1。现在重复判定全部改在**软件内部**完成（与「是否重复」列同一套逻辑），
  汇总表里只保留判定结果，不再依赖表格软件的票号比较，Excel / WPS / 在线预览结果一致。
- 顺带修复：汇总表的公式单元格补写**计算结果缓存值**。原先用腾讯文档等「不重算公式」的在线预览
  打开时，统计栏会显示 0 或空白（Excel / WPS 会自动重算，所以本地看不出问题）；
  现在打开即显示正确数值，公式本身仍完整保留、可照常重算。

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
        # ⚠️ 必须显式声明 argtypes / restype：ctypes 默认把返回值当 c_int，
        #    64 位下窗口句柄一旦超过 2^31 就被截断成无效句柄，图标会静默设置失败。
        user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
        user32.GetAncestor.restype = wintypes.HWND
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


def _place_child_window(win, w: int, h: int) -> None:
    """给 Toplevel 设定尺寸并居中。

    ⚠️ `geometry()` 的数字是**物理像素**，而开启 DPI 感知后 Tk 的字号是按真实 DPI
    放大的 —— 直接写逻辑尺寸会让子窗口又小又挤（内容被压成窄条）。所以这里统一
    过一层 K.u()，再钳到屏幕内、居中到屏幕，避免小屏放不下。
    """
    W, H = K.u(w), K.u(h)
    sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
    W, H = min(W, max(320, sw - K.u(20))), min(H, max(240, sh - K.u(60)))
    x, y = max(0, (sw - W) // 2), max(0, (sh - H) // 3)
    win.geometry("%dx%d+%d+%d" % (W, H, x, y))


class UpdateProgressDialog:
    """更新进度框：展示阶段、进度条、下载速度、已下载大小与详细日志。

    支持中途取消：下载阶段点「取消更新」（或点窗口 ✕）即中断下载并清理临时文件，
    当前版本不受影响；进入「保存新版本」的临界区后不可取消（约 1~2 秒）。
    """

    def __init__(self, parent, on_success=None):
        self.parent = parent
        self.on_success = on_success     # 更新成功、切换新版本前由主线程调用
        self.queue = queue.Queue()
        self._closed = False
        self.cancel = threading.Event()   # 置位 = 用户请求取消
        self._locked = False              # True = 已进入不可取消的落盘阶段
        self._cancelling = False          # 已点过取消，正在等下载线程收尾
        self.win = tk.Toplevel(parent)
        self.win.title("更新进度")
        _place_child_window(self.win, 500, 430)
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
        sk = K.SKIN
        K.style_ttk(sk)
        self.win.configure(bg=sk.bg)
        body = tk.Frame(self.win, bg=sk.bg)
        body.pack(fill=tk.BOTH, expand=True, padx=K.u(12), pady=K.u(10))

        self.stage_var = tk.StringVar(value="准备中…")
        self.pct_var = tk.StringVar(value="0%")
        self.size_var = tk.StringVar(value="0.0 / 0.0 MB")
        self.speed_var = tk.StringVar(value="— KB/s")

        # 进度卡片：阶段文案（左）+ 百分比（右）+ 进度条 + 大小/速度
        card = K.Card(body, sk)
        card.pack(fill=tk.X, pady=(0, K.u(sk.card_gap)))
        head = tk.Frame(card.body, bg=sk.card)
        head.pack(fill=tk.X)
        tk.Label(head, textvariable=self.stage_var, bg=sk.card, fg=sk.text,
                 font=K.f(sk.fs_body, True)).pack(side=tk.LEFT)
        tk.Label(head, textvariable=self.pct_var, bg=sk.card, fg=sk.accent_d,
                 font=K.f(sk.fs_body, True)).pack(side=tk.RIGHT)
        self.bar = K.RoundProgress(card.body, sk, height=sk.bar_h)
        self.bar.pack(fill=tk.X, pady=(K.u(8), 0))
        stats = tk.Frame(card.body, bg=sk.card)
        stats.pack(fill=tk.X, pady=(K.u(6), 0))
        tk.Label(stats, textvariable=self.size_var, bg=sk.card, fg=sk.muted,
                 font=K.f(sk.fs_small)).pack(side=tk.LEFT)
        tk.Label(stats, textvariable=self.speed_var, bg=sk.card, fg=sk.muted,
                 font=K.f(sk.fs_small)).pack(side=tk.RIGHT)

        # 按钮条先 pack：它固定在底部，日志卡片随后吃掉剩余高度（顺序反了按钮被挤没）
        btn_row = tk.Frame(body, bg=sk.bg)
        btn_row.pack(side=tk.BOTTOM, fill=tk.X, pady=(K.u(10), 0))
        self.btn_cancel = K.RoundButton(btn_row, sk, "取消更新", self.request_cancel,
                                        kind="danger", width=104, height=34,
                                        font_size=sk.fs_body)
        self.btn_cancel.pack(side=tk.LEFT)
        self.btn_close = K.RoundButton(btn_row, sk, "关闭", self._close, kind="ghost",
                                       width=104, height=34, font_size=sk.fs_body)
        self.btn_close.pack(side=tk.LEFT, padx=(K.u(8), 0))
        self.btn_close.configure_state("disabled")

        # 详细进度日志（占满剩余高度）
        log_card = K.Card(body, sk, auto=False, fill=True)
        log_card.pack(fill=tk.BOTH, expand=True)
        tk.Label(log_card.body, text="详细进度", bg=sk.card, fg=sk.muted,
                 font=K.f(sk.fs_body, True)).pack(anchor="w")
        tk.Frame(log_card.body, bg=sk.border, height=1).pack(fill=tk.X, pady=K.u(8))
        self.txt = K.LogView(log_card.body, sk, max_lines=400)
        self.txt.pack(fill=tk.BOTH, expand=True)

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
        if not self.btn_cancel.is_enabled():
            return
        self._cancelling = True
        self.cancel.set()
        # 按钮置灰并提示，避免重复点击；真正的收尾由工作线程回报后完成
        self.btn_cancel.configure_state("disabled", "正在取消…")
        self.stage_var.set("正在取消更新…")
        self._append("收到取消请求，正在断开下载并清理临时文件…")

    def _on_window_close(self):
        """点 ✕：能取消就取消（不关窗，等收尾），否则按状态处理。"""
        if self.btn_close.is_enabled():
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
        self._poll_job = None        # _poll 的 after 句柄（关闭时必须取消）
        self._locked = True          # 关闭一切取消入口
        self.stage_var.set("已取消更新")
        try:
            self.btn_cancel.pack_forget()   # 取消入口消失，只留「关闭」
        except Exception:
            pass
        self.btn_close.configure_state("normal")

    def _close(self):
        # 下载还没断开时，不允许直接关窗（先取消、等收尾）
        if self._cancelling and not self._locked:
            self._append("正在取消，请稍候…")
            return
        if self._closed:
            return
        self._closed = True
        if self._poll_job is not None:
            try:
                self.win.after_cancel(self._poll_job)
            except Exception:
                pass
            self._poll_job = None
        try:
            self.win.destroy()
        except Exception:
            pass

    def _switch(self):
        """更新成功：关掉进度框，再把「切换到新版本」交给调用方（主线程内执行）。"""
        self._close()
        if self.on_success:
            try:
                self.on_success()
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
            self._poll_job = self.win.after(80, self._poll)

    def _apply(self, ev):
        t = ev.get("type")
        if t == "stage":
            self.stage_var.set(ev.get("text", ""))
            self._append(ev.get("text", ""))
        elif t == "progress":
            w, tot = ev.get("written", 0), ev.get("total", 0)
            pct = (w / tot * 100) if tot else 0
            self.bar.set_value(pct)
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
            self.btn_cancel.configure_state("disabled", "保存中…")
        elif t == "cancelled":
            self._finish_cancelled()
        elif t == "done":
            ok = ev.get("ok", False)
            cancelled = self.cancel.is_set()
            self._locked = True          # 收尾阶段不再允许取消
            if ok:
                self.stage_var.set("更新完成，正在切换到新版本…")
                self._append("更新完成，即将切换到新版本。")
                if self.on_success is not None:
                    # 先让进度框显示约 1 秒再切换。⚠️ 必须由**主线程**排定（Tkinter 只能在
                    # 主线程调用）：原先是在下载线程里直接 root.after(...)，属跨线程碰 Tk。
                    try:
                        self.win.after(1000, self._switch)
                    except tk.TclError:
                        pass
            elif cancelled:
                self._finish_cancelled()
            else:
                self.stage_var.set("更新失败")
                self._append("更新失败，请重试或手动更新。")
            self.btn_cancel.configure_state("disabled")
            self.btn_close.configure_state("normal")

    def _append(self, text):
        # 更新日志是纯文本，不做分级着色（tag 显式给 None）
        self.txt.append(text, None)


def _start_update_flow(root: tk.Tk, download_url: str, version: tuple):
    """在进度框中执行更新；成功后短暂展示“更新完成”再关闭主程序，由新版本接管。

    取消 / 失败时主程序**照常继续运行**（不 destroy），进度框停在可关闭状态。
    """
    # 「更新成功后退出主程序」由进度框在主线程里回调完成（见 UpdateProgressDialog._switch），
    # 不要在下载线程里直接碰 Tk。
    dlg = UpdateProgressDialog(root, on_success=root.destroy)

    def _worker():
        def on_event(ev):
            dlg.emit(**ev)          # emit 只入队，由主线程 _poll 消费
        try:
            perform_update(download_url, version, on_event=on_event, cancel=dlg.cancel)
        except Exception as e:
            # 兜底：工作线程异常退出就再没人报 done，进度框会卡在「不可关闭」状态
            dlg.emit({"type": "detail", "text": f"更新过程出错：{e}"})
            dlg.emit({"type": "done", "ok": False})
        # 成功 / 取消 / 失败：界面收尾一律由主线程根据事件完成

    threading.Thread(target=_worker, daemon=True).start()


# 更新检查的互斥守卫：避免「启动静默检查」与手动点「检查更新」同时弹出两个对话框
_update_dialog_open = threading.Event()
# 处理任务是否正在进行（进行中就不再弹更新提示，免得打断用户）
_PROCESSING = threading.Event()

# 后台线程 → 主线程的回调队列（见 _ui_post / InvoiceQrToolApp._pump_ui）
_UI_QUEUE: "queue.Queue" = queue.Queue()


def _ui_post(root, fn):
    """把回调排到**主线程**执行（后台线程调用，本身不碰 Tk）。

    ⚠️ Tkinter 只能在创建它的线程（主线程）里调用。更新检查跑在后台线程，绝不能直接
    ``root.after`` / ``messagebox`` —— 那属于跨线程操作 Tk，可能偶发卡死或崩溃。
    """
    _UI_QUEUE.put(fn)


def _begin_update_check() -> bool:
    """尝试占用更新检查权。返回 False 表示已有一次检查在进行中，调用方应直接返回。"""
    if _update_dialog_open.is_set():
        return False
    _update_dialog_open.set()
    return True


def _end_update_check():
    _update_dialog_open.clear()


def _handle_update_result(root: tk.Tk, result, manual: bool):
    """**主线程**：按检查结果显示提示 / 询问是否更新，最后释放检查锁。"""
    try:
        if not result:
            if manual:
                messagebox.showinfo(
                    "检查更新", "暂时无法连接到更新服务器（或当前已是最新）。")
            return
        version, download_url, notes = result
        if version <= _parse_version(__VERSION__):
            if manual:
                messagebox.showinfo("检查更新", f"当前已是最新版本 v{__VERSION__}。")
            return
        if not manual and _PROCESSING.is_set():
            # 等网络返回这段时间里用户可能已经点了「开始处理」，那就别打断他了
            return
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


def _check_update_async(root: tk.Tk, manual: bool = False):
    """后台线程：只负责「取版本号」，界面动作交回主线程。

    可从任意线程调用（内部自己起线程），手动 / 静默两种模式共用一个实现。
    """
    if manual:
        if not _begin_update_check():
            return
    elif _PROCESSING.is_set() or not _begin_update_check():
        return

    def _worker():
        try:
            result = get_latest_release()
        except Exception:
            result = None
        _ui_post(root, lambda: _handle_update_result(root, result, manual))

    threading.Thread(target=_worker, daemon=True).start()


def check_and_prompt_update(root: tk.Tk):
    """检查更新，若有新版本则弹窗询问是否更新。供 UI 按钮调用。"""
    _check_update_async(root, manual=True)


def _silent_startup_check(root: tk.Tk):
    """启动后静默检查更新；发现新版本且用户确认则更新。"""
    _check_update_async(root, manual=False)


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

    h, w = img.shape[:2]
    # ① 原图本身就超预算时先等比降采样。
    # zxing 对小二维码靠「放大重试」，放大后的位图是实体内存（宽×高×3 字节），
    # 而超大扫描件/照片单张原图就有几百 MB；缩到预算内再识别，二维码本身仍然够大。
    if h * w > QR_MAX_SOURCE_PIXELS:
        k = (QR_MAX_SOURCE_PIXELS / float(h * w)) ** 0.5
        img = cv2.resize(
            img, (max(1, int(w * k)), max(1, int(h * k))),
            interpolation=cv2.INTER_AREA,
        )
        h, w = img.shape[:2]

    codes = []

    # 1. zxing-cpp 识别能力更强，先尝试；对小二维码会自动多尺度放大重试
    if zxingcpp is not None:
        # ② 放大到几倍，由「放大后位图的像素预算」推导，而不是只看原图长边：
        #    小图照样能放大到 4×，大图则不再无谓放大，单张占用有上限。
        scales = [s for s in (1, 2, 3, 4)
                  if h * w * s * s <= QR_MAX_SCALE_PIXELS] or [1]

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
    """若文本包含 URL，返回提取到的完整 URL（已削掉尾随中文标点）；否则返回 None"""
    match = URL_RE.search(text)
    if not match:
        return None
    url = match.group(0).rstrip(URL_TAIL_JUNK)
    return url or None


class _TaskCtx:
    """并发任务共享的**只读**上下文。刻意不含任何会被写入的共享状态，
    这样多个线程并发处理时就不存在统计竞态。"""

    __slots__ = ("folder", "pdf_dir", "img_dir", "convert_pdf", "cancel")

    def __init__(self, folder, pdf_dir, img_dir, convert_pdf, cancel):
        self.folder = folder
        self.pdf_dir = pdf_dir
        self.img_dir = img_dir
        self.convert_pdf = convert_pdf
        self.cancel = cancel if cancel is not None else CancelToken()


def _next_free(used: set, first: str, alt: str | None = None, make=None) -> str:
    """在 ``used``（小写名字集合）中取一个未被占用的名字，并登记到 ``used``。

    先试 ``first``；冲突且给了 ``alt`` 就改试 ``alt``；仍然冲突则依次用 ``make(i)``
    （i 从 2 起）生成候选，直到不冲突。默认在命中候选后追加 ``_i``。

    编号放在哪儿由调用方决定，所以这个内核能同时服务两种场景：
      - ``_unique_out_bases``：编号加在末尾（``a_jpg`` → ``a_jpg_2``）；
      - ``_unique_path``：编号插在扩展名之前（``x.jpg`` → ``x_2.jpg``）。
    """
    cand = first
    if cand.lower() in used and alt is not None:
        cand = alt
    if cand.lower() in used:
        base, i = cand, 1
        while cand.lower() in used:
            i += 1
            cand = make(i) if make is not None else f"{base}_{i}"
    used.add(cand.lower())
    return cand


def _unique_out_bases(files: list) -> dict:
    """为每张图片分配唯一的输出基名（不带扩展名）。

    为什么必须做：`a.jpg` 和 `a.png` 去掉扩展名后基名都是 `a`，会指向同一个
    `PDF/a.pdf`。串行处理时是"后覆盖前"，并发时会变成**两个线程同时写同一个文件 → 文件损坏**。
    这里预先给冲突的名字追加扩展名（再冲突就加序号）来消解。
    """
    used: set = set()
    mapping = {}
    for fname in files:
        stem, ext = os.path.splitext(fname)
        tag = ext.lstrip(".").lower()
        mapping[fname] = _next_free(used, stem, alt=f"{stem}_{tag}")
    return mapping


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


_KIND_LABEL = {
    "success": "✓ 已下载 PDF",
    "skipped": "✓ 已存在（跳过下载）",
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


def _is_valid_pdf_file(path: str) -> bool:
    """路径上是否已有一份**可复用**的 PDF（存在、非空、且确实是 PDF）。

    只看「文件在不在」不够：下载中断或请求被拦截时可能留下 0 字节或一段 HTML，
    复用它会一路错到汇总阶段，还不如重下。所以顺带看一眼 PDF 文件头
    （规范允许 %PDF- 出现在前 1024 字节内，故不要求严格开头）。
    """
    try:
        if not os.path.isfile(path) or os.path.getsize(path) < 5:
            return False
        with open(path, "rb") as fh:
            return b"%PDF-" in fh.read(1024)
    except OSError:
        return False


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
    already = _is_valid_pdf_file(pdf_path)
    if already:
        # 这一份上次已经下过（中途「停止」后再跑、隔天补几张再跑都会碰到）：
        # 直接复用，不再请求一次。同一平台短时间内并发请求多了容易被限流，
        # 全量重下既慢又容易让本来正常的票据变成「未下载」。
        lines.append(f"  -> 已存在 PDF，跳过下载：{os.path.basename(pdf_path)}")
    else:
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

    # 4) 可选：PDF 转 JPG（这一份如果已经转出过图片，也不再重复渲染一遍）
    if ctx.convert_pdf and ctx.img_dir:
        first_img = os.path.join(ctx.img_dir, f"{out_base}_第1页.jpg")
        if already and os.path.exists(first_img):
            lines.append("  -> 图片已存在，跳过转换")
        else:
            try:
                for img_path in convert_pdf_to_images(pdf_path, ctx.img_dir, out_base):
                    lines.append(f"  -> 已生成图片：{os.path.basename(img_path)}")
            except Exception as e:
                lines.append(f"  -> PDF 转图片失败：{e}")

    return finish("skipped" if already else "success")


# =====================================================================
# PDF 前置转换（独立工具）
#   只要目标文件夹里出现 PDF，就把 PDF 逐页转成 JPG，与原有图片一起收进
#   「处理后」子文件夹，再把后续识别/下载/汇总全部放到该子文件夹里进行。
#   全为图片时完全跳过，保持原有行为不变。
# =====================================================================

def _unique_path(name: str, used: set) -> str:
    """在 used（已占用的文件名小写集合）中为 name 取一个不冲突的文件名。

    used 里存的是**完整文件名**（与 ``_unique_out_bases`` 存基名不同）：这样
    `a.jpg` 与 `a.png` 能共存，只有真正同名才让位。命中则在扩展名之前追加序号
    （如 原文件名_1.jpg 已被原图片占用 → 原文件名_1_2.jpg）。
    """
    stem, ext = os.path.splitext(name)
    return _next_free(used, name, make=lambda i: f"{stem}_{i}{ext}")


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
        dest_name = _unique_path(name, used)
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
    """处理整个文件夹（对外入口，见下方 _process_folder_impl）。

    ⚠️ 这里必须保证「无论发生什么，最后一定有 ("done",) 入队」：界面靠它复位按钮与
    状态，漏一次就永久卡在「处理中…」，只能重启软件（打包后没有控制台，异常也看不见）。
    所以把兜底放在最外层，业务逻辑全在 _process_folder_impl 里。
    """
    try:
        _process_folder_impl(
            folder, open_after, convert_pdf, summarize, log_queue,
            cancel=cancel, workers=workers, preconvert=preconvert,
        )
    except Exception as e:
        try:
            log_queue.put(("log", f"处理过程出现未预期的错误，本次任务已中止：{e}"))
            log_queue.put(("done",))
        except Exception:
            pass


def _process_folder_impl(
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
        # 第 4 位是并发路数：界面已不再展示，保留在队列里供诊断用
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
            "total": 0, "success": 0, "skipped": 0, "no_pdf": 0, "unrecognized": 0,
            "other": 0, "error": 0, "cancelled": 0,
            "workers": workers, "elapsed": 0.0,
        }))
        log_queue.put(("done",))
        return

    # 输出基名唯一化：避免 a.jpg 与 a.png 同时写同一个 a.pdf（并发下会写坏文件）
    out_bases = _unique_out_bases(files)
    ctx = _TaskCtx(folder, pdf_dir, img_dir, convert_pdf, cancel)

    log(f"开始处理：共 {total} 张图片。")
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
                log(f"  …处理中：已完成 {done}/{total} 张，剩余 {left} 张")

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
        "skipped": kind_count["skipped"],
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
    if stats["skipped"]:
        summary.insert(3, f"✓ 已存在 PDF 直接复用（未重新下载）：{stats['skipped']} 张")
    if stats["error"]:
        summary.append(f"✗ 处理出错（详见上方日志）：{stats['error']} 张")
    if stats["cancelled"]:
        summary.append(f"■ 因「停止」未处理：{stats['cancelled']} 张")
    summary.append(
        f"耗时：{elapsed:.1f} 秒（平均 {elapsed / max(1, done):.2f} 秒/张）"
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


class InvoiceQrToolApp:
    # 日志区保留的最大行数：上千张图的长任务（每张 2~4 行 + 每 3 秒心跳）会堆到几万行，
    # 滚动越来越卡、也白占内存。超过就裁掉最早的，只留最近的这些行。
    _LOG_MAX_LINES = 5000

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"发票二维码识别下载工具 v{__VERSION__}")
        K.setup_scale(self.root)          # 按实际 DPI 推导缩放比（150% → 1.5）
        self.root.configure(bg=K.SKIN.bg)
        apply_window_icon(self.root)

        self.folder_var = tk.StringVar()
        self.open_after_var = tk.BooleanVar(value=True)
        self.convert_pdf_var = tk.BooleanVar(value=False)
        self.summarize_var = tk.BooleanVar(value=False)
        # 进度卡左侧的状态文案。沿用 progress_var 这个变量名：外部脚本与自测
        # 一直是按这个名字读界面状态的。
        self.progress_var = tk.StringVar(value="就绪")
        self.elapsed_var = tk.StringVar(value="")      # 进度卡右侧「已用 mm:ss」

        self.log_queue: queue.Queue = queue.Queue()
        self.worker_thread: threading.Thread | None = None
        self.last_stats: dict | None = None
        self.cancel_token: CancelToken | None = None
        self._busy = False
        self._started_at: float | None = None   # 本次任务开始时刻（算「已用」）
        self._elapsed_job = None                # 每秒刷新的 after 句柄
        # 三个常驻轮询的 after 句柄：关窗时必须显式取消。否则 root.destroy() 之后
        # 已排队的回调还会被 Tcl 唤起，报 `invalid command name "...._poll_log"`。
        self._poll_job = None
        self._pump_job = None
        self._closing = False

        self._build_ui()
        # 兜底清理：不管是谁销毁窗口（关闭按钮 / 外部脚本 / 自测），
        # 都要把常驻轮询的 after 任务取消掉，否则 Tcl 事后会唤起已失效的回调。
        self.root.bind("<Destroy>", self._on_root_destroy, add="+")
        self._poll_log()
        self._pump_ui()
        # 关闭窗口时若任务还在跑，先让用户确认（避免误关导致半途而废）
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        # 启动后静默检查更新（仅发现新版本时弹窗，无网络/无更新时不打扰）
        threading.Thread(
            target=_silent_startup_check, args=(self.root,), daemon=True
        ).start()

    # ───────────────────────── 搭界面 ─────────────────────────
    # 外观是「卡片化」：浅灰底 + 圆角白卡 + 蓝色主色。所有配色与尺寸都取自
    # ui_kit.SKIN，改外观只改 ui_kit 里那一处，这里一行都不用动。
    def _build_ui(self):
        sk = K.SKIN
        K.style_ttk(sk)
        self._build_topbar()
        body = tk.Frame(self.root, bg=sk.bg)
        body.pack(fill=tk.BOTH, expand=True, padx=K.u(sk.page_x),
                  pady=(K.u(sk.page_top), K.u(sk.page_bottom)))
        self._build_picker(body)
        self._build_options(body)
        self._build_actions(body)
        self._build_progress_card(body)
        self._build_log_card(body)
        self._apply_geometry()

    def _apply_geometry(self):
        """设定窗口尺寸：默认用 SKIN.win，但**钳到可用屏幕内**。

        小屏笔记本（1366×768 @125%，逻辑高只有 614）装不下 690；不钳的话
        窗口底部会顶出屏幕、日志区整块看不到。
        """
        sk = K.SKIN
        avail_w = int(self.root.winfo_screenwidth() / max(1.0, K.SCALE)) - 60
        avail_h = int(self.root.winfo_screenheight() / max(1.0, K.SCALE)) - 90
        w, h = sk.win
        w, h = max(320, min(w, avail_w)), max(240, min(h, avail_h))
        mw, mh = sk.win_min
        self.root.minsize(K.u(min(mw, w)), K.u(min(mh, h)))
        self.root.geometry("%dx%d" % (K.u(w), K.u(h)))

    def _build_topbar(self):
        """顶栏：应用图标 + 名称 + 版本徽标，右侧「使用说明 / 检查更新」。"""
        sk = K.SKIN
        bar = tk.Frame(self.root, bg=sk.card, height=K.u(sk.topbar_h))
        bar.pack(fill=tk.X)
        bar.pack_propagate(False)
        logo = tk.Canvas(bar, bg=sk.card, highlightthickness=0, bd=0,
                         width=K.u(sk.logo_s), height=K.u(sk.logo_s))
        K.logo_mark(logo, K.u(1), K.u(1), K.u(sk.logo_s) - K.u(2), sk)
        logo.pack(side=tk.LEFT, padx=(K.u(sk.page_x + 2), K.u(10)))
        tk.Label(bar, text="发票二维码识别下载工具", bg=sk.card, fg=sk.text,
                 font=K.f(sk.fs_title, True)).pack(side=tk.LEFT)
        tk.Label(bar, text=f"v{__VERSION__}", bg=sk.accent_l, fg=sk.accent_d,
                 font=K.f(sk.fs_small - 1, True), padx=K.u(8),
                 pady=K.u(2)).pack(side=tk.LEFT, padx=K.u(10))
        K.make_link(bar, sk, "检查更新", self._check_update).pack(
            side=tk.RIGHT, padx=(K.u(6), K.u(14)))
        K.make_link(bar, sk, "使用说明", self._show_help).pack(side=tk.RIGHT)
        tk.Frame(self.root, bg=sk.border, height=1).pack(fill=tk.X)

    def _build_picker(self, parent):
        """目标文件夹卡片：说明标签 + 输入框 + 「浏览…」。"""
        sk = K.SKIN
        card = K.Card(parent, sk)
        card.pack(fill=tk.X, pady=(0, K.u(sk.card_gap)))
        tk.Label(card.body, text="目标文件夹", bg=sk.card, fg=sk.muted,
                 font=K.f(sk.fs_body, True)).pack(anchor="w")
        row = tk.Frame(card.body, bg=sk.card)
        row.pack(fill=tk.X, pady=(K.u(6), 0))
        self.ent_path = ttk.Entry(row, style="P.TEntry", font=K.f(sk.fs_body),
                                  textvariable=self.folder_var)
        self.ent_path.pack(side=tk.LEFT, fill=tk.X, expand=True)
        # ⚠️ 先让输入框量出自己的高度，「浏览…」再对齐它：ttk.Entry 的高度是
        #    「上下内边距 + 字体行高」，硬编码一个数字永远差几像素。
        self.ent_path.update_idletasks()
        self.btn_browse = K.RoundButton(
            row, sk, "浏览…", self._browse_folder, kind="ghost",
            width=max(76, int(sk.btn_h * 2.2)),
            height=self.ent_path.winfo_reqheight() / max(1.0, K.SCALE),
            font_size=sk.fs_body)
        self.btn_browse.pack(side=tk.LEFT, padx=(K.u(8), 0))

    def _build_options(self, parent):
        """处理选项卡片：三个自绘复选框（两个「处理完成后…」并排，转图单独一行）。"""
        sk = K.SKIN
        card = K.Card(parent, sk)
        card.pack(fill=tk.X, pady=(0, K.u(sk.card_gap)))
        tk.Label(card.body, text="处理选项", bg=sk.card, fg=sk.muted,
                 font=K.f(sk.fs_body, True)).pack(anchor="w", pady=(0, K.u(4)))
        grid = tk.Frame(card.body, bg=sk.card)
        grid.pack(fill=tk.X)
        K.RoundCheck(grid, sk, "处理完成后打开文件夹",
                     self.open_after_var).grid(row=0, column=0, sticky="w",
                                               pady=K.u(2))
        K.RoundCheck(grid, sk, "处理完成后汇总发票（Excel）",
                     self.summarize_var).grid(row=0, column=1, sticky="w",
                                              padx=(K.u(40), 0), pady=K.u(2))
        K.RoundCheck(grid, sk, "将下载的 PDF 转换为图片（JPG，长边 2000px）",
                     self.convert_pdf_var).grid(row=1, column=0, columnspan=2,
                                                sticky="w", pady=K.u(2))

    def _build_actions(self, parent):
        """操作区：整行蓝色主按钮「▶ 开始处理」+ 右侧「停止」。"""
        sk = K.SKIN
        row = tk.Frame(parent, bg=sk.bg)
        row.pack(fill=tk.X, pady=(0, K.u(sk.card_gap)))
        self.btn_start = K.RoundButton(row, sk, "开始处理", self._start_processing,
                                       kind="primary", font_size=sk.fs_btn,
                                       icon="▶", height=sk.btn_h)
        self.btn_start.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.btn_stop = K.RoundButton(row, sk, "停止", self._stop_processing,
                                      kind="danger", width=int(sk.btn_h * 2.6),
                                      font_size=sk.fs_btn2, height=sk.btn_h)
        self.btn_stop.pack(side=tk.LEFT, padx=(K.u(sk.card_gap), 0))
        self.btn_stop.configure_state("disabled", "停止")

    def _build_progress_card(self, parent):
        """进度卡片：左侧状态文案 + 右侧「已用 mm:ss」+ 圆角进度条。"""
        sk = K.SKIN
        card = K.Card(parent, sk)
        card.pack(fill=tk.X, pady=(0, K.u(sk.card_gap)))
        head = tk.Frame(card.body, bg=sk.card)
        head.pack(fill=tk.X)
        tk.Label(head, textvariable=self.progress_var, bg=sk.card, fg=sk.text,
                 font=K.f(sk.fs_body, True)).pack(side=tk.LEFT)
        tk.Label(head, textvariable=self.elapsed_var, bg=sk.card, fg=sk.muted,
                 font=K.f(sk.fs_small)).pack(side=tk.RIGHT)
        self.progress = K.RoundProgress(card.body, sk, height=sk.bar_h)
        self.progress.pack(fill=tk.X, pady=(K.u(8), K.u(2)))

    def _build_log_card(self, parent):
        """日志卡片：占满剩余高度，标题行右侧放「清空」。"""
        sk = K.SKIN
        card = K.Card(parent, sk, auto=False, fill=True)
        card.pack(fill=tk.BOTH, expand=True)
        head = tk.Frame(card.body, bg=sk.card)
        head.pack(fill=tk.X)
        tk.Label(head, text="处理日志", bg=sk.card, fg=sk.text,
                 font=K.f(sk.fs_body, True)).pack(side=tk.LEFT)
        K.RoundButton(head, sk, "清空", self._clear_log, kind="ghost",
                      width=64, height=26, font_size=sk.fs_small).pack(side=tk.RIGHT)
        tk.Frame(card.body, bg=sk.border, height=1).pack(fill=tk.X, pady=K.u(8))
        self.log_view = K.LogView(card.body, sk, max_lines=self._LOG_MAX_LINES)
        self.log_view.pack(fill=tk.BOTH, expand=True)

    # ───────────────────────── 状态与计时 ─────────────────────────
    def _elapsed_text(self) -> str:
        if self._started_at is None:
            return self.elapsed_var.get()
        sec = max(0, int(time.time() - self._started_at))
        return f"已用 {sec // 60:02d}:{sec % 60:02d}"

    def _tick_elapsed(self):
        """每秒刷新右侧「已用 mm:ss」；任务结束后自动停表。"""
        if not self._busy or self._started_at is None:
            self._elapsed_job = None
            return
        self.elapsed_var.set(self._elapsed_text())
        try:
            self._elapsed_job = self.root.after(1000, self._tick_elapsed)
        except tk.TclError:
            self._elapsed_job = None

    def _stop_polling(self):
        """取消三个常驻 after 任务（日志轮询 / UI 泵 / 计时器）。

        句柄在各自的重排处保存，这里统一取消并置 None —— 幂等，可重复调用。
        """
        for name in ("_poll_job", "_pump_job", "_elapsed_job"):
            job = getattr(self, name, None)
            if job is not None:
                try:
                    self.root.after_cancel(job)
                except Exception:
                    pass            # 窗口已销毁时 Tcl 会拒绝，忽略即可
                setattr(self, name, None)

    def _on_root_destroy(self, event=None):
        """<Destroy> 兜底：只在 root 自己被销毁时清理（子控件销毁也会触发本事件）。"""
        if event is not None and getattr(event, "widget", None) is not self.root:
            return
        self._closing = True
        self._stop_polling()

    def _start_elapsed_timer(self):
        self._started_at = time.time()
        self.elapsed_var.set("已用 00:00")
        self._tick_elapsed()

    def _stop_elapsed_timer(self):
        if self._elapsed_job is not None:
            try:
                self.root.after_cancel(self._elapsed_job)
            except Exception:
                pass
            self._elapsed_job = None
        self._started_at = None

    def _browse_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folder_var.set(path)

    def _check_update(self):
        if self._busy:
            self._log("处理任务进行中，暂不检查更新（避免打断当前任务）。")
            return
        self._log("正在检查更新…")
        # 网络请求与弹窗都在内部处理（主线程调用，内部自行起后台线程）
        check_and_prompt_update(self.root)

    def _show_help(self):
        sk = K.SKIN
        win = tk.Toplevel(self.root)
        win.title("使用说明 / 更新记录")
        _place_child_window(win, 880, 640)
        win.configure(bg=sk.bg)
        try:
            win.transient(self.root)
            win.grab_set()
        except Exception:
            pass
        apply_window_icon(win)
        # 先 pack 底部的按钮条，再让卡片吃掉剩余空间（顺序反了按钮会被挤没）
        btn_row = tk.Frame(win, bg=sk.bg)
        btn_row.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, K.u(10)))
        K.RoundButton(btn_row, sk, "关闭", win.destroy, kind="ghost",
                      width=110, height=34, font_size=sk.fs_body).pack()
        card = K.Card(win, sk, auto=False, fill=True)
        card.pack(fill=tk.BOTH, expand=True, padx=K.u(sk.page_x),
                  pady=K.u(sk.page_top))
        view = K.LogView(card.body, sk)
        view.pack(fill=tk.BOTH, expand=True)
        view.set_text(USAGE_TEXT + "\n\n" + CHANGELOG_TEXT)

    def _clear_log(self):
        """清空日志区内容（不影响正在运行的任务），便于开始下一个任务。"""
        self.log_view.clear()

    def _log(self, msg: str):
        now = time.strftime("%H:%M:%S")
        # 级别按**正文**判定（tag_for 会自动跳过 [HH:MM:SS] 前缀），
        # 这样「=== / ✓ / ⚠ / ✗」的着色规则不会因加了时间戳而失效。
        self.log_view.append(f"[{now}] {msg}", K.LogView.tag_for(msg))

    def _poll_log(self):
        if self._closing:
            return
        try:
            while True:
                item = self.log_queue.get_nowait()
                if item[0] == "log":
                    self._log(item[1])
                elif item[0] == "progress":
                    current, total = item[1], item[2]
                    # 第 4 位是并发路数：界面已不再展示，保留在队列里供诊断用
                    if total > 0:
                        self.progress.set_value((current / total) * 100)
                        self.progress_var.set(f"正在处理 {current} / {total} 份")
                    else:
                        # total == 0：文件夹里没有可处理的图片，不是「跑完了」
                        self.progress.reset()
                        self.progress_var.set("没有需要处理的图片")
                elif item[0] == "stats":
                    self.last_stats = item[1]
                elif item[0] == "done":
                    self._finish_task()
        except queue.Empty:
            pass
        except tk.TclError:
            return          # 窗口已销毁，停止轮询
        try:
            self._poll_job = self.root.after(100, self._poll_log)
        except tk.TclError:
            pass

    def _pump_ui(self):
        """主线程消费「后台线程 → 主线程」的回调队列（见 _ui_post）。

        更新检查跑在后台线程，它拿到的结果必须回到主线程才能碰界面；这是那条通道的
        消费端。放在这里是因为它和其他轮询一样依赖主线程的事件循环。
        """
        if self._closing:
            return
        try:
            while True:
                fn = _UI_QUEUE.get_nowait()
                try:
                    fn()
                except Exception:
                    pass
        except queue.Empty:
            pass
        try:
            self._pump_job = self.root.after(60, self._pump_ui)
        except tk.TclError:
            pass

    def _finish_task(self):
        """任务收尾：复位按钮与状态，弹出统计。"""
        self._busy = False
        self.cancel_token = None
        _PROCESSING.clear()
        s = self.last_stats
        # 「停止」不等于「完成」：进度条要停在真正处理到的位置（琥珀色），
        # 而不是跳满 100% —— 否则用户会以为整批都跑完了。
        stopped = bool(s and s.get("cancelled"))
        if stopped:
            total = s.get("total") or 0
            done = max(0, total - (s.get("cancelled") or 0))
            self.progress.set_value(
                (done / total * 100) if total else 0.0, stopped=True
            )
            self.progress_var.set(f"已停止：{done} / {total} 份")
        elif s and not (s.get("total") or 0):
            # 文件夹里没有可处理的图片：空进度 + 明确文案，别显示成跑满
            self.progress.reset()
            self.progress_var.set("没有需要处理的图片")
        else:
            self.progress.set_value(100)
            if s:
                self.progress_var.set(f"已完成 {s['total']} / {s['total']} 份")
        # 停表：优先用工作线程实测的耗时，拿不到就用界面自己的计时
        el = (s or {}).get("elapsed") or 0.0
        if el:
            self.elapsed_var.set(
                f"已用 {int(el) // 60:02d}:{int(el) % 60:02d}")
        else:
            self.elapsed_var.set(self._elapsed_text())
        self._stop_elapsed_timer()
        self.btn_start.configure_state("normal", "开始处理")
        self.btn_stop.configure_state("disabled", "停止")
        self._log("--- 任务结束 ---")
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
        if s.get("skipped"):
            msg += f"\n已存在 PDF 直接复用（未重新下载）：{s['skipped']} 张"
        cancelled = s.get("cancelled") or 0
        if cancelled:
            msg += f"\n因「停止」未处理：{cancelled} 张"
        elapsed = s.get("elapsed") or 0.0
        if elapsed:
            processed = max(1, s["total"] - cancelled)
            msg += (
                f"\n\n耗时：{elapsed:.1f} 秒"
                f"（平均 {elapsed / processed:.2f} 秒/张）"
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

        self.btn_start.configure_state("disabled", "处理中…")
        self.btn_stop.configure_state("normal", "停止")
        self.progress.set_value(0)
        self.progress_var.set("正在准备…")
        self._start_elapsed_timer()
        self._log("=== 开始处理 ===")

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
        self.btn_stop.configure_state("disabled", "正在停止…")
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
        # 先立标志再取消句柄：标志挡住「取消瞬间又自己排了一个」的竞态，
        # 这样 destroy 之后不会再有回调被 Tcl 唤起。（<Destroy> 绑定会再兜一次）
        self._closing = True
        self._stop_polling()
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
