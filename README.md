# 发票二维码识别下载工具（Invoice QR Downloader）

一个 Windows 单文件桌面小工具，用于**批量处理含有二维码的发票 / 票据图片**：自动识别二维码、按规则下载对应 PDF、转图、并按需汇总成 Excel。无需安装 Python 环境，**双击 `发票二维码工具.exe` 即用**。

## 它能做什么

1. 识别指定文件夹下所有图片中的二维码（基于 `zxing-cpp`，对截图 / 小图自动 1~4 倍放大，识别率高）。
2. 二维码为**网址** → 自动下载对应 PDF，保存到 `PDF/` 子目录，文件名与原始图片同名。
   - 内置两种下载策略：直接 PDF 链接；以及需要先进入展示页再解析 `idBase` 的发票平台（如江苏财政电子票据）。
3. 二维码为**纯数字** → 忽略。
4. 未识别到二维码 / 无法识别 → 原图片自动重命名为 `未识别-原文件名`。
5. 网址但**无可下载的 PDF** → 重命名为 `未下载原文件名`。
6. 同一网址**重复**出现 → 图片名前加 `重复` 标识（可与 `未下载` 叠加，如 `重复未下载原文件名`）。
7. 可选开关：**处理完成后自动打开文件夹**。
8. 可选开关：**将下载的 PDF 转为图片**（长边 2000px，短边自适应），保存到 `PDF/图片/`，命名 `原名_第N页.png`。
9. 可选开关：**处理完成后汇总发票**（生成 Excel）。目前针对**江苏省医疗收费票据**模板，提取字段：交款人、票据号码、开票日期、金额合计（小写）、医保统筹基金支付；输出 `PDF/发票汇总_时间戳.xlsx`（含金额、医保统筹合计行）。

## 使用方法

1. 把待处理的发票图片放进同一个文件夹。
2. 双击 `发票二维码工具.exe`。
3. 点「浏览…」选择该文件夹。
4. 按需勾选：
   - 处理完成后打开文件夹
   - 将下载的 PDF 转换为图片（长边 2000px）
   - 处理完成后汇总发票（生成 Excel）
5. 点「开始处理」，在日志区查看进度。
6. 处理完成后，PDF 在 `PDF/` 文件夹，转换图片在 `PDF/图片/`，汇总表在 `PDF/发票汇总_*.xlsx`。

界面内还有「使用说明」「更新记录」两个按钮，可随时查看。

## 命名规则速查

| 情况 | 结果 |
| --- | --- |
| 网址 + 下载成功 | `PDF/<原名>.pdf` |
| 网址 + 无 PDF 可下载 | 图片改名 `未下载<原名>` |
| 纯数字二维码 | 忽略 |
| 无 / 无法识别二维码 | 图片改名 `未识别-<原名>` |
| 重复网址 | 图片加 `重复` 前缀 |
| 上述可叠加 | 例如 `重复未下载<原名>` |

## 文件结构

```
invoice_qr_tool.py          # 源码（tkinter GUI + 处理逻辑）
README.md                   # 本说明
使用说明.txt / 更新日志.txt   # 随软件附带的说明与更新记录
outputs/
  └─ 发票二维码工具.exe       # 打包产物（单文件，见 GitHub Releases 下载）
```

## 技术栈

- Python 3.13 + `tkinter`（GUI）
- `zxing-cpp`（二维码识别）、`OpenCV`（图像预处理）
- `PyMuPDF`（PDF 转图片）、`pdfplumber` + `openpyxl`（发票字段提取与 Excel 汇总）
- `requests`（下载）
- `PyInstaller` 打包为 `--onefile --windowed` 单文件可执行程序

## 从源码构建

需要 Windows 与本机 Python 3.13（含 `tkinter`）：

```bash
python -m venv envs/default
envs/default/Scripts/python.exe -m pip install opencv-python zxingcpp pymupdf pdfplumber openpyxl requests pyinstaller
envs/default/Scripts/python.exe -m PyInstaller --onefile --windowed --name 发票二维码工具 \
  --hidden-import numpy --hidden-import cv2 --hidden-import zxingcpp \
  --hidden-import pymupdf --hidden-import requests --hidden-import pdfplumber \
  --hidden-import openpyxl --hidden-import pdfminer --hidden-import pdfminer.high_level \
  --hidden-import openpyxl.styles invoice_qr_tool.py
# 产物位于 dist/发票二维码工具.exe
```

## 说明与免责

- 汇总功能目前针对**江苏省医疗收费票据**模板；其它模板可能提取不全（单元格留空或标“解析失败”）。
- 本工具仅用于本地批量处理，请妥善保管包含个人票据信息的图片与输出文件。
- EXE 体积较大（约 100MB+，含全部依赖），因此通过 GitHub Releases 分发，而非直接提交进仓库。
