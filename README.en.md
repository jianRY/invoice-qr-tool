# Invoice QR Downloader (发票二维码识别下载工具)

A single-file Windows desktop tool that **batch-processes invoice / receipt images containing QR codes**: it recognizes the QR codes, downloads the corresponding PDFs by the rules, converts them to images, and (optionally) summarizes them into an Excel sheet. No Python environment required — **just double-click `发票二维码工具.exe` (Invoice QR Downloader.exe)**.

## What it does

1. Recognizes QR codes in all images under a chosen folder (powered by `zxing-cpp`, with automatic 1–4× upscaling for screenshots / small codes → high recognition rate).
2. QR code is a **URL** → automatically downloads the corresponding PDF into the `PDF/` subfolder, keeping the original image's filename.
   - Two download strategies are built in: a direct PDF link, and invoice platforms that first show a display page and require parsing `idBase` (e.g. Jiangsu fiscal e-receipts).
3. QR code is **pure digits** → ignored.
4. No QR / unrecognizable → the image is auto-renamed to `未识别-原文件名` (`Unrecognized-…`).
5. URL but **no downloadable PDF** → renamed to `未下载原文件名` (`NotDownloaded…`).
6. Same URL appears **repeatedly** → prefixed with `重复` (`Duplicate`). Can stack with `未下载`, e.g. `重复未下载原文件名`.
7. Optional toggle: **open the folder automatically** when done.
8. Optional toggle: **convert downloaded PDFs to images** (long edge 2000px, short edge auto), saved to `PDF/图片/` as `原名_第N页.png`.
9. Optional toggle: **summarize invoices into Excel** when done. Currently tuned for **Jiangsu Province medical fee receipt** templates, extracting: payer, receipt number, issue date, total amount (lowercase), and medical insurance pooled-fund payment, plus **critical-illness insurance payment** / **medical assistance payment** (conditional columns, emitted only when recognized); outputs `PDF/发票汇总_时间戳.xlsx`. The `是否重复` column cross-references row numbers (e.g. 「与第3、5行重复」); to its right sit two blocks — **「统计（剔重后）」** (receipt count / total amount / pooled-fund total /〔critical-illness total〕/〔medical assistance total〕/ compensable amount, all computed **after removing duplicate receipts**, i.e. only the 1st copy of each receipt number) and **「重复票据」** (duplicate count / duplicate amount total / per-category duplicate totals, counting only the 2nd and later copies). Duplicate detection runs **inside the tool**, not through Excel formulas: receipt numbers are 15+ digits and spreadsheet `COUNTIF`/`SUMIF` coerce numeric-looking text to numbers, truncating them to 15 significant digits and misjudging every receipt as the same one (see changelog v4.9.0).
10. A successfully recognized image whose PDF was downloaded gets a `1` prefix added to its filename (e.g. `1invoice001.jpg`), matching the other status prefixes so you can spot completed items at a glance.
11. When processing finishes, a **results summary** popup (also written to the log) reports: total images recognized, how many downloaded a PDF, how many recognized-but-no-PDF, how many pure-digit-ignored, and how many failed to recognize (duplicates counted separately).

## Usage

1. Put the invoice images you want to process into one folder.
2. Double-click `发票二维码工具.exe`.
3. Click **Browse…** to select that folder.
4. Toggle as needed:
   - Open folder when done
   - Convert downloaded PDFs to images (long edge 2000px)
   - Summarize invoices into Excel when done
5. Click **Start**, and watch the progress in the log area.
6. When finished: PDFs are in `PDF/`, converted images in `PDF/图片/`, and the summary in `PDF/发票汇总_*.xlsx`.
7. To update, click the **Check for Updates** button (or just confirm the prompt shown by the silent startup check).

The UI also has **Usage**, **Changelog**, and **Check for Updates** buttons, always accessible.

## Naming rules at a glance

| Case | Result |
| --- | --- |
| URL + download OK | `PDF/<original>.pdf`, and the image renamed `1<original>` |
| URL + no PDF available | image renamed `未下载<original>` |
| Pure-digit QR | ignored |
| No / unrecognizable QR | image renamed `未识别-<original>` |
| Duplicate URL | image prefixed `重复` |
| Combinable | e.g. `重复未下载<original>`, `重复1<original>` |

## Auto-update & rollback

- On startup the app silently checks the latest Release of the public repo. If a newer version is found, it prompts you; click **Yes** to download and replace automatically (or click **Check for Updates** anytime).
- The update source is the public repo's Release — **no account or token required**.
- **Automatic rollback on failure**: the current executable is backed up before replacement (`_backup.exe`); if the new version fails its startup self-check, the updater script automatically restores the previous version, relaunches it, and shows a **"Update rolled back"** warning dialog so you are never left with a broken or silently-failed app.
- **Update progress dialog**: a dedicated progress window shows the update stage, a progress bar, live download speed (MB/s / KB/s), downloaded size, and a detailed log; when done it replaces and relaunches automatically — no blocking popup.
- **Faster startup**: heavy dependencies (OpenCV / PyMuPDF / pdfplumber / openpyxl / zxing-cpp) are lazily imported only when their feature is actually used; the GUI and the update check no longer initialize them at launch.

### How the auto-update works

- The app reads the public repo's latest Release (`releases/latest`) and compares versions semantically.
- When a higher version is found, it downloads the asset named `InvoiceQRDownloader_X.Y.Z.exe` from that Release to a temp folder, copies it next to the current exe as `_pending.exe`, then a self-launched `.bat` script replaces the current exe after the old process exits and relaunches the new one with a `--pending` self-check flag.
- The new process confirms it started successfully, then clears the backup (update succeeds). If it never confirms, the script restores the backup (rollback).

## Download the executable

Go to the repo's **Releases** page to download the latest (the repo is public, no login needed):

- `InvoiceQRDownloader_4.7.0.exe` — main program (single file, double-click to run; the GitHub attachment uses an ASCII name because Release asset names can't be Chinese — the in-app title and local file are still the Chinese name `发票二维码工具.exe`)
- `InvoiceQR_Usage.txt` — usage instructions
- `InvoiceQR_Changelog.txt` — changelog

> Chinese versions `使用说明.txt` / `更新日志.txt` are also provided at the repo root, identical in content to the attachments above.

## File structure

```
invoice_qr_tool.py          # source (tkinter GUI + processing logic)
README.md                   # this file (English)
README.zh.md                # Chinese documentation
使用说明.txt / 更新日志.txt   # bundled Chinese usage & changelog
outputs/
  └─ 发票二维码工具.exe       # packaged artifact (single file; see GitHub Releases)
```

## Tech stack

- Python 3.13 + `tkinter` (GUI)
- `zxing-cpp` (QR recognition), `OpenCV` (image preprocessing)
- `PyMuPDF` (PDF→image), `pdfplumber` + `openpyxl` (field extraction & Excel summary)
- `requests` (download)
- `PyInstaller` packaged as `--onefile --windowed` single executable

## Build from source

Requires Windows and a local Python 3.13 (with `tkinter`):

```bash
python -m venv envs/default
envs/default/Scripts/python.exe -m pip install opencv-python zxingcpp pymupdf pdfplumber openpyxl requests pyinstaller
envs/default/Scripts/python.exe -m PyInstaller --onefile --windowed --name 发票二维码工具 \
  --hidden-import numpy --hidden-import cv2 --hidden-import zxingcpp \
  --hidden-import pymupdf --hidden-import requests --hidden-import pdfplumber \
  --hidden-import openpyxl --hidden-import pdfminer --hidden-import pdfminer.high_level \
  --hidden-import openpyxl.styles invoice_qr_tool.py
# artifact at dist/发票二维码工具.exe
```

## Notes & disclaimer

- The summary feature is currently tuned for **Jiangsu Province medical fee receipt** templates; other templates may extract incompletely (cells left blank or marked "解析失败").
- This tool is for local batch processing only. Keep images and output files containing personal receipt information secure.
- The EXE is large (~100MB+, all dependencies bundled), so it is distributed via GitHub Releases rather than committed to the repo.
