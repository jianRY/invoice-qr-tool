# -*- coding: utf-8 -*-
"""发票二维码识别下载工具 · 应用图标生成器（方案 B：取景框 + 二维码）

几何绘制（PIL），不用 AI 生图：
  · 设计坐标系 1024，超采样 SS=3 后 LANCZOS 缩小 → 边缘干净、小尺寸不糊
  · 圆角方底 + 对角渐变 + 顶部柔光；主体带投影浮起
  · 配色取自 website/assets/style.css 的品牌色（--ink / --accent）

产物：
  app_icon.ico           多尺寸 ICO（16/24/32/48/64/128/256），打包与运行时使用
  app_icon_1024.png      1024 母版留档（改尺寸/改色可重生成，不必重画）
  website/assets/app-icon.png   官网 logo + favicon

用法：python make_icon.py
"""
import os
import numpy as np
from PIL import Image, ImageDraw, ImageFilter

DESIGN, SS = 1024, 3
S = DESIGN * SS
HERE = os.path.dirname(os.path.abspath(__file__))

# 品牌色（website/assets/style.css）
BG_DARK = "#0f172a"   # --ink（深色端）
BG_BLUE = "#1e40af"   # 蓝色端
INK     = "#0f172a"   # 二维码模块
CYAN    = "#22d3ee"   # 取景角标（--accent 提亮）

ICO_SIZES = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (24, 24), (16, 16)]


# ---------------------------------------------------------------- 基础工具
def sc(v):
    return int(round(v * SS))


def hex2rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def gradient(size, c1, c2):
    """对角线性渐变。"""
    y, x = np.mgrid[0:size, 0:size]
    t = (x + y) / (2.0 * size - 2.0)
    arr = np.zeros((size, size, 3), dtype=np.float32)
    for i in range(3):
        arr[:, :, i] = c1[i] * (1 - t) + c2[i] * t
    return Image.fromarray(arr.astype(np.uint8), "RGB")


def drop_shadow(mask, blur=16, alpha=0.36, dx=4, dy=14):
    sh = mask.filter(ImageFilter.GaussianBlur(sc(blur))).point(lambda v: int(v * alpha))
    out = Image.new("L", (S, S), 0)
    out.paste(sh, (sc(dx), sc(dy)))
    return out


def composite_black(img, alpha_mask):
    img.alpha_composite(Image.merge("RGBA", (
        Image.new("L", (S, S), 0), Image.new("L", (S, S), 0),
        Image.new("L", (S, S), 0), alpha_mask)))


def thick_line(md, pts, width, fill=255):
    """带圆头的粗折线。"""
    w = sc(width)
    md.line([(sc(x), sc(y)) for x, y in pts], fill=fill, width=w, joint="curve")
    r = w / 2.0
    for x, y in pts:
        md.ellipse((sc(x) - r, sc(y) - r, sc(x) + r, sc(y) + r), fill=fill)


# ------------------------------------------------- 二维码图案（9x9 模块）
def _qr_modules():
    mods = set()
    for r in range(3):                       # 三个定位角（3x3 环，中间留空）
        for c in range(3):
            if (r, c) != (1, 1):
                mods.add((r, c))
                mods.add((r, c + 6))
                mods.add((r + 6, c))
    mods |= {
        (3, 0), (4, 1), (5, 2), (0, 3), (1, 4), (2, 5),
        (3, 3), (3, 5), (4, 4), (5, 3), (5, 5),
        (3, 6), (4, 7), (5, 8), (6, 5), (7, 4), (8, 3),
        (6, 6), (7, 7), (8, 8), (6, 8), (8, 6), (7, 3),
        (3, 7), (4, 8), (2, 7), (1, 8), (7, 1), (8, 2),
        (0, 5), (5, 7), (8, 5), (6, 3),
    }
    return {m for m in mods if 0 <= m[0] <= 8 and 0 <= m[1] <= 8}


QR9 = _qr_modules()


def draw_qr(md, x, y, size, color, gap=0.12):
    """在 (x,y) 处画 size×size 的 9x9 二维码，gap 为模块间留白比例。"""
    cell = size / 9.0
    g = cell * gap / 2.0
    for (r, c) in QR9:
        md.rounded_rectangle(
            (sc(x + c * cell + g), sc(y + r * cell + g),
             sc(x + (c + 1) * cell - g), sc(y + (r + 1) * cell - g)),
            radius=sc(g * 0.9), fill=color)


# ---------------------------------------------------------------- 绘制
def build_icon():
    # 1) 圆角方底 + 对角渐变
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    plate = Image.new("L", (S, S), 0)
    ImageDraw.Draw(plate).rounded_rectangle(
        (sc(28), sc(28), sc(DESIGN - 28), sc(DESIGN - 28)), radius=sc(212), fill=255)
    img.paste(gradient(S, hex2rgb(BG_DARK), hex2rgb(BG_BLUE)).convert("RGBA"),
              (0, 0), plate)

    # 2) 顶部柔光（椭圆 + 大半径模糊，白色低透明）
    gloss = Image.new("L", (S, S), 0)
    ImageDraw.Draw(gloss).ellipse(
        (sc(-180), sc(-320), sc(DESIGN + 180), sc(470)), fill=255)
    gloss = gloss.filter(ImageFilter.GaussianBlur(sc(46))).point(lambda v: int(v * 0.26))
    gloss = Image.composite(gloss, Image.new("L", (S, S), 0), plate)
    img.alpha_composite(Image.merge("RGBA", (
        Image.new("L", (S, S), 255), Image.new("L", (S, S), 255),
        Image.new("L", (S, S), 255), gloss)))

    # 3) 中央白色圆角方块（二维码底），带投影
    sq = Image.new("L", (S, S), 0)
    ImageDraw.Draw(sq).rounded_rectangle(
        (sc(312), sc(312), sc(712), sc(712)), radius=sc(64), fill=255)
    composite_black(img, drop_shadow(sq, blur=16, alpha=0.36, dx=4, dy=14))
    white = Image.new("RGBA", (S, S), (255, 255, 255, 255))
    white.putalpha(sq)
    img.alpha_composite(white)

    # 4) 二维码 + 四角青色取景角标
    d = ImageDraw.Draw(img)
    draw_qr(d, 352, 352, 320, hex2rgb(INK) + (255,))
    for (cx, cy, sx, sy) in ((204, 204, 1, 1), (820, 204, -1, 1),
                             (204, 820, 1, -1), (820, 820, -1, -1)):
        thick_line(d, [(cx, cy), (cx + sx * 176, cy)], 62, hex2rgb(CYAN) + (255,))
        thick_line(d, [(cx, cy), (cx, cy + sy * 176)], 62, hex2rgb(CYAN) + (255,))

    return img.resize((DESIGN, DESIGN), Image.LANCZOS)


def main():
    icon = build_icon()
    ico = os.path.join(HERE, "app_icon.ico")
    master = os.path.join(HERE, "app_icon_1024.png")
    icon.save(master)
    icon.save(ico, format="ICO", sizes=ICO_SIZES)

    web = os.path.join(HERE, "website", "assets", "app-icon.png")
    if os.path.isdir(os.path.dirname(web)):
        icon.resize((512, 512), Image.LANCZOS).save(web, optimize=True)

    for p in (ico, master, web):
        if os.path.exists(p):
            print("  %-46s %8.1f KB" % (os.path.relpath(p, HERE), os.path.getsize(p) / 1024))
    print("图标生成完成（方案 B：取景框 + 二维码）")


if __name__ == "__main__":
    main()
