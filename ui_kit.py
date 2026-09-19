# -*- coding: utf-8 -*-
"""ui_kit —— Tkinter 美化控件库（卡片化方案，无第三方依赖）。

配色与形态集中在 SKIN 单例：改外观只改那一处，绘制代码不用动。
主程序 invoice_qr_tool.py 使用的就是 SKIN（= PRESETS["标准"]），
PRESETS 里另存着「紧凑 / 舒适」两档密度，供预览脚本对比或日后微调。

⚠️ 使用前提（顺序不能变）：
    在 `import tkinter` **之前** 调用 enable_dpi_awareness()，否则 150%/200% 缩放下
    Tk 逻辑像素与屏幕物理像素不一致。本模块 import 时已自动调用，
    调用方只要保证 `import ui_kit` 早于 `import tkinter` 即可。

设计要点：
  * 尺寸一律过 u()（乘 DPI 缩放）；字号用「点」，Tk 自动按 DPI 换算（且只接受整数点值）。
  * Tk 无圆角/投影 → 全部 Canvas 自绘。
"""
import re
import ctypes

# ⚠️ 立即生效（在本模块 import tkinter 之前）：任何 `import ui_kit` 都自动获得 DPI 感知。
#    调用方必须保证 ui_kit 是本程序里最早接触 tkinter 的模块（别先 import tkinter）。
def _apply_dpi_awareness():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)      # PER_MONITOR_DPI_AWARE
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()       # 老系统兜底
        except Exception:
            pass


_apply_dpi_awareness()

import tkinter as tk
from tkinter import ttk


def enable_dpi_awareness():
    """同上（公开别名）。注意：必须在 import tkinter 之前调用才有效。"""
    _apply_dpi_awareness()


# ───────────────────────── 缩放与字体 ─────────────────────────
SCALE = 1.0          # 由 setup_scale(root) 按实际 DPI 覆盖


def setup_scale(root):
    """按屏幕实际 DPI 推导缩放比（150% → 1.5）。"""
    global SCALE
    try:
        SCALE = root.winfo_fpixels("1i") / 96.0
    except Exception:
        SCALE = 1.0
    return SCALE


def u(v):
    """逻辑尺寸 → 物理像素。"""
    return int(round(v * SCALE))


FONT_UI = "Microsoft YaHei UI"


def f(size, bold=False):
    """字号用「点」，Tk 按 DPI 自动换算。⚠️ 只接受整数点值。"""
    size = int(round(size))
    return (FONT_UI, size, "bold") if bold else (FONT_UI, size)


def parent_bg(widget, fallback="#FFFFFF"):
    try:
        return widget.cget("bg")
    except Exception:
        return fallback


# ───────────────────────── 配色与形态 ─────────────────────────
class Skin:
    """配色与形态参数：所有控件都从这里取值。"""

    def __init__(self, name="卡片化", **kw):
        self.name = name
        # 色板
        self.bg = "#F4F6FA"
        self.card = "#FFFFFF"
        self.border = "#E4E8F0"
        self.text = "#1F2937"
        self.muted = "#6B7280"
        self.faint = "#9AA3AF"
        self.accent = "#2563EB"
        self.accent_d = "#1D4ED8"
        self.accent_l = "#EEF2FF"
        self.track = "#E8ECF3"
        self.ok = "#15803D"
        self.warn = "#B45309"
        self.bad = "#B91C1C"
        # 形态（圆角）
        self.radius_card = 10
        self.radius_btn = 9
        self.shadow = True
        # ── 字号（点）⚠️ Tk 只接受整数点值，非整数会被 f() 四舍五入，别写 9.5 ──
        self.fs_title = 12      # 顶栏应用名
        self.fs_body = 10       # 正文 / 卡片标题 / 输入框 / 复选框
        self.fs_small = 9       # 辅助文字（已用时间、清空、版本徽标）
        self.fs_log = 10        # 日志正文
        self.fs_btn = 12        # 主按钮
        self.fs_btn2 = 10       # 次按钮
        # ── 垂直节奏（逻辑像素；用的时候一律过 u()）──
        self.topbar_h = 52      # 顶栏高
        self.logo_s = 26        # 顶栏图标边长
        self.page_x = 16        # 页面左右边距
        self.page_top = 8       # 页面顶边距
        self.page_bottom = 12   # 页面底边距
        self.card_pad = 14      # 卡片内边距
        self.card_gap = 10      # 卡片之间 / 分组间距
        self.btn_h = 42         # 主/次按钮高
        self.check_s = 17       # 复选框边长
        self.bar_h = 9          # 进度条高
        self.entry_pad_y = 11   # 输入框上下内边距（输入框高由它决定，浏览按钮再对齐输入框）
        # ── 窗口 ──
        self.win = (900, 690)
        self.win_min = (780, 560)
        self.__dict__.update(kw)

    def palette(self):
        return {k: v for k, v in self.__dict__.items() if isinstance(v, str)}


# 唯一配色：卡片化（浅灰底 + 白卡 + 蓝色主色）。想换外观改这里一处即可。
SKIN = Skin()

# 三档密度：**窗口尺寸完全一致**，只差字号与控件/间距大小，方便横向对比。
# 换档只换 Skin，所有绘制代码零改动。
PRESETS = {
    "紧凑": Skin("紧凑",
                 fs_title=11, fs_body=10, fs_small=9, fs_log=9,
                 fs_btn=11, fs_btn2=10,
                 topbar_h=46, logo_s=24, page_x=14, page_top=6, page_bottom=10,
                 card_pad=12, card_gap=8, btn_h=36, check_s=16, bar_h=8,
                 entry_pad_y=9),
    "标准": Skin("标准"),
    "舒适": Skin("舒适",
                 fs_title=13, fs_body=11, fs_small=10, fs_log=10,
                 fs_btn=13, fs_btn2=12,
                 topbar_h=58, logo_s=30, page_x=18, page_top=10, page_bottom=14,
                 card_pad=16, card_gap=12, btn_h=48, check_s=19, bar_h=11,
                 entry_pad_y=12),
}
PRESET_ORDER = ("紧凑", "标准", "舒适")
SKIN = PRESETS["标准"]


# ───────────────────────── 绘制原语 ─────────────────────────
def rr(cv, x1, y1, x2, y2, r, fill, tags=None):
    """圆角矩形填充：4 个扇形 + 2 个矩形拼合（outline 同色，避免接缝）。"""
    tags = tags or ()
    r = max(0, r)
    kw = dict(fill=fill, outline=fill, tags=tags)
    d = 2 * r
    if r > 0:
        cv.create_arc(x1, y1, x1 + d, y1 + d, start=90, extent=90, style="pieslice", **kw)
        cv.create_arc(x2 - d, y1, x2, y1 + d, start=0, extent=90, style="pieslice", **kw)
        cv.create_arc(x1, y2 - d, x1 + d, y2, start=180, extent=90, style="pieslice", **kw)
        cv.create_arc(x2 - d, y2 - d, x2, y2, start=270, extent=90, style="pieslice", **kw)
    cv.create_rectangle(x1 + r, y1, x2 - r, y2, **kw)
    cv.create_rectangle(x1, y1 + r, x2, y2 - r, **kw)


def rr_border(cv, x1, y1, x2, y2, r, border, fill, tags=None):
    """带 1px 描边的圆角矩形 = 外圈描边色 + 内缩 1px 填充色。"""
    rr(cv, x1, y1, x2, y2, r, border, tags)
    rr(cv, x1 + 1, y1 + 1, x2 - 1, y2 - 1, max(0, r - 1), fill, tags)


def text_w(cv, text, font):
    """量文字宽度（临时画一个再删）。"""
    t = cv.create_text(-9999, -9999, text=text, font=font)
    b = cv.bbox(t)
    cv.delete(t)
    return (b[2] - b[0]) if b else 0


def shadow(cv, x1, y1, x2, y2, r, base_bg="#F4F6FA"):
    """伪投影：卡片下方叠 3 层渐淡圆角块（Tk 无真阴影）。"""
    for i, tone in enumerate(("#E7EBF3", "#ECEFF6", "#F1F4F9"), start=1):
        rr(cv, x1 + i, y1 + i + 1, x2 + i, y2 + i + 1, r, tone)


def logo_mark(cv, x, y, size, sk, fill=None, line_color="#FFFFFF"):
    """应用图标：圆角方块 + 3 条横线（模拟票据）。"""
    fill = fill or sk.accent
    rr(cv, x, y, x + size, y + size, size * 0.25, fill)
    for i, t in enumerate((0.36, 0.50, 0.64)):
        cv.create_line(x + size * 0.26, y + size * t,
                       x + size * (0.74 if i < 2 else 0.58), y + size * t,
                       fill=line_color, width=u(1.6), capstyle="round")


# ───────────────────────── 控件 ─────────────────────────
class Card(tk.Canvas):
    """圆角白卡：外部布局正常 pack/grid，内部内容塞进 self.body。

    auto=True  → 高度由内容决定（普通卡片）
    fill=True  → 高度由父容器分配（日志卡片）
    """

    def __init__(self, master, sk, radius=None, pad=None, auto=True, fill=False):
        self.sk = sk
        self.pad = u(sk.card_pad if pad is None else pad)
        self.auto, self.fill_mode = auto, fill
        self.radius = u(radius if radius is not None else sk.radius_card)
        bg = parent_bg(master, sk.bg)
        super().__init__(master, bg=bg, highlightthickness=0, bd=0,
                         height=u(40) if auto else u(60))
        self.body = tk.Frame(self, bg=sk.card)
        self._win = self.create_window(self.pad, self.pad, window=self.body,
                                       anchor="nw")
        self.body.bind("<Configure>", self._on_body)
        self.bind("<Configure>", self._on_canvas)

    def _on_body(self, e):
        if self.auto:
            want = e.height + 2 * self.pad
            if want != self.winfo_height():
                self.configure(height=want)

    def _on_canvas(self, e):
        w = max(1, e.width - 2 * self.pad)
        if self.fill_mode:
            self.itemconfigure(self._win, width=w,
                               height=max(1, e.height - 2 * self.pad))
        else:
            self.itemconfigure(self._win, width=w)
        self._paint(e.width, e.height)

    def _paint(self, w, h):
        if not self.winfo_exists():
            return
        self.delete("cardbg")
        if w < 4 or h < 4:
            return
        if self.sk.shadow:
            shadow(self, 0, 0, w - u(3), h - u(3), self.radius)
        rr_border(self, 0, 0, w - 1, h - 1, self.radius, self.sk.border,
                  self.sk.card, ("cardbg",))
        self.tag_lower("cardbg")


class RoundButton(tk.Canvas):
    """自绘圆角按钮。kind: primary | ghost | danger。"""

    def __init__(self, master, sk, text="", command=None, kind="primary",
                 height=None, width=None, font_size=12, icon=None):
        self.sk = sk
        self.text, self.command, self.kind, self.icon = text, command, kind, icon
        self.font_size = font_size
        self._enabled = True
        self._hover = False
        bg = parent_bg(master, sk.card)
        h = u(height if height is not None else sk.btn_h)
        kw = dict(height=h, bg=bg, highlightthickness=0, bd=0)
        if width:
            kw["width"] = u(width)
        super().__init__(master, **kw)
        self.configure(cursor="hand2")
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Enter>", self._enter)
        self.bind("<Leave>", self._leave)
        self.bind("<Button-1>", self._click)

    def configure_state(self, state="normal", text=None):
        """state: normal | disabled；text 给了就换文案。"""
        self._enabled = (state == "normal")
        if text is not None:
            self.text = text
        self.configure(cursor="hand2" if self._enabled else "arrow")
        self._draw()

    def is_enabled(self) -> bool:
        """当前是否可点（更新进度框据此判断按钮处于哪一态）。"""
        return self._enabled

    def _colors(self):
        sk = self.sk
        if not self._enabled:
            if self.kind == "ghost":
                return "#FBFCFE", sk.faint, sk.border
            return "#DCE4F5", "#9FB4E0", None
        if self.kind == "primary":
            return sk.accent_d if self._hover else sk.accent, "#FFFFFF", None
        if self.kind == "danger":
            return "#FEF2F2" if self._hover else sk.card, sk.bad, sk.border
        return "#F3F5F9" if self._hover else sk.card, sk.text, sk.border

    def _draw(self):
        if not self.winfo_exists():
            return
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 4 or h < 4:
            return
        fill, fg, bd = self._colors()
        if bd:
            rr_border(self, 1, 1, w - 2, h - 2, u(self.sk.radius_btn), bd, fill)
        else:
            rr(self, 1, 1, w - 2, h - 2, u(self.sk.radius_btn), fill)
        cx, cy = w / 2, h / 2
        if self.icon:
            tw = text_w(self, self.text, f(self.font_size, True))
            iw = text_w(self, self.icon, f(self.font_size - 3))
            total = iw + u(10) + tw
            x0 = cx - total / 2
            self.create_text(x0, cy, text=self.icon, font=f(self.font_size - 3),
                             fill=fg, anchor="w")
            self.create_text(x0 + iw + u(10), cy, text=self.text,
                             font=f(self.font_size, True), fill=fg, anchor="w")
        else:
            self.create_text(cx, cy, text=self.text, font=f(self.font_size, True),
                             fill=fg)

    def _enter(self, _e):
        self._hover = True
        self._draw()

    def _leave(self, _e):
        self._hover = False
        self._draw()

    def _click(self, _e):
        if self._enabled and self.command:
            self.command()


class RoundCheck(tk.Canvas):
    """自绘复选框（18px 圆角方块 + 对勾）。"""

    def __init__(self, master, sk, text="", variable=None, font_size=None):
        self.sk = sk
        self.text = text
        self.var = variable or tk.BooleanVar(value=False)
        self.font_size = font_size or sk.fs_body
        bg = parent_bg(master, sk.card)
        self._size = u(18)
        self._tw = None
        super().__init__(master, bg=bg, highlightthickness=0, bd=0,
                         height=self._size, width=u(400), cursor="hand2")
        self._var_trace = self.var.trace_add("write", self._on_var)
        # ⚠️ 界面重建会销毁控件，但 trace 还挂在共享的 BooleanVar 上 → 必须在销毁时摘掉，
        #    否则之后每次 set() 都会调已销毁控件的 _draw()，报 invalid command name
        self.bind("<Destroy>", self._on_destroy)
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Button-1>", self._toggle)

    def _on_var(self, *_):
        self._draw()

    def _on_destroy(self, e):
        if e.widget is not self:
            return
        try:
            self.var.trace_remove("write", self._var_trace)
        except Exception:
            pass

    def _toggle(self, _e):
        self.var.set(not self.var.get())

    def _draw(self):
        if not self.winfo_exists():
            return
        self.delete("all")
        h = self.winfo_height()
        s = self._size
        y = (h - s) / 2
        x = u(2)
        if self.var.get():
            rr(self, x, y, x + s, y + s, u(4), self.sk.accent)
            self.create_line(x + s * 0.26, y + s * 0.52, x + s * 0.44, y + s * 0.70,
                             x + s * 0.76, y + s * 0.30, fill="#FFFFFF",
                             width=u(2), capstyle="round", joinstyle="round")
        else:
            rr_border(self, x, y, x + s, y + s, u(4), "#CBD3DF", self.sk.card)
        self.create_text(x + s + u(9), h / 2, text=self.text,
                         font=f(self.font_size), fill=self.sk.text, anchor="w")


class RoundProgress(tk.Canvas):
    """圆角进度条；停止态用琥珀色。"""

    def __init__(self, master, sk, height=10, width=None):
        bg = parent_bg(master, sk.card)
        kw = dict(bg=bg, highlightthickness=0, bd=0, height=u(height))
        if width:
            kw["width"] = u(width)
        super().__init__(master, **kw)
        self.sk = sk
        self._value = 0.0
        self._stopped = False
        self.bind("<Configure>", lambda e: self._draw())

    def set_value(self, pct, stopped=False):
        try:
            pct = float(pct)
        except (TypeError, ValueError):
            pct = 0.0
        self._value = max(0.0, min(100.0, pct))
        self._stopped = bool(stopped)
        self._draw()

    def reset(self):
        self.set_value(0)

    def _draw(self):
        if not self.winfo_exists():
            return
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 4 or h < 4:
            return
        track = self.sk.track
        fill = self.sk.warn if self._stopped else self.sk.accent
        rr(self, 0, 1, w - 1, h - 2, (h - 3) / 2, track)
        fw = (w - 1) * self._value / 100.0
        if fw >= h - 2:
            rr(self, 0, 1, fw, h - 2, (h - 3) / 2, fill)
        elif fw > 2:
            rr(self, 0, 1, 2 * (h - 3) / 2 + 2, h - 2, (h - 3) / 2, fill)


_TS_RE = re.compile(r"^\[\d{2}:\d{2}:\d{2}\]\s*")


class LogView(tk.Frame):
    """带分级着色的日志区（✓ 绿 / ⚠ 琥珀 / ✗ 红 / 标题蓝）。

    max_lines > 0 时自动裁剪最早的行：上千张图的长任务（每张 2~4 行 + 每 3 秒
    心跳）会堆到几万行，滚动越来越卡、也白占内存。
    """

    def __init__(self, master, sk, pad=0, max_lines=0):
        super().__init__(master, bg=sk.card)
        self.sk = sk
        self.max_lines = max(0, int(max_lines))
        self._since_trim = 0
        self.txt = tk.Text(self, bd=0, highlightthickness=0, bg=sk.card,
                           fg=sk.text, font=f(sk.fs_log), wrap="word",
                           padx=u(2), pady=u(2), spacing1=u(1), spacing3=u(2))
        self.txt.tag_configure("ok", foreground=sk.ok)
        self.txt.tag_configure("warn", foreground=sk.warn)
        self.txt.tag_configure("bad", foreground=sk.bad)
        self.txt.tag_configure("hl", foreground=sk.accent_d)
        self.txt.tag_configure("dim", foreground=sk.faint)
        self.sb = ttk.Scrollbar(self, orient="vertical", style="P.Vertical.TScrollbar",
                                command=self.txt.yview)
        self.txt.configure(yscrollcommand=self.sb.set)
        self.sb.pack(side=tk.RIGHT, fill=tk.Y, padx=(u(4), 0))
        self.txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.txt.configure(state=tk.DISABLED)

    @staticmethod
    def tag_for(msg):
        # 判级只看正文：主程序写的是「[HH:MM:SS] 正文」，加了时间戳后
        # startswith("===") 会失效，所以先剥掉可选的时间戳前缀。
        body = _TS_RE.sub("", msg, count=1)
        if "✓" in body:
            return "ok"
        if "⚠" in body:
            return "warn"
        if "✗" in body:
            return "bad"
        if body.startswith("===") or body.startswith("---"):
            return "hl"
        return None

    def append(self, msg, tag="auto"):
        """追加一行。tag="auto" 按内容判级；显式传 None 表示不着色。"""
        if tag == "auto":
            tag = self.tag_for(msg)
        self.txt.configure(state=tk.NORMAL)
        self.txt.insert(tk.END, msg + "\n", tag or ())
        self.txt.see(tk.END)
        self.txt.configure(state=tk.DISABLED)
        self._since_trim += 1
        if self.max_lines and self._since_trim >= 200:
            # 每 200 行才查一次总行数（不逐行算，省开销）
            self._since_trim = 0
            self.trim()

    def trim(self):
        """行数超限时裁掉最早的部分。"""
        if not self.max_lines:
            return
        try:
            total = int(self.txt.index("end-1c").split(".")[0])
            if total > self.max_lines:
                self.txt.configure(state=tk.NORMAL)
                self.txt.delete("1.0", f"{total - self.max_lines}.0")
                self.txt.configure(state=tk.DISABLED)
        except Exception:
            pass

    def clear(self):
        """清空全部内容。"""
        self._since_trim = 0
        self.txt.configure(state=tk.NORMAL)
        self.txt.delete("1.0", tk.END)
        self.txt.configure(state=tk.DISABLED)

    def set_text(self, text):
        """整段装载纯文本（不做分级着色），用于「使用说明」这类静态内容。"""
        self._since_trim = 0
        self.txt.configure(state=tk.NORMAL)
        self.txt.delete("1.0", tk.END)
        self.txt.insert(tk.END, text)
        self.txt.configure(state=tk.DISABLED)
        self.txt.see("1.0")

    def load(self, lines):
        self.txt.configure(state=tk.NORMAL)
        self.txt.delete("1.0", tk.END)
        for text, tag in lines:
            self.txt.insert(tk.END, text + "\n", tag or ())
        self.txt.see(tk.END)
        self.txt.configure(state=tk.DISABLED)


def style_ttk(sk):
    """把 ttk 的 Entry / Scrollbar 调成与皮肤一致的扁平风。"""
    st = ttk.Style()
    try:
        st.theme_use("clam")
    except Exception:
        pass
    st.configure("P.TEntry", fieldbackground=sk.card, foreground=sk.text,
                 insertcolor=sk.text, bordercolor=sk.border, lightcolor=sk.border,
                 darkcolor=sk.border, padding=(u(8), u(sk.entry_pad_y)),
                 relief="flat")
    st.configure("P.Vertical.TScrollbar", gripcount=0, background="#C9D2DE",
                 darkcolor="#C9D2DE", lightcolor="#C9D2DE",
                 troughcolor="#F6F8FB", bordercolor="#F6F8FB",
                 arrowcolor=sk.card, arrowsize=u(10))


def make_link(parent, sk, text, command, font_size=None, bg=None):
    """顶栏/卡片头的文字链接：手型光标 + 悬停变蓝。"""
    lb = tk.Label(parent, text=text, bg=bg or sk.card, fg=sk.muted,
                  font=f(font_size or sk.fs_body), cursor="hand2")
    lb.bind("<Enter>", lambda _e: lb.configure(fg=sk.accent))
    lb.bind("<Leave>", lambda _e: lb.configure(fg=sk.muted))
    lb.bind("<Button-1>", lambda _e: command())
    return lb
