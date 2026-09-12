"""
补全弹窗 UI：基于 tkinter 的无焦点悬浮列表。

要点：
  - 窗口用 overrideredirect + 不抢焦点，避免打断 VBE 输入；
  - 字号自动跟随 VBE 编辑器设置（读注册表）并比其小一号；
  - 位置紧贴 VBE 代码窗的光标（用 Win32 GetGUIThreadInfo 取 caret 屏幕坐标）；
  - 用 Canvas 自绘每一行，把【命中的字符标红】——Listbox 做不到部分着色；
  - 所有 UI 操作通过 root.after 回到主线程执行；
  - 键盘选择/确认由 main.py 的全局钩子统一处理；
  - 鼠标单击/双击也可确认补全。
"""

import tkinter as tk
import tkinter.font as tkfont

# 相对 VBE 字号的偏移（-1 = 小一号）
FONT_DELTA = -1
MIN_FONT_SIZE = 8
MAX_FONT_SIZE = 24
MAX_ROWS = 8

# ---- 配色 ----
COLOR_BG = "#FFFFF0"        # 未选中行底色（沿用原来的淡黄）
COLOR_FG = "#202020"        # 未选中行文字
COLOR_HIT = "#C00000"       # 命中字符：红（用户要求"更醒目"）
COLOR_SEL_BG = "#2A5CAA"    # 选中行底色
COLOR_SEL_FG = "#FFFFFF"    # 选中行文字
COLOR_SEL_HIT = "#FFD24A"   # 选中行命中字符：暖黄
#                              不用红色是因为红字压在深蓝底上对比度太差、反而看不清，
#                              换成暖黄后在任何背景下都跳出来。
COLOR_NUM = "#909090"       # 序号（未选中）
COLOR_NUM_SEL = "#B9CDE8"   # 序号（选中行）

PAD_X = 6                   # 行内左右留白
NUM_GAP = 6                 # 序号与名字之间的间距


def _vbe_font_info():
    """读取 VBE 编辑器字体设置，返回 (FontFace, FontHeight像素)；失败返回 None。"""
    try:
        import winreg
        for ver in ("7.1", "7.0", "6.0"):
            try:
                k = winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER,
                    "Software\\Microsoft\\VBA\\%s\\Common" % ver,
                )
                face, height = None, None
                try:
                    face, _ = winreg.QueryValueEx(k, "FontFace")
                except Exception:
                    pass
                try:
                    height, _ = winreg.QueryValueEx(k, "FontHeight")
                except Exception:
                    pass
                winreg.CloseKey(k)
                if height:
                    return (face or "Consolas"), int(height)
            except Exception:
                continue
    except Exception:
        pass
    return None


def split_by_hits(name, positions):
    """把名字按"命中/未命中"切成若干段，供逐段上色。

    返回 [(片段文本, 是否命中), ...]，相邻同类字符会合并成一段，
    这样 Canvas 只需要画很少几个文本项（而不是一个字符一个）。

      split_by_hits("dataSheet", [2, 4])
      -> [("da", False), ("t", True), ("a", False), ("S", True), ("heet", False)]
    """
    hit = set(positions or ())
    out = []
    i = 0
    n = len(name)
    while i < n:
        j = i + 1
        is_hit = i in hit
        while j < n and (j in hit) == is_hit:
            j += 1
        out.append((name[i:j], is_hit))
        i = j
    return out


def format_row(index, name):
    """列表第 index 行（0-based）的显示文本：`1. num` 这样带序号。

    序号只用于显示与"按数字键选词"，真正的补全值仍按【下标】从 matches 里取，
    所以不会因为前缀而把 "1. num" 写进编辑器。

    注：Canvas 渲染上线后这里不再参与绘制，仅保留给测试与调试使用。
    """
    return "%d. %s" % (index + 1, name)


def caret_screen_rect():
    """取当前前台窗口光标的屏幕坐标 (x, top, bottom)；失败返回 None。

    win32gui 未暴露 GetGUIThreadInfo，这里用 ctypes 直接调 Win32 API。
    """
    try:
        import ctypes
        from ctypes import wintypes

        user32 = ctypes.windll.user32

        class GUITHREADINFO(ctypes.Structure):
            _fields_ = [
                ("cbSize", wintypes.DWORD),
                ("flags", wintypes.DWORD),
                ("hwndActive", wintypes.HWND),
                ("hwndFocus", wintypes.HWND),
                ("hwndCapture", wintypes.HWND),
                ("hwndMenuOwner", wintypes.HWND),
                ("hwndMoveSize", wintypes.HWND),
                ("hwndCaret", wintypes.HWND),
                ("rcCaret", wintypes.RECT),
            ]

        hwnd = user32.GetForegroundWindow()
        if not hwnd:
            return None
        tid = user32.GetWindowThreadProcessId(hwnd, None)
        if not tid:
            return None

        g = GUITHREADINFO()
        g.cbSize = ctypes.sizeof(GUITHREADINFO)
        if not user32.GetGUIThreadInfo(tid, ctypes.byref(g)):
            return None
        if not g.hwndCaret:
            return None

        p_top = wintypes.POINT(g.rcCaret.left, g.rcCaret.top)
        p_bot = wintypes.POINT(g.rcCaret.left, g.rcCaret.bottom)
        user32.ClientToScreen(g.hwndCaret, ctypes.byref(p_top))
        user32.ClientToScreen(g.hwndCaret, ctypes.byref(p_bot))
        return p_top.x, p_top.y, p_bot.y
    except Exception:
        return None


class Popup:
    """用 Canvas 自绘的候选列表。

    为什么不用 Listbox：Listbox 每一行只能有一种前景色，没法把"命中的字
    符"单独标红。Canvas 可以逐段 create_text，想怎么着色就怎么着色。

    每行画三样东西：选中底色 → 灰色序号 → 名字（按命中位置切成若干段，
    命中的段用高亮色）。行高固定，最多 MAX_ROWS 行，不足留白，
    避免列表随匹配数跳动。
    """

    def __init__(self, root):
        self.root = root
        self.win = None
        self.canvas = None
        self.completer = None
        self.font_obj = tkfont.Font(family="Consolas", size=10)
        # 是否"应该"显示。show() 置 True、hide() 置 False，
        # 用来取消已排队但尚未执行的 _show，避免"按了 Tab 又冒出来"。
        self._want_visible = False
        self._pending_show = None
        self._rows = []          # [(name, [命中下标...]), ...]
        self._sel = 0
        self._row_h = 18
        self._build()

    def _build(self):
        self.win = tk.Toplevel(self.root)
        self.win.withdraw()
        self.win.overrideredirect(True)
        self.win.wm_attributes("-topmost", True)
        self.canvas = tk.Canvas(
            self.win, highlightthickness=0, bd=0, bg=COLOR_BG,
        )
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.canvas.bind("<ButtonRelease-1>", self._on_click)
        self.canvas.bind("<Double-Button-1>", self._on_double)

    # ---- 鼠标 ----
    def _row_at(self, y):
        """y 坐标落在第几行（0-based）；落在空白区返回 None。"""
        if self._row_h <= 0:
            return None
        idx = int(y // self._row_h)
        if 0 <= idx < len(self._rows):
            return idx
        return None

    def _on_click(self, event):
        if self.completer is None:
            return
        idx = self._row_at(event.y)
        if idx is None:          # 点到空白行：不确认（与原来"空行占位"行为一致）
            return
        # 走 completer.pick：它会同步 selected 再 accept，且越界时自动忽略
        self.completer.pick(idx)

    def _on_double(self, event):
        # 单击已经确认并收起了，这里兜底：万一还可见就再确认一次
        if self.completer is None:
            return
        idx = self._row_at(event.y)
        if idx is None:
            return
        self.completer.pick(idx)

    # ---- 字体：跟随 VBE 设置，比其小一号 ----
    def _apply_font(self):
        info = _vbe_font_info()
        if info:
            face, px = info
            pt = None
            try:
                per_pt = self.root.winfo_fpixels("1p")
                if per_pt and per_pt > 0:
                    pt = px / per_pt
            except Exception:
                pt = None
            if not pt:
                pt = px * 72.0 / 96.0
            size = max(MIN_FONT_SIZE, min(MAX_FONT_SIZE, pt + FONT_DELTA))
            family = face
        else:
            family, size = "Consolas", 10
        try:
            self.font_obj.configure(family=family, size=size)
        except Exception:
            try:
                self.font_obj.configure(size=size)
            except Exception:
                pass

    def _cancel_pending_show(self):
        if self._pending_show:
            try:
                self.root.after_cancel(self._pending_show)
            except Exception:
                pass
            self._pending_show = None

    # ---- 显示 ----
    def show(self, matches, selected, completer):
        self.completer = completer
        self._want_visible = True
        self._cancel_pending_show()
        self._pending_show = self.root.after(0, self._show, matches, selected)

    def _show(self, matches, selected):
        self._pending_show = None
        # 已被 hide() 取消（例如刚按过 Tab 确认），就不要再冒出来
        if not self._want_visible:
            return
        self._apply_font()   # 每次弹出都刷新，VBE 改字号后即时生效
        # 取命中位置：由 Completer.trigger 填好的 {名字: [下标...]}
        hits = {}
        try:
            hits = getattr(self.completer, "match_hits", None) or {}
        except Exception:
            hits = {}
        self._rows = [(m, list(hits.get(m) or ())) for m in matches]
        self._sel = selected or 0
        self._draw()
        self._position()
        self.win.deiconify()
        self.win.lift()

    def _draw(self):
        """把当前 _rows 画到 Canvas 上（每次全量重绘，最多 8 行，开销可忽略）。"""
        c = self.canvas
        c.delete("all")
        try:
            self._row_h = self.font_obj.metrics("linespace") or 18
        except Exception:
            self._row_h = 18
        row_h = self._row_h

        # 先定宽：Canvas 里 create_text 不会自动撑开，得自己算
        num_w = max([self._measure("%d." % (i + 1))
                     for i in range(len(self._rows))] or [self._measure("1.")])
        text_w = max([self._measure(name) for name, _ in self._rows] or [0])
        width = PAD_X * 2 + num_w + NUM_GAP + text_w + 2
        height = MAX_ROWS * row_h
        c.config(width=width, height=height)

        for i, (name, hit_pos) in enumerate(self._rows):
            y0 = i * row_h
            ym = y0 + row_h // 2          # 垂直居中基线
            on = (i == self._sel)
            bg = COLOR_SEL_BG if on else COLOR_BG
            fg = COLOR_SEL_FG if on else COLOR_FG
            hit_fg = COLOR_SEL_HIT if on else COLOR_HIT
            num_fg = COLOR_NUM_SEL if on else COLOR_NUM

            if on:
                c.create_rectangle(0, y0, width, y0 + row_h,
                                   fill=bg, outline="")
            # 序号（仅显示，不影响补全值）
            c.create_text(PAD_X, ym, anchor="w", text="%d." % (i + 1),
                          fill=num_fg, font=self.font_obj)

            # 名字：按命中位置切成"普通段 / 命中段"交替，逐段上色
            x = PAD_X + num_w + NUM_GAP
            for seg, is_hit in split_by_hits(name, hit_pos):
                c.create_text(x, ym, anchor="w", text=seg,
                              fill=hit_fg if is_hit else fg,
                              font=self.font_obj)
                x += self._measure(seg)

    def _measure(self, text):
        try:
            return self.font_obj.measure(text)
        except Exception:
            return len(text) * 7

    def _position(self):
        # 优先贴着 VBE 光标；取不到就退回鼠标位置
        rect = caret_screen_rect()
        if rect is None:
            try:
                import win32api
                mx, my = win32api.GetCursorPos()
                rect = (mx, my, my + 16)
            except Exception:
                rect = (200, 200, 216)
        x, top, bottom = rect

        w = int(self.canvas.cget("width") or 200)
        h = int(self.canvas.cget("height") or (MAX_ROWS * self._row_h))
        w = min(560, max(180, w))

        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        if x + w > sw:
            x = max(0, sw - w)
        # 光标下方放得下就放下面，否则翻到光标上方
        if bottom + h + 4 <= sh:
            y = bottom + 2
        else:
            y = max(0, top - h - 2)
        self.win.geometry("%dx%d+%d+%d" % (w, h, x, y))

    def update_selection(self, selected):
        self.root.after(0, self._update_selection, selected)

    def _update_selection(self, selected):
        self._select(selected)

    def _select(self, selected):
        if not self._rows:
            return
        self._sel = (selected or 0) % len(self._rows)
        self._draw()

    def contains_point(self, px, py):
        """判断屏幕坐标 (px, py) 是否落在弹窗矩形内（主线程调用才安全）。

        留 4px 余量，避免点在条目边缘时被误判为“弹窗外”而误收起。
        """
        try:
            if self.win is None:
                return False
            if self.win.state() == "withdrawn":
                return False
            margin = 4
            rx = self.win.winfo_rootx() - margin
            ry = self.win.winfo_rooty() - margin
            w = self.win.winfo_width() + 2 * margin
            h = self.win.winfo_height() + 2 * margin
            return (rx <= px <= rx + w) and (ry <= py <= ry + h)
        except Exception:
            return False

    def hide(self):
        # 1) 立刻标记为"不该显示"，并取消尚未执行的 _show（消除 after 队列时序问题）
        self._want_visible = False
        self._cancel_pending_show()
        # 2) 立即收起（主线程调用时生效；失败也不影响下面的兜底）
        try:
            self._hide()
        except Exception:
            pass
        # 3) 兜底：再排一次，确保跨线程调用也不会漏
        try:
            self.root.after(0, self._hide)
        except Exception:
            pass

    def _hide(self):
        self._want_visible = False
        self._rows = []
        if self.win:
            try:
                self.win.withdraw()
            except Exception:
                pass
