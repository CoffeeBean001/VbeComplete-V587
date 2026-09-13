"""
补全弹窗 UI：基于 tkinter 的无焦点悬浮列表。

要点：
  - 窗口用 overrideredirect + 不抢焦点，避免打断 VBE 输入；
    （v41 起额外加 WS_EX_NOACTIVATE：点击/拖动它都【不会】抢走 VBE 前台焦点，
     所以拖滚动条拖多久都不会被焦点逻辑误收起。）
  - 字号自动跟随 VBE 编辑器设置（读注册表）并比其小一号；
  - 位置紧贴 VBE 代码窗的光标（用 Win32 GetGUIThreadInfo 取 caret 屏幕坐标）；
  - 用 Canvas 自绘每一行，把【命中的字符标红】——Listbox 做不到部分着色；
  - 所有 UI 操作通过 root.after 回到主线程执行；
  - 键盘选择/确认由 main.py 的全局钩子统一处理；
  - 鼠标单击/双击也可确认补全；纵向滚动条可拖。
"""

import tkinter as tk
import tkinter.font as tkfont

# 相对 VBE 字号的偏移（-1 = 小一号）
FONT_DELTA = -1
MIN_FONT_SIZE = 8
MAX_FONT_SIZE = 24
MAX_VISIBLE_ROWS = 15          # 一屏最多显示几行（候选不足时列表会相应变短）
NAME_CHARS = 30                # 列表宽度【固定】容纳这么多字符（超长用 ... 省略，不做横向滚动）
ELLIPSIS = "..."               # 超长名字的省略号（仿 IDEA：放不下的字符用 ... 代替）

# ---- 配色 ----
COLOR_BG = "#FFFFF0"        # 未选中行底色（沿用原来的淡黄）
COLOR_FG = "#202020"        # 未选中行文字
COLOR_HIT = "#C00000"       # 命中字符：红（用户要求"更醒目"）
COLOR_SEL_BG = "#2A5CAA"    # 选中行底色
COLOR_SEL_FG = "#FFFFFF"    # 选中行文字
COLOR_SEL_HIT = "#FFD24A"   # 选中行命中字符：暖黄
#                              不用红色是因为红字压在深蓝底上对比度太差、反而看不清，
#                              换成暖黄后在任何背景下都跳出来。
COLOR_HOVER_BG = "#EAF1FB"  # 鼠标悬停行底色（浅蓝，与选中的深蓝区分）
COLOR_DETAIL_BG = "#F4F3EC" # 详情行（完整名字）底色
COLOR_DETAIL_FG = "#333333" # 详情行文字
COLOR_DETAIL_SEP = "#D9D5C6"  # 详情行与列表之间的分隔线

SCROLL_W = 7                # 纵向滚动条宽度（只在候选数超过一屏时出现）
SCROLL_PAD = 2              # 滚动条与名字之间的留白
COLOR_SCROLL_TRACK = "#E8E4D0"
COLOR_SCROLL_THUMB = "#B0A98C"

PAD_X = 6                   # 行内左右留白


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
    """列表第 index 行（0-based）的显示文本：v49 起就是名字本身。

    数字键选词已移除、序号也不再显示，所以显示文本 = 名字；保留 index 参数
    只为兼容调用方。仅作测试 / 调试用途 —— 真正的补全值始终按【下标】从
    matches 里取，与显示文本无关。
    """
    return name


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
    每行画两样东西：选中底色 → 名字（按命中位置切成若干段，命中的段用高亮色）。
    v49 起不再画序号（数字键选词已移除），名字从列表最左侧开始。

    布局（v42 起，仿 IDEA）：
      - 列表宽度【固定】= NAME_CHARS(30) 个字符，不再随内容伸缩，也没有横向
        滚动条；名字超过 30 个字符用 ... 省略。
      - 一屏最多 MAX_VISIBLE_ROWS(15) 行；候选不足时列表高度随之缩短，超过
        15 行才出现【纵向】滚动条（贴列表右边界，支持滚轮与鼠标拖动）。
      - 名字被省略时，把当前"活动"候选（鼠标悬停优先，否则上下键选中项）的
        【完整名字】显示在列表下方的"详情行"里（窗口为此行按需加宽），这样
        选中的超长名字永远能看全 —— 与 IDEA 显示完整补全项的思路一致。
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
        self._rows = []          # [(name, [命中下标...]), ...] 当前这一屏
        self._sel = 0            # 选中行在【本屏】内的行号（0-based）
        self._total = 0          # 候选总数（可能大于一屏）
        self._top = 0            # 本屏第一项在候选列表中的下标
        self._hover = None       # 鼠标悬停的本屏行号（0-based）或 None
        self._row_h = 18
        self._last_size = None   # 上次画出的 (宽,高)，用于判断是否需要重新定位
        # 纵向滚动条几何与拖拽状态（_draw 时填充，供鼠标命中测试使用）
        self._v_track = None      # (x0, y0, x1, y1) 纵向轨道
        self._v_thumb_y0 = 0      # 纵向滑块顶 y
        self._v_th = 0            # 纵向滑块高
        self._v_travel = 0        # 纵向滑块可移动像素
        self._v_max_top = 0       # view_top 最大值（= 候选总数 - 可见行）
        self._drag = None         # "v" / None：当前是否正在拖纵向滚动条
        self._drag_off_v = 0      # 按下点距纵向滑块顶的偏移
        self._pressed = False     # 鼠标左键是否正按在弹窗上（拖拽 / 点击中）
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
        # 鼠标：按下判定是否落在滚动条上（是则进入拖拽），松开确认行；
        # 无按键的移动（悬停）用于把被省略的完整名字显示到详情行。
        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_motion)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.canvas.bind("<Double-Button-1>", self._on_double)
        self.canvas.bind("<Motion>", self._on_hover)
        self.canvas.bind("<Leave>", self._on_leave)

    # ---- 鼠标 ----
    def _row_at(self, y):
        """y 坐标落在第几行，返回它在【完整候选列表】里的下标；空白区返回 None。

        换算成真实下标是为了让 completer.pick(绝对下标) 继续可用 ——
        鼠标点第 3 行，点中的是"当前这一屏的第 3 行"，而不是全局第 3 项。
        """
        if self._row_h <= 0:
            return None
        idx = int(y // self._row_h)
        if 0 <= idx < len(self._rows):
            return self._top + idx
        return None

    def _screen_row_at(self, y):
        """y 坐标落在本屏第几行（0-based，只看这一屏）；空白区/详情行返回 None。"""
        if self._row_h <= 0:
            return None
        row = int(y // self._row_h)
        if 0 <= row < len(self._rows):
            return row
        return None

    def _on_click(self, event):
        if self.completer is None:
            return
        idx = self._row_at(event.y)
        if idx is None:          # 点到空白行：不确认
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

    def _on_hover(self, event):
        """鼠标移动到某一行（无按键）→ 更新详情行要显示的完整名字。"""
        if self._drag is not None:
            return
        row = self._screen_row_at(event.y)
        if row == self._hover:
            return
        self._hover = row
        self._draw()

    def _on_leave(self, event):
        """鼠标移出弹窗 → 详情行回到跟随"键盘选中项"。"""
        if self._hover is None:
            return
        self._hover = None
        self._draw()

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

    def _sync_from_completer(self, matches=None, selected=None):
        """把"要画什么"同步成 Completer 现在的视口状态。

        【单一数据源在 Completer 侧】：它持有完整候选与 view_top，UI 只负责照着画。
        这样滚动不用两边同步簿记 —— 上下键改变了 view_top 之后，这里直接问它要
        最新的那一屏即可。

        matches/selected 只作为【兜底】：老式/测试用的假 completer 没有这些接口时
        退化为"传进来什么画什么"。
        """
        rows = sel = None
        total = top = None
        try:
            rows = [str(m) for m in (self.completer.visible_rows() or [])]
        except Exception:
            rows = None
        try:
            sel = int(self.completer.view_selection())
        except Exception:
            sel = None
        try:
            total, top = self.completer.view_info()
        except Exception:
            total, top = None, 0

        if rows is None:
            rows = [str(m) for m in (matches or [])]
        if len(rows) > MAX_VISIBLE_ROWS:        # 兜底路径的保护：不该发生，但别画到框外
            rows = rows[:MAX_VISIBLE_ROWS]
        if sel is None:
            sel = int(selected or 0)
        if total is None:
            total, top = len(rows), 0
        if top is None:
            top = 0

        hits = {}
        try:
            hits = getattr(self.completer, "match_hits", None) or {}
        except Exception:
            hits = {}

        self._rows = [(m, list(hits.get(m) or ())) for m in rows]
        self._total = total
        self._top = top
        self._sel = sel if rows else 0

    def _show(self, matches, selected):
        self._pending_show = None
        # 已被 hide() 取消（例如刚按过 Tab 确认），就不要再冒出来
        if not self._want_visible:
            return
        self._apply_font()   # 每次弹出都刷新，VBE 改字号后即时生效
        self._hover = None   # 新弹窗：悬停状态归零
        self._sync_from_completer(matches, selected)
        self._draw()
        self._position()
        self.win.deiconify()
        self.win.lift()
        self._apply_noactivate()

    def _apply_noactivate(self):
        """让弹窗【永远不抢焦点】。

        弹窗是 overrideredirect + topmost 的悬浮窗，但默认点击它仍会"激活"它、
        从而把 VBE 从"前台窗口"挤走。一旦 VBE 不再是前台窗口，main.py 的焦点去抖
        （约 600ms）后会判定"已离开 VBE"并自动收起弹窗 —— 表现就是
        "用鼠标拖滚动条拖一小会儿，整个列表框连同滚动条一起消失"。

        加 WS_EX_NOACTIVATE 扩展样式后：弹窗仍能接收鼠标事件（滚动条可拖、行可点），
        但【永远不成为激活窗口】，VBE 始终保住前台身份，焦点驱动逻辑不会再误收起它。
        """
        try:
            import ctypes
            from ctypes import wintypes
            user32 = ctypes.windll.user32
            GWL_EXSTYLE = -20
            WS_EX_NOACTIVATE = 0x08000000
            hwnd = int(self.win.winfo_id())
            if hwnd == 0:
                return
            for name, at, rt in (
                ("GetWindowLongPtrW", [wintypes.HWND, ctypes.c_int], ctypes.c_void_p),
                ("SetWindowLongPtrW",
                 [wintypes.HWND, ctypes.c_int, ctypes.c_void_p], ctypes.c_void_p),
                ("GetWindowLongW", [wintypes.HWND, ctypes.c_int], ctypes.c_int),
                ("SetWindowLongW", [wintypes.HWND, ctypes.c_int, ctypes.c_int],
                 ctypes.c_int),
            ):
                fn = getattr(user32, name, None)
                if fn is not None:
                    try:
                        fn.argtypes = at
                        fn.restype = rt
                    except Exception:
                        pass
            getf = getattr(user32, "GetWindowLongPtrW", None) or user32.GetWindowLongW
            setf = getattr(user32, "SetWindowLongPtrW", None) or user32.SetWindowLongW
            ex = getf(hwnd, GWL_EXSTYLE)
            ex = int(ex) if ex else 0
            setf(hwnd, GWL_EXSTYLE, ctypes.c_void_p(ex | WS_EX_NOACTIVATE))
        except Exception:
            pass

    # ---- 文本截断（仿 IDEA：放不下就用 ... 省略）----
    def _truncate_to_width(self, text, max_w):
        """把 text 截到不超过 max_w 像素。

        返回 (显示文本, 是否被截断, 保留的前缀长度)。放得下就原样返回；
        放不下就二分找最长前缀 prefix，使 `prefix + ELLIPSIS` 仍放得下。
        """
        if self._measure(text) <= max_w:
            return text, False, len(text)
        ell_w = self._measure(ELLIPSIS)
        if ell_w > max_w:
            return ELLIPSIS, True, 0           # 极端：连省略号都放不下
        lo, hi = 0, len(text)
        while lo < hi:
            mid = (lo + hi + 1) // 2
            if self._measure(text[:mid]) + ell_w <= max_w:
                lo = mid
            else:
                hi = mid - 1
        return text[:lo] + ELLIPSIS, True, lo

    def _draw(self):
        """把当前这一屏画到 Canvas 上。

        布局（v42 起，仿 IDEA）：
          - 行数 = 实际候选数，但最多 MAX_VISIBLE_ROWS(=15) 行；候选不足时列表框
            高度跟着缩短（显示几行就几行高）。
          - 宽度【固定】= NAME_CHARS(30) 个字符（+ 滚动条），不再随内容伸缩，
            也没有横向滚动条；名字超过 30 字用 ... 省略。
          - 纵向滚动条贴在列表框右边界；支持鼠标在滑块/轨道上按住拖动。
          - 被省略的"活动"候选（悬停优先，否则选中项）会把【完整名字】显示在
            列表下方一行"详情行"里，窗口为此行按需加宽。
        """
        c = self.canvas
        c.delete("all")
        # 每次重绘先清空滚动条几何；下面只在需要时回填
        self._v_track = None
        try:
            self._row_h = self.font_obj.metrics("linespace") or 18
        except Exception:
            self._row_h = 18
        row_h = self._row_h

        count = len(self._rows)
        n_rows = max(1, min(count, MAX_VISIBLE_ROWS))

        # 固定宽度：NAME_CHARS 个字符的像素宽（与具体字体无关）
        try:
            cw = self._measure("x" * NAME_CHARS) / float(NAME_CHARS)
        except Exception:
            cw = 7.0
        text_w = float(NAME_CHARS) * cw
        # 名字列起点：v49 去掉序号列后，名字从列表最左侧开始
        name_x0 = PAD_X

        # 纵向滚动条：候选超过一屏才需要
        vscroll = self._total > MAX_VISIBLE_ROWS
        sb_v_w = (SCROLL_W + SCROLL_PAD) if vscroll else 0
        # 列表区（含右留白）宽度 —— 选中底色、滚动条都以此为界
        list_w = name_x0 + text_w + sb_v_w + PAD_X
        list_bg_right = name_x0 + text_w + PAD_X

        # 详情行：当前"活动"候选被省略时，显示其完整名字（悬停优先，否则选中项）
        active = None
        if count:
            if self._hover is not None and 0 <= self._hover < count:
                active = self._hover
            elif 0 <= self._sel < count:
                active = self._sel
        detail_text = None
        if active is not None:
            name = self._rows[active][0]
            if self._measure(name) > text_w:
                detail_text = name
        detail_h = row_h if detail_text else 0
        detail_w = (PAD_X + self._measure(detail_text) + PAD_X) if detail_text else 0

        width = int(max(list_w, detail_w))
        height = int(n_rows * row_h + detail_h)
        v_track_bottom = int(n_rows * row_h) - 1

        size_changed = (getattr(self, "_last_size", None) != (width, height))
        c.config(width=width, height=height)
        self._last_size = (width, height)

        for i, (name, hit_pos) in enumerate(self._rows):
            y0 = i * row_h
            ym = y0 + row_h // 2          # 垂直居中基线
            on = (i == self._sel)
            hovered = (self._hover == i)
            if on:
                bg = COLOR_SEL_BG
            elif hovered:
                bg = COLOR_HOVER_BG
            else:
                bg = COLOR_BG
            fg = COLOR_SEL_FG if on else COLOR_FG
            hit_fg = COLOR_SEL_HIT if on else COLOR_HIT

            if on or hovered:
                c.create_rectangle(0, y0, list_bg_right, y0 + row_h,
                                   fill=bg, outline="")
            # 固定宽度：超过 30 字用 ... 省略；命中下标只保留仍在可见前缀内的
            disp, _trunc, prefix = self._truncate_to_width(name, text_w)
            x = name_x0
            for seg, is_hit in split_by_hits(disp, [p for p in hit_pos if p < prefix]):
                c.create_text(x, ym, anchor="w", text=seg,
                              fill=hit_fg if is_hit else fg,
                              font=self.font_obj)
                x += self._measure(seg)

        # 详情行：完整名字铺在列表下方（窗口已按需加宽）
        if detail_text:
            dy0 = n_rows * row_h
            c.create_rectangle(0, dy0, width, dy0 + row_h,
                               fill=COLOR_DETAIL_BG, outline="")
            c.create_line(0, dy0, width, dy0, fill=COLOR_DETAIL_SEP)
            c.create_text(PAD_X, dy0 + row_h // 2, anchor="w", text=detail_text,
                          fill=COLOR_DETAIL_FG, font=self.font_obj)

        if vscroll:
            self._draw_vscrollbar(list_w, sb_v_w, v_track_bottom)

        # 尺寸变了（详情行出现/消失）就重新定位，保证贴着光标
        if size_changed and self.win is not None:
            try:
                if self.win.state() != "withdrawn":
                    self._position()
            except Exception:
                pass

    def _draw_vscrollbar(self, list_w, sb_v_w, v_track_bottom):
        """右侧（列表边界处）纵向滚动条：轨道 + 表示当前位置的滑块。

        overrideredirect 弹窗不抢焦点，真正的 tk.Scrollbar 拖不动，所以自绘
        指示器；v40 起支持鼠标在滑块 / 轨道上按住左键拖动（见 _on_press / _on_motion），
        拖动时改 Completer.view_top。滚动条一律贴在【列表区】右边界，不随详情行
        的加宽而漂移。
        """
        c = self.canvas
        track_x1 = list_w - SCROLL_PAD
        track_x0 = track_x1 - SCROLL_W
        c.create_rectangle(track_x0, 1, track_x1, v_track_bottom,
                           fill=COLOR_SCROLL_TRACK, outline="")

        visible = max(1, len(self._rows))
        total = max(visible, self._total)
        th = max(int((v_track_bottom - 1) * visible / float(total)) - 2, 12)
        travel = max(1, v_track_bottom - 1 - th - 2)
        top = self._top
        max_top = max(1, total - visible)
        y0 = 1 + int(travel * (float(top) / max_top))
        c.create_rectangle(track_x0 + 1, y0, track_x1 - 1, y0 + th,
                           fill=COLOR_SCROLL_THUMB, outline="")
        # 记录几何，供鼠标拖动命中测试
        self._v_track = (track_x0, 1, track_x1, v_track_bottom)
        self._v_thumb_y0 = y0
        self._v_th = th
        self._v_travel = travel
        self._v_max_top = max_top

    # ---- 鼠标拖动纵向滚动条 ----
    def _point_in(self, x, y, r):
        return bool(r) and (r[0] <= x <= r[2]) and (r[1] <= y <= r[3])

    def _on_press(self, event):
        """鼠标按下：先判定是否落在纵向滚动条上；是则进入拖拽，否则松开时确认行。"""
        self._pressed = True
        if self.completer is None:
            return
        if (self._v_track is not None and self._v_max_top > 0
                and self._point_in(event.x, event.y, self._v_track)):
            thumb = (self._v_track[0] + 1, self._v_thumb_y0,
                     self._v_track[2] - 1, self._v_thumb_y0 + self._v_th)
            if self._point_in(event.x, event.y, thumb):
                self._drag_off_v = event.y - self._v_thumb_y0
            else:
                self._drag_off_v = self._v_th // 2
            self._drag = "v"
            self._apply_vscroll_from_y(event.y)
            return
        # 落在行上：等待松开确认（_on_release）

    def _on_motion(self, event):
        if self._drag == "v":
            self._apply_vscroll_from_y(event.y)

    def _on_release(self, event):
        self._pressed = False
        # 拖拽滚动条时，松开只是结束拖拽，不确认候选
        if self._drag is not None:
            self._drag = None
            return
        # 普通按下并松开（没碰滚动条）= 点中某行 → 确认
        self._on_click(event)

    def _apply_vscroll_from_y(self, y):
        if self._v_track is None or self._v_max_top <= 0:
            return
        ty = max(self._v_track[1],
                 min(y - self._drag_off_v, self._v_track[1] + self._v_travel))
        top = int(round((ty - self._v_track[1]) / float(self._v_travel) * self._v_max_top))
        top = max(0, min(top, self._v_max_top))
        if self.completer is not None:
            try:
                self.completer.set_view_top(top)
            except Exception:
                pass

    def _measure(self, text):
        try:
            return self.font_obj.measure(text)
        except Exception:
            return len(text) * 7

    def _position(self):
        # 优先贴着 VBE 光标；取不到就退回鼠标位置。
        # 窗口宽高直接取画布算好的值 —— 已贴合内容（含滚动条）。
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
        h = int(self.canvas.cget("height") or (MAX_VISIBLE_ROWS * self._row_h))

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
        # 上下键滚动后"这一屏"可能整屏都换了，所以这里不能只挪高亮，
        # 必须回 Completer 重新取一次当前窗口（含新的 view_top 与行内容）。
        self._sync_from_completer(None, selected)
        self._draw()

    def scroll_by(self, delta):
        """滚动 delta 行（正数往下、负数往上），由滚轮 / PageUp 这类调用。

        实现上直接转给 completer.move：它会在移动选中项的同时按需滚动窗口，
        两边状态不会走偏。
        """
        if self.completer is None or not self._rows:
            return
        try:
            self.completer.move(delta)
        except Exception:
            pass

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

    def mouse_inside(self):
        """当前鼠标是否落在弹窗矩形内（主线程调用）。"""
        try:
            import win32api
            x, y = win32api.GetCursorPos()
        except Exception:
            return False
        return self.contains_point(x, y)

    def is_busy(self):
        """是否正在与弹窗交互（左键按下 / 拖滚动条）。

        交互期间【绝不允许】被外部逻辑（焦点变化、点弹窗外判定）收起 ——
        保证"选词前 UI 一直可见"。
        """
        return bool(self._pressed) or (self._drag is not None)

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
        self._top = 0
        self._total = 0
        self._sel = 0
        self._hover = None
        self._pressed = False
        self._drag = None
        self._last_size = None
        if self.win:
            try:
                self.win.withdraw()
            except Exception:
                pass
