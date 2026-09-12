"""
补全弹窗 UI：基于 tkinter 的无焦点悬浮列表框。

要点：
  - 窗口用 overrideredirect + 不抢焦点，避免打断 VBE 输入；
  - 字号自动跟随 VBE 编辑器设置（读注册表）并比其小一号；
  - 位置紧贴 VBE 代码窗的光标（用 Win32 GetGUIThreadInfo 取 caret 屏幕坐标）；
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


def format_row(index, name):
    """列表第 index 行（0-based）的显示文本：`1. num` 这样带序号。

    序号只用于显示与"按数字键选词"，真正的补全值仍按【下标】从 matches 里取，
    所以不会因为前缀而把 "1. num" 写进编辑器。
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
    def __init__(self, root):
        self.root = root
        self.win = None
        self.listbox = None
        self.completer = None
        self.font_obj = tkfont.Font(family="Consolas", size=10)
        # 是否"应该"显示。show() 置 True、hide() 置 False，
        # 用来取消已排队但尚未执行的 _show，避免"按了 Tab 又冒出来"。
        self._want_visible = False
        self._pending_show = None
        self._build()

    def _build(self):
        self.win = tk.Toplevel(self.root)
        self.win.withdraw()
        self.win.overrideredirect(True)
        self.win.wm_attributes("-topmost", True)
        self.listbox = tk.Listbox(
            self.win, font=self.font_obj, height=MAX_ROWS, width=40,
            activestyle="dotbox", bg="#FFFFF0", exportselection=False,
        )
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.listbox.bind("<ButtonRelease-1>", self._on_click)
        self.listbox.bind("<Double-Button-1>", self._on_double)

    def _on_click(self, event):
        if self.completer is None:
            return
        # 点到"空行占位"不当作确认（列表不足 MAX_ROWS 行时用空行补位）
        sel = self.listbox.curselection()
        if sel:
            try:
                if self.listbox.get(sel[0]) == "":
                    return
            except Exception:
                pass
        self.completer.accept()

    def _on_double(self, event):
        if self.completer is None:
            return
        sel = self.listbox.curselection()
        if sel:
            try:
                if self.listbox.get(sel[0]) == "":
                    return
            except Exception:
                pass
        self.completer.accept()

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
        self.listbox.delete(0, tk.END)
        # 每条前面带 1/2/3... 序号：按对应数字键即可直接选中并写入编辑器。
        # 序号只是显示层，确认时仍按【下标】取 self.matches[index]，
        # 所以不会因为加了前缀而把 "1. num" 这种文本写进代码。
        for i, m in enumerate(matches):
            self.listbox.insert(tk.END, format_row(i, m))
        # 至少显示 MAX_ROWS 行：不足用空行占位，避免列表随匹配数跳动，
        # 也方便用方向键浏览（用户要求"不足 8 个剩下的空起"）。
        while self.listbox.size() < MAX_ROWS:
            self.listbox.insert(tk.END, "")
        self._select(selected)
        self._position()
        self.win.deiconify()
        self.win.lift()

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

        items = list(self.listbox.get(0, tk.END))
        max_len = max([len(s) for s in items], default=10)
        try:
            char_w = self.font_obj.measure("0") or 7
            row_h = self.font_obj.metrics("linespace") or 18
        except Exception:
            char_w, row_h = 7, 18

        w = min(560, max(180, int((max_len + 4) * char_w)))
        h = min(len(items), MAX_ROWS) * row_h + 4

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
        if self.listbox.size() == 0:
            return
        idx = selected % self.listbox.size()
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(idx)
        self.listbox.activate(idx)
        self.listbox.see(idx)

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
        if self.win:
            try:
                self.win.withdraw()
            except Exception:
                pass
