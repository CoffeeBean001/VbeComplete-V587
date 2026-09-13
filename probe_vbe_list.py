# -*- coding: utf-8 -*-
"""实地探测「VBE 自带的自动列出成员列表」到底是个什么窗口。

用法：
    python probe_vbe_list.py            # 监视 60 秒，或按回车提前结束
    python probe_vbe_list.py 20         # 监视 20 秒

步骤：
  1. 运行本脚本；
  2. 切到 VBE 代码窗，随便敲点东西让 VBE 自己的提示列表弹出来
     （最稳的是敲一个对象加点号，比如 `Debug.`；或者按 Ctrl+J）；
  3. 回来按回车，或等它自己结束。

脚本会先把「没有列表时」的窗口拍个基线，之后凡是新冒出来的窗口都会连同
类名 / 标题 / 位置 / 窗口风格一起打出来 —— 那一行就是 VBE 的提示列表。
把输出贴回来，就能把精确的类名写进 vbe_bridge._VBE_LIST_CLASSES_BASE。

只读：不改焦点、不碰代码、不发按键。
"""
import ctypes
import sys
import time
from ctypes import wintypes

u32 = ctypes.windll.user32
u32.GetWindowTextW.argtypes = [wintypes.HWND, ctypes.c_wchar_p, ctypes.c_int]
u32.GetClassNameW.argtypes = [wintypes.HWND, ctypes.c_wchar_p, ctypes.c_int]

GW_OWNED = 3
GWL_STYLE = -16
WS_POPUP = 0x80000000
WS_CAPTION = 0x00C00000
VBE_MAIN_CLASS = "wndclass_desked_gsk"


def wcls(h):
    b = ctypes.create_unicode_buffer(256)
    u32.GetClassNameW(h, b, 256)
    return b.value


def wtxt(h):
    b = ctypes.create_unicode_buffer(512)
    u32.GetWindowTextW(h, b, 512)
    return b.value


def wrect(h):
    r = wintypes.RECT()
    u32.GetWindowRect(h, ctypes.byref(r))
    return (r.left, r.top, r.right - r.left, r.bottom - r.top)


def wpid(h):
    p = wintypes.DWORD()
    u32.GetWindowThreadProcessId(h, ctypes.byref(p))
    return int(p.value)


def wstyle(h):
    return u32.GetWindowLongW(h, GWL_STYLE) & 0xFFFFFFFF


def describe(h, why="", caret=None):
    st = wstyle(h)
    info = [
        "  hwnd   = 0x%X" % h,
        "  class  = %r" % wcls(h),
        "  title  = %r" % wtxt(h)[:60],
        "  rect   = %s  (left, top, w, h)" % (wrect(h),),
        "  style  = 0x%08X  popup=%s caption=%s visible=%s"
        % (st, bool(st & WS_POPUP), bool(st & WS_CAPTION),
           bool(u32.IsWindowVisible(h))),
    ]
    if caret is not None:
        info.append("  caret  = %s" % (caret,))
    print("[%s]" % why)
    print("\n".join(info))


def snapshot_hwnds(main_h):
    """当前「VBE 主窗口 OWNED 的窗口」+「同进程的顶层窗口」句柄集合。"""
    out = {}

    h = u32.GetWindow(main_h, GW_OWNED)
    n = 0
    while h and n < 128:
        out[int(h)] = "owned"
        h = u32.GetWindow(h, GW_OWNED)
        n += 1

    vbe_pid = wpid(main_h)
    seen = []

    def cb(h, _lp):
        if int(h) != main_h and wpid(h) == vbe_pid:
            seen.append(int(h))
        return True

    EnumWindowsProc = ctypes.WINFUNCTYPE(
        ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    u32.EnumWindows(EnumWindowsProc(cb), 0)
    for hh in seen:
        out.setdefault(hh, "same-proc")
    return out


def caret_point():
    """当前前台窗口光标的 (x, top, bottom)；拿不到返回 None。"""
    try:
        class GUITHREADINFO(ctypes.Structure):
            _fields_ = [
                ("cbSize", wintypes.DWORD), ("flags", wintypes.DWORD),
                ("hwndActive", wintypes.HWND), ("hwndFocus", wintypes.HWND),
                ("hwndCapture", wintypes.HWND), ("hwndMenuOwner", wintypes.HWND),
                ("hwndMoveSize", wintypes.HWND), ("hwndCaret", wintypes.HWND),
                ("rcCaret", wintypes.RECT),
            ]

        hwnd = u32.GetForegroundWindow()
        if not hwnd:
            return None
        tid = u32.GetWindowThreadProcessId(hwnd, None)
        g = GUITHREADINFO()
        g.cbSize = ctypes.sizeof(GUITHREADINFO)
        if not u32.GetGUIThreadInfo(tid, ctypes.byref(g)) or not g.hwndCaret:
            return None
        p_top = wintypes.POINT(g.rcCaret.left, g.rcCaret.top)
        p_bot = wintypes.POINT(g.rcCaret.left, g.rcCaret.bottom)
        u32.ClientToScreen(g.hwndCaret, ctypes.byref(p_top))
        u32.ClientToScreen(g.hwndCaret, ctypes.byref(p_bot))
        return p_top.x, p_top.y, p_bot.y
    except Exception:
        return None


def main():
    try:
        secs = float(sys.argv[1]) if len(sys.argv) > 1 else 60.0
    except Exception:
        secs = 60.0

    main_h = u32.FindWindowW(VBE_MAIN_CLASS, None)
    if not main_h:
        print("没找到 VBE 主窗口（class=%s）。" % VBE_MAIN_CLASS)
        print("请先把 Excel 的 VBE 打开（Alt+F11），再运行本脚本。")
        return 1
    main_h = int(main_h)
    print("VBE 主窗口 hwnd=0x%X pid=%d" % (main_h, wpid(main_h)))

    print("\n--- 拍基线（此刻 VBE 还没弹列表）---")
    base = snapshot_hwnds(main_h)
    for h, why in sorted(base.items()):
        print("0x%-10X %-10s class=%-26r title=%r"
              % (h, why, wcls(h), wtxt(h)[:40]))
    print("基线共 %d 个窗口。" % len(base))

    print("\n--- 开始监视（%.0f 秒）。现在切到 VBE，让它的提示列表弹出来 ---"
          % secs)
    print("（回来按回车可提前结束）")
    deadline = time.time() + secs
    found = []
    try:
        while time.time() < deadline:
            cur = snapshot_hwnds(main_h)
            for h, why in cur.items():
                if h in base or h in found:
                    continue
                found.append(h)
                caret = caret_point()
                describe(h, why="新窗口 %s" % why, caret=caret)
                if caret is not None:
                    pt = wintypes.POINT(
                        caret[0], caret[2] + max(4, (caret[2] - caret[1]) // 2))
                    at = u32.WindowFromPoint(pt)
                    print("  光标下方那点 (x=%d, y=%d) 上的窗口 = 0x%X class=%r"
                          % (pt.x, pt.y, int(at or 0), wcls(at) if at else ""))
                print()
            time.sleep(0.3)
    except KeyboardInterrupt:
        pass

    if not found:
        print("\n没抓到任何新窗口。确认一下：VBE 的「自动列出成员」是开着的吗")
        print("（VBE 里 工具 → 选项 → 编辑器 → 自动列出成员）。")
    else:
        print("\n抓到 %d 个新窗口。把上面的输出贴回来即可。" % len(found))
    return 0


if __name__ == "__main__":
    sys.exit(main())
