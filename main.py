"""
VbeComplete 主入口（修复版）

修复要点（之前"只能弹一次"的根因）：
  1. pynput 的 on_press/on_release 回调一旦 return False，监听器会被永久停止
     （pynput 源码：if f(*args) is False: raise StopException()）。
     所以这里所有回调一律 return True，绝不返回 False。
  2. pynput 的回调里抛出任何异常，监听器也会被停止。
     所以所有回调体都用 try/except 包住，绝不让异常逃逸。
  3. 要"吞掉"按键（不让 Tab/Enter/↑/↓ 传到 VBE），必须用 Windows 专用的
     win32_event_filter + listener.suppress_event()，而不是 return False。
  4. 所有改变 UI / 调 COM 的动作都通过队列交给 tkinter 主线程执行（线程安全）。
  5. 增加看门狗：监听器若意外死亡会自动重启，保证能一直用。

操作方式：
  - 输入标识符字符自动弹出候选（只提示当前作用域可见的名字）；
  - ↑ / ↓   选择候选（循环，无需移动鼠标），再按 Tab 写回编辑器；
    候选超过一屏（最多 15 条）时，上下键会带着窗口一起滚动，右侧显示滚动条；
  - 滚轮    指针放在弹窗上时，上下滚动 = 上下选词（右侧的滚动条也可用鼠标拖动）；
  - Tab     确认补全（唯一确认键；回车/空格按普通输入处理，不当确认）；
  - Shift+Enter  在当前行下方【新起一行】：当前行不拆分（光标右侧的代码留在
    原行），新行缩进与上一行代码起始位置对齐，光标落在新行缩进之后。等价于
    "先把光标移到行尾再按回车"，但一步到位（光标在行中间时同样适用）；
  - Esc     取消；鼠标单击/双击候选也可确认；
  - ← / →   移动光标即收起列表（与"鼠标点到别处"同义），按键照常放行；
  - Ctrl+Space  手动触发。
  - 自动配对：敲 `(` 自动补 `)`、敲 `"` 自动补另一个 `"`，光标停在**中间**；
    光标右边已经有那个右半边时（打完 `"abc` 再按 `"`）就只是【跨过去】，
    不会补出 `""` 双份。注释里不做，有选区时不做。
  注（v49）：数字键 1~9 不再用于选词，候选也不再显示序号 —— 弹窗开着时
  照常输入数字，定义 `s1` / `arr17` 这类名字不会被列表抢走。

  v55：VBE 自己弹「自动列出成员」时我们让位（写 `UserForm1.` 后输入成员名、
  `Dim x As ` 后填类型名、或按 Ctrl+J），只让位那一会儿，名字一打完就恢复。
  那套判据走光标的【语法位置】（见 engine.vbe_list_expected）。

  v63：再加一路【精确】判据 —— 真机实测 VBE 的提示窗（NameListWndClass /
  PopupTipWndClass）是预建复用的真窗口，直接查可见性就行，于是 VBE 在任何
  位置弹的提示都拦得住（最典型的是敲逗号后弹的【参数信息】）。另外按
  Ctrl+Shift+I（VBE 唤出参数信息）也会收起我们的。想彻底关掉让位：设
  VBECOMPLETE_NO_YIELD=1。想关掉自动配对：设 VBECOMPLETE_NO_AUTOPAIR=1。

  v64：提示只在【代码窗格】里有。以前判断"在不在 VBE"只看前台窗口标题，而
  属性窗口 / 工程窗口 / 窗体设计器都是 VBE 主窗口的子窗口 —— 于是那些非写代码
  的工作区里也会弹提示（用户报的：在属性窗口里改属性值时冒出候选）。现在改用
  精确判据：查 VBE 线程里"当前拥有键盘焦点"的窗口，看它父链上有没有代码窗格
  （窗口类 VbaWindow，见 vbe_bridge.vbe_code_pane_focused）。焦点不在代码窗格
  时：不弹候选、并收起已挂着的候选窗；自动配对 / Shift+Enter / Ctrl+Space
  一律不接管（免得把字符写进代码窗格）。判不出来时退回老判据，绝不"拿不准就拦"。
"""

import os
import sys
import time
import ctypes
import queue
import tkinter as tk

try:
    from pynput.keyboard import Listener, Key
except ImportError:
    sys.stderr.write("缺少依赖 pynput，请先运行：pip install -r requirements.txt\n")
    raise

try:
    from pynput.mouse import Listener as MouseListener
except ImportError:
    sys.stderr.write("缺少依赖 pynput，请先运行：pip install -r requirements.txt\n")
    raise

import engine
import vbe_bridge
from vbe_bridge import VbeBackend, com_backoff_remaining
from ui import Popup, MAX_VISIBLE_ROWS
from log import log as _log, log_boot as _log_boot
from log import LOG_ENABLED as _LOG_ENABLED


# ---------------- Windows 虚拟键码 ----------------
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
WM_SYSKEYDOWN = 0x0104
WM_SYSKEYUP = 0x0105

VK_TAB = 0x09
VK_RETURN = 0x0D
VK_SHIFT = 0x10         # 只用来查"Shift 是否按住"（Shift+Enter = 新起一行）
VK_ESCAPE = 0x1B
VK_SPACE = 0x20
VK_LEFT = 0x25
VK_UP = 0x26
VK_RIGHT = 0x27
VK_DOWN = 0x28
VK_PROCESSKEY = 0xE5   # IME 组字过程中的按键（vkCode=229），不携带真实字符
VK_CONTROL = 0x11
VK_MENU = 0x12         # Alt

NAV_VKS = (VK_UP, VK_DOWN)

# KBDLLHOOKSTRUCT.flags 的 bit4：这一下是【注入】的（SendInput / keybd_event
# 发出），不是人真按的。v57 起我们自己兜底重发按键时会带这个标志 —— 必须放行，
# 否则会自己截自己、无限递归。别人家的自动化按键同理，一律不拦。
LLKHF_INJECTED = 0x10

# 兜底重发字符用的 SendInput 结构（KEYEVENTF_UNICODE：直接给字符，不碰键盘
# 布局 / Shift，中英文布局下都准）。
INPUT_KEYBOARD = 1
KEYEVENTF_UNICODE = 0x0004
KEYEVENTF_KEYUP = 0x0002


class _KEYBDINPUT(ctypes.Structure):
    _fields_ = [("wVk", ctypes.c_ushort),
                ("wScan", ctypes.c_ushort),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", ctypes.c_size_t)]


class _INPUT_UNION(ctypes.Union):
    _fields_ = [("ki", _KEYBDINPUT), ("pad", ctypes.c_ubyte * 32)]


class _INPUT(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong), ("u", _INPUT_UNION)]


def send_char(ch):
    """把一个字符原样"还给"系统（keydown+keyup），等价于人敲了一下。

    只用于兜底：自动配对没做成时，按键已经被我们吞了，必须让用户这一下
    【真的输入进去】—— 绝不能"吞掉按键却什么都没发生"。
    用 KEYEVENTF_UNICODE 直接发字符，绕开键盘布局与 Shift 状态。
    """
    try:
        evts = (_INPUT * 2)()
        for i, flags in enumerate((KEYEVENTF_UNICODE,
                                   KEYEVENTF_UNICODE | KEYEVENTF_KEYUP)):
            e = evts[i]
            e.type = INPUT_KEYBOARD
            e.u.ki.wVk = 0
            e.u.ki.wScan = ord(ch)
            e.u.ki.dwFlags = flags
            e.u.ki.time = 0
            e.u.ki.dwExtraInfo = 0
        n = ctypes.windll.user32.SendInput(
            2, ctypes.byref(evts), ctypes.sizeof(_INPUT))
        return int(n) == 2
    except Exception:
        return False


def send_vk(vk):
    """把一个【虚拟键】原样"还给"系统（keydown+keyup）。

    只用于兜底：Shift+↑/↓（跳出候选列表并移动光标）那一下已经被我们吞了，
    而对应的 COM 动作没能落成（Excel 正忙 / 代理失效）。此时必须让这一下按键
    真的发生 —— 绝不能"吞掉按键却什么都没发生"（v57 的老规矩）。
    注入的事件带 LLKHF_INJECTED，钩子会放行，不会自己截自己。
    """
    try:
        evts = (_INPUT * 2)()
        for i, flags in enumerate((0, KEYEVENTF_KEYUP)):
            e = evts[i]
            e.type = INPUT_KEYBOARD
            e.u.ki.wVk = int(vk)
            e.u.ki.wScan = 0
            e.u.ki.dwFlags = flags
            e.u.ki.time = 0
            e.u.ki.dwExtraInfo = 0
        n = ctypes.windll.user32.SendInput(
            2, ctypes.byref(evts), ctypes.sizeof(_INPUT))
        return int(n) == 2
    except Exception:
        return False


def _is_newline_shortcut(vk, shift_down):
    """Shift+Enter 是否应触发"在当前行下方新起一行"（纯判断，便于单测）。

    只认 Enter：Shift+Tab（反缩进）等一律不管。小键盘回车在 VBE 里同样是
    VK_RETURN，因此与主键盘一视同仁。
    """
    return vk == VK_RETURN and bool(shift_down)


def _is_line_nav_shortcut(vk, shift_down):
    """Shift+↑ / Shift+↓ 是否应触发"跳出候选列表 + 光标上/下移一行"（纯判断）。

    v78：只认 ↑/↓ 这两个导航键 + Shift 按住。**必须在"在列表里选词"那条
    分支之前判**，否则 Shift+↑ 会被当成普通 ↑ 去列表里选词 —— 那就又回到
    用户抱怨的"没法把光标移到上下行"了。
    """
    return vk in NAV_VKS and bool(shift_down)


def _pair_char_for_key(vk, shift_down):
    """这个键敲出来的是不是需要自动配对的符号 —— 返回 '(' / '"' / None。

    刻意去问系统「这个 vk 在当前键盘布局下的基础字符是什么」
    （MapVirtualKeyW），而不是写死 vk 常量：布局不同也不会配错。
    小键盘（NumLock 下按 9 就是数字 9）一律不管。
    """
    if not shift_down:
        return None
    try:
        if 0x60 <= vk <= 0x69:                       # 小键盘
            return None
        base = ctypes.windll.user32.MapVirtualKeyW(vk, 2) & 0xFFFF
        if not base:
            return None
        ch = chr(base)
    except Exception:
        return None
    if ch == "9":
        return "("
    if ch == "0":
        return ")"          # 右半边已在光标右边时"跨过去"，不重复插
    if ch in ("'", '"'):
        return '"'
    return None


def _is_text_key(vk):
    """这个键按下去是否【本可能】往代码里写一个字符。纯函数。

    v79 用途：数"按了这类键、文档却没变"的次数 —— 那就是**输入法正在组字**的
    信号（拼音字母被输入法吃掉、还没上屏）。组字状态下"回车"是**上屏键**，
    绝不能被我们抢去当换行用（见 state["pending_keys"] / _enter_indent_key_ok）。
    这是输入法无关的判据：不依赖任何 IME 接口，中/日/韩输入法都一样。
    """
    try:
        vk = int(vk)
    except Exception:
        return False
    if vk == 0x20:                                 # Space
        return True
    if 0x30 <= vk <= 0x39:                         # 0-9
        return True
    if 0x41 <= vk <= 0x5A:                         # A-Z
        return True
    if 0x60 <= vk <= 0x69:                         # 小键盘 0-9
        return True
    if 0xBA <= vk <= 0xC0 or 0xDB <= vk <= 0xDE:   # OEM 标点 ;=,.-/`[] 与引号
        return True
    return False


def _mod_down():
    """Ctrl 或 Alt 是否被按住（组合键不该触发自动配对）。"""
    try:
        u32 = ctypes.windll.user32
        ctrl = bool(u32.GetAsyncKeyState(VK_CONTROL) & 0x8000)
        alt = bool(u32.GetAsyncKeyState(VK_MENU) & 0x8000)
        return ctrl or alt
    except Exception:
        return False


def _shift_down():
    """Shift 此刻是否被按住 —— 直接问系统，不靠按键事件自己累积。

    为什么不用 on_press 维护一个布尔量（ctrl / alt 就是那么做的）：
    pynput 的 win32_event_filter 跑在【系统钩子线程】里（_util/win32.py 的
    _handler -> _convert 同步调用它），而 on_press 是钩子把消息 post 给监听器
    消息循环之后才执行的（keyboard/_win32.py 的 _process）—— 两者不同线程、
    且隔着一次排队。"按住 Shift 再按 Enter"这两个事件紧挨着，等队列里的
    on_press 把状态记下来往往已经晚了。
    GetAsyncKeyState 拿的是 OS 的实时按键状态（pynput 自己也用它判定修饰键），
    与线程 / 队列无关。取不到（非 Windows 等）就返回 False —— 那时 Shift+Enter
    退化成 VBE 原生的回车行为，不会吞掉用户的按键。
    """
    try:
        return bool(ctypes.windll.user32.GetAsyncKeyState(VK_SHIFT) & 0x8000)
    except Exception:
        return False


# 收起键：左右方向键。
# 与"鼠标点到别处"同义——光标离开了正在拼的那个词，补全上下文就失效了，
# 列表不该继续挂着挡视线。按键本身【放行】（不吞），光标该移还得移。
DISMISS_VKS = (VK_LEFT, VK_RIGHT)

# 确认键：**只用 Tab**。
# 回车刻意不确认——回车在 VBE 里是换行，误触发代价太大（用户要求去掉）。
CONFIRM_VKS = (VK_TAB,)

# 候选窗可见时，Shift+↑ / Shift+↓ = **跳出候选列表**，把光标上/下移一行（v78）。
#
# 用户口径：「当提示词列表框出现的时候，按 ↑↓ 会一直在列表框选词，不能将光标
# 移动上下行。帮我新增 Shift+↑，可以跳出列表框，并将光标上移一行；Shift+↓ 同理。」
#
# 为什么不用"放行按键让 VBE 自己动光标"：VBE 原生的 Shift+↑/↓ 是**扩展选区**
# （按住 Shift 按一下会把上一行整行选进去），并不是"光标干净地上移一行"。所以
# 这里由我们吞键 + 用 COM 落光标，见 VbeBackend.move_caret_line。
# 想关掉（Shift+↑/↓ 原样交给 VBE，即原生"扩展选区"）：
#   set VBECOMPLETE_NO_SHIFT_NAV=1
try:
    SHIFT_NAV_JUMP = (os.environ.get("VBECOMPLETE_NO_SHIFT_NAV",
                                     "0").strip() != "1")
except Exception:
    SHIFT_NAV_JUMP = True

# 回车自动缩进（v79）。
#
# 用户口径：「写了一整行注释，回车换行后光标落在行首、不缩进 —— 希望能忽略上方
# 注释行（可能有多行），跟上方第一个非注释行对齐」；「写 For / Do / If / With /
# Sub / Function / Type / Enum 这些结构时，回车能自动缩进」。
#
# 只在【这一下回车确实要改缩进】时才由我们接管换行：判据 enter_indent_wanted
# 是纯函数，跑在键盘钩子线程里（只读轮询留下的快照，绝不碰 COM）；普通代码行、
# 空行、光标不在行尾、行内有 Tab —— 一律不接管，原样交给 VBE，风险为零。
# 万一 COM 没写成，_enter_indent_here 会把这一下回车原样还给系统。
# 想关掉（回车完全回到 VBE 原生）：set VBECOMPLETE_NO_ENTER_INDENT=1
try:
    ENTER_INDENT = (os.environ.get("VBECOMPLETE_NO_ENTER_INDENT",
                                   "0").strip() != "1")
except Exception:
    ENTER_INDENT = True

# 钩子里判"要不要接管回车"时，轮询快照最多允许多旧（秒）。
# 快照是每轮轮询刷新的（100ms 一次），但空闲时会降频到 2s。超过这个年纪就
# 一律不接管：宁可这一次不做缩进，也绝不拿旧行文本去赌用户正在敲的那一行。
ENTER_CTX_MAX_AGE = 0.8

# ---- v79b：块头那行按回车，顺手把【块收尾】也补出来 ----
# 用户的追加要求：「写 if i=1 then，按回车之后，能自动帮我补全后面的 end if，
# 而且光标缩进；with 语句块也是这样，自动补全 end with」。
# 只对"真需要收尾"的块生效（Sub / Function / Property / Type / Enum / If…Then /
# With / Select Case -> End Xxx；For -> Next、Do -> Loop、While -> Wend）；
# Else / ElseIf…Then / Case 是块【内】的分支，条件编译 #If 要配 #End If —— 都不补。
# 下面已经挂着同一个收尾、或这块的下文已经在写了（下一行缩进不比块头浅）也不补，
# 见 vbe_bridge.closer_needed —— 免得把用户已有的代码顶开。
# 想关掉（只留缩进、不补收尾）：set VBECOMPLETE_NO_AUTOCLOSE=1
try:
    AUTO_CLOSE_BLOCK = (os.environ.get("VBECOMPLETE_NO_AUTOCLOSE",
                                       "0").strip() != "1")
except Exception:
    AUTO_CLOSE_BLOCK = True

# 弹窗可见时这些键由 win32_filter 接管（↑/↓ 导航、Tab 确认、Esc 取消），
# 且会被 suppress_event() 吞掉。on_press 里绝不能抢先收起弹窗，否则"按方向键
# 选词"会退化成"按方向键弹窗消失"。
#
# Enter 刻意【不在】这里：回车不当确认键，按回车应当照常换行（并顺带收起弹窗）。
_VK_HANDLED_WHEN_VISIBLE = (Key.up, Key.down, Key.tab, Key.esc)

# 这些修饰键按下时不收起弹窗
_MOD_KEYS = (Key.shift, Key.shift_r, Key.caps_lock, Key.cmd, Key.cmd_r)

# 不依赖键盘事件，因此中文/输入法/粘贴都能可靠触发。
# 100ms 对打字而言几乎无感，且能在连续快速输入时自然合并掉中间态。
POLL_INTERVAL_MS = 100

# 焦点去抖：连续这么多次读数一致，才认定焦点真的变了（reconcile 每 300ms 一次，
# 即约 600ms）。输入法的候选/组字窗口可能短暂抢走前台窗口，若立即判定
# "离开 VBE"，会卸载钩子、收起弹窗、清空轮询基线，导致中文输入永远不触发。
FOCUS_DEBOUNCE = 2

# 弹窗刚出现后的保护期（秒）：这段时间内忽略"非标识符按键"引发的收起。
# 输入法组字/候选过程会产生杂散按键，不应把刚弹出的列表立刻关掉；
# 真正的判定随后由轮询（读编辑器文本）重新给出。
POPUP_GRACE_SEC = 0.15

# 焦点离开 VBE 后的宽限期（秒）。这段时间内行为与"仍在 VBE"完全一致：
# 输入法的候选/组字窗口会短暂抢走前台窗口，若一失焦就停轮询、释放 COM，
# 中文输入就会漏触发。超过宽限期才认定用户真的离开了，转入"空闲轮询 +
# 释放 COM"。
#
# 刻意取得比较短（2 秒）：关工作簿 / 退 Excel 时焦点必然离开 VBE，宽限期越
# 长，我们攥着 Excel 的 COM 引用不放的时间就越久，退出流程（尤其 VBAProject
# 变脏时的写回）被拖得越慢。去抖由 FOCUS_DEBOUNCE（约 600ms）负责，2 秒的
# 宽限对输入法抖动已足够。
FOCUS_GRACE_SEC = 2.0

# 空闲（已离开 VBE 超过宽限期）时的轮询间隔（毫秒）。
# 关工作簿 / 退出 Excel 时如果还按 100ms 猛打 COM，会把关闭拖慢好几秒，
# 甚至把刚退出的 Excel 又唤醒一次。空闲期降到 0.5 次/秒足够兜底。
POLL_INTERVAL_IDLE_MS = 2000

# 钩子抑制冷却期（秒）：当"上一个非 VBE 的前台窗口已被销毁"——即那个窗口
# 是被关掉的、焦点是被它甩回 VBE 的——在此窗口期内绝不挂载系统级钩子。
#
# 用途：关 VSCode / PyCharm 时，IDE 窗口销毁后焦点会被动落回 VBE，若立刻
# 重新挂上 WH_KEYBOARD_LL / WH_MOUSE_LL，仍在收尾（写设置、退子进程）的
# IDE 进程会被钩子拖慢，表现为「窗口关了但要 2~3 秒才真正退出」。给一段
# 冷却期，等那个进程彻底死透再恢复正常挂载；冷却期内 VBE 内不弹补全，
# 代价极小（用户刚关完 IDE，不会马上回头敲 VBA）。
#
# 刻意取得比 VSCode 收尾略长（2.5 秒），宁可多等半秒也不要再拖慢一次关闭。
HOOK_SUPPRESS_SEC = 2.5

# v64：「提示只在【代码窗格】里出现」这道闸门的开关。
#
# 判据是"VBE 线程里当前拥有键盘焦点的窗口，父链上有没有代码窗格（类名
# VbaWindow）"，见 in_vbe_code_area。若你的宿主编译器窗口类名与实测不同、
# 或这道闸门在你机器上误判（该弹的时候不弹），用这个开关退回旧行为
# （前台窗口标题里有 "Microsoft Visual Basic" 就算数）：
#   set VBECOMPLETE_NO_CODE_AREA_GATE=1
try:
    CODE_AREA_GATE = (
        os.environ.get("VBECOMPLETE_NO_CODE_AREA_GATE", "0").strip() != "1")
except Exception:
    CODE_AREA_GATE = True

# 自动配对：敲 `(` 补 `)`、敲 `"` 补另一个 `"`，光标落在中间。
#
# VBE 原生【不会】自动闭合括号 / 引号（VBE_Extras、Rubberduck 都把它当增强功能
# 往外加），所以由我们补上不会出现"两边都补"的双份；VBE 只在【行尾按回车时】
# 才补缺失的右引号，与逐字符输入不冲突。
# 想关掉：set VBECOMPLETE_NO_AUTOPAIR=1
try:
    AUTO_PAIR = os.environ.get("VBECOMPLETE_NO_AUTOPAIR", "0").strip() != "1"
except Exception:
    AUTO_PAIR = True

# 文本变化后，标识符缓存最快多久允许重解析一次（秒）。
# 太小会每敲一个键都全量解析所有模块（大工程会卡），太大则新声明的变量
# 迟迟不进候选。0.4s 兼顾：连打时最多 2.5 次/秒。
ID_REFRESH_MIN_SEC = 0.4

try:
    import win32gui as _win32gui

    _gui = _win32gui
except Exception:
    _gui = None

# 单实例互斥体句柄（必须保持引用，否则会被 GC 释放导致锁失效）
_mutex_handle = None


def acquire_single_instance():
    """保证只允许一个实例在跑。

    多个实例会各自弹一个列表框：A 实例确认后收起了自己的，
    B 实例的可能还留在屏幕上（表现为"按 Tab 列表框不消失"）。
    """
    global _mutex_handle
    try:
        import ctypes

        ERROR_ALREADY_EXISTS = 183
        k32 = ctypes.windll.kernel32
        k32.CreateMutexW.argtypes = [ctypes.c_void_p, ctypes.c_bool, ctypes.c_wchar_p]
        k32.CreateMutexW.restype = ctypes.c_void_p
        h = k32.CreateMutexW(None, False, "VbeComplete_SingleInstance_Mutex")
        if not h:
            return True  # 拿不到就算了，不阻止启动
        _mutex_handle = h
        return k32.GetLastError() != ERROR_ALREADY_EXISTS
    except Exception:
        return True


def in_vbe_code_pane():
    """判断当前前台窗口是否是 VBE（粗判据：只看前台窗口标题）。

    只用于「要不要挂系统级钩子」这个粗粒度决定：属性窗口 / 工程窗口 / 设计器
    也都是 VBE 主窗口的子窗口，它们在这里一律返回 True（挂上钩子是安全的，
    因为真正会动代码的动作另有 in_vbe_code_area 把关）。
    """
    if _gui is None:
        return False
    try:
        hwnd = _gui.GetForegroundWindow()
        title = _gui.GetWindowText(hwnd)
        return "Microsoft Visual Basic" in title
    except Exception:
        return False


def in_vbe_code_area():
    """焦点是否真的在 VBE 的【代码窗格】上（v64，精确判据）。

    与 in_vbe_code_pane（前台窗口标题）的分工：标题判据分不出"VBE 里的哪个
    子窗口"—— 属性窗口、工程窗口、窗体设计器全都满足它。而用户明确要求：
    **非写代码的工作区里不要出现提示**（也不用代他写代码）。

    所以凡是"会弹候选窗"或"会往代码里写字符"的动作，一律改用本判据：
      * 轮询发现代码窗格文本变化 -> 该不该弹候选窗；
      * Shift+Enter 新起一行、敲 `(`/`"` 自动配对；
      * Ctrl+Space 手动唤出候选。

    实现走 vbe_bridge.vbe_code_pane_focused()：查【VBE 线程】里当前拥有键盘
    焦点的窗口，看它的父链上有没有代码窗格（类名 VbaWindow）。用线程内部焦点
    而不是前台窗口，是因为中文输入法的候选窗会短暂抢走前台窗口。
    拿不准时返回 None -> 退回标题判据（绝不在拿不准时一刀拦死）。

    想退回旧行为（前台是 VBE 就认）：set VBECOMPLETE_NO_CODE_AREA_GATE=1。
    若你的宿主代码窗格类名不同（不是 VbaWindow）：
    set VBECOMPLETE_VBE_FRAME_CLASS=... 或直接用上面那个开关关掉本判据。
    """
    if not CODE_AREA_GATE:
        return in_vbe_code_pane()
    try:
        st = vbe_bridge.vbe_code_pane_focused()
    except Exception:
        st = None
    if st is None:
        return in_vbe_code_pane()
    return bool(st)


def vbe_focus_class():
    """VBE 线程里当前拥有键盘焦点的那个窗口的类名（拿不到返回 None）。

    只用于诊断日志：排查"某处什么都不弹"时，靠它一眼看出焦点到底落在哪个
    窗口上（`VbaWindow`=代码窗格 / `wndclass_pbrs`=属性窗口 / `PROJECT`=工程
    窗口 / `NameListWndClass`、`PopupTipWndClass`=VBE 自带的提示窗）。
    """
    try:
        h = vbe_bridge._focused_hwnd_in_vbe()
        return vbe_bridge._window_class_name(h) if h else None
    except Exception:
        return None


def foreground_window_closing():
    """检测当前前台窗口是否正在关闭 / 已无响应。

    关闭 VSCode / PyCharm 等 IDE 时，其主窗口在退出过程中（保存提示、
    销毁窗口、进程收尾）常常仍是前台窗口，且主线程被阻塞、对消息无响应。
    此时若误把它当成"有效前台"去挂载全局钩子，该 IDE 的退出流程就会被
    我们的 WH_KEYBOARD_LL / WH_MOUSE_LL 系统级钩子挡住 —— 表现就是
    「PyCharm / VSCode 关不掉 / 卡死」。

    探测手段：对前台窗口发一个 WM_NULL，并用 SMTO_ABORTIFHUNG 让系统在
    窗口挂起时立即放弃（超时返回 0）。正在关闭 / 已挂起的窗口会超时 ->
    判定为"正在关闭"，此时绝不挂载钩子。

    注意：
      - 前台无有效窗口（已关完、回到桌面）也视为"无有效前台"，不挂载；
      - 前台是正常 VBE 代码窗时必然响应正常 -> 返回 False，不误伤；
      - 拿不到窗口 API 时退化返回 False（保持原可用行为）。
    """
    if _gui is None:
        return False
    try:
        import ctypes
        hwnd = _gui.GetForegroundWindow()
        if not hwnd or not _gui.IsWindow(hwnd):
            return True      # 前台无有效窗口 -> 不挂载
        user32 = ctypes.windll.user32
        WM_NULL = 0x0000
        SMTO_ABORTIFHUNG = 0x0002
        SMTO_BLOCK = 0x0001
        result = ctypes.c_ulong(0)
        # 200ms 足够区分"正常响应"与"退出中卡住"，又不拖慢轮询。
        ok = user32.SendMessageTimeoutW(
            ctypes.c_void_p(hwnd), WM_NULL, 0, 0,
            SMTO_ABORTIFHUNG | SMTO_BLOCK, 200, ctypes.byref(result))
        return ok == 0      # 超时 / 挂起 -> 窗口正在退出
    except Exception:
        return False


def _reconcile_resources(focused, released, backoff, last_focus_time, now, grace,
                         closing=False, suppress=False):
    """纯函数：根据焦点/释放状态决定本周期的资源动作。

    返回 (mount_hooks, release_com)：
      - mount_hooks：仅在「焦点在 VBE」且「前台窗口没有正在关闭」且
        「不在钩子抑制冷却期」时为真。离开 VBE 必须立刻拆钩子，绝不能因
        released / COM 退避而重新挂载（否则系统级钩子会干扰其它程序，导致
        PyCharm / VSCode 关不掉）。此外有两道保险：
          (a) closing=True：即使焦点去抖仍维持"在 VBE"，只要检测到前台窗口
              正在关闭（正在退出的 IDE），也强制不挂载——绕过去抖窗口；
          (b) suppress=True：检测到"上一个非 VBE 前台窗口已被销毁"（即它是
              被关掉、焦点被甩回 VBE 的），冷却期内强制不挂载，避免刚挂上的
              钩子拖慢那个仍在收尾的 IDE 进程（表现为关窗口后 2~3 秒才退出）。
      - release_com：离开 VBE 超过宽限期、或 COM 正处失败退避时为真。

    把「钩子挂载」和「COM 释放」两条原本纠缠在一起的逻辑解耦，避免
    再次出现「离开 VBE 后钩子被反复重新挂载」的回归。
    """
    mount_hooks = bool(focused) and not closing and not suppress
    release_com = (not released) and (
        backoff > 0
        or (not focused and (now - last_focus_time) > grace)
    )
    return mount_hooks, release_com


def _poll_mod_switch(last, snap):
    """两次快照之间是不是【换了模块】——而不是用户敲了字。

    快照 = (行号, 行文本, 模块名)（v64 起多了模块名）。换模块 —— Ctrl+Tab 切
    代码窗、点工程树换模块、属性改动导致 VBE 改写别的模块 —— 会让"同一行 +
    文本不同"看起来像一次编辑，其实用户一个字都没敲，不该据此弹候选。

    模块名只要有任一侧拿不到（宿主不返回 / 旧式后端只给二元组）就【不否决】：
    宁可照旧触发，也绝不因为拿不到模块名把补全整个哑掉。
    """
    try:
        if len(last) < 3 or len(snap) < 3:
            return False
        a = str(last[2] or "")
        b = str(snap[2] or "")
        return bool(a) and bool(b) and a.lower() != b.lower()
    except Exception:
        return False


def _poll_action(typed_sep):
    """纯函数：本轮轮询该做什么 —— 'hide' 还是 'trigger'。

    typed_sep：这一下敲的是空格 / 标点（v52，别拿右边的词来补全）。

    v55 注：「VBE 自己也在弹列表」的让位不在这里判 —— VBE 那个列表画在代码
    窗格上、没有独立窗口，探不到；改为在引擎里按光标的语法位置判（光标停在
    `标识符.` 之后时 VBE 必然弹成员列表），见 engine.vbe_list_expected。
    """
    if typed_sep:
        return "hide"
    return "trigger"


def main():

    _log_boot()

    if not acquire_single_instance():
        return

    try:
        root = tk.Tk()
        root.withdraw()
        backend = VbeBackend()
        popup = Popup(root)
        # 滚动窗口行数以 UI 为准（两边必须是同一个数，否则窗口会露出半行）
        completer = engine.Completer(backend, popup, view_rows=MAX_VISIBLE_ROWS)
    except Exception:
        raise

    action_queue = queue.Queue()
    state = {
        "kbd": None,          # 键盘钩子（仅 VBE 聚焦时挂载）
        "mouse": None,        # 鼠标钩子（仅 VBE 聚焦时挂载）
        "vbe_focused": False, # 上次（去抖后）的焦点状态，用于“切换时”做一次资源清理
        "raw_focused": None,  # 最近一次原始焦点读数，用于去抖
        "focus_streak": 0,    # 同一读数连续出现的次数
        "last_snap": None,    # 上次轮询到的 (行号, 行文本, 模块名)，用于检测内容变化
        "cur_ctx": None,      # v79：(时刻, 行号, 行文本, 光标列)，每轮刷新，给回车自动缩进判据用
        "pending_keys": 0,    # v79：按了"可能写字"的键、文档却还没变的次数（输入法组字中）
        "_popup_sig": None,   # 上次看到的 VBE 提示窗矩形（v66：变了就重摆候选窗）
        "ctrl": False,
        "alt": False,
        "swallowed": set(),   # 被吞掉的 keydown 的 vk，用于吞掉配对 keyup
        "trigger_queued": False,   # 队列里是否已有一个待执行的 trigger
        "last_focus_time": 0.0,   # 最近一次"确实在 VBE 里"的时刻
        "released": False,        # 离开 VBE 后是否已释放过 COM 资源
        "non_vbe_hwnd": None,     # 最近一次"不在 VBE"时记录的前台窗口句柄
        "hook_suppress_until": 0.0,  # 钩子抑制冷却到期时刻（>now 表示抑制中）
    }

    def _com_allowed():
        """此刻是否允许访问 COM。

        - 焦点在 VBE：允许（除非正处于失败退避期）；
        - 焦点已离开且超过宽限期：进入空闲态，轮询降频；
        - COM 连续失败退避中：一律不碰（Excel 正在关闭/重启）。
        """
        if com_backoff_remaining() > 0:
            return False
        if state.get("vbe_focused"):
            return True
        # 启动后还一次都没判定出"焦点在 VBE"（比如窗口标题不含
        # "Microsoft Visual Basic" 的宿主）：照旧轮询，不能让工具整个哑掉。
        if not state.get("last_focus_time"):
            return True
        return (time.time() - state.get("last_focus_time", 0.0)
                <= FOCUS_GRACE_SEC)

    def post(fn, *args):
        """把动作丢给主线程队列（线程安全）。"""
        action_queue.put((fn, args))

    def _run_trigger():
        state["trigger_queued"] = False
        # v72：轮询（自动）路径弹出来的候选窗 —— VBE 的成员列表一冒出来就该让位
        state["popup_manual"] = False
        completer.trigger(True)

    def post_trigger():
        """把"内容变了，重整候选"排进主线程队列，并做【合并】。

        连打 / 连删时 poll_editor 每 100ms 就会排一次，而每次 trigger 都要打
        COM（读上下文 + 解析标识符 + 现场扫描），单次可能几十到几百毫秒。若
        全部排队执行，主线程会被连续占住，表现为"弹窗明显滞后于编辑器"，
        也让"回退删字时的旧片段"在屏幕上多挂一会儿。

        合并后同一时刻只会有一个待执行的 trigger；它执行时读的是【当时的】
        编辑器状态，所以既不会漏哪次变化，也不会积压。
        """
        if state.get("trigger_queued"):
            return
        state["trigger_queued"] = True
        post(_run_trigger)

    def _new_line_here():
        """Shift+Enter：在当前行下方新起一行（缩进对齐上一行）。

        等价于用户"先把光标移到本行末尾，再按回车"，但一步到位 —— 当前行
        【不拆分】（光标右侧的代码留在原行），新行复制上一行的行首空白，
        光标落到新行缩进之后。细节见 VbeBackend.new_line_below。

        顺序：先收起弹窗（光标要换行了，列表留着没意义），再动文本。
        v79 补上兜底：后端没写成（COM 抽风 / 不在代码窗）就把这一下按键原样
        还给系统 —— 绝不"按了 Shift+Enter 却没换行"（VBE 里 Shift+Enter 与
        回车同义，所以补一个纯回车即可）。
        """
        try:
            completer.hide()
        except Exception:
            pass
        _ok = False
        try:
            _ok = backend.new_line_below()
        except Exception:
            _ok = False
        if not _ok:
            _log("newline: COM 未落成 -> 按键原样还给系统")
            send_vk(VK_RETURN)

    def _enter_indent_here():
        """回车自动缩进（v79）+ 自动补块收尾（v79b）：收起弹窗 -> 算出缩进、代出新行。

        与 _new_line_here 的区别只在"缩进怎么来"：这里是 smart 模式 ——
        整行注释则忽略它、跟上方第一个非注释行对齐，块结构开头则缩进一级。

        块头（If…Then / With / Sub …）还会在后面补一行同级的收尾（End If /
        End With / End Sub…），光标停在中间那行 —— 用户直接写块内容就行。
        要不要补由后端按 vbe_bridge.block_closer + closer_needed 决定（下面已经
        挂着同一个收尾、或这块的下文已经在写 -> 不补）。

        后端返回 False（光标不在行尾 / 空行 / 引号没闭合 / COM 抽风…）就把
        这一下回车【原样还给系统】—— 绝不让用户"按了回车却没换行"。
        """
        try:
            completer.hide()
        except Exception:
            pass
        _ok = False
        try:
            _ok = backend.new_line_below(smart=True,
                                         auto_close=AUTO_CLOSE_BLOCK)
        except Exception:
            _ok = False
        if not _ok:
            _log("enterindent: 未接管 -> 回车原样还给系统")
            send_vk(VK_RETURN)

    def _move_caret_here(delta):
        """Shift+↑/↓（候选窗可见时）：收起我们的窗，光标上/下移一行（v78）。

        顺序与 _new_line_here 一致：先收窗（光标要换行了，列表留着没意义），
        再用 COM 落光标。COM 没落成（Excel 正忙 / 代理临时失效）就把这一下按键
        原样还给系统 —— 绝不"吞掉按键却什么都没发生"；代价是那一下退化成 VBE
        原生的 Shift+↑/↓（扩展选区）。
        """
        try:
            completer.hide()
        except Exception:
            pass
        _ok = False
        try:
            _ok = backend.move_caret_line(delta)
        except Exception:
            _ok = False
        if not _ok:
            _log("navline: COM 未落成 -> 按键原样还给系统 (delta=%d)" % delta)
            send_vk(VK_UP if delta < 0 else VK_DOWN)

    def _auto_pair_here(ch):
        """自动配对（`(` -> `()`、`"` -> `""`、右半边已存在则跨过去）。

        【只在主线程执行】—— COM 只能在主线程碰（见 vbe_bridge._get_vbe_cached
        的线程守卫）。按键此时已经被钩子吞掉了，所以这里必须收尾：
        写成功了最好；写不成（注释里 / 有选区 / COM 恰好抽风）就把这一下
        原样还给系统，让 VBE 按原生行为插入 —— 绝不丢用户的按键。
        """
        try:
            if backend.insert_pair(ch):
                return
            _log("autopair: 未接管 %r -> 原样还给系统" % ch)
        except Exception:
            _log("autopair: insert_pair 异常 -> 原样还给系统")
        send_char(ch)

    def drain_actions():
        """主线程循环消费动作队列。"""
        try:
            while True:
                fn, args = action_queue.get_nowait()
                try:
                    fn(*args)
                except Exception:
                    # 绝不静默：一次异常就是一次"输入了却什么都没弹"。诊断模式下
                    # 把动作名与栈记下来，省得下次又靠猜（v65 教训）。
                    try:
                        import traceback as _tb
                        _log("action %r 抛异常:\n%s"
                             % (getattr(fn, "__name__", fn), _tb.format_exc()))
                    except Exception:
                        pass
        except queue.Empty:
            pass
        root.after(10, drain_actions)

    def _enter_indent_key_ok(vk):
        """这一下回车要不要由我们接管（v79，跑在键盘钩子线程里，绝不碰 COM）。

        钩子线程没有 COM 单元（v56 血泪教训），而"要不要吞这一下按键"必须在
        按键当场就定下来 —— 所以这里只吃【轮询留下的快照】：
          * 快照必须够新（ENTER_CTX_MAX_AGE）；过期就一律不接管，宁可这一次不
            做缩进，也不拿旧行文本去赌用户正在敲的那一行；
          * 真正"值不值得接管"交给纯函数 vbe_bridge.enter_indent_wanted
            （整行注释 / 块结构开头 + 光标在行尾 + 行内无 Tab）。
        其余情况一律返回 False：这一下回车原样交给 VBE，零风险。
        """
        try:
            if not ENTER_INDENT or vk != VK_RETURN:
                return False
            if _shift_down() or _mod_down():
                return False
            if int(state.get("pending_keys") or 0) > 0:
                # 有"按下去、却没落到文档里"的按键 —— 多半是输入法正在组字，
                # 这一下回车是【上屏】用的（打拼音时按回车把字送进文档）。
                # 抢了它就会变成"字没上屏、反而多出一行"，所以一律不接管。
                return False
            if com_backoff_remaining() > 0 or not in_vbe_code_area():
                return False
            ctx = state.get("cur_ctx")
            if not ctx:
                return False
            ts, _ln, text, ec = ctx
            if time.time() - float(ts) > ENTER_CTX_MAX_AGE:
                return False
            return bool(vbe_bridge.enter_indent_wanted(text, ec))
        except Exception:
            return False

    def win32_filter(msg, data):
        """
        Windows 低层键盘钩子的事件过滤器（运行在钩子线程）。

        返回 True  = 事件正常传播；
        调用 listener.suppress_event() = 系统级吞掉该按键（VBE 收不到）。

        注意：这里绝对不能让异常逃逸（否则按键行为会异常），
        但 suppress_event() 抛出的异常必须正常向上传播，所以放在 try 之外。
        """
        action = None
        suppress = False
        try:
            # v57：注入的按键一律放行（含我们自己兜底重发的那一下），
            # 否则会自己截自己 -> 无限递归。
            if msg in (WM_KEYDOWN, WM_SYSKEYDOWN, WM_KEYUP, WM_SYSKEYUP):
                try:
                    if int(data.flags) & LLKHF_INJECTED:
                        return True
                except Exception:
                    pass
            vk = data.vkCode

            if msg in (WM_KEYDOWN, WM_SYSKEYDOWN):
                # Shift+Enter = 在当前行下方新起一行（缩进对齐上一行）。
                # 与"弹窗是否可见"无关：弹窗开着也照样接管（顺带收起弹窗）。
                # 只在【焦点真在代码窗格】且【COM 没在失败退避】时才吞键；
                # 否则放行，让 VBE 按原生行为处理（原生 Shift+Enter == 回车），
                # 绝不"吞了按键却什么都没发生"。
                # v64：判据从"前台是 VBE"收紧成"焦点在代码窗格"—— 在属性窗口 /
                # 工程窗口 / 窗体设计器里打字时不能把内容写进代码窗格。
                _pair_ch = (_pair_char_for_key(vk, _shift_down())
                            if AUTO_PAIR else None)
                # v79：数"按了可能写字的键、文档却没变"的次数 —— 输入法正在组字
                # 的信号（拼音被输入法吃掉、还没上屏）。组字状态下回车是【上屏
                # 键】，见 _enter_indent_key_ok：那时绝不接管回车。
                # 文档一变（轮询看到）就把计数清零，所以正常打字时它总是 0。
                if _is_text_key(vk) and not _mod_down() and in_vbe_code_area():
                    state["pending_keys"] = min(
                        99, int(state.get("pending_keys") or 0) + 1)
                if (_is_newline_shortcut(vk, _shift_down())
                        and com_backoff_remaining() <= 0
                        and in_vbe_code_area()):
                    action = (_new_line_here, ())
                    suppress = True
                elif _enter_indent_key_ok(vk):
                    # v79 回车自动缩进：只有"整行注释 / 块结构开头 + 光标在行尾
                    # + 行内无 Tab"才走得到这里（其余一律由上面那个判据挡掉，
                    # 回车原样交给 VBE）。真正的写入在主线程，写不成会把这一下
                    # 回车原样还给系统。
                    action = (_enter_indent_here, ())
                    suppress = True
                elif (_pair_ch
                        and not _mod_down()
                        and com_backoff_remaining() <= 0
                        and in_vbe_code_area()):
                    # 自动配对 —— v57 改【主线程执行】（见 _auto_pair_here）。
                    #
                    # v56 这里曾是同步调用 backend.insert_pair() 才吞键，理由是
                    # "先确知写进去了再吞"。但它犯了个致命错误：**钩子线程不
                    # 能碰 COM**。那个线程没有 COM 单元，调用必然失败；失败又
                    # 走 _com_fail() 把全局退避打满（1s/2s/5s），连主线程的提示
                    # 轮询一起挡住 -> 敲一下 `(` 之后提示要等 2~3 秒才出来，
                    # 而配对本身还一次都没生效过。
                    #
                    # 现在的做法：先吞键（判断条件全是纯 Win32 / 本地状态，
                    # 钩子线程里做是安全的），真正的写入丢给主线程。万一主
                    # 线程也写不进去，_auto_pair_here 会把这一下原样还给系统
                    # —— 依然不会"吞掉按键却什么都没发生"。
                    action = (_auto_pair_here, (_pair_ch,))
                    suppress = True
                elif completer.is_visible():
                    if not in_vbe_code_pane():
                        # 焦点已离开 VBE：收起弹窗，按键照常放行
                        post(completer.hide)
                    elif vk == VK_ESCAPE:
                        action = (completer.hide, ())
                        suppress = True
                    elif _is_line_nav_shortcut(vk, _shift_down()):
                        # v78：Shift+↑/↓ = 跳出候选列表 + 光标上/下移一行。
                        # 刻意【不】走下面的"在列表里选词"分支 —— 这正是用户要的
                        # "跳出列表框"的手感。按键由我们接管（COM 落光标），因为
                        # VBE 原生的 Shift+↑/↓ 是扩展选区、不是单纯移动光标。
                        # 开关关掉（或 COM 退避中）则不吞键：原样交给 VBE。
                        if (SHIFT_NAV_JUMP
                                and com_backoff_remaining() <= 0
                                and in_vbe_code_area()):
                            post(completer.hide)
                            action = (_move_caret_here,
                                      (-1 if vk == VK_UP else 1,))
                            suppress = True
                    elif vk in NAV_VKS:
                        action = (completer.move, (-1 if vk == VK_UP else 1,))
                        suppress = True
                    elif vk in DISMISS_VKS:
                        # 左/右：移动光标即放弃本次补全，收起弹窗。
                        # 刻意【不】suppress —— 按键要照常传给 VBE 让光标移动。
                        # 也刻意不用 action：这里没有需要同步的后续动作。
                        # 收起后不会被轮询弹回来：poll_editor 只在「行号+行文本
                        # 」变化时才触发，单纯移动光标不改变它。
                        post(completer.hide)
                    elif vk in CONFIRM_VKS:
                        action = (completer.accept, ())
                        suppress = True
                if suppress:
                    state["swallowed"].add(vk)

            elif msg in (WM_KEYUP, WM_SYSKEYUP):
                # 吞掉与 keydown 配对的 keyup，避免 VBE 收到半截按键
                if vk in state["swallowed"]:
                    state["swallowed"].discard(vk)
                    suppress = True

        except Exception:
            return True

        if action is not None:
            post(action[0], *action[1])
        if suppress:
            # 注意：焦点驱动重构后，键盘监听器存于 state["kbd"]（不是 "listener"）。
            # 取错键会导致 suppress_event 永不执行 -> 方向键/Tab/Esc 没被吞掉，
            # 既传到 VBE 又触发 on_press 收起弹窗，"上下键选词"因此失效。
            lst = state.get("kbd")
            if lst is not None:
                lst.suppress_event()   # 抛异常 -> 由 pynput 转成 hook 返回 1
        return True

    def _is_vk_char(key, ch):
        """按键是否是某个字母（Ctrl+J 这类组合键用）。

        Ctrl 按住时 pynput 有时给不出 char（ToUnicode 会折算成控制字符），
        所以 vk 与 char 两条路都认。
        """
        try:
            if getattr(key, "vk", None) == ord(ch.upper()):
                return True
        except Exception:
            pass
        try:
            c = getattr(key, "char", None)
            return bool(c) and c.lower() == ch.lower()
        except Exception:
            return False

    def _is_ident_char(key):
        """是否是标识符字符（字母/数字/下划线）。空格、标点、功能键都不算。"""
        try:
            ch = getattr(key, "char", None)
            return bool(ch) and (ch.isalnum() or ch == "_")
        except Exception:
            return False

    def on_press(key):
        # 绝不 return False（否则监听器会被停止）；绝不让异常逃逸。
        try:
            if key in (Key.ctrl_l, Key.ctrl_r):
                state["ctrl"] = True
            elif key in (Key.alt_l, Key.alt_r):
                state["alt"] = True
            elif key in _MOD_KEYS:
                pass
            elif key == Key.space and state["ctrl"] and not state["alt"]:
                # 手动触发：用户点名要我们的列表。即便光标停在 VBE 也会弹列表
                # 的位置（`标识符.` 之后）也照弹 —— 引擎只对【轮询自动触发】
                # 让位，手动唤出永远有效。
                # v64：但焦点得真在代码窗格上（属性窗口里按 Ctrl+Space 不弹）。
                if in_vbe_code_area():
                    # v72：手动唤出的是用户点名要的列表，不该被下面那条
                    # "VBE 成员列表出现就让位"的兜底收掉 —— 否则就成了
                    # "按了 Ctrl+Space 却没反应"（比被遮住更让人以为工具坏了）。
                    state["popup_manual"] = True
                    post(completer.trigger)
            elif state["ctrl"] and not state["alt"] and _is_vk_char(key, "j"):
                # Ctrl+J / Ctrl+Shift+J = VBE 唤出它自己的「列出属性/方法」。
                # 这是唯一能确定"VBE 列表马上要出现"的时刻（文本没变，轮询看
                # 不出来），这里直接收起我们的，让它弹。
                post(completer.hide)
            elif state["ctrl"] and _is_vk_char(key, "i"):
                # Ctrl+Shift+I（以及 Ctrl+I）= VBE 唤出「参数信息 / 快速信息」。
                # 真机实测（v63）：Ctrl+Shift+I 后 VBE 的参数签名窗
                # PopupTipWndClass 立刻可见、Esc 立刻不可见 —— 确实有效的是
                # 带 Shift 那个。与 Ctrl+J 同理：文本没变、轮询看不出来，所以
                # 直接收起我们的，让 VBE 那个弹（用户口径：VBE 自带弹窗优先）。
                post(completer.hide)
            elif completer.is_visible() and not _is_ident_char(key) \
                    and key not in _VK_HANDLED_WHEN_VISIBLE:
                # 弹窗开着时按了非标识符、非导航键（空格/标点/退格/换行等）
                # → 收起弹窗，按键本身照常传给 VBE。
                # ↑/↓/Tab/Esc 由 win32_filter 接管（导航/确认/取消）且已被吞掉，
                # 这里不再抢先收起，否则“按方向键选词”会退化成“方向键把弹窗消掉”。
                # 注：内容变化引起的“该弹/该收”由 poll_editor 轮询判定——
                # 键盘事件在输入法(IME)下拿不到真实中文字符，不能作为唯一依据。
                # VK_PROCESSKEY(229) 是 IME 组字过程键、不携带真实字符，
                # 必须跳过，否则组字时会把弹窗误关。
                vk = getattr(key, "vk", None)
                if vk == VK_PROCESSKEY:
                    pass
                else:
                    ch = getattr(key, "char", None)
                    if ch is not None or key in (Key.backspace, Key.enter):
                        # 弹窗刚出现时有短暂保护期：输入法组字/候选过程会产生
                        # 杂散按键，不应把刚弹出的列表立刻关掉。
                        # 真正的"该不该弹"随后由轮询读编辑器文本重新判定。
                        if (time.time() - getattr(completer, "shown_at", 0.0)
                                >= POPUP_GRACE_SEC):
                            post(completer.hide)
        except Exception:
            pass
        return True

    def on_release(key):
        # 内容变化不再由键盘事件触发（改由 poll_editor 轮询），这里只维护修饰键。
        try:
            if key in (Key.ctrl_l, Key.ctrl_r):
                state["ctrl"] = False
            elif key in (Key.alt_l, Key.alt_r):
                state["alt"] = False
        except Exception:
            pass
        return True

    def on_mouse_click(x, y, button, pressed):
        """鼠标点击：弹窗可见且点击落在弹窗外时收起（点击弹窗内由列表框绑定确认）。"""
        if pressed and completer.is_visible():
            post(completer.maybe_hide_on_outside_click, x, y)
        return True

    def scroll_popup(x, y, dx, dy):
        """指针停在弹窗上时，滚轮用来上下浏览候选（纵向滚动）。

        必须在主线程里判 contains_point —— 窗口几何查询跨线程读会得到半截状态，
        所以整个判断都塞进 post 里，而不是在监听器里算好再 post 结果。
        pynput 约定 dy > 0 向上滚、dy < 0 向下滚。

        v42 起列表不再有横向滚动（超长名字用 ... 省略，完整名字显示在详情行），
        所以滚轮只做纵向。
        """
        try:
            if not completer.is_visible():
                return
            if not popup.contains_point(x, y):
                return
        except Exception:
            return
        if dy != 0:
            completer.move(-1 if dy > 0 else 1)

    def on_mouse_scroll(x, y, dx, dy):
        if completer.is_visible():
            post(scroll_popup, x, y, dx, dy)
        return True

    def _make_kbd():
        lst = Listener(
            on_press=on_press,
            on_release=on_release,
            win32_event_filter=win32_filter,
        )
        lst.daemon = True
        return lst

    def _make_mouse():
        lst = MouseListener(on_click=on_mouse_click, on_scroll=on_mouse_scroll)
        lst.daemon = True
        return lst

    def _stop_and_clear(key):
        lst = state.get(key)
        if lst is not None:
            try:
                if lst.is_alive():
                    lst.stop()
            except Exception:
                pass
        state[key] = None

    def reconcile():
        """焦点驱动：仅当 VBE 代码窗是前台窗口时才挂键盘/鼠标全局钩子。

        之前一直挂着 WH_KEYBOARD_LL 全局钩子，会干扰 PyCharm 等程序（导致
        PyCharm 关不掉）。改为：在 VBE 里才挂载钩子；切到其它程序（PyCharm 等）
        立刻卸载钩子，并收起弹窗、释放 COM 资源。这样既消除冲突，又让“打开 VBE
        自动生效”成为可能（工具常驻，进 VBE 自动挂载）。
        """
        try:
            raw_focused = in_vbe_code_pane()
        except Exception:
            raw_focused = False
        # 若拿不到窗口 API，退化成“始终挂载”（保持可用，但会失去冲突修复）
        if _gui is None:
            raw_focused = True

        # 记录“不在 VBE”时的前台窗口句柄：用于区分「IDE 关掉、焦点被甩回 VBE」
        # （应抑制钩子）与「用户主动 Alt+Tab 回 VBE」（应正常挂载）。
        # 输入法候选窗口短暂抢焦点时也会记录，但它不会触发“进入 VBE”的转折，
        # 且转折用的去抖会把这类抖动过滤掉，不会误判。
        if not raw_focused and _gui is not None:
            try:
                state["non_vbe_hwnd"] = _gui.GetForegroundWindow()
            except Exception:
                state["non_vbe_hwnd"] = None

        # 焦点去抖：连续 FOCUS_DEBOUNCE 次读数一致才认定焦点真的变了。
        # 输入法的候选/组字窗口可能短暂抢走前台窗口；若立刻判定“离开 VBE”，
        # 会卸载钩子、收起弹窗（并曾清空轮询基线），中文输入就永远不触发。
        if raw_focused == state.get("raw_focused"):
            state["focus_streak"] = state.get("focus_streak", 0) + 1
        else:
            state["focus_streak"] = 1
            state["raw_focused"] = raw_focused
        if state["focus_streak"] >= FOCUS_DEBOUNCE:
            focused = raw_focused
        else:
            focused = state["vbe_focused"]   # 抖动期间维持原状态，不做任何切换

        # 焦点状态切换时，做一次资源清理 / 刷新
        if focused != state["vbe_focused"]:
            state["vbe_focused"] = focused
            # 刻意不再清空 last_snap：中文经输入法提交时焦点可能短暂离开，
            # 若清空基线，提交上去的中文会被当成“首次记录”而不触发——
            # 这正是「中文开头不提示」的成因之一。是否弹列表由
            # “同一行内文本变化”这条规则把关，不会因此乱弹。
            if focused:
                backend._cache = None            # 回到 VBE：强制刷新标识符
                state["released"] = False
                # v79：刚切回 VBE，把"按了键却没落到文档里"的计数清零 ——
                # 那些键是敲在别的程序里的，不该让我们以为输入法在组字。
                state["pending_keys"] = 0
                # 上一个“非 VBE”前台窗口已销毁 -> 说明它是被关掉、焦点被甩回
                # VBE 的（典型如关掉 VSCode/PyCharm）。若此刻立刻挂系统钩子，
                # 会拖慢那个仍在收尾（写设置、退子进程）的进程，表现为关窗口后
                # 2~3 秒才真正退出。给一段冷却期，期间不挂载钩子。
                prev = state.get("non_vbe_hwnd")
                if prev and _gui is not None:
                    try:
                        if not _gui.IsWindow(prev):
                            state["hook_suppress_until"] = (
                                time.time() + HOOK_SUPPRESS_SEC)
                    except Exception:
                        pass
            else:
                # 正在与弹窗交互（按住鼠标 / 拖滚动条）或鼠标正停在弹窗上时，
                # 不收起 —— 保证"选词前 UI 一直可见"（焦点抖动绝不误伤）。
                if completer.is_visible() and not popup.is_busy() \
                        and not popup.mouse_inside():
                    post(completer.hide)
                # 刻意【不】立刻 release：输入法抖动已由 FOCUS_DEBOUNCE 过滤，
                # 但在 VBE 与其它窗口之间来回切是很常见的操作，每次都清掉
                # VBE 缓存就要重新 GetActiveObject（对 Excel 加一次引用）。
                # 改为"真离开超过宽限期"才释放，由下面的 idle 检查执行。

        # 钩子（全局键盘 + 全局鼠标）【只在 VBE 里才是必需的】：
        # pynput 装的是系统级 WH_KEYBOARD_LL / WH_MOUSE_LL，只要挂着，
        # 全系统输入事件都要先经过我们的钩子线程；别的程序退出、拆窗口时
        # 一旦被它挡住就会卡死（即「PyCharm / VSCode 关不掉」）。
        # 因此挂载与否【只能】由「焦点是否在 VBE」决定，绝不能因为 released /
        # COM 退避等其它状态而重新挂载——之前那版就是在这里踩了坑。
        #
        # 额外保险：即便焦点去抖仍维持"在 VBE"，只要检测到前台窗口正在关闭
        # （正在退出的 IDE），也强制不挂钩。否则关闭 PyCharm/VSCode 时，焦点
        # 离开 VBE 的去抖窗口（约 600ms）内钩子仍挂着，IDE 退出会被挡死。
        closing = foreground_window_closing()
        # 钩子抑制冷却期：刚因“别的程序关掉”而落回 VBE 时，压住钩子挂载，
        # 等那个进程彻底退出再恢复（见上文进入 VBE 时设 hook_suppress_until）。
        hook_suppressed = bool(state.get("hook_suppress_until", 0.0) > time.time())
        mount_hooks, release_com = _reconcile_resources(
            state["vbe_focused"], state["released"], com_backoff_remaining(),
            state.get("last_focus_time", 0.0), time.time(), FOCUS_GRACE_SEC,
            closing=closing, suppress=hook_suppressed)
        if mount_hooks:
            state["last_focus_time"] = time.time()
            k = state.get("kbd")
            if k is None or not k.is_alive():
                try:
                    k = _make_kbd()
                    k.start()
                    state["kbd"] = k
                except Exception:
                    pass
            m = state.get("mouse")
            if m is None or not m.is_alive():
                try:
                    m = _make_mouse()
                    m.start()
                    state["mouse"] = m
                except Exception:
                    pass
        else:
            _stop_and_clear("kbd")
            _stop_and_clear("mouse")
            if release_com:
                state["released"] = True
                backend.release()
        root.after(300, reconcile)

    def poll_editor():
        """轮询编辑器内容变化来触发补全（问题「中文开头不提示」的根本解法）。

        之前靠键盘事件判断是否输入了标识符字符，但中文经输入法(IME)输入时，
        键盘钩子往往只收到 VK_PROCESSKEY 或空字符（尤其 on_release），
        导致中文永远不触发。改为直接读编辑器里的文本：只要当前行内容变了，
        就按「光标前是否为标识符字符」决定弹还是收——与输入法完全无关，
        中英文、粘贴、自动替换一律可靠。

        刻意不受 vbe_focused 门控：输入法的候选/组字窗口可能短暂抢走前台
        窗口，一旦门控，中文输入期间轮询就整个停摆。轮询只做 COM 读取
        （不是全局低层钩子），不会像键盘钩子那样干扰其它程序，可以常开。
        VBE 不可用时 snapshot() 返回 None，自然退化为空转。
        """
        try:
            # v74：提示窗签名【每轮都采样】，不再"只在候选窗可见时采"。
            #
            # v66 起这句是写在 if completer.is_visible() 里面的。于是候选窗
            # 不可见期间发生的任何变化（VBE 弹出/收起签名窗）都不会被记录 ——
            # 下次候选窗弹出时，签名可能恰好等于那个【陈旧】的缓存值，后面那句
            # `_sig != _popup_sig` 便永远看不见这次变化 ⇒ 候选窗一直压在签名窗
            # 上，直到用户自己挪光标。用户报的「instr 参数信息很长、与我们候选
            # 窗重叠」走的就是这条路：先在实参里打出候选（签名窗在 ⇒ 缓存记下
            # 它），随后删字重打（中途候选窗收起，缓存不动），签名窗在同一位置
            # 重新出现 —— 签名没变，"变化检测"瞎了。
            #
            # 代价可忽略：提示窗句柄是缓存好的，一轮只是几次只读 Win32 调用
            # （FindWindow + IsWindowVisible + GetWindowRect），量级远小于同
            # 一轮里 backend.snapshot() 那次 COM 读取。
            try:
                _sig = tuple(vbe_bridge.vbe_popup_rects())
                _sig_changed = _sig != state.get("_popup_sig")
                state["_popup_sig"] = _sig
            except Exception:
                _sig, _sig_changed = (), False
            # 【v76】状态轨迹：只在开日志时产生，且只在状态【真的变化】时写一行。
            #
            # 排查"VBE 的成员列表和我们候选窗冲突"这类问题时，下面三件事必须能一眼
            # 分开，否则只能靠猜：
            #   * VBE提示窗=()     -> 根本没探到提示窗（句柄缓存 / 类名 / 宿主 PID）
            #   * 手动=True        -> Ctrl+Space 的标记还挂着（v72 起它会让兜底跳过让位）
            #   * 两者几何明明相交 -> 窗不知道自己叠着（避让/自愈那一侧的问题）
            # 注意：取几何是额外几次只读调用，所以必须用 _LOG_ENABLED 包住。
            if _LOG_ENABLED:
                try:
                    _trace = (completer.is_visible(),
                              state.get("popup_manual"),
                              bool(completer.yield_pending),
                              completer.yield_given_up,
                              _sig,
                              popup.geometry_now() if completer.is_visible()
                              else None)
                    if _trace != state.get("_trace"):
                        state["_trace"] = _trace
                        _log("poll: trace 候选窗可见=%r 手动=%r 待定让位=%r"
                             " 已放弃让位=%r 候选窗几何=%r VBE提示窗=%r"
                             % (_trace[0], _trace[1], _trace[2], _trace[3],
                                _trace[5], _trace[4]))
                except Exception:
                    pass
            # v64：候选窗只允许活在"焦点真在代码窗格"的时候。焦点一旦落到属性
            # 窗口 / 工程窗口 / 窗体设计器（都不是写代码的地方），立刻收起 ——
            # 否则它会继续挂在那些工作区里，而且 Tab / ↑ / ↓ 还会被我们吞掉、
            # 作用到代码窗格上。判据拿不准（None）时不动手，保持旧行为。
            if completer.is_visible():
                try:
                    if vbe_bridge.vbe_code_pane_focused() is False:
                        _log("poll: 焦点不在代码窗格 -> 收起候选窗")
                        post(completer.hide)
                except Exception:
                    pass
                # v66：候选窗挂着的时候，VBE 随时可能弹出/收起【参数信息】窗
                # （形参签名）—— 它就在光标正下方，与我们默认的位置重叠，要让开。
                try:
                    # v72 兜底：冒出来的是 VBE 的【成员列表】时，要做的不是
                    # "挪开"而是"整个让位" —— 它和我们同质（也是一份候选
                    # 列表，也吃 Tab / ↑ / ↓ / Enter），叠在一起既遮挡又抢
                    # 键盘。语法判据（engine.vbe_list_expected）已经在 VBE
                    # 弹窗之前拦住了 `对象.` / `As` / `New` / With 块的
                    # `.成员`；这里兜的是它预测不到的路径（输入法、粘贴、
                    # 超大模块，以及 VBE 自己在别处弹的列表）。判据只有一条：
                    # **VBE 的列表真的画在屏幕上了**。
                    #
                    # 【v76】不再区分"手动唤出"。用户口径（原话）："VBE 的弹窗和
                    # 我们项目弹窗冲突时，屏蔽我们项目弹窗……直接消失让位。"
                    #
                    # v72 当初加 popup_manual 是为了防"按了 Ctrl+Space 却没反应"，
                    # 而那个担心只在【VBE 列表不在】时才成立 —— 那种情况下下面这个
                    # any(...) 本来就是 False，压根走不到 hide。所以去掉这个前置
                    # 条件不会把 v72 的问题带回来；反过来，VBE 的成员列表本身就是
                    # 一份候选列表，它画在屏幕上时用户的目的已经达到，这时候还挂着
                    # 我们的窗（topmost）只会把它整个盖住。
                    if any(c in vbe_bridge.VBE_YIELD_CLASSES
                           for (c, _l, _t, _r, _b) in _sig):
                        _log("poll: VBE 成员列表出现 %r -> 让位 (手动=%r)"
                             % (_sig, state.get("popup_manual")))
                        post(completer.hide)
                    else:
                        # 两条判据任一命中就重摆（都不命中则什么都不做 ——
                        # 不会周期性重设 geometry 造成闪烁）：
                        #   * 签名变化  —— 事件驱动：签名窗刚出现 / 收起 / 换行；
                        #   * 实测仍重叠 —— 状态自愈（v74）：兜住"事件漏了"的
                        #     情况。只做几次矩形比较，拿不到几何一律 False。
                        try:
                            _clash = popup.clashes_with_popup()
                        except Exception:
                            _clash = False
                        if _sig_changed or _clash:
                            _log("poll: 重摆候选窗"
                                 " (提示窗=%r 签名变化=%r 实测重叠=%r)"
                                 % (_sig, _sig_changed, _clash))
                            # 注意调的是 UI（popup）而不是 completer：重摆是画布
                            # 的事，引擎那层没有这个方法。
                            post(popup.reposition)
                except Exception:
                    pass
            # v73：待定让位的【宽限确认】。
            #
            # 我们刚在"VBE 会弹成员列表"的位置让了位（`标识符.`、`As`/`New`
            # 之后、With 块的 `.成员`），现在看 VBE 到底弹没弹：
            #   * 弹了 -> 认这个让位，什么都不做；
            #   * 没弹 -> 把候选窗补回来。
            # 为什么必须有这一步：VBE 的自动列出成员要先解析出对象的类型，
            # With 块里解析不了（地址写错、模块有待编译错误）它一个窗都不弹 ——
            # 那时"让位"就等于"什么都不给"，正是用户报的"在 With 块里输入
            # .Size 什么都不弹了"。
            #
            # 只探测 Win32 窗口可见性（只读），且只在宽限到期那一刻探一次。
            try:
                _yp = completer.yield_pending
                if _yp and time.time() >= _yp[0]:
                    # 复用本轮开头采的那份签名 —— 同一 tick 内两处判据必须看
                    # 同一时刻的窗口状态，别再探一次（v74）。
                    _ml = any(c in vbe_bridge.VBE_YIELD_CLASSES
                              for (c, _l, _t, _r, _b) in _sig)
                    if completer.confirm_yield(_ml):
                        _log("poll: 让位宽限到 -> VBE 列表在（或已离开该处），保持让位")
                    else:
                        _log("poll: 让位宽限到 -> VBE 没弹成员列表，自己补上候选")
                        # 这次窗是我们自己补的，VBE 的列表一冒出来照样该让位
                        state["popup_manual"] = False
                        post(lambda: completer.trigger(False))
            except Exception:
                pass
            snap = backend.snapshot() if _com_allowed() else None
            if snap is not None and len(snap) > 3:
                # v79：给"回车自动缩进"留一份【每轮都刷新】的光标上下文
                # (时刻, 行号, 行文本, 光标列)。刻意不复用 last_snap —— 那份只在
                # 文本变化时才更新，行号与列可能早就过期（移动光标不算变化），
                # 拿它判"光标是否在行尾"会误判成"在行尾"而抢走一次拆行。
                state["cur_ctx"] = (time.time(), snap[0], snap[1], snap[3])
            # 诊断心跳（只在 VBECOMPLETE_LOG=1 时产生，约每 3 秒一行）。
            # 排查"某处输入什么都不弹"时，它一眼分清是【没读到快照】（COM 读到
            # None：多行选区 / 退避 / 不可用）、【焦点不在代码窗格】、还是
            # 【读了却没触发】。
            if time.time() >= state.get("_hb_next", 0.0):
                state["_hb_next"] = time.time() + 3.0
                _log("poll: heartbeat snap=%r focused=%r focus_cls=%r"
                     " popup=%r popup_cls=%r yield=%r"
                     " backoff=%.2f visible=%r kbd=%r"
                     % (None if snap is None else
                        (snap[0], (snap[2] if len(snap) > 2 else None),
                         (snap[1] or "")[-45:]),
                        vbe_bridge.vbe_code_pane_focused(), vbe_focus_class(),
                        vbe_bridge.vbe_popup_visible(),
                        vbe_bridge.vbe_popup_info(),
                        vbe_bridge.vbe_yield_visible(),
                        com_backoff_remaining(),
                        completer.is_visible(),
                        state.get("kbd") is not None))
            if snap is not None:
                last = state.get("last_snap")
                if last is None:
                    state["last_snap"] = snap      # 首次只记录，不触发
                    state["pending_keys"] = 0      # v79：文档状态未知 -> 计数清零
                elif last[:3] != snap[:3]:
                    # v79：文档真的变了 -> "按了键却还没落到文档里"的计数清零
                    # （那些键确实写进去了，输入法没有在组字）。
                    state["pending_keys"] = 0
                    # 只比前三个元素（行号 / 行文本 / 模块名）：第 4 个是 v79 加
                    # 的【光标列】，光标的移动不算"内容变化"—— 否则在代码里点来
                    # 点去都会被弹窗打扰（这条一直以来的行为必须守住）。
                    state["last_snap"] = snap
                    # v64：换模块（Ctrl+Tab 切代码窗、点工程树换模块）会让
                    # "同一行 + 文本不同"看起来像一次编辑，其实一个字都没敲
                    # —— 那不是输入的上下文，收起挂着的候选窗、也别弹新的。
                    if _poll_mod_switch(last, snap):
                        _log("poll: 换了模块 -> 收起（不是输入）")
                        post(completer.hide)
                    # 必须是「同一行内的文本变化」才算真的输入了内容。
                    # 单纯换到别的行（鼠标点击、方向键换行）不应弹列表，
                    # 否则在代码里点来点去会被弹窗打扰。
                    elif last[0] == snap[0]:
                        # 文本变了 -> 让标识符缓存尽快失效，保证"刚声明的
                        # 变量"也能进候选。限流：最快 ID_REFRESH_MIN_SEC
                        # 重解析一次，避免每敲一个键都全量解析所有模块。
                        if (time.time() - getattr(backend, "_cache_time", 0.0)
                                > ID_REFRESH_MIN_SEC):
                            backend.invalidate_identifiers()
                        # v52：这一下敲的是空格/标点（行【变长】且插进来的
                        # 这段不含标识符字符）-> 直接收起，别去 trigger。
                        # 否则当它敲在某个标识符【前面】时，右边那个标识符会被
                        # 当成"正在输入的词"拿来补全 —— 形参首字符前打空格把
                        # 形参自己提示出来，正是这么来的（详见 engine 里
                        # typed_separator 的说明）。删字的路径是"变短"，不受影响。
                        if _poll_action(
                                engine.typed_separator(last[1],
                                                       snap[1])) == "hide":
                            _log("poll: 敲了空格/标点 -> 收起")
                            post(completer.hide)
                        else:
                            # True: 校验光标前是否标识符字符（不在拼标识符则收起）
                            _fst = vbe_bridge.vbe_code_pane_focused()
                            if _fst is False:
                                # v64：焦点不在代码窗格（属性窗口 / 工程窗口 /
                                # 窗体设计器）—— 那里的"文本变化"不是用户在拼
                                # 标识符，别在非写代码的工作区弹提示。
                                _log("poll: 焦点不在代码窗格 -> 不触发"
                                     " (focused=%r focus_cls=%r popup=%r"
                                     " popup_cls=%r)"
                                     % (_fst, vbe_focus_class(),
                                        vbe_bridge.vbe_popup_visible(),
                                        vbe_bridge.vbe_popup_info()))
                                post(completer.hide)
                            else:
                                _log("poll: line %s changed -> trigger"
                                     " (focused=%r focus_cls=%r popup=%r"
                                     " popup_cls=%r)"
                                     % (snap[0], _fst, vbe_focus_class(),
                                        vbe_bridge.vbe_popup_visible(),
                                        vbe_bridge.vbe_popup_info()))
                                post_trigger()
        except Exception:
            pass
        # 已离开 VBE（超过宽限期）、或 COM 正处于失败退避（宿主多半在关闭）
        # -> 降频轮询，别在关文件/退 Excel 时猛打 COM。
        idle = (com_backoff_remaining() > 0
                or (not state.get("vbe_focused")
                    and time.time() - state.get("last_focus_time", 0.0)
                    > FOCUS_GRACE_SEC))
        root.after(POLL_INTERVAL_IDLE_MS if idle else POLL_INTERVAL_MS,
                   poll_editor)

    root.after(10, drain_actions)
    root.after(0, reconcile)
    root.after(0, poll_editor)

    try:
        root.mainloop()
    except Exception:
        pass


if __name__ == "__main__":
    main()
