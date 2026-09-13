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
  注（v49）：数字键 1~9 不再用于选词，候选也不再显示序号 —— 弹窗开着时
  照常输入数字，定义 `s1` / `arr17` 这类名字不会被列表抢走。
"""

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
from ui import Popup, MAX_VISIBLE_ROWS, caret_screen_rect
from log import log as _log, log_boot as _log_boot


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

NAV_VKS = (VK_UP, VK_DOWN)


def _is_newline_shortcut(vk, shift_down):
    """Shift+Enter 是否应触发"在当前行下方新起一行"（纯判断，便于单测）。

    只认 Enter：Shift+Tab（反缩进）等一律不管。小键盘回车在 VBE 里同样是
    VK_RETURN，因此与主键盘一视同仁。
    """
    return vk == VK_RETURN and bool(shift_down)


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
    """判断当前前台窗口是否是 VBE 代码窗。"""
    if _gui is None:
        return False
    try:
        hwnd = _gui.GetForegroundWindow()
        title = _gui.GetWindowText(hwnd)
        return "Microsoft Visual Basic" in title
    except Exception:
        return False


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


def _vbe_list_up():
    """VBE 自己是不是正在显示「自动列出成员」列表（是则我们让位）。

    拿不准 / 出任何异常一律返回 False —— 宁可两个列表同时出现，也不能让
    工具整个哑掉。判据见 vbe_bridge.vbe_list_visible。
    """
    try:
        return vbe_bridge.vbe_list_visible(caret_screen_rect())
    except Exception:
        return False


def _poll_action(typed_sep, vbe_list_up):
    """纯函数：本轮轮询该做什么 —— 'hide' 还是 'trigger'。

    - typed_sep  ：这一下敲的是空格 / 标点（v52，别拿右边的词来补全）；
    - vbe_list_up：VBE 自己的提示列表已经弹出来了（v54）。

    后者与前者同等对待：直接收起、不去 trigger。两个列表同时出现会互相
    遮挡、还抢键盘，VBE 已经给了就让它给。
    """
    if typed_sep or vbe_list_up:
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
        "last_snap": None,    # 上次轮询到的 (行号, 行文本)，用于检测内容变化
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
        """
        try:
            completer.hide()
        except Exception:
            pass
        try:
            backend.new_line_below()
        except Exception:
            pass

    def drain_actions():
        """主线程循环消费动作队列。"""
        try:
            while True:
                fn, args = action_queue.get_nowait()
                try:
                    fn(*args)
                except Exception:
                    pass
        except queue.Empty:
            pass
        root.after(10, drain_actions)

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
            vk = data.vkCode

            if msg in (WM_KEYDOWN, WM_SYSKEYDOWN):
                # Shift+Enter = 在当前行下方新起一行（缩进对齐上一行）。
                # 与"弹窗是否可见"无关：弹窗开着也照样接管（顺带收起弹窗）。
                # 只在【焦点在 VBE 代码窗】且【COM 没在失败退避】时才吞键；
                # 否则放行，让 VBE 按原生行为处理（原生 Shift+Enter == 回车），
                # 绝不"吞了按键却什么都没发生"。
                if (_is_newline_shortcut(vk, _shift_down())
                        and com_backoff_remaining() <= 0
                        and in_vbe_code_pane()):
                    action = (_new_line_here, ())
                    suppress = True
                elif completer.is_visible():
                    if not in_vbe_code_pane():
                        # 焦点已离开 VBE：收起弹窗，按键照常放行
                        post(completer.hide)
                    elif vk == VK_ESCAPE:
                        action = (completer.hide, ())
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
                # v54：Ctrl+Space 也是 VBE 唤列表的键，它已经弹了就别再弹
                if in_vbe_code_pane() and not _vbe_list_up():
                    post(completer.trigger)
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
            snap = backend.snapshot() if _com_allowed() else None
            if snap is not None:
                # v54：VBE 自己也在弹列表 -> 让位（两个列表同时出现会
                # 互相遮挡、抢键盘）。文本变没变都要查：按 Ctrl+J 手动
                # 唤出 VBE 列表时文本是不变的，只有这里能拦住。
                vbe_up = _vbe_list_up()
                if vbe_up and completer.is_visible():
                    _log("poll: VBE 自带列表已出现 -> 收起")
                    post(completer.hide)
                last = state.get("last_snap")
                if last is None:
                    state["last_snap"] = snap      # 首次只记录，不触发
                elif last != snap:
                    state["last_snap"] = snap
                    # 必须是「同一行内的文本变化」才算真的输入了内容。
                    # 单纯换到别的行（鼠标点击、方向键换行）不应弹列表，
                    # 否则在代码里点来点去会被弹窗打扰。
                    if last[0] == snap[0]:
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
                                engine.typed_separator(last[1], snap[1]),
                                vbe_up) == "hide":
                            _log("poll: %s -> 收起"
                                 % ("VBE 自带列表已出现" if vbe_up
                                    else "敲了空格/标点"))
                            post(completer.hide)
                        else:
                            # True: 校验光标前是否标识符字符（不在拼标识符则收起）
                            _log("poll: line %s changed -> trigger" % (snap[0],))
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
