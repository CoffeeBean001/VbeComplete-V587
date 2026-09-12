"""
VBE 后端：通过 pywin32 与正在运行的 Excel VBE 交互。

只在这一层依赖 pywin32 / Excel，引擎层不依赖，方便单测。
需要：Excel 已开启「信任对 VBA 工程对象模型的访问」。
"""

import os
import re
import time
import unicodedata

from log import log as _log

import parser as vba_parser
from engine import replace_word

_CACHE_TTL = 2.0  # 标识符缓存刷新间隔（秒）

# 未声明即使用的"隐式变量"（不写 Option Explicit 时的用法）是否纳入提示。
# 关闭后行为回到旧版：只提示 Dim/Const/Sub 等显式声明出来的名字。
ENABLE_IMPLICIT_IDENTIFIERS = True
# 隐式变量的作用域粒度：
#   "module" —— 本模块任意位置都能提示（最宽松，但会跨过程泄漏同名变量）
#   "proc"   —— 按"首次出现所在过程"限定作用域（更贴近 VBA 语义，避免跨过程泄漏）
IMPLICIT_SCOPE = "proc"

# 合法标识符（含中文）的模块名/窗体名/类名，用于把组件名纳入候选
_RE_PLAIN_IDENT = re.compile(r"^[^\W\d]\w*$")

# 跨工程作用域：
#   False（默认）—— 只收集【当前活动工程】里的标识符。
#     VBE 里常常同时开着好几个工程（多个工作簿、加载宏 .xlam、个人宏工作簿
#     Personal.xlsb）。除了显式"工具-引用"之外，一个工程里的窗体名、函数名、
#     全局变量名在另一个工程里根本不可见，提示它们纯属噪音（典型症状：在当前
#     工程里敲一个名字，冒出来的却是另一个工作簿里的 UserForm1 / 全局变量）。
#   True —— 恢复旧行为：收集 VBE 里所有工程的标识符。
ENABLE_CROSS_PROJECT = False

def _co_free():
    """释放 pywin32 缓存的 COM 代理（仅在显式开启时才会真正调用）。

    说明：CoFreeUnusedLibraries 只释放"当前无人引用"的 COM 库，**并不能**释放
    我们持有的 Excel/VBE 引用，对"让 Excel 退出"其实没有帮助；反而在 Excel
    正在关闭时反复调用会触发 COM 库的卸载/重载，是"关文件卡几秒、关完又被
    拉起"的嫌疑点之一。因此默认不再自动调用，需要排障时用环境变量
    VBECOMPLETE_COFREE=1 打开。pythoncom 延迟导入，避免无 pywin32 环境报错。
    """
    import os
    if os.environ.get("VBECOMPLETE_COFREE", "0").strip() != "1":
        return
    try:
        import pythoncom
        pythoncom.CoFreeUnusedLibraries()
    except Exception:
        pass


# ---------------------------------------------------------------------------
# VBE 行列坐标处理
#
# 现象：按 Tab 确认后，光标偶尔落在"词的中间"而不是词尾。根因是行内含
# 制表符（Tab 缩进）时，字符位置 与 VBE 的列 可能不一致：
#   - 若 VBE 的列=字符位置（Tab 计 1 列）：我们按 Python 字符数定位应无偏差；
#   - 若 VBE 的列=显示位置（Tab 展开为多格）：按字符数 SetSelection 会少算
#     若干格，光标就落在词中间。
#
# 处理（双保险，无需预知 VBE 语义）：
#   1. _vbe_measure：向"远超行尾"的列 SetSelection，读回停靠列 E 来探测
#      该行列语义与 Tab 宽度（字符语义会被钳制到 len+1，显示语义会钳制到
#      显示宽度+1，二者数值不同，可区分）。
#   2. 写回时若行内含 Tab，先把整行的 Tab 展开为空格再 ReplaceLine ——
#      该行此后不含 Tab，字符位置与列恒等，光标定位在任意语义下都精确。
#      唯一的副作用：被补全的那一行的缩进从 Tab 变成等宽空格（纯视觉，
#      VBA 语义不变，视觉宽度也不变）。
# ---------------------------------------------------------------------------

def _is_ident_tail(ch):
    """标识符尾字符判定（Unicode 感知：中文等字母/数字/下划线都算）。"""
    return ch.isalnum() or ch == "_"


def _is_wide(ch):
    """是否会被等宽编辑器当成 2 格宽的全角字符（CJK/全角）。"""
    return unicodedata.east_asian_width(ch) in ("W", "F")


def _disp_width(line_text, tabw, wide2, upto=None):
    """把 line_text[:upto] 按显示宽度展开（Tab 到下一个 tab stop）。"""
    w = 0
    end = len(line_text) if upto is None else min(upto, len(line_text))
    for ch in line_text[:end]:
        if ch == "\t":
            w += tabw - (w % tabw)
        else:
            w += 2 if (wide2 and _is_wide(ch)) else 1
    return w


def _col_to_char_index(line_text, disp_col, tabw, wide2):
    """显示列(1-based) -> 字符下标(0-based，光标停在该字符前)。

    若 disp_col 落在行尾之后则返回 len(line_text)。
    """
    off = 0
    for i, ch in enumerate(line_text):
        if disp_col == off + 1:
            return i
        if ch == "\t":
            off += tabw - (off % tabw)
        else:
            off += 2 if (wide2 and _is_wide(ch)) else 1
    return len(line_text)


def _fit_semantics(line_text, end_col):
    """用"某个已知的行尾列 end_col"反推 (sem, tabw, wide2)。

    end_col 必须来自 VBE 自己报告的列（光标停靠列或行尾列）。
      end_col <= len+1 -> 列没有超过字符数，只能是"字符列"语义；
      否则            -> 必然是"显示列"语义，用候选 (tabw, wide2) 拟合。
    """
    n = len(line_text)
    if end_col <= n + 1:
        return "char", 4, False
    # 候选顺序：优先最常见的 tabw=4。行内无 Tab 时 tabw 不影响结果，
    # 若不把 4 放前面会先命中 2，导致缓存里的 tabw 看起来"很怪"。
    for wide2 in (False, True):
        for tabw in (4, 8, 2, 3):
            try:
                if _disp_width(line_text, tabw, wide2) + 1 == end_col:
                    return "disp", tabw, wide2
            except Exception:
                continue
    # 兜底：假设只有 Tab 造成展开
    ntabs = line_text.count("\t")
    if ntabs > 0:
        extra = (end_col - 1) - n
        if extra % ntabs == 0:
            t = extra // ntabs + 1
            if 1 <= t <= 32:
                return "disp", t, False
    # 兜底：行内含全角字符却无法用候选参数复现时，优先假定"全角占 2 列"——
    # 比当成 1 列更贴近 VBE 实际行为，避免光标落在中文变量名中间。
    if any(_is_wide(c) for c in line_text):
        return "disp", 4, True
    return "disp", 4, False


def _detect_semantics_passive(line_text, ec):
    """【完全不移动光标】地判定列语义；判不出来返回 None。

    原理：字符列语义下 VBE 报告的列不可能超过 len+1。若 ec 已经 > len+1，
    则它必然是显示列，直接用它拟合 (tabw, wide2) 即可。
    打字时光标几乎总在行尾，这条路径能覆盖绝大多数真实输入，
    因此可以避免为了探测而移动光标（某些编辑器会把光标还原错/还原延迟，
    表现为"光标自己跳到代码中间"）。
    """
    if not line_text:
        return None
    if ec <= len(line_text) + 1:
        return None
    return _fit_semantics(line_text, ec)


# 列语义是"编辑器级"属性（不随行变化），因此探测一次即可长期复用。
_sem_cache = {"info": None, "probe_broken": False}


def _probe_semantics(cm, cp, line_no, line_text):
    """安全地探测列语义：移动光标探测后必定还原，并校验还原结果。

    某些编辑器（如 WPS 的 VBA 兼容编辑器）对 SetSelection 的处理与 Office
    不同：越界列可能不被钳制到行尾，或还原是异步的——探测列若没被正确还原，
    就会出现"光标自己跳到代码中间"。因此这里：
      1) 有缓存就直接返回，整个会话最多探测一次；
      2) try/finally 保证一定执行还原，且还原后读回校验（不符再还原一次）；
      3) 校验始终失败 -> 标记 probe_broken，此后永久禁用探测，退回被动判定。
    """
    if _sem_cache["probe_broken"]:
        return _sem_cache["info"]
    cached = _sem_cache["info"]
    if cached is not None:
        return cached
    try:
        sl, sc, el, ec = cp.GetSelection()
    except Exception:
        return None
    info = None
    try:
        info = _vbe_measure(cm, cp, line_no, line_text)
    except Exception:
        info = None
    finally:
        # 无论探测成功与否，都必须把光标放回原处
        restored = False
        for _ in range(3):
            try:
                cp.SetSelection(sl, sc, el, ec)
                cur = tuple(cp.GetSelection())
                if cur == (sl, sc, el, ec):
                    restored = True
                    break
            except Exception:
                break
        if not restored:
            # 该编辑器无法可靠还原探测列：以后不再探测，避免光标乱跳
            _sem_cache["probe_broken"] = True
    if info is not None:
        _sem_cache["info"] = info
    return info


def _vbe_measure(cm, cp, line_no, line_text):
    """探测 VBE 对当前行的"列"语义与 Tab 宽度（不恢复光标，调用方负责）。

    返回 (sem, tabw, wide2)：
      sem == 'char' -> VBE 列 == 字符位置（Tab 计 1），直接按字符数用列；
      sem == 'disp' -> VBE 列 == 显示列（Tab 展开），需按 (tabw, wide2) 换算。
    探测法：把光标设到第 len+100 列（远超行尾），读回实际停靠列 E，再拟合。

    注意：本函数会【移动光标】。除 apply_completion（随后本来就要设置光标）
    之外，请一律改用 _probe_semantics，它会自动还原并校验。
    """
    n = len(line_text)
    try:
        cp.SetSelection(line_no, n + 100, line_no, n + 100)
        _sl, _sc, _el, ec = cp.GetSelection()
    except Exception:
        return "char", 4, False
    return _fit_semantics(line_text, ec)


def _find_completion_end(text, completion, near0, fallback0):
    """在(可能被 VBE 规范化过的)行文本里定位补全词 completion 的结尾字符下标。

    要求找到的 occurrence 前后都不是标识符字符（整词匹配）；
    若行内多处出现同名，取起点最接近 near0 的那个。
    找不到则回退 fallback0（裁剪到行尾）。
    """
    L = len(completion)
    best = None
    i = 0
    while True:
        i = text.find(completion, i)
        if i < 0:
            break
        ok_l = (i == 0) or (not _is_ident_tail(text[i - 1]))
        ok_r = (i + L >= len(text)) or (not _is_ident_tail(text[i + L]))
        if ok_l and ok_r:
            if best is None or abs(i - near0) < abs(best - near0):
                best = i
        i += L
    if best is not None:
        return best + L
    return min(fallback0, len(text))


def _skip_into_parens(actual, end0, orig_line, orig_caret_col):
    """补全词之后若紧跟一个「新出现」的左括号，把光标下标移到它右侧。

    VBE 会自动为过程声明补括号：把 `Function gCalc` 写回后，行里会冒出一对
    `()`。光标停在词尾就等于停在左括号【左侧】，用户还得手动右移一格才能写
    参数。这里检测"原行光标后本来不是 `(`、写回后却出现 `(`"，就把光标送进
    括号里；像 `x = gCalc(1)` 这种括号本来就有的调用点则保持不变。

    actual  : 写回后读回的实际行文本
    end0    : 补全词结尾的字符下标（0-based，即光标应停的位置）
    orig_line / orig_caret_col : 写回前的行文本与光标列（1-based 字符列）
    """
    try:
        i = int(orig_caret_col) - 1
        orig_after = orig_line[i:i + 1] if 0 <= i < len(orig_line) else ""
        if orig_after == "(":
            return end0
        j = end0
        n = len(actual)
        while j < n and actual[j] == " ":
            j += 1
        if j < n and actual[j] == "(":
            return j + 1
    except Exception:
        pass
    return end0


def _get_vbe():
    import win32com.client
    xl = win32com.client.GetActiveObject("Excel.Application")
    return xl.VBE


def _caret_position(vbe):
    """只读地取当前光标位置，返回 (module_name, (line_no, char_col)) 或 (None, None)。

    只调用 GetSelection、绝不 SetSelection，因此不会移动光标
    （WPS 等编辑器对 SetSelection 还原不可靠，移动光标会导致光标乱跳）。

    列语义优先用被动判定，其次用会话缓存；两者都拿不到时，仅当该行是
    "纯 ASCII 且无 Tab"（此时 VBE 列必然等于字符列）才直接换算，否则放弃——
    宁可不排除，也不要因为列换算不准而抹掉错误的词。
    """
    try:
        cp = vbe.ActiveCodePane
        if cp is None:
            return None, None
        cm = cp.CodeModule
        _sl, _sc, el, ec = cp.GetSelection()
        if el <= 0:
            return None, None
        line_text = cm.Lines(el, 1)
        info = _detect_semantics_passive(line_text, ec)
        if info is None:
            info = _sem_cache.get("info")
        if info is not None:
            sem, tabw, wide2 = info
            if sem == "disp":
                col = _col_to_char_index(line_text, ec, tabw, wide2) + 1
            else:
                col = min(ec, len(line_text) + 1)
        elif "\t" not in line_text and not any(_is_wide(c) for c in line_text):
            col = min(ec, len(line_text) + 1)
        else:
            return None, None

        # 模块名必须与收集时的 comp.Name 同源，否则比对失败会导致 caret 被丢弃、
        # 光标处的词被当成隐式变量收录（自我提示"打什么提示什么"）。
        # CodeModule.Name 在部分宿主里与 VBComponent.Name 不一致，故优先取后者。
        mod_name = ""
        try:
            mod_name = str(cp.CodeModule.Parent.Name)
        except Exception:
            pass
        if not mod_name:
            try:
                mod_name = str(cm.Name)
            except Exception:
                mod_name = ""
        return mod_name, (el, col)
    except Exception:
        return None, None


def _active_vb_project(vbe):
    """返回当前活动工程（VBProject 对象）；取不到返回 None。

    VBE.ActiveVBProject 直接对应"当前代码窗所属工程"；个别宿主/WPS 上该属性
    可能不可用，则退到 CodeModule.Parent(VBComponent).Collection.Parent 这条
    链路。两条都拿不到时返回 None（调用方退回"全部工程"的旧行为）。
    """
    try:
        p = vbe.ActiveVBProject
        if p is not None:
            return p
    except Exception:
        pass
    try:
        cp = vbe.ActiveCodePane
        if cp is None:
            return None
        return cp.CodeModule.Parent.Collection.Parent
    except Exception:
        return None


def _project_name(proj):
    try:
        return str(proj.Name)
    except Exception:
        return ""


def _components_of(proj):
    try:
        out = []
        for comp in proj.VBComponents:
            out.append(comp)
        return out
    except Exception:
        return []


def _collect_components(vbe, caret_mod=None):
    """返回本次应参与补全的组件列表（默认只含【活动工程】的组件）。

    两道保险，宁可退回"全部工程"也不能让补全整个失效：
      1) 活动工程读不到组件（工程被密码保护/锁定）；
      2) 已知光标所在模块名，却不在活动工程的组件里——说明取到的"活动工程"
         不是真正正在编辑的那个，退回全工程（旧行为）。
    """
    if not ENABLE_CROSS_PROJECT:
        proj = _active_vb_project(vbe)
        if proj is not None:
            comps = _components_of(proj)
            if comps:
                if caret_mod:
                    names = []
                    for c in comps:
                        try:
                            names.append(str(c.Name).lower())
                        except Exception:
                            names.append("")
                    if str(caret_mod).lower() not in names:
                        _log("collect: 活动工程 %r 不含模块 %r，退回全工程"
                             % (_project_name(proj), caret_mod))
                        proj = None
                        comps = []
            if comps:
                _log("collect: 仅收集活动工程 %r（%d 个组件）"
                     % (_project_name(proj), len(comps)))
                return comps
    comps = []
    try:
        for p in vbe.VBProjects:
            comps.extend(_components_of(p))
    except Exception:
        return []
    _log("collect: 收集全部工程（%d 个组件）" % len(comps))
    return comps


# ---------------------------------------------------------------------------
# Excel / VBE 的 COM 代理生命周期  ——  「关文件卡 3~10 秒 + 关完又启动 Office」
# 的病根就在这里
#
# 症状的关键线索：**改过 VBA 代码才卡，只改工作表内容不卡**。
#
# 机理：只要我们的进程持有 Excel.Application（哪怕只是间接经由 VBE 对象），
# Excel 就退不干净 —— 主窗口关了但 EXCEL.EXE 仍驻留后台，之后再打一次 COM
# 就会把它「唤醒」重新弹窗，看起来就是"关完又启动 Office"。
# 而「改过 VBA 才慢」正是因为 VBAProject 变脏后，Excel 退出时要把 VBA 工程
# 写回文件，这一步要与 VBA 环境交互，被外部 COM 引用钉住就会拖上好几秒；
# 没改 VBA 就不必写回，直接退出所以完全无感。
#
# 第一轮之所以没问题，是因为当时每次都 `_get_vbe()` 新建、用完即弃 —— 局部
# 变量在函数返回时引用计数归零，COM 引用立刻释放。后来改成模块级长缓存，
# 引用从此常驻，病就来了。
#
# 对策（三道，缺一不可）：
#   1) 代理设 TTL，到期主动丢弃 —— 回应用完即弃的语义，但不牺牲打字性能；
#   2) 宿主窗口（Excel 主窗口 / VBE 窗口）消失后，绝不再碰 COM；
#   3) 连续失败按档退避，不在 Excel 关闭窗口期猛打 COM。
# ---------------------------------------------------------------------------

# 缓存的 VBE/Excel COM 代理最长存活时间（秒）。0 或负数 = 彻底不缓存。
# 想回到"每次新建、用完即弃"（最保守）就设成 0。
try:
    VBE_PROXY_TTL = float(
        os.environ.get("VBECOMPLETE_PROXY_TTL", "1.0").strip() or "1.0")
except Exception:
    VBE_PROXY_TTL = 1.0

# 宿主窗口探测门。类名对不上的宿主（如 WPS）可能误判，用环境变量关掉：
#   VBECOMPLETE_NO_WINDOW_GATE=1
_WINDOW_GATE_DISABLED = (
    os.environ.get("VBECOMPLETE_NO_WINDOW_GATE", "0").strip() == "1")

_vbe_cache = {"obj": None, "at": 0.0}
_wnd_state = {"seen": False, "miss": 0}

# Excel 主窗口 / VBE 主窗口的类名
_HOST_WINDOW_CLASSES = ("XLMAIN", "wndclass_desked_gsk")


def _probe_host_window():
    """纯 Win32 探测宿主窗口是否还在（完全不碰 COM）。"""
    import ctypes
    u32 = ctypes.windll.user32
    fn = u32.FindWindowW
    fn.restype = ctypes.c_void_p
    fn.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p]
    for cls in _HOST_WINDOW_CLASSES:
        if fn(cls, None):
            return True
    return False


def _host_window_alive():
    """宿主窗口是否还在（可 monkeypatch，测试会替换 _probe_host_window）。

    刻意带"从未见过就别拦"的兜底：若这台机器上从来没探测到过 Excel/VBE
    窗口，很可能是类名不匹配的宿主（WPS 等），此时盲目拦截会让工具整个
    失效。只有"曾经见过、现在连续两次都看不到"才认定宿主真的退出了。
    """
    if _WINDOW_GATE_DISABLED:
        return True
    try:
        alive = _probe_host_window()
    except Exception:
        return True
    if alive:
        if _wnd_state["miss"] >= 2:
            com_reset()               # 宿主回来了 -> 立刻解除退避
        _wnd_state["seen"] = True
        _wnd_state["miss"] = 0
        return True
    if not _wnd_state["seen"]:
        return True                   # 从未见过：不拦，避免误伤其它宿主
    _wnd_state["miss"] += 1
    return _wnd_state["miss"] < 2     # 连续 2 次才认定真的没了


def _release_vbe_proxy(collect=False):
    """丢弃缓存的 VBE/Excel COM 代理 —— 这是让 Excel 能真正退出的关键。

    collect=True 时额外跑一次 gc.collect()：刚写过代码、或确认宿主已退出时
    用，确保 COM 包装对象被立刻回收而不是等 GC 想起来。
    """
    _vbe_cache["obj"] = None
    _vbe_cache["at"] = 0.0
    if collect:
        try:
            import gc
            gc.collect()
        except Exception:
            pass


# COM 调用失败退避（秒）。
#
# Excel 正在关闭 / 正在重启时，每次 GetActiveObject 都可能被 COM 拖住数秒
# （等服务器响应、甚至触发重新激活）。若仍按 100ms 猛冲，表现为：
#   - 关文件要卡 3~10 秒；
#   - 文件关完之后 Excel/Office 又被 COM 唤醒启动一次。
# 因此连续失败时按档退避（比早期版本更狠），成功后立刻恢复。
_COM_BACKOFF = (1.0, 2.0, 5.0, 10.0, 30.0)
_com_state = {"fails": 0, "until": 0.0}


def com_backoff_remaining():
    """还需退避多少秒（0 表示可以正常访问 COM）。"""
    return max(0.0, _com_state["until"] - time.time())


def com_reset():
    """重置退避状态（仅供测试 / 手动恢复使用）。"""
    _com_state["fails"] = 0
    _com_state["until"] = 0.0


def _com_ok():
    _com_state["fails"] = 0
    _com_state["until"] = 0.0


def _com_fail():
    _com_state["fails"] += 1
    i = min(_com_state["fails"] - 1, len(_COM_BACKOFF) - 1)
    _com_state["until"] = time.time() + _COM_BACKOFF[i]
    _log("com: 第 %d 次失败，退避 %.1fs"
         % (_com_state["fails"], _COM_BACKOFF[i]))


def _get_vbe_cached():
    """取 VBE 对象（**短命**代理，TTL 到期主动放手）。

    注意与"长缓存"的取舍：复用代理确实省一次 GetActiveObject，但代价是我们
    的进程会一直持有 Excel.Application —— 而只要持有它，Excel 就退不干净
    （关完又启动 Office），且 VBAProject 变脏时的写回会被拖成 3~10 秒。

    所以这里只做**短缓存**：TTL（默认 1 秒）一到就丢弃，让 Excel 随时有
    机会干净退出；重建成本只是一次 ROT 查询，对打字毫无影响。

    取到后先访问一次 ActiveCodePane 验证仍可用；Excel/VBE 已关闭会抛异常，
    此时丢弃缓存并返回 None。
    """
    if com_backoff_remaining() > 0:
        return None
    # 宿主窗口门【每次都查】：FindWindow 是纯 Win32 调用，开销可忽略。
    # 刻意放在缓存检查【之前】——只在新建立时才查的话，宿主已经退出、我们却
    # 还在用旧代理继续打 COM，照样会把正在退出的 Excel 拽住。
    if not _host_window_alive():
        if _wnd_state["miss"] >= 2:
            _release_vbe_proxy(collect=True)
            _log("com: 宿主窗口已消失 -> 释放 COM 代理并暂停访问")
        _com_fail()
        return None
    v = _vbe_cache.get("obj")
    if v is not None:
        # TTL 从"创建时刻"起算，命中不续期 —— 否则只要一直在打字就永不超时，
        # 又变回长期持有。TTL <= 0 表示彻底不缓存（每次新建、用完即弃）。
        expired = (VBE_PROXY_TTL <= 0
                   or time.time() - _vbe_cache.get("at", 0.0) > VBE_PROXY_TTL)
        if expired:
            _release_vbe_proxy()
            v = None
        else:
            try:
                v.ActiveCodePane
                _com_ok()
                return v
            except Exception:
                _release_vbe_proxy()
                _com_fail()
                v = None
    try:
        v = _get_vbe()
    except Exception:
        _com_fail()
        return None
    _vbe_cache["obj"] = v
    _vbe_cache["at"] = time.time()
    _com_ok()
    return v


def _classify(line_text, caret_col):
    """判断光标前是否处于字符串/注释内。

    caret_col 必须是「字符列」(1-based)。若传入 VBE 原始的「显示列」，
    行内含全角字符时它会大于 len(line_text)，下面的 line_text[i] 就会
    IndexError——所以调用方必须先完成显示列->字符列换算。
    这里再做一次钳制兜底，保证本函数本身永不因越界抛异常。
    """
    n = caret_col - 1 if caret_col > 1 else 0
    if n > len(line_text):
        n = len(line_text)
    in_str = False
    comment_pos = -1
    i = 0
    while i < n:
        c = line_text[i]
        if c == '"':
            if in_str:
                if i + 1 < n and line_text[i + 1] == '"':  # 转义引号
                    i += 2
                    continue
                in_str = False
            else:
                in_str = True
        elif c == "'" and not in_str:
            comment_pos = i
        i += 1
    return in_str, (not in_str and comment_pos != -1)


def _proc_of_line(cm, line_no):
    """返回该行所在的过程名；不在任何过程内（模块声明区）则返回 None。

    优先用「代码文本往上找最近的 Sub/Function/Property」（纯 Python，
    行为确定、可单测）；COM 的 ProcOfLine 只作兜底——它在 pywin32 下
    的 byref 参数（ProcKind）常常抛异常而拿不到值。
    """
    # 1) 文本扫描：取 1..line_no 行，从光标行往上找最近的过程头
    try:
        head = cm.Lines(1, line_no)
        name = vba_parser.proc_at_line(head, line_no)
        if name:
            return str(name)
    except Exception:
        pass

    # 2) 兜底：COM ProcOfLine（vbext_pk_Proc=0 / Let=1 / Set=2 / Get=3）
    for kind in (0, 1, 2, 3):
        try:
            name = cm.ProcOfLine(line_no, kind)
            if name:
                return str(name)
        except Exception:
            continue
    return None


# VBIDE 组件类型常量（只关心"标准模块"与"其他"两类）
VBE_CT_STDMODULE = 1      # vbext_ct_StdModule


def _is_std_module(comp_type):
    """标准模块（.bas）之外，类模块/窗体/文档模块的成员都只能
    实例或限定名访问，跨模块不做裸名提示。"""
    try:
        return int(comp_type) == VBE_CT_STDMODULE
    except Exception:
        return True   # 取不到类型时按标准模块处理（最宽松，不误伤）


class VbeBackend:
    def __init__(self):
        self._cache = None
        self._cache_time = 0
        self._declared_names = set()

    def release(self):
        """Excel/VBE 关闭或离开 VBE 时调用：清空标识符缓存并释放 COM 资源。

        这样切到其它程序（如 PyCharm）或关掉 Excel 后，占用的内存会及时回收，
        也不会有悬空的 COM 引用拖住 Excel 进程退出。
        """
        self._cache = None
        self._cache_time = 0
        # 丢掉缓存的 VBE 对象 —— 这是 Excel 能否真正退出的关键一步。
        # 刻意【不】调 CoFreeUnusedLibraries —— 它释放不了我们持有的引用，
        # 在 Excel 关闭期间调用反而会拖慢/惊扰 COM（详见 _co_free 的说明）。
        # 这里用 collect=True 主动 gc，确保 COM 包装对象立刻回收。
        _release_vbe_proxy(collect=True)

    def snapshot(self):
        """轻量读取「当前行号 + 行文本」，用于轮询检测内容变化。

        刻意不做列语义探测（那会临时移动光标），因此足够廉价，可高频调用。
        返回 (line_no, line_text)；不在代码窗/有选区/取不到时返回 None。
        """
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return None
            cp = vbe.ActiveCodePane
            if cp is None:
                return None
            sl, _sc, el, _ec = cp.GetSelection()
            if sl != el:      # 多行选择：不参与补全
                return None
            cm = cp.CodeModule
            snap = (sl, cm.Lines(sl, 1))
            _com_ok()
            return snap
        except Exception:
            _release_vbe_proxy()
            _com_fail()
            return None

    # ---- 标识符缓存 ----
    def invalidate_identifiers(self):
        """文本内容变化后调用：丢弃标识符缓存，下次取用时会重新解析。

        避免「刚声明的变量不提示」——缓存里可能还是若干秒前解析的旧内容，
        新敲出来的名字还不在里面。只清 _cache、不动 _cache_time，
        以便调用方用 _cache_time 做重解析的限流。
        """
        self._cache = None

    def get_identifiers(self):
        now = time.time()
        if self._cache is not None and now - self._cache_time < _CACHE_TTL:
            return self._cache
        ids = self._collect_identifiers()
        self._cache = ids
        self._cache_time = now
        return ids

    def get_declared_names(self):
        """返回工程里真实声明过的名字（小写集合），供引擎剔除"提示自己"的幻影。

        只有被 Dim/Const/Sub/Function/Type/Enum 及组件名声明过的才算；隐式变量
        （用到即存在）不算。缓存随 get_identifiers 一并刷新。
        """
        if getattr(self, "_declared_names", None) is None:
            self.get_identifiers()
        return getattr(self, "_declared_names", set())

    def _collect_identifiers(self):
        # 用缓存的 VBE（会顺带做失败退避），避免每次收集都新建
        # Excel.Application 代理 —— 那会让 Excel 退出时要清理的引用越堆越多。
        vbe = _get_vbe_cached()
        if vbe is None:
            return []
        modules = []
        records = []
        # 光标位置（用于排除"正在输入的词"被当成隐式变量收录）
        # 模块名比对失败就等于没拿到光标，会出现"打什么就提示什么"的自我提示，
        # 因此取不到模块名时宁可对所有模块应用（误伤极小），也不要丢弃 caret。
        caret_mod, caret = _caret_position(vbe)
        _log("collect: caret_mod=%r caret=%r" % (caret_mod, caret))
        # 注意：早先这里会把"声明行正在声明的名字"从【全部候选】里按名字整体剔除
        # （连前缀一起），用于杜绝"定义变量时提示变量自己"。但这会把合法的前缀
        # 补全也误杀——在 numArr 声明行输入 num，numArr 就再也提示不出来（bug：
        # 输入 num 只能提醒 num、不提醒 numArr）。正确的语义是"精确匹配自己才隐藏、
        # 前缀照常提示"，这由引擎层 in_decl_position + ctx["decl_names"] 的精确自我
        # 剔除负责（输入完整 numArr 才隐藏，输入前缀 num 仍正常提示 numArr）。
        # 因此解析/收集层不再做任何声明行全局剔除，名字一律收录，交给引擎按"正在
        # 输入的词"与"当前声明行的名字"决定是否剔除。
        # "类型名"集合：`Dim f As <这里>` 用得上——窗体名、类模块名、标准模块名、
        # 以及代码里定义的 Type / Enum 名。
        type_names = set()
        # 真实"声明"过的名字（Dim/Const/Sub/Function/Type/Enum 及组件名）。
        # 供引擎层区分"真名字"与"幻影"：回退删字回退出来的未定义词（如 numA）
        # 只可能是隐式残留 / 正在敲的词本身，引擎据此剔除"提示自己"。
        declared_names = set()
        try:
            for comp in _collect_components(vbe, caret_mod):
                try:
                    try:
                        mod_name = str(comp.Name)
                    except Exception:
                        continue
                    # 组件名本身（模块名 / 窗体名 / 类名 / ThisWorkbook、Sheet1 等
                    # 文档模块名）也是工程内可以直接引用的标识符：
                    #   Module1.Foo / UserForm1.Show / Dim c As New Class1
                    # 记为模块级、非私有 -> 本工程任意位置都能提示。
                    #
                    # 必须在读 CodeModule【之前】收录：新建的窗体 / 类模块通常
                    # 一行代码都没有（CountOfLines == 0）。放在读代码之后会被
                    # `if count <= 0: continue` 整个跳过 —— 症状正是"模块名能
                    # 提示、窗体名死活不提示"。工程被密码保护、CodeModule 取不到
                    # 时同理，名字照样该提示。
                    if _RE_PLAIN_IDENT.match(mod_name):
                        records.append((mod_name, mod_name, None, False))
                        type_names.add(mod_name)
                        declared_names.add(mod_name.lower())
                    modules.append(mod_name)
                    cm = comp.CodeModule
                    count = cm.CountOfLines
                    if count <= 0:
                        continue
                    code = cm.Lines(1, count)
                    # 模块名对不上就等于没拿到光标，会出现"打什么就提示什么"
                    # 的自我提示；取不到模块名时宁可对所有模块应用（误伤极小），
                    # 也不要丢弃 caret。
                    apply_caret = caret if (
                        not caret_mod
                        or str(caret_mod).lower() == str(mod_name).lower()
                    ) else None
                    is_std = _is_std_module(comp.Type)
                    # 代码里自定义的 Type / Enum 名同样是类型名
                    for tn in vba_parser.extract_type_enum_names(code):
                        type_names.add(tn)
                    recs = vba_parser.extract_records(
                        code, module=mod_name, is_std_module=is_std,
                        caret=apply_caret)
                    records.extend(recs)
                    # "真名集合"必须剔除【光标所在声明行正在声明的那些名字】：
                    # 那是用户正在敲 / 正在删的词，不是"工程里已存在的真名字"。
                    #
                    # 不剔除会怎样：回退删字时（Dim numArr 退成 Dim numA），
                    # 文本里此刻确实写着 Dim numA，解析层会把它当成一个真实的
                    # 声明收进真名集合；引擎"不是真名就剔除自身"的防护随即失效，
                    # numA 又被当作一条补全建议提示出来 —— 正是用户反馈的
                    # "回退快的时候还是提示 numA"。
                    caret_decl = set()
                    if apply_caret:
                        try:
                            caret_decl = set(
                                str(n).lower() for n in
                                (vba_parser.decl_names_at_caret(
                                    code, apply_caret) or ()))
                        except Exception:
                            caret_decl = set()
                    declared_names.update(
                        str(r[0]).lower() for r in recs
                        if str(r[0]).lower() not in caret_decl)
                    # 隐式变量（没写 Option Explicit 时"用到即存在"）：
                    # 只补 extract_records 没声明过的名字，避免重复与作用域冲突。
                    imp = []
                    if ENABLE_IMPLICIT_IDENTIFIERS:
                        imp = vba_parser.extract_implicit_records(
                            code, module=mod_name, is_std_module=is_std,
                            declared=recs,
                            scope=IMPLICIT_SCOPE,
                            caret=apply_caret)
                        records.extend(imp)
                    _log("  mod=%-16s type=%-3s std=%-5s recs=%-4d imp=%-4d caret_applied=%s"
                         % (mod_name, comp.Type, is_std, len(recs), len(imp),
                            bool(apply_caret)))
                except Exception:
                    continue
        except Exception:
            return []
        self._type_names = type_names
        self._declared_names = declared_names
        # 只在数量变化时记日志，避免每 2 秒刷一行
        if len(records) != getattr(self, "_last_id_count", -1):
            self._last_id_count = len(records)
        return records

    def get_type_names(self):
        """返回可作为「数据类型名」使用的名字（窗体/类模块/标准模块名、Type/Enum 名）。

        供引擎在 `Dim x As |` 这类位置过滤候选：那里只该冒出类型名，
        不该冒出变量与过程名。
        """
        return getattr(self, "_type_names", set())

    # ---- 当前光标上下文 ----
    def get_context(self):
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return None
            cp = vbe.ActiveCodePane
            if cp is None:
                return None
            sl, sc, el, ec = cp.GetSelection()
            if sl != el:  # 存在多行选择，不触发
                return None
            cm = cp.CodeModule
            line_text = cm.Lines(sl, 1)

            # 顺序是关键：必须「先换算列，再判定字符串/注释」。
            # VBE 的 ec 是显示列（中文等全角占 2 格），而 line_text 是字符序列。
            # 若直接拿 ec 去 _classify 索引 line_text，行内含全角时 ec-1 会
            # 超过 len(line_text)-1 -> IndexError -> get_context 返回 None
            # -> 中文输入永远不弹列表（纯英文行显示列==字符列，所以一直正常）。
            ec_char = ec
            sem, tabw, wide2 = "char", 4, False
            has_wide = any(_is_wide(c) for c in line_text)
            if ec > 0 and ("\t" in line_text or has_wide):
                # 判定顺序（越靠前越安全）：
                #   1) 被动判定：完全不移动光标（打字时光标多在行尾，这条最常用）；
                #   2) 会话缓存：同一编辑器探测一次后长期复用；
                #   3) 才移动光标探测（_probe_semantics 会自动还原+校验，
                #      还原不可靠的编辑器会永久禁用探测，避免光标乱跳）。
                info = _detect_semantics_passive(line_text, ec)
                if info is None:
                    info = _sem_cache.get("info")
                if info is None:
                    info = _probe_semantics(cm, cp, sl, line_text)
                if info is None:
                    info = ("char", 4, False)
                sem, tabw, wide2 = info
                if sem == "disp":
                    idx = _col_to_char_index(line_text, ec, tabw, wide2)
                    ec_char = idx + 1
                else:
                    ec_char = min(ec, len(line_text) + 1)
            elif ec > 0:
                ec_char = min(ec, len(line_text) + 1)
            in_str, in_comment = _classify(line_text, ec_char)
            # 声明里 `As` 之后的数据类型名位置：正在填类型名，不该提示变量名。
            # 即便后面的变量名前缀匹配上已有标识符，也应保持静默。
            in_type = vba_parser.is_caret_in_type_position(line_text, ec_char)
            # 正在"起新名字"的声明行（Dim/Const/Sub/Function/Public 变量...）：
            # 此刻任何候选都是噪音，且正在输入的词本身常已存在于标识符池
            # （它在别处被用过、或别的模块有同名声明）-> 表现为提示出自己。
            # 这里直接整行静默，不弹窗。
            in_decl = vba_parser.is_caret_in_declaration(line_text, ec_char)

            # 当前声明行正在声明的名字集合（小写）。供引擎在 in_decl_position
            # 时精确剔除"正在输入的那个词本身"（打全名才隐藏，前缀照常提示），
            # 从而修好"输入 num 只提醒 num、不提醒 numArr"的前缀补全误杀。
            # 只在声明行上算，平时为空、零额外开销。
            decl_names = []
            if in_decl:
                try:
                    full = cm.Lines(1, cm.CountOfLines)
                    dn = vba_parser.decl_names_at_caret(full, (sl, ec_char))
                    if dn:
                        decl_names = sorted(dn)
                except Exception:
                    decl_names = []

            try:
                module_name = str(cm.Name)
            except Exception:
                module_name = None
            return {
                "line_no": sl,
                "caret_col": ec_char,          # 字符列（1-based）
                "line_text": line_text,
                "in_string": in_str,
                "in_comment": in_comment,
                "in_type_position": in_type,
                "in_decl_position": in_decl,
                "decl_names": decl_names,
                "proc_name": _proc_of_line(cm, sl),
                "module_name": module_name,
            }
        except Exception:
            return None

    # ---- 应用补全 ----
    def apply_completion(self, line_no, word_start_col, caret_col, completion):
        """把 [word_start_col, caret_col) 区间的单词替换为 completion 并写回 VBE。

        word_start_col / caret_col 均为 1-based 字符列（与 Python 行文本一致）。
        写回后把光标放到补全词尾部：
          - 若该行(含 Tab/全角)在 VBE 里按"显示列"计，就把字符列换算成显示列，
            否则光标会因 Tab 展开的格差落在词的中间；
          - 若 VBE 按"字符列"计（探测结果 'char'），直接用字符列。
        """
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return None
            cp = vbe.ActiveCodePane
            if cp is None:
                return None
            cm = cp.CodeModule
            line_text = cm.Lines(line_no, 1)
            new_line = replace_word(line_text, word_start_col, caret_col, completion)

            # 词尾字符下标（0-based）
            start0 = max(0, word_start_col - 1)
            end0 = start0 + len(completion)

            cm.ReplaceLine(line_no, new_line)
            # 读回实际行（VBE 可能做规范化，内容略有变化）
            actual = cm.Lines(line_no, 1)
            if actual != new_line:
                end0 = _find_completion_end(actual, completion,
                                            end0 - len(completion), end0)
            end0 = min(end0, len(actual))
            # VBE 自动补的括号（`Function gCalc` -> `Function gCalc()`）：
            # 光标要落进括号里，否则停在左括号左侧还得手动右移一格。
            end0 = _skip_into_parens(actual, end0, line_text, caret_col)
            end0 = min(end0, len(actual))

            # 关键：列语义探测必须基于「写回之后」的实际行。
            # 典型故障：用户只打了首字母（纯 ASCII，如 abc中文 只打了 a），
            # 旧行不含任何全角字符 -> 探测被跳过、默认 'char'（字符列）；
            # 但补全后的新行含中文，VBE 按显示列计（中文占 2 格），
            # 于是按字符列定位光标会少算格差，光标落在变量名中间。
            sem, tabw, wide2 = "char", 4, False
            if "\t" in actual or any(_is_wide(c) for c in actual):
                # 用带缓存的探测：同一编辑器只探测一次，且会自动还原光标。
                info = _probe_semantics(cm, cp, line_no, actual)
                if info is not None:
                    sem, tabw, wide2 = info

            def _set_caret(text, off):
                """按当前列语义把光标放到字符下标 off 之后。"""
                if sem == "disp":
                    col = _disp_width(text, tabw, wide2, upto=off) + 1
                else:
                    col = off + 1
                cp.SetSelection(line_no, col, line_no, col)

            _set_caret(actual, end0)
            # 二次确认：个别情况下 VBE 是在我们读完行【之后】才补括号的，
            # 那一对 `()` 会把光标"挤"到左括号左侧。再读一次行，若它又变了、
            # 且光标处正是新出现的 `(`，就右移一格把光标放进括号。
            try:
                actual2 = cm.Lines(line_no, 1)
            except Exception:
                actual2 = actual
            if actual2 != actual:
                end0b = _skip_into_parens(actual2, end0, line_text, caret_col)
                if end0b != end0:
                    end0 = min(end0b, len(actual2))
                    try:
                        _set_caret(actual2, end0)
                    except Exception:
                        pass
            return new_line
        except Exception:
            return None
        finally:
            # 刚用 ReplaceLine 改过 VBA 代码（VBAProject 变脏），此刻手上还攥着
            # CodeModule / CodePane / VBE 的引用，而 Excel 退出时正是要把这个
            # 工程写回文件 —— 立刻放手，别把它的退出流程钉住。
            # （"只改工作表不卡、改过 VBA 才卡"的针对性处理。）
            _release_vbe_proxy()
