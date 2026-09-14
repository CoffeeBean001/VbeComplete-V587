"""
VBE 后端：通过 pywin32 与正在运行的 Excel VBE 交互。

只在这一层依赖 pywin32 / Excel，引擎层不依赖，方便单测。
需要：Excel 已开启「信任对 VBA 工程对象模型的访问」。
"""

import difflib
import os
import re
import threading
import time
import unicodedata

from log import log as _log

import parser as vba_parser
import vba_builtins
from engine import replace_word

_CACHE_TTL = 2.0  # 标识符缓存刷新间隔（秒）

# 未声明即使用的"隐式变量"（不写 Option Explicit 时的用法）是否纳入提示。
# 关闭后行为回到旧版：只提示 Dim/Const/Sub 等显式声明出来的名字。
ENABLE_IMPLICIT_IDENTIFIERS = True
# 隐式变量的作用域粒度：
#   "module" —— 本模块任意位置都能提示（最宽松，但会跨过程泄漏同名变量）
#   "proc"   —— 按"首次出现所在过程"限定作用域（更贴近 VBA 语义，避免跨过程泄漏）
IMPLICIT_SCOPE = "proc"

# VBA 语言自带的名字（内建函数 / 常用内建常量 / 内建数据类型）是否纳入提示。
# 关闭后回到旧行为：只提示工程代码里出现过的名字（外加组件名、窗体控件名；
# 关键字另见下面的 ENABLE_VBA_KEYWORDS，同样关掉才是完全旧行为）。
ENABLE_VBA_BUILTINS = True

# 语言关键字 / 保留字（Sub / Dim / If / For / Set / And …）是否纳入提示（v62）。
#
# 独立成一个开关（而不是并进 ENABLE_VBA_BUILTINS）：关键字与内建函数是两拨
# 东西 —— 前者是语法骨架、后者是可调用的库成员。实测下来若觉得关键字提示偏吵，
# 只关这一个即可，内建函数照常提示。
ENABLE_VBA_KEYWORDS = True
# 内建名字挂靠的"模块名"。
#
# 这是一个【虚拟模块】——VBA 运行时库。之所以要给它一个名字而不是留空：
#   1) 这些名字在工程里任何模块都能裸名引用，按模块级 priv=False 收录即可；
#   2) 引擎判"回声候选是不是真名字"时会问 declared_elsewhere（名字是不是
#      在【别的模块】声明过）。把内建名字的归属记成 VBA，那一问天然为真，
#      内建名字就不会被"只在本模块声明过 + 现场文本扫不到 -> 当幽灵剔掉"
#      这条规则误杀（详见 engine._name_really_exists）。
_BUILTIN_MODULE = "VBA"

# 合法标识符（含中文）的模块名/窗体名/类名，用于把组件名纳入候选
_RE_PLAIN_IDENT = re.compile(r"^[^\W\d]\w*$")
# 行首空白（空格 / Tab）：新起一行时用它对齐上一行的代码起始位置。
_RE_LEADING_WS = re.compile(r"[ \t]*")

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


def _leading_ws(line_text):
    """行首空白（空格 / Tab）。

    新起一行时用它对齐上一行的代码起始位置 —— VBE 自己按回车也遵循这个规则
    （新行继承上一行缩进），我们在光标位于行中间时代劳，必须手工复刻。
    """
    m = _RE_LEADING_WS.match(line_text or "")
    return m.group(0) if m else ""


def _indent_end_col(indent):
    """缩进【之后】那一列（1-based）——新行光标要停在这里。

    缩进只由空格与 Tab 组成、不含全角字符，因此：
      * 编辑器按【字符列】计时 col = len(indent) + 1；
      * 编辑器按【显示列】计时，Tab 要展开到下一个 tab stop。
    列语义是编辑器级属性（与 apply_completion 共用 _sem_cache）；拿不到缓存时
    按字符列算 —— 缩进全是空格（最常见）时两种算法结果一致。
    """
    info = _sem_cache.get("info")
    if info:
        try:
            sem, tabw, wide2 = info
            if sem == "disp":
                return _disp_width(indent, tabw, wide2) + 1
        except Exception:
            pass
    return len(indent) + 1


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


# ---------------------------------------------------------------------------
# v63：VBE 自己弹的提示窗（参数信息 / 列出成员）此刻是否可见
# ---------------------------------------------------------------------------
# 让位机制的第二路判据。v55 那套走的是【光标语法位置】（`标识符.` 之后、`As ` /
# `New ` 之后），只能覆盖"自动列出成员 / 类型列表"；VBE 在别处弹的提示它一概不
# 认 —— 最典型的就是【参数信息】：`MsgBox "已完成！",` 敲下逗号后 VBE 弹出参数
# 签名，我们的弹窗同时也冒出来，两个窗叠在同一处（用户 v63 报的 bug）。
#
# 真机实测（本机 Office VBE，2026-09-14）：
#   * VBE 的提示窗是【预建复用】的真窗口 —— 什么都没按的时候它们就在顶层窗口
#     列表里，只是 IsWindowVisible=False；发 Ctrl+Shift+I（参数信息）后立刻
#     变 True、Esc 后立刻回 False。
#   * 类名是 VB IDE 专有的两个：
#       NameListWndClass —— 列出成员 / 列出常量
#       PopupTipWndClass —— 参数信息（Quick Info）
#   * 开销：全量枚举（EnumWindows + GetClassName）1.75ms/次、按进程过滤
#     0.81ms/次；而句柄缓存之后只查 IsWindowVisible，0.0005ms/次。所以这里缓存
#     句柄，只在"缓存为空"或"句柄失效（Excel/VBE 重启过）"时重新枚举，且限流。
#
# ⚠️ v55 曾以"给 VBE 发 Ctrl+J 后全系统零新窗口"推断【VBE 的列表没有独立窗口】。
# 那个结论只对"【新建】窗口"成立 —— 实际上窗口早就建好了，只是被显示/隐藏，
# 所以枚举"新窗口"永远看不到它。这次改成直接查已知类名的可见性，才拿到真相。
#
# 想追加别的类名（不同 VBE 版本类名可能不一样）：
#   set VBECOMPLETE_VBE_POPUP_CLASS=ClassA,ClassB
VBE_POPUP_CLASSES = ("NameListWndClass", "PopupTipWndClass")
try:
    _extra_popup_cls = tuple(
        c.strip() for c in
        os.environ.get("VBECOMPLETE_VBE_POPUP_CLASS", "").split(",")
        if c.strip())
    if _extra_popup_cls:
        VBE_POPUP_CLASSES = tuple(VBE_POPUP_CLASSES) + _extra_popup_cls
except Exception:
    pass

# 缓存：命中的窗口句柄列表 + 上次枚举时刻（找不到时限流重扫）
_vbe_popup_state = {"hwnds": [], "at": 0.0}
_VBE_POPUP_RESCAN_SEC = 1.0

# ---------------------------------------------------------------------------
# v65：让位只针对【成员列表】，不针对【参数信息】
# ---------------------------------------------------------------------------
# 用户口径（v65 澄清 v63 的原意）："只有弹成员列表的时候才让位，弹形参签名
# 不需要让位。"
#
# 为什么该这么分：
#   * NameListWndClass（列出成员 / 列出常量）本身就是一份【候选列表】—— 与我们
#     的候选窗完全同质：同样画在光标下方、同样吃 Tab / ↑↓ / Enter 选词。两个
#     列表叠在一起既遮挡又抢键盘，让位是对的（v55 的原意）。
#   * PopupTipWndClass（参数信息 = 形参签名）只是一行只读提示：不给候选、不吃
#     键盘。而它出现的位置恰恰是"正在填实参"——**最需要候选的时候**。这时候
#     让位，就变成了"这里输入任何字符都无提醒"（v65 报的 bug）。
#
# 想调整（比如某版本的成员列表类名不同）：
#   set VBECOMPLETE_VBE_YIELD_CLASS=ClassA,ClassB
VBE_YIELD_CLASSES = ("NameListWndClass",)
try:
    _extra_yield_cls = tuple(
        c.strip() for c in
        os.environ.get("VBECOMPLETE_VBE_YIELD_CLASS", "").split(",")
        if c.strip())
    if _extra_yield_cls:
        VBE_YIELD_CLASSES = tuple(VBE_YIELD_CLASSES) + _extra_yield_cls
except Exception:
    pass


def _enum_vbe_popup_windows():
    """枚举顶层窗口，返回类名命中 VBE_POPUP_CLASSES 的 [(hwnd, cls)]。"""
    import ctypes
    u32 = ctypes.windll.user32
    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool,
                                     ctypes.c_void_p, ctypes.c_void_p)
    buf = ctypes.create_unicode_buffer(256)
    out = []

    def _cb(hwnd, _lp):
        try:
            if u32.GetClassNameW(hwnd, buf, 256):
                cls = buf.value
                if cls in VBE_POPUP_CLASSES:
                    out.append((hwnd, cls))
        except Exception:
            pass
        return True

    u32.EnumWindows(WNDENUMPROC(_cb), 0)
    return out


def _hwnd_alive(hwnd):
    import ctypes
    try:
        return bool(ctypes.windll.user32.IsWindow(hwnd))
    except Exception:
        return False


def _host_process_id():
    """宿主（Excel / VBE）进程 PID；找不到返回 0。纯 Win32，不碰 COM。"""
    import ctypes
    import ctypes.wintypes
    u32 = ctypes.windll.user32
    fn = u32.FindWindowW
    fn.restype = ctypes.c_void_p
    fn.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p]
    for cls in _HOST_WINDOW_CLASSES:
        try:
            hwnd = fn(cls, None)
        except Exception:
            hwnd = 0
        if hwnd:
            pid = ctypes.wintypes.DWORD()
            u32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            return int(pid.value or 0)
    return 0


def _window_class_name(hwnd):
    import ctypes
    try:
        buf = ctypes.create_unicode_buffer(256)
        if ctypes.windll.user32.GetClassNameW(hwnd, buf, 256):
            return buf.value
    except Exception:
        pass
    return ""


def _popup_rects_now(hwnds):
    """这些句柄里正在显示的 VBE 提示窗，返回 [(类名, 左, 上, 右, 下), ...]。

    逐个复核三件事（句柄值会被 Windows 回收复用给别的窗口，只凭"还记得这个
    句柄"就相信它，一旦复用就会出现"永远以为 VBE 在弹提示、我们的窗再也不弹"
    的灾难）：
      1. 类名仍是 VBE 的提示窗类；
      2. 属于宿主 Excel/VBE 进程；
      3. 可见且尺寸正常。

    返回【位置】而不只是 bool/尺寸：v66 起 UI 要拿它做避让（把候选窗挪到
    参数信息窗下方），所以必须知道屏幕坐标。冷路径开销可忽略：缓存里最多
    两三个句柄，逐句柄 GetWindowRect。
    """
    import ctypes
    import ctypes.wintypes
    u32 = ctypes.windll.user32
    rect = ctypes.wintypes.RECT()
    host_pid = _host_process_id()
    out = []
    for hwnd in hwnds:
        try:
            cls = _window_class_name(hwnd)
            if cls not in VBE_POPUP_CLASSES:
                continue
            if host_pid:
                pid = ctypes.wintypes.DWORD()
                u32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                if int(pid.value or 0) != host_pid:
                    continue
            if not u32.IsWindowVisible(hwnd):
                continue
            if not u32.GetWindowRect(hwnd, ctypes.byref(rect)):
                continue
            l, t = int(rect.left), int(rect.top)
            r, b = int(rect.right), int(rect.bottom)
            if r - l <= 0 or b - t <= 0:
                continue
            out.append((cls, l, t, r, b))
        except Exception:
            continue
    return out


def _popup_details_now(hwnds):
    """这些句柄里正在显示的 VBE 提示窗，返回 [(类名, 宽, 高), ...]。

    与 _popup_rects_now 是同一套判据（这里只是把矩形投影成"宽高"），供诊断
    日志与让位判据使用 —— 明细能区分【成员列表】(NameListWndClass) 与
    【参数信息】(PopupTipWndClass)，两者的结论不同（v65）。
    """
    return [(c, r - l, b - t) for (c, l, t, r, b) in _popup_rects_now(hwnds)]


def _vbe_popup_hwnds():
    """取（并维护）VBE 提示窗的句柄缓存；返回当前可用的句柄列表。

    只在"缓存为空"或"句柄失效（Excel/VBE 重启过）"时重新枚举，且限流 ——
    枚举全量顶层窗口 1.75ms/次，命中缓存后只查 IsWindowVisible，快三个数量级。
    """
    cached = _vbe_popup_state["hwnds"]
    hs = [h for h in cached if _hwnd_alive(h)]
    if len(hs) != len(cached):
        _vbe_popup_state["hwnds"] = hs         # 有句柄失效 -> 作废重找
    if not hs:
        now = time.time()
        if now - _vbe_popup_state["at"] < _VBE_POPUP_RESCAN_SEC:
            return []
        _vbe_popup_state["at"] = now
        hs = [h for h, _cls in _enum_vbe_popup_windows()]
        _vbe_popup_state["hwnds"] = hs
    return hs


def vbe_popup_info():
    """当前可见的 VBE 提示窗明细 [(类名, 宽, 高), ...]；探不到返回 []。

    只读 Win32 查询，不碰 COM。专供诊断日志（区分成员列表 / 参数信息）。
    """
    try:
        return _popup_details_now(_vbe_popup_hwnds())
    except Exception:
        return []


def vbe_popup_visible():
    """VBE 自带的提示窗此刻是否可见（v63）。探不到一律 False（不拦）。

    只读 Win32 查询，不碰 COM。Excel / VBE 没开时自然返回 False。
    ⚠️ 语义是"**任何**提示窗"（含参数信息），只用于诊断与兼容；
    引擎真正据以让位的是 vbe_yield_visible()（v65 起只认成员列表）。
    """
    try:
        return bool(_popup_details_now(_vbe_popup_hwnds()))
    except Exception:
        return False


def vbe_yield_visible():
    """VBE 的【成员列表】窗此刻是否可见 —— 引擎据此让位（v65）。

    只认 VBE_YIELD_CLASSES（默认只有 NameListWndClass）。参数信息
    （PopupTipWndClass）**不算**：它不给候选、不吃键盘，而它出现时用户正在填
    实参，正是最需要候选的时候 —— 在那里让位就成了"输入什么都不弹"（v65）。

    只读 Win32，探不到一律 False（不让位）。
    """
    try:
        for _cls, _w, _h in _popup_details_now(_vbe_popup_hwnds()):
            if _cls in VBE_YIELD_CLASSES:
                return True
        return False
    except Exception:
        return False


def vbe_popup_rects():
    """当前可见的 VBE 提示窗的屏幕矩形 [(类名, 左, 上, 右, 下), ...]（v66）。

    UI 拿它做【避让】：VBE 的参数信息窗（形参签名）就画在光标正下方，而我们的
    候选窗默认也在那儿（光标底边 +2px）—— 两窗一重叠就互相遮挡。知道它画在哪，
    就能把候选窗下移到它下面。

    用户口径（v66）："把所有这种我们候选窗和参数信息重叠的情况，都改成把候选窗
    下移到签名提示的下方"；并明确【自定义函数】的参数信息也要一并处理。
    VBE 对自己工程里的过程同样会弹这个窗（它认识过程签名），所以这里**不区分
    "内建还是自定义"** —— 一律按"可见的提示窗"避让即可，天然覆盖自定义函数。

    只读 Win32、不碰 COM；探不到返回 []（UI 就按原位置显示，绝不让"拿不准"
    影响显示 —— v54 的教训）。
    """
    try:
        return _popup_rects_now(_vbe_popup_hwnds())
    except Exception:
        return []


# ---------------------------------------------------------------------------
# v64：焦点是不是真的在【代码窗格】上
# ---------------------------------------------------------------------------
# 症状（用户报）：在【属性窗口】里改属性值时，候选窗偶尔会冒出来 —— 那里根本
# 不是写代码的地方，用户不希望在这些非代码工作区出现提示。
#
# 根因：判断"是否在 VBE 里"一直是【前台窗口标题里有没有 "Microsoft Visual
# Basic"】（main.in_vbe_code_pane）。而属性窗口 / 工程窗口 / 窗体设计器全都是
# VBE 主窗口（wndclass_desked_gsk）里的子窗口，标题判据对它们一律返回 True。
# 于是两件事都会发生：
#   * 轮询发现代码窗格那行的文本变了（切代码窗格、属性改动导致 VBE 改写代码…）
#     就照样弹候选窗；
#   * 自动配对（敲 `(` / `"`）与 Shift+Enter 也会把内容写进【代码窗格】，
#     而用户此刻其实在属性窗口里打字。
#
# 精确判据（真机实测 2026-09-14，本机 Office VBE）：
#   * 代码窗格的窗口类是 **VbaWindow**，标题形如
#     `下单填写模板.xlsm - UserForm1 (代码)`；它挂在 MDIClient 之下：
#       VbaWindow < MDIClient < wndclass_desked_gsk
#   * 属性窗口的窗口类是 **wndclass_pbrs**（标题 `属性 - CommandButton2`），
#     里面的编辑框 / 列表是 Edit / ListBox / ComboBox / SysTabControl32；
#   * 工程窗口是 **PROJECT**（标题 `工程 - VBAProject`）；
#   * 窗体设计器是 **DesignerWindow**，里面装着被设计的窗体（ThunderDFrame）。
#   所以"焦点窗口的父链上有没有 VbaWindow"是一个干净、与语言无关的判据。
#
# 为什么用 GetGUIThreadInfo【VBE 线程】而不是 GetForegroundWindow：
#   中文输入法的候选/组字窗口会短暂抢走【前台窗口】，但 VBE 线程内部的焦点
#   窗口不受影响 —— 用它就不会在中文输入时把代码窗格误判成"没焦点"。
#
# 探不到（VBE 主窗口找不到 / API 失败 / 焦点窗口为 0）一律返回 None（拿不准），
# 由调用方退回老判据 —— 绝不"拿不准就拦"（v54 就是这么把工具搞成什么都不弹的）。
VBE_FRAME_CLASSES = ("wndclass_desked_gsk",)
try:
    _extra_frame_cls = tuple(
        c.strip() for c in
        os.environ.get("VBECOMPLETE_VBE_FRAME_CLASS", "").split(",")
        if c.strip())
    if _extra_frame_cls:
        VBE_FRAME_CLASSES = tuple(VBE_FRAME_CLASSES) + _extra_frame_cls
except Exception:
    pass

CODE_PANE_CLASSES = ("VbaWindow",)

_focus_state = {"frame": 0}
_GUI_CACHE = {"cls": None, "ready": False}


def _focus_protos():
    """给这几个 Win32 调用设好原型并缓存（句柄是 64 位，restype 不设会被截断）。"""
    import ctypes
    import ctypes.wintypes as wt
    if _GUI_CACHE["ready"]:
        return ctypes.windll.user32
    u32 = ctypes.windll.user32
    u32.FindWindowW.restype = ctypes.c_void_p
    u32.FindWindowW.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p]
    u32.GetParent.restype = ctypes.c_void_p
    u32.GetParent.argtypes = [ctypes.c_void_p]
    u32.GetWindowThreadProcessId.restype = wt.DWORD
    u32.GetWindowThreadProcessId.argtypes = [ctypes.c_void_p,
                                             ctypes.POINTER(wt.DWORD)]
    u32.GetGUIThreadInfo.restype = wt.BOOL
    u32.GetGUIThreadInfo.argtypes = [wt.DWORD, ctypes.c_void_p]
    u32.IsWindow.argtypes = [ctypes.c_void_p]

    class _GUITHREADINFO(ctypes.Structure):
        _fields_ = [("cbSize", wt.DWORD),
                    ("flags", wt.DWORD),
                    ("hwndActive", wt.HWND),
                    ("hwndFocus", wt.HWND),
                    ("hwndCapture", wt.HWND),
                    ("hwndMenuOwner", wt.HWND),
                    ("hwndMoveSize", wt.HWND),
                    ("hwndCaret", wt.HWND),
                    ("rcCaret", wt.RECT)]

    _GUI_CACHE["cls"] = _GUITHREADINFO
    _GUI_CACHE["ready"] = True
    return u32


def _vbe_frame_hwnd():
    """VBE 主窗口句柄（缓存 + 失效重找）；找不到返回 0。纯 Win32。"""
    u32 = _focus_protos()
    h = int(_focus_state.get("frame") or 0)
    if h and u32.IsWindow(h):
        return h
    h = 0
    for cls in VBE_FRAME_CLASSES:
        try:
            got = u32.FindWindowW(cls, None)
        except Exception:
            got = 0
        if got:
            h = int(got)
            break
    _focus_state["frame"] = h
    return h


def _focused_hwnd_in_vbe():
    """VBE 线程里"当前拥有键盘焦点的窗口"；取不到返回 0。"""
    u32 = _focus_protos()
    frame = _vbe_frame_hwnd()
    if not frame:
        return 0
    tid = int(u32.GetWindowThreadProcessId(frame, None) or 0)
    if not tid:
        return 0
    import ctypes
    import ctypes.wintypes as wt
    gti = _GUI_CACHE["cls"]()
    gti.cbSize = ctypes.sizeof(gti)
    if not u32.GetGUIThreadInfo(tid, ctypes.byref(gti)):
        return 0
    return int(gti.hwndFocus or 0)


def _parent_hwnd(h):
    """父窗口句柄（取不到 / 没有父窗口返回 0）。单独一层，便于单测打桩。"""
    try:
        u32 = _focus_protos()
        return int(u32.GetParent(h) or 0)
    except Exception:
        return 0


def vbe_code_pane_focused():
    """焦点此刻是否落在 VBE 的【代码窗格】上。

    返回：
      True  —— 焦点窗口的父链上有代码窗格（类名 VbaWindow）；
      False —— 焦点在 VBE 里，但不在代码窗格（属性窗口 / 工程窗口 / 设计器…）；
      None  —— 拿不准（VBE 主窗口找不到、API 失败、焦点窗口为空）。

    调用方约定：None 必须退回老判据（前台窗口标题），绝不当成 False 去拦 ——
    否则在类名不符的宿主上会把工具整个哑掉。

    只读 Win32 查询，不碰 COM；每次约 0.02ms（FindWindow 有缓存）。
    """
    try:
        frame = _vbe_frame_hwnd()
        if not frame:
            return None
        h = _focused_hwnd_in_vbe()
        if not h:
            return None
        # v65：VBE 自己的提示窗（成员列表 NameListWndClass / 参数信息
        # PopupTipWndClass）会短暂拿到键盘焦点。它们只可能在"正在编辑代码"时
        # 出现，属于代码窗格的临时浮层 —— 必须算作"还在代码窗格"。否则它们一
        # 冒出来，本判据立刻变 False，把我们的候选窗当场收掉：表现就是"在填实参
        # 的位置输入任何字符都不提示"（v65 报的 bug，第二处根因）。
        if _window_class_name(h) in VBE_POPUP_CLASSES:
            return True
        for _ in range(12):
            if not h:
                break
            if _window_class_name(h) in CODE_PANE_CLASSES:
                return True
            if h == frame:
                break
            p = _parent_hwnd(h)
            if not p or p == h:
                break
            h = p
        return False
    except Exception:
        return None


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


def _on_main_thread():
    """当前是否在【主线程】上 —— 只有主线程允许碰 COM（见 _get_vbe_cached）。"""
    try:
        return threading.current_thread() is threading.main_thread()
    except Exception:
        return True


_warned_thread = {"n": 0}


def _get_vbe_cached():
    """取 VBE 对象（**短命**代理，TTL 到期主动放手）。

    v57 重大约束：**只允许主线程调用**。
    键盘低层钩子运行在系统创建的【钩子线程】上，那个线程没有 COM 单元、
    也不该被阻塞（钩子回调超时会被 Windows 直接摘掉）。在那里调 COM 必然
    失败，更糟的是失败会走 _com_fail() 把全局退避打成 1s/2s/5s —— 连主线程
    的正常轮询一起挡住，表现为"敲一下 `(` 之后提示要等 2~3 秒才出来"。
    所以这里直接拒绝，并【不记失败】，避免污染退避状态。

    注意与"长缓存"的取舍：复用代理确实省一次 GetActiveObject，但代价是我们
    的进程会一直持有 Excel.Application —— 而只要持有它，Excel 就退不干净
    （关完又启动 Office），且 VBAProject 变脏时的写回会被拖成 3~10 秒。

    所以这里只做**短缓存**：TTL（默认 1 秒）一到就丢弃，让 Excel 随时有
    机会干净退出；重建成本只是一次 ROT 查询，对打字毫无影响。

    取到后先访问一次 ActiveCodePane 验证仍可用；Excel/VBE 已关闭会抛异常，
    此时丢弃缓存并返回 None。
    """
    if not _on_main_thread():
        # 钩子线程（或其它后台线程）来要 COM：直接拒绝，且【不】计失败 ——
        # 责任在调用方（它本就不该在这里调），不该让主线程的正常访问连坐。
        if _warned_thread["n"] < 3:
            _warned_thread["n"] += 1
            _log("com: 非主线程请求 COM -> 拒绝（不记失败，请改在主线程调用）")
        return None
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
VBE_CT_MSFORM = 3         # vbext_ct_MSForm（UserForm 窗体）

# 取窗体控件名时最多列出的控件数。正常窗体都在这个量级以内；
# 设上限只是为了让"设计器接口异常返回天量条目"时不会把候选池撑爆。
_FORM_CTL_LIMIT = 300


def _is_std_module(comp_type):
    """标准模块（.bas）之外，类模块/窗体/文档模块的成员都只能
    实例或限定名访问，跨模块不做裸名提示。"""
    try:
        return int(comp_type) == VBE_CT_STDMODULE
    except Exception:
        return True   # 取不到类型时按标准模块处理（最宽松，不误伤）


def _is_msform(comp_type):
    """是否 UserForm 窗体组件。"""
    try:
        return int(comp_type) == VBE_CT_MSFORM
    except Exception:
        return False


def _form_control_names(comp):
    """窗体（UserForm）设计器里的控件名列表。

    【为什么必须走 Designer】控件名只存在于【设计器数据】里，窗体代码模块
    （CodeModule）里通常一个字都看不到 —— 除非那个控件挂过事件过程
    （那种情况下才会以 `Label1_Click` 的形式出现在代码里）。所以用户"刚拖
    上去一个 Label1，希望输入 la 就能提示"只能从这里拿。

    【为什么不递归】实测 `Designer.Controls` 已经是【扁平全集】：容器控件
    （Frame / Page）自身也在列表里，它们的子控件同样被列出。真机验证：一个
    含 Frame1/Frame2 的窗体，顶层 9 条已经覆盖了 Frame 里的全部控件；若再
    对 Frame 递归一次，Label1/TextBox1/BOMList 会被重复收集一遍（这也是
    VBA 里控件名必须全窗体唯一的语义所决定的）。所以只取顶层，一次到位。

    任何一步失败都只返回已收到的部分：窗体处于运行态、宿主不支持 Designer
    接口时都可能取不到，绝不能因此让整个候选收集流程崩掉。
    """
    out = []
    try:
        cs = comp.Designer.Controls
        n = int(cs.Count)
    except Exception:
        return out
    if n > _FORM_CTL_LIMIT:
        n = _FORM_CTL_LIMIT
    for i in range(n):
        try:
            nm = str(cs.Item(i).Name)
        except Exception:
            continue
        if nm:
            out.append(nm)
    return out


def _caret_char_col(cm, cp, line_no, line_text, ec):
    """VBE 给的光标列 ec（显示列）-> (字符列 1-based, sem, tabw, wide2)。

    VBE 的 ec 是显示列（中文等全角占 2 格），而 line_text 是字符序列；行内含
    Tab / 全角时必须换算，否则索引会越过行尾或落在中文中间。换算结果连同本次
    的列语义一起返回，供"写回后摆光标"复用同一套算法。

    判定顺序（越靠前越安全，与 get_context 一致）：
      1) 被动判定：完全不移动光标（打字时光标多在行尾，这条最常用）；
      2) 会话缓存：同一编辑器探测一次后长期复用；
      3) 才移动光标探测（_probe_semantics 会自动还原+校验，还原不可靠的编辑器
         会永久禁用探测，避免光标乱跳）。
    """
    if ec <= 0:
        return 1, "char", 4, False
    if not ("\t" in line_text or any(_is_wide(c) for c in line_text)):
        return min(ec, len(line_text) + 1), "char", 4, False
    info = _detect_semantics_passive(line_text, ec)
    if info is None:
        info = _sem_cache.get("info")
    if info is None:
        info = _probe_semantics(cm, cp, line_no, line_text)
    if info is None:
        info = ("char", 4, False)
    sem, tabw, wide2 = info
    if sem == "disp":
        return (_col_to_char_index(line_text, ec, tabw, wide2) + 1,
                sem, tabw, wide2)
    return min(ec, len(line_text) + 1), sem, tabw, wide2


def _pair_insertion(line_text, col0, open_ch, close_ch):
    """算出自动配对后的 (新行文本, 光标字符下标 0-based)。

    新行文本为 None 表示"一个字符都不用写，只把光标挪到 caret_off"—— 右半边
    已经在光标右边时（打完 `s = "abc` 再按 `"`），再插一对就会变成 `""`，
    此时应当像现代编辑器那样"跨过去"。

    引号是特殊的一个键（既开又闭），所以额外看【光标左边引号数的奇偶】：
    奇数说明正处在未闭合的字符串里，这一下是收尾的那一半，只补一个。
    """
    if open_ch == '"':
        if col0 < len(line_text) and line_text[col0] == '"':
            return None, col0 + 1                       # 跨过去
        inside = line_text[:col0].count('"') % 2 == 1
        if inside:
            return line_text[:col0] + '"' + line_text[col0:], col0 + 1
        return line_text[:col0] + '""' + line_text[col0:], col0 + 1
    return (line_text[:col0] + open_ch + close_ch + line_text[col0:],
            col0 + 1)


def _locate_inserted(after, written, open_ch, close_ch, col0):
    """在【VBE 改写过的行】里重新找出我们插入的那对符号，返回光标该落在的
    位置（0-based："开符号之后"）；找不到返回 None。

    为什么需要：VBE 有「自动语法检测」，会把整行重新格式化 —— 等号两边补空格
    （`x=1` -> `x = 1`）、关键字改大小写（`if a then` -> `If a Then`）、参数逗号
    后补空格（`foo(1,2)` -> `foo(1, 2)`）、标识符与引号之间补空格
    （`w""` -> `w ""`）。行一变长，写入前算好的光标位置就偏了，表现为「光标
    没落在括号/引号中间」。真机实测（v58）：`    w=4` 里 `w` 之后敲 `"`，得到
    `    w "" = 4`，光标停在 7（第一个引号之前）而不是 8（两引号之间）。

    做法：把「我们写入的版本 written」与「VBE 改写后的版本 after」做序列对齐
    （difflib），再把 written 中 open_ch 的下标 col0 **映射**到 after 里的对应
    位置 —— 而不是去猜哪个 diff 块是"我们加的"。因为 VBE 的改写可能把我们的
    符号卷进 equal 块（比如它只在符号前补了个空格），那时"找 insert 块"就会
    落空。
    """
    want = open_ch + close_ch
    try:
        sm = difflib.SequenceMatcher(None, written, after, autojunk=False)
        p = None
        for tag, i1, i2, j1, _j2 in sm.get_opcodes():
            if tag == "equal":
                if i1 <= col0 < i2:
                    p = j1 + (col0 - i1)
                    break
            elif i1 <= col0 <= i2:      # replace / insert / delete
                p = j1
                break
        if p is None:
            return None
        if after[p:p + len(want)] == want:
            return p + 1
        if after[p:p + 1] == open_ch:
            return p + 1
        # 对齐点附近再找一下（VBE 把一对符号拆开或挪过位置）
        k = after.find(want, max(0, p - 2), p + len(want) + 3)
        if k >= 0:
            return k + 1
    except Exception:
        pass
    return None


class VbeBackend:
    def __init__(self):
        self._cache = None
        self._cache_time = 0
        self._declared_names = set()
        # 收集标识符那一刻，光标处那个词的原文（小写）。用来识别【回声】：
        # 用户正在输入 / 正在回退删除的词，不该被当成工程里的真名字。
        self._caret_word = ""
        # 名字(小写) -> 声明它的模块名集合(小写)。让引擎能区分
        # 【声明在别的模块】（跨模块 Public，可当证据）与【只在我正在编辑的
        # 这个模块里声明过】（一律以现场文本为准，见 declared_elsewhere）。
        self._declared_by_module = {}
        # 由【工程结构 / 语言运行时】而非代码文本决定其存在的名字（小写）：
        #   * 组件名（标准模块 / 窗体 / 类模块 / ThisWorkbook、Sheet1 等文档模块名）
        #     与窗体控件名 —— 窗体代码里没写过 UserForm1，UserForm1 照样是合法
        #     引用；刚拖上去的 Label1 一行代码都还没有，它的名字照样存在于设计器
        #     里（v60）；
        #   * VBA 语言自带的名字（内建函数 / 内建常量 / 内建数据类型 v61，语言
        #     关键字 v62）—— 它们由语言运行时提供，用户代码里没调用过 MsgBox
        #     也照样能写 MsgBox、新模块第一行没写过 If 也照样能补 If。
        # 引擎靠它避免把这类名字当成"回声幻影"剔掉（见 engine._name_really_exists）。
        self._structural_names = set()
        # 真正作为"语言自带词汇"收录的那些（小写，v61 内建函数/常量/类型 +
        # v62 语言关键字）。引擎对这批名字的模糊匹配收紧一档（纯分散命中要求
        # 输入至少 4 个字符，详见 engine.trigger 里的说明）。
        self._builtin_names = set()

    def release(self):
        """Excel/VBE 关闭或离开 VBE 时调用：清空标识符缓存并释放 COM 资源。

        这样切到其它程序（如 PyCharm）或关掉 Excel 后，占用的内存会及时回收，
        也不会有悬空的 COM 引用拖住 Excel 进程退出。
        """
        self._cache = None
        self._cache_time = 0
        self._caret_word = ""
        self._declared_by_module = {}
        self._structural_names = set()
        self._builtin_names = set()
        # 丢掉缓存的 VBE 对象 —— 这是 Excel 能否真正退出的关键一步。
        # 刻意【不】调 CoFreeUnusedLibraries —— 它释放不了我们持有的引用，
        # 在 Excel 关闭期间调用反而会拖慢/惊扰 COM（详见 _co_free 的说明）。
        # 这里用 collect=True 主动 gc，确保 COM 包装对象立刻回收。
        _release_vbe_proxy(collect=True)

    def snapshot(self):
        """轻量读取「模块名 + 当前行号 + 行文本」，用于轮询检测内容变化。

        刻意不做列语义探测（那会临时移动光标），因此足够廉价，可高频调用。
        返回 (line_no, line_text, module_name)；不在代码窗/有选区/取不到时返回 None。

        v64：多带一个【模块名】。主线程靠它区分"用户真在敲字"和"换了个代码
        窗格/模块"—— 后者会让"同一行 + 文本不同"看起来像一次编辑（Ctrl+Tab
        切代码窗、点工程树换模块都会），其实一个字都没敲，不该弹候选。
        模块名取不到就带空串，调用方据此不做否决（绝不让工具整个哑掉）。
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
            try:
                mod = str(cm.Name or "")
            except Exception:
                mod = ""
            snap = (sl, cm.Lines(sl, 1), mod)
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
        self._caret_word = ""

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

    def names_outside_caret(self, caret, scope_only=False):
        """把光标处的词抹掉后，【当前模块】里还出现过的名字（小写集合）。

        `name_exists_outside_caret` 的批量版：一次读取模块文本，供引擎在剔除
        "回声候选"时对多个候选复用 —— 否则每个候选都要打一次 COM。

        刻意只读【当前模块】：回声只可能产生于正在编辑的这个模块；全工程扫描
        既没必要，又会在多模块大工程上明显变慢。

        scope_only=True（v51）：只收集【从光标处看得见】的出现 —— 模块级行
        （含过程/Type 声明头行、模块声明区）与光标所属过程的行；*别的过程体内
        的出现一律不算*。

        为什么需要这个窄窗口：用户正在【过程内声明一个名字】（形参 / 局部 Dim）
        时，这个名字是过程局部的。若现场证据仍是"全模块出现过"，那么"别的过程
        里的同名局部变量"就会被当成"这个名字真实存在"的铁证，回声防护随即放行，
        于是用户把别的过程的变量名原样敲/回退/粘贴到形参位置时，弹窗提示出自己
        ——正是用户报的"只有一模一样才会提示自己"。收窄到可见窗口后，那种出现
        不再是证据，候选照常按回声剔除。
        默认 False = 全模块扫描（模块级声明的场景仍走这一路：模块级名字本就
        全模块可见，任何出现都是它的合法引用，见 v37 语义）。
        """
        # 读不到就返回 None（= 现场证据不可用），与"真的没有名字"（空集）
        # 区分开 —— 调用方据此决定是保守放过还是照常判定。
        try:
            line_no, col = int(caret[0]), int(caret[1])
        except Exception:
            return None
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return None
            cp = vbe.ActiveCodePane
            if cp is None:
                return None
            cm = cp.CodeModule
            if cm.CountOfLines <= 0:
                return set()
            code = cm.Lines(1, cm.CountOfLines)
        except Exception:
            return None
        try:
            src = vba_parser._blank_ident_at(code, line_no, col)
            src = vba_parser._mask_strings_and_comments(src)
        except Exception:
            src = code
        try:
            if not scope_only:
                return set(m.group(0).lower()
                           for m in re.finditer(r"[^\W\d]\w*", src))
            rows = src.split("\n")
            if not (1 <= line_no <= len(rows)):
                # 行号对不上（理论上不该发生）：保守退回全模块扫描，
                # 宁可少剔除候选，也不要凭错的窗口把真名字误杀。
                return set(m.group(0).lower()
                           for m in re.finditer(r"[^\W\d]\w*", src))
            # 逐行标注归属：声明头行归属【它自己声明的那个过程】（与
            # parser._line_proc_map 一致），但它在下面按模块级收集 —— 过程的
            # 声明本身就是模块级语句，`Public Sub update()` 这种模块级过程名
            # 必须在这里找得到，否则在别的过程里就再也搜不到它了。
            owners = []
            cur = None
            for raw in rows:
                s = raw.strip()
                if vba_parser._RE_ANY_PROC_END.match(s):
                    cur = None
                m_head = (vba_parser._RE_SUB.match(s)
                          or vba_parser._RE_PROP.match(s))
                if m_head:
                    cur = m_head.group(1)
                owners.append((cur, bool(m_head)))
            keep = owners[line_no - 1][0] or ""
            if not keep:
                # 光标不在任何过程里（模块声明区）：没有可收窄的窗口。
                return set(m.group(0).lower()
                           for m in re.finditer(r"[^\W\d]\w*", src))
            kl = str(keep).lower()
            names = set()
            for (owner, is_header), raw in zip(owners, rows):
                if owner is None or is_header or str(owner).lower() == kl:
                    names.update(m.group(0).lower()
                                 for m in re.finditer(r"[^\W\d]\w*", raw))
            return names
        except Exception:
            return None

    def name_exists_outside_caret(self, name, caret):
        """把光标处的词抹掉后，【当前模块】里 name 是否还在别处出现。

        这是判断"正在输入的词是不是自己的回声"的最可靠依据：

        * 打 Name到一半/打全时，光标处那几个字符本身会被当成"用到过的隐式
          变量"收录进候选池。若这个名字在工程里再没别处出现，把它提示出来
          就是纯粹的自我提示（回退出来的 numA 就是典型）。
        * 反过来，只要它在别处真实存在（哪怕解析层没能识别出它的声明），
          那它就是个合法候选 —— 用户把它打全时**必须继续提示**，否则会出现
          "输入 a / ar 都提示 arr，一把 arr 打全列表就消失"。

        刻意只在【当前模块】里查：回声只可能产生于正在编辑的这个模块；
        全工程扫描既没必要，又会在多模块大工程上明显变慢。

        刻意只读文本做一次线性扫描（不触发标识符重解析），因此很廉价，
        可以每次按键都调用。
        """
        if not name:
            return False
        return str(name).lower() in (self.names_outside_caret(caret)
                                      or ())

    def caret_word_at_collect(self):
        """返回"收集标识符那一刻光标处的词"（小写，可能为空串）。

        标记符池是带缓存的：调用方拿到的候选可能来自 0.4 秒前的解析。那时
        光标处的词，就是用户此刻可能正在回退删除的那个词 —— 引擎用它把
        "回退途中残留的更长的旧片段"从候选里剔掉（v43）。
        """
        return getattr(self, "_caret_word", "") or ""

    def declared_elsewhere(self, name, module_name):
        """name 是否在【其它模块】里被真实声明过。

        True  = 别的模块声明过（跨模块 Public 证据）；
        False = 只在当前模块声明过，或压根没声明过 -> 请以现场文本为准；
        None  = 后端不提供这项信息 -> 调用方退回旧证据（旧式 / 测试后端）。

        分工：跨模块的 Public 名字只有声明集合认得它（现场扫不到），靠 True；
        而只在【正在编辑的这个模块】里声明过的名字，一律以现场文本为准 ——
        用户完全可能刚把它删掉，声明集合却还是 0.4 秒前解析的快照，信它就会把
        "回退途中残留的更长的旧片段"保下来（v43 修的就是这个）。
        """
        by_mod = getattr(self, "_declared_by_module", None)
        if not isinstance(by_mod, dict):
            return None
        m = str(module_name or "").lower()
        for owner in (by_mod.get(str(name).lower()) or ()):
            if owner != m:
                return True
        return False

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
        # 每次收集都重置：没有光标信息时它就是空串（不误伤任何候选）
        self._caret_word = ""
        _decl_by_mod = {}
        type_names = set()
        # 真实"声明"过的名字（Dim/Const/Sub/Function/Type/Enum 及组件名）。
        # 供引擎层区分"真名字"与"幻影"：回退删字回退出来的未定义词（如 numA）
        # 只可能是隐式残留 / 正在敲的词本身，引擎据此剔除"提示自己"。
        declared_names = set()
        # 结构性名字：组件名 + 窗体控件名（v60）+ 语言自带名字（v61/v62）。
        # 详见 __init__ 里的说明。
        structural_names = set()
        # 真正作为"VBA 内建名字"收录的那些（v61，小写）。与清单的差别：工程里
        # 已经自己声明过同名时以用户的为准，那些名字【不】算内建 —— 于是引擎
        # 对它们收紧匹配时不会连用户自己的定义一起收紧。
        builtin_names = set()
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
                        structural_names.add(mod_name.lower())
                        _decl_by_mod.setdefault(
                            mod_name.lower(), set()).add(mod_name.lower())
                    modules.append(mod_name)
                    # 窗体控件名（v60）：只收集【光标所在的那个窗体】的。
                    #   * 控件名只在它自己的窗体代码模块里能裸名引用（跨模块
                    #     得写 UserForm1.Label1），所以别的窗体的控件名收进来
                    #     也永远不会可见，白白占池子；
                    #   * 取一次控件要过 Designer 接口（真机实测约 23ms/窗体），
                    #     只对正在编辑的窗体取，开销小到可以忽略。
                    # 必须在读 CodeModule【之前】收集：新建窗体一行代码都没有，
                    # 落在 `count <= 0: continue` 之后就永远收不到了。
                    if (_is_msform(comp.Type) and caret_mod
                            and str(caret_mod).lower() == str(mod_name).lower()):
                        for _cn in _form_control_names(comp):
                            if not _RE_PLAIN_IDENT.match(_cn):
                                continue
                            records.append((_cn, mod_name, None, True))
                            declared_names.add(_cn.lower())
                            structural_names.add(_cn.lower())
                            _decl_by_mod.setdefault(
                                _cn.lower(), set()).add(str(mod_name).lower())
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
                    # 记下【这一刻光标处的词】：它多半是用户正在输入 / 正在
                    # 回退删除的词。回退删字时，缓存的池子里还留着它更长的
                    # 旧版本，引擎靠这个词把那个幽灵剔掉（v43）。
                    # 只在【确知光标在哪个模块】时才记：caret_mod 为空时
                    # 同一个 caret 会被套用到所有模块，这里最后赋值的那个
                    # 词可能来自任意模块，拿它当"光标旧词"就会误伤候选。
                    if apply_caret and caret_mod:
                        try:
                            self._caret_word = vba_parser.ident_at_caret(
                                code, apply_caret[0],
                                apply_caret[1]).lower()
                        except Exception:
                            self._caret_word = ""
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
                    for _r in recs:
                        _n = str(_r[0]).lower()
                        if _n not in caret_decl:
                            _decl_by_mod.setdefault(
                                _n, set()).add(str(mod_name).lower())
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

        # ---- VBA 语言自带的名字（v61）----
        # 用户报："VBA 本身自带的函数也加入提示列表，现在不能提示"。
        #
        # 这些名字由【语言运行时】决定存在，与工程代码文本无关，所以三件事都要做：
        #   1) 进候选池 —— 否则模糊匹配无从命中；
        #   2) 进 structural_names —— 引擎判"回声候选是不是真名字"时最先看它。
        #      打 ms 想补 MsgBox 时，ms 是 msgbox 的连续子串 -> 命中回声候选；
        #      而代码里若从没出现过 MsgBox，现场文本扫不到、工程里也没声明过，
        #      少了这一路证据候选必被踢掉 —— 症状正是"打 ms 死活不提示 MsgBox"；
        #   3) 进 _decl_by_mod（挂到 VBA 这个"别的模块"下），让 declared_elsewhere
        #      也认它们 —— 现场证据不可用（COM 读不到）时的兜底路径同样放行。
        #
        # 已声明过的同名不重复收录：工程里真有 Function Left(...) 时以用户的定义
        # 为准，免得列表里出现两条同名。
        if ENABLE_VBA_BUILTINS:
            for _bn in vba_builtins.BUILTIN_NAMES:
                _bl = _bn.lower()
                if _bl in declared_names:
                    continue
                records.append((_bn, _BUILTIN_MODULE, None, False))
                declared_names.add(_bl)
                structural_names.add(_bl)
                builtin_names.add(_bl)
                _decl_by_mod.setdefault(_bl, set()).add(_BUILTIN_MODULE)
            # 内建数据类型名：既进候选池（`As |` 位置的候选是从可见标识符里筛出来
            # 的，不在池子里就永远显示不出来），也进 type_names（那一位置只该冒
            # 类型名）。与函数清单重合的（Date / String 等）由上面的 declared_names
            # 去重挡掉，不会再收一遍。
            for _tn in vba_builtins.BUILTIN_TYPE_NAMES:
                _tl = _tn.lower()
                type_names.add(_tn)
                if _tl in declared_names:
                    continue
                records.append((_tn, _BUILTIN_MODULE, None, False))
                declared_names.add(_tl)
                structural_names.add(_tl)
                builtin_names.add(_tl)
                _decl_by_mod.setdefault(_tl, set()).add(_BUILTIN_MODULE)

        # ---- 语言关键字 / 保留字（v62）----
        # 用户要求："把 vba 里面的所有关键字都纳入到提示词里"。
        #
        # 走的是与内建名字**完全相同**的四路承接，原因也一样：关键字同样由语言
        # 本身决定存在，与代码文本无关。比如输入 if 想补 If —— 而此刻代码里恰好
        # 一个 If 都还没写过（新模块的第一行），现场文本扫不到、工程里也没声明过，
        # 少任何一路都会被回声防护当成幻影剔掉（v60 窗体名、v61 内建函数的同一个坑）。
        #
        # 唯一区别是开关独立（ENABLE_VBA_KEYWORDS）：关掉关键字不影响内建函数。
        # 关键字同样进 builtin_names —— 它们是语言自带的公共词汇（数量大、什么字母
        # 组合都凑得出子序列），要跟内建函数一样收紧模糊匹配，见 engine.trigger。
        if ENABLE_VBA_KEYWORDS:
            for _kn in vba_builtins.BUILTIN_KEYWORDS:
                _kl = _kn.lower()
                if _kl in declared_names:
                    continue
                records.append((_kn, _BUILTIN_MODULE, None, False))
                declared_names.add(_kl)
                structural_names.add(_kl)
                builtin_names.add(_kl)
                _decl_by_mod.setdefault(_kl, set()).add(_BUILTIN_MODULE)

        self._type_names = type_names
        self._declared_names = declared_names
        self._declared_by_module = _decl_by_mod
        self._structural_names = structural_names
        self._builtin_names = builtin_names
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

    def get_structural_names(self):
        """返回由【工程结构 / 语言运行时】决定其存在的名字（小写集合）。

        三类：组件名（模块 / 窗体 / 类 / 文档模块名）+ 窗体控件名 + VBA 语言自带
        的名字（内建函数 / 内建常量 / 内建数据类型 / 关键字）。

        与 get_declared_names 的区别：后者是"代码文本里真声明过什么"，前者是
        "本来就有、与代码文本无关"。用户刚拖上去的 Label1、一行代码都没有的新
        窗体 UserForm1、代码里从没调用过的 MsgBox、新模块第一行就想补的 If，
        都属于前者而不属于后者。

        引擎用它来判断"回声候选是不是真名字"——这类名字不能走"只在本模块声明
        过、现场文本又扫不到 -> 当幽灵剔掉"的规则（v60 组件名与控件名、v61 VBA
        内建名字：打 ms 要能补出从没写过的 MsgBox；v62 关键字：打 if 要能补出
        新模块里还没写过的 If）。
        """
        return set(getattr(self, "_structural_names", set()) or set())

    def get_builtin_names(self):
        """返回真正作为"语言自带词汇"收录的那些（小写集合，v61 / v62）。

        含：VBA 内建函数 / 内建常量 / 内建数据类型（v61）+ 语言关键字（v62）。

        与 vba_builtins 的清单不同：工程里已经自己声明过同名时以用户的定义为准，
        那些名字不会出现在这里 —— 引擎据此对"语言自带的公共词汇"收紧模糊匹配
        （纯分散命中要求输入至少 4 个字符），而不会连用户自己的同名定义一起收紧。

        后端不提供该接口时引擎不做任何收紧（旧式 / 测试后端行为不变）。
        """
        return set(getattr(self, "_builtin_names", set()) or set())

    def vbe_popup_visible(self):
        """VBE 自带的提示窗（参数信息 / 列出成员）此刻是否可见（v63）。

        ⚠️ v65 起语义澄清：这是"**任何**提示窗"，含参数信息（形参签名），
        只用于诊断；引擎真正据以让位的是 vbe_yield_visible()。保留本方法是为了
        与旧式 / 测试后端和诊断日志兼容。纯 Win32 查询、不碰 COM。
        """
        return vbe_popup_visible()

    def vbe_yield_visible(self):
        """VBE 的【成员列表】窗此刻是否可见 —— 引擎据此让位（v65）。

        用户口径（v65 澄清 v63）："只有弹成员列表的时候才让位，弹形参签名不需要
        让位。" 成员列表（NameListWndClass）与我们同质（也是候选列表、也吃键盘），
        让位；参数信息（PopupTipWndClass）只是只读签名，不该让我们消失。

        见模块级 vbe_yield_visible()。纯 Win32、不碰 COM，探不到一律 False
        （不让位）—— 后端不提供本方法时引擎退回 vbe_popup_visible() 的老行为。
        """
        return vbe_yield_visible()

    def vbe_popup_info(self):
        """当前可见的 VBE 提示窗明细 [(类名, 宽, 高), ...]（v65 诊断用）。

        与 vbe_popup_visible() 同一套判据，只是把命中的窗描述出来，让日志能
        区分【成员列表】(NameListWndClass) 与【参数信息】(PopupTipWndClass)。
        纯 Win32 只读，探不到返回 []。引擎只在真正要"让位"那一刻记一行日志，
        正常路径不调用，无开销。
        """
        return vbe_popup_info()

    def vbe_popup_rects(self):
        """当前可见的 VBE 提示窗矩形 [(类名, 左, 上, 右, 下), ...]（v66，UI 避让用）。

        UI 用它把候选窗挪到参数信息窗（形参签名）下方，两个窗不再重叠。
        自定义过程的参数信息用的是同一个窗，一并覆盖。见模块级同名函数。
        """
        return vbe_popup_rects()

    # ---- 自动配对：输入 ( / " 自动补右半边 ----
    #
    # VBE 原生【不会】自动闭合括号和引号（VBE_Extras / Rubberduck 都把它当增强
    # 功能往外加），所以这里补上不会出现"两边都补"的双份。唯一要注意的是 VBE
    # 会在【行尾按回车时】补齐缺失的右引号，那与逐字符输入无关，不受影响。
    _PAIR_CLOSE = {"(": ")", '"': '"', ")": ")"}

    @staticmethod
    def _write_line(cm, line_no, text):
        """把 text 写回第 line_no 行。返回是否成功。

        VBE 允许光标停在【模块最后一行的下一行】（一个尚不存在的"虚拟空行"），
        此时 `Lines()` 读出来是空串、看起来一切正常，但 `ReplaceLine()` 会抛
        "无效的过程调用或参数"。真机实测（模块 9 行、光标在第 10 行）确实如此。
        这一行得先用 `InsertLines()` 真的建出来，再写内容。
        """
        try:
            n = cm.CountOfLines
        except Exception:
            n = None                  # 拿不到行数（桩对象等）-> 按普通行处理
        try:
            if n is not None and line_no > n:
                if line_no != n + 1:
                    return False      # 差得离谱，不猜，交给调用方兜底
                cm.InsertLines(line_no, text)
            else:
                cm.ReplaceLine(line_no, text)
            return True
        except Exception:
            return False

    def insert_pair(self, open_ch):
        """在光标处插入一对符号（`(` -> `()`、`"` -> `""`），光标落在中间。

        返回 True = "这一下我代为输入了"（调用方据此吞掉原按键）；
        返回 False = "我没管"（在注释里 / 有选区 / COM 出问题），调用方必须
        **放行按键**让 VBE 按原生行为处理 —— 绝不吞掉用户的按键却什么都没发生。
        """
        close_ch = self._PAIR_CLOSE.get(open_ch)
        if not close_ch:
            return False
        new_line = None
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return False
            cp = vbe.ActiveCodePane
            if cp is None:
                return False
            cm = cp.CodeModule
            sl, sc, el, ec = cp.GetSelection()
            if sl != el or sc != ec:
                return False          # 有选区：替换语义交给 VBE，别插
            line_text = cm.Lines(sl, 1)
            col, sem, tabw, wide2 = _caret_char_col(
                cm, cp, sl, line_text, ec)
            col0 = max(0, min(col - 1, len(line_text)))
            if _classify(line_text, col0 + 1)[1]:
                return False          # 注释里不自动配对（VBE 也不）
            if open_ch == ")":
                # 右半边已经在光标右边 -> 只跨过去（补完 `()` 后再按 `)` 不该
                # 多出一个）；否则【放行】，让用户真的插入一个右括号。
                if col0 >= len(line_text) or line_text[col0] != ")":
                    return False
                new_line, caret_off = None, col0 + 1
            else:
                new_line, caret_off = _pair_insertion(
                    line_text, col0, open_ch, close_ch)
            if new_line is not None:
                if not self._write_line(cm, sl, new_line):
                    return False
                try:
                    actual = cm.Lines(sl, 1)
                except Exception:
                    actual = new_line
                if actual != new_line:
                    # VBE 把整行重新格式化过了（见 _locate_inserted 的说明）：
                    # 在改写后的行里重新定位我们插入的那对符号，把光标钉到
                    # 它中间；实在认不出来才退到行尾附近（绝不越界）。
                    off = _locate_inserted(actual, new_line, open_ch, close_ch,
                                           col0)
                    if off is not None:
                        caret_off = off
                    else:
                        caret_off = min(caret_off, len(actual))
            else:
                # 右半边已经在光标右边：只把光标挪过去，一个字符都不写
                actual = line_text
                caret_off = min(caret_off, len(actual))
            if sem == "disp":
                col = _disp_width(actual, tabw, wide2, upto=caret_off) + 1
            else:
                col = caret_off + 1
            cp.SetSelection(sl, col, sl, col)
            return True
        except Exception:
            return False
        finally:
            # 与 apply_completion 同理：写过代码（VBAProject 变脏），立刻放手。
            if new_line is not None:
                _release_vbe_proxy()

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
            ec_char, sem, tabw, wide2 = _caret_char_col(
                cm, cp, sl, line_text, ec)
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
            # 只在声明行（含 v53 的续行）上算，平时为空、开销极小。
            decl_names = []
            # v53：声明【续行】也算声明行。形参常常独占一行（长签名被拆成
            # `Sub Foo( _` + 缩进的形参行），续行本身没有 Sub/Dim 关键字，
            # is_caret_in_declaration() 看不出来；但用户在那里打 / 退格 /
            # 粘贴形参名，与在首行上是同一件事——不把这条逻辑行的声明名算出来，
            # 引擎就没法拦住"形参提示形参自己"（用户报的正是续行上的退格）。
            # 判定很便宜：只额外读上一行，看它是否以续行符 `_` 结尾。
            cont_decl = False
            if not in_decl and sl > 1:
                try:
                    cont_decl = vba_parser.is_continuation_line(
                        cm.Lines(sl - 1, 1))
                except Exception:
                    cont_decl = False
            if in_decl or cont_decl:
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
    def apply_completion(self, line_no, word_start_col, end_col, completion):
        """把 [word_start_col, end_col) 区间的单词替换为 completion 并写回 VBE。

        word_start_col / end_col 均为 1-based 字符列（与 Python 行文本一致）。
        end_col 通常是光标列（光标左边那截已输入的前缀）；光标停在标识符【开头】
        （按 Delete 从名字开头删字符，v44）或【中间】（从名字中间删字符，v46）时
        它才大于光标列——此时要把光标右边那段残留名字一起换掉，否则会在残留词
        前面插入候选、拼出双份名字。

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
            new_line = replace_word(line_text, word_start_col, end_col, completion)

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
            end0 = _skip_into_parens(actual, end0, line_text, end_col)
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
                end0b = _skip_into_parens(actual2, end0, line_text, end_col)
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

    # ---- 新起一行（Shift+Enter） ----
    def new_line_below(self):
        """在光标所在行的【下方】新起一行：缩进与上一行对齐，光标落在缩进之后。

        等价于"先把光标移到本行末尾，再按回车"，但一步到位 —— 而且当前行
        【不拆分】（光标右侧的代码留在原行）。VBE 只在"光标已在行尾"时按回车
        才继承上一行缩进；光标停在行中间时按回车会把行拆开，所以必须由我们代劳。

        实现上刻意不依赖光标列，也不移动光标：直接读本行文本、取它的行首空白
        作为新行内容，再把光标放到新行缩进之后。

        成功返回 True；不在代码窗 / 取不到 COM / 出任何异常都返回 False ——
        调用方据此"不吞键"，让 VBE 按原生行为处理，绝不吞掉用户的换行。
        """
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return False
            cp = vbe.ActiveCodePane
            if cp is None:
                return False
            cm = cp.CodeModule
            sl, _sc, el, _ec = cp.GetSelection()
            # 有选区时以选区【末行】为基准（正常情况下 sl == el）。
            anchor = max(int(sl), int(el))
            if anchor <= 0:
                return False
            indent = _leading_ws(cm.Lines(anchor, 1))
            # InsertLines 在 anchor+1 处插入（anchor == CountOfLines 时即追加到末尾），
            # 已有行自动下移 —— 正是"在本行下方新起一行"。
            cm.InsertLines(anchor + 1, indent)
            col = _indent_end_col(indent)
            cp.SetSelection(anchor + 1, col, anchor + 1, col)
            return True
        except Exception:
            return False
        finally:
            # 与 apply_completion 同理：写过代码（VBAProject 变脏），立刻放开
            # 手里的 COM 代理，别拖住 Excel 的退出/写回。
            _release_vbe_proxy()
