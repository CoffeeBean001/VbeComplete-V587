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


def _env_flag(name, default):
    """三态环境变量开关：`1/true/yes/on` -> True，`0/false/no/off` -> False，
    未设置 / 认不出来 -> default。

    用意：下面那些"收不收某类名字"的开关，默认值写死在代码里（跟着用户口径走），
    但要能不改源码就翻过来 —— 双击 run.bat 之前先 `set XXX=1` 即可。
    认不出的值一律回落到 default（绝不因为写错一个环境变量就把功能关掉）。
    """
    raw = os.environ.get(name)
    if raw is None:
        return default
    v = raw.strip().lower()
    if v in ("1", "true", "yes", "on"):
        return True
    if v in ("0", "false", "no", "off"):
        return False
    return default


# 未声明即使用的"隐式变量"（不写 Option Explicit 时的用法）是否纳入提示。
# 关闭后行为回到旧版：只提示 Dim/Const/Sub 等显式声明出来的名字。
ENABLE_IMPLICIT_IDENTIFIERS = True
# 隐式变量的作用域粒度：
#   "module" —— 本模块任意位置都能提示（最宽松，但会跨过程泄漏同名变量）
#   "proc"   —— 按"首次出现所在过程"限定作用域（更贴近 VBA 语义，避免跨过程泄漏）
IMPLICIT_SCOPE = "proc"

# ★v75：隐式变量要不要【尊重模块自己写的 Option Explicit】。
#
# 用户口径："没有 Option Explicit 时，用过的变量名就算已经被定义、可以直接提示；
# 写了 Option Explicit 是强制声明，没声明过的变量名在后续就不该被提示。"
#
# 判定是【按模块】做的 —— Option Explicit 本来就是模块级语句：
#   * 没写它的模块：完全保持原样（用过的名字照旧进池子，用户明确说这样好）；
#   * 写了它的模块：只收"真声明过"的名字（Dim / Const / Sub / Function / Type /
#     Enum / 形参）；"只用过、没声明"的名字不再进池子 —— 那种写法在该模块里
#     是编译错误，提示它只会带出坏代码。
#
# 想回到旧行为（不分模块，隐式变量一律收）：
#   set VBECOMPLETE_IMPLICIT_HONOR_EXPLICIT=0
IMPLICIT_HONOR_OPTION_EXPLICIT = _env_flag(
    "VBECOMPLETE_IMPLICIT_HONOR_EXPLICIT", True)

# VBA 语言自带的名字（内建函数 / 内建数据类型）是否纳入提示。
#
# ★v77：**默认关闭**（用户口径）。原话："我觉得提示词太多了……你把提示 vba 本身带的
# 关键字、函数这块都删掉吧。保留能提示变量名、自定义的函数/过程名、窗体/控件名、
# 模块名。其他的那些个提示词我也不怎么用到，提示一大堆看起来也不舒服。"
#
# 于是默认口径收敛成一句话：**只提示"你这个工程里存在的东西"** ——
#   变量名（含没写 Option Explicit 时用过的名字）、自定义 Sub/Function/Property、
#   组件名（模块 / 窗体 / 类模块）、窗体控件名；
#   VBA 自带的四批（内建函数 + 数据类型、语言关键字、vb* 枚举、宿主 xl*/mso* 枚举）
#   一律不进候选池。
#
# ⚠️ 只改【收集侧开关】，四份清单本身一个字都没裁（见 vba_builtins.py）——
# 清单保持完整，测试/诊断仍能看到全貌，需要时一行环境变量立刻收回来。
# 想开回来：`set VBECOMPLETE_VBA_BUILTINS=1`（内建函数与 `Dim x As <类型>` 一起回来）。
ENABLE_VBA_BUILTINS = _env_flag("VBECOMPLETE_VBA_BUILTINS", False)

# ★v70：VBA 内建的 **枚举常量**（`vbOK` / `vbCrLf` / `vbYes` … 共 102 条）是否纳入提示。
#
# 用户口径（v70）："我不怎么使用 VBA 内置枚举，提示出来对我有干扰 —— 把枚举关掉，
# 函数那些要保留。" 所以它与"内建函数"【分家】：函数（MsgBox / Left / Split …）
# 与内建数据类型（Long / String …）照常提示，只把这一批枚举常量摘出去。
# 想开回来：`set VBECOMPLETE_VBA_CONSTANTS=1`。
ENABLE_VBA_CONSTANTS = _env_flag("VBECOMPLETE_VBA_CONSTANTS", False)

# 语言关键字 / 保留字（Sub / Dim / If / For / Set / And …）是否纳入提示（v62）。
#
# 独立成一个开关（而不是并进 ENABLE_VBA_BUILTINS）：关键字与内建函数是两拨
# 东西 —— 前者是语法骨架、后者是可调用的库成员。你想单收一批时用得上。
#
# ★v77：**默认关闭**（用户口径："vba 本身带的关键字、函数这块都删掉"）。
# 想开回来：`set VBECOMPLETE_VBA_KEYWORDS=1`。
# 与 ENABLE_VBA_BUILTINS 是两个独立开关，所以"只收关键字、不收内建函数"也做得到。
ENABLE_VBA_KEYWORDS = _env_flag("VBECOMPLETE_VBA_KEYWORDS", False)
# 内建名字挂靠的"模块名"。
#
# 这是一个【虚拟模块】——VBA 运行时库。之所以要给它一个名字而不是留空：
#   1) 这些名字在工程里任何模块都能裸名引用，按模块级 priv=False 收录即可；
#   2) 引擎判"回声候选是不是真名字"时会问 declared_elsewhere（名字是不是
#      在【别的模块】声明过）。把内建名字的归属记成 VBA，那一问天然为真，
#      内建名字就不会被"只在本模块声明过 + 现场文本扫不到 -> 当幽灵剔掉"
#      这条规则误杀（详见 engine._name_really_exists）。
_BUILTIN_MODULE = "VBA"

# 宿主类型库的枚举常量是否纳入提示（v67）。
#
# 用户报："输入 vb 会提示一堆 VBA 枚举值，输入 xl 却不提示任何 xl 开头的枚举值。"
# 前者是 v61 收的 VBA 内建常量（vb*），后者是【宿主 Excel 类型库】的枚举常量
# （xl*）—— 那是一整套库，从来没收过，于是同一个动作在 xl 上一个提示都没有。
#
# 这些常量与内建名字同源：由【工程引用的类型库】决定存在，与代码文本无关
# （代码里从没写过 xlUp，打 xl 也该补得出来）。所以走完全相同的四路承接。
#
# 与内建名字的关键区别是【匹配口径】—— 它们数量极大（Excel + Office 两个库
# 实测 4538 条），若像工程内名字那样接受跳步匹配，输入 3~4 个字母就会带出
# 几百条无关项。所以：
#   * 只收"有统一家族前缀"的库（Excel -> xl，Office -> mso），
#     没有家族前缀的库（stdole 的 Checked / Gray / Color 这种通用词）一概不收；
#   * 只认【从头开始的连续前缀】命中，且输入长度不低于家族前缀长度 ——
#     于是 xl -> xl*、xlu -> xlUp；而 ms 只出 MsgBox，绝不带出 mso*。
#
# ★v70：**默认关闭**（用户口径："xl 开头这些枚举也不要提示，我不怎么用"）。
# 想开回来：`set VBECOMPLETE_HOST_ENUMS=1`。关着的时候收集段整段不跑
# （连类型库都不会去加载，零开销），引擎那边也拿不到宿主清单。
ENABLE_HOST_ENUMS = _env_flag("VBECOMPLETE_HOST_ENUMS", False)
# 家族前缀的覆盖率门槛：某前缀能盖住该库这么多比例的枚举成员，才算"家族前缀"。
_HOST_ENUM_FAMILY_PCT = 0.5
# 允许的家族前缀白名单（小写，逗号分隔）。留空 = 自动推导（推荐）。
#   例：想只留 Excel 的 xl*，设 VBECOMPLETE_HOST_ENUM_FAMILIES=xl
_HOST_ENUM_FAMILIES = tuple(
    p.strip().lower() for p in
    os.environ.get("VBECOMPLETE_HOST_ENUM_FAMILIES", "").split(",") if p.strip())
# 宿主常量挂靠的虚拟模块名（同 _BUILTIN_MODULE 的道理，见那边的说明）。
_HOST_ENUM_MODULE = "宿主库"

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


# ---- 回车后的自动缩进（v79） ----
#
# 用户口径两条：
#   1) "我在一整行写入了注释，然后按回车换行后，光标会默认在行首、不会缩进；
#      帮我加个能忽略上方注释行（可能有多行注释）、跟上方第一个非注释行对齐的"；
#   2) "写 for / do / if / with / sub / function / type / enum 这些结构时，
#      回车能自动缩进"（块的开头行之后，新行要深一级）。
#
# 下面几个都是【纯函数】：不碰 COM，因此既能被单测直接钉住，也能在键盘钩子
# 线程里安全调用 —— 钩子线程不许碰 COM（v56 的教训），而"这一下回车要不要由
# 我们接管"必须在钩子线程里就定下来（否则要么吞了键才后悔，要么每次都吞）。

# 块结构【开头】：这一行往下的内容都属于新的一级，回车换行后新行要缩进一级。
#
# 判定只看两件事，不做语法分析：
#   * "以 Then 结尾" —— 覆盖 `If x Then` / `ElseIf x Then` / `#If VBA7 Then`；
#     单选 If（`If x Then y = 1`）以别的词结尾，天然被排除在外；
#   * "行首关键字在白名单里" —— Sub / Function / Property / Type / Enum /
#     For / Do / While / With / Else / Case / Select Case。
# 闭合行（Next / Loop / Wend / End If / End Sub …）刻意不在表里：它们本来就
# 缩进在正确的层级上，新行跟它们齐平就是对的结果。
_BLOCK_FIRST_WORDS = ("sub", "function", "property", "type", "enum",
                      "for", "do", "while", "with", "else", "case")
# 可以出现在声明关键字前面的修饰词，判断"行首关键字"时要跳过它们
# （`Private Sub Foo()` / `Public Function Bar()` / `Static Sub`）。
_MOD_FIRST_WORDS = ("public", "private", "friend", "static")
# 行尾就是 Then（前面得有分隔，别把 `Something` 认成 Then）
_RE_TAIL_THEN = re.compile(r"(?:^|\s)then\s*$")

# 块结构【收尾】（v79b 用户新要求：写 `If i = 1 Then` 按回车，后面自动补 `End If`）。
# 键是 _block_kind 认出来的规范化类型名，值是补在下一行（缩进回到块头同级）的收尾。
#   * 只有"真需要收尾"的块才在这儿 —— Else / ElseIf…Then / Case 是块【内】的分支，
#     往下一行只是同级的分支体，不该替用户写收尾；
#   * For 用 Next（VBA 允许不带变量名）、Do 用 Loop、While 用 Wend；其余统一 End Xxx；
#   * 条件编译（#If / #Else）刻意不在表里：它必须配 `#End If`，替用户写个裸的
#     `End If` 反而把代码写错，宁可让用户自己敲。
_BLOCK_CLOSERS = {
    "sub": "End Sub",
    "function": "End Function",
    "property": "End Property",
    "type": "End Type",
    "enum": "End Enum",
    "if": "End If",
    "with": "End With",
    "select": "End Select",
    "for": "Next",
    "do": "Loop",
    "while": "Wend",
}


def _code_part(line_text):
    """一行代码里"去掉字符串与注释之后"的部分（行尾空白也去掉）。纯函数。

    用 parser 里那份掩码（字符串字面量 + `'` 注释换成等长空格），因此
    `If x = "a Then" Then` 里字符串中的 Then 不会被误当成块头，
    `If x Then  '注释` 尾部的注释也不会破坏"以 Then 结尾"的判定。
    整行 `Rem ...` 注释按空串处理。
    """
    t = line_text or ""
    if not t.strip():
        return ""
    _st = t.lstrip().lower()
    if _st == "rem" or _st.startswith("rem ") or _st.startswith("rem\t"):
        return ""
    try:
        return vba_parser._mask_strings_and_comments(t).rstrip()
    except Exception:
        return t.rstrip()


def _is_comment_only(line_text):
    """整行是不是注释（`'...` 或 `Rem ...`）—— 一行里没有任何代码。纯函数。"""
    t = (line_text or "").strip()
    if not t:
        return False
    if t.startswith("'"):
        return True
    low = t.lower()
    return low == "rem" or low.startswith("rem ") or low.startswith("rem\t")


def _block_kind(line_text):
    """这一行是【哪一种】块结构的开头？返回规范化的类型名；不是块头返回 None。

    返回值的三种含义：
      * `"branch"`    —— Else / ElseIf…Then / Case：要缩进一级，但**没有收尾**；
      * `"directive"` —— #If / #Else 条件编译：要缩进一级，收尾必须由用户写
                        `#End If`，我们不代劳（见 _BLOCK_CLOSERS 的说明）；
      * 其余（sub / function / property / type / enum / if / with / select /
        for / do / while）—— 都在 _BLOCK_CLOSERS 里有对应的收尾行。
    """
    code = _code_part(line_text)
    if not code:
        return None
    if code.endswith("_"):
        # 续行符：这条语句还没写完，下一行是它的延续，不是"块里的第一行"。
        return None
    low = " ".join(code.lower().split())
    hashed = low.startswith("#")
    words = low.split(" ")
    i = 0
    while i < len(words) and words[i] in _MOD_FIRST_WORDS:
        i += 1
    first = words[i].lstrip("#") if i < len(words) else ""
    if first in ("else", "elseif", "case"):
        if first == "elseif" and not _RE_TAIL_THEN.search(low):
            return None                    # 只写 `ElseIf x` 没写 Then —— 不算块头
        return "directive" if hashed else "branch"
    if _RE_TAIL_THEN.search(low):
        # `If x Then` / `ElseIf x Then` / `#If VBA7 Then`：行尾就是 Then。
        # 单选 If（`If x Then y = 1`）以别的词结尾，天然被排除在外。
        return "directive" if hashed else "if"
    if first == "select":
        ok = len(words) > i + 1 and words[i + 1] == "case"
        return ("directive" if hashed else "select") if ok else None
    if first in _BLOCK_FIRST_WORDS:
        return "directive" if hashed else first
    return None


def opens_block(line_text):
    """这一行是不是"块结构的开头"——回车换行后新行该缩进一级。纯函数。

    例：`For i = 1 To 10` / `For Each c In rng` / `Do While x` / `If x Then` /
    `ElseIf y Then` / `Else` / `Case 1` / `Select Case x` / `With rng` /
    `Sub Foo()` / `Private Function Bar() As Long` / `Property Get X` /
    `Type Foo` / `Enum Color` / `#If VBA7 Then` / `#Else` -> True；
    `If x Then y = 1`（单选 If）、`Next i`、`End If`、`Loop`、`Wend`、
    `x = 1`、`Exit For`、`DoEvents`、`foo(1, 2)`、`Debug.Print x`、
    以续行符 `_` 结尾的行 -> False。
    """
    return _block_kind(line_text) is not None


def block_closer(line_text):
    """块头那行【该补的收尾】；不需要收尾时返回 None。纯函数。

    `Sub Foo()` -> "End Sub"、`If x Then` -> "End If"、`With rng` -> "End With"、
    `Select Case x` -> "End Select"、`For i = 1 To 10` -> "Next"、`Do` -> "Loop"、
    `While x` -> "Wend"、`Private Function F()` -> "End Function"；
    `ElseIf y Then` / `Else` / `Case 1` / `#If VBA7 Then` / 整行注释 /
    普通行 -> None（它们要么是块内的分支、要么该由用户自己写收尾）。
    """
    return _BLOCK_CLOSERS.get(_block_kind(line_text))


def _indent_width(raw):
    """行首空白占几个字符位（Tab 按 4 列算）。纯函数。"""
    n = 0
    for ch in (raw or ""):
        if ch == " ":
            n += 1
        elif ch == "\t":
            n += 4
        else:
            break
    return n


def _is_closer_text(line_text, closer):
    """这一行是不是 closer 那一行（`End If` / `Next i` / `Loop While x`…）。纯函数。

    先用 _code_part 去掉注释：`Next  '下一个` 也算 `Next`。
    """
    if not closer:
        return False
    s = _code_part(line_text).strip().lower()
    if not s:
        return False
    want = closer.strip().lower()
    return s == want or s.startswith(want + " ")


def _closer_owner_above(lines_above, kind, indent):
    """上方有没有一个【同类、仍未闭合】的块头，正好缩进在 indent 列？（纯函数）

    lines_above: 块头行【上方】各行的文本（越靠近它的越靠后）
    kind:        _block_kind 认出来的类型名（for / if / with / do / while /
                 select / type / enum / sub / function / property）
    indent:      待判定的收尾行缩进宽度

    只干一件事：回答"下面那个更浅的同类收尾，是外层块（嵌套的 For/If/Do/With…）
    的还是本块的"。从近到远扫，用"同类收尾 +1 / 同类块头 -1"的括号配对法，
    遇到一个"还开着"的同类块头时，看它缩进是不是正好等于那个收尾的缩进
    —— 是，就说明那个收尾属于它。

    为什么需要（v79c，用户报的双层 For）：内层 `For` 下面压着的是【外层】的
    `Next`，缩进比内层浅。单看缩进"比块头浅就补"是对的，但用户自己把收尾写歪
    一级的情形也长这样 —— 再补一条就成 `Next` 双份（编译报错）。加上这一问，
    两种情形就分开了：外层真有一个同类块头在那一列 -> 补；没有 -> 不补。
    """
    closer = _BLOCK_CLOSERS.get(kind or "")
    if not closer:
        return False
    depth = 0
    for raw in reversed(list(lines_above or ())):
        if not (raw or "").strip():
            continue
        if _is_closer_text(raw, closer):
            depth += 1                         # 更近处已经闭合了一层
            continue
        if _block_kind(raw) != kind:
            continue
        if depth == 0:
            return _indent_width(raw) == int(indent)
        depth -= 1
    return False


def _own_closer_below(lines_below, closer, kind, base):
    """下面是不是已经挂着【本块自己的】收尾？（v80a，纯函数）

    lines_below: 光标行下方各行（越靠近它的越靠前）
    closer:      本块的收尾文本（`Next` / `End If` / `End Type`…）
    kind:        `_block_kind` 认出来的块头类型
    base:        块头那行的缩进【宽度】

    与 `_closer_owner_above` 正好对称的一问：那一问管"下方第一个非空行是
    【外层块】的收尾"，这一问管"下方压着的那些同级代码里，到底有没有本块的
    收尾"。只看【和块头同一缩进】的行，用"同类块头 +1 / 同类收尾 -1"配平：

      * 同缩进的同类块头（`For` 里嵌 `For`、用户还没缩进）-> 计一层，它的
        `Next` 只闭合它自己；
      * 配平到 0 时遇到的同类收尾 -> 就是【本块】的，下面已经挂着了；
      * 同缩进的其它东西（`If` / `With` / `x = 1` / 别的收尾）-> 本块早就
        结束了（那一行就是它的边界），收尾不在下面；
      * 缩进比块头浅的行 -> 已经出了这一层，收尾不在下面；
      * 缩进更深的行 -> 属于嵌套块，与"本块的收尾在哪"无关，跳过；更深的
        同类块头 / 收尾另用 deeper 配平（免得把内层的 `Next` 当成"用户写歪
        的收尾"而少补一条）；
      * 空行 / 整行注释 / 块内分支行（`Else` / `ElseIf…Then` / `Case…`，
        含 `#Else` / `#ElseIf…Then`）-> 都不是"本块到此为止"的边界，跳过。

    为什么要这一问（v80a，用户报的）：先写了一个 `For`（回车自动补了 `Next`），
    再回到它【上面】补一个新的 `For` 并回车 —— 新 `For` 下面压着的正是原来那个
    `For`（同级）。旧口径"同级一律不补（只有模块级缩进 0 才补）"，于是新 `For`
    的 `Next` 永远补不出来，用户看到的就是"只换行缩进、不补 `Next`"。

    ★v82（用户报的"`If…Then` 补过一次 `End If`，写完 `Else` 再回到 `Then` 那行
    回车又补一个"）：`Else` / `Case` 这类**分支行**本来就在本块【里面】，可 v81
    起分支行会被拉回块头那一带（`Else` 与 `If` 齐平；`Case` 从 v83 起比 `Select
    Case` 深一级）—— 齐平的那种长得跟"同级的别的代码"一模一样，旧判据一眼认定
    "本块已结束了"，收尾就当没看见，回车再补一条。注释同理（注释根本不参与块
    结构）。这两类行现在一律跳过，扫描继续往下走 —— 跳过是**先跳后判缩进**，
    所以分支行放在哪一级都不影响这一问（第 63.8 节钉着）。
    """
    depth = 0        # 同缩进的同类块头（本块下面嵌着的同类块）
    deeper = 0       # 真正嵌套（缩进更深）的同类块头
    for raw in (lines_below or ()):
        if not (raw or "").strip():
            continue
        if _is_comment_only(raw) or _branch_word(raw) is not None:
            # ★v82：注释 / 块内分支行都不是"本块到此为止"的边界（详见 docstring）。
            continue
        w = _indent_width(raw)
        if w < base:
            return False                       # 出了这一层：本块收尾不在下面
        if _is_closer_text(raw, closer):
            if w == base:
                if depth == 0:
                    return True                # 这个收尾就是本块的
                depth -= 1
            elif deeper > 0:
                deeper -= 1                    # 闭合的是更深的那个同类块
            else:
                # 孤零零一个更深的同类收尾：位置不对，但它是用户自己写的 ——
                # 保守起见不再补（宁可少补，也不要 `Loop` / `Next` 双份把代码
                # 写坏）。这与 v79b 的老口径一致。
                return True
            continue
        if _block_kind(raw) == kind:
            if w == base:
                depth += 1                     # 同缩进的同类嵌套块头
            else:
                deeper += 1
            continue
        if w == base:
            return False                       # 同缩进的其它东西 -> 本块已结束
    return False


def closer_needed(lines_below, closer, base_indent, kind=None, lines_above=()):
    """下面到底要不要补 closer（v79b；嵌套判定 v79c；同级判定 v80a，纯函数）。

    lines_below: 光标行【下方】各行的文本（越靠近它的越靠前）
    closer:      block_closer 算出来的收尾文本（空 / None -> 不补）
    base_indent: 块头那行的行首缩进【宽度】
    kind:        块头类型（可选；给了才能判"更浅的同类收尾是不是外层块的"）
    lines_above: 块头行上方各行（可选，同上）

    传了 kind（生产路径上一定会传）时的口径：
      * 下方【没有】本块自己的收尾 -> 补（怎么算"没有"，见 _own_closer_below：
        它按缩进分类扫下方各行，同级同类块头配平、同级别的东西算边界、
        空行 / 注释 / 分支行跳过、更深的行跳过）；
      * 下方第一个非空行就是本块的同类收尾、且缩进比块头【更浅】-> 再问一句
        `_closer_owner_above`："那是外层块的收尾（本块该补）还是用户写歪的
        （不补）"；
      * 其余 -> 不补。

    不传 kind 时保持 v79b 老口径（只看"下方第一个非空行"）：
      * 它比块头缩进得【更浅】-> 补（但若它就是本块的收尾且更浅，一律不补）；
      * 它跟块头齐平或更深 -> 不补；唯一的例外是【块头在模块级（缩进 0）】，
        那里"齐平"的只可能是下一个声明/过程头，收尾该补；
      * 下方根本没有非空行（文件到底了）-> 补。
    空行一律跳过（下方可能隔着一堆空行才挂着真正的收尾）。

    不传 kind/lines_above（旧式调用）时，"更浅的同类收尾"一律按【不补】处理 ——
    这是最保守的一档，与 v79b 完全一致。

    ★v80a 变更：传了 kind 时，"同级"一档不再是"只有模块级（缩进 0）才补"，
    改成问一句 `_own_closer_below` —— "下面到底有没有本块自己的收尾"。
    别的分档一律不动。

    ★v82 变更（不再多补一条收尾）：`_own_closer_below` 现在把【整行注释】和
    【块内分支行】也跳过 —— `Else` / `ElseIf…Then` / `Case…` 是块内的分支，
    不是"本块已结束"的边界。旧判据把它们当边界，于是 `If…Then` 补过一次
    `End If`、写完 `Else` 后再回到 `Then` 那行回车，会【再补一个】`End If`
    （用户报的正是这个；`Select Case` 里的 `Case` 完全同理）。这条只让"补"
    更保守，别的分档一字未动。
    """
    if not closer:
        return False
    base = int(base_indent)
    first = None
    for raw in (lines_below or ()):
        if (raw or "").strip():
            first = raw
            break
    if kind is not None:
        # 【v80a 口径】下面挂着本块自己的收尾 -> 不补；没有 -> 补。
        # 这一问把"同级代码"分成三类：同类嵌套块头（配平）、本块的收尾、
        # 以及"本块早就结束了"的边界 —— 详见 _own_closer_below。
        if _own_closer_below(lines_below, closer, kind, base):
            return False
        if first is not None:
            w = _indent_width(first)
            if _is_closer_text(first, closer) and w < base:
                # 更浅的同类收尾：可能是【外层块】的（双层 For 的 `Next`），
                # 也可能是用户自己写歪的 —— 分这一下的活交给它（v79c）。
                return _closer_owner_above(lines_above, kind, w)
        return True
    # 【老口径】kind 没传（v79b 那套调用方式）：一个字不动，第 58.9 节钉着。
    if first is None:
        return True
    w = _indent_width(first)
    if _is_closer_text(first, closer):
        if w >= base:
            return False                       # 本块的收尾已经挂在下面了
        # 比块头浅：要么是外层块的收尾（本块该补），要么是用户写歪的
        return _closer_owner_above(lines_above, kind, w)
    if w != base:
        return w < base
    # 同级、且不是本块的收尾：模块级 -> 本块还没内容，补；过程内 -> 维持
    # v79b 老口径（不补），免得把用户正在写的块体顶开。
    return base == 0


def next_line_indent(cur_line, lines_above=(), unit="    "):
    """回车换行后，新行该有的行首缩进（纯函数）。

    cur_line:    光标所在行的文本
    lines_above: 它上方各行的文本（越靠近它的越靠后）
    unit:        一级缩进（默认 4 个空格）

    规则（用户 v79 口径）：
      * 参考行 = cur_line 自己；但当它是【整行注释】或空行时，往上找第一个
        既不是整行注释也不是空行的行 —— 这就是"忽略上方注释行（可能有多行），
        跟上方第一个非注释行对齐"；
      * 新行缩进 = 参考行的行首空白；参考行若是【块结构开头】，再深一级；
      * 连参考行都找不到（首行 / 上方全是注释与空行）→ 就用 cur_line 自己的
        行首空白，也就是 VBE 的原生行为 —— 绝不比原来更差。
    """
    ref = None
    if (cur_line or "").strip() and not _is_comment_only(cur_line):
        ref = cur_line
    else:
        for t in reversed(list(lines_above or ())):
            if (t or "").strip() and not _is_comment_only(t):
                ref = t
                break
    if ref is None:
        return _leading_ws(cur_line)
    base = _leading_ws(ref)
    return base + unit if opens_block(ref) else base


# ── v81：块【内】分支行自动对齐（Else / ElseIf…Then / Case… / #Else / #ElseIf）
# 用户口径："我在写了一行 if 之后，回车会自动缩进和补充 end if，然后我再写 else，
# 此时的 else 是缩进状态，不能自动对齐上面的 if。我想在我打完 else，或 elseif
# （if 里的多分支）再按回车后，else / elseif 能够自动对齐上方的 if。其他类似的结构，
# 比如 select case，里面可以多分支的，都帮我做成这样能自动对齐的。"
#
# 为什么会歪：`Else` 是 _block_kind 认的块头（要缩进一级），所以这一下回车被我们
# 接管了 —— 而 VBE 原生回车【会】把刚敲完的那一行拉回它该在的层级，我们一接管，
# 那个自动对齐就没机会跑。于是 `Else` 一直停在"块体那一级"（上面自动补完 End If 后
# 光标停的那一行）。这里补的正是这一下：回车时先把分支行拉回它该在的层级，
# 再让新行深一级 —— 就是 VBE 原生会做的那件事。
#
# 归属关系在 VBA 里没有歧义：Else / ElseIf…Then 属于最近的还开着的 `If … Then`，
# Case… 属于最近的还开着的 `Select Case`，#Else / #ElseIf 属于最近的还开着的 `#If … Then`。
_BRANCH_OWNER = {"else": "if", "elseif": "if", "case": "select"}

# ★v83：分支行对齐到【块头往下第几级】。归属者一样，但层级不一样 ——
#     If x Then
#         ...
#     Else                      <- 与 `If` 【齐平】（0）
#     End If
#     Select Case x
#         Case 1                <- 比 `Select Case` 【深一级】（1）
#             ...
#     End Select
# 这是 VBA 的惯用缩进，也正是 VBE 原生回车会整理成的样子。v81 把 `Case` 也拉到
# `Select Case` 同级了（用户报的："case 应该是相对于 select 缩进的"）；`#Else` /
# `#ElseIf` 与 `#If … Then` 齐平、不分深浅，所以条件编译这一档同样是 0。
_BRANCH_OWNER_DEPTH = {"else": 0, "elseif": 0, "case": 1}
# 条件编译的收尾。_BLOCK_CLOSERS 里刻意没收 `#If`（那个收尾必须由用户自己写），
# 但"往上找还开着的 #If"这一步照样要用到它，所以单独放在这里。
_DIRECTIVE_CLOSER = "#End If"


def _directive_word(line_text):
    """`#` 开头的条件编译指令的第一段（不含 `#`）：`#If VBA7 Then` -> 'if'。纯函数。"""
    code = _code_part(line_text)
    if not code:
        return None
    low = " ".join(code.lower().split())
    if not low.startswith("#"):
        return None
    rest = low[1:].strip()
    return rest.split(" ")[0] if rest else None


def _branch_word(line_text):
    """这一行是哪种【块内分支】？返回 'else' / 'elseif' / 'case'；不是 -> None。纯函数。

    与 _block_kind 的 'branch' 一档严格对应（`#` 开头的条件编译分支返回同一个词，
    调用方再按"有没有 `#`"决定去找哪类块头）：
      * `Else` / `Case 1` / `Case Else` / `Case Is > 3` -> 'else' / 'case'；
      * `ElseIf y Then`          -> 'elseif'（少了 Then 不算，与 _block_kind 同口径）；
      * `Else If y Then`         -> 'else'   （VBA 里这是"Else + 一行新 If"，不是 ElseIf）；
      * `End If` / `Next` / 普通行 / 续行（行尾 `_`）-> None。
    """
    code = _code_part(line_text)
    if not code:
        return None
    if code.endswith("_"):
        return None
    low = " ".join(code.lower().split())
    words = low.split(" ")
    i = 0
    while i < len(words) and words[i] in _MOD_FIRST_WORDS:
        i += 1
    first = words[i].lstrip("#") if i < len(words) else ""
    if first == "elseif":
        return "elseif" if _RE_TAIL_THEN.search(low) else None
    if first in ("else", "case"):
        return first
    return None


def _unclosed_owner_above(lines_above, kind):
    """上方最近一个【还开着】的 kind 类块头在哪？返回它的行首空白；没有 -> None。纯函数。

    lines_above: 当前行【上方】各行（越靠近它的越靠后）
    kind:        'if' / 'select' / 'directive'（_block_kind 的口径）

    括号配对法（与 _closer_owner_above 同一套思路，只是方向相反）：从近到远扫，
    收尾（`End If` / `End Select` / `#End If`）计 +1，同类的块头计 -1；扫到
    "count 已经是 0 的同类块头"，它就是这一档分支的归属者。

    ⚠️ 为什么 `ElseIf` / `Else` 不算块头：它们和 `If` 共用同一个 `End If`，要是也
    计一层，配平永远对不上（`If` / `Else` / `End If` 会配成 -2）。`_block_kind` 对
    它们返回的是 'branch' 而不是 'if'，天然被跳过 —— 这正是这里必须用 `_block_kind`
    而不是"行首是不是 If"来判的原因。
    """
    closer = _DIRECTIVE_CLOSER if kind == "directive" else _BLOCK_CLOSERS.get(kind)
    if not closer:
        return None
    depth = 0
    for raw in reversed(list(lines_above or ())):
        if not (raw or "").strip():
            continue
        if _is_closer_text(raw, closer):
            depth += 1
            continue
        if kind == "directive":
            # 条件编译：只有 `#If … Then` 开一层；`#Else` / `#ElseIf` 是兄弟分支，
            # 别的指令（`#Const` / `Option …` 之类）一律不看。
            if _directive_word(raw) != "if":
                continue
        elif _block_kind(raw) != kind:
            continue
        if depth == 0:
            return _leading_ws(raw)
        depth -= 1
    return None


def branch_align(line_text, lines_above, unit="    "):
    """块内分支行该对齐到哪里（v81；层级按分支分档 v83，纯函数）。

    line_text:   光标所在行的文本（正在写的那一行，光标在行尾）
    lines_above: 它上方各行的文本（越靠近它的越靠后）
    unit:        一级缩进

    返回 `(本行应有的行首空白, 新行应有的行首空白)`；不是分支行、或上方找不到它
    所属的块头（`If x Then y = 1` 这种单选 If、`Else` 敲在 `If` 前面、`Case` 敲在
    `Select Case` 前面…）-> 返回 None，调用方维持原行为（新行 = 本行缩进 + 一级）。
    **绝不猜**：找不到归属者就当没这回事，宁可不对齐也不要挪错。

    层级由 `_BRANCH_OWNER_DEPTH` 定（v83）：
      * `Else` / `ElseIf … Then` -> 与 `If … Then` 齐平，新行深一级；
      * `Case …` / `Case Else` / `Case Is …` -> 比 `Select Case` **深一级**，
        新行（块体）再深一级；
      * `#Else` / `#ElseIf … Then` -> 与 `#If … Then` 齐平，新行深一级。
    缩进单位是 `"    "` 时就是 0 / 4 / 8 格与 4 / 8 / 12 格的区别；用 Tab 或
    2 空格缩进的工程跟着 `unit` 走（`unit * 1`）。
    """
    word = _branch_word(line_text)
    if not word:
        return None
    hashed = _code_part(line_text).lstrip().startswith("#")
    kind = "directive" if hashed else _BRANCH_OWNER.get(word)
    if not kind:
        return None
    owner = _unclosed_owner_above(lines_above, kind)
    if owner is None:
        return None
    mine = owner + unit * _BRANCH_OWNER_DEPTH.get(word, 0)
    return mine, mine + unit


def _indent_unit_for(lines):
    """一级缩进用什么字符串：默认 4 个空格；这个工程里用 Tab 缩进就跟着用 Tab。

    可用 `set VBECOMPLETE_INDENT_UNIT=2`（2 个空格）或 `=tab` 覆盖。
    """
    try:
        raw = (os.environ.get("VBECOMPLETE_INDENT_UNIT") or "").strip()
    except Exception:
        raw = ""
    if raw:
        if raw.lower() in ("tab", "\\t", "\t"):
            return "\t"
        if raw.isdigit():
            return " " * max(1, min(16, int(raw)))
    for t in lines or ():
        if "\t" in _leading_ws(t):
            return "\t"
    return "    "


def enter_indent_wanted(line_text, caret_ec):
    """钩子线程用的纯判据：这一下回车值不值得由我们接管（v79）。

    只有两种情形才接管，其余一律放行给 VBE 原生回车：
      * 整行注释（要忽略它、跟上方第一个非注释行对齐）；
      * 块结构开头（新行要深一级）。
    另有四条保守门槛（拿不准就不抢回车）：
      * 空行 / 只有空白：交给 VBE（它本来就继承上一行的缩进）；
      * 行内有 Tab：列语义算不准，不拿"光标是否在行尾"去赌；
      * 光标不在行尾：那是"拆行"，必须让 VBE 原生处理；
      * caret_ec 拿不到（0）：不接管。
    """
    t = line_text or ""
    if not t.strip():
        return False
    if "\t" in t:
        return False
    try:
        if int(caret_ec or 0) < len(t):
            return False
    except Exception:
        return False
    return bool(_is_comment_only(t) or opens_block(t))


def _unterminated_string(line_text):
    """这一行的【代码部分】是否"引号没闭合"（VBE 原生回车会替用户补上右引号）。

    注释里出现的引号不算（`x = 1  '他说"你好` 是好好的），字符串里的 `""`
    按 VBA 的转义写法跳过去。判定为 True 时调用方【不接管】回车 —— 让 VBE
    自己把右引号补上，那是它原生就有的、很有用的行为。
    """
    t = line_text or ""
    in_str = False
    i, n = 0, len(t)
    while i < n:
        c = t[i]
        if c == "'" and not in_str:
            break                       # 注释开始，后面的引号都不是代码
        if c == '"':
            if in_str and i + 1 < n and t[i + 1] == '"':
                i += 2                  # 字符串里的转义引号 ""
                continue
            in_str = not in_str
        i += 1
    return in_str


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


def implicit_gate_applies(code):
    """这个模块的"隐式变量"要不要被 Option Explicit 闸掉（v75）。

    True  = 该模块写了 `Option Explicit` ⇒ 只收真声明过的名字，
            "用过但没声明"的名字不进候选池（那种写法在该模块里是编译错误）。
    False = 照旧收（没写 Option Explicit，或者开关被关掉了）。

    纯函数（只看文本 + 两个开关，不碰 COM），便于单测 —— 这里最容易出错的
    是开关极性，而不是文本判定本身。
    """
    if not ENABLE_IMPLICIT_IDENTIFIERS or not IMPLICIT_HONOR_OPTION_EXPLICIT:
        return False
    try:
        return vba_parser.has_option_explicit(code)
    except Exception:
        return False


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

# 缓存：命中的窗口句柄列表 + 上次枚举时刻。
# v80b 起这是【周期性重扫】的间隔（不再是"只缓存为空时才扫"）：VBE 的提示窗
# 可能是用到才建的，一旦第一次枚举时它还没出现，旧口径会让它永远进不了缓存。
# 想更省（或排障时想更灵敏）：set VBECOMPLETE_VBE_POPUP_RESCAN=2.0
_vbe_popup_state = {"hwnds": [], "at": 0.0}
try:
    _VBE_POPUP_RESCAN_SEC = float(
        os.environ.get("VBECOMPLETE_VBE_POPUP_RESCAN", "").strip() or 1.0)
except Exception:
    _VBE_POPUP_RESCAN_SEC = 1.0
if not (0.05 < _VBE_POPUP_RESCAN_SEC < 60.0):
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

    枚举全量顶层窗口约 1.75ms —— 每轮轮询都枚举不可取（轮询约 20 次/秒），
    所以缓存句柄、平时只查 IsWindowVisible（快三个数量级）。

    ★v80b：缓存必须能【长出】新句柄。原来只在"缓存为空"时重扫，于是只要第一次
    枚举时某个提示窗还没建出来（VBE 的部分提示窗是【用到才建】的），它就永远进
    不了缓存 —— 表现是"我们探不到那个窗"，于是既不让位的兜底、避让也拿不到矩形，
    候选窗就一直压着它（用户报的"我们的弹窗遮挡 VBE 的形参提示列表"，还"不是每次
    都能复现"—— 取决于工具启动那一刻那个窗建出来了没有）。
    现在按 _VBE_POPUP_RESCAN_SEC（默认 1s）周期性重扫，句柄集合变了就换掉；
    代价是每秒一次 EnumWindows（约 0.2% CPU）。
    """
    cached = _vbe_popup_state["hwnds"]
    hs = [h for h in cached if _hwnd_alive(h)]
    if len(hs) != len(cached):
        _vbe_popup_state["hwnds"] = hs         # 有句柄失效 -> 作废重找
    now = time.time()
    if now - _vbe_popup_state["at"] >= _VBE_POPUP_RESCAN_SEC:
        _vbe_popup_state["at"] = now
        found = [h for h, _cls in _enum_vbe_popup_windows()]
        if set(found) != set(hs):
            # 句柄集合变化只在"提示窗被建出来 / 被销毁"时发生，很少 —— 值得记一行，
            # 它一眼能分清"探不到 VBE 提示窗"到底是【窗没建出来】还是【别的判据】。
            _log("vbe-popup: 提示窗句柄集合 %d -> %d 个 %r"
                 % (len(hs), len(found), [_window_class_name(h)
                                          for h in found]))
            _vbe_popup_state["hwnds"] = found
        hs = found
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


def _proc_owns_line(cm, name, line_no):
    """过程 name 是否真的把第 line_no 行圈在里面（归属校验）。"""
    try:
        return vba_parser.proc_owns_line(cm.Lines(1, line_no), line_no, name)
    except Exception:
        return False


def _proc_of_line(cm, line_no):
    """返回该行所在的过程名；不在任何过程内（模块声明区）则返回 None。

    优先用「代码文本往上找最近的 Sub/Function/Property」（纯 Python，
    行为确定、可单测）；COM 的 ProcOfLine 只作兜底——它在 pywin32 下
    的 byref 参数（ProcKind）常常抛异常而拿不到值。

    ⚠️ v79d：光标【自己这一行】就是过程头时，一律按"没有当前过程"（None）算。
    判据是 `parser.is_proc_header_line`，必须放在最前面 —— 它要同时压住下面
    两条支路（文本扫描会把这一行自己当成过程头；COM 兜底也会这么认，且
    `proc_owns_line` 对过程头那一行恰好返回 True，拦不住）。不这么做的话：
    正在写新过程头 `Sub test`、而模块里已有 `Sub test()` 时，那一行的
    "当前过程"就是 test，test 的局部变量（targetsheet）会被当成当前过程的
    提示出来 —— 用户报的泄漏。过程是"从过程头【之后】开始"的，头那一行是声明。

    ⚠️ v79c 修两处（都在兜底那一支，但都直接影响"作用域"判断）：
      1. pywin32 下 `ProcOfLine` 返回的是【元组】(名字, ProcKind)（ProcKind 是
         byref 出参），旧代码直接 `str(name)` 得到的是 `"('插入工分', 0)"` 这种
         垃圾字符串。它不是 None，于是 ctx["proc_name"] 在【模块级/End Sub 行】
         上变成一个既非空、又不等于任何真实过程名的值 —— 作用域过滤里
         `scope_only` 会被它误判成 True（本该按 v37 口径"模块级声明不受限"），
         过滤阈值也跟着错位。
      2. 更要命的是【不能盲信 VBE】：实测它会把手伸到过程外 —— `End Sub`
         那一行、以及它下面直到下一个过程头之间的所有行，都归给【前一个】
         过程；模块声明区里第一个过程头【之前】的行，归给【第一个】过程。
         照单全收就等于：在模块级 / 新起一个函数的位置上，前一个函数的局部
         变量全部变成"当前过程的"——正是用户报的"别的函数的变量泄漏到本
         位置"。所以拿到名字后必须做一次归属校验（proc_owns_line）：该过程头
         在光标上方、且两者之间没有它的收尾行，才算真的在里面。

    ⚠️ v79e：返回的是【作用域键】，正常情况下就是过程名；同一模块里存在同名
    过程时带 `#序号`（见 parser._proc_scope_key）。那种状态下"过程名"不足以把
    两个过程分开 —— 记录侧也是这么写键的，两边必须一致，否则第二个过程里会
    提示出第一个过程的局部变量（用户报的"过程级的变量都泄露到其他过程"）。
    """
    # 0) v79d：光标这一行自己就是过程头 -> 不在任何 procedure body 里
    try:
        if vba_parser.is_proc_header_line(cm.Lines(line_no, 1)):
            return None
    except Exception:
        pass

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
        except Exception:
            continue
        if isinstance(name, (tuple, list)):
            name = name[0] if name else ""      # (名字, ProcKind)
        name = str(name or "").strip()
        if not name:
            continue
        if _proc_owns_line(cm, name, line_no):  # 归属校验：别信"邻居过程"
            # v79e：同名过程（模块里两个 `Sub test`）的键带 `#序号`，与记录侧
            # 保持一致；文本扫描取不到键就退回裸名（宁可少提示，也不乱提示）。
            try:
                key = vba_parser.proc_at_line(cm.Lines(1, line_no), line_no)
                if key:
                    return str(key)
            except Exception:
                pass
            return name
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


def _norm_code(s):
    """把一行代码压成"只看内容"的形式：去掉所有空白 + 统一小写。

    只用于【在 VBE 重排过的行里找回我们的插入点】：VBE 的自动语法检测会补空格
    （`x=1` -> `x = 1`、`foo(1,2)` -> `foo(1, 2)`）、改大小写（`if a then` ->
    `If a Then`、`rgb` -> `RGB`）。这些都是同一条代码的另一种写法，压平之后两边
    的前缀才可比。
    """
    return "".join(ch for ch in s if not ch.isspace()).lower()


def _locate_inserted(after, written, open_ch, close_ch, col0):
    """在【VBE 改写过的行】里重新找出我们插入的那对符号，返回光标该落在的
    位置（0-based："开符号之后"）；找不到返回 None。

    为什么需要：VBE 有「自动语法检测」，会把整行重新格式化 —— 等号两边补空格
    （`x=1` -> `x = 1`）、关键字改大小写（`if a then` -> `If a Then`）、参数逗号
    后补空格（`foo(1,2)` -> `foo(1, 2)`）、标识符与引号之间补空格
    （`w""` -> `w ""`）。行一变长，写入前算好的光标位置就偏了，表现为「光标
    没落在括号/引号中间」。真机实测（v58）：`    w=4` 里 `w` 之后敲 `"`，得到
    `    w "" = 4`，光标停在 7（第一个引号之前）而不是 8（两引号之间）。

    ★v79 换成本法：**按"压平后的前缀"直接定位**（见 _norm_code）。我们插入的
    位置左边是用户已经打好的代码（written[:col0]），把它压平（去空白 + 小写）
    当指纹，再去 after 里找"左边压平后和这枚指纹一模一样"的那一对符号。好处：
      * VBE 在【任何位置】补空格 / 改大小写都不影响指纹；
      * 一行里有好几对括号时也不会认错 —— 用户报的正是这种行：
        `targetSheet.Cells(i, "b").Interior.Color = RGB()`，`Cells(...)` 是
        第一对、`RGB()` 是第二对，把"第几对"或纯差分对齐一挪就会指到别处。
    差分对齐（difflib，v58 的老办法）留作兜底：指纹也对不上（VBE 真改了内容）
    时再退回它。
    """
    want = open_ch + close_ch
    try:
        pre = _norm_code(written[:col0])
    except Exception:
        pre = None
    # ---- 首选：压平前缀指纹 ----
    if pre is not None:
        try:
            hits = []
            k = after.find(want)
            while k >= 0:
                if _norm_code(after[:k]) == pre:
                    hits.append(k)
                k = after.find(want, k + 1)
            if hits:
                # 理论上只该命中一处（指纹含整行前缀）；真撞上重复时取离原位
                # 最近的 —— VBE 重排只挪几个字符，不会把这一对搬到半行之外。
                near = min(col0, len(after))
                return min(hits, key=lambda x: abs(x - near)) + 1
        except Exception:
            pass
    # ---- 兜底一：差分对齐（v58 的老办法） ----
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
        if p is not None:
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
    # ---- 兜底二：连那一对符号都没了（VBE 把它删/改没了），至少把光标放到
    #      "前缀压平后对齐"的位置上 —— 也就是用户刚打完那段代码之后，绝不乱扔。
    if pre is not None:
        try:
            for k in range(len(after), -1, -1):
                if _norm_code(after[:k]) == pre:
                    return k
        except Exception:
            pass
    return None


def enum_family_prefix(names):
    """一个类型库的枚举成员里，覆盖率达标的最长【小写家族前缀】（长度 >= 2）。

    用来回答"这个库的枚举常量有没有统一的写法"：
      * Excel 库 -> "xl"（97% 的成员都以 xl 打头）—— 用户本来就得敲 xl 才找得到；
      * Office 库 -> "mso"（83%）；
      * stdole  -> ""（Checked / Gray / Color 各说各话，没有共同前缀）。

    没有家族前缀的库【整库丢弃】：那种库的成员都是 Checked / Default / Color
    这类通用词，收进来只会污染用户自己的变量候选（用户口径：宁可少提示，也不
    要噪音）。

    为什么用"覆盖率"而不是"最长公共前缀"：Excel 库里还夹着 rgb* / sigdet* 等
    零散成员，最长公共前缀只有 1 个字符，判不出任何东西。取"能盖住一半以上成员
    的最长前缀"才能稳稳落在 xl / mso 上。

    前缀越长覆盖率越低（子集关系），所以一旦某长度不达标就可以直接停。
    """
    lows = [str(n).lower() for n in names if n]
    if not lows:
        return ""
    total = len(lows)
    best = ""
    for k in range(2, 25):
        counts = {}
        for x in lows:
            if len(x) >= k:
                p = x[:k]
                counts[p] = counts.get(p, 0) + 1
        if not counts:
            break
        p, n = max(counts.items(), key=lambda kv: kv[1])
        if n < total * _HOST_ENUM_FAMILY_PCT:
            break
        best = p
    return best


def host_enum_family(names):
    """(家族前缀, 该前缀下的成员名列表)。没有家族前缀则 ("", [])。"""
    p = enum_family_prefix(names)
    if not p:
        return "", []
    if _HOST_ENUM_FAMILIES and p not in _HOST_ENUM_FAMILIES:
        return "", []      # 白名单模式：不在清单里的家族整库跳过
    return p, [str(n) for n in names
               if str(n).lower().startswith(p)]


def _tlb_member_name(ti, memid):
    """ITypeInfo.GetNames 的兼容封装（各版本 pywin32 的参数个数不同）。"""
    try:
        return ti.GetNames(memid)
    except TypeError:
        return ti.GetNames(memid, 256)


def _tlb_enum_members(guid, major, minor):
    """按 GUID + 版本加载已注册的类型库，返回其【常量枚举】的全部成员名。

    只读、不碰 VBE 内容；任何一步失败都返回 []（绝不抛 —— 这条路上出异常会
    让整个候选池收集失败，得不偿失）。

    注意 VBA 自己的库加载不了：它没有注册成类型库（在 VBE7.DLL 里），
    LoadRegTypeLib 会报"库没有注册"。VBA 自带的名字本来就有 vba_builtins.py
    那份静态清单兜着，不依赖这里。
    """
    try:
        import pythoncom
        import pywintypes
    except Exception:
        return []
    try:
        tlb = pythoncom.LoadRegTypeLib(
            pywintypes.IID(str(guid)), int(major), int(minor), 0)
    except Exception:
        return []
    out = []
    try:
        count = tlb.GetTypeInfoCount()
    except Exception:
        return []
    for i in range(count):
        try:
            ti = tlb.GetTypeInfo(i)
            attr = ti.GetTypeAttr()
            if int(attr.typekind) != int(pythoncom.TKIND_ENUM):
                continue
            for j in range(attr.cVars):
                try:
                    vd = ti.GetVarDesc(j)
                    nm = _tlb_member_name(ti, vd.memid)[0]
                except Exception:
                    continue
                if nm and _RE_PLAIN_IDENT.match(nm):
                    out.append(nm)
        except Exception:
            continue
    return list(dict.fromkeys(out))


def _host_enum_refs():
    """当前 VBE 里各工程引用的类型库 -> [(名称, GUID, 主版本, 次版本), ...]。

    读不到（Excel 没开 / COM 失败）返回 None —— 与"读完发现没有引用"区分开，
    调用方据此决定要不要缓存结果。
    """
    try:
        vbe = _get_vbe_cached()
        if vbe is None:
            return None
        projects = []
        try:
            active = vbe.ActiveVBProject
        except Exception:
            active = None
        if active is not None:
            projects.append(active)
        else:
            try:
                col = vbe.VBProjects
                for i in range(1, col.Count + 1):
                    try:
                        projects.append(col.Item(i))
                    except Exception:
                        continue
            except Exception:
                projects = []
        refs = []
        seen = set()
        for pj in projects:
            try:
                col = pj.References
                n = col.Count
            except Exception:
                continue
            for i in range(1, n + 1):
                try:
                    r = col.Item(i)
                    key = (str(r.GUID).upper(), int(r.Major), int(r.Minor))
                    if key in seen:
                        continue
                    seen.add(key)
                    refs.append((str(getattr(r, "Name", "") or ""),) + key)
                except Exception:
                    continue
        return refs
    except Exception:
        return None


# 宿主枚举常量的进程级缓存：加载三个类型库约 15ms，而引用清单几乎不变，
# 所以按"引用签名"缓存 —— 用户新增/删除引用后签名一变就重新枚举一次。
_host_enum_cache = {"key": None, "items": []}


def host_enum_constants():
    """工程引用的类型库里那些枚举常量 -> [(规范名, 最短输入长度), ...]。

    最短输入长度 = 该库的家族前缀长度（xl / mso 各 2、3）。引擎据此要求
    "输入至少这么长、且候选从头开始以它开头"才提示这类名字 ——
    于是 xl -> xl*、xlu -> xlUp，而 ms 只出 MsgBox、绝不带出 mso*。

    顺序稳定、已去重；取不到一律返回 []（引擎行为与旧版完全一致）。
    """
    if not ENABLE_HOST_ENUMS:
        return []
    if os.environ.get("VBECOMPLETE_NO_HOST_ENUMS", "0").strip() == "1":
        return []
    try:
        refs = _host_enum_refs()
    except Exception:
        refs = None
    if not refs:
        return []
    key = tuple(sorted(refs))
    if _host_enum_cache["key"] == key:
        return _host_enum_cache["items"]
    items = []
    fams = []
    for name, guid, major, minor in refs:
        members = _tlb_enum_members(guid, major, minor)
        if not members:
            continue
        fam, kept = host_enum_family(members)
        if not fam:
            _log("host-enums: %-12s 无家族前缀 -> 整库跳过（%d 个枚举成员）"
                 % (name, len(members)))
            continue
        fams.append("%s->%s(%d)" % (name, fam, len(kept)))
        for m in kept:
            items.append((m, len(fam)))
    # 去重：同名的以先出现的为准（Excel 里也有 xl* 同名项）
    uniq = {}
    for m, ml in items:
        uniq.setdefault(m.lower(), (m, ml))
    out = list(uniq.values())
    _host_enum_cache["key"] = key
    _host_enum_cache["items"] = out
    _log("host-enums: %d 条（%s）" % (len(out), " ".join(fams)))
    return out


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
        # 宿主类型库的枚举常量（v67）：{小写名: 最短输入长度}。
        # 只收"有家族前缀"的库（Excel -> xl*、Office -> mso*），引擎对这批名字
        # 只认"从头开始的连续前缀"命中，且输入长度不低于家族前缀长度 ——
        # 它们数量太大（实测 4538 条），放开跳步匹配会淹没正常候选。
        self._host_enum_names = {}

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
        self._host_enum_names = {}
        # 丢掉缓存的 VBE 对象 —— 这是 Excel 能否真正退出的关键一步。
        # 刻意【不】调 CoFreeUnusedLibraries —— 它释放不了我们持有的引用，
        # 在 Excel 关闭期间调用反而会拖慢/惊扰 COM（详见 _co_free 的说明）。
        # 这里用 collect=True 主动 gc，确保 COM 包装对象立刻回收。
        _release_vbe_proxy(collect=True)

    def snapshot(self):
        """轻量读取「模块名 + 当前行号 + 行文本 + 光标列」，用于轮询检测内容变化。

        刻意不做列语义探测（那会临时移动光标），因此足够廉价，可高频调用。
        返回 (line_no, line_text, module_name, caret_col)；不在代码窗/多行选择/
        取不到时返回 None。

        v64：多带一个【模块名】。主线程靠它区分"用户真在敲字"和"换了个代码
        窗格/模块"—— 后者会让"同一行 + 文本不同"看起来像一次编辑（Ctrl+Tab
        切代码窗、点工程树换模块都会），其实一个字都没敲，不该弹候选。
        模块名取不到就带空串，调用方据此不做否决（绝不让工具整个哑掉）。

        v79：再带一个【光标列 ec】。只给"回车自动缩进"那道判据用（见
        enter_indent_wanted）—— 它必须在【键盘钩子线程】里就决定要不要接管
        这一下回车，而钩子线程不许碰 COM，只能读轮询留下的这份快照。
        ⚠️ 变更检测仍然只比前三个元素（行号/行文本/模块名）：光标的移动不算
        "内容变化"，否则点来点去都会弹候选窗（见 main._poll_mod_switch 附近）。
        """
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return None
            cp = vbe.ActiveCodePane
            if cp is None:
                return None
            sl, _sc, el, ec = cp.GetSelection()
            if sl != el:      # 多行选择：不参与补全
                return None
            cm = cp.CodeModule
            try:
                mod = str(cm.Name or "")
            except Exception:
                mod = ""
            try:
                ec = int(ec)
            except Exception:
                ec = 0
            snap = (sl, cm.Lines(sl, 1), mod, ec)
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
            # blank_decl=True：这里**必须**把光标处的词抹掉 —— 哪怕它是这一行正在
            # 声明的名字（v79e）。本函数的语义是"抹掉光标处的词之后，模块里还剩
            # 哪些名字"，那个词自己不能当证据；而 _blank_ident_at 默认会放过
            # "正在声明的名字"（为了不改动文本、不破坏行结构），所以在这一路要
            # 显式要求抹掉。不加这个参数会怎样：光标停在 `Sub zzq` 上（zzq 别处
            # 都没有）时，zzq 把自己算成"别处出现过"的铁证，回声防护放行，
            # 弹窗把用户正在敲的字原样提示回来 —— 第 21.4 节钉的正是这条。
            src = vba_parser._blank_ident_at(code, line_no, col, True)
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
        # 宿主类型库枚举常量 -> 最短输入长度（v67），见下面的收集段。
        host_min = {}
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
                    #
                    # v75：写了 Option Explicit 的模块整段跳过（详见
                    # IMPLICIT_HONOR_OPTION_EXPLICIT 的说明）—— "用过但没声明"
                    # 在那里是编译错误，收进来只会提示出坏代码。
                    imp = []
                    if ENABLE_IMPLICIT_IDENTIFIERS:
                        if implicit_gate_applies(code):
                            _log("  mod=%-16s Option Explicit -> 不收集隐式变量"
                                 % mod_name)
                        else:
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
        #
        # 【v70】这一批里 **枚举常量（vb*）默认不收** —— 用户口径："我不怎么使用
        # VBA 内置枚举，提示出来对我有干扰；函数那些要保留。" 于是函数（MsgBox /
        # Left / Split …）与内建数据类型照常收，只有 vb* 那一组由
        # ENABLE_VBA_CONSTANTS 裁掉（默认 False；`set VBECOMPLETE_VBA_CONSTANTS=1`
        # 开回来）。注意裁剪放在【收集侧】而不是改清单本身：清单保持完整，
        # 测试与诊断仍能看到全貌，开关一开立刻生效。
        #
        # 【v77】连**内建函数与内建数据类型**也默认不收了（ENABLE_VBA_BUILTINS
        # 默认 False）—— 用户口径："vba 本身带的关键字、函数这块都删掉；保留变量名、
        # 自定义函数/过程名、窗体/控件名、模块名。" 于是默认口径下候选池里
        # 只剩"这个工程里真实存在的东西"。
        #
        # ⚠️ 连带的**唯一**后果（其余一个字没动）：`Dim x As <这里>` 的候选里不再
        # 出现 Long / String / Integer 这类内建类型名 —— 因为 `As` 位置的候选是
        # "先出候选、再按 type_names 过滤"，不在池子里就永远显示不出来。
        # 那一位置现在只剩你自己的类型：窗体 / 类模块 / 模块名 + Type / Enum 名。
        # 想要回内建类型（含 MsgBox 那批函数）：`set VBECOMPLETE_VBA_BUILTINS=1`。
        if ENABLE_VBA_BUILTINS:
            _bn_src = vba_builtins.BUILTIN_FUNCTIONS
            if ENABLE_VBA_CONSTANTS:
                _bn_src = _bn_src + vba_builtins.BUILTIN_CONSTANTS
            for _bn in _bn_src:
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
        #
        # 【v77】**默认关闭**（用户口径："vba 本身带的关键字也删掉"）—— 于是默认
        # 口径下打 `su` 不会冒出 `Sub`、打 `if` 不会冒出 `If`，候选池里只剩你这个
        # 工程里真实存在的名字。想开回来：`set VBECOMPLETE_VBA_KEYWORDS=1`。
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

        # ---- 宿主类型库的枚举常量（v67）----
        # 用户报："输入 vb 会提示一堆 VBA 枚举值，输入 xl 却一个 xl 开头的枚举值
        # 都不提示。" vb* 是 v61 收的 VBA 内建常量；xl* 属于【宿主 Excel 类型库】，
        # 一直没收 —— 于是同一个动作用在 xl 上什么都出不来。
        #
        # 与内建名字同源：由工程【引用的类型库】决定存在，代码文本里从没写过
        # xlUp 也照样该补得出来。所以四路承接完全一样（池子 / declared_names /
        # structural_names / _decl_by_mod）—— 少任何一路，回声防护都会把
        # "代码里从未出现过的 xlUp" 当幽灵剔掉（v60 窗体名、v61 内建函数、
        # v62 关键字都栽在这一件事上，这是第四次）。
        #
        # 唯一的差别在【匹配口径】：后端把 {小写名: 最短输入长度} 交给引擎，
        # 引擎只认"从头开始的连续前缀"且输入长度不低于该值。理由见 engine.trigger。
        if ENABLE_HOST_ENUMS:
            for _hname, _hmin in host_enum_constants():
                _hl2 = _hname.lower()
                if _hl2 in declared_names:
                    continue
                records.append((_hname, _HOST_ENUM_MODULE, None, False))
                declared_names.add(_hl2)
                structural_names.add(_hl2)
                _decl_by_mod.setdefault(_hl2, set()).add(_HOST_ENUM_MODULE)
                host_min[_hl2] = _hmin

        self._type_names = type_names
        self._declared_names = declared_names
        self._declared_by_module = _decl_by_mod
        self._structural_names = structural_names
        self._builtin_names = builtin_names
        self._host_enum_names = host_min
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
        ⚠️ v77：第三类默认是空的（那几个开关都默认关）⇒ 默认口径下只有
        **组件名 + 窗体控件名**。这一路证据本身一个字没删，开关打开就回来。

        与 get_declared_names 的区别：后者是"代码文本里真声明过什么"，前者是
        "本来就有、与代码文本无关"。用户刚拖上去的 Label1、一行代码都没有的新
        窗体 UserForm1、代码里从没调用过的 MsgBox、新模块第一行就想补的 If，
        都属于前者而不属于后者。

        引擎用它来判断"回声候选是不是真名字"——这类名字不能走"只在本模块声明
        过、现场文本又扫不到 -> 当幽灵剔掉"的规则（v60 组件名与控件名、v61 VBA
        内建名字：打 ms 要能补出从没写过的 MsgBox；v62 关键字：打 if 要能补出
        新模块里还没写过的 If）。

        ⚠️ 直接交回内部集合（**只读，调用方不得修改**）：v67 收进宿主枚举常量后
        这个集合有 4600 多条，而引擎的回声防护会对每个候选查一次 —— 每次调用都
        复制一份的话，输入 xl（一次带出 2000 多个候选）光复制就要 2 秒。
        """
        return getattr(self, "_structural_names", None) or set()

    def get_builtin_names(self):
        """返回真正作为"语言自带词汇"收录的那些（小写集合，v61 / v62）。

        含：VBA 内建函数 / 内建常量 / 内建数据类型（v61）+ 语言关键字（v62）。

        与 vba_builtins 的清单不同：工程里已经自己声明过同名时以用户的定义为准，
        那些名字不会出现在这里 —— 引擎据此对"语言自带的公共词汇"收紧模糊匹配
        （纯分散命中要求输入至少 4 个字符），而不会连用户自己的同名定义一起收紧。

        后端不提供该接口时引擎不做任何收紧（旧式 / 测试后端行为不变）。
        """
        return getattr(self, "_builtin_names", None) or set()

    def get_host_enum_names(self):
        """宿主类型库的枚举常量 -> {小写名: 最短输入长度}（v67）。

        内容：工程【引用的类型库】里的枚举常量，且只收有家族前缀的库
        （Excel 的 xl*、Office 的 mso*）。值是"该家族前缀的长度"，引擎据此
        要求输入至少这么长、且候选从头开始以它开头才提示 ——
        于是 xl -> xl*、xlu -> xlUp，而 ms 只出 MsgBox、绝不带出 mso*。

        为什么必须单独给一份而不是并进 get_builtin_names：内建那批（400 来个）
        是"收紧一档跳步匹配"，这批（4500 多条）得直接【只认前缀】—— 同一条
        规则套两批名字，轻则噪音爆炸（实测输入 ms 会带出 2365 条 mso*），
        重则把内建名字该有的跳步补全也一起砍掉。

        工程里已经自己声明过同名的以用户的定义为准，那些名字不会出现在这里。

        后端不提供该接口时引擎不做任何收紧（旧式 / 测试后端行为不变）。
        约定：键已经是小写；直接交回内部字典（**只读，调用方不得修改**）——
        它有 4500 多条，每次调用复制一份纯属浪费。
        """
        if not getattr(self, "_host_enum_names", None):
            # 空就现收一次（顺带把标识符池建起来）—— __init__ 里它初始化为 {}，
            # 光判 None 会让"还没收集过"被当成"没有宿主常量"。
            try:
                self.get_identifiers()
            except Exception:
                return {}
        return getattr(self, "_host_enum_names", None) or {}

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
                    # v79 起把这一路记进日志：VBE 到底怎么重排的、我们又把它
                    # 认到了哪儿 —— 以后再出"光标没落在括号里"，一眼就能看出是
                    # 定位错了还是 VBE 根本没重排（仅 run_debug.bat 下有开销）。
                    _log("autopair: VBE 重排了整行 -> 重新定位"
                         " (写入=%r 实际=%r col0=%d 定位=%r)"
                         % (new_line, actual, col0, off))
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
    # 往上读多少行来"跳过注释行"（够用即可：注释再长也不会超过这个数）。
    _INDENT_LOOKBACK = 200

    def _lines_above(self, cm, line_no, limit=None):
        """取第 line_no 行【上方】的各行文本（最近的在最后）。

        只读、且只在"回车由我们接管"这一条少见路径上调用，行数封顶
        _INDENT_LOOKBACK —— 免得在几千行的模块里把半个模块都读进来。
        """
        n = self._INDENT_LOOKBACK if limit is None else int(limit)
        start = max(1, int(line_no) - n)
        cnt = int(line_no) - start
        if cnt <= 0:
            return []
        try:
            text = cm.Lines(start, cnt)
        except Exception:
            return []
        lines = str(text).split("\n")
        if lines and lines[-1] == "":
            lines.pop()                    # Lines() 末尾带回车，会多一个空元素
        return [l.rstrip("\r") for l in lines]

    def _lines_below(self, cm, line_no, limit=None):
        """取第 line_no 行【下方】的各行文本（最近的在最前）。

        v79b 补收尾时用：得知道下面是不是已经挂着 `End If` / `Next` 了，
        免得替用户补出一个重复的收尾。同样只读、行数封顶 _INDENT_LOOKBACK。
        """
        n = self._INDENT_LOOKBACK if limit is None else int(limit)
        try:
            total = int(cm.CountOfLines)
        except Exception:
            return []
        start = int(line_no) + 1
        cnt = min(n, total - start + 1)
        if cnt <= 0:
            return []
        try:
            text = cm.Lines(start, cnt)
        except Exception:
            return []
        lines = str(text).split("\n")
        if lines and lines[-1] == "":
            lines.pop()                    # Lines() 末尾带回车，会多一个空元素
        return [l.rstrip("\r") for l in lines]

    def new_line_below(self, smart=False, auto_close=False, align_branch=True):
        """在光标所在行的【下方】新起一行，光标落在缩进之后。

        等价于"先把光标移到本行末尾，再按回车"，但一步到位 —— 而且当前行
        【不拆分】（光标右侧的代码留在原行）。VBE 只在"光标已在行尾"时按回车
        才继承上一行缩进；光标停在行中间时按回车会把行拆开，所以必须由我们代劳。

        smart=False（Shift+Enter，v78 起的老口径）：不依赖光标列、也不移动
        光标，直接读本行文本、取它的行首空白作为新行缩进。

        smart=True（回车自动缩进，v79）—— 由用户那两条要求而来：
          * 本行是【整行注释】：忽略它（以及上方连续的注释行），跟上方第一个
            非注释行对齐；
          * 本行是【块结构开头】（For / Do / If…Then / With / Sub / Function /
            Type / Enum / Select Case …）：新行缩进一级。
        缩进由纯函数 next_line_indent 算；这一模式下多了四道保险，任一不满足
        就返回 False 让调用方把回车【原样还给系统】：
          * 有选区：多行选择按回车是删除/覆盖，语义完全不同；
          * 光标不在本行行尾：那是"拆行"，必须交给 VBE 原生回车；
          * 本行是空行（只有空白）：让 VBE 自己继承上一行缩进，别跟用户手动
            排好的缩进较劲；
          * 本行引号没闭合：VBE 原生回车会替用户补上右引号，别抢这个活。

        auto_close=True（回车自动补收尾，v79b；嵌套判定 v79c）—— 仅对 smart 模式生效：
        块头是 `If x Then` / `With rng` / `Sub Foo()` / `Type Foo` / `Enum E` 这类
        【真需要收尾】的，再在其后补一行同级的 `End If` / `End With` / `Next` /
        `End Type`…，光标停在中间那行（缩进正好），用户直接在块里写内容。要不要补
        由纯函数 block_closer + closer_needed 决定：下面已经挂着本块的收尾、或者
        这块的下文已经在写了 —— 都不补，免得顶开已有代码。

        align_branch=True（分支行自动对齐，v81；层级分档 v83）—— 仅对 smart 模式生效：
        本行是块【内】的分支（`Else` / `ElseIf … Then` / `Case …`，含条件编译的
        `#Else` / `#ElseIf`）时，先把它【自己】拉回该在的层级（`Else` / `ElseIf` 与
        `If … Then` 齐平、`#Else` / `#ElseIf` 与 `#If … Then` 齐平、**`Case` 比
        `Select Case` 深一级**），新行再深一级。
        要不要对齐、对齐到哪，由纯函数 branch_align 决定：上方找不到那个"还开着的块头"
        就什么都不做（宁可不对齐，也不要挪错）。**只改行首空白**，行内一个字符都不碰。

        ⚠️ v79c：内层块下面压着的往往是【外层】的收尾（双层 `For` 里的 `Next`），
        它的缩进正好比内层块头浅一格 —— 单看缩进会误判成"已经有收尾了"。所以
        closer_needed 还会查"上方有没有一个同类、未闭合、缩进正好对上它的块头"，
        有就说明那个收尾属于外层，本块照样得补。

        成功返回 True；不在代码窗 / 取不到 COM / 出任何异常都返回 False ——
        调用方据此把这一下按键原样还给系统（回车绝不能吞掉）。
        """
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return False
            cp = vbe.ActiveCodePane
            if cp is None:
                return False
            cm = cp.CodeModule
            sl, sc, el, ec = cp.GetSelection()
            line_text = None
            closer_text = None             # 要补的块收尾（smart + auto_close 才可能有）
            # ⚠️ v81：这个必须在 if smart 之外初始化 —— 放在 smart 分支里的话，
            # smart=False（Shift+Enter）走到下面那句 `if fixed_line is not None`
            # 会抛 NameError，被 `except -> return False` 吞成"Shift+Enter 不换行"
            # （第 29.4~29.7 / 56.8e / 57.4h 当场全红，和 v79b 的 closer_text 同坑）。
            fixed_line = None              # 分支行要拉回它所属块头的缩进
            if smart:
                if int(sl) != int(el) or int(sc) != int(ec):
                    return False           # 有选区：不接管
                anchor = int(sl)
                if anchor <= 0:
                    return False
                line_text = cm.Lines(anchor, 1)
                if not str(line_text).strip():
                    return False           # 空行：交给 VBE
                if _unterminated_string(line_text):
                    return False           # 让 VBE 去补右引号
                col, _sem, _tw, _w2 = _caret_char_col(
                    cm, cp, anchor, line_text, ec)
                if line_text[col - 1:].strip():
                    return False           # 光标后面还有内容 -> 是"拆行"
                above = self._lines_above(cm, anchor)
                unit = _indent_unit_for(above + [line_text])
                indent = next_line_indent(line_text, above, unit)
                if align_branch:
                    _al = branch_align(line_text, above, unit)
                    if _al is not None:
                        fixed_line, indent = _al
                if auto_close:
                    # 块头才谈得上收尾；"要不要真的补"交给纯函数判（下面可能
                    # 已经挂着同一个收尾、或者这块的下文已经在写了）。
                    # v79c：把块头类型与上方各行一并交给它 —— 判"更浅的同类收尾
                    # 是不是外层块的"（双层 For 的 `Next` 就长这样）要用。
                    _close = block_closer(line_text)
                    if _close:
                        base = _leading_ws(line_text)
                        if closer_needed(self._lines_below(cm, anchor), _close,
                                         _indent_width(base),
                                         kind=_block_kind(line_text),
                                         lines_above=above):
                            closer_text = base + _close
            else:
                # 有选区时以选区【末行】为基准（正常情况下 sl == el）。
                anchor = max(int(sl), int(el))
                if anchor <= 0:
                    return False
                indent = _leading_ws(cm.Lines(anchor, 1))
            if fixed_line is not None:
                # v81：分支行（Else / ElseIf…Then / Case…）先拉回它该在的层级
                # （v83：Case 比 Select Case 深一级，其余与块头齐平）。
                # 只动【行首空白】，行内一个字符都不碰；写不成就不动、也不插新行，
                # 让调用方把这一下回车原样还给系统。
                _raw = str(line_text).rstrip("\r\n")
                _ws = _leading_ws(_raw)
                if _ws != fixed_line and not self._write_line(
                        cm, anchor, fixed_line + _raw[len(_ws):]):
                    return False
            # InsertLines 在 anchor+1 处插入（anchor == CountOfLines 时即追加到末尾），
            # 已有行自动下移 —— 正是"在本行下方新起一行"。
            cm.InsertLines(anchor + 1, indent)
            if closer_text:
                # 收尾插在"新起的空行"下面：块头下面留出缩进好了的一行给用户写
                # 内容，收尾再往下、缩进回到块头同级。原有行整体再下移一行。
                cm.InsertLines(anchor + 2, closer_text)
            col = _indent_end_col(indent)
            cp.SetSelection(anchor + 1, col, anchor + 1, col)
            return True
        except Exception:
            return False
        finally:
            # 与 apply_completion 同理：写过代码（VBAProject 变脏），立刻放开
            # 手里的 COM 代理，别拖住 Excel 的退出/写回。
            _release_vbe_proxy()

    # ---- 按行移动光标（Shift+↑ / Shift+↓） ----
    def move_caret_line(self, delta):
        """把代码窗光标上（delta<0）/ 下（delta>0）移 delta 行，尽量保持同一列。

        【为什么不用"放行按键让 VBE 自己动光标"】VBE 里 Shift+↑ / Shift+↓ 的
        原生语义是**扩展选区**——按住 Shift 按一下会把上一行整行选进去，光标
        并没"干净地"上移一行。v78 用户口径是"跳出候选列表，并把光标上移 /
        下移一行"，所以这里由我们用 COM 直接落光标，那一对按键由调用方吞掉。

        列刻意沿用 GetSelection 给的列值原样透传：VBE 的列与 SetSelection 的列
        是同一套语义（显示列，Tab 展开），因此同一列值就是"同一视觉列"；目标
        行较短时由 VBE 自己钳到行尾（不在这里做字符数换算，避免 Tab 行算错）。

        成功返回 True；取不到 COM / 光标行无效 / 目标行越界 / 出任何异常都返回
        False —— 调用方据此把这一下按键原样还给系统，绝不吞掉用户的按键。
        """
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return False
            cp = vbe.ActiveCodePane
            if cp is None:
                return False
            cm = cp.CodeModule
            sl, sc, el, ec = cp.GetSelection()
            # 光标在选区的【活动端】：正向选择在 (el, ec)，反向选择在 (sl, sc)。
            if int(el) >= int(sl):
                line, col = int(el), int(ec)
            else:
                line, col = int(sl), int(sc)
            target = line + int(delta)
            if line <= 0 or target <= 0 or target > int(cm.CountOfLines):
                return False
            cp.SetSelection(target, col, target, col)
            return True
        except Exception:
            return False
        finally:
            # 与 new_line_below 同理：立刻放开手里的 COM 代理，别拖住 Excel 退出。
            _release_vbe_proxy()
