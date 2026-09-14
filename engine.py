"""
补全引擎（纯逻辑，不依赖 Excel / 界面，可单测）。

职责：
  - 根据"当前光标所在行 + 列"算出光标前的单词；
  - 用该单词前缀过滤全部标识符，得到候选列表；
  - 管理弹窗显示、上下选择、接受补全、取消；
  - 接受时调用 backend.apply_completion 完成实际插入（由具体后端实现）。

backend 需实现：
  get_context() -> dict | None
      {line_no, caret_col, line_text, in_string, in_comment,
       in_type_position, in_decl_position, proc_name, module_name}
      不在代码窗时返回 None
  get_identifiers() -> list[记录]
      记录形态（_normalize_scoped 全兼容）：
        - "x"                   纯名字（视为全局可见，便于单测）
        - ("x", scope)          旧式二元组（scope=过程名或 None）
        - ("x", mod, proc, priv) 完整记录
  apply_completion(line_no, word_start_col, end_col, completion)
      end_col = 补全词范围的右端（1-based，开区间）。通常就是光标列；
      光标停在标识符【开头】或【中间】时大于光标列——前者要盖住光标右边
      那个残留标识符，后者要盖住同一个名字的后半截。取词见 extract_word_at。
"""

import os
import re
import time

from log import log as _log
from log import LOG_ENABLED as _LOG_ENABLED

# VBE 自带提示列表的让位开关（v55）。
#
# 光标停在 VBE 自己会弹「自动列出成员」的位置时我们让位（详见
# vbe_list_expected）。若某台机器上 VBE 的自动列出成员是关掉的、或你想
# 永远用我们的列表，设 VBECOMPLETE_NO_YIELD=1 关掉这套让位。
try:
    YIELD_TO_VBE_LIST = (
        os.environ.get("VBECOMPLETE_NO_YIELD", "0").strip() != "1")
except Exception:
    YIELD_TO_VBE_LIST = True

# 标识符（含中文）：首字符为字母/下划线，其后可跟字母/数字/下划线。
# [^\W\d] = 非“非单词字符”且非数字 = 字母或下划线（Unicode 感知，含中文）。
_IDENT = re.compile(r"[^\W\d]\w*$")

# 确认补全后的"静默期"（秒）：这段时间内忽略新的触发请求。
# 目的：Tab 确认后，可能还有排队中的 trigger（来自刚才那次按键），
# 若不屏蔽，弹窗会在收起后立刻又冒出来。
SUPPRESS_AFTER_ACCEPT = 0.35


def extract_word_before(line_text, caret_col):
    """返回 (word, start_col)，均为 1-based；无则返回 (None, None)。

    只看光标【左边】。从名字中间/结尾删字符都在这个模型内；从名字【开头】
    删则左邻是空格或括号，取不到词——那种情形由 extract_word_at 兜住。
    """
    prefix = line_text[:caret_col - 1] if caret_col and caret_col > 1 else ""
    m = _IDENT.search(prefix)
    if not m:
        return None, None
    return m.group(0), m.start() + 1


# 正向标识符（从某个下标起，首字符必须是字母/下划线/中文）。
_IDENT_FWD = re.compile(r"[^\W\d]\w*")


def extract_word_at(line_text, caret_col):
    """返回 (word, start_col, end_col)，均为 1-based；无则返回 (None, None, None)。

    当前词 = **光标所在处那一个完整标识符**，而不是"光标左边那半截"。

    v44：不再只看"光标前的词"。用户把光标停在名字最左边、按 Delete 从头删字符
    时，光标左边是空格（或 `(`,`=` 之类），只看左边永远取不到词，弹窗因此永远
    不出现——用户报的正是这个。那种情形改为取光标【右边】紧挨的标识符，
    替换范围 [caret_col, end_col) 覆盖整个标识符，这样 Tab 确认时是把残留的
    名字【整段】换掉，而不是把候选插在光标处、拼出 `abc` + `bcDef` 这种双份名字。

    v46：光标左边有词时，**还要往右吞掉紧邻的标识符字符**。用户把光标停在名字
    【中间】删字符时（test23456789 删掉中间那个 t -> tes|23456789），若只取左边
    那半截 `tes`，候选就会按 `tes` 的前缀算，把 `test` / `test02` 提示出来，
    完全无视光标右边还留着的 `23456789`——用户报的正是这个。当前词取【整段名字】
    （左半截 + 右半截）后，`tes23456789` 既不匹配 `test` 也不匹配 `test02`，
    只剩真正那一整段名字（Tab 可一次把整段换掉）。

    规则三条：

      1. 光标左边有词 -> 取它，再往右吞掉紧邻的标识符字符，得到整段名字，
         替换范围 [start_col, end_col)；
      2. 光标左边没有词，但光标【右边】紧挨着一个标识符（光标正好停在该标识符
         开头）-> 取这个标识符，替换范围 [caret_col, end_col)；
      3. 两边都没有（光标夹在两个非标识符之间）才返回 None，由调用方收起弹窗。
    """
    if not line_text or not caret_col or caret_col < 1:
        return None, None, None
    word, start = extract_word_before(line_text, caret_col)
    i = caret_col - 1                     # 光标右边第一格的 0-based 下标
    n = len(line_text)
    if word:
        # v46：把光标右边紧邻的标识符字符一并吞进来（同一个名字的后半截）。
        j = i
        while j < n and (line_text[j].isalnum() or line_text[j] == "_"):
            j += 1
        if j > i:
            return line_text[start - 1:j], start, j + 1
        return word, start, caret_col
    # 兜底：光标停在标识符开头（左边是分隔符，右边是标识符字符）
    if i >= n:
        return None, None, None
    m = _IDENT_FWD.match(line_text, i)
    if not m:
        return None, None, None
    return m.group(0), m.start() + 1, m.end() + 1


def _ident_char_before_caret(ctx):
    """光标是否紧贴着标识符字符（左边一格，或右边一格）。

    用于"编辑器内容变化"触发的路径：若刚输入的是空格/标点/换行，
    说明用户并没有在拼标识符，应当收起弹窗（保持与按空格收起一致的行为）。

    v44：判据从"光标【前】一个字符"放宽到"光标前【或】后一个字符"。因为
    光标停在名字开头时（按 Delete 从头删字符），左邻恰恰是空格/括号，只认
    左边就会把"正在拼标识符"误判成"已经打完"，弹窗永远弹不出来。右边紧挨
    标识符字符，同样说明光标正贴在某个标识符上。

    这不会破坏"打空格收起"：在词尾敲空格后，光标右边通常是行尾或括号，
    右边没有标识符字符，依旧收起；具体该弹该收仍由 extract_word_at 定夺。
    """
    caret = ctx.get("caret_col") or 0
    if caret <= 1:
        return True   # 行首：交给后续取词逻辑（取不到词自然会收起）
    text = ctx.get("line_text") or ""
    i = caret - 2
    if i < 0 or i >= len(text):
        return True
    ch = text[i]
    if ch.isalnum() or ch == "_":
        return True
    j = caret - 1                     # 光标右边那一格
    if j < len(text):
        ch = text[j]
        return bool(ch.isalnum() or ch == "_")
    return False


def line_edit_inserted(old_line, new_line):
    """对比同一行的前后两版文本，返回本次【纯插入】的内容。

    只认一种形态：新文本 = 旧文本在某个位置塞进一段新字符、其余原样（手工
    输入正是这种）。删除、替换、多处改动一律返回 None —— 返回 None 不代表
    没变化，只是"这一步改动无法用单一插入来解释"。

    刻意只看文本、不看光标位置：后端快照为了廉价，只给 (行号, 行文本)，
    没有列号。靠最长公共前缀 + 最长公共后缀定位改动段同样能得出结论，
    且不引入额外的 COM 开销。
    """
    if old_line is None or new_line is None:
        return None
    if len(new_line) <= len(old_line):
        return None                      # 没变长 -> 是删除或替换
    n = len(old_line)
    p = 0                                # 最长公共前缀
    while p < n and old_line[p] == new_line[p]:
        p += 1
    s = 0                                # 最长公共后缀（不许越过前缀，避免重叠）
    limit = n - p
    while s < limit and old_line[n - 1 - s] == new_line[len(new_line) - 1 - s]:
        s += 1
    if p + s != n:
        return None                      # 中段还有别的差异 -> 不是纯插入
    return new_line[p:len(new_line) - s]


def typed_separator(old_line, new_line):
    """这次编辑是不是"往行里敲了一段不含标识符字符的内容"（空格 / 标点）。

    v52 要治的现象：把光标停在某个标识符【最前面】（形参首字符前就是这种），
    敲一下空格 -> 弹窗把右边那个标识符原样提示出来（形参提示形参自己）。

    链路是这样的：`extract_word_at` 有一条兜底——光标左边没有词、但右边紧挨
    着标识符时取右边的词（v44 为"光标停在词首、按 Delete 从头删字"引入的），
    而 trigger 的入场检查 `_ident_char_before_caret` 也据此放宽到"光标前【或】
    后一个字符是标识符即可"。于是做完这次编辑后的那一帧快照，
    "左边是分隔符、右边是标识符"——**与"从词首删字"长得一模一样**，
    单看快照永远区分不开。

    唯一的差别在【怎么变的】：
      * 从词首删字 -> 行【变短】（是删除）；
      * 敲空格/标点  -> 行【变长】，且插进来的这一段不含标识符字符。
    这里就是靠这个差别把它拦下来的。删字的路径完全不受影响。
    """
    ins = line_edit_inserted(old_line, new_line)
    if not ins:
        return False
    return not any(ch.isalnum() or ch == "_" for ch in ins)


# ---------------------------------------------------------------------------
# v55：VBE 自己会弹「自动列出成员」的位置 —— 我们让位
# ---------------------------------------------------------------------------
#
# VBE 的提示列表【没有独立窗口】：现场实测（发 Ctrl+J 后全系统扫描）一个
# 新窗口都没出现，代码窗格底下连个编辑控件子窗口都没有 —— 那个列表是 VBE
# 自己画在代码窗格上的。所以窗口探测（找类名 / 找属主弹窗 / 探光标下方窗口）
# 既探不到真列表，又会把别的程序的无标题栏浮窗误当成它（v54 就这么把工具
# 搞成"不管输入什么都不弹"）。
#
# 改成看【光标所在处的语法位置】。VBE 会自己弹列表的位置里，最确定、也最常
# 跟我们撞车的就是成员访问：光标停在 `标识符.` 之后、正在输入成员名 ——
# `UserForm1.` / `Me.` / `Sheet1.Range(...).` 全是这种。
#
# 只认这两种，宁可少拦也别再弄成"一直不弹"：让位只在那些位置成立，成员名 /
# 类型名一打完（空格、等号、换行）立刻恢复。

# 点号左边允许出现的"能取成员的东西"末尾字符
_MEMBER_OWNER_TAIL = "_)]}"


def _in_comment_or_string(line_text, col):
    """光标（1-based 列）左边是否落在字符串 / 注释里（极简扫描）。

    VBE 在注释和字符串里不会弹列表，那种位置我们照旧弹自己的。
    """
    if not line_text or not col or col < 1:
        return False
    in_str = False
    n = min(col - 1, len(line_text))
    for i in range(n):
        c = line_text[i]
        if in_str:
            if c == '"':
                in_str = False
        elif c == '"':
            in_str = True
        elif c == "'":
            return True                  # 注释一直延续到行尾
    return in_str


def vbe_member_list_expected(line_text, caret_col):
    """光标是否正停在 VBE 会自己弹【成员列表】的位置（`标识符.` 之后）。

    判据：光标左边形如 `xxx.` + 正在输入的成员名（成员名可以还没开始打）。
    纯文本判定，不碰窗口、不碰 COM，可单测。
    """
    if not line_text or not caret_col or caret_col < 1:
        return False
    i = min(caret_col - 1, len(line_text))
    # 1) 跳过光标左边正在输入的成员名
    j = i - 1
    while j >= 0 and (line_text[j].isalnum() or line_text[j] == "_"):
        j -= 1
    # 2) 成员名左边应当是点号（`X .` 这种带空格的写法也认）
    k = j
    while k >= 0 and line_text[k] in " \t":
        k -= 1
    if k < 0 or line_text[k] != ".":
        return False
    # 3) 点号左边得是能取成员的东西：标识符 / ) / ] / }
    m = k - 1
    while m >= 0 and line_text[m] in " \t":
        m -= 1
    if m < 0:
        return False
    if not (line_text[m].isalnum() or line_text[m] in _MEMBER_OWNER_TAIL):
        return False
    # 4) `1.5` 这种小数点：点号左边整段是纯数字，VBE 不弹列表
    s = m
    while s >= 0 and (line_text[s].isalnum() or line_text[s] == "_"):
        s -= 1
    if line_text[s + 1:m + 1].isdigit():
        return False
    # 5) 注释 / 字符串里 VBE 不弹，我们照旧弹
    if _in_comment_or_string(line_text, k + 1):
        return False
    return True


def _in_type_list_position(line_text, caret_col):
    """光标是否停在 `As ` / `New ` 之后 —— VBE 在这里会弹【类型列表】。

    `Dim x As ` / `Set c = New ` 之后敲的第一个字母就会把 VBE 的类型列表唤出来，
    跟我们的弹窗撞在同一处，同样要让位。判据刻意简单：光标左边（去掉正在输入
    的那个词之后）以 `As` 或 `New` 收尾即可，不去看整句是不是合法声明 ——
    `As` / `New` 是关键字，正常代码里不会有别的以它收尾的位置。
    """
    if not line_text or not caret_col or caret_col < 1:
        return False
    i = min(caret_col - 1, len(line_text))
    j = i - 1
    while j >= 0 and (line_text[j].isalnum() or line_text[j] == "_"):
        j -= 1
    head = line_text[:j + 1].rstrip().lower()
    if not (head.endswith(" as") or head.endswith(" new")):
        return False
    # 注释 / 字符串里 VBE 不弹，我们照旧弹
    return not _in_comment_or_string(line_text, j + 2)


def vbe_list_expected(line_text, caret_col):
    """光标是否正停在 VBE 会自己弹列表的位置（是则我们让位）。

    两种位置：
      1. 成员访问 —— 光标停在 `标识符.` 之后（含成员名输入中）；
      2. 类型位置 —— 光标停在 `As ` / `New ` 之后填类型名。
    """
    return (vbe_member_list_expected(line_text, caret_col)
            or _in_type_list_position(line_text, caret_col))


def replace_word(line_text, word_start_col, caret_col, completion):
    """把 [word_start_col, caret_col) 区间的单词替换为 completion，返回新行文本。"""
    s = word_start_col - 1
    e = caret_col - 1
    return line_text[:s] + completion + line_text[e:]


def _normalize_scoped(ids):
    """把 backend 返回的记录统一成 (name, module, proc, priv)。

    兼容三种形态：
      - "x" 纯名字              -> ("x", None, None, False)  全局可见
      - ("x", "ProcA") 旧二元组  -> ("x", None, "ProcA", False)
      - ("x", "M1", "ProcA", T) -> 原样（真实 VBE 完整记录）
    module=None 表示"未知/旧式"：模块级时引擎视作全局可见，
    过程级时只要过程名匹配即可见（不限制模块）。
    """
    out = []
    for item in ids or []:
        if isinstance(item, (tuple, list)) and len(item) >= 4:
            out.append((item[0], item[1], item[2], bool(item[3])))
        elif isinstance(item, (tuple, list)) and len(item) >= 2:
            out.append((item[0], None, item[1], False))
        else:
            out.append((item, None, None, False))
    return out


def filter_identifiers_by_scope(scoped_ids, current_proc, current_module=None):
    """
    按作用域过滤标识符，返回当前位置可见的名字列表（保序去重）。

    规则（VBA 官方可见性语义）：
      过程内局部变量 / 参数（proc 非空）：
        - 仅当 current_proc == proc，且模块一致（module=None 视为通配）时可见；
      模块级（proc 为空）：
        - module 未知（None）或 current_module 未知（None）：视为全局可见（旧式）；
        - 定义模块 == 当前模块：一律可见（本模块 Private/Public 都看得到）；
        - 定义模块 != 当前模块：仅 priv=False（Public）的可见；
          priv=True（Private）的隐藏 —— 这就是"不提示其他模块的私有变量"；
      同名遮蔽优先级：本过程局部(4) > 本模块模块级(3) > 其他模块 Public/旧式全局(2)。
    """
    scoped = _normalize_scoped(scoped_ids)
    cur = (current_proc or "").lower() or None
    cm = (current_module or "").lower() or None

    ranked = {}   # name.lower -> (rank, display_name)

    def consider(name, rank):
        key = name.lower()
        if key not in ranked or rank > ranked[key][0]:
            ranked[key] = (rank, name)

    for name, module, proc, priv in scoped:
        pl = (proc or "").lower() or None
        ml = (module or "").lower() or None
        if pl is not None:
            # 过程内局部变量/参数
            if cur and pl == cur and (cm is None or ml is None or ml == cm):
                consider(name, 4)
            continue
        # 模块级
        if ml is None or cm is None:
            consider(name, 2)               # 旧式/信息不足：全局可见
        elif ml == cm:
            consider(name, 3)               # 本模块模块级（含 Private）
        elif not priv:
            consider(name, 2)               # 其他模块的 Public
        # 其他模块的 Private -> 直接跳过，不提示

    return [ranked[k][1] for k in ranked]


def _word_boundaries(name):
    """名字里的"词边界"下标集合。

    边界 = 首字符 / 下划线之后 / 小写→大写的切换处（camelCase 驼峰）。
    例如 dataSheet -> {0, 4}，DataArr -> {0, 4}。

    【必须用原始 name，不能先 lower】——大小写切换信息一丢，
    就再也认不出 camelCase 的词首了，`ds` 想命中 `dataSheet` 的 D 和 S 就无从谈起。
    """
    b = set()
    if not name:
        return b
    b.add(0)
    prev = name[0]
    for i in range(1, len(name)):
        ch = name[i]
        if prev == "_" or (prev.islower() and ch.isupper()):
            b.add(i)
        prev = ch
    return b


def _greedy_positions(low, query, bounds):
    """无法整块连续命中时，逐字符挑位置：优先词边界，其次紧接上一个命中。

    档位越小越优先（同档取更靠前的）：
      0 既接在上一个命中之后、又落在词边界
      1 落在词边界（首字母缩略型：ds -> dataSheet）
      2 紧接上一个命中（尽量连成块）
      3 其它

    【v59 修复】纯贪心会【漏判】：挑位置时只看"这一档更优"，不管后面还凑不凑
    得齐，于是 query 明明是子序列也会被判成不匹配。真实案例（用户报的）：
      getSettingTitle3 里输入 getse3 —— 第 3 个字符 t 贪心跳到词边界 T(10)
      （档 1 优于接着用的 t(2) 档 2），再往后就没有 s 了，直接 return None；
      可 [0,1,2,3,4,15]（g-e-t-s-e-…-3）本来就成立，理应提示。
    修法：先反向算出每个字符【最靠右能放的下标】maxpos[t]，它保证 query[t+1:]
    在其后仍能匹配完；正向挑位置时把候选限制在 i..maxpos[t] 之内。
    这样既保住"优先词边界"的观感，又不可能再漏判（只要 query 真的是子序列，
    区间一定非空，因为 low[maxpos[t]+1:] 是 low[i:] 的后缀）。
    """
    n = len(low)
    k = len(query)
    # 反向可行性上界：maxpos[t] = query[t] 的最靠右合法位置
    maxpos = [0] * k
    limit = n - 1
    for t in range(k - 1, -1, -1):
        j = low.rfind(query[t], 0, limit + 1)
        if j < 0:
            return None                      # 连子序列都不是
        maxpos[t] = j
        limit = j - 1
    pos = []
    i = 0
    for t, ch in enumerate(query):
        hi = maxpos[t]
        best = None
        best_key = None
        for j in range(i, hi + 1):
            if low[j] != ch:
                continue
            if j == i and j in bounds:
                key = (0, j)
            elif j in bounds:
                key = (1, j)
            elif j == i:
                key = (2, j)
            else:
                key = (3, j)
            if best_key is None or key < best_key:
                best_key, best = key, j
        if best is None:
            return None
        pos.append(best)
        i = best + 1
    return pos


def fuzzy_match(name, query):
    """模糊匹配：query 的字符按顺序出现在 name 里即算命中（不要求开头）。

    返回 (score, positions)；不匹配返回 None。

      score     排序键（元组，越大越靠前）：
                  (kind, -首个命中位置, -命中跨度, -名字长度)
                kind: 4 完全相等 > 3 前缀 > 2 词首缩略 > 1 连续块 > 0 分散
      positions 命中的字符下标（升序），供 UI 高亮

    匹配策略（命中位置怎么挑，直接决定高亮好不好看）：
      1) 优先【整块连续】——query 作为子串出现时直接用它，
         视觉上就是"我打的这几个字母连在一起"。这是最常见的直觉。
         例：arr -> dataArr 命中 Arr（而不是散着的 a..r..r）。
      2) 否则逐字符贪心，优先落在词边界上（首字母缩略型）。
         例：ts -> dataSheet 命中 t 和 S（datasheet 里没有连续的 ts）。
    """
    if not name or not query:
        return None
    low = name.lower()
    q = query.lower()
    if len(q) > len(low):
        return None
    bounds = _word_boundaries(name)

    contiguous = False
    # 1) 整块连续：取"最靠前、且尽量落在词边界"的那一次出现
    starts = []
    i = low.find(q)
    while i >= 0:
        starts.append(i)
        i = low.find(q, i + 1)
    if starts:
        best = min(starts, key=lambda s: (0 if s == 0 else
                                          (1 if s in bounds else 2), s))
        positions = list(range(best, best + len(q)))
        contiguous = True
    else:
        # 2) 分散：逐字符按档位挑
        positions = _greedy_positions(low, q, bounds)
        if positions is None:
            return None

    if low == q:
        kind = 4                                  # 完全相等
    elif low.startswith(q):
        kind = 3                                  # 前缀
    elif all(p in bounds for p in positions):
        kind = 2                                  # 词首缩略（ds -> dataSheet）
    elif contiguous:
        kind = 1                                  # 连续块
    else:
        kind = 0                                  # 分散
    span = positions[-1] - positions[0] + 1
    score = (kind, -positions[0], -span, -len(low))
    return score, positions


# v68：宿主类型库枚举常量放行"跳步模糊"的门槛 —— 输入里必须出现过一段
# 【连续】的名字片段，且长度不小于这个值。见 Completer.trigger 里那段说明。
# 取 4 与 v61 内建名字"纯分散命中要 >= 4 个字符"那档保持一致。
HOST_ENUM_FUZZY_RUN = 4


def _longest_run(positions):
    """命中下标（升序）里最长的那段连续长度。

    `xlworkfaul` 打在 `xlWorkbookDefault` 上是 [0,1,2,3,4,9,12,13,14,15]，
    最长连续段是 6（`xlwork`；后面 `faul` 又是 4）—— 说明用户确实在打这个名字，
    只是中间跳过了 `book`。而 `xlce` 打在 `xlVAlignCenter` 上只有断断续续的单点，
    最长连续段 1 —— 那就只是"碰巧凑得出子序列"，不算。
    """
    if not positions:
        return 0
    best = cur = 1
    for a, b in zip(positions, positions[1:]):
        cur = cur + 1 if b == a + 1 else 1
        if cur > best:
            best = cur
    return best


def _name_exists_outside_caret(backend, word, line_no, caret_col):
    """把光标处的词抹掉之后，这个 name 在工程里还存不存在。

    这是区分「真实存在的标识符」与「正在输入的回声」的唯一可靠问法，
    由后端实现（可选接口 `backend.name_exists_outside_caret(name, caret)`）。

    后端没实现时退化为 False —— 保守起见按"不存在"处理，等价于旧的
    "不是声明过的名字就剔除本身"语义，行为不会比修复前更糟。
    """
    hook = getattr(backend, "name_exists_outside_caret", None)
    if not callable(hook) or not word or not line_no or not caret_col:
        return False
    try:
        return bool(hook(word, (int(line_no), int(caret_col))))
    except Exception:
        return False


# 弹窗一屏显示几行。**这是窗口高度，不是候选上限** —— 候选再多也不会被丢弃，
# 超出的排在窗口外，靠上下键 / 滚轮滚动查看。
#
#
# v49 恢复 15：v47 取 9 只是为了对齐「数字键 1~9 选词」，该功能已移除。
VIEW_ROWS = 15


class Completer:
    def __init__(self, backend, ui, view_rows=VIEW_ROWS):
        self.backend = backend
        self.ui = ui
        self.visible = False
        self.matches = []
        self.selected = 0
        self.ctx = None
        self.word_is_complete = False
        self._suppress_until = 0.0
        self.shown_at = 0.0       # 最近一次真正弹出的时刻（供"刚弹出保护期"使用）
        # 当前候选各自的命中下标：{名字: [下标, ...]}，供 UI 把命中的字符标红。
        # 由 trigger() 填写、hide() 清空。UI 通过 completer 读取，不占用 show() 签名。
        self.match_hits = {}
        # ---- 滚动视口 ----
        # 弹窗高度固定显示 view_rows 行，但候选永远【不会】被截断：
        # 多余的排在窗口外，靠上下键/滚轮滚动查看。view_top 是当前窗口的起始下标。
        self.view_rows = max(1, int(view_rows or VIEW_ROWS))
        self.view_top = 0
        # 结构性名字的本次触发缓存（v67）。见 _structural_names 的说明。
        self._struct_cache = None

    def _clamp_top(self):
        """把窗口起点夹到合法区间（不能越过列表末尾）。"""
        n = len(self.matches)
        if n <= self.view_rows:
            self.view_top = 0
            return
        last_top = n - self.view_rows
        if self.view_top < 0:
            self.view_top = 0
        elif self.view_top > last_top:
            self.view_top = last_top

    def _scroll_to_selected(self):
        """滚动最小的幅度，让选中项落在窗口内（尽量保持在中间之外不折腾用户）。"""
        self._clamp_top()
        n = len(self.matches)
        if n <= self.view_rows:
            self.view_top = 0
            return
        if self.selected < self.view_top:
            self.view_top = self.selected
        elif self.selected >= self.view_top + self.view_rows:
            self.view_top = self.selected - self.view_rows + 1
        self._clamp_top()

    def visible_rows(self):
        """当前这一屏要画出来的候选（最多 view_rows 个，不会掐掉剩下的）。"""
        top = max(0, min(self.view_top, max(0, len(self.matches) - 1)))
        return list(self.matches[top:top + self.view_rows])

    def view_selection(self):
        """选中项在【当前这一屏】里的行号（0-based），UI 用它决定高亮哪一行。"""
        return max(0, self.selected - self.view_top)

    def view_info(self):
        """(候选总数, 当前窗口起点) —— UI 画滚动条用。"""
        return (len(self.matches), max(0, self.view_top))

    def set_view_top(self, top):
        """鼠标拖动纵向滚动条时，按拖到的位置设置窗口起点 view_top。

        top 会被夹到合法区间；同时把选中项夹进当前可见窗口，
        保证高亮始终可见、Tab 确认的是看得见的候选。
        """
        if not self.visible or not self.matches:
            return
        self.view_top = max(0, int(top))
        self._clamp_top()
        n = len(self.matches)
        lo = self.view_top
        hi = min(self.view_top + self.view_rows - 1, n - 1)
        if self.selected < lo:
            self.selected = lo
        elif self.selected > hi:
            self.selected = hi
        if self.selected < 0:
            self.selected = 0
        self.ui.update_selection(self.selected)

    def _type_name_set(self):
        """后端给出的"类型名"集合（小写），供 `As` 之后的位置使用。

        可选接口：backend.get_type_names() -> 可迭代的名字序列。
        后端没实现就返回空集，类型名位置保持静默（等同 v22 行为）。
        """
        hook = getattr(self.backend, "get_type_names", None)
        if not callable(hook):
            return set()
        try:
            return set(str(n).lower() for n in (hook() or ()))
        except Exception:
            return set()

    def _live_names_outside_caret(self, ctx=None, scope_only=False):
        """现场扫描当前模块：抹掉光标处的词之后，里面还出现过的名字（小写）。

        可选接口 backend.names_outside_caret(caret, scope_only=False)。
        返回 None 表示【现场证据不可用】（后端没实现 / 读不到），
        此时调用方退回"声明集合"这一路证据（行为等同旧版，不会更糟）。

        scope_only=True（v51）：只要【光标处可见的那个作用域】里的出现 ——
        模块级行 + 光标所属过程的行，别的过程体内的一律不算。用于"正在过程内
        声明一个名字"的场景（形参 / 局部 Dim），详见 trigger 里的说明。
        后端没实现这个参数时退回全模块扫描（旧行为，不会更糟）。
        """
        if ctx is None:
            ctx = self.ctx or {}
        hook = getattr(self.backend, "names_outside_caret", None)
        if not callable(hook):
            return None
        line_no, col = ctx.get("line_no"), ctx.get("caret_col")
        if not line_no or not col:
            return None
        try:
            try:
                got = hook((int(line_no), int(col)), scope_only)
            except TypeError:
                got = hook((int(line_no), int(col)))   # 旧式单参实现
            if got is None:
                return None       # 后端读不到（COM 失败）
            return set(str(n).lower() for n in got)
        except Exception:
            return None

    def _live_evidence_available(self):
        """后端是否支持"现场扫描当前模块"（names_outside_caret）。

        幽灵判定（候选恰好等于"收集时光标处的词"就剔除）只在现场证据可用时
        才敢下结论：拿不到现场证据时，那个词也可能确实是别的模块里的真名字，
        宁可保守放过，也不要误杀候选。
        """
        return callable(getattr(self.backend, "names_outside_caret", None))

    def _caret_word_at_collect(self):
        """标识符池被解析那一刻，光标处的词（小写；后端不提供则空串）。"""
        hook = getattr(self.backend, "caret_word_at_collect", None)
        if not callable(hook):
            return ""
        try:
            return str(hook() or "").lower()
        except Exception:
            return ""

    def _name_really_exists(self, name, ctx=None, live=None, declared=None):
        """这个名字在工程里是【真实存在】的，还是只是光标处的回声？

        两条互为补充的证据，命中任一即认为真实存在：

        1. `get_declared_names()` 里有它 —— 全工程范围内被 Dim/Const/Sub/
           Function/Type/Enum 真正声明过。跨模块的 Public 名字靠这一条。
        2. 现场扫描当前模块（把光标处的词抹掉后）还能找到它 —— 隐式变量
           （用到即存在）、以及解析层没能归类的声明，靠这一条兜住。
           v51：调用方可能传入【按作用域收窄过】的 live（只含模块级 + 光标
           所属过程里的出现）。此时"别的过程里的同名局部变量"不算证据 ——
           它不是光标处的合法引用，只是文字巧合（详见 trigger 里的 scope_only）。

        只用 1 会漏判隐式变量（正是 v35"打全 arr 列表消失"的成因）；
        只用 2 会漏判跨模块的 Public 名字（`names_outside_caret` 刻意只扫
        当前模块以省开销）。所以两条都要。

        v43 加固：缓存里的"收集那一刻光标处的词"**不算**存在证据。它很可能
        正是用户此刻在回退删除的那个词（候选池还没刷新），若把它当真名字保
        下来，就会"回退时提示出更长的旧片段"。所以这条要排在声明集合之前
        否决 —— 但只要现场能扫到（真在别处存在），仍然算数（先判 live）。

        live / declared 可由调用方传入复用，避免每个候选重复打 COM。
        """
        low = str(name).lower()
        if not low:
            return False
        # v60：组件名（模块 / 窗体 / 类 / 文档模块名）、窗体控件名，以及 v61 加
        # 上的 VBA 语言自带名字（内建函数 / 常量 / 数据类型）—— 它们的共同点是
        # 【由工程结构或语言运行时决定存在，与代码文本无关】，所以直接判为真名字，
        # 不走下面那套"现场文本 / 声明归属"的证据链。
        #
        # 不加这一条会怎样（用户报的两个 bug）：
        #   * 插入一个 UserForm1 后 VBE 会自动打开该窗体的代码窗口，用户在
        #     【窗体自己的代码里】输入 use —— 而 use 是 userform1 的连续子串，
        #     于是 UserForm1 命中"回声候选"；接着 declared_elsewhere 说它"只声明
        #     在本模块"，现场文本里又（通常）压根没出现过 UserForm1 这个词，两条
        #     证据都不成立，候选被剔掉 —— 症状就是"窗体名死活不提示"，而同样的
        #     输入在别的模块里却一切正常（那里 declared_elsewhere 返回 True）。
        #     刚拖上去的 Label1 同理。
        #   * 输入 ms 想补 MsgBox —— ms 同样是 msgbox 的连续子串；只要这个模块里
        #     还没写过 MsgBox，现场就扫不到它，候选照样被剔。VBA 内建名字由语言
        #     提供，代码里没调用过也永远合法，必须走这一路。
        if low in self._structural_names():
            return True
        if declared is None:
            try:
                declared = self._declared_names()
            except Exception:
                declared = set()
        if live is None:
            live = self._live_names_outside_caret(ctx)
        if live is None:
            # 现场证据不可用：退回"声明集合"这一路（旧行为），
            # 绝不因为拿不到现场数据就把候选误杀。
            return low in declared
        if low in live:
            return True                      # 现场就找得到 -> 铁证
        # 后端能区分"声明在哪个模块"时（v43）：声明证据只认【别的模块】。
        # 只在当前正在编辑的模块里声明过的名字，一律以现场文本为准 —— 现场
        # 扫不到就说明它已被删掉，声明集合里那一份只是 0.4 秒前的旧快照。
        ext = None
        hook = getattr(self.backend, "declared_elsewhere", None)
        if callable(hook):
            try:
                ext = hook(name, (ctx or self.ctx or {}).get("module_name"))
            except Exception:
                ext = None
        if ext is True:
            return True                      # 别的模块声明过 -> 跨模块真名字
        if ext is False:
            return False                     # 只在本模块声明过，而现场已经没有
        # 幽灵判定只在【现场证据可用】时才敢下结论（见 _live_evidence_available）：
        # 候选恰好等于【收集那一刻光标处的词】、而现场又扫不到它 —— 它就是用户
        # 正在回退删除的那个词的旧版本，缓存里的声明集合很可能还残留着它。
        if (self._live_evidence_available()
                and low == self._caret_word_at_collect()):
            return False
        return low in declared

    def _module_level_names(self):
        """标识符池里在【模块级】存在的名字（小写集合）。

        v53 用：判断"光标处的这个词是不是全模块可见的真名字"。判据是标识符池里
        有一条该名字的记录、且其过程名为空（模块声明区的 Dim/Const/Sub/Function…
        或组件名）。过程内的局部变量、形参（proc 非空）都不算。

        为什么不问后端要一个专门接口：`get_identifiers()` 是【带缓存】的，
        同一次 trigger 早已取过（作用域过滤用的就是它），这里只是换个角度筛一遍，
        一次 COM 都不会多打。
        """
        try:
            recs = _normalize_scoped(self.backend.get_identifiers())
        except Exception:
            return set()
        return set(str(n).lower() for n, _m, p, _pv in recs if not p)

    def _declared_names(self):
        """工程里【真实声明】过的名字（小写集合），供 trigger 区分"真名字"与"幻影"。

        只有被 Dim / Const / Sub / Function / Type / Enum 等真正声明过的名字才算；
        隐式变量（不写 Option Explicit 时"用到即存在"）不算——否则无法剔除"回退
        删字回退出来的"未定义词（例如 numA）。

        可选接口：backend.get_declared_names() -> 可迭代名字序列。后端没实现时退回
        "全部可见标识符都算已声明"，仅用于旧式/测试后端兜底，不影响真实路径
        （真实 VbeBackend 会提供 get_declared_names）。
        """
        hook = getattr(self.backend, "get_declared_names", None)
        if callable(hook):
            try:
                return set(str(n).lower() for n in (hook() or ()))
            except Exception:
                pass
        # 兜底（旧式 / 测试后端）：把【全部可见标识符】都算已声明。
        # 记录可能是 "x" / ("x", scope) / ("x", mod, proc, priv) 三种形态，
        # 必须用 _normalize_scoped 取出【名字】——直接 str(整条记录) 会得到
        # "('num', 'M1', 'ProcA', False)" 这种垃圾，等于永远查不到，
        # 会把本该保留的前缀候选全部误杀。
        try:
            return set(str(r[0]).lower()
                       for r in _normalize_scoped(self.backend.get_identifiers()))
        except Exception:
            return set()

    def _structural_names(self):
        """由【工程结构 / 语言运行时】决定其存在的名字（小写集合）。

        三类：组件名 + 窗体控件名（v60）+ VBA 语言自带的名字（v61：内建函数 /
        内建常量 / 内建数据类型）。

        v60。与 _declared_names 的分工：
          * _declared_names 答的是"代码文本里真声明过什么"；
          * 本方法答的是"本来就有、与代码文本无关的东西"。
        刚拖上去的 Label1（一个字代码都还没有）、一行代码都没有的新窗体
        UserForm1、代码里从没调用过的 MsgBox，只在前者里出现，不在后者里。

        为什么需要单独一路证据：回声防护会先假定"以当前词为连续子串的候选"
        可能是被删剩下的旧版本，再用"它在工程里真实存在吗"来裁决。裁决依赖
        现场文本与声明归属，而【结构性命名的存在与代码文本无关】——窗体代码里
        从不出现 UserForm1 这个词，它照样是合法引用。少了这一路证据就会误杀
        （用户报的"在窗体里输入 use 死活不提示 UserForm1"，详见
        _name_really_exists）。

        可选接口；后端不提供时返回空集，旧式/测试后端行为完全不变。

        ⚠️ v67：结果按【一次触发】缓存。回声防护会对每个候选问一次
        _name_really_exists()，而它第一步就要查这个集合 —— 候选一多就是
        O(候选数 × 集合大小)。v67 收进宿主枚举常量后集合涨到 4600 多条，
        输入 xl 这种会一次带出 2000 多个候选，逐次重建集合直接拖到 2 秒
        （真机实测 2135ms）。触发入口处会把它置空，所以每次触发仍然是最新
        快照（候选池本身按 _CACHE_TTL 刷新，刷新粒度不变）。
        """
        if self._struct_cache is not None:
            return self._struct_cache
        hook = getattr(self.backend, "get_structural_names", None)
        if not callable(hook):
            self._struct_cache = set()
            return self._struct_cache
        try:
            got = hook()
        except Exception:
            got = None
        # 后端直接交出内部集合时不复制 —— 复制一份 4600 条的集合在每个候选上
        # 各来一遍，正是上面那个 2 秒的来源。
        if isinstance(got, (set, frozenset)):
            self._struct_cache = got
        else:
            self._struct_cache = set(str(n).lower() for n in (got or ()))
        return self._struct_cache

    def _builtin_names(self):
        """后端给出的"VBA 语言自带名字"集合（小写）。可选接口。

        v61。引擎用它给这批名字单独收紧一档模糊匹配 —— 它们是语言自带的公共
        词汇（内建函数 / 常量 / 数据类型，近 300 个），什么字母组合都能在里面
        找到子序列。详见 trigger 里的说明。

        后端不提供时返回空集 -> 不做任何收紧（旧式 / 测试后端行为完全不变）。
        """
        hook = getattr(self.backend, "get_builtin_names", None)
        if not callable(hook):
            return set()
        try:
            return set(str(n).lower() for n in (hook() or ()))
        except Exception:
            return set()

    def _host_enum_names(self):
        """后端给出的"宿主类型库枚举常量" -> {小写名: 最短输入长度}。可选接口。

        v67。用户报"输入 vb 会提示一堆 VBA 枚举值，输入 xl 却一个 xl 开头的
        都不提示"：vb* 是 v61 收的 VBA 内建常量，xl* 属于宿主 Excel 类型库，
        一直没收。收进来的同时必须收紧匹配口径 —— 这批名字数量极大
        （实测 Excel + Office 共 4538 条），若像工程内名字那样接受跳步匹配，
        三四个字母的输入就能带出几百条无关项（实测 输入 ms -> 2365 条 mso*，
        输入 count -> 130 条、输入 cell -> 174 条）。所以这批【只认前缀】：
        候选必须从头开始以所输的词开头，且输入长度不低于家族前缀长度
        （xl / mso）。详见 trigger 里的快速路径。

        后端不提供时返回 {} -> 不做任何限制（旧式 / 测试后端行为完全不变）。
        约定：后端给的键必须已经是小写（VbeBackend 就是这么给的）。
        """
        hook = getattr(self.backend, "get_host_enum_names", None)
        if not callable(hook):
            return {}
        try:
            got = hook()
        except Exception:
            return {}
        return got if isinstance(got, dict) else {}

    def _vbe_popup_showing(self):
        """VBE 自带的【成员列表】此刻是否显示着 —— 是则我们让位（v63→v65）。可选接口。

        与 vbe_list_expected（语法判据）的分工：
          * vbe_list_expected 猜的是"VBE 大概会在这里弹列表"——只覆盖成员访问与
            As/New 类型位置，VBE 弹别的东西（最典型的【参数信息】：`MsgBox "x",`
            敲逗号后弹出的参数签名）它一概不认；
          * 本方法看的是"VBE 的提示窗现在是不是真的画在屏幕上"——后端直接查
            VBE 那两个预建复用窗口（NameListWndClass / PopupTipWndClass）的可见性。

        两者是【并集】：任一说该让位就让位。语法判据保留原样（它覆盖"列表马上
        要出现"那一小段窗口可见之前的时序），本方法补齐它漏掉的场景。

        ⚠️ v65 收窄：**只有成员列表让位，参数信息（形参签名）不让位。**
        用户 v65 澄清 v63 的原意是"只有弹成员列表的时候才让位"。
        理由：成员列表本身就是候选列表 —— 与我们同质，叠在一起既遮挡又抢键盘；
        而参数信息只是一行只读签名，不吃键盘，且它出现时用户正在填实参，
        正是最需要候选的时候（在那儿让位 = "输入任何字符都无提醒"）。

        接口取值顺序：优先 vbe_yield_visible()（新，只认成员列表）；
        后端没有时退回 vbe_popup_visible()（旧，任何提示窗）—— 旧式 / 测试后端
        行为不变（与 v62 一致）。两者都没有则返回 False。
        """
        for _name in ("vbe_yield_visible", "vbe_popup_visible"):
            hook = getattr(self.backend, _name, None)
            if not callable(hook):
                continue
            try:
                return bool(hook())
            except Exception:
                return False
        return False

    def _vbe_popup_desc(self):
        """VBE 提示窗的明细，仅用于日志（后端没提供该接口时返回 None）。"""
        hook = getattr(self.backend, "vbe_popup_info", None)
        if not callable(hook):
            return None
        try:
            return hook()
        except Exception:
            return None

    def trigger(self, require_ident_before_caret=False):
        """尝试弹出补全列表。

        require_ident_before_caret=True：由「编辑器内容变化」轮询触发时使用。
        此时无法从按键判断用户输的是什么，需要检查光标前一个字符——若是空格/
        标点/换行，说明不在拼标识符，应当收起弹窗（等效于按空格收起）。
        键盘事件触发路径不需要该检查，按键本身已能区分。

        采用轮询而不仅靠键盘事件的原因：中文经输入法(IME)输入时，键盘钩子拿
        到的往往是 VK_PROCESSKEY 或空字符（尤其 on_release），导致中文永不触发。
        直接读编辑器文本变化则与输入法无关，中英文一律可靠。
        """
        # 刚确认过补全：静默期内不再弹，避免"收起后立刻又冒出来"
        if time.time() < self._suppress_until:
            return
        # 结构性名字的集合按"一次触发"取一份快照（v67）：本次触发里所有回声
        # 判定复用同一份，别再逐个候选重建（那会退化成 O(候选数 × 集合大小)）。
        self._struct_cache = None
        ctx = self.backend.get_context()
        if ctx is None:
            self.hide()
            return
        if ctx.get("in_string") or ctx.get("in_comment"):
            self.hide()
            return
        # 声明里 `As` 类型名位置：正在填数据类型名（如 `Dim x As Inte...`）。
        # 这里只提示真正的"类型名"（窗体名 / 类模块名 / 标准模块名 / Type / Enum 名），
        # 变量与过程名依旧静默（v22 的初衷：填类型名时冒出一堆变量名是噪音）。
        in_type = bool(ctx.get("in_type_position"))
        type_names = self._type_name_set() if in_type else None
        # v55：光标停在 `标识符.` 之后输入成员名时，VBE 自己会弹「自动列出成员」
        # 列表 -> 我们让位，不跟它抢同一块地方（两个列表会互相遮挡、还抢键盘）。
        # 只拦【轮询自动触发】：手动 Ctrl+Space 是用户点名要我们的列表，照旧弹。
        if require_ident_before_caret and YIELD_TO_VBE_LIST \
                and vbe_list_expected(ctx.get("line_text"),
                                      ctx.get("caret_col")):
            _log("trigger: VBE 自带列表位置 -> 让位")
            self.hide()
            return
        # v63：VBE 自己弹着的东西【真的显示在屏幕上】时也让位。v65 收窄了范围。
        #
        # 上面那条走的是"语法位置"猜测，只覆盖成员访问（`标识符.`）与类型位置
        # （`As `/`New `）。VBE 在别处弹的提示它一概不认，于是两个窗叠在一起。
        #
        # 这条是精确判据：后端直接查 VBE 那两个预建复用提示窗
        # （NameListWndClass / PopupTipWndClass）的可见性，与光标语法位置无关。
        #
        # ⚠️ v65：只让【成员列表】。用户澄清 v63 的原意——"只有弹成员列表的时候
        # 才让位，弹形参签名不需要让位"。成员列表是候选列表（同质：遮挡 + 抢键盘），
        # 而参数信息（形参签名）只是只读提示、不吃键盘，且它出现时用户正在填实参，
        # 正是最需要候选的时候；在那儿让位就是"输入任何字符都无提醒"（v65 的 bug）。
        if require_ident_before_caret and YIELD_TO_VBE_LIST \
                and self._vbe_popup_showing():
            # 记下【是哪个窗】在让位（成员列表 / 参数信息），排查"某处什么都不弹"
            # 时必须能一眼分开（v65）。
            _log("trigger: VBE 成员列表窗可见 -> 让位 %r"
                 % (self._vbe_popup_desc(),))
            self.hide()
            return
        if require_ident_before_caret and not _ident_char_before_caret(ctx):
            self.hide()
            return
        word, start_col, end_col = extract_word_at(ctx["line_text"],
                                                    ctx["caret_col"])
        if not word or not (word[0].isalpha() or word[0] == "_"):
            self.hide()
            return
        # 只保留当前作用域可见的标识符（屏蔽其他过程局部变量 + 其他模块 Private）
        visible_ids = filter_identifiers_by_scope(
            self.backend.get_identifiers(),
            ctx.get("proc_name"),
            ctx.get("module_name"),
        )
        # 候选 = 输入串按顺序出现在名字里即命中（模糊匹配，不要求从头开始）。
        # 例：dataArr / dataSheet 两个变量——
        #   输入 ts -> dataSheet（datasheet 里没有连续的 "ts"，按词边界命中 t 和 S）
        #   输入 ta -> 两者都中（"ta" 在两边都是连续块）
        # 关键：就算单词已完整等于某个候选（如已把 cell 打全），也保留在列表里，
        # 弹窗继续显示，直到用户按 Tab 确认 / 鼠标点列表外 / 改成不再匹配的词。
        #
        # 打分排序：精确 > 前缀 > 词首缩略 > 连续块 > 分散，同级再比
        # "命中越靠前 / 跨度越小 / 名字越短"越好——保证最像的那个永远排第一，
        # 不会因为放开模糊匹配就让列表变成一锅粥。
        scored = []
        # v67：宿主类型库的枚举常量（xl* / mso* …）走【前缀专用】通道。
        #
        # 为什么不能跟工程内名字一起丢进 fuzzy_match：这批名字有 4500 多条、
        # 全是长复合词，任何三四个字母都能在里面凑出子序列。实测（真实工程）：
        #   输入 ms    -> 2365 条 mso*（用户要的其实是 MsgBox）
        #   输入 count -> 130 条 xlCount / xlCountryCode …
        #   输入 cell  -> 174 条 xlCell*
        #   输入 open  -> 492 条
        # 全是噪音，而用户口径是"宁可少提示，也不要噪音"。
        #
        # 所以这批只用【从头开始的连续前缀】命中，且输入长度不低于家族前缀长度
        # （后端给的下限：xl 是 2、mso 是 3）：
        #   xl / xlu / xlup -> xlUp 家族          ✓ 用户点名要的
        #   ms              -> 只出 MsgBox         ✓ mso* 被前缀长度挡在外面
        #   count / cell    -> 一条都不多出        ✓
        #
        # 顺带一个性能好处：前缀命中不必跑 _word_boundaries（正则），4500 个名字
        # 各跑一遍会把每次按键拖到 ~10ms，走快捷路径实测 ~2.8ms。
        #
        # 【v68】只认前缀又太死：用户报"输入 xlworkfaul，不提示 xlWorkbookDefault"。
        # 那不是前缀（xl work f aul，中间跳过了 book），但除了它整个 4538 条里
        # 一条都不匹配 —— 命中是唯一的，谈不上噪音，纯粹是被这条规则一刀切掉了。
        #
        # 放宽的口径：**输入必须从家族前缀开始**（xl / mso，这条不能松：正是它
        # 把 ms -> 2365 条 mso*、count -> 130 条、open -> 492 条这些噪音整片挡掉的，
        # 实测放开后 ms / count / open / cell / ar / vb 依旧是 0 条），
        # 然后再要求输入里【出现过一段连续的名字片段】（长度 >= HOST_ENUM_FUZZY_RUN）。
        # 意思是"你真的打过这个名字里的一段"，而不是拿三两个字母去凑子序列：
        #   xlworkfaul -> xlWorkbookDefault   连续块 6（xlwork / faul）  ✓ 用户要的
        #   xlms       -> xlMSDOS             连续块 4                    ✓
        #   xlco / xlce / xla / xlms / xlm    输入不够 4 个字符，
        #                                     退化成纯前缀，与 v67 完全一致（无新增噪音）
        #   xlcell     -> 仍以 xlCellType* 前缀命中排在最前，多出来的几条跳步命中
        #                 （xlLastCell 之类）排在 kind 0 档，不会顶掉首选
        #
        # 为什么是"连续块"而不是"输入够长就行"：实测放行纯跳步后，xlcell 会从
        # 15 条涨到 89 条、xlcount 从 4 条涨到 59 条 —— 全是"碰巧凑得出子序列"的
        # 长复合词。"连续块 >= 4"这条把短输入的噪音压回前缀水平，同时保住了
        # "用户确实在打这个名字"的那批（阈值 4 与 v61 内建名字那档一致）。
        host_min = self._host_enum_names()
        if not host_min:
            for i in visible_ids:
                r = fuzzy_match(i, word)
                if r is not None:
                    scored.append((r[0], i, r[1]))
        else:
            _lw = len(word)
            _qlow = word.lower()
            _hits = list(range(_lw))
            for i in visible_ids:
                _low = i.lower()
                _need = host_min.get(_low, 0)
                if _need:
                    if _lw < _need:
                        continue
                    if _low.startswith(_qlow):
                        # 前缀命中：kind 3（前缀）/ 4（完全相同），命中位置从 0 起
                        _kind = 4 if _lw == len(_low) else 3
                        scored.append(((_kind, 0, -_lw, -len(_low)), i, _hits))
                        continue
                    # v68：非前缀。先要求输入本身就从家族前缀开始（`xl...`）。
                    if not _qlow.startswith(_low[:_need]):
                        continue
                    r = fuzzy_match(i, word)
                    if r is None or _longest_run(r[1]) < HOST_ENUM_FUZZY_RUN:
                        continue
                    # 跳步命中的评分照 fuzzy_match 的来（通常是 kind 0 分散），
                    # 于是它天然排在所有前缀命中之后，不会顶掉首选。
                    scored.append((r[0], i, r[1]))
                    continue
                r = fuzzy_match(i, word)
                if r is not None:
                    scored.append((r[0], i, r[1]))
        # v61：VBA 内建名字（内建函数 / 常量 / 数据类型）+ v62 语言关键字，
        # 单独收紧一档匹配。
        #
        # 它们是【语言自带的公共词汇】，近 400 个，而且几乎任何字母组合都能在
        # 里面凑出子序列；若与"你自己工程里的名字"一样接受跳步匹配，任何 1~3 个
        # 字符的输入都会冒出一大串无关项 —— 真机实测（候选池 373 条）：
        #   输入 ar  -> vbAbortRetryIgnore / VarType / Partition 全挤进列表；
        #   输入 vbc -> 29 条，其中 22 条是这种"碰巧凑得出"的。
        #
        # 收紧规则：kind >= 1（连续块 / 词首缩略 / 前缀 / 精确）一律放行，纯分散
        # 命中（kind == 0）则要求输入至少 4 个字符。效果：
        #   ms / msg -> MsgBox           （前缀）      照常
        #   vbc      -> vbCrLf           （前缀）      照常
        #   spli     -> Split            （前缀）      照常
        #   instrr   -> InStrRev         （连续块）    照常
        #   formcur  -> FormatCurrency   （跳步，7 字符）照常
        #   slct     -> Select           （跳步，4 字符）照常
        #   ar       -> vbAbortRetryIgnore（跳步，2 字符）被挡掉 —— 正是要滤掉的噪音
        # 这条只作用于【内建公共词汇】；你自己工程里的名字不受限
        # （getse3 -> getSettingTitle3 照旧）。
        #
        # v62 起语言关键字与内建名字【共用这一档】（后端 get_builtin_names 把两者
        # 一并给出）：关键字同样是"打前缀就能补全"的用法，没有理由区别对待。
        if len(word) < 4:
            _bl = self._builtin_names()
            if _bl:
                scored = [t for t in scored
                          if t[0][0] >= 1 or t[1].lower() not in _bl]
        # sorted 稳定：同分时保持后端原来的顺序
        scored.sort(key=lambda t: t[0], reverse=True)
        matches = [name for _, name, _ in scored]
        # 每个候选的命中下标，交给 UI 高亮（键是名字，值是下标列表）
        self.match_hits = {name: pos for _, name, pos in scored}
        if in_type:
            matches = [i for i in matches if i.lower() in type_names]
        else:
            # 自我提示防护（声明行 / 用法行一律生效）：
            # 正在输入的词有可能只是"自己的回声"——光标处那几个字符被当成
            # 隐式变量收录了。此时把它提示出来等于"打什么提示什么"（回退出来
            # 的 numA 就是典型），必须剔除。
            #
            # 判定依据不能用"候选剩几个"（那是 v30 的老毛病：会把"打全名"误杀），
            # 也不能只用"是否真实声明过"——后者在真实工程里会漏判：名字明明在
            # 代码里到处在用（或声明在解析层没能归类的位置），却因为不在
            # 声明集合里被打全时一脚踢掉，正是"输入 arr 列表反而消失"的成因。
            #
            # 【v37】也不能反过来用"它是不是当前声明行正在起的新名字"来剔除
            # （那是 v31~v36 的老规则）：它同样会误杀——工程里已经有
            # `Function test()`，你在 `Sub test` 这一行把 test 打全，那不是回声，
            # 是货真价实的同名函数，必须继续提示。老规则的唯一作用是在
            # "正在声明的名字恰好又被本过程用到"时多剔除一次，而那种情况
            # 提示出来也无害（Tab 确认后写的就是同样的字符）。
            #
            # 正确问法只有一个：**这个名字在工程里真实存在吗？**
            #   存在 -> 合法候选，打全照常提示（arr / num / test 都保留）；
            #   不存在 -> 它只是回声，剔除（numA 不提示）。
            # 回声候选：候选名【包含当前词这一整段连续片段】——
            #   * 候选 == 当前词         精确回声（v30 起就管了）；
            #   * 候选 以当前词开头      从【结尾】删：剩下的是前缀（v43）；
            #   * 候选 以当前词结尾      从【开头】删：剩下的是后缀（v45）；
            #   * 候选 中间夹着当前词    头尾都删：剩下的是中段（v45）。
            # 统一判据：**当前词是候选名的连续子串**。用户删除字符只会让光标处的
            # 词变短，剩下的必然是原词的连续一段，不可能"跳着剩"——
            # 所以"连续子串"恰好等价于"这个候选可能是被删剩下的旧版本"。
            #
            # 为什么会有"残留片段"：标识符池是【带缓存】的（最快也要
            # ID_REFRESH_MIN_SEC 才重解析一次）。用户按住 Delete 连删时，池子里
            # 还留着"文字还长"那一刻解析出来的名字，于是弹窗会提示出一个比编辑器
            # 里实际内容【更长】的旧片段 —— 用户报的正是这个：
            #   从结尾删 -> 删到只剩 studengd 却提示 studengd...sgjd；
            #   从开头删 -> 删到只剩 ngdflk... 却提示完整的 studengd...sgjd。
            #
            # 判定候选是不是"真名字"，用两路证据（见 _name_really_exists）：
            #   1) 声明集合（跨模块 Public 名字只有它认得）；
            #   2) 现场扫描当前模块（隐式变量 / 解析层没归类的声明靠它）。
            # 加固：等于"收集那一刻光标处的词"、且现场扫不到的候选，一律当幽灵
            # 剔除 —— 缓存里的声明集合可能还残留着它（那一刻它确实写在声明行上），
            # 不能让它把幽灵保下来。
            #
            # 性能：只有"缓存里没有它"或"它就是那个光标旧词"时才需要现场扫描。
            # 正常打字路径候选都在声明集合里，一次 COM 都不会多打。
            low_word = word.lower()
            # 光标处【正在声明的名字】（形参 / 局部 Dim / 过程名 / Type 成员…），
            # 由后端解析当前声明行得到（ctx["decl_names"]）。
            decl_at_caret = set(str(n).lower()
                                for n in (ctx.get("decl_names") or ()))
            # 【v53】正在【起新名字】的位置不提示自己。
            #
            # 用户报的现象：形参位置上按退格 / 敲空格 / 把别的过程的变量名粘贴
            # 过来，弹窗都会把光标处这个名字原样提示出来（提示自己）。v51 把现场
            # 证据按作用域收窄、v52 拦住了"敲空格"那一路，但这两条判的都是
            # 【这次编辑是怎么变的】——粘贴一个完整名字与手动敲出最后一个字符，
            # 前后两帧快照逐字符等价，永远区分不开。堵不住的那部分只能从【语义】
            # 上堵：光标处这个词如果就是本行正在声明的名字（形参 / 局部 Dim /
            # 局部 Const），用户此刻是在【造名字】而不是【引用名字】，把它原样
            # 提示回去没有任何意义，还会盖住他刚敲的字。
            #
            # 只剔除【与光标处词完全相同】的候选：输入 num 仍要照常提示 numArr
            # （v37 的教训——按"本行声明的名字"整体剔除会把前缀补全一起误杀）。
            #
            # 模块级真名字除外（v37 语义：`Sub test` 那一行把 test 打全，工程里
            # 真有 Function test() 时照常提示；v51 的 updateValue 同理）：它们在
            # 模块级真的存在，提示出来是"重用已有名字"而不是"提示自己"。过程内
            # 的局部变量 / 形参则一律剔除——那正是"别的过程的变量名泄露到本位置"
            # 的来源（用户报的就是这种）。
            if low_word in decl_at_caret:
                if low_word not in self._module_level_names():
                    matches = [m for m in matches if m.lower() != low_word]
            # 若"正在输入的词"本身就是这里正在声明的名字，且声明发生在【过程内】，
            # 那它此刻的作用域是过程局部的 —— 现场证据也必须限定在同一个窗口里
            # 取（模块级 + 本过程），别的过程里的出现不算数。
            #
            # 不这么做会怎样（用户报的 bug）：ProcedureA 里有局部变量 username，
            # 你在 ProcedureB 的形参位置打 / 回退 / 粘一个同名的 username ——
            # 逐字符看，"别的过程里出现过 username"会被当成"这个名字真实存在"
            # 的铁证，于是弹窗把用户刚敲进去的字原样提示出来（提示自己）。
            # 只有【一模一样】才会触发：换成任何别处都没有的新名字，现场扫不到，
            # 回声防护照常把它剔掉 —— 正好对应用户描述的"不相同就不提示"。
            #
            # 模块级声明不受此限（proc_name 为空）：模块级名字本来就全模块可见，
            # 任何出现都是它的合法引用（v37 语义：`Sub test` 打全仍要提示）。
            scope_only = bool(ctx.get("proc_name") and low_word in decl_at_caret)

            def _is_echo_candidate(name):
                # 连续子串：前缀（结尾删）/ 后缀（开头删）/ 中段（头尾都删）
                # 都算回声。是否保留由下面的 _name_really_exists 用现场文本
                # 与声明证据裁决 —— 真名字（本模块别处出现 / 别的模块声明过）
                # 一律照常提示。
                return low_word in name.lower()

            echoes = [m for m in matches if _is_echo_candidate(m)]
            if echoes:
                try:
                    declared = self._declared_names()
                except Exception:
                    declared = set()
                # 现场证据【每次都要取】：判断"声明在本模块的名字是不是已被
                # 删掉"只能靠现场文本（声明集合最长滞后 0.4s）。代价是每次触发
                # 多读一次当前模块文本，远比全量重解析便宜。
                # scope_only：过程内声明（形参 / 局部 Dim）只看本过程与模块级
                # 的出现，别把别的过程的同名局部变量当证据（见上面 scope_only）。
                live = self._live_names_outside_caret(ctx, scope_only=scope_only)
                matches = [
                    m for m in matches
                    if not _is_echo_candidate(m)
                    or self._name_really_exists(m, ctx, live=live,
                                                declared=declared)
                ]
        # 注意：这里【不要】再加"唯一候选恰好等于所输词就收起"的收尾规则。
        #
        # 那条规则（v30 为修"回退键提示自己"引入）误伤了正常场景：变量 arr 已定义，
        # 输入 a / ar 都提示 arr，把 arr 打全反倒把列表关掉了 —— 用户明确要求
        # "输入完整，提示保留"，否则没法 Tab 确认、也看不出自己写对了没有。
        #
        # 它原本要防的"提示自己"，现在由上面那段「回声防护」更精确地覆盖：
        # 凡"以当前词开头"的候选，只要拿不出"真实存在"的证据（现场扫不到、
        # 别的模块也没声明过），就一律剔除 —— 既管"就是当前词本身"，
        # 也管"回退删字时残留的更长的旧版本"（v43）。
        # 真实存在的名字打全了就该继续提示，与打前缀行为一致。
        # 诊断：把"原始记录里有哪些同名命中、各属于哪个模块/过程"一并记下，
        # 这样日志能直接区分是"压根没收录"还是"收录了但被作用域过滤掉"。
        #
        # ⚠️ v67：这一段的开销与【候选池大小】成正比（它对全池再跑一遍
        # fuzzy_match，含正则分词）。池子 367 条时约 0.7ms，无所谓；v67 收进
        # 4500 多条宿主枚举常量后变成 ~18ms/次 —— 而它唯一的用途就是喂日志。
        # 所以只在开着日志（run_debug.bat）时才跑；日常使用完全不付这份开销。
        if _LOG_ENABLED:
            _log("trigger: line=%s col=%s word=%r proc=%r module=%r"
                 " visible=%d matches=%r"
                 % (ctx.get("line_no"), ctx.get("caret_col"), word,
                    ctx.get("proc_name"), ctx.get("module_name"),
                    len(visible_ids), matches))
            try:
                hits = [r for r in self.backend.get_identifiers()
                        if fuzzy_match(str(r[0]), word) is not None]
                _log("  raw hits for %r: %s" % (word, hits[:40]))
                _log("  hit positions: %s" % (
                    {k: v for k, v in list(self.match_hits.items())[:10]},))
            except Exception:
                pass
        if not matches:
            self.hide()
            return
        self.matches = matches
        # 按 Tab 即可直接确认当前已输入的完整名字；否则选中列表首项。
        exact = word.lower()
        sel = 0
        for idx, m in enumerate(matches):
            if m.lower() == exact:
                sel = idx
                break
        self.selected = sel
        # 换了输入词就回到列表开头，再按选中项的位置把窗口挪过去 ——
        # 打全名时选中项可能排在很后面（几十条候选里的第 20 个），
        # 不挪的话它在窗口外，用户看不见自己选了什么。
        self.view_top = 0
        self._scroll_to_selected()
        self.ctx = {
            "line_no": ctx["line_no"],
            "caret_col": ctx["caret_col"],
            "word": word,
            "word_start_col": start_col,
            # 替换范围右端：正常等于 caret_col；光标停在标识符【开头】或
            # 【中间】时大于它（要连光标右边那段残留名字一起换掉，
            # 见 extract_word_at），否则会在残留词前面插入候选、拼出双份名字。
            "word_end_col": end_col,
        }
        self.visible = True
        self.shown_at = time.time()
        self.word_is_complete = word.lower() in set(i.lower() for i in visible_ids)
        # 只把【当前这一屏】交给 UI 去画：行数最多 view_rows（默认 9）行，
        # 候选不足时还会更短，列表不会越滚越长；而完整的候选始终留在 self.matches
        # 里，滚动即可到达。
        self.ui.show(self.visible_rows(), self.view_selection(), self)

    def can_space_confirm(self):
        """
        空格键能否用于确认补全。

        只有当"光标前的词还不是完整定义的标识符"时才允许用空格补全，
        否则（例如已打完 ws，后面还有 wsName 这个候选）空格应当正常放行，
        避免把 'Set ws = ...' 误改成 'Set wsName = ...'。
        """
        return bool(self.visible and self.matches and not self.word_is_complete)

    def move(self, delta):
        if not self.visible or not self.matches:
            return
        n = len(self.matches)
        self.selected = (self.selected + delta) % n
        # 选中项跟着滚动走：走到窗口下沿时窗口下移一行（而不是整屏跳），
        # 首项按上键绕到末项时直接滚到底部。
        self._scroll_to_selected()
        self.ui.update_selection(self.selected)

    def accept(self):
        if not self.visible or not self.matches:
            return
        chosen = self.matches[self.selected % len(self.matches)]
        ctx = self.ctx
        try:
            end_col = ctx.get("word_end_col") or ctx.get("caret_col")
            self.backend.apply_completion(
                ctx["line_no"], ctx["word_start_col"], end_col, chosen
            )
        except Exception:
            # 插入失败也要保证状态复位，否则弹窗会卡住不再触发
            pass
        finally:
            # 进入静默期：后续排队中的 trigger 不会再把它弹回来
            self._suppress_until = time.time() + SUPPRESS_AFTER_ACCEPT
            self.hide()

    def pick(self, index):
        """按绝对下标直接选中并确认（鼠标点击 / 程序调用）。

        index 为 0-based，对应 `matches` 中的绝对位置（UI 在点击时会把"屏内行号"
        换算回绝对下标再调用，所以这里不需要知道窗口视图）。
        越界（列表没那么多项）返回 False 且不改变状态。
        """
        if not self.visible or not self.matches:
            return False
        if index < 0 or index >= len(self.matches):
            return False
        self.selected = index
        self.accept()
        return True

    def is_visible(self):
        return self.visible

    def current_matches(self):
        return list(self.matches)

    def hide(self):
        self.visible = False
        self.matches = []
        self.selected = 0
        self.view_top = 0
        self.ctx = None
        self.word_is_complete = False
        self.match_hits = {}
        self.ui.hide()

    def maybe_hide_on_outside_click(self, px, py):
        """鼠标点击了弹窗之外时收起。

        由 main.py 的全局鼠标监听在后台线程调用，经 action 队列转回主线程执行
        （popup 的几何查询只有主线程才安全）。点击弹窗内部由列表框绑定处理确认，
        这里仅当“落在弹窗外”才收起。
        """
        if not self.visible:
            return
        # 正在拖滚动条：直接返回，不收起（确保"选词前 UI 一直可见"）。
        # 加了 WS_EX_NOACTIVATE 后弹窗不会抢焦点，本路径本就不该触发；
        # 这里再兜一道保险，避免任何边界情况下拖拽途中把列表框收掉。
        if getattr(self.ui, "_drag", None) is not None:
            return
        try:
            inside = self.ui.contains_point(px, py)
        except Exception:
            inside = False
        if not inside:
            self.hide()
