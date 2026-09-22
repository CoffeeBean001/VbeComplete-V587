"""
VBA 标识符提取器。

从一段 VBA 模块源码文本中，提取出可用于自动补全的所有标识符：
  - 变量：Dim / Private / Public / Global / Static / Friend (+ ReDim)
  - 常量：Const
  - 过程：Sub / Function（含参数名）
  - 属性：Property Get / Let / Set（含参数名）
  - API 声明：Declare Sub / Function（含参数名）
  - 自定义类型与枚举：Type / Enum（**只收类型名**；块成员只能限定访问
    —— p.field / E.Member，不收录成裸名候选，见 in_block 分支）

纯文本解析，不依赖 Office，可在任何环境单测。

标识符记录为四元组：(name, module, proc, priv)
  - name   ：标识符名
  - module ：所在模块名；None 表示"未知/旧式"，引擎视为全局可见
  - proc   ：所属过程名；模块级标识符为 None
  - priv   ：模块级标识符是否"仅本模块可见"（Private）；
             过程内的局部变量/参数此字段无意义（恒 False）

可见性事实（依据微软官方文档，2026 复核）：
  - 模块级 Dim / Static 变量：默认等效 Private
  - 模块级 Const：默认 Private（类模块里永远只能 Private）
  - Sub / Function / Property / Declare：默认 Public
  - Type / Enum：默认 Public（标准模块与类模块均如此；类模块 Public Enum
    写入类型库，工程内全局可见）
  - 显式 Private -> 仅本模块；显式 Public / Global / Friend(工程内) -> 全局
"""

import bisect
import re

# 标识符（含中文）：首字符字母/下划线，其后字母/数字/下划线。Unicode 感知。
_IDENT = r"[^\W\d]\w*"

# VBA / Office 内置常量（vbTextCompare / xlUp / msoFileDialogFolderPicker ...）。
# 它们既没有声明也"只被读取"，用法扫描会把它们当隐式变量收进来 —— 纯噪音。
# 这类名字一律有固定的库前缀 + 大写首字母，按前缀排除足够安全。
_RE_BUILTIN_CONST = re.compile(r"^(?:vb|xl|mso|db|wd|pp|ol|ac|wpp)[A-Z]")

# 行继续符 `_`：VBA 语法要求它【前面至少有一个空格 / Tab】，且其后到行尾只有空白。
# 必须校验"前面是空白"——只写 `_\s*\n` 会把【以 _ 结尾的标识符】一起吞掉。
# 真实事故（v59）：用户 pub3 里 `Public Enum E` 的成员名是 A_ / B_ / ... / EZ_，
# 每行都以 "Z_" 结尾，于是 6 行全被并进下一行，`End Enum` 被一起吞掉 —— 之后
# 整个模块（206 行、25 个过程）都被当成 Enum 块成员解析：过程名收不全（只收
# 每行行首那个词），假名字 Function / If / With / s 混进候选池且 priv=False
# （全工程可见），过程内局部变量还被记成模块级，跨过程泄漏。
_RE_CONTINUATION = re.compile(r"(?<=[ \t])_[ \t]*\r?\n")  # 行继续符 _
_RE_DECL = re.compile(r"\b(?:Dim|Private|Public|Global|Friend|Static|ReDim)\b", re.I)
_RE_CONST = re.compile(r"\bConst\b", re.I)
_RE_SUB = re.compile(
    r"^\s*(?:(?:Public|Private|Friend|Global|Static)\s+)*(?:Sub|Function)\s+("
    + _IDENT + r")\s*(?:\(|$)", re.I)
_RE_PROP = re.compile(
    r"^\s*(?:(?:Public|Private|Friend|Global|Static)\s+)*Property\s+(?:Get|Let|Set)\s+("
    + _IDENT + r")\s*(?:\(|$)", re.I)
# 事件声明：Public Event Foo(ByVal x As Long)
_RE_EVENT = re.compile(
    r"^\s*(?:(?:Public|Private|Friend|Global|Static)\s+)*Event\s+(" + _IDENT
    + r")\s*(?:\(|$)", re.I)
# PtrSafe 是 64 位 Office 下 Declare 的修饰关键字，需一并容忍。
# 前面的访问修饰符（Public / Private / Friend / Global）也必须吃进来：本正则
# 在 extract_records 里是用 .match() 从头匹配的，`Private Declare PtrSafe
# Function gApi Lib "k" ...` 行首是 Private，旧写法 \bDeclare 匹配不上，会掉进
# 下面的"变量声明"分支 —— 收录出假名字 "Declare"，真正的 gApi2 反而丢了（v59）。
_RE_DECLARE = re.compile(
    r"\s*(?:(?:Public|Private|Friend|Global)\s+)*\bDeclare\s+"
    r"(?:PtrSafe\s+)?(?:Sub|Function)\s+(" + _IDENT + r")\s*", re.I)
_RE_TYPE = re.compile(
    r"^\s*(?:(?:Public|Private|Friend|Global)\s+)*Type\s+(" + _IDENT + r")\b", re.I)
_RE_ENUM = re.compile(
    r"^\s*(?:(?:Public|Private|Friend|Global)\s+)*Enum\s+(" + _IDENT + r")\b", re.I)
_RE_LEADING_IDENT = re.compile(_IDENT)

# 过程结束 / 块结束
_RE_PROC_END = re.compile(r"\bEnd\s+(?:Sub|Function|Property)\b", re.I)
_RE_BLOCK_END = re.compile(r"\bEnd\s+(?:Type|Enum)\b", re.I)

# 用于 proc_at_line：从某行往上找所属过程
_RE_ANY_PROC_END = re.compile(r"^\s*End\s+(?:Sub|Function|Property|Type|Enum)\b", re.I)

# 行首修饰关键字检测（在已屏蔽字符串/注释的文本上使用）
_RE_LEAD_PRIVATE = re.compile(r"^\s*(?:Private)\b", re.I)
_RE_LEAD_PUBLIC = re.compile(r"^\s*(?:Public|Global|Friend)\b", re.I)

# 关键词 -> 分类（决定"默认可见性"）
KIND_PROC = "proc"      # Sub/Function/Property/Declare：标准模块默认 Public
KIND_VAR = "var"        # 模块级 Dim/Static 变量：默认 Private
KIND_CONST = "const"    # Const：默认 Private
KIND_TYPE = "type"      # Type / Enum：默认 Public（两种模块都是）
KIND_MEMBER = "member"  # Type/Enum 成员：只用于分类，当前不收录（见 in_block）
KIND_LOCAL = "local"    # 过程内局部变量/参数：仅本过程可见


def _is_private(kind, has_priv_kw, has_pub_kw, is_std_module):
    """判定一条模块级声明的可见性是否为"仅本模块"（priv=True）。"""
    if kind in (KIND_PROC, KIND_VAR, KIND_CONST):
        if has_priv_kw:
            return True
        if has_pub_kw:
            # Public 变量/常量/过程：标准模块 -> 工程可见；类/文档模块
            # 的成员只能通过实例/限定名访问，跨模块不可裸写 -> 视为模块私有
            return not is_std_module
        if kind == KIND_PROC:
            # Sub/Function/Property/Declare 不带关键字默认 Public
            # （类/文档模块里的过程是成员，跨模块不能裸调 -> 私有）
            return not is_std_module
        return True   # Dim / Static / Const 默认 Private
    if kind == KIND_TYPE:
        # Type / Enum 默认 Public（含类模块）；显式 Private 则模块私有
        return has_priv_kw
    return False      # member/local：priv 字段无意义，调用处另定


def scan_code_states(code):
    """逐字符标注：每个下标处的字符是否落在【字符串 / 注释】里。

    返回一个与 `code` **等长**的 bool 列表（True = 该下标在字符串或注释里）。

    ★v93：这是这套扫描规则的【唯一实现】—— `_mask_strings_and_comments`
    （要一份抹白后的文本去跑正则）与 `engine._in_comment_or_string`（要问
    "光标左边那个字符在不在字符串/注释里"）都从它派生。

    为什么必须抽出来：`_mask_strings_and_comments` 的产物是**逐 token 一格**
    的空格（一整个字符串 / 一整行注释都只落一个空格），**不是逐字符**等长，
    因此调用方【不能】拿它的下标去对应原文的下标。engine 那边一开始就是这么
    误用的（`masked[n-1]` 直接 IndexError）。规则本身又很细（`""` 转义、
    `'` 到行尾），写第二份必然再次分叉 —— 所以只能有一份。

    规则（与 VBE 一致的常见约定，不含行继续符等跨行情形）：
      * `"` 开启字符串，直到配对的 `"`；串内 `""` 是一个转义的引号，跳过；
      * `'` 开启注释，直到行尾（含 `\\n` 之前）；
      * 其余字符都不在里面。
    刻意**不认 `Rem`** —— 它需要"行首 / 语句边界"这条额外的语法前提，
    那个前提属于调用方的判据（见 engine._member_list_dot_col 的守卫）。
    """
    n = len(code)
    states = [False] * n
    i = 0
    while i < n:
        c = code[i]
        if c == '"':
            states[i] = True
            i += 1
            while i < n:
                states[i] = True
                if code[i] == '"':
                    if i + 1 < n and code[i + 1] == '"':  # 转义引号 ""
                        states[i + 1] = True
                        i += 2
                        continue
                    i += 1
                    break
                i += 1
            continue
        if c == "'":  # 注释到行尾
            while i < n and code[i] != "\n":
                states[i] = True
                i += 1
            continue
        i += 1
    return states


def _mask_strings_and_comments(code):
    """把字符串字面量和注释替换为空格，避免误匹配。

    ⚠️ 产物长度【不等】于原文本：一整个字符串 / 一整行注释只落一个空格
    （逐 token，不是逐字符）。需要按下标对齐的场合请用 scan_code_states。
    """
    states = scan_code_states(code)
    out = []
    i, n = 0, len(code)
    while i < n:
        if states[i]:
            # 一整段连续的字符串 / 注释 -> 合成一个空格
            while i < n and states[i]:
                i += 1
            out.append(" ")
            continue
        out.append(code[i])
        i += 1
    return "".join(out)


# 出现在【变量名之前】的修饰关键字：`Public WithEvents clk As MSForms.CommandButton`
# 里的 WithEvents 必须跳过，否则收录到的是 "WithEvents" 这个假名字，真正的 clk
# 反而丢了（v59 真实事故：cls 模块的 Public WithEvents clk）。
_RE_LEAD_MODIFIER = re.compile(r"^(?:WithEvents)\s+", re.I)


def _decl_names(rest):
    """从声明右侧（如 'x As Long, y(), z As String' 或 'a = 5, b = 6'）提取标识符名。"""
    names = []
    for seg in rest.split(","):
        seg = seg.strip()
        # 跳过 WithEvents 这类"名字之前"的修饰符（逐段跳过，逗号后面那段不受影响）
        seg = _RE_LEAD_MODIFIER.sub("", seg, count=1).strip()
        m = _RE_LEADING_IDENT.match(seg)
        if m:
            names.append(m.group(0))
    return names


def _params_inside(paren_text):
    """从括号内文本提取参数名：ByVal x As Integer, ByRef y As String, Optional z。

    跳过 ByVal/ByRef/Optional/ParamArray 等修饰符，取真正的参数名。
    """
    names = []
    for part in paren_text.split(","):
        part = part.strip()
        if not part:
            continue
        # 形如 ByVal ws As Worksheet / Optional retries As Long = 3
        m = re.match(r"(?:ByVal|ByRef|Optional|ParamArray)\s+(" + _IDENT + r")", part, re.I)
        if m:
            names.append(m.group(1))
            continue
        m = _RE_LEADING_IDENT.match(part)
        if m:
            names.append(m.group(0))
    return names


def extract_records(code, module=None, is_std_module=True, caret=None):
    """
    返回 [(name, module, proc, priv), ...]。

    module 为 None 时仍可调用（用于单元测试/旧逻辑），priv 会按默认规则计算，
    但引擎侧把 module=None 的记录当作"旧式全局可见"处理。

    caret=(line_no, col)：光标位置。解析层【不】据此排除任何名字 —— 名字一律
    收录，交给引擎按"正在输入的词"决定是否剔除（ctx["decl_names"] 精确剔除光标
    处那一个词、前缀照常提示；幻影词由回声防护剔除）。见下面那段注释与
    _blank_ident_at 的 v79e 说明。
    """
    masked = _mask_strings_and_comments(code)
    # 合并行继续符
    masked = _RE_CONTINUATION.sub(" ", masked)

    # 注意：早先这里会按 caret 行把"正在声明的名字"从候选里剔除（连前缀一起）。
    # 但这会把合法的【前缀补全】也误杀：在 numArr 的声明行上输入 num，numArr
    # 就再也提示不出来（bug：输入 num 只能提醒 num、不提醒 numArr）。正确的语义是
    # "精确匹配自己才隐藏、前缀照常提示"，而这由引擎层 in_decl_position 的精确
    # 自我剔除负责（输入完整 numArr 才隐藏，输入前缀 num 仍正常提示 numArr）。
    # 因此解析层不做任何 caret 排除，名字一律收录，交给引擎按"正在输入的词"
    # 决定是否剔除——这样既修好了前缀补全，又没丢"定义变量时不提示自己"。
    # （v79d 曾把光标压着的声明名抹成下划线来"顺手"排除掉它，反而破坏了行结构，
    #   见 _blank_ident_at 的 v79e 说明；那件事本来就该由引擎那两道防线做。）

    result = []
    seen = set()
    cur_proc = None      # 当前所在过程【的作用域键】（None = 模块声明区）
    in_block = None      # 'Type' / 'Enum'

    # v79e：同名过程会让"过程名"失去区分度（作用域过滤就是拿它比对的），
    # 因此按出现次序生成作用域键：第 1 次用原名，第 2 次起带 `#序号`。
    _lines = masked.split("\n")
    _occ = {}            # 小写过程名 -> 已出现的次数

    # v79e：光标正压在"这条声明正在声明的名字"上时，文本【一个字都不动】
    # （见 _blank_ident_at）—— 动它会让这一行不再是声明行，紧随其后的整个过程体
    # 失去过程归属，局部/隐式变量降级成模块级、全模块到处可见。这个"打什么提示
    # 什么"的防护由引擎那两道防线负责，解析层不掺和。

    def add(name, proc, priv):
        if not name:
            return
        key = (name.lower(), (proc or "").lower(), (module or "").lower())
        if key not in seen:
            seen.add(key)
            result.append((name, module, proc, priv))

    def proc_key(name):
        """刚看到的过程头名字 -> 作用域键（第 2 次同名起带序号）。"""
        k = str(name).lower()
        _occ[k] = _occ.get(k, 0) + 1
        return _proc_scope_key(name, _occ[k])

    for raw in _lines:
        line = raw.strip()
        if not line:
            continue

        # 过程结束：End Sub / End Function / End Property
        if _RE_PROC_END.match(line):
            cur_proc = None
            in_block = None
            continue

        # 块结构：Type / Enum
        m_type = _RE_TYPE.match(line)
        m_enum = _RE_ENUM.match(line)
        if m_type or m_enum:
            kind = KIND_TYPE
            in_block = "Type" if m_type else "Enum"
            priv = _is_private(kind, bool(_RE_LEAD_PRIVATE.match(line)),
                               bool(_RE_LEAD_PUBLIC.match(line)), is_std_module)
            name = m_type.group(1) if m_type else m_enum.group(1)
            add(name, None, priv)          # 类型/枚举名
            continue
        if _RE_BLOCK_END.match(line):
            in_block = None
            continue

        if in_block:
            # 块成员（Type 字段 / Enum 成员）【不收录】——v59 按用户要求改：
            #   * Type 字段只能通过变量限定访问（p.field），裸名根本不可引用，
            #     提示出来毫无意义；
            #   * Enum 成员虽然 VBA 允许裸名引用，但既然枚举名本身已经限定了
            #     （E.A_），列表里再倒出一堆成员名只会把候选池搞吵。
            # 点号之后的成员列表交给 VBE 自带的「自动列出成员」（v55 我们让位）。
            # 注意这里仍然要 continue：块内的行不能被当成模块级变量声明。
            # （extract_implicit_records 早就是这么做的，两条路径现在一致。）
            continue

        has_priv_kw = bool(_RE_LEAD_PRIVATE.match(line))
        has_pub_kw = bool(_RE_LEAD_PUBLIC.match(line))

        # Declare 声明（API）——一定在模块级
        m_decl = _RE_DECLARE.match(line)
        if m_decl:
            priv = _is_private(KIND_PROC, has_priv_kw, has_pub_kw, is_std_module)
            add(m_decl.group(1), None, priv)
            pm = re.search(r"\((.*)\)", line)
            if pm:
                for p in _params_inside(pm.group(1)):
                    add(p, cur_proc, False)
            continue

        # Sub / Function：过程名模块级，参数属于该过程
        m_sub = _RE_SUB.match(line)
        if m_sub:
            proc = m_sub.group(1)
            priv = _is_private(KIND_PROC, has_priv_kw, has_pub_kw, is_std_module)
            add(proc, None, priv)
            cur_proc = proc_key(proc)      # 第 2 次同名起带 `#序号`
            pm = re.search(r"\((.*)\)", line)
            if pm:
                for p in _params_inside(pm.group(1)):
                    add(p, cur_proc, False)
            continue

        # Property Get/Let/Set
        m_prop = _RE_PROP.match(line)
        if m_prop:
            proc = m_prop.group(1)
            priv = _is_private(KIND_PROC, has_priv_kw, has_pub_kw, is_std_module)
            add(proc, None, priv)
            cur_proc = proc_key(proc)
            pm = re.search(r"\((.*)\)", line)
            if pm:
                for p in _params_inside(pm.group(1)):
                    add(p, cur_proc, False)
            continue

        # Const 常量：作用域取决于声明位置
        if _RE_CONST.search(line):
            kind = KIND_CONST
            priv = _is_private(kind, has_priv_kw, has_pub_kw, is_std_module)
            body = _RE_CONST.sub("", line, count=1).strip()
            body = re.sub(r"^(?:Private|Public|Global|Friend)\b", "", body, flags=re.I).strip()
            for nm in _decl_names(body):
                # 过程内 Const 永远是局部
                if cur_proc:
                    add(nm, cur_proc, False)
                else:
                    add(nm, None, priv)
            continue

        # 普通变量声明 Dim/Private/...（非 Const）
        if _RE_DECL.search(line):
            kind = KIND_VAR
            priv = _is_private(kind, has_priv_kw, has_pub_kw, is_std_module)
            body = _RE_DECL.sub("", line, count=1).strip()
            for nm in _decl_names(body):
                # 过程内 Dim/Static = 局部
                if cur_proc:
                    add(nm, cur_proc, False)
                else:
                    add(nm, None, priv)
            continue

    return result


# ---------------------------------------------------------------------------
# 隐式变量（不写 Option Explicit 时"用到即存在"的变量）
#
# 这类变量没有任何 Dim/Const 声明，只在代码里被赋值或被 For 遍历时"诞生"。
# 因此从「赋值目标」与「For/For Each 循环变量」反推它们：
#     x = 1            Set ws = ...        Let y = ...
#     arr(0) = 1       For i = 1 To 10     For Each c In Cells
# 只有真正被写入过得名字才算——仅出现在表达式里（如 MsgBox 用户名）的名字
# 不会收录，避免把一堆内置符号当成候选造成噪音。
# ---------------------------------------------------------------------------

# 不作为隐式变量收录的关键字 / 内置对象与函数名（小写比较）
_IMPLICIT_STOPWORDS = frozenset("""
abs and array asc atn beep boolean byval byref call case cbool cbyte ccur cdate cdbl cdec
choose chr cint clng collection const cos createobject csng cstr curdir currency date dateadd
datediff datepart dateserial datevalue day ddb debug declare deftype dim dir do doevents double
each else elseif end enum erase err error event exit exp false filecopy fileattr filelen
filedatetime fix for format freefile friend function fv get getobject global gosub goto hex
hour iif if imp implements in input instr instrrev int integer ipmt irr is isarray isdate
isempty iserror ismissing isnull isnumeric isobject join kill lbound lcase left len let
lineinput like load loc lock log long loop lset ltrim me mid minute mirr mkdir mod month msgbox
name new next not now nper null object oct on open option or pm pmt ppmt preserve print private
property public put pv raiseevent randomize rate redim rem remove reset resume return rgb right
rmdir rnd rset rtrim savepicture seek select set sgn shell sin single sln space spc sqr static
stop str strcomp string sub switch tab tan then time timer timeserial timevalue to trim true
type typeof ubound ucase unload until val variant vartype weekday while with withevents write
xor year application worksheetfunction activeworkbook activesheet activecell selection range
cells worksheets workbooks sheets charts chartobjects thisworkbook rows columns activewindow
explicit compare binary database module attribute
vb_name vb_creatable vb_predeclaredid vb_exposed vb_description vb_helpid vb_varhelpid
vb_user_memid vb_procdata
round split replace inputbox empty nothing strreverse environ cverr qbcolor filter typename
second weekdayname monthname strconv formatcurrency formatdatetime formatnumber formatpercent
byte longlong longptr decimal object worksheet workbook chart
as step optional paramarray wend lib alias ptrsafe vba6 vba7 win32 win64 mac text
""".split())

# 整行跳过：声明行与不可能产生隐式变量的语句
_RE_SKIP_LINE = re.compile(
    r"^\s*(?:#|Rem\b|Option\b|Def[A-Za-z]*\b|Implements\b|Sub\b|Function\b|Property\b|"
    r"Type\b|Enum\b|Declare\b|End\b|With\b|On\b|Exit\b|GoSub\b|GoTo\b|Resume\b|"
    r"Case\b|Select\b|Do\b|Loop\b|Next\b|Wend\b|While\b|Erase\b|Open\b|Close\b|"
    r"Print\b|Write\b|Input\b|Get\b|Put\b|Seek\b|Lock\b|Unlock\b|Name\b|Kill\b|"
    r"MkDir\b|RmDir\b|ChDir\b|ChDrive\b|FileCopy\b|RaiseEvent\b|Stop\b)", re.I)

# If x = 1 Then y = 2 / ElseIf ... Then ... / Else y = 3 里的赋值也要能抓到
_RE_IF_THEN = re.compile(r"^\s*(?:Else)?If\b.*?\bThen\b(.*)$", re.I | re.S)
_RE_ELSE_ONLY = re.compile(r"^\s*Else\b(.*)$", re.I)

_RE_IMP_ASSIGN = re.compile(
    r"^\s*(?:Let\s+|Set\s+)?(" + _IDENT + r")\s*(?:\([^()]*\))?\s*=(?!=)", re.I)
_RE_IMP_FOR_EACH = re.compile(r"^\s*For\s+Each\s+(" + _IDENT + r")\s+In\b", re.I)
_RE_IMP_FOR = re.compile(r"^\s*For\s+(" + _IDENT + r")\s*=", re.I)


def _implicit_names_in_stmt(stmt):
    """从一条语句里提取"被写入"的隐式变量名（可能为空）。"""
    s = stmt.strip()
    if not s:
        return []
    m = _RE_IMP_FOR_EACH.match(s) or _RE_IMP_FOR.match(s) or _RE_IMP_ASSIGN.match(s)
    if not m:
        return []
    name = m.group(1)
    if name.lower() in _IMPLICIT_STOPWORDS:
        return []
    return [name]


# 这些前导词后面的名字不是变量（类型名 / 库名 / 别名 / 跳转标签等）
_RE_USAGE_SKIP_LINE = re.compile(r"^\s*(?:#|Declare\b|Attribute\b|Option\b|Implements\b)", re.I)

_PREV_WORD_SKIP = frozenset(
    "as new lib alias implements goto gosub then else byval byref optional paramarray".split())


def _blank_type_enum_blocks(masked_code):
    """把 Type / Enum 块【内部的成员行】整行抹成等长空格（v61）。

    为什么需要：块成员（枚举成员 / 自定义类型字段）只能通过 `E.Member` /
    `p.field` 限定访问，裸名根本不可引用，因此不该进候选池（v59 用户拍板）。
    extract_records 那时起就不收它们了，但"只读用法"扫描（_usage_candidates）
    是【整篇文本】扫标识符的、没有块的概念 —— 于是 `Public Enum E` 里每行的
    成员名（A_ / AA_ / …）又被当成"用到了却没声明"的隐式变量收回来，v59 的
    修复等于没落地。真机实测：工程里那个模块又冒出 A_ / AA_ / AZ_ / BA_ /
    BZ_ / CA_ / CZ_ / DA_ / DZ_ / EA_ / EZ_ / Z_ 共 12 个。

    Type / Enum 声明行与 End 行【保留】：前者由 extract_records 负责收类型名，
    后者让 extract_implicit_records 里的 in_block 状态能正常复位（否则块之后
    的正常代码会被整段当成"块内"跳过）；End / Enum 这类关键字本来就在
    _IMPLICIT_STOPWORDS 里，不会被误收。

    抹成【等长空格】而不是删行：_usage_candidates 返回的是字符偏移，调用方拿它
    配 _line_proc_map 定位过程归属，删行/改长度会让后面所有偏移整体错位。
    """
    out = []
    in_block = False
    for raw in masked_code.split("\n"):
        line = raw.strip()
        if in_block:
            if _RE_BLOCK_END.match(line):
                in_block = False
                out.append(raw)                 # End Type / End Enum：保留
            else:
                out.append(" " * len(raw))      # 成员行：等长空格
            continue
        if _RE_TYPE.match(line) or _RE_ENUM.match(line):
            in_block = True
        out.append(raw)
    return "\n".join(out)


def _usage_candidates(masked_code, excluded):
    """扫描"只被读取、没被写过"的隐式变量，返回 [(name, 字符偏移), ...]。

    带偏移是为了让调用方能按"首次出现所在过程"给它们限定作用域 —— 一律记成
    模块级会让 A 过程里读到的名字泄漏到 B 过程（表现为"一个函数里提示出
    别的函数的变量"）。

    VBA 里凡是不是关键字/内置符号、又没有任何声明的裸标识符，就是隐式变量。
    因此这里把整段代码里所有"解释不掉"的裸标识符都收进来，再排除：
      - 关键字/内置符号（_IMPLICIT_STOPWORDS）
      - 已声明或已被赋值判定收录的名字（excluded）
      - 成员访问 `.xxx`、调用/下标 `xxx(`、命名参数 `xxx:`
      - As / New / Lib / Alias 等前导词后面的类型名与限定名
    """
    found = []
    seen = set()
    # 编译指令 / API 声明 / 模块属性行不参与用法扫描（否则函数名、库名会被当变量）。
    #
    # 【关键】必须抹成【等长空格】而不是空串 —— 本函数返回的偏移是与
    # _line_proc_map 算出的 offsets 配对使用的（调用方靠它定位"这个名字属于
    # 哪个过程"）。早期实现把整行替换成 ""，使后续所有字符偏移整体前移；
    # 真实模块头部动辄有 5 行以上 `Attribute VB_xxx` + `Option Explicit`
    # （累计可前移 150+ 字符），于是偏移与 offsets 彻底错位 ——
    # 只读用法扫出来的隐式变量会被算到【别的过程】甚至【模块级】，
    # 表现为"一个过程里的变量泄漏到另一个过程（含其形参位置）"。
    # 抹成空格后长度不变，偏移始终与 offsets 对齐。
    kept_parts = []
    for l in masked_code.split("\n"):
        if _RE_USAGE_SKIP_LINE.match(l):
            kept_parts.append(" " * len(l))     # 等长空格：屏蔽内容、保留偏移
        else:
            kept_parts.append(l)
    masked_code = "\n".join(kept_parts)
    n = len(masked_code)
    for m in re.finditer(_IDENT, masked_code):
        name = m.group(0)
        low = name.lower()
        if low in excluded or low in _IMPLICIT_STOPWORDS or low in seen:
            continue
        if _RE_BUILTIN_CONST.match(name):
            continue                       # vbTextCompare / xlUp / msoXxx ...
        s, e = m.start(), m.end()
        j = s - 1
        while j >= 0 and masked_code[j] in " \t":
            j -= 1
        if j >= 0 and masked_code[j] in ".!":
            continue                       # 成员访问
        k = e
        while k < n and masked_code[k] in " \t":
            k += 1
        if k < n and masked_code[k] in "(:":
            continue                       # 调用 / 下标 / 命名参数 / 标签
        if j >= 0:
            w_end = j + 1
            w_start = w_end
            while w_start > 0 and (masked_code[w_start - 1].isalnum()
                                   or masked_code[w_start - 1] == "_"):
                w_start -= 1
            if masked_code[w_start:w_end].lower() in _PREV_WORD_SKIP:
                continue                   # As Long / New Collection / ...
        seen.add(low)
        found.append((name, m.start()))
    return found


# 行继续符（行尾的 ` _`）
_RE_CONT_AT_END = re.compile(r"[ \t]_\s*$")

# 声明行前导修饰词：Public / Private / Global / Friend / Static / Dim / Const /
# WithEvents。判定"这是不是一条变量声明"之前要先剥掉它们，否则
# `Public Sub Foo()` 会因为含 Public 而被误判成变量声明行。
_RE_LEAD_MODIFIER = re.compile(
    r"^\s*(?:Public|Private|Global|Friend|Static|Dim|Const|WithEvents)\s+", re.I)
_RE_LEAD_MODIFIER_STRIP = re.compile(
    r"^\s*(?:Public|Private|Global|Friend|Static|Dim|Const|WithEvents)\b", re.I)

# 剥掉修饰词后若以此开头，说明这条语句【不是】变量/常量声明：
#   Sub/Function/Property/Type/Enum/Declare/Event —— 过程、类型、API 声明；
#   ReDim —— 只是重新分配已有数组，不产生新名字，反而需要提示已有数组名；
#   With/End/Option/Implements/Attribute/Rem/# —— 其它语句与编译指令。
_RE_NOT_VAR_DECL = re.compile(
    r"^\s*(?:Sub|Function|Property|Type|Enum|Declare|Event|With|End|Option|"
    r"Implements|Attribute|ReDim|Rem\b|#|Def[A-Za-z]*\b)", re.I)



def _logical_line_at(lines, idx):
    """返回包含第 idx 行（0-based）的、合并了行继续符后的【逻辑行】文本。

    `Public gA As Long, _` / `       gB As Long` 这种跨行声明，光标落在续行上时
    该行本身不含 Dim/Public，若只看物理行就识别不出"正在声明"，自我提示因此漏网。
    """
    start = idx
    while start > 0 and _RE_CONT_AT_END.search(lines[start - 1].rstrip("\r")):
        start -= 1
    end = idx
    last = len(lines) - 1
    while end < last and _RE_CONT_AT_END.search(lines[end].rstrip("\r")):
        end += 1
    return " ".join(_RE_CONT_AT_END.sub("", lines[i].rstrip("\r"))
                    for i in range(start, end + 1))


def is_continuation_line(line_text):
    """这一行是否以行继续符 `_` 结尾（因此【下一行】是它的续行）。

    v53：形参常常独占一行（长签名被 VBE 拆成 `Sub Foo( _` + 缩进的形参行）。
    续行本身没有 Sub/Dim 之类的声明关键字，is_caret_in_declaration() 看不出来，
    但用户在那里打 / 退格 / 粘贴形参名与在首行上完全是一回事。后端据此把续行
    也算作"声明行"，把该逻辑行正在声明的名字（含形参）交给引擎。

    只看一行文本，不做任何跨行读取——调用方（vbe_bridge.get_context）只需
    额外读一行就能判定。
    """
    return bool(_RE_CONT_AT_END.search((line_text or "").rstrip("\r")))


def _inside_type_block(lines, idx):
    """第 idx 行（0-based）是否位于 Type / Enum 块内部（用于识别成员行）。"""
    depth = 0
    for i in range(idx):
        l = lines[i].strip()
        if _RE_TYPE.match(l) or _RE_ENUM.match(l):
            depth += 1
        elif _RE_BLOCK_END.match(l):
            depth = max(0, depth - 1)
    return depth > 0


def _decl_line_names(line):
    """若 line 是【任何声明语句】，返回其中正在被声明的名字集合（小写）。

    覆盖：变量/常量（Dim/Static/Const/Public|Private|Global ...）、
    过程（Sub/Function/Property）、API（Declare）、事件（Event）、
    自定义类型与枚举（Type/Enum）。参数名也算——它们同样是"正在起的新名字"。

    注意：函数名/过程名/类型名同样必须排除。否则 `Public Function gCalc()`
    里正在输入的 gCalc 会被自己的声明记录命中，表现为"输入函数名提示函数名"。
    """
    if not line or not line.strip():
        return None
    # 1) 过程 / 属性 / 事件 / API 声明
    m = _RE_SUB.match(line) or _RE_PROP.match(line) or _RE_EVENT.match(line)
    if m:
        return _proc_decl_names(line, m.group(1))
    m = _RE_DECLARE.search(line)
    if m:
        return _proc_decl_names(line, m.group(1))
    # 2) Type / Enum 声明行
    m = _RE_TYPE.match(line) or _RE_ENUM.match(line)
    if m:
        return {m.group(1).lower()}
    # 3) 变量 / 常量声明
    s = line
    for _ in range(4):                      # 连续修饰词（Public Static x ...）
        m2 = _RE_LEAD_MODIFIER.match(s)
        if not m2:
            break
        s = s[m2.end():]
    if _RE_NOT_VAR_DECL.match(s):
        return None                          # ReDim / With / End / Option ... 不算
    if _RE_CONST.search(line):
        body = _RE_CONST.sub("", line, count=1)
    elif _RE_DECL.search(line):
        body = _RE_DECL.sub("", line, count=1)
    else:
        return None
    body = _RE_LEAD_MODIFIER_STRIP.sub("", body)
    names = set()
    for seg in body.split(","):
        mm = _RE_LEADING_IDENT.match(seg.strip())
        if mm:
            names.add(mm.group(0).lower())
    return names or None


def _proc_decl_names(line, proc_name):
    """过程/事件/API 声明行：过程名 + 全部参数名。"""
    names = {str(proc_name).lower()}
    pm = re.search(r"\((.*)\)", line)
    if pm:
        for p in _params_inside(pm.group(1)):
            names.add(p.lower())
    return names


def decl_names_at_caret(code, caret):
    """光标若位于【声明行】上，返回该行正在声明的名字集合（小写）。

    用于兜底剔除「定义变量/函数时提示其本身」的自我提示（第二道防线）。
    第一道防线是 is_caret_in_declaration() —— 它让弹窗在声明行上直接静默；
    这里负责它覆盖不到的情形：续行声明的第二行、Type/Enum 成员行等
    "本行没有声明关键字"的位置。

    注意不能只排除"光标处那一个词"——用户输入 `g` 时光标处的词是 `g`，
    而声明的变量是 `gCounter`，两者不等，仍会被收录并自我提示。
    因此整行（逻辑行，含续行）声明的名字一并排除。

    必须在**原始文本**上判断：_mask_strings_and_comments 与续行合并都会
    改变长度/行号，用它们之后的行号会错位（同 v20 的教训）。
    """
    if not caret:
        return None
    try:
        line_no = int(caret[0])
    except Exception:
        return None
    lines = code.split("\n")
    if not (1 <= line_no <= len(lines)):
        return None
    idx = line_no - 1
    names = _decl_line_names(_logical_line_at(lines, idx))
    if names:
        return names
    # Type / Enum 成员行：`    x As Long`，本行没有任何声明关键字，
    # 只能靠向上回溯确认正处在 Type/Enum 块里。
    stripped = _mask_strings_and_comments(lines[idx]).strip()
    m = _RE_LEADING_IDENT.match(stripped)
    if (m and m.group(0).lower() not in ("end", "type", "enum")
            and _inside_type_block(lines, idx)):
        return {m.group(0).lower()}
    return None


def is_caret_in_declaration(line_text, caret_col=None):
    """光标是否处在"正在起新名字"的声明语句上（第一道防线）。

    命中即【彻底不弹窗】。理由：这些位置上用户是在**造名字**而不是**引用
    名字**，任何候选都是噪音；更要命的是正在输入的那个词本身往往已经在
    标识符池里（它在别处被读过、或在别的模块里有同名 Public 声明），
    于是出现"输入函数名/公共变量名时提示出它自己"。

    覆盖：Dim / Static / Const / Public|Private|Global 变量与常量、
    Sub / Function / Property / Declare / Event / Type / Enum。
    不覆盖：ReDim（那里要提示已有的数组名）、普通语句、Type 成员行
    （成员行由 decl_names_at_caret 兜底）。

    只依赖光标所在行，不需要整个模块文本，因此可在 get_context 里廉价调用。
    """
    if not line_text or not line_text.strip():
        return False
    line = _mask_strings_and_comments(line_text).strip()
    if not line:
        return False
    # 1) 过程 / 属性 / 事件 / API
    if _RE_SUB.match(line) or _RE_PROP.match(line) or _RE_EVENT.match(line):
        return True
    if _RE_DECLARE.search(line):
        return True
    # 2) Type / Enum 声明行
    if _RE_TYPE.match(line) or _RE_ENUM.match(line):
        return True
    # 3) 变量 / 常量声明
    s = line
    for _ in range(4):
        m = _RE_LEAD_MODIFIER.match(s)
        if not m:
            break
        s = s[m.end():]
    if _RE_NOT_VAR_DECL.match(s):
        return False
    return bool(_RE_CONST.search(line) or _RE_DECL.search(line))


def _caret_decl_names(code, caret):
    """decl_names_at_caret 的旧名（保留兼容）。"""
    return decl_names_at_caret(code, caret)


def is_caret_in_type_position(line_text, caret_col):
    """光标是否落在声明（`Dim`/`Const`/`Declare`/`Sub`/`Function`/`Property`/
    `Type` 成员等）里 `As` 关键字之后的【数据类型名】位置。

    例如 `Dim x As |`（| 为光标）正在输入类型名，此刻不应提示任何变量名。
    判定依据：在「光标所在的声明片段」内、光标之前是否出现过 `As` 关键字。

    声明片段 = 从光标往前，跳过括号内的内容、以最近的顶层逗号（同
    `Dim a As Long, b As Integer` 的分隔）为界切出的一段：
      - `Dim a As Long, b As |` 命中第二个片段 `b As ` -> 命中（类型位置）；
      - `Dim x` 光标在变量名上、片段 `x` 不含 As -> 不命中（照常提示新变量名）；
      - `Dim arr(0 To 9) As In` 片段为整行（括号内无顶层逗号）含 As -> 命中。

    VBA 里 `As` 只出现在类型声明的语义中（参数类型、返回值类型、变量/常量/成员
    类型、`Open ... As #n` 的文件号等），因此「光标前有 As」即可稳健地判定为
    「正在填类型名」，无需逐一枚举声明关键字。先屏蔽字符串/注释再扫描，避免
    字符串字面量里出现的 "As" 造成误判。
    """
    if not line_text:
        return False
    if not caret_col or caret_col <= 1:
        return False
    # 先屏蔽字符串/注释再判断，避免 `"As ..."` 里的 As 误判。
    # ⚠️ 必须用【逐字符状态】再拼前缀，不能拿 _mask_strings_and_comments 的
    # 产物去切片 —— 它是逐 token 一个空格、长度与原文不等（v93 踩过）。
    _states = scan_code_states(line_text)
    end = caret_col - 1
    if end < 0:
        end = 0
    if end > len(line_text):
        end = len(line_text)
    prefix = "".join(" " if _states[i] else line_text[i] for i in range(end))
    # ★v93：`New` 也是类型位置（`Set c = New |` / `Set c = New Coll|`，VBE
    # 在这儿弹类型列表）。原实现只认 `As`，于是 `Set c = New ` 处
    # in_type_position=False —— 我们不再把候选收窄到"类型名"，一堆变量名 /
    # 过程名冒出来（正是用户报的"填类型名时提示一堆变量名"那类噪音）。
    #
    # 判据与 engine._type_list_kw_end 同口径：先跳过【紧贴光标】的那个词
    # （用户正在打的类型名），再跳空白，然后要求收尾是整词 `New`。
    _q = prefix
    if _q and (_q[-1].isalnum() or _q[-1] == "_"):
        while _q and (_q[-1].isalnum() or _q[-1] == "_"):
            _q = _q[:-1]
    _q = _q.rstrip(" \t")
    if re.search(r"\bNew$", _q, re.I):
        return True
    # 括号深度跟踪：从光标往前，遇到顶层逗号即当前声明片段起点。
    # 顶层逗号 = 不在任何括号里（数组维度 / 过程参数 / 下标）的逗号。
    depth = 0
    seg_start = 0
    i = len(prefix) - 1
    while i >= 0:
        ch = prefix[i]
        if ch in ")]":
            depth += 1
        elif ch in "([":
            depth -= 1
            if depth < 0:
                depth = 0
        elif ch == "," and depth == 0:
            seg_start = i + 1
            break
        i -= 1
    segment = prefix[seg_start:]
    return bool(re.search(r"\bAs\b", segment, re.I))


def _struct_decl_name_span(line):
    """若 line 是 Sub/Function/Property/Event/Type/Enum/Declare 声明行，
    返回它正在声明的那个【名字】在行内的 (start, end)；否则 None。

    v79d/v79e：这些名字是解析器的**语法锚点**。让这一行继续被认作"声明行"
    比"少收一个名字"重要得多 —— 一旦它不再是声明行，这条声明【之后】的代码
    全部失去过程归属，局部变量与隐式变量统统降级成"模块级"，表现就是
    "整个模块到处都能提示出别的过程的变量"。
    所以 v79e 的口径是：**文本一个字都不动**，只在记录层把这一个名字排除
    （见 caret_decl_name_at / extract_records）。
    """
    for rx in (_RE_SUB, _RE_PROP, _RE_EVENT, _RE_TYPE, _RE_ENUM):
        try:
            m = rx.match(line)
        except Exception:
            m = None
        if m:
            return (m.start(1), m.end(1))
    # API 声明：`Private Declare PtrSafe Function gApi Lib "k" ()`
    m = _RE_DECLARE.search(line)
    if m:
        return (m.start(1), m.end(1))
    return None


def caret_decl_name_at(text, line_no, col):
    """光标处的标识符【恰好是这一行正在声明的名字】时返回它（原文），否则 ""。

    v79e：判"哪个标识符盖住了光标"的规则与 _blank_ident_at / ident_at_caret
    完全一致（含"列号越界钳到行尾"），只是额外要求它落在
    `_struct_decl_name_span` 给出的那个名字区间上。

    用途：收集时把"用户正在敲的那个声明名"从候选池里排除。排除放在【记录层】
    而不是【文本层】—— 见 _blank_ident_at 的 v79e 说明。
    """
    if not text or not line_no or not col:
        return ""
    try:
        line_no = int(line_no)
        col = int(col)
    except Exception:
        return ""
    lines = text.split("\n")
    if not (1 <= line_no <= len(lines)):
        return ""
    line = lines[line_no - 1]
    c = col - 1
    if c < 0:
        return ""
    if c > len(line):
        c = len(line)
    anchor = _struct_decl_name_span(line)
    if not anchor:
        return ""
    for m in re.finditer(_IDENT, line):
        if m.start() <= c <= m.end():
            return m.group(0) if m.start() == anchor[0] else ""
    return ""


def _blank_ident_at(text, line_no, col, blank_decl=False):
    """把 (line_no, col) 处（均为 1-based）的标识符抹成等长空格。

    用途：排除"用户此刻正在输入的那个词"。它还没写完、此前也从未被使用过，
    若被当成隐式变量收录，就会出现"打什么就提示什么"的自我提示——
    尤其是单字母（i）或单字（我）这类短词，表现最明显。
    抹成空格而非删除，是为了保持所有字符偏移不变。

    ⚠️ v79e：光标处的词若是这一行**正在声明的名字**（`Sub xxx` / `Type xxx` /
    `Enum xxx` …），**一个字都不动**，直接原样返回（`blank_decl=True` 时例外，
    见下）。

    v79d 那版把它抹成**等长下划线**（想让这一行继续被认作声明行）。方向没错，
    做法有害，两个后果都是真机实测出来的：
      1. **单字符名字会变成行继续符**：`Sub t` 被抹成 `Sub _` 之后，行尾正好是
         "空格 + 下划线 + 换行"——而 _RE_CONTINUATION 就是这个形状（它要求
         `_` 前面是空白、后面到行尾只有空白）⇒ 这一行被当成续行，把【下一行】
         吞进同一行。合并后的文本既不再是过程头（`Sub` 后面跟的是别的东西），
         后面整段过程体也失去归属 ⇒ 那些隐式变量降级成【模块级】，模块里
         任何位置都能提示出来（用户报的"过程级的变量都泄露到其他过程"）。
         实测：`Sub t` 那一行 + 下一行被合并，模块行数 3 -> 2。
      2. 光标停在过程头上时，该过程的记录会挂到占位符名下（`Sub ____`），
         与光标回到过程体内时用的真实名字对不上，变量忽隐忽现。
    所以"正在声明的名字"这一支不再改文本 —— 行结构、过程归属完全不动。
    "打什么不提示什么"由引擎那两道现成防线负责：`ctx["decl_names"]`
    （decl_at_caret 精确剔除光标处那个词）+ 回声防护（拿"现场证据"判它是不是
    真名字）。⚠️ 与之配套的一处必须同步：names_outside_caret（现场证据扫描）
    要传 `blank_decl=True` 强制抹词 —— 否则那个词会把自己算成"别处也出现过"
    的证据，幻影词（第 21.4 节）又会被提示回来。

    其余情形（普通变量 / 赋值目标 / 形参名 / 表达式里的词）照旧抹等长空格。
    自我提示的两道防线（vbe_bridge 的 caret_decl / 引擎的 decl_at_caret）读的是
    【未抹】的原文，不受影响。
    """
    if not line_no or not col:
        return text
    if not blank_decl and caret_decl_name_at(text, line_no, col):
        return text                    # 正在声明的名字：文本一个字不动（v79e）
    try:
        line_no = int(line_no)
        col = int(col)
    except Exception:
        return text
    lines = text.split("\n")
    if not (1 <= line_no <= len(lines)):
        return text
    line = lines[line_no - 1]
    c = col - 1
    if c < 0:
        return text
    # 部分宿主（WPS / 不同 Office 版本）返回的列号可能大于行长，钳到行尾。
    # 只处理越界，不做"回退一格"猜测——`c-1 == ident.end()` 既可能是"正在输入"
    # 也可能是"已打完且后面有空格"，无法区分，猜就会误抹已有标识符。
    if c > len(line):
        c = len(line)
    for m in re.finditer(_IDENT, line):
        if m.start() <= c <= m.end():
            lines[line_no - 1] = (line[:m.start()]
                                  + " " * (m.end() - m.start())
                                  + line[m.end():])
            break
    return "\n".join(lines)


def ident_at_caret(text, line_no, col):
    """返回 (line_no, col) 处（1-based）标识符的原文；不在标识符内则返回 ""。

    与 _blank_ident_at 用同一套"哪个标识符盖住了光标"的判定，区别只在于这里
    返回原文而不是抹掉它。

    v43 用途：后端在收集标识符时记下"那一刻光标处的词"。这个词很可能是用户
    正在输入 / 正在回退删除的词，不是工程里的真名字；引擎据此识别并剔除
    "回退删字时残留的旧片段"（见 engine.Completer.trigger 的回声防护）。
    """
    if not text or not line_no or not col:
        return ""
    lines = text.split("\n")
    if not (1 <= int(line_no) <= len(lines)):
        return ""
    line = lines[int(line_no) - 1]
    c = int(col) - 1
    if c < 0:
        return ""
    if c > len(line):
        c = len(line)
    for m in re.finditer(_IDENT, line):
        if m.start() <= c <= m.end():
            return m.group(0)
    return ""


# `Option Explicit`（强制变量声明）—— 模块级语句，必须出现在所有过程之前。
# 只在【行首】（允许缩进）认，且必须过字符串/注释掩码（见 has_option_explicit）。
_RE_OPTION_LINE = re.compile(r"(?im)^[ \t]*Option\b")
_RE_OPTION_EXPLICIT = re.compile(r"(?im)^[ \t]*Option[ \t]+Explicit\b")


def has_option_explicit(code):
    """模块是否写了 `Option Explicit`（= 强制声明变量）。

    纯文本判定，不碰 COM、可单测。判据要点：

      * 只在【行首】（允许前导空白）认 —— `Option Explicit` 是模块级语句，
        不会出现在语句中间；
      * **先过 _mask_strings_and_comments**：字符串里的 `"Option Explicit"` 与
        注释里的 `' Option Explicit` 都不算。真实工程里"临时注释掉这一行"很常见，
        认错了会把整个模块的隐式变量误砍掉（而那是用户明确要保留的行为）；
      * 大小写不敏感（`option explicit` 同样合法）；
      * 只认 Explicit —— `Option Base 1` / `Option Compare Text` /
        `Option Private Module` 都与"是否必须声明变量"无关，不算数。

    为什么要判它：写了 Option Explicit 的模块里，"用过但没声明"的名字是
    **编译错误**，不该再进候选池（见 vbe_bridge.IMPLICIT_HONOR_OPTION_EXPLICIT）。
    """
    if not code:
        return False
    # 廉价预筛：绝大多数模块连 Option 语句都没有，连掩码都不用做。
    if not _RE_OPTION_LINE.search(code):
        return False
    try:
        return bool(_RE_OPTION_EXPLICIT.search(_mask_strings_and_comments(code)))
    except Exception:
        # 判定失败一律当"没写"—— 宁可多给提示，也不要因为一个解析意外
        # 把整个模块的隐式变量砍掉。
        return False


def extract_implicit_records(code, module=None, is_std_module=True,
                             declared=None, scope="module", include_usage=True,
                             caret=None):
    """提取"未声明但被使用过"的隐式变量（不写 Option Explicit 的场景）。

    返回 [(name, module, proc, priv), ...]，与 extract_records 同构：
      - 只收录没有出现在 declared 里的名字（已声明的由 extract_records 负责，
        它带更准确的作用域）；
  - scope == "module"：所有隐式变量记为模块级且 priv=True —— 本模块任意位置
    可见、跨模块不提示（隐式变量本就是"松散"用法，全模块可提示更实用）。
  - scope == "proc"  （推荐，本项目默认启用）：
    * 赋值 / For / For Each 产生的变量，记为「首次出现所在过程的局部变量」，
      更贴近 VBA 语义（未声明变量的作用域是其首次出现的过程）。同名变量在
      不同过程里各记各的，因此不会跨过程泄漏（函数B的 num2 不会出现在函数A）。
    * 只读用法（只被读取、从未被赋值）的隐式变量仍记为模块级，本模块内任意
      过程可见，避免"松散全局变量"因陈旧过程标签而丢失。

    判定来源：赋值目标（x = ... / Set x = ... / Let x = ... / x(i) = ...）
    与 For / For Each 循环变量。字符串、注释、声明行均已排除。

    declared：已声明名字。可为名称序列（旧式：一律视为全模块可见而排除），
    也可为 extract_records 的记录序列 [(name, module, proc, priv), ...]
    （推荐）——此时按作用域精确排除：模块级声明全模块排除；过程级声明
    （Dim / 参数）仅在同名过程内排除，避免「其它过程的同名参数挡掉本过程
    的隐式变量」——真实案例：主过程里的 num 被 Function f(ws, num) 的参数
    挡住，导致主过程深处输入 n 提示不出 num。

    caret=(line_no, col)：光标位置（1-based）。该位置上的标识符会被忽略，
    避免"正在输入的词"被当成已使用过的变量收录（自我提示）。

    ⚠️ 本函数【不自己判 Option Explicit】—— 那份判定（连同开关）住在调用方
    `vbe_bridge.implicit_gate_applies`：写了 Option Explicit 的模块干脆不该调
    本函数（"用过但没声明"在那里是编译错误），先问一次能省掉整模块解析。
    """
    # 光标处的词必须在【原始文本】上抹掉，不能先 mask 再抹。
    #
    # 原因：VBE 给的列号基于原始文本，而 _mask_strings_and_comments 会把字符
    # 串/注释压缩掉（例如 "Scripting.FileSystemObject" 28 字符 -> 2 个空格），
    # **不保证等长**（真实工程实测有 18 行长度不一致，最多差 26 字符）。
    # 若先 mask，列号相对 mask 文本会整体右移，_blank_ident_at 就抹错位置，
    # 光标处正在输入的词会被当成"已使用过的隐式变量"收录，表现为自我提示
    # （打 n 提示 n）。真实案例：UserForm3 第 61 行，caret=(61,54)。
    src = code
    if caret:
        try:
            src = _blank_ident_at(src, caret[0], caret[1])
        except Exception:
            pass
    masked = _mask_strings_and_comments(src)
    masked = _RE_CONTINUATION.sub(" ", masked)
    # v61：Type / Enum 块内的成员行整行抹空 —— 否则下面的"只读用法"扫描会把
    # 枚举成员名当成隐式变量收回来（详见 _blank_type_enum_blocks）。
    masked = _blank_type_enum_blocks(masked)

    # v79e：同名过程（用户正在敲一个与既有过程同名的过程头）—— 记录里必须能把
    # 两次出现分开，否则第二个过程会提示出第一个过程的隐式变量。
    _occ = {}                # 小写过程名 -> 已出现次数
    # 光标正压着的"正在声明"的名字照样收（见 extract_records 的说明）：
    # 防自我提示由引擎的 decl_names + 回声防护负责，解析层不动文本、也不排除。

    # 已声明名字。支持两种传法：
    #   - 名称序列（旧式）：视为全模块可见，一律排除；
    #   - 记录序列 [(name, module, proc, priv), ...]：按作用域精确排除。
    # 后者是修复「别的函数的同名参数把本过程的隐式变量挡掉」的关键：
    # 主过程里的 num 与 Private Function f(ws, num) 的参数同名时，若把 num
    # 当成全局已声明而跳过，主过程里就永远提示不出 num（真实案例）。
    declared_global = set()     # 模块级声明（proc 为空）：全模块排除
    declared_by_proc = {}       # 过程级声明（Dim / 参数）：仅同名过程内排除

    def _note_declared(name, proc):
        if name is None:
            return
        n = str(name).lower()
        if proc:
            declared_by_proc.setdefault(str(proc).lower(), set()).add(n)
        else:
            declared_global.add(n)

    def _declare_from_records():
        for n, _m, p, _pr in extract_records(code, module, is_std_module):
            _note_declared(n, p)

    if declared is None:
        _declare_from_records()
    else:
        items = list(declared)
        # 记录序列 [(name, module, proc, priv), ...] 能携带作用域，按作用域精确排除。
        # 旧式的"纯名字序列"做不到：它一旦被当成全模块已声明，别的过程的同名
        # 参数（如 Function f(ws, num)）就会把本过程的 num 彻底挡掉。
        # 因此遇到旧式传参时，改为按本模块重新计算，保证调用方即使未同步
        # 升级也不会漏提示。
        if items and all(isinstance(i, (tuple, list)) for i in items):
            for item in items:
                _note_declared(item[0] if len(item) > 0 else None,
                               item[2] if len(item) > 2 else None)
        else:
            _declare_from_records()

    # 供"只读用法"扫描排除用（保留旧式的全量排除，避免引入跨过程泄漏）
    declared_all = set(declared_global)
    for _s in declared_by_proc.values():
        declared_all |= _s

    def _already_declared(name, proc):
        n = name.lower()
        if n in declared_global:
            return True
        if proc:
            return n in declared_by_proc.get(str(proc).lower(), ())
        return False

    result = []
    seen = set()
    cur_proc = None
    in_block = False
    implicit_names = set()   # 已被隐式收录的名字，供"只读用法"扫描排除（不混入 declared）

    def add(name, proc):
        key = (name.lower(), (proc or "").lower(), (module or "").lower())
        if key in seen:
            return
        seen.add(key)
        result.append((name, module, proc, True))

    for raw in masked.split("\n"):
        line = raw.strip()
        if not line:
            continue
        if _RE_PROC_END.match(line):
            cur_proc = None
            continue
        m_sub = _RE_SUB.match(line) or _RE_PROP.match(line)
        if m_sub:
            _n59 = m_sub.group(1)
            _k59 = str(_n59).lower()
            _occ[_k59] = _occ.get(_k59, 0) + 1
            cur_proc = _proc_scope_key(_n59, _occ[_k59])
            continue
        if _RE_TYPE.match(line) or _RE_ENUM.match(line):
            in_block = True
            continue
        if _RE_BLOCK_END.match(line):
            in_block = False
            continue
        if in_block:
            continue      # Type/Enum 成员行不是隐式变量
        if _RE_DECL.search(line) or _RE_CONST.search(line) or _RE_SKIP_LINE.match(line):
            continue

        # 单行 If ... Then y = 2 [Else z = 3] / Else y = 2
        stmts = [line]
        m_if = _RE_IF_THEN.match(line) or _RE_ELSE_ONLY.match(line)
        if m_if:
            tail = m_if.group(1)
            stmts = [p.strip() for p in re.split(r"\bElse\b", tail)]
        elif line[:3].lower() == "if ":
            continue

        for stmt in stmts:
            for name in _implicit_names_in_stmt(stmt):
                if (_already_declared(name, cur_proc)
                        or name.lower() in _IMPLICIT_STOPWORDS):
                    continue
                # 按"名字+所在过程"去重：同一名字在不同过程里各算各的局部变量
                # （VBA 未声明变量的作用域是"首次出现的过程"），不再跨过程合并，
                # 否则会出现"函数2定义的 num 在函数1里提示不出来"的反向 bug。
                implicit_names.add(name.lower())
                add(name, cur_proc if scope == "proc" else None)

    # 只被读取、没被写过的隐式变量（如 total = 总行数 * 2 里的 总行数）
    #
    # 作用域与赋值路径保持一致：记为「首次出现所在过程」的局部变量。
    # 旧实现一律记为模块级（proc=None），于是 A 过程里只读用到的名字会出现在
    # B 过程里 —— 也就是"一个函数里提示出别的函数的变量"。VBA 里未声明变量
    # 的作用域本就是其首次出现的过程，因此按过程限定才是对的。
    if include_usage:
        offsets, line_proc = _line_proc_map(masked)
        for name, pos in _usage_candidates(masked, declared_all | implicit_names):
            implicit_names.add(name.lower())
            add(name, _proc_at_offset(offsets, line_proc, pos)
                if scope == "proc" else None)

    return result


def _line_proc_map(masked):
    """返回 (每行起始字符偏移, 每行所属过程名) —— 供"只读用法"定位作用域。

    行号必须在【续行合并之后】的文本上算：_RE_CONTINUATION 会把 ` _\\n` 压成
    空格，行数与原始文本已经不同，直接用原始行号会整体错位。
    v79e：过程名同样按 _proc_scope_key 变成作用域键（第 2 次同名起带 `#序号`），
    必须与 extract_implicit_records 主循环里的键一致 —— 否则"只读用法"扫出来
    的名字会挂到对不上的过程键上（表现为变量忽隐忽现）。
    """
    lines = masked.split("\n")
    _occ = {}
    offsets = []
    line_proc = []
    off = 0
    cur = None
    for raw in lines:
        offsets.append(off)
        off += len(raw) + 1
        line = raw.strip()
        if _RE_PROC_END.match(line):
            cur = None
        m_sub = _RE_SUB.match(line) or _RE_PROP.match(line)
        if m_sub:
            _n = m_sub.group(1)
            _k = str(_n).lower()
            _occ[_k] = _occ.get(_k, 0) + 1
            cur = _proc_scope_key(_n, _occ[_k])
        line_proc.append(cur)
    return offsets, line_proc


def _proc_at_offset(offsets, line_proc, pos):
    """字符偏移 -> 该行所属过程名。"""
    if not offsets:
        return None
    i = bisect.bisect_right(offsets, pos) - 1
    if 0 <= i < len(line_proc):
        return line_proc[i]
    return None


def extract_type_enum_names(code):
    """返回代码里用 Type / Enum 定义的类型名（去重，保留出现顺序）。

    这些名字可以作为数据类型使用（`Dim x As MyType`），因此在 `As` 之后的
    类型名位置应当参与提示。
    """
    out = []
    seen = set()
    if not code:
        return out
    masked = _RE_CONTINUATION.sub(" ", _mask_strings_and_comments(code))
    for raw in masked.split("\n"):
        m = _RE_TYPE.match(raw.strip()) or _RE_ENUM.match(raw.strip())
        if m and m.group(1).lower() not in seen:
            seen.add(m.group(1).lower())
            out.append(m.group(1))
    return out


def extract_scoped(code):
    """
    旧式接口：返回 [(name, scope), ...]。

    scope 语义（为向后兼容旧调用方/单测保留）：
      - None      ：模块级
      - "过程名"  ：该过程内部的局部变量与参数

    注意：这里不携带模块与 Private/Public 信息；需要完整作用域请用 extract_records。
    """
    out = []
    seen = set()
    for name, _module, proc, _priv in extract_records(code, module=None, is_std_module=True):
        scope = proc  # 模块级 proc=None -> None
        key = (name.lower(), scope.lower() if scope else None)
        if key not in seen:
            seen.add(key)
            out.append((name, scope))
    return out


def is_proc_header_line(line_text):
    """这一行是不是【过程头】（`Sub xxx` / `Function xxx` / `Property Get|Let|Set xxx`）。

    v79d：给作用域判定用。光标正停在过程头那一行时，这一行**不属于任何过程的
    "体内"** —— 它是声明，不是过程体。若把它算成"当前过程"，就会出现：
    正在写新过程头 `Sub test`，而模块里已有一个 `Sub test()`，于是那个过程的
    局部变量（如 targetsheet）被当成"当前过程的"提示出来（用户报的泄漏）。
    过程头必须【严格在光标上方】才算"把光标圈在里面"。

    只认"名字后面是 `(` 或行尾"的形式，与 proc_at_line 的判据完全一致
    （`Sub` / `Function` / `Property Get` 是保留字，不可能出现在过程体语句开头，
    所以过程体内不会有行被误判成过程头）。
    """
    if not line_text:
        return False
    s = _mask_strings_and_comments(line_text).strip()
    if not s:
        return False
    try:
        return bool(_RE_SUB.match(s) or _RE_PROP.match(s))
    except Exception:
        return False


def _proc_header_name(line):
    """行是过程头（Sub / Function / Property Get|Let|Set）则返回其名字，否则 None。

    与 proc_at_line / proc_owns_line 用的是同一对正则（_RE_SUB / _RE_PROP），
    保证"数次数"与"找最近的过程头"两处认的是同一批行。
    """
    if not line:
        return None
    s = str(line).strip()
    if not s:
        return None
    m = _RE_SUB.match(s) or _RE_PROP.match(s)
    return m.group(1) if m else None


def _proc_scope_key(name, occurrence):
    """过程在【作用域过滤】里用的键。

    VBA 不允许同一模块里出现两个同名过程（编译期报"二义性名称"），但用户
    正在敲新过程头时会短暂处于这种状态（真实案例：模块里已有 `Sub test()`，
    用户又敲了一个 `Sub test()` 做试验）。记录里只存"过程名"的话，这两个过程
    的局部变量在过滤时无法区分 —— 第二个过程的体内就会提示出第一个过程的
    局部变量，表现正是"过程级的变量泄漏到其他过程"。

    口径：**该名字在模块里第 1 次出现时用原名，第 2 次及以后写成 `名字#序号`**
    （序号 = 第几次出现，按过程头从上到下数，从 1 起）。

    为什么第 1 次不带序号（而不是"重名才全部带序号"）：
      * 名字唯一时（正常代码）键必须与旧版完全一致 —— 既有记录、测试、日志
        都不受影响；
      * 更要紧的是两侧的文本范围不同：记录侧拿的是【整个模块】，而光标侧
        （vbe_bridge._proc_of_line）为了省开销只读【光标以上】的那一段文本。
        "重名才带序号"会让第一段文本数不到下面的重名，两侧对不上键，把第一个
        过程的变量一起挡掉（实测踩到）。"第 1 次用原名"则对前缀截断天然稳定：
        第 1 次出现的前缀里永远只有它自己。
      * 真机验证（模块2，两个 `Sub test`）：第 1 个过程内 proc='test'、
        第 2 个过程内 proc='test#2'，记录侧 targetSheet 归 'test' -> 第 2 个
        过程不再提示它，第 1 个过程照常提示。
    """
    if not name:
        return name
    try:
        occ = int(occurrence)
    except Exception:
        occ = 1
    if occ <= 1:
        return name
    return "%s#%d" % (name, occ)


def proc_key_plain(name):
    """把 `名字#序号` 拆回名字本身；本来就没有序号则原样返回。

    给"只要名字"的地方用：COM 兜底（ProcOfLine 给的是裸名）与 proc_owns_line
    的比对。
    """
    s = str(name or "")
    i = s.rfind("#")
    if i > 0 and s[i + 1:].isdigit():
        return s[:i]
    return s


def _proc_key_at(lines, idx, name):
    """过程头位于第 idx 行（0-based）、名字为 name 时，返回它的作用域键。

    序号按【本次扫描到的文本】里"这个名字第几次出现"算（见 _proc_scope_key）：
    记录侧扫整个模块、光标侧只扫光标以上那一段，第 1 次出现两边都数得到 1，
    因此即使文本范围不同也稳定一致。
    """
    if not name:
        return name
    occ = 0
    last = min(int(idx), len(lines) - 1)
    for j in range(0, last + 1):
        n = _proc_header_name(lines[j])
        if n and n.lower() == str(name).lower():
            occ += 1
    return _proc_scope_key(name, max(occ, 1))


def proc_at_line(code, line_no):
    """返回第 line_no 行（1-based）所属过程【的作用域键】；在模块声明区则返回 None。

    做法：从该行往上扫描，遇到最近的 Sub/Function/Property 头就返回其名字；
    若先遇到 End Sub/Function/Property/Type/Enum，说明在模块声明区，返回 None。

    返回值通常是过程名；只有当同一模块里存在【同名过程】时才带 `#序号`
    后缀（见 _proc_scope_key）—— 那种状态下"过程名"不足以区分两个过程，
    作用域过滤必须能分开，否则第二个过程会提示出第一个过程的局部变量。

    ⚠️ 本函数回答的是"这一行归哪个过程"（过程头那一行归它自己，这是"归属"的
    自然口径，proc_owns_line 也据此回答）。**作用域判定不要直接用它** ——
    光标停在过程头上时，正确的"当前作用域"是模块级（那一行是声明，不是过程体），
    见 `is_proc_header_line` 与 vbe_bridge._proc_of_line 的 v79d 说明。

    纯文本实现，不依赖 COM（VBE 的 ProcOfLine 在 pywin32 下常因 byref
    参数 ProcKind 抛异常而取不到值），且可单测。
    """
    if not code or not line_no or line_no < 1:
        return None
    try:
        line_no = int(line_no)
    except Exception:
        return None
    lines = _mask_strings_and_comments(code).split("\n")
    idx = min(line_no, len(lines)) - 1
    while idx >= 0:
        line = lines[idx].strip()
        if line:
            m = _RE_SUB.match(line) or _RE_PROP.match(line)
            if m:
                return _proc_key_at(lines, idx, m.group(1))
            if _RE_ANY_PROC_END.match(line):
                return None
        idx -= 1
    return None


def proc_owns_line(code, line_no, name):
    """过程 name 是否真的把第 line_no 行（1-based）圈在里面。纯函数。

    与 proc_at_line 的分工：
      * proc_at_line 答"从这一行往上，最近的过程头是谁"；
      * 本函数答"这个名字的过程是不是真的包含这一行" —— 从该行往上找，先遇到
        它的过程头 -> True；先遇到任何过程收尾（End Sub/Function/Property）
        -> False（说明这一行其实在过程【外】：`End Sub` 那行本身、过程之间的
        空行、模块声明区）。
    name 为空 / 找不到它的过程头 -> False（宁可不认，也不认错）。
    name 允许带 `#序号` 后缀（作用域键），比对时只看名字部分。

    为什么需要它（v79c）：VBE 的 CodeModule.ProcOfLine 会把手伸到过程外 ——
    实测 `End Sub` 及其下方直到下一个过程头之间的行都归给【前一个】过程，
    声明区里第一个过程头之前的行归给【第一个】过程。盲信它，就会出现"在模块级
    / 新函数的位置上，前一个函数的局部变量全被当成当前过程的" —— 也就是
    "别的函数的变量泄漏到本位置"。
    """
    if not code or not line_no or line_no < 1 or not name:
        return False
    want = proc_key_plain(name).strip().lower()
    if not want:
        return False
    try:
        line_no = int(line_no)
    except Exception:
        return False
    lines = _mask_strings_and_comments(code).split("\n")
    idx = min(line_no, len(lines)) - 1
    while idx >= 0:
        line = lines[idx].strip()
        if line:
            m = _RE_SUB.match(line) or _RE_PROP.match(line)
            if m and m.group(1).lower() == want:
                return True
            if _RE_ANY_PROC_END.match(line):
                return False
        idx -= 1
    return False


def extract_identifiers(code):
    """返回模块中所有可补全标识符名（去重，保留首次出现顺序）。"""
    out = []
    seen = set()
    for name, _scope in extract_scoped(code):
        if name.lower() not in seen:
            seen.add(name.lower())
            out.append(name)
    return out


def extract_scoped_from_modules(modules):
    """旧式接口：modules: list[str] -> 合并去重后的 [(name, scope), ...]。"""
    result = []
    seen = set()
    for src in modules:
        for name, scope in extract_scoped(src):
            key = (name.lower(), scope.lower() if scope else None)
            if key not in seen:
                seen.add(key)
                result.append((name, scope))
    return result


def extract_from_modules(modules):
    """旧式接口：modules: list[str]，返回合并去重后的标识符名列表。"""
    all_ids = []
    for src in modules:
        all_ids.extend(extract_identifiers(src))
    # 再次去重（跨模块）
    seen = set()
    out = []
    for n in all_ids:
        if n.lower() not in seen:
            seen.add(n.lower())
            out.append(n)
    return out
