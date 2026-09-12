"""
VBA 标识符提取器。

从一段 VBA 模块源码文本中，提取出可用于自动补全的所有标识符：
  - 变量：Dim / Private / Public / Global / Static / Friend (+ ReDim)
  - 常量：Const
  - 过程：Sub / Function（含参数名）
  - 属性：Property Get / Let / Set（含参数名）
  - API 声明：Declare Sub / Function（含参数名）
  - 自定义类型与枚举：Type / Enum（含成员名）

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

_RE_CONTINUATION = re.compile(r"_\s*\n")  # 行继续符 _
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
# PtrSafe 是 64 位 Office 下 Declare 的修饰关键字，需一并容忍
_RE_DECLARE = re.compile(
    r"\bDeclare\s+(?:PtrSafe\s+)?(?:Sub|Function)\s+(" + _IDENT + r")\s*", re.I)
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
KIND_MEMBER = "member"  # Type/Enum 成员：跟随所属块的可见性
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


def _mask_strings_and_comments(code):
    """把字符串字面量和注释替换为等长的空格，避免误匹配。"""
    out = []
    i, n = 0, len(code)
    while i < n:
        c = code[i]
        if c == '"':
            out.append(" ")
            i += 1
            while i < n:
                if code[i] == '"':
                    if i + 1 < n and code[i + 1] == '"':  # 转义引号 ""
                        i += 2
                        continue
                    i += 1
                    break
                i += 1
            out.append(" ")
            continue
        if c == "'":  # 注释到行尾
            while i < n and code[i] != "\n":
                out.append(" ")
                i += 1
            continue
        out.append(c)
        i += 1
    return "".join(out)


def _decl_names(rest):
    """从声明右侧（如 'x As Long, y(), z As String' 或 'a = 5, b = 6'）提取标识符名。"""
    names = []
    for seg in rest.split(","):
        seg = seg.strip()
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

    caret=(line_no, col)：光标位置。若光标正处在变量/常量声明行上，该行声明的
    名字会被排除——它们尚未定义完成，否则会出现"输入 g 就提示 gCounter 本身"
    的自我提示（显式声明路径此前一直缺少光标排除，只有隐式变量路径有）。
    """
    masked = _mask_strings_and_comments(code)
    # 合并行继续符
    masked = _RE_CONTINUATION.sub(" ", masked)

    # 注意：早先这里会按 caret 行把"正在声明的名字"从候选里剔除（连前缀一起）。
    # 但这会把合法的【前缀补全】也误杀：在 numArr 的声明行上输入 num，numArr
    # 就再也提示不出来（bug：输入 num 只能提醒 num、不提醒 numArr）。正确的语义是
    # "精确匹配自己才隐藏、前缀照常提示"，而这由引擎层 in_decl_position 的精确
    # 自我剔除负责（输入完整 numArr 才隐藏，输入前缀 num 仍正常提示 numArr）。
    # 因此解析层不再做任何 caret 排除，名字一律收录，交给引擎按"正在输入的词"
    # 决定是否剔除——这样既修好了前缀补全，又没丢"定义变量时不提示自己"。

    result = []
    seen = set()
    cur_proc = None      # 当前所在过程名（None = 模块声明区）
    in_block = None      # 'Type' / 'Enum'
    block_priv = False   # 当前 Type/Enum 块的可见性（成员继承）

    def add(name, proc, priv):
        if not name:
            return
        key = (name.lower(), (proc or "").lower(), (module or "").lower())
        if key not in seen:
            seen.add(key)
            result.append((name, module, proc, priv))

    for raw in masked.split("\n"):
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
            block_priv = priv
            name = m_type.group(1) if m_type else m_enum.group(1)
            add(name, None, priv)          # 类型/枚举名
            continue
        if _RE_BLOCK_END.match(line):
            in_block = None
            continue

        if in_block:
            # 成员行：member As Type  或  member = value（枚举）
            m = _RE_LEADING_IDENT.match(line)
            if m and m.group(0).lower() not in ("type", "enum", "end"):
                add(m.group(0), None, block_priv)   # 成员继承块可见性
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
            cur_proc = proc
            pm = re.search(r"\((.*)\)", line)
            if pm:
                for p in _params_inside(pm.group(1)):
                    add(p, proc, False)
            continue

        # Property Get/Let/Set
        m_prop = _RE_PROP.match(line)
        if m_prop:
            proc = m_prop.group(1)
            priv = _is_private(KIND_PROC, has_priv_kw, has_pub_kw, is_std_module)
            add(proc, None, priv)
            cur_proc = proc
            pm = re.search(r"\((.*)\)", line)
            if pm:
                for p in _params_inside(pm.group(1)):
                    add(p, proc, False)
            continue

        # Const 常量：作用域取决于声明位置
        if _RE_CONST.search(line):
            kind = KIND_CONST
            priv = _is_private(kind, has_priv_kw, has_pub_kw, is_std_module)
            body = _RE_CONST.sub("", line, count=1).strip()
            body = re.sub(r"^(?:Private|Public|Global|Friend)\b", "", body, flags=re.I).strip()
            for seg in body.split(","):
                m = _RE_LEADING_IDENT.match(seg.strip())
                if m:
                    # 过程内 Const 永远是局部
                    if cur_proc:
                        add(m.group(0), cur_proc, False)
                    else:
                        add(m.group(0), None, priv)
            continue

        # 普通变量声明 Dim/Private/...（非 Const）
        if _RE_DECL.search(line):
            kind = KIND_VAR
            priv = _is_private(kind, has_priv_kw, has_pub_kw, is_std_module)
            body = _RE_DECL.sub("", line, count=1).strip()
            for seg in body.split(","):
                m = _RE_LEADING_IDENT.match(seg.strip())
                if m:
                    # 过程内 Dim/Static = 局部
                    if cur_proc:
                        add(m.group(0), cur_proc, False)
                    else:
                        add(m.group(0), None, priv)
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
    # 编译指令 / API 声明 / 模块属性行不参与用法扫描（否则函数名、库名会被当变量）
    kept = "\n".join("" if _RE_USAGE_SKIP_LINE.match(l) else l
                     for l in masked_code.split("\n"))
    masked_code = kept
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
    # 先屏蔽字符串/注释再判断，避免 `"As ..."` 里的 As 误判
    masked = _mask_strings_and_comments(line_text)
    end = caret_col - 1
    if end < 0:
        end = 0
    if end > len(masked):
        end = len(masked)
    prefix = masked[:end]
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


def _blank_ident_at(text, line_no, col):
    """把 (line_no, col) 处（均为 1-based）的标识符抹成等长空格。

    用途：排除"用户此刻正在输入的那个词"。它还没写完、此前也从未被使用过，
    若被当成隐式变量收录，就会出现"打什么就提示什么"的自我提示——
    尤其是单字母（i）或单字（我）这类短词，表现最明显。
    抹成空格而非删除，是为了保持所有字符偏移不变。
    """
    if not line_no or not col:
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
            cur_proc = m_sub.group(1)
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
                if _already_declared(name, cur_proc) or name.lower() in _IMPLICIT_STOPWORDS:
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
    """
    offsets = []
    line_proc = []
    off = 0
    cur = None
    for raw in masked.split("\n"):
        offsets.append(off)
        off += len(raw) + 1
        line = raw.strip()
        if _RE_PROC_END.match(line):
            cur = None
        m_sub = _RE_SUB.match(line) or _RE_PROP.match(line)
        if m_sub:
            cur = m_sub.group(1)
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


def proc_at_line(code, line_no):
    """返回第 line_no 行（1-based）所属的过程名；在模块声明区则返回 None。

    做法：从该行往上扫描，遇到最近的 Sub/Function/Property 头就返回其名字；
    若先遇到 End Sub/Function/Property/Type/Enum，说明在模块声明区，返回 None。

    纯文本实现，不依赖 COM（VBE 的 ProcOfLine 在 pywin32 下常因 byref
    参数 ProcKind 抛异常而取不到值），且可单测。
    """
    if not code or not line_no or line_no < 1:
        return None
    lines = _mask_strings_and_comments(code).split("\n")
    idx = min(line_no, len(lines)) - 1
    while idx >= 0:
        line = lines[idx].strip()
        if line:
            m = _RE_SUB.match(line) or _RE_PROP.match(line)
            if m:
                return m.group(1)
            if _RE_ANY_PROC_END.match(line):
                return None
        idx -= 1
    return None


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
