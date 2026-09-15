"""VBA 语言自带的名字：内建函数、内建常量、内建数据类型、语言关键字。

这些名字【不在用户的 VBE 工程代码里】—— 它们由 VBA 语言运行时提供，代码文本里
一个字都找不到（除非用户恰好调用过 MsgBox / Left 之类）。所以必须像组件名、窗体
控件名那样，作为一路独立的来源喂给候选池（见 vbe_bridge._collect_identifiers）。

收录范围：
  * FUNCTIONS  —— VBA 库的内建过程与函数（类型转换 / 字符串 / 日期 / 数组 / 数学 /
                  文件 IO / 交互 / 注册表 / 财务），含 Open、Close、Print 这类以
                  语句形式出现的 VBA 内建过程；
  * CONSTANTS  —— 常用的 vb* 内建常量（换行符、MsgBox 参数与返回值、StrConv、
                  日期常量、VarType、文件属性、颜色等）；
  * TYPE_NAMES —— VBA 内建数据类型（`Dim x As <这里>` 用得上）；
  * KEYWORDS   —— 语言关键字 / 保留字（v62，按用户要求："把 vba 里面的所有
                  关键字都纳入到提示词里"）：过程与声明（Sub / Function / Dim /
                  Set …）、流程控制（If / For / Select …）、字面量（True /
                  Nothing …）、运算符（And / Mod …）。

【v61 与 v62 的口径变化】v61 时关键字是刻意不收的（用户当时讨厌 If/For/Sub 这类
名字混进候选池）；v62 用户明确要求收进来，于是单列 KEYWORDS 一组，并配独立的开关
`vbe_bridge.ENABLE_VBA_KEYWORDS` —— 想关掉只关这一组，不影响内建函数/常量。

【v70 的口径变化】用户要求"**内置枚举（vb*）不要提示，内建函数保留**"——
理由是内置枚举他基本不用，提示出来只是干扰。于是 CONSTANTS 与 FUNCTIONS 拆成两组
独立导出（BUILTIN_CONSTANTS / BUILTIN_FUNCTIONS），收不收由
`vbe_bridge.ENABLE_VBA_CONSTANTS`（**默认 False**）决定；同一口径下宿主类型库的
枚举常量（xl* / mso*）也默认不收（`vbe_bridge.ENABLE_HOST_ENUMS` 默认 False）。
⚠️ 本模块的清单本身【不裁剪】，开关只在 vbe_bridge 的收集段生效 ——
这样测试 / 诊断仍能看到完整名单，也方便用户随时把开关打开。

【v77 的口径变化】用户接着说"**我觉得提示词太多了……vba 本身带的关键字、函数这块
都删掉吧。保留能提示变量名、自定义的函数/过程名、窗体/控件名、模块名**"——
于是 FUNCTIONS / TYPE_NAMES / KEYWORDS 也改成**默认不收**
（`vbe_bridge.ENABLE_VBA_BUILTINS` / `ENABLE_VBA_KEYWORDS` 均默认 False）。
默认口径下候选池只剩"这个工程里真实存在的东西"。
⚠️ 同样**一个字都不裁**（清单保持完整、收集门都还在），四个开关彼此独立，
想收哪一批就 `set VBECOMPLETE_xxx=1`。见回归测试第 54 节。

刻意【不收】的：
  * 过时的 Def* 类型声明关键字（DefBool / DefInt / … / DefVar，11 个）—— 早在
    VB6 时代就已废弃，实际代码里几乎不会手写；
  * 歧义过大的语句名（Name / Width / Line / Spc / Tab）—— 它们在代码里更常作为
    变量名或属性名出现，提示出来容易误导（v61 起就没有收录，v62 也没放回来）；
  * Option Base / Option Compare 的参数名（Base / Compare / Binary / Text /
    Database）—— 不是保留字，且都是极常见的普通变量名；
  * vbKey* 键盘常量（约 70 个）—— 前缀高度集中、实际极少手写；
  * Excel / Office 对象模型（Application / Worksheets / Range …）—— 那是宿主库、
    不是 VBA 自带，且 VBE 自己会提示点号后的成员。

清单按 VBA 文档 + VBE 对象浏览器（VBA 库）整理；名字用 VBA 惯例的大小写（插入
到代码里就是这个形式）。全部成员都是合法标识符，无需再做正则校验。
"""

# --------------------------------------------------------------------------
# 内建函数 / 过程（VBA 库成员）
# --------------------------------------------------------------------------
FUNCTIONS = (
    # ---- 类型转换与判定 ----
    "CBool", "CByte", "CCur", "CDate", "CDbl", "CDec", "CInt", "CLng",
    "CLngLng", "CLngPtr", "CSng", "CStr", "CVar", "CVErr",
    "IsArray", "IsDate", "IsEmpty", "IsError", "IsMissing", "IsNull",
    "IsNumeric", "IsObject",
    "TypeName", "VarType",

    # ---- 字符串 ----
    "Asc", "AscB", "AscW", "Chr", "ChrB", "ChrW",
    "Format", "FormatCurrency", "FormatDateTime", "FormatNumber",
    "FormatPercent",
    "Hex", "Oct", "InStr", "InStrRev", "Join", "LCase", "Left", "LeftB",
    "Len", "LenB", "LTrim", "Mid", "MidB", "Replace", "Right", "RightB",
    "RTrim", "Space", "Split", "Str", "StrComp", "StrConv", "String",
    "StrReverse", "Trim", "UCase", "Val", "Filter",

    # ---- 数学 ----
    "Abs", "Atn", "Cos", "Exp", "Fix", "Int", "Log", "Rnd", "Round", "Sgn",
    "Sin", "Sqr", "Tan",

    # ---- 日期与时间 ----
    "Date", "DateAdd", "DateDiff", "DatePart", "DateSerial", "DateValue",
    "Day", "Hour", "Minute", "Month", "MonthName", "Now", "Second", "Time",
    "Timer", "TimeSerial", "TimeValue", "Weekday", "WeekdayName", "Year",

    # ---- 数组 ----
    "Array", "LBound", "UBound",

    # ---- 交互 / 宿主 ----
    "MsgBox", "InputBox", "DoEvents", "Shell", "AppActivate", "SendKeys",
    "Command", "Environ", "CurDir",

    # ---- 对象 ----
    "CallByName", "Choose", "CreateObject", "GetObject", "IIf", "Partition",
    "QBColor", "RGB", "Switch",

    # ---- 错误处理 ----
    "Error",

    # ---- 文件与输入输出 ----
    "Dir", "EOF", "FileAttr", "FileDateTime", "FileLen", "FreeFile",
    "GetAttr", "SetAttr", "FileCopy", "Kill", "MkDir", "RmDir", "ChDir",
    "ChDrive",
    "Open", "Close", "Get", "Put", "Input", "Seek", "Lock", "Unlock",
    "Loc", "LOF", "Print", "Write", "Reset", "LSet", "RSet",
    "LoadPicture", "SavePicture",

    # ---- 注册表设置 ----
    "SaveSetting", "GetSetting", "GetAllSettings", "DeleteSetting",

    # ---- 财务 ----
    "DDB", "FV", "IPmt", "IRR", "MIRR", "NPer", "NPV", "Pmt", "PPmt", "PV",
    "Rate", "SLN", "SYD",

    # ---- 窗体（VBA 库内建过程）----
    "Load", "Unload",

    # ---- 其它内建过程（语句形式）----
    "Beep", "Erase", "Randomize",
)

# --------------------------------------------------------------------------
# 常用内建常量（VBA 库常量）
# --------------------------------------------------------------------------
_CONST_STRINGS = (
    "vbCr", "vbCrLf", "vbLf", "vbNewLine", "vbNullChar", "vbNullString",
    "vbTab", "vbBack", "vbFormFeed", "vbVerticalTab",
)
_CONST_MSGBOX = (
    # 按钮组合
    "vbOKOnly", "vbOKCancel", "vbAbortRetryIgnore", "vbYesNoCancel",
    "vbYesNo", "vbRetryCancel",
    # 图标
    "vbCritical", "vbQuestion", "vbExclamation", "vbInformation",
    # 默认按钮
    "vbDefaultButton1", "vbDefaultButton2", "vbDefaultButton3",
    "vbDefaultButton4",
    # 模态 / 外观
    "vbApplicationModal", "vbSystemModal", "vbMsgBoxHelpButton",
    "vbMsgBoxSetForeground", "vbMsgBoxRight", "vbMsgBoxRtlReading",
    # 返回值
    "vbOK", "vbCancel", "vbAbort", "vbRetry", "vbIgnore", "vbYes", "vbNo",
)
_CONST_COMPARE = (
    "vbBinaryCompare", "vbTextCompare", "vbDatabaseCompare",
    "vbUpperCase", "vbLowerCase", "vbProperCase",
    "vbWide", "vbNarrow", "vbKatakana", "vbHiragana",
    "vbUnicode", "vbFromUnicode",
)
_CONST_DATE = (
    "vbGeneralDate", "vbLongDate", "vbShortDate", "vbLongTime", "vbShortTime",
    "vbSunday", "vbMonday", "vbTuesday", "vbWednesday", "vbThursday",
    "vbFriday", "vbSaturday", "vbUseSystemDayOfWeek",
    "vbUseSystem", "vbFirstJan1", "vbFirstFourDays", "vbFirstFullWeek",
)
_CONST_VARTYPE = (
    "vbEmpty", "vbNull", "vbInteger", "vbLong", "vbSingle", "vbDouble",
    "vbCurrency", "vbDate", "vbString", "vbObject", "vbError", "vbBoolean",
    "vbVariant", "vbDataObject", "vbDecimal", "vbByte", "vbLongLong",
    "vbUserDefinedType", "vbArray",
)
_CONST_FILE = (
    "vbNormal", "vbReadOnly", "vbHidden", "vbSystem", "vbVolume",
    "vbDirectory", "vbArchive", "vbAlias",
)
_CONST_COLOR = (
    "vbBlack", "vbRed", "vbGreen", "vbYellow", "vbBlue", "vbMagenta",
    "vbCyan", "vbWhite",
)
_CONST_MISC = (
    "vbObjectError",
)

CONSTANTS = (_CONST_STRINGS + _CONST_MSGBOX + _CONST_COMPARE + _CONST_DATE
             + _CONST_VARTYPE + _CONST_FILE + _CONST_COLOR + _CONST_MISC)

# --------------------------------------------------------------------------
# 内建数据类型（`Dim x As <这里>`）
# --------------------------------------------------------------------------
TYPE_NAMES = (
    "Boolean", "Byte", "Currency", "Date", "Double", "Integer", "Long",
    "LongLong", "LongPtr", "Object", "Single", "String", "Variant",
)

# --------------------------------------------------------------------------
# 语言关键字 / 保留字（v62）
# --------------------------------------------------------------------------
# 按 VBA 官方关键字表整理，另补三个常写的（PtrSafe / Alias / Explicit）：
#   * PtrSafe / Alias —— 64 位 Declare 的修饰关键字，写 API 声明时要用；
#   * Explicit       —— Option Explicit（Excel 新模块默认带上这一行）。
# 与上面几组重合的（String / Date / Error / Erase / Get / Open …）由
# BUILTIN_KEYWORDS 的去重挡掉，不会重复进池。
KEYWORDS = (
    # ---- 过程、声明与可见性 ----
    "Sub", "Function", "Property", "Declare", "PtrSafe", "Alias", "Lib",
    "Public", "Private", "Friend", "Global", "Static", "Const", "Dim",
    "ReDim", "Preserve", "Implements", "Event", "WithEvents", "RaiseEvent",
    "Optional", "ParamArray", "ByVal", "ByRef", "Attribute", "Option",
    "Explicit", "Enum", "Type", "As", "New", "Me", "Is", "Like",
    "Set", "Let", "Call",

    # ---- 流程控制 ----
    "If", "Then", "Else", "ElseIf", "End", "Select", "Case", "For", "Each",
    "In", "To", "Step", "Next", "Do", "Loop", "While", "Wend", "Until",
    "Exit", "GoTo", "GoSub", "Return", "On", "Resume", "Stop", "With",
    "TypeOf", "Rem", "Debug",

    # ---- 字面量 / 空值 ----
    "True", "False", "Nothing", "Empty", "Null",

    # ---- 运算符 ----
    "And", "Or", "Not", "Xor", "Eqv", "Imp", "Mod", "AddressOf",
)

# --------------------------------------------------------------------------
# 对外导出
# --------------------------------------------------------------------------
# v70：函数与常量【分开】导出 —— 用户口径是"内置枚举常量不收，内建函数照常收"。
# 只拆导出、不动清单本身（见文件头 v70 说明）。
BUILTIN_FUNCTIONS = tuple(dict.fromkeys(FUNCTIONS))
BUILTIN_CONSTANTS = tuple(dict.fromkeys(CONSTANTS))

# 两组并集（静态全量清单）：保留给"想知道完整名单"的调用方（测试 / 诊断）。
# ⚠️ 收集侧**不要**直接用它 —— 走 vbe_bridge 的 ENABLE_VBA_CONSTANTS 裁剪。
BUILTIN_NAMES = tuple(dict.fromkeys(FUNCTIONS + CONSTANTS))

# 内建数据类型名（同样进候选池，供 `As |` 位置使用）。
BUILTIN_TYPE_NAMES = tuple(dict.fromkeys(TYPE_NAMES))

# 语言关键字（单独一组：开关独立、也便于测试里单独取舍）。
# 与 BUILTIN_NAMES 的同名项（Error / Erase / Get / Open …）由收集侧的统一去重挡掉。
BUILTIN_KEYWORDS = tuple(dict.fromkeys(KEYWORDS))

# 去重后的小写集合，供需要快速查询的调用方使用。
BUILTIN_LOWER = frozenset(n.lower() for n in BUILTIN_NAMES)
BUILTIN_TYPE_LOWER = frozenset(n.lower() for n in BUILTIN_TYPE_NAMES)
BUILTIN_KEYWORD_LOWER = frozenset(n.lower() for n in BUILTIN_KEYWORDS)
