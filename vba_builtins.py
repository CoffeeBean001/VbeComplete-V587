"""VBA 语言自带的名字：内建函数、常用内建常量、内建数据类型。

这些名字【不在用户的 VBE 工程代码里】—— 它们由 VBA 语言运行时提供，代码文本里
一个字都找不到（除非用户恰好调用过 MsgBox / Left 之类）。所以必须像组件名、窗体
控件名那样，作为一路独立的来源喂给候选池（见 vbe_bridge._collect_identifiers）。

收录范围（按用户口径："VBA 本身自带的函数"）：
  * FUNCTIONS  —— VBA 库的内建过程与函数（类型转换 / 字符串 / 日期 / 数组 / 数学 /
                  文件 IO / 交互 / 注册表 / 财务），含 Open、Close、Print 这类以
                  语句形式出现的 VBA 内建过程；
  * CONSTANTS  —— 常用的 vb* 内建常量（换行符、MsgBox 参数与返回值、StrConv、
                  日期常量、VarType、文件属性、颜色等）；
  * TYPE_NAMES —— VBA 内建数据类型（`Dim x As <这里>` 用得上）。

刻意【不收】的：
  * 语言关键字（If / For / Sub / Function / Dim / Set / New / Nothing …）——
    用户明确讨厌这类名字混进候选池（v59 修过"Function/If/With 混进候选池"这个
    bug），VBE 自身也不把它们当成员提示；
  * 歧义过大的语句名（Name / Width / Line / Spc / Tab）—— 它们在代码里更常作为
    变量名或属性名出现，提示出来容易误导；
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
# 对外导出
# --------------------------------------------------------------------------
# 进候选池的名字（函数 + 常量），保留规范大小写、顺序稳定。
BUILTIN_NAMES = tuple(dict.fromkeys(FUNCTIONS + CONSTANTS))

# 内建数据类型名（同样进候选池，供 `As |` 位置使用）。
BUILTIN_TYPE_NAMES = tuple(dict.fromkeys(TYPE_NAMES))

# 去重后的小写集合，供需要快速查询的调用方使用。
BUILTIN_LOWER = frozenset(n.lower() for n in BUILTIN_NAMES)
BUILTIN_TYPE_LOWER = frozenset(n.lower() for n in BUILTIN_TYPE_NAMES)
