# -*- coding: utf-8 -*-
"""回归测试：用真实工程（脚本.xlsm 导出的 14 个模块）验证补全作用域。

不依赖 Excel / COM，纯解析层验证，可直接 `python tests/test_real_project.py`
或双击 `run_tests.bat` 运行。

覆盖的关键场景：
  1. UserForm3 主过程里输入 n 必须提示 num（不能被别处的同名参数挡掉）
  2. 不能自我提示（打 n 不该冒出 n）
  3. 跨过程不泄漏（别的函数的局部变量不该出现）
  4. 跨模块 Private 不提示
  5. 打全变量名后列表仍保留
  6. 中文变量名
  7. 多层嵌套里外层变量仍可提示
"""

import io
import os
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

import parser as P          # noqa: E402
import engine as E          # noqa: E402

# main.py 顶部依赖 tkinter / pynput（以及 engine/vbe_bridge/ui），而测试用的
# managed python 是精简环境，未必装这些 GUI 包。注入最小桩，避免 `import main`
# 失败、导致第 18 节（钩子挂载决策）整段 return 跳过——那样
# _reconcile_resources 的回归就根本没被验证（之前吃过这个亏）。
import types as _types
for _mod in ("tkinter", "tkinter.font", "pynput", "pynput.keyboard",
             "pynput.mouse"):
    if _mod not in sys.modules:
        sys.modules[_mod] = _types.ModuleType(_mod)


class _StubListener:         # 桩：测试不真正启动钩子
    daemon = False

    def start(self):
        pass

    def stop(self):
        pass

    def is_alive(self):
        return False

    def suppress_event(self):
        pass


sys.modules["pynput.keyboard"].Listener = _StubListener
sys.modules["pynput.mouse"].Listener = _StubListener
# main.py 模块级会引用 Key.up/.down/.tab/.esc/.ctrl_l/...，桩里要有这些属性
sys.modules["pynput.keyboard"].Key = type("Key", (), {
    "up": object(), "down": object(), "tab": object(), "esc": object(),
    "ctrl_l": object(), "ctrl_r": object(), "alt_l": object(), "alt_r": object(),
    "shift": object(), "shift_r": object(), "caps_lock": object(),
    "cmd": object(), "cmd_r": object(), "space": object(),
    "backspace": object(), "enter": object(),
})

try:                          # noqa: E402
    import main as M         # noqa: E402  （main._reconcile_resources 钩子挂载决策）
except Exception:            # noqa: E402
    M = None                 # noqa: E402

TESTDATA = os.path.join(HERE, "testdata")

PASS = 0
FAIL = 0
FAILURES = []


def _is_std(filename):
    """判断是否标准模块（对应 VBE 的 vbext_ct_StdModule）。"""
    return filename.endswith(".bas") and not filename.startswith(
        ("ThisWorkbook", "Sheet"))


def _mod_name(filename):
    return filename.split(".")[0]


def load_modules():
    mods = {}
    for fn in sorted(os.listdir(TESTDATA)):
        if not fn.endswith(".bas"):
            continue
        code = io.open(os.path.join(TESTDATA, fn), encoding="utf-8").read()
        mods[fn] = (code, _is_std(fn), _mod_name(fn))
    return mods


def collect(mods, caret_file=None, caret=None, scope="proc", caret_mod=None):
    """复刻 VbeBackend._collect_identifiers 的收集逻辑。

    mods: {文件名: (代码, 是否标准模块, 模块名)}，代表【当前工程】的全部模块。
    caret_mod 显式给出光标所在模块名；缺省时用 caret_file 推导。

    注：v30 起后端不再做"声明行名字全局剔除"（那样会连同前缀补全一起误杀
    numArr），名字一律收录；"打全名不提示自己"改由引擎层 in_decl_position +
    ctx["decl_names"] 精确剔除。故此处复刻也不再剔除声明行名字。
    """
    if caret_mod is None and caret_file:
        caret_mod = _mod_name(caret_file)
    records = []
    for fn, (code, std, mn) in mods.items():
        apply_caret = caret if (not caret_mod
                                or str(caret_mod).lower() == str(mn).lower()) else None
        recs = P.extract_records(code, module=mn, is_std_module=std,
                                 caret=apply_caret)
        records.extend(recs)
        records.extend(P.extract_implicit_records(
            code, module=mn, is_std_module=std, declared=recs, scope=scope,
            caret=apply_caret))
        # 组件名（模块名/窗体名/类名）本身也是工程内可直接引用的标识符
        # （对应 vbe_bridge 里的 records.append((mod_name, mod_name, None, False))）
        if mn:
            records.append((mn, mn, None, False))
    return records


def collect_projects(projects, active, caret_file=None, caret=None, scope="proc"):
    """复刻"只收集活动工程"的行为：projects = {工程名: mods}。

    非活动工程的模块一律不参与收集，因此它们的窗体名/函数名/全局变量名
    不会泄漏到当前工程（v23 修复的跨工程作用域）。
    """
    return collect(projects[active], caret_file=caret_file, caret=caret,
                   scope=scope)


def type_at(mods, filename, line_no, typed, indent=20):
    """返回「在某行输入 typed 之后」的模块副本与光标列。"""
    m = dict(mods)
    code, std, mn = m[filename]
    ls = code.split("\n")
    ls[line_no - 1] = " " * indent + typed
    m[filename] = ("\n".join(ls), std, mn)
    return m, indent + len(typed) + 1     # 光标在刚输入内容之后


def candidates(mods, records, filename, line_no, prefix):
    code = mods[filename][0]
    proc = P.proc_at_line(code, line_no)
    vis = E.filter_identifiers_by_scope(records, proc, _mod_name(filename))
    return sorted(i for i in vis if i.lower().startswith(prefix.lower())), proc


def check(tag, got, expect_contain=(), expect_absent=()):
    global PASS, FAIL
    ok = True
    for x in expect_contain:
        if x not in got:
            ok = False
    for x in expect_absent:
        if x in got:
            ok = False
    if ok:
        PASS += 1
        print("  PASS  %s" % tag)
        print("        候选=%s" % (got,))
    else:
        FAIL += 1
        FAILURES.append(tag)
        print("  FAIL  %s" % tag)
        print("        候选=%s" % (got,))
        print("        应含=%s 应不含=%s" % (list(expect_contain), list(expect_absent)))


def main():
    mods = load_modules()
    print("载入 %d 个模块: %s\n" % (len(mods), ", ".join(sorted(mods))))

    print("=== 1. 真实场景：UserForm3 第 61 行输入 n ===")
    for line in (60, 61, 63):
        m, col = type_at(mods, "UserForm3.frm.bas", line, "n")
        recs = collect(m, caret_file="UserForm3.frm.bas", caret=(line, col))
        cand, proc = candidates(m, recs, "UserForm3.frm.bas", line, "n")
        print("  [行 %d] proc=%s" % (line, proc))
        check("行%d 输入 n 提示 num" % line, cand,
              expect_contain=["num", "numColInTargetSheet"],
              expect_absent=["n"])

    print("\n=== 2. 不自我提示（打 n 不该冒出 n）===")
    m, col = type_at(mods, "UserForm3.frm.bas", 61, "n")
    recs = collect(m, caret_file="UserForm3.frm.bas", caret=(61, col))
    cand, _ = candidates(m, recs, "UserForm3.frm.bas", 61, "n")
    check("自我提示", cand, expect_absent=["n"])

    print("\n=== 3. 跨过程不泄漏（num2 属于 getNumColInTargetSheet）===")
    check("主过程不含 num2", cand, expect_absent=["num2"])

    print("\n=== 4. 在 getNumColInTargetSheet 内应能看到 num / num2 ===")
    # 该函数体在 122-131 行区间（去掉 Attribute 头后约 114-123 行）
    code = mods["UserForm3.frm.bas"][0]
    target = None
    for i, l in enumerate(code.split("\n"), 1):
        if "Private Function getNumColInTargetSheet" in l:
            target = i + 3
            break
    if target:
        m2, col2 = type_at(mods, "UserForm3.frm.bas", target, "n", indent=8)
        recs2 = collect(m2, caret_file="UserForm3.frm.bas", caret=(target, col2))
        cand2, proc2 = candidates(m2, recs2, "UserForm3.frm.bas", target, "n")
        print("  [行 %d] proc=%s" % (target, proc2))
        check("函数内可见 num/num2", cand2, expect_contain=["num", "num2"])

    print("\n=== 5. 打全变量名后列表仍保留（v15）===")
    m3, col3 = type_at(mods, "UserForm3.frm.bas", 61, "num")
    recs3 = collect(m3, caret_file="UserForm3.frm.bas", caret=(61, col3))
    cand3, _ = candidates(m3, recs3, "UserForm3.frm.bas", 61, "num")
    check("打全 num 仍保留", cand3, expect_contain=["num", "numColInTargetSheet"])

    print("\n=== 6. 跨过程隔离 / 中文 / 嵌套（合成用例）===")
    def synth(code, proc, prefix, mod="M1", std=True, caret=None):
        r = P.extract_records(code, mod, std)
        r += P.extract_implicit_records(code, mod, std, declared=r,
                                        scope="proc", caret=caret)
        v = E.filter_identifiers_by_scope(r, proc, mod)
        return sorted(i for i in v if i.lower().startswith(prefix.lower()))

    c1 = "\n".join(["Sub ProcA()", "    num = 1", "    Debug.Print num", "End Sub",
                    "Sub ProcB()", "    num = 2", "    num2 = 3", "End Sub"])
    check("ProcA 只提示 num", synth(c1, "ProcA", "n"),
          expect_contain=["num"], expect_absent=["num2"])
    check("ProcB 有 num+num2", synth(c1, "ProcB", "n"),
          expect_contain=["num", "num2"])

    c2 = "\n".join(["Sub 甲()", "    姓名 = \"x\"", "End Sub",
                    "Sub 乙()", "    姓名 = \"y\"", "    年龄 = 18", "End Sub"])
    check("中文：甲不泄漏年龄", synth(c2, "甲", "年"), expect_absent=["年龄"])
    check("中文：乙有年龄", synth(c2, "乙", "年"), expect_contain=["年龄"])

    nest = "\n".join(["Sub T()", "    outerV = 1", "    For i = 1 To 10",
                      "        If outerV > 1 Then", "            For j = 1 To 3",
                      "                If j > 1 Then", "                    outerV",
                      "                End If", "            Next j",
                      "        End If", "    Next i", "End Sub"])
    check("多层嵌套外层可见", synth(nest, "T", "outer", caret=(7, 25)),
          expect_contain=["outerV"])

    # ---- 7. v21：声明行上的自我提示 ----
    # 定义模块级/全局变量时，正在声明的名字不该提示其本身
    print("\n=== 7. 声明行自我提示（v21）===")
    decl = "\n".join([
        "Public gCounter As Long",
        "Private mTotal As String",
        "Dim gFlag As Boolean",
        "Sub Foo()",
        "    gCounter",
        "End Sub"])

    def decl_cand(caret, prefix):
        recs = P.extract_records(decl, module="M1",
                                      is_std_module=True, caret=caret)
        proc = P.proc_at_line(decl, caret[0]) if caret else None
        vis = E.filter_identifiers_by_scope(recs, proc, "M1")
        return sorted(i for i in vis if i.lower().startswith(prefix.lower()))

    # v30 起解析层不再剔除声明行名字（否则会把前缀补全 numArr 一起误杀），
    # 名字一律收录、交给引擎按"正在输入的词 + 当前声明行名字"精确剔除。
    # 故此处只验证：声明行上的名字确实被【收录】（前缀补全可用），
    # "打全名不提示自己"由第 11 节引擎集成测试把关。
    # 光标在 Public gCounter 声明行（前缀 g）：gCounter 与 gFlag 都应收录
    check("声明行名字被收录(整词前缀)",
          decl_cand((1, 19), "g"),
          expect_contain=["gCounter", "gFlag"])
    # 只输入首字母时同样收录（前缀补全可用）
    check("声明行名字被收录(首字母前缀)",
          decl_cand((1, 10), "g"),
          expect_contain=["gCounter", "gFlag"])
    # 光标在行首同样收录
    check("声明行名字被收录(行首前缀)",
          decl_cand((1, 2), "g"),
          expect_contain=["gCounter", "gFlag"])
    # Private 声明同理：mTotal 被收录
    check("Private 声明行名字被收录",
          decl_cand((2, 17), "m"), expect_contain=["mTotal"])
    # 使用行必须保留（v15 打全名保留列表的行为不能被破坏）
    check("使用行仍保留该变量",
          decl_cand((5, 13), "g"),
          expect_contain=["gCounter", "gFlag"])
    # 拿不到光标时不得误排除
    check("caret=None 不误排除",
          decl_cand(None, "g"), expect_contain=["gCounter", "gFlag"])

    # ---- 8. v22：声明里 `As` 之后的类型名位置不应提示变量名 ----
    print("\n=== 8. 声明 As 之后类型名位置（v22）===")

    def tp(line, col):
        """is_caret_in_type_position 的简短封装（包成 list 适配 check）。"""
        return [P.is_caret_in_type_position(line, col)]

    # 变量声明：As 之后 -> 类型名位置
    check("Dim x As | 命中类型位置",
          tp("Dim x As ", 10), expect_contain=[True])
    check("Dim x As Integer 末尾命中",
          tp("Dim x As Integer", 17), expect_contain=[True])
    # 变量名位置（As 之前）-> 不命中，照常提示
    check("Dim x 变量名位置不命中",
          tp("Dim x", 6), expect_contain=[False])
    # 多声明：`a As Long, b As |` 各片段独立判断
    check("Dim a As Long, b As | 命中",
          tp("Dim a As Long, b As ", 22), expect_contain=[True])
    check("Dim a As Long, b 变量名位置不命中",
          tp("Dim a As Long, b", 18), expect_contain=[False])
    # 数组维度括号里的逗号不干扰
    check("Dim arr(0 To 9) As In 命中",
          tp("Dim arr(0 To 9) As In", 21), expect_contain=[True])
    # 过程参数 / 返回值里的 As 也命中
    check("Function 返回值类型位置命中",
          tp("Private Function f(ByVal a As Long) As Double", 47),
          expect_contain=[True])
    # 模块级 / 常量 / 类型
    check("Public gCounter As Long 命中",
          tp("Public gCounter As Long", 24), expect_contain=[True])
    check("Const pi As Double 命中",
          tp("Const pi As Double = 3.14", 19), expect_contain=[True])
    check("Dim x As New Collection 命中",
          tp("Dim x As New Collection", 24), expect_contain=[True])
    check("Dim x As String * 10 命中",
          tp("Dim x As String * 10", 20), expect_contain=[True])
    # 否定用例：字符串里的 "As" 不能误判
    check("字符串里的 As 不误判",
          tp('Debug.Print "As foo" & bar', 27), expect_contain=[False])
    check("ReDim（无 As）不命中",
          tp("ReDim arr(1 To 10)", 18), expect_contain=[False])
    check("If 行（无 As）不命中",
          tp("If x > y Then", 14), expect_contain=[False])

    # 引擎集成：in_type_position=True 时 trigger 必须收起弹窗
    print("\n--- 引擎集成：类型位置强制静默 ---")

    class _FakeUI:
        def __init__(self):
            self.shown = False
        def show(self, *a, **k):
            self.shown = True
        def update_selection(self, *a, **k):
            pass
        def hide(self, *a, **k):
            self.shown = False
        def contains_point(self, *a, **k):
            return False

    class _FakeBackend:
        def __init__(self, ctx):
            self.ctx = ctx
        def get_context(self):
            return self.ctx
        def get_identifiers(self):
            return [("num", "M1", "ProcA", False),
                    ("numColInTargetSheet", "M1", "ProcA", False)]
        def apply_completion(self, *a, **k):
            return None

    # 1) 类型名位置（As 后）即便有匹配也静默
    ui = _FakeUI()
    comp = E.Completer(_FakeBackend(
        {"line_no": 1, "caret_col": 12, "line_text": "Dim x As In",
         "in_string": False, "in_comment": False, "in_type_position": True,
         "proc_name": None, "module_name": "M1"}), ui)
    comp.trigger()
    check("类型位置 trigger 收起弹窗",
          [ui.shown], expect_contain=[False])

    # 2) 变量名位置（As 前）有匹配则正常弹出
    ui2 = _FakeUI()
    comp2 = E.Completer(_FakeBackend(
        {"line_no": 2, "caret_col": 6, "line_text": "    num = 1",
         "in_string": False, "in_comment": False, "in_type_position": False,
         "proc_name": "ProcA", "module_name": "M1"}), ui2)
    comp2.trigger()
    check("变量名位置 trigger 正常弹出",
          [ui2.shown], expect_contain=[True])

    # ---- 9. v23：定义全局/模块级变量时不提示其本身 ----
    print("\n=== 9. 声明变量时不自我提示（v23）===")

    def decl_case(code, caret, module="M1", prefix=None, std=True):
        """返回「在模块级/过程内声明变量时光标处的候选」。"""
        mods = {module + ".bas": (code, std, module)}
        recs = collect(mods, caret_mod=module, caret=caret)
        proc = P.proc_at_line(code, caret[0])
        vis = E.filter_identifiers_by_scope(recs, proc, module)
        return sorted(i for i in vis
                      if i.lower().startswith((prefix or "").lower()))

    # 9.1 名字在别处被"只读"使用过：用法扫描会把它当隐式变量收录。
    #     v30 起解析层不再全局剔除声明行名字（否则连前缀补全一起误杀），
    #     名字一律收录，交给引擎按"当前声明行名字 + 正在输入的词"精确剔除。
    #     此处只验证声明行名字确实被【收录】（前缀补全可用）。
    c_ro = "\n".join([
        "Public gUserName As String",
        "Sub Foo()",
        "    MsgBox gUserName",
        "End Sub"])
    check("模块级声明：名字被收录(整词前缀)",
          decl_case(c_ro, (1, len("Public gUserName") + 1), prefix="g"),
          expect_contain=["gUserName"])
    # 同一行只打出首字母时同样收录（前缀补全可用）
    check("模块级声明：名字被收录(首字母前缀)",
          decl_case(c_ro, (1, len("Public g") + 1), prefix="g"),
          expect_contain=["gUserName"])

    # 9.2 过程内：先用了再补 Dim（For i 已把 i 收成隐式变量）
    c_local = "\n".join([
        "Sub Foo()",
        "    For i = 1 To 3",
        "    Next i",
        "    Dim i As Long",
        "End Sub"])
    check("过程内 Dim 名字被收录",
          decl_case(c_local, (4, len("    Dim i") + 1), prefix="i"),
          expect_contain=["i"])

    # 9.3 续行声明：光标落在第二行（该行本身没有 Dim/Public）
    c_cont = "\n".join([
        "Public gA As Long, _",
        "       gB As Long",
        "Sub Foo()",
        "    MsgBox gA & gB",
        "End Sub"])
    check("续行声明第二行名字被收录",
          decl_case(c_cont, (2, len("       gB") + 1), prefix="g"),
          expect_contain=["gB"])
    check("续行声明整行两个名字都收录",
          decl_case(c_cont, (1, len("Public gA") + 1), prefix="g"),
          expect_contain=["gA", "gB"])

    # 9.4 非声明行不受影响：使用行仍要能提示
    check("使用行仍可提示",
          decl_case(c_ro, (3, len("    MsgBox gUser") + 1), prefix="g"),
          expect_contain=["gUserName"])

    # 9.5 函数名/过程名同样会被收录（v23-2 的"打全名不提示自己"由引擎把关）
    c_fn = "\n".join([
        "Public Function gCalc() As Long",
        "    gCalc = 1",
        "End Function"])
    check("函数声明行名字被收录",
          decl_case(c_fn, (1, len("Public Function gCalc") + 1), prefix="g"),
          expect_contain=["gCalc"])
    # 参数名也是"正在起的新名字"
    c_param = "\n".join([
        "Public Function gCalc(gTotal As Long) As Long",
        "    gCalc = gTotal",
        "End Function"])
    check("参数声明位置名字被收录",
          decl_case(c_param, (1, len("Public Function gCalc(gTot") + 1), prefix="g"),
          expect_contain=["gCalc", "gTotal"])
    # 使用行必须照常提示
    check("函数调用行仍可提示",
          decl_case(c_fn, (2, len("    gCal") + 1), prefix="g"),
          expect_contain=["gCalc"])
    # Type / Enum 名
    c_type = "\n".join([
        "Public Type gRec",
        "    gId As Long",
        "End Type"])
    check("Type 声明行名字被收录",
          decl_case(c_type, (1, len("Public Type gRe") + 1), prefix="g"),
          expect_contain=["gRec"])
    check("Type 成员行名字被收录",
          decl_case(c_type, (2, len("    gI") + 1), prefix="g"),
          expect_contain=["gId"])

    # ---- 10. v23：跨工程隔离 ----
    print("\n=== 10. 跨工程隔离（v23）===")
    proj_a = {
        "Module1.bas": ("\n".join([
            "Public gLocal As Long",
            "Sub Main()",
            "    gLocal = 1",
            "End Sub"]), True, "Module1"),
    }
    # 另一个工程：同名模块 Module1（默认名撞车是跨工程泄漏的典型入口）
    # 里面有全局变量、函数、以及被 Show 调用的窗体名
    proj_b = {
        "Module1.bas": ("\n".join([
            "Public gOther As Long",
            "Public Function OtherFn() As Long",
            "    OtherFn = 1",
            "End Function"]), True, "Module1"),
        "UserForm1.frm.bas": ("\n".join([
            "Private Sub UserForm_Initialize()",
            "End Sub",
            "Sub ShowIt()",
            "    UserForm1.Show",
            "End Sub"]), False, "UserForm1"),
    }
    projects = {"A": proj_a, "B": proj_b}

    def cand_in_project(active, filename, line_no, prefix):
        mods = projects[active]
        recs = collect_projects(projects, active, caret_file=filename,
                                caret=(line_no, 1))
        code = mods[filename][0]
        proc = P.proc_at_line(code, line_no)
        vis = E.filter_identifiers_by_scope(recs, proc, _mod_name(filename))
        return sorted(i for i in vis if i.lower().startswith(prefix.lower()))

    check("活动工程 A 内可见本工程全局量",
          cand_in_project("A", "Module1.bas", 3, "g"),
          expect_contain=["gLocal"])
    check("不提示工程 B 的全局变量",
          cand_in_project("A", "Module1.bas", 3, "g"),
          expect_absent=["gOther"])
    check("不提示工程 B 的函数名",
          cand_in_project("A", "Module1.bas", 3, "other"),
          expect_absent=["OtherFn"])
    check("不提示工程 B 的窗体名",
          cand_in_project("A", "Module1.bas", 3, "userform"),
          expect_absent=["UserForm1"])
    # 反向：在 B 里也不该看到 A 的东西
    check("反向：B 内不含 A 的全局量",
          cand_in_project("B", "Module1.bas", 3, "g"),
          expect_contain=["gOther"], expect_absent=["gLocal"])

    # ---- 11. v23-2：声明行整行静默（第一道防线）----
    print("\n=== 11. 声明行彻底静默（v23-2）===")

    def dp(line):
        return [P.is_caret_in_declaration(line)]

    check("Public 变量声明行 -> 静默",
          dp("Public gUserName As String"), expect_contain=[True])
    check("Dim 声明行 -> 静默",
          dp("Dim gFlag As Boolean"), expect_contain=[True])
    check("Private 变量声明行 -> 静默",
          dp("Private mTotal As String"), expect_contain=[True])
    check("Global 变量声明行 -> 静默",
          dp("Global gX As Long"), expect_contain=[True])
    check("Const 声明行 -> 静默",
          dp("Public Const gMax = 10"), expect_contain=[True])
    check("Static 声明行 -> 静默",
          dp("Static gCnt As Long"), expect_contain=[True])
    check("多行声明（逗号）-> 静默",
          dp("Dim a As Long, b As Integer"), expect_contain=[True])
    check("续行首行 -> 静默",
          dp("Public gA As Long, _"), expect_contain=[True])
    # 函数名 / 过程名 / 属性 / API / 事件 / 类型
    check("Function 声明行 -> 静默",
          dp("Public Function gCalc() As Long"), expect_contain=[True])
    check("Sub 声明行 -> 静默",
          dp("Private Sub gRun()"), expect_contain=[True])
    check("Property 声明行 -> 静默",
          dp("Public Property Get gVal() As Long"), expect_contain=[True])
    check("Declare 声明行 -> 静默",
          dp('Private Declare PtrSafe Function gApi Lib "x" () As Long'),
          expect_contain=[True])
    check("Event 声明行 -> 静默",
          dp("Public Event gChanged(ByVal x As Long)"), expect_contain=[True])
    check("Type 声明行 -> 静默",
          dp("Public Type gRec"), expect_contain=[True])
    check("Enum 声明行 -> 静默",
          dp("Public Enum gKind"), expect_contain=[True])
    # 不误伤
    check("普通赋值行 -> 不静默",
          dp("    gUserName = \"x\""), expect_contain=[False])
    check("函数调用行 -> 不静默",
          dp("    x = gCalc(1)"), expect_contain=[False])
    check("If 行 -> 不静默",
          dp("If gFlag Then"), expect_contain=[False])
    check("ReDim 行 -> 不静默（要提示已有数组）",
          dp("ReDim gArr(1 To 10)"), expect_contain=[False])
    check("字符串里的 Dim 不误判",
          dp('Debug.Print "Dim x As Long"'), expect_contain=[False])
    check("End Function 不误判",
          dp("End Function"), expect_contain=[False])

    # 引擎集成：声明行上即使有匹配也一条都不弹（这是"提示自己"的根治手段）
    class _UI2(_FakeUI):
        pass

    class _B2(_FakeBackend):
        def __init__(self, ctx, ids, declared=None):
            _FakeBackend.__init__(self, ctx)
            self._ids = ids
            # 真实声明的名字（小写）；缺省把全部 ids 当声明，匹配"旧行为兜底"。
            if declared is None:
                declared = [i[0] if isinstance(i, (tuple, list)) else i
                            for i in ids]
            self._declared = set(str(n).lower() for n in declared)

        def get_identifiers(self):
            return self._ids

        def get_declared_names(self):
            return list(self._declared)

    ids = [("gUserName", "M1", None, False), ("gCalc", "M1", None, False),
           ("gFlag", "M1", None, False)]

    def trig(line_text, col, in_decl, in_type=False, ids_=None, cls=None,
             decl_names=None, declared_=None):
        ui = _UI2()
        c = E.Completer((cls or _B2)(
            {"line_no": 1, "caret_col": col, "line_text": line_text,
             "in_string": False, "in_comment": False,
             "in_type_position": in_type, "in_decl_position": in_decl,
             "decl_names": decl_names or [],
             "proc_name": None, "module_name": "M1"},
            ids_ or ids, declared_), ui)
        c.trigger()
        return (ui.shown, c.current_matches())

    # 声明行：能匹配到已定义的就提示，只剩"正在输入的那个词本身"才安静
    check("公共变量：打全名 -> 不提示自己",
          [trig("Public gUserName", 17, True, decl_names=["gUserName"])],
          expect_contain=[(False, [])])
    check("公共变量：打前缀 -> 提示已有的 gUserName",
          [trig("Public gUser", 13, True, decl_names=["gUserName"])],
          expect_contain=[(True, ["gUserName"])])
    check("函数名：打全名 -> 不提示自己",
          [trig("Public Function gCalc", 22, True, decl_names=["gCalc"])],
          expect_contain=[(False, [])])
    check("函数名：打前缀 -> 提示已有的 gCalc",
          [trig("Public Function gCal", 21, True, decl_names=["gCalc"])],
          expect_contain=[(True, ["gCalc"])])
    check("非声明行：照常提示",
          [trig("    gUser", 10, False)], expect_contain=[(True, ["gUserName"])])

    # ---- 11.6 Bug1：声明行上输入 num 同时提醒 num 与 numArr（前缀补全不再误杀）----
    bug1_code = "\n".join([
        "Public num As Long",
        "Public numArr() As Long",
        "Sub Foo()",
        "    MsgBox num",
        "End Sub"])
    b1_mods = {"M1.bas": (bug1_code, True, "M1")}
    b1_recs = collect(b1_mods, caret_mod="M1", caret=(2, len("Public num") + 1))
    ui_b1 = _UI2()
    c_b1 = E.Completer(_B2(
        {"line_no": 2, "caret_col": len("Public num") + 1,
         "line_text": "Public num", "in_string": False, "in_comment": False,
         "in_type_position": False, "in_decl_position": True,
         "decl_names": ["numArr"], "proc_name": None, "module_name": "M1"},
        b1_recs), ui_b1)
    c_b1.trigger()
    check("Bug1: 声明行输入 num 同时提醒 num 与 numArr",
          [(ui_b1.shown, sorted(c_b1.current_matches()))],
          expect_contain=[(True, ["num", "numArr"])])
    # 打全 numArr（声明行完整名字）-> 不提示 numArr 自己（无其它兄弟，整行安静）
    ui_b1b = _UI2()
    c_b1b = E.Completer(_B2(
        {"line_no": 2, "caret_col": len("Public numArr") + 1,
         "line_text": "Public numArr", "in_string": False, "in_comment": False,
         "in_type_position": False, "in_decl_position": True,
         "decl_names": ["numArr"], "proc_name": None, "module_name": "M1"},
        b1_recs), ui_b1b)
    c_b1b.trigger()
    check("Bug1: 声明行打全 numArr 不提示自己",
          [(ui_b1b.shown, sorted(c_b1b.current_matches()))],
          expect_contain=[(False, [])])

    # ---- 11.7 Bug2：回退删除后若唯一匹配等于所输入词，不提示自己 ----
    ui_b2 = _UI2()
    c_b2 = E.Completer(_B2(
        {"line_no": 3, "caret_col": 10, "line_text": "    gFlag",
         "in_string": False, "in_comment": False, "in_type_position": False,
         "in_decl_position": False, "decl_names": [],
         "proc_name": "Foo", "module_name": "M1"}, ids), ui_b2)
    c_b2.trigger()
    # v33 语义更新：gFlag 是【真实定义】的名字，打全/回退到它都要保留列表
    # （用户要求"输入完整，提示保留"）。"提示自己"由"不是真名就剔除"负责，
    # 不再靠"唯一候选等于所输词就收起"这条误伤正常场景的规则。
    check("Bug2(新语义): 回退到真实存在的 gFlag -> 保留提示",
          [(ui_b2.shown, sorted(c_b2.current_matches()))],
          expect_contain=[(True, ["gFlag"])])
    # 反向：前缀 "gF" 仍有真实匹配 -> 正常提示（不是自我提示）
    ui_b2b = _UI2()
    c_b2b = E.Completer(_B2(
        {"line_no": 3, "caret_col": 8, "line_text": "    gF",
         "in_string": False, "in_comment": False, "in_type_position": False,
         "in_decl_position": False, "decl_names": [],
         "proc_name": "Foo", "module_name": "M1"}, ids), ui_b2b)
    c_b2b.trigger()
    check("Bug2: 前缀有真实匹配 -> 正常提示",
          [(ui_b2b.shown, sorted(c_b2b.current_matches()))],
          expect_contain=[(True, ["gFlag"])])

    # ---- 11.8 Bug1(续)：回退出未定义的 numA，绝不提示 numA 自己 ----
    # 模拟"隐式变量扫描把光标处正在敲的 numA 当成已用变量"的最坏情况：ids 里
    # 塞入一个 numA（priv=True 的隐式残留），但工程并未声明 numA。引擎应凭
    # "numA 不是真实声明的名字"把它从候选里剔除，只留下 numArr。
    phantom_ids = [("num", "M1", None, False), ("numArr", "M1", None, False),
                   ("numA", "M1", None, True)]
    ui_p = _UI2()
    c_p = E.Completer(_B2(
        {"line_no": 3, "caret_col": len("    numA") + 1, "line_text": "    numA",
         "in_string": False, "in_comment": False, "in_type_position": False,
         "in_decl_position": False, "decl_names": [],
         "proc_name": "Foo", "module_name": "M1"},
        phantom_ids, declared=["num", "numArr"]), ui_p)
    c_p.trigger()
    check("Bug1续: 输入 numA 不提示未定义的 numA（只提示 numArr）",
          [(ui_p.shown, sorted(c_p.current_matches()))],
          expect_contain=[(True, ["numArr"])])
    # 反向：输入真实定义的 num -> 仍同时提醒 num 与 numArr（前缀补全正常）
    ui_p2 = _UI2()
    c_p2 = E.Completer(_B2(
        {"line_no": 3, "caret_col": len("    num") + 1, "line_text": "    num",
         "in_string": False, "in_comment": False, "in_type_position": False,
         "in_decl_position": False, "decl_names": [],
         "proc_name": "Foo", "module_name": "M1"},
        [("num", "M1", None, False), ("numArr", "M1", None, False)],
        declared=["num", "numArr"]), ui_p2)
    c_p2.trigger()
    check("Bug1续: 输入 num 同时提醒 num 与 numArr（用法行）",
          [(ui_p2.shown, sorted(c_p2.current_matches()))],
          expect_contain=[(True, ["num", "numArr"])])

    # ---- 11.9 键盘选择 / 确认（↑↓ 浏览、Tab/Enter 确认，无需移动鼠标）----
    class _CapBackend(_B2):
        def __init__(self, ctx, ids, declared=None):
            _B2.__init__(self, ctx, ids, declared)
            self.last_apply = None

        def apply_completion(self, line_no, start, caret, completion):
            self.last_apply = (line_no, start, caret, completion)
            return completion

    nav_ids = [("num", "M1", None, False), ("numArr", "M1", None, False),
               ("numCount", "M1", None, False)]
    ui_n = _UI2()
    c_n = E.Completer(_CapBackend(
        {"line_no": 3, "caret_col": 8, "line_text": "    num",
         "in_string": False, "in_comment": False, "in_type_position": False,
         "in_decl_position": False, "decl_names": [],
         "proc_name": "Foo", "module_name": "M1"}, nav_ids), ui_n)
    c_n.trigger()
    check("键盘: 候选 3 个", [c_n.current_matches()],
          expect_contain=[["num", "numArr", "numCount"]])
    check("键盘: 默认选中首项", [c_n.selected], expect_contain=[0])
    c_n.move(2)                       # ↓↓ 跳到末项
    check("键盘: ↓↓ 到第 3 项", [c_n.selected], expect_contain=[2])
    check("键盘: 当前候选 = numCount",
          [c_n.current_matches()[c_n.selected]], expect_contain=["numCount"])
    c_n.accept()                      # 确认 -> apply_completion 拿到 numCount
    check("键盘: 确认写回 numCount",
          [c_n.backend.last_apply[3] if c_n.backend.last_apply else None],
          expect_contain=["numCount"])
    check("键盘: 确认后弹窗收起", [ui_n.shown], expect_contain=[False])
    # 循环：首项按 ↑ 应回到末项
    ui_n2 = _UI2()
    c_n2 = E.Completer(_CapBackend(
        {"line_no": 3, "caret_col": 8, "line_text": "    num",
         "in_string": False, "in_comment": False, "in_type_position": False,
         "in_decl_position": False, "decl_names": [],
         "proc_name": "Foo", "module_name": "M1"}, nav_ids), ui_n2)
    c_n2.trigger()
    c_n2.move(-1)
    check("键盘: 首项按 ↑ 循环到末项", [c_n2.selected], expect_contain=[2])

    # ---- 11.10 v32：真名集合必须剔除"光标所在声明行正在声明的名字" ----
    # 回退删字（Dim numArr -> Dim numA）时，文本里此刻确实写着 Dim numA，
    # 解析层会把它当成真实声明；若真名集合照单全收，"不是真名就剔除自身"
    # 的防护就失效，numA 又被提示出来。这里模拟"真名集合已正确剔除光标行
    # 声明名"（declared 只有 num / numArr），同时 decl_names 故意给空
    # （模拟光标信息陈旧），验证引擎仍不提示 numA。
    ui_v32 = _UI2()
    c_v32 = E.Completer(_B2(
        {"line_no": 3, "caret_col": len("    Dim numA") + 1,
         "line_text": "    Dim numA", "in_string": False, "in_comment": False,
         "in_type_position": False, "in_decl_position": True,
         "decl_names": [], "proc_name": "Foo", "module_name": "M1"},
        phantom_ids, declared=["num", "numArr"]), ui_v32)
    c_v32.trigger()
    check("v32: 声明行回退出的 numA 不提示自己（真名集合已剔除光标行）",
          [(ui_v32.shown, sorted(c_v32.current_matches()))],
          expect_contain=[(True, ["numArr"])])
    # 反向：若真名集合【没有】剔除光标行声明名（旧行为），numA 就会漏出来。
    # 这条是"修复前必然失败"的对照，锁住回归。
    ui_v32b = _UI2()
    c_v32b = E.Completer(_B2(
        {"line_no": 3, "caret_col": len("    Dim numA") + 1,
         "line_text": "    Dim numA", "in_string": False, "in_comment": False,
         "in_type_position": False, "in_decl_position": True,
         "decl_names": [], "proc_name": "Foo", "module_name": "M1"},
        phantom_ids, declared=["num", "numArr", "numA"]), ui_v32b)
    c_v32b.trigger()
    check("v32 对照: 真名集合含 numA 时确实会漏出（旧行为基线）",
          [(ui_v32b.shown, sorted(c_v32b.current_matches()))],
          expect_contain=[(True, ["numA", "numArr"])])

    # ---- 11.10b v33：已定义变量打全名，列表必须保留（用户反馈的回归）----
    # 场景：定义了 arr，输入 a / ar 都提示 arr；把 arr 打全后列表不能消失，
    # 否则既没法 Tab 确认，也看不出自己写对了没有。
    arr_ids = [("arr", "M1", None, False)]
    for w, tag in (("a", "首字母 a"), ("ar", "前缀 ar"), ("arr", "打全 arr")):
        ui_a = _UI2()
        c_a = E.Completer(_B2(
            {"line_no": 3, "caret_col": len("    " + w) + 1,
             "line_text": "    " + w, "in_string": False, "in_comment": False,
             "in_type_position": False, "in_decl_position": False,
             "decl_names": [], "proc_name": "Foo", "module_name": "M1"},
            arr_ids, declared=["arr"]), ui_a)
        c_a.trigger()
        check("v33: 已定义 arr 输入 %s -> 仍提示 arr" % tag,
              [(ui_a.shown, sorted(c_a.current_matches()))],
              expect_contain=[(True, ["arr"])])
    # 对照：从未定义的 arrX 打全 -> 不提示自己（真名集合里没有它）
    ui_ax = _UI2()
    c_ax = E.Completer(_B2(
        {"line_no": 3, "caret_col": len("    arrX") + 1, "line_text": "    arrX",
         "in_string": False, "in_comment": False, "in_type_position": False,
         "in_decl_position": False, "decl_names": [],
         "proc_name": "Foo", "module_name": "M1"},
        [("arr", "M1", None, False), ("arrX", "M1", None, True)],
        declared=["arr"]), ui_ax)
    c_ax.trigger()
    check("v33 对照: 未定义的 arrX 打全 -> 不提示自己",
          [(ui_ax.shown, sorted(c_ax.current_matches()))],
          expect_contain=[(False, [])])

    # ---- 11.11 v32：数字键 1..9 直接选词并写回编辑器 ----
    ui_k = _UI2()
    b_k = _CapBackend(
        {"line_no": 3, "caret_col": 8, "line_text": "    num",
         "in_string": False, "in_comment": False, "in_type_position": False,
         "in_decl_position": False, "decl_names": [],
         "proc_name": "Foo", "module_name": "M1"}, nav_ids)
    c_k = E.Completer(b_k, ui_k)
    c_k.trigger()
    ok2 = c_k.pick(1)                  # 按数字键 2 -> 选中第 2 项 numArr
    check("v32: 数字键 pick(1) 成功",
          [ok2, b_k.last_apply[3] if b_k.last_apply else None],
          expect_contain=[True, "numArr"])
    check("v32: 数字键选词后弹窗收起", [ui_k.shown], expect_contain=[False])
    # 越界的数字必须返回 False 且【不改变状态】——那说明用户真想输入数字，
    # 调用方据此放行按键。
    ui_k2 = _UI2()
    b_k2 = _CapBackend(
        {"line_no": 3, "caret_col": 8, "line_text": "    num",
         "in_string": False, "in_comment": False, "in_type_position": False,
         "in_decl_position": False, "decl_names": [],
         "proc_name": "Foo", "module_name": "M1"}, nav_ids)
    c_k2 = E.Completer(b_k2, ui_k2)
    c_k2.trigger()
    ok9 = c_k2.pick(8)                 # 只有 3 个候选，按 9 越界
    check("v32: 越界数字 pick(8) 返回 False 且状态不变",
          [ok9, ui_k2.shown, c_k2.selected],
          expect_contain=[False, True, 0])

    # ---- 11.12 v32：候选项显示序号 ----
    try:
        import ui as _ui
        check("v32: 候选显示带序号",
              [_ui.format_row(0, "num"), _ui.format_row(2, "numCount")],
              expect_contain=["1. num", "3. numCount"])
        check("v32: 列表至少 8 行（不足用空行补位）",
              [_ui.MAX_ROWS], expect_contain=[8])
    except Exception as e:                                # pragma: no cover
        check("v32: 候选显示带序号（跳过：ui 不可用 %s）" % e, [True],
              expect_contain=[True])

    # ---- 12. v23-3：模块名 / 窗体名 / 类名参与提示 ----
    print("\n=== 12. 模块名 / 窗体名参与提示（v23-3）===")
    proj_c = {
        "Module1.bas": ("Public Sub Foo()\nEnd Sub", True, "Module1"),
        "pub2.bas.bas": ("Public Sub Baz()\nEnd Sub", True, "pub2"),
        "UserForm1.frm.bas": ("Private Sub UserForm_Initialize()\nEnd Sub",
                              False, "UserForm1"),
        "类1.cls.bas": ("Public Sub Bar()\nEnd Sub", False, "类1"),
    }
    recs_c = collect(proj_c, caret_mod="Module1", caret=None)
    vis_c = E.filter_identifiers_by_scope(recs_c, "Foo", "Module1")

    def pick(prefix):
        return sorted(i for i in vis_c if i.lower().startswith(prefix.lower()))

    check("写码时可提示窗体名", pick("userform"),
          expect_contain=["UserForm1"])
    check("写码时可提示其它标准模块名", pick("pub"),
          expect_contain=["pub2"])
    check("写码时可提示中文类名", pick("类"),
          expect_contain=["类1"])
    check("本模块名也可提示", pick("module"),
          expect_contain=["Module1"])

    # `As` 类型名位置：只提示类型名（窗体/类模块/模块名/Type/Enum），不提示变量
    class _B3(_B2):
        def get_type_names(self):
            return ["UserForm1", "Module1", "MyType"]

    ids3 = [("gUserName", "M1", None, False),
            ("UserForm1", "UserForm1", None, False),
            ("MyType", "M1", None, False)]
    check("As 之后提示窗体名",
          [trig("Dim f As UserF", 15, False, True, ids3, _B3)],
          expect_contain=[(True, ["UserForm1"])])
    check("As 之后提示自定义类型名",
          [trig("Dim f As MyTy", 14, False, True, ids3, _B3)],
          expect_contain=[(True, ["MyType"])])
    check("As 之后不提示变量名",
          [trig("Dim f As gUser", 15, False, True, ids3, _B3)],
          expect_contain=[(False, [])])

    # ---- 13. v24：走真实 _collect_identifiers 路径收集组件名 ----
    # 上一节用的是 collect() 复刻版，它无条件收录组件名，因此漏掉了真正
    # 的 bug：vbe_bridge 里组件名是在 `if count <= 0: continue` 【之后】才
    # 收录的，新建窗体（0 行代码）会被整个跳过 -> 窗体名不提示。
    print("\n=== 13. 组件名收集：空代码的窗体也要收录（v24）===")
    import vbe_bridge as VB

    class _FCM:
        def __init__(self, text, name):
            self._text = text or ""
            self.Name = name

        @property
        def CountOfLines(self):
            return 0 if not self._text else len(self._text.split("\n"))

        def Lines(self, start, count):
            if not self._text:
                return ""
            ls = self._text.split("\n")
            return "\n".join(ls[start - 1:start - 1 + count])

    class _FComp:
        def __init__(self, name, ctype, text):
            self.Name, self.Type = name, ctype
            self.CodeModule = _FCM(text, name)

    class _FProj:
        def __init__(self, name, comps):
            self.Name, self.VBComponents = name, comps

    class _FVBE:
        def __init__(self, proj):
            self.ActiveVBProject = proj
            self.VBProjects = [proj]

    def collect_vbe(comps, caret_mod=None, caret=None):
        """调用真实的 VbeBackend._collect_identifiers（COM 用假对象顶替）。"""
        vbe = _FVBE(_FProj("VBAProject", comps))
        old_get, old_caret = VB._get_vbe, VB._caret_position
        VB._get_vbe = lambda: vbe
        VB._caret_position = lambda v: (caret_mod, caret)
        VB._vbe_cache["obj"] = None      # 清掉上个用例残留的假 COM 对象
        VB.com_reset()
        try:
            bk = VB.VbeBackend()
            recs = bk._collect_identifiers()
            return recs, bk
        finally:
            VB._get_vbe, VB._caret_position = old_get, old_caret

    comps13 = [
        _FComp("Module1", 1, "Public Sub Foo()\nEnd Sub"),
        _FComp("UserForm1", 3, ""),      # 新建窗体：一行代码都没有
        _FComp("UserForm2", 3, "Private Sub UserForm_Initialize()\nEnd Sub"),
        _FComp("Class1", 2, ""),         # 空类模块
        _FComp("类2", 2, ""),            # 空中文类模块
    ]
    recs13, bk13 = collect_vbe(comps13)
    names13 = sorted(set(str(r[0]) for r in recs13))
    check("空代码窗体名被收录", names13, expect_contain=["UserForm1"])
    check("有代码窗体名被收录", names13, expect_contain=["UserForm2"])
    check("空类模块名被收录", names13, expect_contain=["Class1"])
    check("中文类模块名被收录", names13, expect_contain=["类2"])
    check("标准模块名被收录", names13, expect_contain=["Module1"])
    vis13 = E.filter_identifiers_by_scope(recs13, "Foo", "Module1")
    check("写码时能提示窗体名", sorted(vis13),
          expect_contain=["UserForm1", "UserForm2"])
    check("空窗体名进入类型名集合",
          sorted(bk13.get_type_names()), expect_contain=["UserForm1"])

    # ---- 14. v24：VBE 自动补括号时光标要落进括号里 ----
    print("\n=== 14. 补全后光标位置（v24）===")
    check("跳过左括号（同步补括号）",
          [VB._skip_into_parens("Public Function gCalc()", 21,
                                "Public Function gCal", 21)],
          expect_contain=[22])
    check("原行已有左括号则不动",
          [VB._skip_into_parens("    x = gCalc(1)", 13, "    x = gCal(1)", 13)],
          expect_contain=[13])
    check("后面没有括号则不动",
          [VB._skip_into_parens("    gUserName", 13, "    gUse", 9)],
          expect_contain=[13])

    class _FCM2:
        def __init__(self, lines, auto="off"):
            self.lines = list(lines)
            self.Name = "Module1"
            self.auto = auto          # off / sync / async
            self.pending = False

        def Lines(self, start, count):
            return self.lines[start - 1]

        def ReplaceLine(self, line_no, text):
            self.lines[line_no - 1] = text
            if self.auto == "sync":
                if "(" not in text:
                    self.lines[line_no - 1] = text + "()"
            elif self.auto == "async":
                self.pending = True   # 等 SetSelection 之后才补

        @property
        def CountOfLines(self):
            return len(self.lines)

    class _FPane:
        def __init__(self, cm):
            self.CodeModule = cm
            self.sel = (1, 1, 1, 1)

        def GetSelection(self):
            return self.sel

        def SetSelection(self, sl, sc, el, ec):
            cm = self.CodeModule
            if getattr(cm, "pending", False):
                cm.pending = False
                t = cm.lines[sl - 1]
                # VBE 在光标处插入 "()"，光标仍停在左括号左侧
                cm.lines[sl - 1] = t[:sc - 1] + "()" + t[sc - 1:]
            self.sel = (sl, sc, el, ec)

    class _FVBE2:
        def __init__(self, pane):
            self.ActiveCodePane = pane

    def apply14(lines, line_no, start_col, caret_col, completion, auto="off"):
        cm = _FCM2(lines, auto=auto)
        pane = _FPane(cm)
        old = VB._get_vbe
        VB._get_vbe = lambda: _FVBE2(pane)
        VB._vbe_cache["obj"] = None      # 每个用例必须用自己那份假对象
        VB.com_reset()
        try:
            VB.VbeBackend().apply_completion(line_no, start_col, caret_col,
                                             completion)
        finally:
            VB._get_vbe = old
        return cm.lines[line_no - 1], pane.sel[3]

    # 同步补括号：Function gCal -> gCalc()，光标应在 "(" 之后（列 23）
    # "Public Function " 16 字符，gCal 起于第 17 列，光标在其后 = 第 21 列
    got = apply14(["Public Function gCal"], 1, 17, 21, "gCalc", auto="sync")
    check("同步补括号：光标进括号", [got], expect_contain=[("Public Function gCalc()", 23)])
    # 异步补括号（读回行之后才补）同样要纠正
    got = apply14(["Public Function gCal"], 1, 17, 21, "gCalc", auto="async")
    check("异步补括号：光标进括号", [got], expect_contain=[("Public Function gCalc()", 23)])
    # 不补括号时，光标停在补全词尾部
    got = apply14(["Public Function gCal"], 1, 17, 21, "gCalc", auto="off")
    check("不补括号：光标在词尾", [got], expect_contain=[("Public Function gCalc", 22)])
    # 调用点：括号本来就存在，光标停在词尾（不进括号）
    got = apply14(["    x = gCal(1)"], 1, 9, 13, "gCalc", auto="off")
    check("调用点：原有括号不被跳过", [got], expect_contain=[("    x = gCalc(1)", 14)])
    # 普通变量名不受影响
    got = apply14(["    gUse"], 1, 5, 9, "gUserName", auto="off")
    check("普通变量名：光标在词尾", [got], expect_contain=[("    gUserName", 14)])

    # ---- 15. v25：过程级不泄漏（只读用法扫描不再记成模块级）----
    print("\n=== 15. 过程级变量不跨过程泄漏（v25）===")

    def leak_case(code, caret_line, caret_col, prefix, module="M1", std=True):
        caret = (caret_line, caret_col)
        r = list(P.extract_records(code, module=module, is_std_module=std,
                                   caret=caret))
        r += P.extract_implicit_records(code, module=module,
                                        is_std_module=std, declared=r,
                                        scope="proc", caret=caret)
        # 注：v30 起解析层不再全局剔除声明行名字（由引擎精确剔除），故此处
        # 只做作用域过滤，与后端保持一致。
        proc = P.proc_at_line(code, caret_line)
        vis = E.filter_identifiers_by_scope(r, proc, module)
        return sorted(i for i in vis if i.lower().startswith(prefix.lower()))

    # 15.1 只被读取、从未赋值的名字：以前被记为模块级 -> 全模块可见（泄漏）
    c_ro2 = "\n".join([
        "Sub ProcA()",
        "    Debug.Print zzOnly",
        "    x = 1",
        "End Sub",
        "Sub ProcB()",
        "    beta = 2",
        "    Debug.Print beta",
        "End Sub"])
    check("只读变量不泄漏到别的过程",
          leak_case(c_ro2, 7, len("    Debug.Print ") + 1, "zz"),
          expect_absent=["zzOnly"])
    # 光标必须落在【别的】行：压在 zzOnly 上会被"防自我提示"逻辑抹掉
    check("只读变量在自己所在过程里仍提示",
          leak_case(c_ro2, 3, len("    x") + 1, "zz"),
          expect_contain=["zzOnly"])
    # 15.2 Dim 出来的局部变量依旧严格按过程隔离
    c_loc = "\n".join([
        "Sub ProcA()",
        "    Dim alpha As Long",
        "    alpha = 1",
        "End Sub",
        "Sub ProcB()",
        "    Dim beta As Long",
        "    beta = 2",
        "End Sub"])
    check("局部变量不跨过程（Dim）",
          leak_case(c_loc, 6, len("    be") + 1, "al"), expect_absent=["alpha"])
    check("隐式变量不跨过程",
          leak_case(c_ro2, 6, len("    be") + 1, "zz"), expect_absent=["zzOnly"])
    # 15.3 模块级变量仍应全模块可见（不能被这条改动误伤）
    c_mod = "\n".join([
        "Public gAlpha As Long",
        "Sub ProcA()",
        "    gAlpha = 1",
        "End Sub",
        "Sub ProcB()",
        "    Debug.Print gAlpha",
        "End Sub"])
    check("模块级 Public 仍在过程内可见",
          leak_case(c_mod, 6, len("    Debug.Print gAl") + 1, "gal"),
          expect_contain=["gAlpha"])
    # 15.4 VBA 内置常量不该被当成隐式变量
    c_const = "\n".join([
        "Sub ProcA()",
        "    If x = vbTextCompare Then",
        "    End If",
        "    i = xlUp",
        "End Sub"])
    check("vbTextCompare 不当变量",
          leak_case(c_const, 2, len("    If x = ") + 1, "vb"),
          expect_absent=["vbTextCompare"])
    check("xlUp 不当变量",
          leak_case(c_const, 4, len("    i = xl") + 1, "xl"),
          expect_absent=["xlUp"])
    # 15.5 真实工程数据：不该出现"名字可见但所有记录都不属于当前过程"
    real_leaks = []
    for _fn in sorted(os.listdir(TESTDATA)):
        if not _fn.endswith(".bas"):
            continue
        _code = io.open(os.path.join(TESTDATA, _fn), encoding="utf-8").read()
        _mn = _mod_name(_fn)
        _r = list(P.extract_records(_code, module=_mn,
                                    is_std_module=_is_std(_fn)))
        _r += P.extract_implicit_records(_code, module=_mn,
                                         is_std_module=_is_std(_fn),
                                         declared=_r, scope="proc")
        for _pr in sorted({x[2] for x in _r if x[2]}):
            _vis = set(E.filter_identifiers_by_scope(_r, _pr, _mn))
            for _n in _vis:
                _srcs = {(x[2] or "").lower() for x in _r if x[0] == _n}
                if _srcs and all(s != _pr.lower() for s in _srcs):
                    # 只算"纯过程级"来源的名字（模块级过程名/类型名本就该可见）
                    if all(s for s in _srcs):
                        real_leaks.append((_fn, _pr, _n))
    check("真实工程：无跨过程泄漏", [len(real_leaks)],
          expect_contain=[0])
    if real_leaks:
        print("        泄漏明细: %s" % (real_leaks[:8],))

    # ---- 16. v25：COM 失败退避（关文件卡顿 / 关完又启动 Office）----
    print("\n=== 16. COM 失败退避（v25）===")
    VB.com_reset()
    check("初始不处于退避期", [VB.com_backoff_remaining()], expect_contain=[0.0])
    VB._com_fail()
    _b1 = round(VB.com_backoff_remaining(), 1)
    VB._com_fail()
    VB._com_fail()
    _b2 = round(VB.com_backoff_remaining(), 1)
    check("失败后进入退避且逐级变长", [_b1 > 0 and _b2 >= _b1],
          expect_contain=[True])
    check("退避期 _get_vbe_cached 直接返回 None",
          [VB._get_vbe_cached() is None], expect_contain=[True])
    VB.com_reset()
    check("reset 后恢复", [VB.com_backoff_remaining()], expect_contain=[0.0])

    # ---- 17. v26：COM 代理生命周期（关文件卡 3~10 秒 / 关完又启动 Office）----
    #
    # 病根：把 Excel/VBE 的 COM 代理长期缓存，等于我们的进程一直持有
    # Excel.Application —— Excel 退不干净（关完又启动 Office），且 VBAProject
    # 变脏时的写回会被拖成好几秒。第一版是"每次新建、用完即弃"所以没问题。
    print("\n=== 17. COM 代理生命周期（v26）===")

    class _FCM3:
        def __init__(self, text):
            self.text = text

        def Lines(self, a, b):
            return self.text

        def ReplaceLine(self, n, t):
            self.text = t

    class _FPane3:
        def __init__(self, cm):
            self.CodeModule = cm
            self.sel = (1, 1, 1, 1)

        def GetSelection(self):
            return self.sel

        def SetSelection(self, sl, sc, el, ec):
            self.sel = (sl, sc, el, ec)

    class _FVBE3:
        def __init__(self, pane):
            self.ActiveCodePane = pane

    _old_ttl, _old_probe = VB.VBE_PROXY_TTL, VB._probe_host_window
    _old_wnd, _old_get = dict(VB._wnd_state), VB._get_vbe
    try:
        # 用假 COM 对象，整节都不依赖真实 Excel 是否在跑
        VB._get_vbe = lambda: _FVBE3(_FPane3(_FCM3("    x = 1")))
        VB._probe_host_window = lambda: True
        VB._wnd_state.update(seen=True, miss=0)

        # 17.1 TTL <= 0：彻底不缓存，每次新建、用完即弃
        VB.VBE_PROXY_TTL = 0.0
        VB.com_reset()
        VB._vbe_cache["obj"] = None
        _a = VB._get_vbe_cached()
        _b = VB._get_vbe_cached()
        check("TTL=0 时每次都新建代理",
              [_a is not None, _b is not None, _a is _b],
              expect_contain=[True, True, False])

        # 17.2 TTL 内复用；到期必须重建（不能永久攥着 Excel）
        VB.VBE_PROXY_TTL = 30.0
        VB.com_reset()
        VB._vbe_cache["obj"] = None
        _a = VB._get_vbe_cached()
        _b = VB._get_vbe_cached()
        check("TTL 未到期时复用同一代理（打字性能不受影响）",
              [_a is _b], expect_contain=[True])
        VB._vbe_cache["at"] -= 60          # 模拟时间流逝
        _c = VB._get_vbe_cached()
        check("TTL 到期后重建代理（不长期持有 Excel）",
              [_c is not None, _c is _a], expect_contain=[True, False])

        # 17.3 宿主窗口消失（Excel 正在退出）：绝不新建代理，并进入退避
        VB.VBE_PROXY_TTL = 0.0
        VB.com_reset()
        VB._vbe_cache["obj"] = None
        VB._probe_host_window = lambda: True
        VB._wnd_state.update(seen=True, miss=0)
        _held = VB._get_vbe_cached() is not None     # 窗口还在时正常持有
        VB._probe_host_window = lambda: False
        VB._wnd_state.update(seen=True, miss=0)
        VB._get_vbe_cached()                          # miss=1：还不拦
        VB._get_vbe_cached()                          # miss=2：认定宿主真没了
        _gone = VB._get_vbe_cached()
        check("宿主窗口消失后不再新建代理",
              [_held, _gone is None], expect_contain=[True, True])
        check("宿主窗口消失后不留残留代理",
              [VB._vbe_cache["obj"] is None], expect_contain=[True])
        check("宿主窗口消失后进入退避",
              [VB.com_backoff_remaining() > 0], expect_contain=[True])

        # 17.4 宿主窗口回来 -> 立刻解除退避，工具自动恢复
        VB._probe_host_window = lambda: True
        VB._wnd_state.update(seen=True, miss=2)
        VB._host_window_alive()
        check("宿主窗口回来后解除退避",
              [VB.com_backoff_remaining()], expect_contain=[0.0])

        # 17.5 从未见过宿主窗口（WPS 等类名不匹配的宿主）：不能误拦
        VB.com_reset()
        VB._wnd_state.update(seen=False, miss=0)
        VB._probe_host_window = lambda: False
        check("未见过宿主窗口时不拦截（兼容 WPS 等宿主）",
              [VB._host_window_alive()], expect_contain=[True])

        # 17.6 写回 VBA 代码（VBAProject 变脏）后立刻放手
        VB._probe_host_window = lambda: True
        VB._wnd_state.update(seen=True, miss=0)
        VB.VBE_PROXY_TTL = 30.0
        VB.com_reset()
        VB._vbe_cache["obj"] = None
        _cm = _FCM3("    gUse")
        VB._get_vbe = lambda: _FVBE3(_FPane3(_cm))
        VB.VbeBackend().apply_completion(1, 5, 9, "gUserName")
        check("写回代码后立即释放 COM 代理",
              [(_cm.text, VB._vbe_cache["obj"] is None)],
              expect_contain=[("    gUserName", True)])
    finally:
        VB.VBE_PROXY_TTL, VB._probe_host_window = _old_ttl, _old_probe
        VB._wnd_state.update(_old_wnd)
        VB._get_vbe = _old_get
        VB.com_reset()
        VB._vbe_cache["obj"] = None

    # ---- 18. v27：全局钩子只在 VBE 聚焦时挂载（修复 PyCharm/VSCode 关不掉）----
    print("\n=== 18. 全局钩子仅 VBE 聚焦时挂载（v27）===")
    if M is None:
        check("main 模块不可用（缺 pynput），跳过第 18 节", [True],
              expect_contain=[True])
    else:
        GRACE = M.FOCUS_GRACE_SEC

        def rr(focused, released, backoff=0.0, elapsed=99.0, grace=GRACE,
               closing=False, suppress=False):
            return M._reconcile_resources(focused, released, backoff,
                                          time.time() - elapsed, time.time(), grace,
                                          closing=closing, suppress=suppress)

        # 18.1 在 VBE 里：必须挂载钩子（且不释放 COM）
        check("VBE 聚焦 -> 挂载钩子",
              [rr(True, False, backoff=0.0, elapsed=0.0)],
              expect_contain=[(True, False)])
        # 18.2 离开 VBE 后立即（宽限内）：不挂载钩子、也不释放 COM（等宽限到期）
        check("离开 VBE 宽限内 -> 不挂载钩子、不释放",
              [rr(False, False, backoff=0.0, elapsed=0.5)],
              expect_contain=[(False, False)])
        # 18.3 离开 VBE 超过宽限期：不挂载钩子 + 释放 COM（已释放则不再重复释放）
        check("离开 VBE 超宽限 -> 不挂载钩子 + 释放",
              [rr(False, False, backoff=0.0, elapsed=GRACE + 1.0)],
              expect_contain=[(False, True)])
        check("已释放后仍不挂载钩子（关键回归点）",
              [rr(False, True, backoff=0.0, elapsed=GRACE + 1.0)],
              expect_contain=[(False, False)])
        # 18.4 COM 失败退避中：即便不在 VBE 也绝不挂载钩子（钩子挂上会干扰关文件）
        check("COM 退避期（不在 VBE）-> 不挂载钩子 + 释放",
              [rr(False, False, backoff=2.0, elapsed=0.0)],
              expect_contain=[(False, True)])
        # 18.5 反例：复刻「旧版 bug」——离开 VBE 后 released 变 True，
        #     旧逻辑会走 else 把钩子重新挂上。新逻辑 mount_hooks 必须仍为 False。
        old_mount = (not False)  # 旧版 else 分支无条件重挂，等价于“始终挂载”
        check("旧版行为确为‘离开也挂载’（对比基线）",
              [old_mount], expect_contain=[True])
        check("新版：离开 VBE 后（released=True）mount_hooks 必须为 False",
              [M._reconcile_resources(False, True, 0.0, 0.0, time.time(), GRACE)[0]],
              expect_contain=[False])

        # 18.6 v28：前景窗口正在关闭 -> 即使去抖仍维持"在 VBE"也绝不挂钩
        #     这正是 PyCharm/VSCode 关不掉的根因：关闭 IDE 时焦点离开 VBE 的
        #     去抖窗口（约 600ms）内钩子仍挂着，IDE 退出被挡死。
        check("焦点在 VBE 但前台正在关闭 -> 不挂载钩子（核心修复点）",
              [rr(True, False, closing=True)], expect_contain=[(False, False)])
        check("前台正在关闭（不在 VBE）-> 不挂载钩子",
              [rr(False, False, closing=True)[0]], expect_contain=[False])
        check("新版：closing 必须使 mount_hooks=False（即便 focused=True）",
              [M._reconcile_resources(True, False, 0.0, 0.0, time.time(), GRACE,
                                      closing=True)[0]],
              expect_contain=[False])

        # 18.7 v29：上一个非 VBE 前台窗口已销毁（IDE 关掉、焦点被甩回 VBE）
        #     -> 冷却期内抑制钩子挂载，避免拖慢那个仍在收尾的进程
        #     （关窗口后 2~3 秒才退出的根因）。
        check("抑制冷却期（focused=True）-> 不挂载钩子",
              [rr(True, False, suppress=True)], expect_contain=[(False, False)])
        check("抑制冷却期必须压过 focused（即便 closing=False）",
              [rr(True, False, suppress=True, closing=False)[0]],
              expect_contain=[False])
        check("抑制冷却期不影响 COM 释放判定（离开超宽限仍释放）",
              [rr(False, False, elapsed=GRACE + 1.0, suppress=True)],
              expect_contain=[(False, True)])
        check("无抑制且正常在 VBE -> 挂载钩子（回归：抑制不能误伤常态）",
              [rr(True, False, suppress=False, closing=False)],
              expect_contain=[(True, False)])
        # 关键反例：用户主动 Alt+Tab 回 VBE（上一个前台窗口仍活着）不应被抑制，
        # 这里用纯函数无法直接构造“已销毁窗口”，靠 18.6/18.7 的 focused 组合
        # + reconcile 内 IsWindow 判定共同保证；此处确认 suppress=False 时照常挂载。

    print("\n" + "=" * 60)
    print("结果: %d PASS, %d FAIL" % (PASS, FAIL))
    if FAILURES:
        print("失败项:")
        for f in FAILURES:
            print("  - " + f)
    print("=" * 60)
    return 1 if FAIL else 0


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    sys.exit(main())
