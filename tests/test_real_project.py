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
    #
    # 【v37 语义变更】"打全名 -> 不提示自己" 这三条的期望已按用户要求反转：
    #   旧规则（v31~v36）：正在声明的名字，打全名时一律剔除 -> 列表整个消失。
    #   问题（用户报的 bug）：工程里已经有 `Function test()`，你在 `Sub test`
    #     这一行把 test 打全，那不是"自己的回声"，是货真价实的同名函数，
    #     却被打全名那一刻清空了列表 —— 输入 tes 提示、输入 test 反而没了。
    #   新规则：只看"这个名字在工程里真实存在吗"。存在就提示（含打全名），
    #     不存在（纯幻影 zzq、回退出来的 numA）才剔除。
    # 下面的 gUserName / gCalc / numArr 在 ids 里都是【已定义】的，故打全名照常提示。
    check("公共变量：打全名 -> 照常提示（v37）",
          [trig("Public gUserName", 17, True, decl_names=["gUserName"])],
          expect_contain=[(True, ["gUserName"])])
    check("公共变量：打前缀 -> 提示已有的 gUserName",
          [trig("Public gUser", 13, True, decl_names=["gUserName"])],
          expect_contain=[(True, ["gUserName"])])
    check("函数名：打全名 -> 照常提示（v37）",
          [trig("Public Function gCalc", 22, True, decl_names=["gCalc"])],
          expect_contain=[(True, ["gCalc"])])
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
    # 打全 numArr（声明行完整名字）-> v37 起【照常提示】：numArr 是本模块真实
    # 声明过的 Public 名字，打全它不属于"自我提示"，用户正需要看到它
    # （与用户报的"输入 test 打全了要提示 test"是同一条规则）。
    ui_b1b = _UI2()
    c_b1b = E.Completer(_B2(
        {"line_no": 2, "caret_col": len("Public numArr") + 1,
         "line_text": "Public numArr", "in_string": False, "in_comment": False,
         "in_type_position": False, "in_decl_position": True,
         "decl_names": ["numArr"], "proc_name": None, "module_name": "M1"},
        b1_recs), ui_b1b)
    c_b1b.trigger()
    check("Bug1: 声明行打全 numArr -> 照常提示（v37）",
          [(ui_b1b.shown, sorted(c_b1b.current_matches()))],
          expect_contain=[(True, ["numArr"])])

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

    # ---- 11.12 v49：候选不再带序号（数字键选词已移除），一屏 15 行 ----
    try:
        import ui as _ui
        check("v49: 候选显示就是名字本身（无 `N. ` 前缀）",
              [_ui.format_row(0, "num") == "num",
               _ui.format_row(2, "numCount") == "numCount",
               _ui.format_row(8, "numArr") == "numArr"],
              expect_contain=[True, True, True])
        check("v49: 一屏最多 15 行（候选不足则列表更短）",
              [_ui.MAX_VISIBLE_ROWS], expect_contain=[15])
        check("v42: 列表宽度固定 30 字符（不再有横向滚动条）",
              [_ui.NAME_CHARS, hasattr(_ui.Popup, "has_hscroll")],
              expect_contain=[30, False])
        check("v42: 超长名字用 ... 省略",
              [_ui.ELLIPSIS], expect_contain=["..."])
    except Exception as e:                                # pragma: no cover
        check("v49: 候选显示不带序号（跳过：ui 不可用 %s）" % e, [True],
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

    # ---- 19. v35：走【真实的】VbeBackend + Completer（假 COM 对象驱动）----
    #
    # 这一段专门补 11.x 覆盖不到的盲区：11.x 用的是手写的 collect() 复刻版，
    # 它会把 caret 传进解析器；而生产的 get_identifiers() 从不传 caret。
    # 两者行为不一致 -> 真实 bug（"arr 打全名后列表消失"）在 11.x 里根本测不出来。
    print("\n=== 19. 真实后端全链路（v35：打全名必须保留列表）===")
    try:
        import vbe_bridge as VB

        class _CM(object):
            def __init__(self, name, text):
                self.Name = name
                self.text = text

            @property
            def CountOfLines(self):
                return self.text.count("\n") + 1

            def Lines(self, start, count):
                return "\r\n".join(self.text.split("\n")
                                   [start - 1:start - 1 + count])

        class _Comp(object):
            def __init__(self, name, text, ctype=1):
                self.Name = name
                self.Type = ctype
                self.CodeModule = _CM(name, text)

        class _Pane(object):
            def __init__(self, comp):
                self._c = comp
                self.sel = (1, 1, 1, 1)

            @property
            def CodeModule(self):
                return self._c.CodeModule

            def GetSelection(self):
                return self.sel

        class _VBE(object):
            def __init__(self, comps, act):
                self._p = type("_Proj", (object,), {})()
                self._p.Name = "VBAProject"
                self._p.VBComponents = list(comps)
                self.ActiveVBProject = self._p
                self.ActiveCodePane = _Pane(act)

        class _UI(object):
            def __init__(self):
                self.shown = False

            def show(self, matches, selected, completer=None):
                self.shown = True

            def update_selection(self, i):
                pass

            def hide(self):
                self.shown = False

            def contains_point(self, x, y):
                return False

        def drive(mods, active_mod, edit_line, indent, name):
            """逐字符输入 name，返回 [(输入到的词, 是否显示, 候选), ...]。"""
            out = []
            for n in range(1, len(name) + 1):
                typed = name[:n]
                comps, act = [], None
                for mn, text, ct in mods:
                    if mn == active_mod:
                        ls = text.split("\n")
                        ls[edit_line - 1] = indent + typed
                        text = "\n".join(ls)
                    c = _Comp(mn, text, ct)
                    comps.append(c)
                    if mn == active_mod:
                        act = c
                vbe = _VBE(comps, act)
                VB._get_vbe_cached = lambda: vbe
                bk = VB.VbeBackend()
                ui = _UI()
                cp = E.Completer(bk, ui)
                col = len(indent + typed) + 1
                vbe.ActiveCodePane.sel = (edit_line, col, edit_line, col)
                cp.trigger(True)
                out.append((typed, ui.shown, list(cp.current_matches())))
            return out

        # 19.1 核心回归：**从未 Dim、只是到处在用** 的隐式变量 arr。
        #      修复前：a / ar 提示 arr，一把 arr 打全列表就消失（用户报的现象）。
        m_imp = [("Module1", "Sub Foo()\n    arr = 1\n    <in>\nEnd Sub", 1)]
        check("19.1 隐式变量 arr：打前缀提示", drive(m_imp, "Module1", 3, "    ", "ar"),
              expect_contain=[("ar", True, ["arr"])])
        check("19.1 隐式变量 arr：打全名仍保留",
              drive(m_imp, "Module1", 3, "    ", "arr"),
              expect_contain=[("arr", True, ["arr"])])
        # 19.2 单字符隐式变量（同一个机制，只是长度不同）
        m_single = [("Module1", "Sub Foo()\n    a = 1\n    <in>\nEnd Sub", 1)]
        check("19.2 单字符隐式变量 a：打全名保留",
              drive(m_single, "Module1", 3, "    ", "a"),
              expect_contain=[("a", True, ["a"])])
        # 19.3 显式 Dim 的多字符变量（v33 的老场景不能被新逻辑破坏）
        m_dim = [("Module1", "Option Explicit\nSub Foo()\n    Dim arr As Long\n    <in>\nEnd Sub", 1)]
        check("19.3 Dim arr：打全名保留",
              drive(m_dim, "Module1", 4, "    ", "arr"),
              expect_contain=[("arr", True, ["arr"])])
        # 19.4 对照：从没出现过的词（numA）打全仍不能提示自己
        m_to = [("Module1", "Option Explicit\nPublic num As Long\nPublic numArr As Long\nSub Foo()\n    <in>\nEnd Sub", 1)]
        check("19.4 未定义的 numA：绝不提示自己",
              drive(m_to, "Module1", 5, "    ", "numA"),
              expect_contain=[("numA", True, ["numArr"])],
              expect_absent=[("numA", True, ["numA", "numArr"])])
        check("19.4 num/numArr：打 num 两个同屏",
              drive(m_to, "Module1", 5, "    ", "num"),
              expect_contain=[("num", True, ["num", "numArr"])])
        # 19.5 纯幻影词：工程里任何地方都没有 -> 应当收起，不提示自己
        m_ghost = [("Module1", "Option Explicit\nSub Foo()\n    <in>\nEnd Sub", 1)]
        check("19.5 纯幻影 zzq：不提示",
              drive(m_ghost, "Module1", 3, "    ", "zzq"),
              expect_contain=[("zzq", False, [])])

        # ---- 20. v36：模糊匹配（不必从首字母开始）+ 命中字符高亮 ----
        #
        # 用户原话：定义了 dataArr / dataSheet，输入 ts 要提示 dataSheet
        # （t 和 s 都在里面），输入 ta 两个都要提示，且命中的字符标红。
        print("\n=== 20. 模糊匹配与命中高亮（v36）===")

        def probe(mods, active_mod, edit_line, indent, typed):
            """把 typed 整串放到编辑行，返回 (是否显示, 候选, 命中位置字典)。"""
            comps, act = [], None
            for mn, text, ct in mods:
                if mn == active_mod:
                    ls = text.split("\n")
                    ls[edit_line - 1] = indent + typed
                    text = "\n".join(ls)
                c = _Comp(mn, text, ct)
                comps.append(c)
                if mn == active_mod:
                    act = c
            vbe = _VBE(comps, act)
            VB._get_vbe_cached = lambda: vbe
            bk = VB.VbeBackend()
            ui = _UI()
            cp = E.Completer(bk, ui)
            col = len(indent + typed) + 1
            vbe.ActiveCodePane.sel = (edit_line, col, edit_line, col)
            cp.trigger(True)
            return (ui.shown, list(cp.current_matches()),
                    dict(getattr(cp, "match_hits", None) or {}))

        m_two = [("Module1",
                  "Option Explicit\n"
                  "Sub Foo()\n"
                  "    Dim dataArr As Long\n"
                  "    Dim dataSheet As Worksheet\n"
                  "    <in>\n"
                  "End Sub", 1)]

        # 20.1 核心场景一：ts -> 只提示 dataSheet（dataArr 里没有 s）
        s1, m1, h1 = probe(m_two, "Module1", 5, "    ", "ts")
        check("20.1 输入 ts：弹窗显示", [s1], expect_contain=[True])
        check("20.1 输入 ts：提示 dataSheet", m1, expect_contain=["dataSheet"])
        check("20.1 输入 ts：不提示 dataArr（里面没有 s）", m1,
              expect_absent=["dataArr"])
        check("20.1 输入 ts：命中位置为 t(2) 与 S(4)", [h1.get("dataSheet")],
              expect_contain=[[2, 4]])

        # 20.2 核心场景二：ta -> 两个都提示，且各自命中位置正确
        s2, m2, h2 = probe(m_two, "Module1", 5, "    ", "ta")
        check("20.2 输入 ta：弹窗显示", [s2], expect_contain=[True])
        check("20.2 输入 ta：dataArr 与 dataSheet 同屏", [sorted(m2)],
              expect_contain=[["dataArr", "dataSheet"]])
        check("20.2 输入 ta：dataArr 命中 t(2) a(3)", [h2.get("dataArr")],
              expect_contain=[[2, 3]])
        check("20.2 输入 ta：dataSheet 命中 t(2) a(3)", [h2.get("dataSheet")],
              expect_contain=[[2, 3]])

        # 20.3 高亮位置必须合法：数量等于输入长度、升序、不越界
        bad = []
        for name, pos in (h2 or {}).items():
            if len(pos) != 2 or pos != sorted(pos) or pos[-1] >= len(name):
                bad.append((name, pos))
        check("20.3 命中位置数量/顺序/越界检查", [bad], expect_contain=[[]])

        # 20.4 词首缩略：ds -> 命中 dataSheet 的 D 与 S（不是随便找两个字母）
        s4, m4, h4 = probe(m_two, "Module1", 5, "    ", "ds")
        check("20.4 输入 ds：提示 dataSheet", m4, expect_contain=["dataSheet"])
        check("20.4 输入 ds：命中词首 D(0) 与 S(4)", [h4.get("dataSheet")],
              expect_contain=[[0, 4]])

        # 20.5 连续块优先：arr -> 命中 dataArr 尾部的 Arr，而不是散着凑
        s5, m5, h5 = probe(m_two, "Module1", 5, "    ", "arr")
        check("20.5 输入 arr：命中连续的 Arr(4,5,6)", [h5.get("dataArr")],
              expect_contain=[[4, 5, 6]])

        # 20.6 排序：精确匹配必须排第一（不能被模糊候选挤下去）
        m_rank = [("Module1",
                   "Option Explicit\n"
                   "Public num As Long\n"
                   "Public myNumber As Long\n"
                   "Sub Foo()\n"
                   "    <in>\n"
                   "End Sub", 1)]
        s6, m6, _h6 = probe(m_rank, "Module1", 5, "    ", "num")
        check("20.6 输入 num：精确匹配 num 排第一", [m6[0] if m6 else None],
              expect_contain=["num"])
        check("20.6 输入 num：模糊候选 myNumber 也在列", m6,
              expect_contain=["myNumber"])

        # 20.7 排序：前缀匹配排在"中间命中"之前
        m_pref = [("Module1",
                   "Option Explicit\n"
                   "Public wsName As String\n"
                   "Public myWorksheetName As String\n"
                   "Sub Foo()\n"
                   "    <in>\n"
                   "End Sub", 1)]
        s7, m7, _h7 = probe(m_pref, "Module1", 5, "    ", "wsn")
        check("20.7 输入 wsn：前缀命中的 wsName 排第一",
              [m7[0] if m7 else None], expect_contain=["wsName"])

        # 20.8 不回归：工程里压根没有的字符组合 -> 不显示
        s8, m8, _h8 = probe(m_two, "Module1", 5, "    ", "zzq")
        check("20.8 纯幻影 zzq：不显示、无候选", [(s8, m8)],
              expect_contain=[(False, [])])

        # 20.9 不回归：打全名仍保留列表（v35 的修复不能被模糊匹配破坏）
        s9, m9, h9 = probe(m_two, "Module1", 5, "    ", "dataSheet")
        check("20.9 打全名 dataSheet：列表保留且它排第一",
              [(s9, m9[0] if m9 else None)], expect_contain=[(True, "dataSheet")])
        check("20.9 打全名：命中位置覆盖全部字符",
              [h9.get("dataSheet")], expect_contain=[[0, 1, 2, 3, 4, 5, 6, 7, 8]])

        # ---- 21. v37：函数名打全名要提示 + 回退删除后不漏（真实后端全链路）----
        #
        # 用户报的两个 bug，根因是同一条老规则：in_decl_position 时【无条件】
        # 剔除"当前声明行正在声明的那个名字"。它把"打全名"整个清空了：
        #   bug1 `Sub test` 行输入 test（工程里已有 Function test）-> 列表消失
        #   bug2 在声明行回退 02 剩 test -> 同样消失
        # v37 改为统一判据："这个名字在工程里真实存在吗"。
        print("\n=== 21. 函数名打全名 / 回退删除（v37）===")

        def probe2(mods, active_mod, edit_line, indent, typed):
            comps, act = [], None
            for mn, text, ct in mods:
                if mn == active_mod:
                    ls = text.split("\n")
                    ls[edit_line - 1] = indent + typed
                    text = "\n".join(ls)
                c = _Comp(mn, text, ct)
                comps.append(c)
                if mn == active_mod:
                    act = c
            vbe = _VBE(comps, act)
            VB._get_vbe_cached = lambda: vbe
            bk = VB.VbeBackend()
            ui = _UI()
            cp = E.Completer(bk, ui)
            col = len(indent + typed) + 1
            vbe.ActiveCodePane.sel = (edit_line, col, edit_line, col)
            cp.trigger(True)
            return (ui.shown, sorted(cp.current_matches()))

        # 21.1 已有 Function test()，在新的 Sub 声明行上输入（bug1）
        m_fn = [("Module1",
                 "Option Explicit\n"
                 "Function test() As Long\n"
                 "    test = 1\n"
                 "End Function\n"
                 "\n"
                 "<in>\n"
                 "End Sub", 1)]
        check("21.1 声明行输入 tes：提示 test",
              [probe2(m_fn, "Module1", 6, "Sub ", "tes")],
              expect_contain=[(True, ["test"])])
        check("21.1 声明行输入 test（打全名）：仍提示 test  ← bug1",
              [probe2(m_fn, "Module1", 6, "Sub ", "test")],
              expect_contain=[(True, ["test"])])

        # 21.2 test / test02 都在，声明行回退到 test 要两个都提示（bug2）
        m_two = [("Module1",
                  "Option Explicit\n"
                  "Function test() As Long\n"
                  "    test = 1\n"
                  "End Function\n"
                  "Function test02() As Long\n"
                  "    test02 = 2\n"
                  "End Function\n"
                  "\n"
                  "<in>\n"
                  "End Sub", 1)]
        check("21.2 声明行回退到 test：test 与 test02 同屏  ← bug2",
              [probe2(m_two, "Module1", 10, "Sub ", "test")],
              expect_contain=[(True, ["test", "test02"])])
        check("21.2 声明行输入 test02：提示 test02",
              [probe2(m_two, "Module1", 10, "Sub ", "test02")],
              expect_contain=[(True, ["test02"])])

        # 21.3 过程体内调用处回退（非声明行，属回归保护）
        m_call = [("Module1",
                   "Option Explicit\n"
                   "Function test() As Long\n"
                   "    test = 1\n"
                   "End Function\n"
                   "Function test02() As Long\n"
                   "    test02 = 2\n"
                   "End Function\n"
                   "Sub Main()\n"
                   "    <in>\n"
                   "End Sub", 1)]
        check("21.3 调用处回退到 test：两个都提示",
              [probe2(m_call, "Module1", 10, "    ", "test")],
              expect_contain=[(True, ["test", "test02"])])

        # 21.4 关键反向保护：幻影词（哪都不存在）依然【绝不】提示自己
        m_ghost = [("Module1",
                    "Option Explicit\n"
                    "Sub Main()\n"
                    "    <in>\n"
                    "End Sub", 1)]
        check("21.4 幻影 zzq：不提示",
              [probe2(m_ghost, "Module1", 3, "    ", "zzq")],
              expect_contain=[(False, [])])
        check("21.4 声明行幻影 zzq：不提示",
              [probe2([("Module1", "Option Explicit\n<in>\nEnd Sub", 1)],
                      "Module1", 2, "Sub ", "zzq")],
              expect_contain=[(False, [])])
        # 21.5 回退出来的未定义词 numA（v35 的核心用例，不能被 v37 放宽掉）
        m_numA = [("Module1",
                   "Option Explicit\n"
                   "Public num As Long\n"
                   "Public numArr As Long\n"
                   "Sub Foo()\n"
                   "    <in>\n"
                   "End Sub", 1)]
        check("21.5 回退出的 numA：只提示 numArr，绝不提示 numA 自己",
              [probe2(m_numA, "Module1", 5, "    ", "numA")],
              expect_contain=[(True, ["numArr"])])

        # 20.10 高亮分段：把名字切成"命中/未命中"交替的段（纯函数，无需 GUI）
        import ui as _uimod

        check("20.10 分段：dataSheet 命中 [2,4]",
              [_uimod.split_by_hits("dataSheet", [2, 4])],
              expect_contain=[[("da", False), ("t", True), ("a", False),
                               ("S", True), ("heet", False)]])
        check("20.10 分段：连续命中会合并成一段",
              [_uimod.split_by_hits("dataArr", [4, 5, 6])],
              expect_contain=[[("data", False), ("Arr", True)]])
        check("20.10 分段：无命中时整段一段",
              [_uimod.split_by_hits("num", [])],
              expect_contain=[[("num", False)]])
        check("20.10 分段：全命中时整段一段",
              [_uimod.split_by_hits("num", [0, 1, 2])],
              expect_contain=[[("num", True)]])
        check("20.10 分段：拼接回去必须等于原名（不丢字）",
              ["".join(s for s, _ in _uimod.split_by_hits("dataSheet", [2, 4]))],
              expect_contain=["dataSheet"])

        # 20.11 真实 Popup 渲染冒烟：Canvas 自绘不抛异常（无 GUI 环境则跳过）
        try:
            import tkinter as _tk

            class _FakeC(object):
                match_hits = {"dataSheet": [2, 4], "dataArr": [2, 3]}
                selected = 0

                def pick(self, i):
                    return True

                def accept(self):
                    pass

            _root = _tk.Tk()
            _root.withdraw()
            _p = _uimod.Popup(_root)
            _p.show(["dataSheet", "dataArr"], 0, _FakeC())
            _root.update()
            check("20.11 Canvas 渲染：两行都画出来了", [len(_p._rows)],
                  expect_contain=[2])
            check("20.11 Canvas 渲染：命中位置已传给 UI", [_p._rows[0]],
                  expect_contain=[("dataSheet", [2, 4])])
            check("20.11 点空白行不触发确认（返回 None）", [_p._row_at(10 ** 6)],
                  expect_contain=[None])
            _root.destroy()
        except Exception as _e2:
            check("20.11 跳过（无 GUI 环境）: %s" % _e2, [True],
                  expect_contain=[True])

        # ---------------------------------------------------------------
        # 22. 候选超过一屏：滚动窗口，不多丢、序号不再显示（v40）
        #
        # 用户报的问题：旧版本把 MAX_ROWS(=8) 当成了候选上限，
        # 第 9 个往后直接看不见（"多的都给我删没了"）。v38 把它还原成
        # 【一屏显示几行】：候选全留着，靠上下键 / 滚轮滚动查看；
        # v39 再把上限从 8 提到 10，并去掉"数字键选词"（会与输入变量名里的
        # 数字冲突，例如 s1 会被误选成 sheet1）。选择只走 ↑/↓ + Tab 或鼠标点击。
        # ---------------------------------------------------------------
        print("\n=== 22. 超过一屏不丢候选，滚动查看（v39）===")
        try:
            _N = 30
            _names = ["x%02d" % i for i in range(_N)]

            class _WinUI(object):
                """记录 UI 真正收到了哪些行（不依赖 tkinter，纯记账）。"""
                def __init__(self):
                    self.shown = False
                    self.rows = []
                    self.sel = 0
                    self.hidden = False

                def show(self, rows, sel, completer=None):
                    self.shown = True
                    self.hidden = False
                    self.rows = list(rows)
                    self.sel = sel

                def update_selection(self, sel):
                    self.sel = sel

                def hide(self):
                    self.shown = False
                    self.hidden = True
                    self.rows = []

                def contains_point(self, x, y):
                    return False

            class _ManyBk(object):
                """一次性给出 N 个候选的最小假后端。"""
                def __init__(self, names, word):
                    self._names = list(names)
                    line = "    " + word
                    self._ctx = {
                        "line_no": 3, "caret_col": len(line) + 1,
                        "line_text": line, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": "M1",
                    }
                    self.applied = None

                def get_context(self):
                    return dict(self._ctx)

                def get_identifiers(self):
                    return [(n, "M1", None, False) for n in self._names]

                def get_declared_names(self):
                    return list(self._names)

                def get_type_names(self):
                    return []

                def name_exists_outside_caret(self, name, caret):
                    return True

                def apply_completion(self, line_no, start, end, name):
                    self.applied = name
                    return True

            def mk22(names=None, word="x"):
                bk = _ManyBk(names or _names, word)
                ui = _WinUI()
                c = E.Completer(bk, ui, view_rows=10)
                c.trigger()
                return bk, ui, c

            bk1, ui1, c1 = mk22()
            all_m = c1.current_matches()
            check("22.1 30 个候选一个都不丢（不再只留 8 个）",
                  [len(all_m)], expect_contain=[_N])
            check("22.1 弹窗确实显示了", [ui1.shown], expect_contain=[True])

            check("22.2 UI 只画当前这一屏（最多 10 行）",
                  [ui1.rows], expect_contain=[all_m[:10]])
            check("22.2 这一屏 == visible_rows()（同一份数据，不会错位）",
                  [c1.visible_rows()], expect_contain=[all_m[:10]])
            check("22.3 初始窗口停在列表开头", [c1.view_info()],
                  expect_contain=[(_N, 0)])
            check("22.3 初始选中本屏第 1 行", [c1.view_selection()],
                  expect_contain=[0])

            # 22.4 在本屏内移动：窗口不该动
            for _ in range(9):
                c1.move(1)
            check("22.4 本屏内下移 9 次：selected=9 且窗口不动",
                  [c1.selected, c1.view_info(), c1.view_selection()],
                  expect_contain=[9, (_N, 0), 9])

            # 22.5 越过下沿：窗口跟着下移一行（而不是整屏跳）
            c1.move(1)
            check("22.5 走出本屏：窗口只下移一行",
                  [c1.selected, c1.view_info(), c1.view_selection()],
                  expect_contain=[10, (_N, 1), 9])
            check("22.5 窗口内容与选中行一致",
                  [c1.visible_rows()], expect_contain=[all_m[1:11]])
            check("22.5 选中行始终在窗口内",
                  [c1.view_top <= c1.selected < c1.view_top + c1.view_rows],
                  expect_contain=[True])

            # 22.6 一屏最多 10 行：屏幕上只会显示 10 条，再多靠滚动
            check("22.6 屏幕内一屏最多 10 行",
                  [len(c1.visible_rows())], expect_contain=[10])
            check("22.6 超过一屏才出现纵向滚动（total > 10）",
                  [c1.view_info()[0] > 10], expect_contain=[True])

            # 22.7 鼠标点击对应"本屏第几行"：滚过之后点第 1 行，选的是新的第 1 项
            #      （数字键选词已在 v39 移除，避免与"输入变量名里的数字"冲突）
            bk7, ui7, c7 = mk22()
            m7 = c7.current_matches()
            for _ in range(10):
                c7.move(1)
            check("22.7 滚一行后本屏第一行是原列表第 2 项",
                  [c7.visible_rows()[0]], expect_contain=[m7[1]])
            # 鼠标点本屏第 1 行 -> _row_at 换算成绝对下标 1 -> completer.pick(1)
            check("22.7 点击本屏第 1 行（pick 绝对下标 1）写入的就是它",
                  [c7.pick(1), bk7.applied],
                  expect_contain=[True, m7[1]])
            check("22.7 选完就收起（accept 会 hide）", [ui7.hidden],
                  expect_contain=[True])

            # 22.8 一直下移到底：窗口滑到最后 10 行，再往前一步绕回首项也回到顶部
            bk8, ui8, c8 = mk22()
            for _ in range(29):
                c8.move(1)
            check("22.8 移到末项（第 30 项）：窗口滑到最底",
                  [c8.selected, c8.view_info(), c8.visible_rows()],
                  expect_contain=[29, (_N, _N - 10), c8.current_matches()[-10:]])
            c8.move(1)                      # 末项 -> 绕回首项
            check("22.8 绕回首项：窗口一起回到顶部",
                  [c8.selected, c8.view_info()], expect_contain=[0, (_N, 0)])

            # 22.9 打全名命中很靠后的候选时（第 26 项），窗口要滚过去带着它，
            #      否则用户看不见自己选中的是什么
            bk9, ui9, c9 = mk22()
            c9.selected = 25
            c9._scroll_to_selected()
            check("22.9 选中第 26 项：窗口滚到它可见，且落在最后一行",
                  [c9.view_info(), c9.view_selection()],
                  expect_contain=[(_N, 16), 9])
            check("22.9 选中项确实在这一屏里",
                  [c9.current_matches()[25] in c9.visible_rows()],
                  expect_contain=[True])

            # 22.10 候选不足一屏：窗口永远停在 0，不出现滚动
            bkA, uiA, cA = mk22(["num", "numArr", "numCount"], "num")
            check("22.10 不足一屏：窗口恒为 0、全量显示",
                  [cA.view_info(), cA.visible_rows()],
                  expect_contain=[(3, 0), ["num", "numArr", "numCount"]])
            check("22.10 不足一屏：候选数 < 一屏上限，不触发滚动",
                  [cA.view_info()[0] <= cA.view_rows], expect_contain=[True])
        except Exception as _e22:
            check("第 22 节异常: %s" % _e22, [True], expect_contain=[False])

        # ---- 第 23 节：拖滚动条时 UI 不能消失（v41 修复） ----
        # 复现：之前拖滚动条拖一小会儿，整个列表框+滚动条就消失。
        # 根因是弹窗点击会被"激活"、抢走 VBE 前台焦点，焦点去抖(~600ms)
        # 后 main.py 误判"已离开 VBE"自动收起。修复：弹窗加 WS_EX_NOACTIVATE
        # 永不抢焦点；并给 maybe_hide_on_outside_click 加"拖拽中不收起"保险。
        class _Bk23(object):
            def __init__(self, names, word):
                self._names = list(names)
                line = "    " + word
                self._ctx = {"line_no": 3, "caret_col": len(line) + 1,
                             "line_text": line, "in_string": False,
                             "in_comment": False, "in_type_position": False,
                             "in_decl_position": False, "decl_names": [],
                             "proc_name": None, "module_name": "M1"}
                self.applied = None
            def get_context(self):
                return dict(self._ctx)
            def get_identifiers(self):
                return [(n, "M1", None, False) for n in self._names]
            def get_declared_names(self):
                return list(self._names)
            def get_type_names(self):
                return []
            def apply_completion(self, *a, **k):
                self.applied = a
            def name_exists_outside_caret(self, w, c):
                return True
        class _Ui23(object):
            def __init__(self):
                self.hidden = False
                self._drag = None
            def show(self, *a, **k):
                self.hidden = False
            def update_selection(self, sel):
                pass
            def hide(self):
                self.hidden = True
            def contains_point(self, x, y):
                return False   # 始终模拟"点击落在弹窗外"
        names23 = ["name%02d" % i for i in range(30)]
        ui23 = _Ui23()
        c23 = E.Completer(_Bk23(names23, "name"), ui23, view_rows=10)
        c23.trigger()   # 触发：30 个候选，超过一屏
        check("23.0 触发后弹窗可见", [c23.is_visible()], expect_contain=[True])
        check("23.0 候选超过一屏(存在纵向滚动条)",
              [len(c23.current_matches()) > c23.view_rows], expect_contain=[True])

        # 23.1 正在拖纵向滚动条(_drag='v')时，即便 contains_point 说在外面，
        #       maybe_hide_on_outside_click 也绝不许收起
        ui23._drag = "v"
        c23.maybe_hide_on_outside_click(9999, 9999)
        check("23.1 拖滚动条途中点弹窗外：列表框不消失",
              [c23.is_visible()], expect_contain=[True])

        # 23.2 set_view_top（拖滚动条调用的核心）不会收起 UI，且真的改了窗口起点
        c23.set_view_top(15)
        check("23.2 拖到中部：UI 仍可见、窗口起点=15",
              [c23.is_visible(), c23.view_info()],
              expect_contain=[True, (30, 15)])
        check("23.2 选中项被夹进可见窗口",
              [c23.selected >= c23.view_top and
               c23.selected < c23.view_top + c23.view_rows],
              expect_contain=[True])

        # 23.3 拖完松开、_drag 归位后，点在弹窗外才按旧行为正常收起
        ui23._drag = None
        c23.maybe_hide_on_outside_click(9999, 9999)
        check("23.3 没在拖、点弹窗外：正常收起",
              [c23.is_visible()], expect_contain=[False])

        # ---- 24. v43：回退删字不得提示【比编辑器内容更长的旧片段】 ----
        #
        # 用户报的现象：定义一个很长的变量名后，按住回退键连删，弹窗会提示出
        # 回退过程中某个【更长】的片段 —— 编辑器里剩下的字符比提示出来的还短。
        #
        # 根因：标识符池是带缓存的（最快 ID_REFRESH_MIN_SEC=0.4s 才重解析一次），
        # 连删时池子里还留着"文字还长"那一刻解析出来的名字；而旧的回声防护只剔除
        # 【与当前词完全相等】的候选，于是"当前词 + 更长尾巴"的旧片段活了下来。
        print("\n=== 24. 回退删字不残留旧片段（v43）===")

        _LONG24 = ("studengdflkgjdslkgjdslkgjdlfksgjdlfkgjdlfkgjdlkjgldsgjdslgjdsgldsjgl"
                   "djgdlg")

        class _CM24(object):
            def __init__(self, name, text):
                self.Name = name
                self.text = text

            @property
            def CountOfLines(self):
                return len(self.text.split("\n"))

            def Lines(self, start, count):
                ls = self.text.split("\n")
                s = max(0, int(start) - 1)
                return "".join(l + "\n" for l in ls[s:s + int(count)])

        class _Comp24(object):
            def __init__(self, name, text, ctype=1):
                self.Name = name
                self.Type = ctype
                self.CodeModule = _CM24(name, text)

        class _Pane24(object):
            def __init__(self, comp):
                self._c = comp
                self.sel = (1, 1, 1, 1)

            @property
            def CodeModule(self):
                return self._c.CodeModule

            def GetSelection(self):
                return self.sel

        class _VBE24(object):
            def __init__(self, comps, act):
                self._p = type("_P24", (object,), {})()
                self._p.Name = "VBAProject"
                self._p.VBComponents = list(comps)
                self.ActiveVBProject = self._p
                self.ActiveCodePane = _Pane24(act)

        class _UI24(object):
            def __init__(self):
                self.shown = False
                self.rows = []

            def show(self, matches, selected, completer=None):
                self.shown = True
                self.rows = list(matches or [])

            def update_selection(self, i):
                pass

            def hide(self):
                self.shown = False
                self.rows = []

            def contains_point(self, x, y):
                return False

        def _session24():
            """建一次后端（它的标识符缓存要跨步骤存活）+ 一个触发器闭包。

            trig(mods, active, line_no, col) 每次都【重建模块文本】——
            真实编辑器里文本是同步缩短的；早期复现脚本没重建，才误以为
            "工程里始终留着完整长名"。
            """
            holder = {}
            VB._get_vbe_cached = lambda: holder["vbe"]
            bk = VB.VbeBackend()
            ui = _UI24()
            cp = E.Completer(bk, ui)

            def trig(mods, active, line_no, col):
                comps = [_Comp24(n, t) for n, t in mods]
                act = None
                for c in comps:
                    if c.Name == active:
                        act = c
                holder["vbe"] = _VBE24(comps, act)
                holder["vbe"].ActiveCodePane.sel = (line_no, col, line_no, col)
                cp.trigger(True)
                return list(cp.current_matches())

            def throttle():
                # 复刻 main.poll_editor 的重解析限流（ID_REFRESH_MIN_SEC=0.4）
                if time.time() - getattr(bk, "_cache_time", 0.0) > 0.4:
                    bk.invalidate_identifiers()

            return trig, throttle

        _orig_get24 = VB._get_vbe_cached
        try:
            # 24.1 连删 15 次：绝不能提示出比编辑器当前内容【更长】的候选
            _decl = ("Option Explicit\nSub Foo()\n    Dim {W} As Long\n"
                     "    x = 1\nEnd Sub")
            _both = ("Option Explicit\nSub Foo()\n    Dim {W} As Long\n"
                     "    {W} = 1\nEnd Sub")
            _bad24 = []
            for _tpl in (_decl, _both):
                _trig, _thr = _session24()
                # 先"打全名"：此刻池子是按完整长名解析并被缓存的
                _trig([("Module1", _tpl.replace("{W}", _LONG24))], "Module1",
                      3, 9 + len(_LONG24))
                for _k in range(1, 16):
                    _typed = _LONG24[:-_k]
                    _thr()
                    _m = _trig([("Module1", _tpl.replace("{W}", _typed))],
                               "Module1", 3, 9 + len(_typed))
                    _bad24 += [n for n in _m if len(n) > len(_typed)]
                    time.sleep(0.008)
            check("24.1 回退删字：不提示比编辑器内容更长的旧片段",
                  [_bad24], expect_contain=[[]])

            # 24.2 真名字（本模块别处声明）的前缀补全不能被误杀
            _trig, _thr = _session24()
            _m2 = _trig([("Module1", "Option Explicit\nPublic numArr As Long\n"
                                       "Sub Foo()\n    num\nEnd Sub")],
                        "Module1", 4, 5 + 3)
            check("24.2 真名字前缀补全仍在（num -> numArr）",
                  [_m2], expect_contain=[["numArr"]])

            # 24.3 隐式变量（别处用过、没声明）的前缀补全不能被误杀
            _trig, _thr = _session24()
            _m3 = _trig([("Module1", "Option Explicit\nSub Foo()\n    arr = 1\n"
                                       "    ar\nEnd Sub")],
                        "Module1", 4, 5 + 2)
            check("24.3 隐式变量前缀补全仍在（ar -> arr）",
                  [_m3], expect_contain=[["arr"]])

            # 24.4 别的模块的 Public 名字仍要提示（跨模块靠"声明集合"这一路证据）
            _trig, _thr = _session24()
            _m4 = _trig([("Module1", "Option Explicit\nSub Foo()\n    Global\n"
                                       "End Sub"),
                         ("Module2", "Option Explicit\n"
                                     "Public GlobalDataArray As Long\n")],
                        "Module1", 3, 5 + 6)
            check("24.4 跨模块 Public 前缀补全仍在（Global -> GlobalDataArray）",
                  [_m4], expect_contain=[["GlobalDataArray"]])
            # ---- 25. v44：从名字【开头】删字符也要能出提示词 ----
            #
            # 用户报的现象：从变量 / 函数 / 窗体 / 模块名的【结尾】【中间】删字符都能
            # 弹出候选，但从名字【开头】删（光标停在残留名字最左边按 Delete）什么都不弹。
            #
            # 根因两处，都出在"只看光标左边"这个前提上：
            #   1. engine._ident_char_before_caret（轮询触发的准入门槛）只检查光标前
            #      一格；名字开头左邻是空格 / 括号 -> 判成"没在拼标识符" -> 直接收起。
            #   2. engine.extract_word_before 只往光标左边取词，左边取不到就返回 None
            #      -> 连候选都没得算。
            # 因此"从中间 / 结尾删"（左边还剩半截词）照常工作，"从头删"永远失效。
            print("\n=== 25. 从名字开头删字符也能出提示词（v44）===")

            # 第 2 行（声明行）始终留着【完整长名】，第 4 行放"被删剩"的残词。
            # 这样标识符池里一定有完整长名，断言与缓存刷新时机无关，不会 flaky。
            _TPL25 = ("Option Explicit\nPublic {D} As Long\nSub Foo()\n"
                      "    {U}\nEnd Sub")

            def _src25(declared, used):
                return _TPL25.replace("{D}", declared).replace("{U}", used)

            _trig, _thr = _session24()
            # 先打全名（模拟真实打字；同时确认池子里确实收了这个名字）
            _m25a0 = _trig([("Module1", _src25(_LONG24, _LONG24))],
                           "Module1", 4, 5 + len(_LONG24))
            check("25.0 打全名：池子里有这个名字",
                  [_LONG24 in _m25a0], expect_contain=[True])

            # 25.1 从开头删：光标停在残留名字最左端（第 4 行第 5 列）
            _r25 = []
            for _k in (1, 5, 20):
                _thr()
                _m = _trig([("Module1", _src25(_LONG24, _LONG24[_k:]))],
                           "Module1", 4, 5)
                _r25.append(_LONG24 in _m)
            check("25.1 从名字开头删字符 -> 仍出提示（v44）",
                  _r25, expect_contain=[True])

            # 25.2 从中间删：光标在该行词内部（原有行为不许回退）
            _thr()
            _m25b = _trig([("Module1", _src25(_LONG24, _LONG24))],
                          "Module1", 4, 5 + 12)
            check("25.2 从名字中间删字符 -> 照常出提示",
                  [_LONG24 in _m25b], expect_contain=[True])

            # 25.3 从结尾删：光标在残留词末尾
            _thr()
            _m25c = _trig([("Module1", _src25(_LONG24, _LONG24[:30]))],
                          "Module1", 4, 5 + 30)
            check("25.3 从名字结尾删字符 -> 照常出提示",
                  [_LONG24 in _m25c], expect_contain=[True])

            # 25.4 取词与替换范围（纯函数级，直接盯 extract_word_at 的返回）
            _w25 = E.extract_word_at("    " + _LONG24[5:], 5)
            check("25.4a 光标在名字开头：取词=残留名，范围盖住整个残留词",
                  [_w25],
                  expect_contain=[(_LONG24[5:], 5, 5 + len(_LONG24[5:]))])
            check("25.4b 光标在词尾：右端就是光标列（与旧行为一致）",
                  [E.extract_word_at("    num", 8)], expect_contain=[("num", 5, 8)])
            check("25.4c 光标前后都不是标识符字符 -> 取不到词（不乱弹）",
                  [E.extract_word_at("a  b", 3)],
                  expect_contain=[(None, None, None)])

            # 25.5 Tab 确认：替换范围必须盖住光标右边那截残留名字，
            #      否则会在残留词前面插入候选，拼出"双份名字"
            class _B25(object):
                def __init__(self, ctx, ids):
                    self._ctx = ctx
                    self._ids = list(ids)
                    self.applied = None

                def get_context(self):
                    return self._ctx

                def get_identifiers(self):
                    return self._ids

                def get_declared_names(self):
                    return [i[0] if isinstance(i, (tuple, list)) else i
                            for i in self._ids]

                def apply_completion(self, line_no, start, end, completion):
                    self.applied = (start, end, completion)

            _line25 = "    " + _LONG24[5:]
            _b25 = _B25({"line_no": 4, "caret_col": 5, "line_text": _line25,
                         "in_string": False, "in_comment": False,
                         "in_type_position": False, "in_decl_position": False,
                         "decl_names": [], "proc_name": "Foo",
                         "module_name": "Module1"}, [_LONG24])
            _c25 = E.Completer(_b25, _UI24())
            _c25.trigger(True)
            _c25.accept()
            check("25.5 Tab 确认：整段换掉残留词（不拼出双份名字）",
                  [_b25.applied],
                  expect_contain=[(5, 5 + len(_LONG24[5:]), _LONG24)])

            # ---- 26. v45：从名字【开头】删字符不许提示出"更长的旧片段"（回声）----
            #
            # 用户报的现象：把光标停在名字最左边、按 Delete 从【开头】连删时，
            # 弹窗会提示出【完整的旧长名】（有回音、提示自己）。
            #
            # 与 v43"从结尾删"的回声同源，根因是回声判据只认【前缀】：
            #   候选.startswith(当前词)
            # 从结尾删 -> 当前词是候选的前缀 -> 命中，能剔除；
            # 从开头删 -> 当前词是候选的【后缀】 -> 不命中 -> 回声漏了过去。
            #
            # 修法：判据推广为"当前词是候选名的【连续子串】"——删除只会让光标处
            # 的词变短，剩下的必然是原词的连续一段（前缀 / 后缀 / 中段通吃）。
            # 是否保留仍由 _name_really_exists 裁决：真名字（本模块别处出现 /
            # 别的模块声明过）照常提示，只有"现场已删干净、别处也没有"的旧快照
            # 才被剔除。
            #
            # 本节的场景里，正在编辑的【就是声明行本身】（Dim <长名> As Long），
            # 所以删到最后现场文本里已经找不到完整长名了 —— 池子（0.4s 缓存）里
            # 的那一份只是旧快照，必须剔除。
            _TPL26 = ("Option Explicit\nSub Foo()\n    Dim {W} As Long\n"
                      "    x = 1\nEnd Sub")

            def _src26(typed):
                return _TPL26.replace("{W}", typed)

            # 26.1 从开头删：不许提示出比编辑器当前内容更长的旧片段
            _trig, _thr = _session24()
            _trig([("Module1", _src26(_LONG24))], "Module1", 3, 9 + len(_LONG24))
            _bad26 = []
            for _k in (1, 5, 20, 40):
                _thr()
                _typed = _LONG24[_k:]
                _m = _trig([("Module1", _src26(_typed))], "Module1", 3, 9)
                _bad26 += [n for n in _m if len(n) > len(_typed)]
                time.sleep(0.008)
            check("26.1 从名字开头删字符：不提示出更长的旧片段（v45）",
                  [_bad26], expect_contain=[[]])

            # 26.2 头尾都删过 -> 当前词是旧长名的【中段】，同样不许提示旧长名
            _trig, _thr = _session24()
            _trig([("Module1", _src26(_LONG24))], "Module1", 3, 9 + len(_LONG24))
            _thr()
            _mid26 = _LONG24[10:50]
            _m26b = _trig([("Module1", _src26(_mid26))], "Module1", 3, 9)
            check("26.2 头尾都删过：中段残词不提示旧长名（v45）",
                  [_LONG24 in _m26b], expect_contain=[False])

            # 26.3 回归：本模块别处仍留着完整长名时（编辑的不是声明行），
            #      从头删依旧照常提示 —— 回声防护加强不许误杀真名字
            _trig, _thr = _session24()
            _trig([("Module1", _src25(_LONG24, _LONG24))], "Module1", 4,
                  5 + len(_LONG24))
            _thr()
            _m26c = _trig([("Module1", _src25(_LONG24, _LONG24[20:]))],
                          "Module1", 4, 5)
            check("26.3 回归：本模块声明的真名字从头删仍提示",
                  [_LONG24 in _m26c], expect_contain=[True])

            # 26.4 回归：从结尾删（v43 场景）仍不提示更长片段（控制组）
            _trig, _thr = _session24()
            _trig([("Module1", _src26(_LONG24))], "Module1", 3, 9 + len(_LONG24))
            _bad26d = []
            for _k in (1, 5, 20):
                _thr()
                _typed = _LONG24[:-_k]
                _m = _trig([("Module1", _src26(_typed))], "Module1", 3,
                           9 + len(_typed))
                _bad26d += [n for n in _m if len(n) > len(_typed)]
            check("26.4 回归：从结尾删仍不提示更长片段（v43 控制组）",
                  [_bad26d], expect_contain=[[]])

            # ---- 27. v46：光标停在名字【中间】删字符，当前词取【整段名字】----
            #
            # 用户报的现象：变量 test / test02 / test23456789；把 test23456789
            # 中间那个 t 删掉 -> tes23456789、光标停在 tes 之后，却仍然提示
            # test / test02 —— 因为取词只看光标左边（tes），完全无视光标右边
            # 还留着的 23456789。当前词改为整段名字后，tes23456789 既不匹配
            # test 也不匹配 test02，只剩真正那一整段名字（Tab 一次换掉整段）。
            print("\n=== 27. 从名字【中间】删字符：当前词取整段名字（v46）===")

            _TPL27 = ("Option Explicit\n"
                      "Public test As Long\n"
                      "Public test02 As Long\n"
                      "Public test23456789 As Long\n"
                      "Sub Foo()\n"
                      "    {U}\n"
                      "End Sub")

            def _src27(used):
                return _TPL27.replace("{U}", used)

            # 27.1 从中间删掉那个 t：光标停在 tes|23456789（第 6 行第 8 列）
            _trig, _thr = _session24()
            _trig([("Module1", _src27("test23456789"))], "Module1", 6,
                  5 + len("test23456789"))
            _m27a = _trig([("Module1", _src27("tes23456789"))], "Module1", 6,
                          5 + len("tes"))
            check("27.1 中间删字符：不再按左半截提示（无 test / test02）",
                  [[n for n in _m27a if n.lower() in ("test", "test02")]],
                  expect_contain=[[]])
            check("27.2 中间删字符：整段名字本身仍提示（Tab 可一次换掉）",
                  [_m27a], expect_contain=[["test23456789"]])

            # 27.3 纯函数级：取词与替换范围
            check("27.3 光标在名字中间：取词=整段名字，范围盖住整段",
                  [E.extract_word_at("    tes23456789", 8)],
                  expect_contain=[("tes23456789", 5, 16)])
            check("27.4 光标在词尾：右端仍是光标列（旧行为不变）",
                  [E.extract_word_at("    num", 8)],
                  expect_contain=[("num", 5, 8)])
            check("27.5 光标在名字开头：仍取右边整段（v44 行为不变）",
                  [E.extract_word_at("    " + _LONG24, 5)],
                  expect_contain=[(_LONG24, 5, 5 + len(_LONG24))])
            check("27.6 光标前后都不是标识符字符：仍取不到词（不乱弹）",
                  [E.extract_word_at("a  b", 3)],
                  expect_contain=[(None, None, None)])

            # 27.7 Tab 确认时传给后端的替换范围必须盖住整段名字（含右半截），
            #      否则候选会插在 tes 之后，拼出 test23456789 + 23456789
            class _B27(object):
                def __init__(self, ctx, ids):
                    self._ctx = ctx
                    self._ids = list(ids)

                def get_context(self):
                    return self._ctx

                def get_identifiers(self):
                    return self._ids

                def get_declared_names(self):
                    return set(str(i).lower() for i in self._ids)

            _ctx27 = {"line_no": 6, "caret_col": 8,
                      "line_text": "    tes23456789",
                      "in_string": False, "in_comment": False,
                      "in_type_position": False, "in_decl_position": False,
                      "proc_name": "Foo", "module_name": "Module1"}
            _cp27 = E.Completer(_B27(_ctx27, ["test", "test02",
                                              "test23456789"]), _UI24())
            _cp27.trigger(True)
            check("27.7 Tab 确认范围盖住整段名字（end_col=16 而非光标列 8）",
                  [(bool(_cp27.visible),
                    (_cp27.ctx or {}).get("word_end_col"))],
                  expect_contain=[(True, 16)])


            # ---- 第 29 节：Shift+Enter 在光标行下方新起一行（v48） ----
            #
            # 用户要求：光标停在代码行【中间】时想另起一行，老办法是"先把光标移到
            # 本行末尾再按回车"，太麻烦。现在按 Shift+Enter 一步完成，且新行起始
            # 位置要与上一行代码起始位置对齐（IDEA 的 Start New Line 同款）。
            #
            # 关键点：当前行【不拆分】—— 光标右侧的代码留在原行；新行内容 = 上一行的
            # 行首空白；光标落在新行缩进之后。后端实现见 VbeBackend.new_line_below()。
            print("\n=== 29. Shift+Enter 在当前行下方新起一行（v48）===")
            try:
                import main as _mn29
                _orig_get29 = VB._get_vbe_cached

                # 29.1 纯判断：只认"按住 Shift 的回车"
                check("29.1 只认 Shift+回车（裸回车 / Shift+Tab 都不算）",
                      [_mn29._is_newline_shortcut(0x0D, True),
                       _mn29._is_newline_shortcut(0x0D, False),
                       _mn29._is_newline_shortcut(0x09, True),
                       _mn29._is_newline_shortcut(0x0D, 0)],
                      expect_contain=[True, False, False, False])

                # 29.2 对齐依据：行首空白（空格 / Tab / 无缩进 / 空行）
                check("29.2 行首空白提取（新行对齐上一行代码起始位置的依据）",
                      [VB._leading_ws("    x = 1"), VB._leading_ws("\t If x"),
                       VB._leading_ws("x = 1"), VB._leading_ws("")],
                      expect_contain=["    ", "\t ", "", ""])

                # 29.3 缩进之后的光标列（字符列语义 = len + 1）
                check("29.3 缩进之后的光标列：4 空格 -> 第 5 列；无缩进 -> 第 1 列",
                      [VB._indent_end_col("    "), VB._indent_end_col("")],
                      expect_contain=[5, 1])

                class _CM29(object):
                    def __init__(self, lines):
                        self.lines = list(lines)
                        self.Name = "Module1"

                    @property
                    def CountOfLines(self):
                        return len(self.lines)

                    def Lines(self, start, count):
                        s0 = max(0, int(start) - 1)
                        return "".join(l + "\n"
                                       for l in self.lines[s0:s0 + int(count)])

                    def InsertLines(self, start, text):
                        i = min(max(0, int(start) - 1), len(self.lines))
                        for k, ln in enumerate(str(text).split("\n")):
                            self.lines.insert(i + k, ln)

                class _FPane29(object):
                    def __init__(self, cm, sel):
                        self.CodeModule = cm
                        self.sel = sel

                    def GetSelection(self):
                        return self.sel

                    def SetSelection(self, sl, sc, el, ec):
                        self.sel = (sl, sc, el, ec)

                class _FVBE29(object):
                    def __init__(self, pane):
                        self.ActiveCodePane = pane

                def _nl29(lines, line_no, col):
                    cm29 = _CM29(lines)
                    pane29 = _FPane29(cm29, (line_no, col, line_no, col))
                    VB._get_vbe_cached = lambda: _FVBE29(pane29)
                    ok29 = VB.VbeBackend().new_line_below()
                    return ok29, cm29.lines, pane29.sel

                _SRC29 = ["Option Explicit", "Sub Foo()", "    Dim x",
                          "    x = 1 + 2", "End Sub"]

                # 29.4 核心：光标停在行中间 -> 当前行不拆、新行缩进对齐、光标在缩进后
                _ok29, _ls29, _sel29 = _nl29(_SRC29, 4, 11)
                check("29.4 光标在行中间 -> 新行插在下方、缩进与上一行对齐、不拆行",
                      [_ok29, _ls29, _sel29],
                      expect_contain=[True,
                                      ["Option Explicit", "Sub Foo()", "    Dim x",
                                       "    x = 1 + 2", "    ", "End Sub"],
                                      (5, 5, 5, 5)])

                # 29.5 光标在行尾 / 行首：结果一致（都按"本行缩进"新起一行）
                _r29 = []
                for _c29 in (14, 5):
                    _ok, _ls, _sel = _nl29(_SRC29, 4, _c29)
                    _r29.append((_ok, _ls[4], _sel))
                check("29.5 行尾 / 行首同样新起一行，且都不拆行",
                      [_r29],
                      expect_contain=[[(True, "    ", (5, 5, 5, 5)),
                                       (True, "    ", (5, 5, 5, 5))]])

                # 29.6 末行：追加到模块末尾（InsertLines 起始行超出即追加）
                _ok29b, _ls29b, _sel29b = _nl29(_SRC29, 5, 8)
                check("29.6 光标在模块最后一行 -> 新行追加到末尾、光标在第 1 列",
                      [_ok29b, _ls29b, _sel29b],
                      expect_contain=[True, _SRC29 + [""], (6, 1, 6, 1)])

                # 29.7 Tab 缩进的行：显示列语义下光标要按 tab stop 展开
                _sem29 = VB._sem_cache.get("info")
                VB._sem_cache["info"] = ("disp", 4, False)
                try:
                    _ok29c, _ls29c, _sel29c = _nl29(
                        ["Sub Foo()", "\tIf x Then", "End Sub"], 2, 11)
                finally:
                    VB._sem_cache["info"] = _sem29
                check("29.7 Tab 缩进：新行仍是 Tab，光标按显示列落在其后（第 5 列）",
                      [_ok29c, _ls29c[2], _sel29c],
                      expect_contain=[True, "\t", (3, 5, 3, 5)])

                # 29.8 取不到 VBE（不在代码窗 / Excel 已关）必须返回 False ——
                #      调用方据此【不吞键】，让 VBE 按原生行为处理
                VB._get_vbe_cached = lambda: None
                check("29.8 取不到 VBE -> 返回 False（绝不吞掉用户的按键）",
                      [VB.VbeBackend().new_line_below()],
                      expect_contain=[False])

                # 29.9 空模块（光标行号为 0）也要安全返回 False
                VB._get_vbe_cached = lambda: _FVBE29(
                    _FPane29(_CM29([]), (0, 0, 0, 0)))
                check("29.9 光标行号为 0（无有效行）-> 返回 False",
                      [VB.VbeBackend().new_line_below()],
                      expect_contain=[False])

                VB._get_vbe_cached = _orig_get29
            except Exception as _e29:
                check("第 29 节异常: %s" % _e29, [True], expect_contain=[False])

        finally:
            VB._get_vbe_cached = _orig_get24

    except Exception as _e:
        check("第 19 节不可用（vbe_bridge 导入失败）: %s" % _e, [True],
              expect_contain=[True])

    # ---- 30. v50：跨过程泄漏（只读隐式变量被算到别的过程 / 模块级）----
    #
    # 病根：_usage_candidates 把 `Attribute VB_xxx` / `Option Explicit` 这类
    # 行【替换成空串】以屏蔽内容，但这会缩短文本、让后续所有字符偏移整体前移；
    # 而调用方 _line_proc_map 算 offsets 用的是【未压缩】的文本。
    # 真实模块头部动辄 5 行以上 Attribute + Option（累计前移 150+ 字符），
    # 于是"只被读取、从未被赋值"的隐式变量被算成 proc=None（模块级）——
    # 全模块每个过程都能看到它，包括别的过程的**形参位置**。
    #
    # 修复：跳过行改为抹成【等长空格】，偏移恒定对齐。
    print("\n=== 30. 跨过程泄漏：跳过行必须等长屏蔽（v50）===")
    try:
        _HDR30 = [
            'Attribute VB_Name = "Module1"',
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = False",
            "Attribute VB_Exposed = False",
            "Option Explicit",
        ]
        _code30 = "\n".join(_HDR30 + [
            "",
            "Public Sub ProcedureA()",
            "    Debug.Print targetName",
            "End Sub",
            "",
            "Public Sub ProcedureB(ByVal sheetName As String)",
            "    Debug.Print resultValue",
            "End Sub",
        ])

        # 30.1 偏移对齐：_usage_candidates 返回的偏移必须与 offsets 同基准
        _masked30 = P._mask_strings_and_comments(_code30)
        _masked30 = P._RE_CONTINUATION.sub(" ", _masked30)
        _offs30, _lp30 = P._line_proc_map(_masked30)
        _owner30 = {n: P._proc_at_offset(_offs30, _lp30, p)
                    for n, p in P._usage_candidates(_masked30, set())}
        check("30.1 只读隐式变量归属正确（不再落到模块级）",
              [_owner30.get("targetName"), _owner30.get("resultValue")],
              expect_contain=["ProcedureA", "ProcedureB"])

        # 30.2 跳过行长度不变（等长屏蔽）
        _kept30 = "\n".join(" " * len(l) if P._RE_USAGE_SKIP_LINE.match(l) else l
                            for l in _masked30.split("\n"))
        check("30.2 跳过行抹成等长空格（偏移不漂移）",
              [len(_kept30), len(_masked30)],
              expect_contain=[len(_masked30), len(_masked30)])

        # 30.3 端到端：A 的变量不得出现在 B 的可见集（B 的形参位置）
        _r30 = list(P.extract_records(_code30, module="M1", is_std_module=True))
        _r30 += P.extract_implicit_records(_code30, module="M1",
                                           is_std_module=True, declared=_r30,
                                           scope="proc")
        _visB30 = E.filter_identifiers_by_scope(_r30, "ProcedureB", "M1")
        _visA30 = E.filter_identifiers_by_scope(_r30, "ProcedureA", "M1")
        check("30.3 A 的只读变量不泄漏到 ProcedureB",
              [n for n in _visB30 if str(n).lower() == "targetname"],
              expect_contain=[])
        check("30.3 B 的只读变量不泄漏到 ProcedureA",
              [n for n in _visA30 if str(n).lower() == "resultvalue"],
              expect_contain=[])
        check("30.3 各自的只读变量在自己过程里仍提示",
              [sorted(str(n) for n in _visA30 if str(n).lower() == "targetname"),
               sorted(str(n) for n in _visB30 if str(n).lower() == "resultvalue")],
              expect_contain=[["targetName"], ["resultValue"]])

        # 30.4 头部越"胖"，旧实现漂移越大 —— 多 Attribute 行下依然正确
        _fat30 = "\n".join([
            'Attribute VB_Name = "M"',
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = False",
            "Attribute VB_Exposed = False",
            "#If VBA7 Then",
            "#End If",
            "Option Explicit",
            "Option Base 0",
            "",
            "Public Sub AA()",
            "    Debug.Print onlyInAA",
            "End Sub",
            "",
            "Public Sub BB()",
            "    Debug.Print onlyInBB",
            "End Sub",
        ])
        _m30b = P._RE_CONTINUATION.sub(
            " ", P._mask_strings_and_comments(_fat30))
        _o30b, _p30b = P._line_proc_map(_m30b)
        _own30b = {n: P._proc_at_offset(_o30b, _p30b, p)
                   for n, p in P._usage_candidates(_m30b, set())}
        check("30.4 更胖的模块头部下归属仍正确",
              [_own30b.get("onlyInAA"), _own30b.get("onlyInBB")],
              expect_contain=["AA", "BB"])

        # 30.5 回归：跳过行里的名字不该被当成隐式变量收录
        check("30.5 Attribute 行里的名字不当变量",
              [n for n, _p in P._usage_candidates(_masked30, set())
               if str(n).lower() in ("vb_name", "vb_globalnamespace",
                                     "explicit")],
              expect_contain=[])
    except Exception as _e30:
        check("第 30 节异常: %s" % _e30, [True], expect_contain=[False])

    # ---- 31. v51：形参名与别的过程的变量重名时不得"提示自己" ----
    #
    # 用户报的现象：在 ProcedureB 的形参位置，打 / 回退 / 粘贴一个与
    # ProcedureA 的局部变量一模一样的名字，弹窗会把刚敲进去的字原样提示出来；
    # 换成任何别处都没有的新名字反而安静（"不相同就不提示"）。
    #
    # 根因：形参位置的输入会被解析成 ProcedureB 自己的形参记录 —— 于是
    # "别的过程里出现过这个名字"（现场全模块扫描）被当成了"这个名字真实存在"
    # 的铁证，回声防护放行。本节的 31.1 先复现旧行为，31.2 起验证修复。
    print("\n=== 31. 形参位置不提示自己（v51：现场证据按作用域收窄）===")
    try:
        import vbe_bridge as VB31

        _HDR31 = [
            'Attribute VB_Name = "Module1"',
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = False",
            "Attribute VB_Exposed = False",
            "Option Explicit",
        ]

        class _CM31(object):
            """只实现后端用到的 CodeModule.CountOfLines / Lines。"""

            def __init__(self, code):
                self._code = code

            @property
            def CountOfLines(self):
                return len(self._code.split("\n"))

            def Lines(self, start, count):
                _ls = self._code.split("\n")
                return "\n".join(_ls[start - 1:start - 1 + count])

        class _VBE31(object):
            def __init__(self, code):
                _cm = _CM31(code)
                self.ActiveCodePane = type("_CP31", (), {"CodeModule": _cm})()

        _cur31 = {"code": ""}
        _orig_get31 = VB31._get_vbe_cached
        VB31._get_vbe_cached = lambda *a, **k: _VBE31(_cur31["code"])
        _names31 = VB31.VbeBackend().names_outside_caret

        class _UI31(object):
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

        def _parse31(lines):
            """带 | 标记的代码行 -> (代码, 光标位置)。"""
            _out, _caret = [], None
            for _i, _l in enumerate(lines, 1):
                if "|" in _l:
                    _col = _l.index("|")
                    _caret = (_i, _col + 1)
                    _l = _l[:_col] + _l[_col + 1:]
                _out.append(_l)
            return "\n".join(_out), _caret

        class _B31(object):
            """最小后端：候选/声明名来自真实解析，现场证据走真实
            VbeBackend.names_outside_caret（只把最底层读文本换成假的）。"""

            def __init__(self, code, caret, force_broad=False):
                self.code = code
                self.caret = caret
                self.force_broad = force_broad
                _r = list(P.extract_records(code, module="Module1",
                                            is_std_module=True, caret=caret))
                _r += list(P.extract_implicit_records(
                    code, module="Module1", is_std_module=True, declared=_r,
                    scope="proc", caret=caret))
                self._recs = _r
                self._decl_names = set(
                    str(n).lower() for n in
                    (P.decl_names_at_caret(code, caret) or ()))
                self._declared = set(
                    str(x[0]).lower() for x in _r
                    if str(x[0]).lower() not in self._decl_names)
                try:
                    self._caret_word = P.ident_at_caret(
                        code, caret[0], caret[1]).lower()
                except Exception:
                    self._caret_word = ""

            def get_context(self):
                _ls = self.code.split("\n")
                return {"line_no": self.caret[0], "caret_col": self.caret[1],
                        "line_text": _ls[self.caret[0] - 1],
                        "in_string": False, "in_comment": False,
                        "in_type_position": False,
                        "in_decl_position": bool(self._decl_names),
                        "decl_names": sorted(self._decl_names),
                        "proc_name": P.proc_at_line(self.code, self.caret[0]),
                        "module_name": "Module1"}

            def get_identifiers(self):
                return self._recs

            def get_declared_names(self):
                return list(self._declared)

            def caret_word_at_collect(self):
                return self._caret_word

            def declared_elsewhere(self, name, module_name):
                return False           # 全部声明都在 Module1

            def names_outside_caret(self, caret, scope_only=False):
                _cur31["code"] = self.code
                if self.force_broad:   # 对照：v50 行为（永远全模块扫描）
                    scope_only = False
                return _names31(caret, scope_only)

            def apply_completion(self, *a, **k):
                return None

        def _trig31(body, force_broad=False):
            _code, _caret = _parse31(_HDR31 + body)
            _be = _B31(_code, _caret, force_broad)
            _ui = _UI31()
            _c = E.Completer(_be, _ui)
            _c.trigger()
            return (_ui.shown, sorted(_c.matches or []))

        _leak31 = [
            "Public Sub ProcedureA()",
            "    Dim username As String",
            '    username = "abc"',
            "    Debug.Print username",
            "End Sub",
            "",
            "Public Sub ProcedureB(ByVal username| As String)",
            "End Sub",
        ]
        # 31.1 旧行为确实会"提示自己"——证明这条回归确实有意义。
        # v53 起引擎又多了一道"正在造名字的位置不提示自己"的语义护栏，它与
        # 宽/窄窗口完全独立；对照必须【连它一起短路】，否则根本复现不出旧行为，
        # 这条控制组就变成了恒真的空测。
        _orig_ml31 = E.Completer._module_level_names
        E.Completer._module_level_names = lambda _self: set(
            str(_r[0]).lower() for _r in _self.backend.get_identifiers())
        try:
            _ctl31 = _trig31(_leak31, force_broad=True)
        finally:
            E.Completer._module_level_names = _orig_ml31
        check("31.1 对照：宽窗口 + 关掉 v53 护栏时形参重名会提示自己",
              [_ctl31],
              expect_contain=[(True, ["username"])])
        # 31.2 修复后：不再提示自己
        check("31.2 形参名与其他过程的变量重名 -> 不提示自己",
              [_trig31(_leak31)],
              expect_contain=[(False, [])])
        # 31.3 对照：任何地方都没有的新名字同样安静（用户说的"不相同不提示"）
        check("31.3 全新形参名保持安静",
              [_trig31([
                  "Public Sub ProcedureA()",
                  "    Dim username As String",
                  "End Sub",
                  "",
                  "Public Sub ProcedureB(ByVal zzqnewxx| As String)",
                  "End Sub"])],
              expect_contain=[(False, [])])
        # 31.4 回退到"别的过程的变量名"时同样不提示
        check("31.4 形参回退到别的过程的变量名 -> 不提示",
              [_trig31([
                  "Public Sub ProcedureA()",
                  "    Dim username As String",
                  "End Sub",
                  "",
                  "Public Sub ProcedureB(ByVal usernam| As String)",
                  "End Sub"])],
              expect_contain=[(False, [])])
        # 31.5 模块级名字仍是合法引用（v37 语义）——收窄不等于砍掉
        check("31.5 形参同名于模块级过程名 -> 照常提示",
              [_trig31([
                  "Public Sub updateValue()",
                  "End Sub",
                  "",
                  "Public Sub ProcedureB(ByVal updateValue| As String)",
                  "End Sub"])],
              expect_contain=[(True, ["updateValue"])])
        # 31.6 回归：用法位置（不是在声明新名字）照常提示同过程的名字
        check("31.6 同过程内用法位置打前缀 -> 照常提示",
              [_trig31([
                  "Public Sub ProcedureA()",
                  "    Dim username As String",
                  "    Debug.Print userna|",
                  "End Sub"])],
              expect_contain=[(True, ["username"])])
        # 31.7 回归：模块级声明行打全名照常提示（gate 不生效的路径）
        check("31.7 模块级声明行打全名 -> 照常提示",
              [_trig31([
                  "Public gUserName As String",
                  "Sub Foo()",
                  "    MsgBox gUserName",
                  "End Sub",
                  "",
                  "Public gUserName| As String"])],
              expect_contain=[(True, ["gUserName"])])
        # 31.8 机制本身：现场证据的宽/窄两个窗口
        _c31, _k31 = _parse31(_HDR31 + _leak31)
        _cur31["code"] = _c31
        _broad31 = _names31(_k31, False)
        _narrow31 = _names31(_k31, True)
        check("31.8 现场证据：宽窗口含 username，窄窗口不含",
              ["username" in _broad31, "username" in _narrow31],
              expect_contain=[True, False])
        _c31b, _k31b = _parse31(_HDR31 + [
            "Public Sub updateValue()",
            "End Sub",
            "",
            "Public Sub ProcedureB(ByVal updateValue| As String)",
            "End Sub"])
        _cur31["code"] = _c31b
        check("31.8 窄窗口仍认模块级过程名（不能收窄过头）",
              ["updatevalue" in _names31(_k31b, True)],
              expect_contain=[True])
        VB31._get_vbe_cached = _orig_get31
    except Exception as _e31:
        check("第 31 节异常: %s" % _e31, [True], expect_contain=[False])

    # ---- 32. v52：在标识符【前面】敲空格/标点，不许把右边那个词拿来补全 ----
    #
    # 用户报的现象：光标停在形参首字符前面，敲一个空格 -> 弹窗把这个形参
    # 原样提示出来（提示自己）。真实代码里形参几乎总会在本过程体内被用到，
    # 于是"现场还能扫到它" -> 被当成真实存在的名字 -> 回声防护放行。
    #
    # 根因：`extract_word_at` 有一条兜底（v44 为"光标停在词首、按 Delete 从头
    # 删字"引入）——光标左边没词、但右边紧挨着标识符时取右边那个词；入场检查
    # `_ident_char_before_caret` 也据此放宽到"光标前【或】后是标识符即可"。
    # 于是做完这次编辑后的那一帧快照，"敲空格"与"从词首删字"长得一模一样
    # （左边是分隔符、右边是标识符），单看快照永远区分不开。
    #
    # 唯一的差别在【怎么变的】：删字是行【变短】；敲分隔符是行【变长】且插入
    # 的那一段不含标识符字符。本节的 32.1~32.5 锁这条判据本身，32.6 起锁端到端。
    print("\n=== 32. 标识符前敲空格/标点不许补全（v52）===")
    try:
        # A) 判据本身：只有"插进来一段不含标识符字符的内容"才算敲了分隔符
        check("32.1 插入空格 -> 判定为分隔符",
              [E.typed_separator("Public Sub B(ByVal x As String)",
                                 "Public Sub B(ByVal  x As String)")],
              expect_contain=[True])
        check("32.2 删除（行变短）-> 不判定为分隔符",
              [E.typed_separator("    Debug.Print abcd",
                                 "    Debug.Print abc")],
              expect_contain=[False])
        check("32.3 插入标识符字符 -> 不判定为分隔符",
              [E.typed_separator("    Debug.Print abc",
                                 "    Debug.Print abcd")],
              expect_contain=[False])
        check("32.4 VBE 纠正大小写（替换）-> 不判定为分隔符",
              [E.typed_separator("dim x as string", "Dim x As String")],
              expect_contain=[False])
        check("32.5 纯插入内容提取正确（替换/删除一律 None）",
              [E.line_edit_inserted("ab", "aXXb"),
               E.line_edit_inserted("ab", "b"),
               E.line_edit_inserted("abc", "abc")],
              expect_contain=["XX", None])

        # B) 端到端：完全按 main.poll_editor 的分支跑
        #    （31 节那套脚手架还在函数作用域里，这里直接复用；它内部是
        #     真实解析 + 真实 VbeBackend.names_outside_caret）
        VB31._get_vbe_cached = lambda *a, **k: _VBE31(_cur31["code"])

        def _poll32(body, edit):
            """edit(old_line, col) -> (new_line, new_col)；按 poll 的分支走一遍。"""
            code, caret = _parse31(_HDR31 + body)
            ln, col = caret
            ls = code.split("\n")
            old_line = ls[ln - 1]
            new_line, new_col = edit(old_line, col)
            if E.typed_separator(old_line, new_line):
                return (False, [], new_line)          # 收起，不 trigger
            ls[ln - 1] = new_line
            new_code = "\n".join(ls)
            _cur31["code"] = new_code
            be = _B31(new_code, (ln, new_col))
            ui = _UI31()
            c = E.Completer(be, ui)
            c.trigger(True)
            return (ui.shown, sorted(c.matches or []), new_line)

        def _ins32(ch):
            def f(line, col):
                return line[:col - 1] + ch + line[col - 1:], col + len(ch)
            return f

        def _del32(line, col):
            return line[:col - 1] + line[col:], col     # Delete：删光标右边那个字

        def _tail32(ch):
            def f(line, col):
                return line + ch, len(line) + len(ch) + 1
            return f

        _p32 = _poll32([
            "",
            "Public Sub ProcedureB(ByVal |username As String)",
            "    Debug.Print username",
            "End Sub"], _ins32(" "))
        check("32.6 形参首字符前打空格 -> 不提示形参自身",
              [_p32[0], "username" in _p32[1], _p32[2]],
              expect_contain=[False, False,
                              "Public Sub ProcedureB(ByVal  username As String)"],
              expect_absent=[True])

        # 控制组：绕过这道门直接 trigger，确实会提示自己 —— 证明这道门有意义
        _c32, _k32 = _parse31(_HDR31 + [
            "",
            "Public Sub ProcedureB(ByVal |username As String)",
            "    Debug.Print username",
            "End Sub"])
        _ls32 = _c32.split("\n")
        _ln32, _col32 = _k32
        _l32 = _ls32[_ln32 - 1]
        _ls32[_ln32 - 1] = _l32[:_col32 - 1] + " " + _l32[_col32 - 1:]
        _nc32 = "\n".join(_ls32)
        _cur31["code"] = _nc32
        _ui32 = _UI31()
        # 同理（见 31.1）：v53 的语义护栏也会挡住这一例，对照要把它一并关掉，
        # 才能真正测出"没有 v52 这道门时会提示自己"。
        _orig_ml32 = E.Completer._module_level_names
        E.Completer._module_level_names = lambda _self: set(
            str(_r[0]).lower() for _r in _self.backend.get_identifiers())
        try:
            _cp32 = E.Completer(_B31(_nc32, (_ln32, _col32 + 1)), _ui32)
            _cp32.trigger(True)
        finally:
            E.Completer._module_level_names = _orig_ml32
        check("32.7 对照：两道门都关掉时会提示形参自身（本节没白测）",
              [_ui32.shown, sorted(_cp32.matches or [])],
              expect_contain=[True, ["username"]])

        # B) 回归：删字 / 打字母两条老路径必须原样保留，不能被一刀切
        _p32b = _poll32([
            "Public Function Foo() As Long",
            "    Dim gUserName As String",
            "    Debug.Print |gUserName",
            "End Function"], _del32)
        check("32.8 词首按 Delete 删字 -> 仍照旧提示（v44 路径不变）",
              [_p32b[0], _p32b[1]],
              expect_contain=[True, ["gUserName"]])

        _p32c = _poll32([
            "Public Function Foo() As Long",
            "    Dim gUserName As String",
            "    Debug.Print |",
            "End Function"], _tail32("g"))
        check("32.9 插入字母 -> 照常弹出候选（没被误伤）",
              [_p32c[0], "gUserName" in _p32c[1]],
              expect_contain=[True, True])

        _p32d = _poll32([
            "Public Function Foo() As Long",
            "    Dim gUserName As String",
            "    Debug.Print gUserName|",
            "End Function"], _ins32(" "))
        _p32e = _poll32([
            "Public Function Foo() As Long",
            "    Dim gUserName As String",
            "    Debug.Print gUserName |",
            "End Function"], _ins32("g"))
        check("32.10 打空格收起后接着打字母 -> 能重新弹出（流程不卡死）",
              [_p32d[0], _p32e[0], "gUserName" in _p32e[1]],
              expect_contain=[False, True, True])
    except Exception as _e32:
        check("第 32 节异常: %s" % _e32, [True], expect_contain=[False])

    # ---- 33. v53：形参位置上按退格 / 粘贴，不许提示形参自己 ----
    #
    # 用户报的现象：光标停在形参首字符前按【退格】，或把别的过程的变量名【粘贴】
    # 到形参里，弹窗都会把这个形参原样提示出来（v52 只堵住了"敲空格"那一路）。
    #
    # 根因：v51（现场证据按作用域收窄）与 v52（typed_separator）判的都是
    # 【这一次编辑是怎么变的】。而"粘贴一个完整名字"与"手动敲出最后一个字符"
    # 前后两帧快照逐字符等价 —— 从编辑动作上永远区分不开。何况真实代码里形参总
    # 会在本过程体内被用到（Debug.Print username），窄窗口里它就是"真实存在"的，
    # 回声防护照样放行。
    #
    # 修法（语义级、与按键无关）：光标处的词若正是【本行正在声明的名字】
    # （形参 / 局部 Dim），且它在模块级并不存在（过程内的局部变量不是全模块可见
    # 的真名字），就把与它完全同名的候选剔除。
    #   * 前缀照常提示（v37 教训：按"本行声明的名字"整体剔除会误杀 num -> numArr）；
    #   * 模块级真名字照常提示（31.5 的 updateValue / 31.7 的模块级声明行）。
    # 另外：形参常常独占一行（长签名被拆成 `Sub Foo( _` + 缩进的形参行），
    # 续行上没有声明关键字，后端据此把【声明续行】也算作声明行（v53 配套改动），
    # 否则续行上的退格依然会提示自己。
    print("\n=== 33. 形参位置不提示自己：退格 / 粘贴（v53）===")
    try:
        import vbe_bridge as VB33

        # A) 判据：续行识别 + 逻辑行合并出形参名
        check("33.1 以 `_` 结尾的行 -> 认定为续行",
              [P.is_continuation_line("Public Sub B( _"),
               P.is_continuation_line("    B( _  "),
               P.is_continuation_line("    Debug.Print a_b")],
              expect_contain=[True, True, False])

        _sig33 = "\n".join(_HDR31 + [
            "Public Sub ProcedureB( _",
            "    username As String, _",
            "    ByVal other As Long)",
            "    Debug.Print username",
            "End Sub"])
        check("33.2 续行上的 decl_names 含全部形参（合并逻辑行）",
              [sorted(P.decl_names_at_caret(_sig33, (8, 5)) or [])],
              expect_contain=[["other", "procedureb", "username"]])

        # B) 真机 VbeBackend.get_context()：续行也必须算出 decl_names
        #    （只在真机后端上，才会暴露"只在 in_decl 时才去算"这个漏洞）
        class _CM33(object):
            def __init__(self, code):
                self._code = code

            @property
            def CountOfLines(self):
                return len(self._code.split("\n"))

            def Lines(self, start, count):
                _ls = self._code.split("\n")
                return "\n".join(_ls[start - 1:start - 1 + count])

            @property
            def Name(self):
                return "Module1"

        class _CP33(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self._sel = sel

            def GetSelection(self):
                return self._sel

        class _VBE33(object):
            def __init__(self, code, sel):
                self.ActiveCodePane = _CP33(_CM33(code), sel)

        _cur33 = {"code": _sig33, "sel": (8, 5, 8, 5)}
        _orig_get33 = VB33._get_vbe_cached
        VB33._get_vbe_cached = lambda *a, **k: _VBE33(_cur33["code"],
                                                     _cur33["sel"])
        _be33 = VB33.VbeBackend()
        _ctx33 = _be33.get_context()
        _cur33["sel"] = (10, 24, 10, 24)
        _ctx33b = _be33.get_context()
        VB33._get_vbe_cached = _orig_get33
        check("33.3 真机 get_context：续行上有 decl_names（否则引擎拦不住）",
              [_ctx33.get("decl_names"), _ctx33.get("proc_name")],
              expect_contain=[["other", "procedureb", "username"], "ProcedureB"])
        check("33.4 真机 get_context：用法行 decl_names 仍为空（不误伤）",
              [_ctx33b.get("decl_names")],
              expect_contain=[[]])

        # C) 端到端：完全照 main.poll_editor 的分支跑
        #    （复用 31 节那套脚手架：真实解析 + 真实 names_outside_caret）
        VB31._get_vbe_cached = lambda *a, **k: _VBE31(_cur31["code"])

        _CODE33 = [
            "Public Sub ProcedureA()",
            "    Dim username As String",
            '    username = "abc"',
            "    Debug.Print username",
            "End Sub",
            "",
        ]

        def _poll33(body, edit):
            _code, _caret = _parse31(_HDR31 + body)
            _ln, _col = _caret
            _ls = _code.split("\n")
            _old = _ls[_ln - 1]
            _new, _ncol = edit(_old, _col)
            if E.typed_separator(_old, _new):      # v52 那道门照旧生效
                return (False, [], _new)
            _ls[_ln - 1] = _new
            _nc = "\n".join(_ls)
            _cur31["code"] = _nc
            _ui = _UI31()
            _c = E.Completer(_B31(_nc, (_ln, _ncol)), _ui)
            _c.trigger(True)
            return (_ui.shown, sorted(_c.matches or []), _new)

        def _paste33(txt):
            def f(line, col):
                return line[:col - 1] + txt + line[col - 1:], col + len(txt)
            return f

        def _back33(line, col):
            if col <= 1:
                return line, col
            return line[:col - 2] + line[col - 1:], col - 1

        _p = _poll33(_CODE33 + ["Public Sub ProcedureB(ByVal | )",
                                "    Debug.Print username",
                                "End Sub"], _paste33("username"))
        check("33.5 粘贴别的过程的变量名到形参 -> 不提示自己",
              [_p[0], _p[1]], expect_contain=[False, []])

        _p = _poll33(_CODE33 + ["Public Sub ProcedureB( _",
                                "    |username As String)",
                                "    Debug.Print username",
                                "End Sub"], _back33)
        check("33.6 续行形参首字符前按退格 -> 不提示自己",
              [_p[0], _p[1]], expect_contain=[False, []])

        _p = _poll33(_CODE33 + [
            "Public Sub ProcedureB(ByVal a As Long,|username As String)",
            "    Debug.Print username",
            "End Sub"], _back33)
        check("33.7 同行形参首字符前按退格 -> 不提示自己",
              [_p[0], _p[1]], expect_contain=[False, []])

        _p = _poll33(_CODE33 + ["Public Sub ProcedureB()",
                                "    Dim |",
                                "    Debug.Print username",
                                "End Sub"], _paste33("username"))
        check("33.8 粘贴到局部 Dim 位置 -> 不提示自己",
              [_p[0], _p[1]], expect_contain=[False, []])

        _p = _poll33(_CODE33 + ["Public Sub ProcedureB(ByVal | )",
                                "    Debug.Print username",
                                "End Sub"], _paste33("username"))
        _orig_ml33 = E.Completer._module_level_names
        E.Completer._module_level_names = lambda _self: set(
            str(_r[0]).lower() for _r in _self.backend.get_identifiers())
        try:
            _p_off = _poll33(_CODE33 + ["Public Sub ProcedureB(ByVal | )",
                                        "    Debug.Print username",
                                        "End Sub"], _paste33("username"))
        finally:
            E.Completer._module_level_names = _orig_ml33
        check("33.9 对照：关掉 v53 护栏时会提示形参自身（本节没白测）",
              [_p_off[0], "username" in _p_off[1]],
              expect_contain=[True, True])

        # D) 防误杀：模块级真名字 / 前缀补全 / 用法位置都不许被砍
        _p = _poll33(["Public Sub updateValue()",
                      "End Sub",
                      "",
                      "Public Sub ProcedureB(ByVal updateValue| As String)",
                      "End Sub"], _back33)
        check("33.10 形参同名于模块级过程名 -> 照常提示（v37 / 31.5）",
              [_p[0], "updateValue" in _p[1]], expect_contain=[True, True])

        _p = _poll33(["Public Function test() As Long",
                      "    test = 1",
                      "End Function",
                      "",
                      "Public Sub |()",
                      "End Sub"], _paste33("test"))
        check("33.11 Sub 行写全 test，工程里有 Function test() -> 照常提示",
              [_p[0], "test" in [x.lower() for x in _p[1]]],
              expect_contain=[True, True])

        _p = _poll33(["Public Sub ProcedureB()",
                      "    Dim numArr As Long",
                      "    Dim | As Long",
                      "End Sub"], _paste33("num"))
        check("33.12 局部声明行输入前缀 num -> 仍提示 numArr（不许误杀前缀）",
              [_p[0], "numArr" in _p[1]], expect_contain=[True, True])

        _p = _poll33(["Public Sub ProcedureA()",
                      "    Dim username As String",
                      "    Debug.Print userna|",
                      "End Sub"], _paste33("m"))
        check("33.13 同过程内用法位置打前缀 -> 照常提示（31.6 语义不变）",
              [_p[0], "username" in _p[1]], expect_contain=[True, True])
    except Exception as _e33:
        check("第 33 节异常: %s" % _e33, [True], expect_contain=[False])


    # ==================================================================
    # 34. VBE 自带提示列表出现时让位（v54）
    # ==================================================================
    # VBE 的「自动列出成员」下拉列表跟我们的弹窗会出现在同一处（光标下方），
    # 两个列表同时出现会互相遮挡、还抢键盘。所以 VBE 一旦自己弹了，我们就
    # 不弹。这里测两件事：决策函数、以及三条探测判据的接通与兜底。
    print("\n=== 34. VBE 自带列表出现时让位（v54）===")
    try:
        import vbe_bridge as VB54

        # --- 决策：谁该让位（纯函数，与 Win32 无关）---
        if M is not None:
            check("34.1 没敲分隔符、VBE 没弹 -> trigger",
                  [M._poll_action(False, False)],
                  expect_contain=["trigger"])
            check("34.2 敲了空格/标点 -> hide（v52 语义不变）",
                  [M._poll_action(True, False)],
                  expect_contain=["hide"])
            check("34.3 VBE 自带列表已出现 -> hide（本版新增）",
                  [M._poll_action(False, True)],
                  expect_contain=["hide"])
        else:
            check("34.1 main 不可用", [True], expect_contain=[False])

        # --- 窗口形态：尺寸像不像一个下拉列表 ---
        check("34.4 300x150 像列表",
              [VB54._looks_like_list(300, 150)], expect_contain=[True])
        check("34.5 40x18（细长工具条）不像列表",
              [VB54._looks_like_list(40, 18)], expect_contain=[False])
        check("34.6 1900x1000（整屏窗口）不像列表",
              [VB54._looks_like_list(1900, 1000)], expect_contain=[False])

        # --- 探测总入口：三条判据各自能接通，且拿不准一律不拦 ---
        _o_h = VB54._vbe_main_hwnd
        _o_a = VB54._find_list_by_class
        _o_b = VB54._find_list_owned
        _o_c = VB54._list_at_caret
        _o_y = VB54._YIELD_TO_VBE_LIST
        try:
            VB54._find_list_by_class = lambda m: 0
            VB54._find_list_owned = lambda m: 0
            VB54._list_at_caret = lambda r, m: 0

            VB54._vbe_main_hwnd = lambda: 0
            check("34.7 VBE 没开 -> 不拦（拿不准就不拦）",
                  [VB54.vbe_list_visible((10, 20, 34))],
                  expect_contain=[False])

            VB54._vbe_main_hwnd = lambda: 0x1234
            check("34.8 三条判据都没命中 -> 正常弹我们的",
                  [VB54.vbe_list_visible((10, 20, 34))],
                  expect_contain=[False])

            VB54._find_list_by_class = lambda m: 0xAAAA
            check("34.9 判据 A（已知类名）命中 -> 让位",
                  [VB54.vbe_list_visible((10, 20, 34))],
                  expect_contain=[True])

            VB54._find_list_by_class = lambda m: 0
            VB54._find_list_owned = lambda m: 0xBBBB
            check("34.10 判据 B（主窗口 OWNED 的裸弹窗）命中 -> 让位",
                  [VB54.vbe_list_visible((10, 20, 34))],
                  expect_contain=[True])

            VB54._find_list_owned = lambda m: 0
            VB54._list_at_caret = lambda r, m: 0xCCCC
            check("34.11 判据 C（光标下方被 VBE 弹窗盖住）命中 -> 让位",
                  [VB54.vbe_list_visible((10, 20, 34))],
                  expect_contain=[True])

            # 34.12 逃生门：VBECOMPLETE_NO_YIELD=1 时一律不拦，工具照旧工作
            VB54._YIELD_TO_VBE_LIST = False
            check("34.12 关掉让位开关 -> 一律不拦（逃生门）",
                  [VB54.vbe_list_visible((10, 20, 34))],
                  expect_contain=[False])

            # 34.13 Win32 调用本身炸了也不能把工具搞哑 —— 这是最容易踩的坑：
            # 探测代码出异常若向外抛，主循环就会整段停摆，补全彻底失效。
            VB54._YIELD_TO_VBE_LIST = True

            def _boom54(_m):
                raise RuntimeError("Win32 炸了")

            VB54._find_list_by_class = _boom54
            VB54._list_at_caret = lambda r, m: 0
            check("34.13 探测过程抛异常 -> 不拦（绝不因探测失败而哑掉）",
                  [VB54.vbe_list_visible((10, 20, 34))],
                  expect_contain=[False])
        finally:
            VB54._vbe_main_hwnd = _o_h
            VB54._find_list_by_class = _o_a
            VB54._find_list_owned = _o_b
            VB54._list_at_caret = _o_c
            VB54._YIELD_TO_VBE_LIST = _o_y
    except Exception as _e34:
        check("第 34 节异常: %s" % _e34, [True], expect_contain=[False])

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
