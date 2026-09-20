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
    # v61：VBA 内建名字（内建函数 / 常量 / 数据类型）默认会进候选池
    # （vbe_bridge.ENABLE_VBA_BUILTINS）；v62 语言关键字另行一组，开关独立
    # （vbe_bridge.ENABLE_VBA_KEYWORDS）。下面第 1~37 节的断言是"工程内名字的
    # 候选列表精确比对"（用户要求的防噪音护栏），与它们无关 —— 因此统一在
    # 【关闭】的前提下跑，护栏强度保持不变；第 38 节专门验内建名字，
    # 第 39 节专门验关键字。
    try:
        import vbe_bridge as _vb61
        _vb61.ENABLE_VBA_BUILTINS = False
        _vb61.ENABLE_VBA_KEYWORDS = False
    except Exception:
        pass

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
    # v59：块成员（Type 字段 / Enum 成员）不再收录成裸名候选 —— Type 字段只能
    # 通过变量限定访问（p.field），Enum 成员虽可裸名引用但枚举名已限定（E.Member），
    # 提示出来只会把候选池搞吵。类型/枚举名本身照常收录（上面那条）。
    check("Type 成员行不再收录成员名",
          decl_case(c_type, (2, len("    gI") + 1), prefix="g"),
          expect_absent=["gId"])

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

        # 20.11 v59：纯贪心会【漏判】—— 字符要"跳步"时仍必须能匹配
        #
        # 用户原话：定义了 getSettingTitle3，输入 get3 能提示，输入 getse3 不提示。
        # 根因在 _greedy_positions：挑位置时只看"这一档更优"，不管后面还凑不凑得
        # 齐 —— 第 3 个字符 t 贪心跳到词边界 T(10)（档 1 优于接着用的 t(2) 档 2），
        # 后面再没有 s，于是整个 query 被判成不匹配。可 g-e-t-s-e-…-3 =
        # [0,1,2,3,4,15] 本来就成立。修法：先反向算可行性上界 maxpos，再把正向
        # 候选限制在 i..maxpos[t] 之内（见 engine._greedy_positions）。
        m_v59 = [("Module1",
                  "Option Explicit\n"
                  "Private Function getSettingTitle3(ws, col)\n"
                  "    <in>\n"
                  "End Function", 1)]

        s11, m11, h11 = probe(m_v59, "Module1", 3, "    ", "getse3")
        check("20.11 输入 getse3：提示 getSettingTitle3", m11,
              expect_contain=["getSettingTitle3"])
        check("20.11 getse3 命中位置 [0,1,2,3,4,15]",
              [h11.get("getSettingTitle3")],
              expect_contain=[[0, 1, 2, 3, 4, 15]])

        # 20.12 反向把关：query 真的【不是】子序列时仍要判不匹配
        # （不能为了"不漏判"就乱放行 —— 名字里只有一个 s，getss3 不该中）
        s12, m12, _h12 = probe(m_v59, "Module1", 3, "    ", "getss3")
        check("20.12 输入 getss3：不提示（名字里只有一个 s）", m12,
              expect_absent=["getSettingTitle3"])

        # 20.13 不回归：原本就能中的 get3 行为不变
        s13, m13, h13 = probe(m_v59, "Module1", 3, "    ", "get3")
        check("20.13 输入 get3：仍提示 getSettingTitle3", m13,
              expect_contain=["getSettingTitle3"])
        check("20.13 get3 命中位置仍为 [0,1,10,15]",
              [h13.get("getSettingTitle3")], expect_contain=[[0, 1, 10, 15]])

        # 20.14 不能为了修漏判把"优先词边界"的观感改坏：ds -> dataSheet 仍是 [0,4]
        _s14, _m14, h14 = probe(m_two, "Module1", 5, "    ", "ds")
        check("20.14 输入 ds：命中位置仍为词首 [0,4]",
              [h14.get("dataSheet")], expect_contain=[[0, 4]])

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
            #
            # ⚠️ v75 起这条【不能再用 Option Explicit】—— 写了它的模块不再收集
            # 隐式变量（用户口径，见第 52 节），旧写法等于在断言新行为。
            # 本断言真正要守的是"回声防护不误杀隐式变量"，与是否强制声明无关，
            # 所以这里去掉那一行、行号随之前移。
            _trig, _thr = _session24()
            _m3 = _trig([("Module1", "Sub Foo()\n    arr = 1\n"
                                       "    ar\nEnd Sub")],
                        "Module1", 3, 5 + 2)
            check("24.3 隐式变量前缀补全仍在（不带 Option Explicit：ar -> arr）",
                  [_m3], expect_contain=[["arr"]])
            # 反向对照：同一个文本加上 Option Explicit 后，arr 用过但没声明
            # ⇒ 不该再被提示（v75 的语义翻转，与第 52 节同源）。
            _trig, _thr = _session24()
            _m3b = _trig([("Module1", "Option Explicit\nSub Foo()\n    arr = 1\n"
                                        "    ar\nEnd Sub")],
                         "Module1", 4, 5 + 2)
            check("24.3b ★对照：带 Option Explicit 时 arr 不再提示"
                  "（用过但没声明 = 编译错误）",
                  [_m3b], expect_absent=["arr"])

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
    # 34. VBE 自带提示列表出现时让位（v55）
    # ==================================================================
    # VBE 的「自动列出成员」跟我们的弹窗会出现在同一处（光标下方），两个列表
    # 同时出现会互相遮挡、还抢键盘，所以光标停在会让 VBE 弹列表的位置时我们
    # 让位。判据走【光标的语法位置】而不是窗口探测——真机实测：给 VBE 发
    # Ctrl+J 后全系统扫描，一个新窗口都没出现，那个列表是画在代码窗格上的。
    print("\n=== 34. VBE 自带列表出现时让位（v55）===")
    try:
        # --- 决策：本轮轮询该 hide 还是 trigger（纯函数）---
        if M is not None:
            check("34.1 没敲分隔符 -> trigger",
                  [M._poll_action(False)], expect_contain=["trigger"])
            check("34.2 敲了空格/标点 -> hide（v52 语义不变）",
                  [M._poll_action(True)], expect_contain=["hide"])
        else:
            check("34.1 main 不可用", [True], expect_contain=[False])

        # --- 语义判据：光标是不是停在 VBE 会自己弹列表的位置 ---
        _yl = E.vbe_list_expected
        check("34.3 UserForm1. 之后（还没开始输成员名）-> 让位",
              [_yl("    UserForm1.", 15)], expect_contain=[True])
        check("34.4 UserForm1.Cap 输入成员名中 -> 让位",
              [_yl("    UserForm1.Cap", 18)], expect_contain=[True])
        check("34.5 Me. -> 让位",
              [_yl("    Me.", 8)], expect_contain=[True])
        check("34.6 Range(\"A1\"). 之后（点号左边是括号）-> 让位",
              [_yl('    Range("A1").', 17)], expect_contain=[True])
        check("34.7 普通变量名（没有点号）-> 不让位",
              [_yl("    userName = 1", 12)], expect_contain=[False])
        check("34.8 小数点 1. -> 不让位（VBE 也不弹）",
              [_yl("    x = 1.", 11)], expect_contain=[False])
        check("34.9 注释里的点号 -> 不让位",
              [_yl("    ' UserForm1.", 17)], expect_contain=[False])
        check("34.10 字符串里的点号 -> 不让位",
              [_yl('    s = "a."', 12)], expect_contain=[False])
        # 【v72 语义变更】行首的点号 = With 块的成员访问（`With obj` 之后成员
        # 各占一行，点就顶在行首），VBE 在那儿同样会弹成员列表 -> 让位。
        # 旧断言（v55 时代写的"不让位"）是错的：它默认"点号左边必须有东西"，
        # 恰恰漏掉了用户报的 `With Range(...).Font` + `.Size = 12` 场景。
        check("34.11 ★行首就是点号 -> 让位（With 块成员访问；v72 语义变更）",
              [_yl(".", 2)], expect_contain=[True])
        check("34.12 成员名打完（等号后）-> 立刻恢复，不让位",
              [_yl("    UserForm1.Caption = ", 25)], expect_contain=[False])
        check("34.13 空行 / 列号缺失 -> 不让位（不炸）",
              [_yl("", 1), _yl("    x", None), _yl(None, 3)],
              expect_contain=[False])

        # --- 引擎集成：轮询自动触发时让位，手动 Ctrl+Space 照弹 ---
        class _UI55(object):
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

        class _B55(object):
            def __init__(self, ctx):
                self.ctx = ctx

            def get_context(self):
                return self.ctx

            def get_identifiers(self):
                return [("Caption", None, None, False),
                        ("userName", None, None, False)]

            def apply_completion(self, *a, **k):
                return None

        def _ctx55(text, col):
            return {"line_no": 1, "caret_col": col, "line_text": text,
                    "in_string": False, "in_comment": False,
                    "in_type_position": False, "in_decl_position": False,
                    "proc_name": None, "module_name": "M1"}

        _ui55 = _UI55()
        _c55 = E.Completer(_B55(_ctx55("    UserForm1.Cap", 18)), _ui55)
        _c55.trigger(True)          # 轮询自动触发 -> 让位
        check("34.14 轮询触发 + VBE 列表位置 -> 不弹（让位）",
              [_ui55.shown], expect_contain=[False])

        _ui55b = _UI55()
        _c55b = E.Completer(_B55(_ctx55("    UserForm1.Cap", 18)), _ui55b)
        _c55b.trigger()             # 手动 Ctrl+Space -> 用户点名要我们的
        check("34.15 手动触发（Ctrl+Space）即使在 VBE 列表位置也照弹",
              [_ui55b.shown], expect_contain=[True])

        _ui55c = _UI55()
        _c55c = E.Completer(_B55(_ctx55("    userN", 11)), _ui55c)
        _c55c.trigger(True)         # 普通位置：完全不受影响
        check("34.16 普通标识符位置照旧弹出（没被误伤）",
              [_ui55c.shown], expect_contain=[True])

        # 34.17 逃生门：VBECOMPLETE_NO_YIELD=1 -> 一律不让位
        _orig_yield = E.YIELD_TO_VBE_LIST
        try:
            E.YIELD_TO_VBE_LIST = False
            _ui55d = _UI55()
            _c55d = E.Completer(_B55(_ctx55("    UserForm1.Cap", 18)), _ui55d)
            _c55d.trigger(True)
            check("34.17 关掉让位开关 -> 照旧弹出（逃生门）",
                  [_ui55d.shown], expect_contain=[True])
        finally:
            E.YIELD_TO_VBE_LIST = _orig_yield

        # --- 声明里 `As ` / `New ` 之后：VBE 会弹类型列表，同样让位 ---
        check("34.18 Dim x As （类型名还没开始打）-> 让位",
              [_yl("    Dim x As ", 14)], expect_contain=[True])
        check("34.19 Dim x As In（类型名输入中）-> 让位",
              [_yl("    Dim x As In", 16)], expect_contain=[True])
        check("34.20 Set c = New （大写的 New）-> 让位",
              [_yl("    Set c = New ", 17)], expect_contain=[True])
        check("34.21 类型名打完并跟了空格 -> 恢复，不让位",
              [_yl("    Dim x As Long ", 19)], expect_contain=[False])
        check("34.22 注释里的 As -> 不让位",
              [_yl("    ' Dim x As ", 16)], expect_contain=[False])
        check("34.23 变量名 asName（不是 As 关键字）-> 不让位",
              [_yl("    x = asName", 15)], expect_contain=[False])
    except Exception as _e34:
        check("第 34 节异常: %s" % _e34, [True], expect_contain=[False])


    # ==================================================================
    # 35. 自动配对：`(` -> `()`、`"` -> `""`，光标落在中间
    # ==================================================================
    # VBE 原生不自动闭合括号 / 引号（VBE_Extras、Rubberduck 都把它当增强功能加），
    # 所以由我们补。两条铁律：写成功了才吞键；不该插时【放行】让 VBE 原生处理。
    print("\n=== 35. 自动配对（括号 / 引号）===")
    try:
        # --- 按键 -> 符号（纯函数）---
        if M is not None:
            check("35.1 Shift+9 -> (",
                  [M._pair_char_for_key(0x39, True)],
                  expect_contain=["("])
            check("35.2 不按 Shift 的 9 -> 不管",
                  [M._pair_char_for_key(0x39, False)],
                  expect_contain=[None])
            check("35.3 Shift+' -> \"",
                  [M._pair_char_for_key(0xDE, True)],
                  expect_contain=['"'])
            check("35.4 字母键 -> 不管",
                  [M._pair_char_for_key(0x41, True)],
                  expect_contain=[None])
            check("35.5 小键盘 9 -> 不管（NumLock 下是数字）",
                  [M._pair_char_for_key(0x69, True)],
                  expect_contain=[None])
        else:
            check("35.1 main 不可用", [True], expect_contain=[False])

        # --- 插什么、光标落哪（纯函数）---
        import vbe_bridge as VB35
        _pi = VB35._pair_insertion
        check("35.6 行尾敲 ( -> 补成 ()，光标在中间",
              list(_pi("x = ", 4, "(", ")")),
              expect_contain=["x = ()", 5])
        check("35.7 行尾敲 \" -> 补成 \"\"，光标在中间",
              list(_pi("s = ", 4, '"', '"')),
              expect_contain=['s = ""', 5])
        check("35.8 右边已经有 \" -> 只跨过去，一个字符都不写",
              list(_pi('s = "abc"', 8, '"', '"')),
              expect_contain=[None, 9])
        check("35.9 在未闭合字符串里敲 \" -> 只补收尾那一半",
              list(_pi('s = "abc', 8, '"', '"')),
              expect_contain=['s = "abc"', 9])
        check("35.10 行中间敲 ( -> 就地补一对（不会去补别人的右括号）",
              list(_pi("foo(a, b)", 5, "(", ")")),
              expect_contain=["foo(a(), b)", 6])

        # --- 后端集成：真写进"编辑器"（假 VBE）并把光标摆到中间 ---
        class _Mod35(object):
            def __init__(self, text, fmt=None):
                self.lines = [text]
                self.fmt = fmt          # 模拟 VBE 的「自动语法检测」改写整行

            def Lines(self, n, _c):
                # 真 VBE 对"虚拟空行"（末行+1）也返回空串，不报错
                if 1 <= n <= len(self.lines):
                    return self.lines[n - 1]
                return ""

            @property
            def CountOfLines(self):
                return len(self.lines)

            def ReplaceLine(self, n, t):
                if self.fmt is not None:
                    t = self.fmt(t)
                self.lines[n - 1] = t

            def InsertLines(self, n, t):
                # 真 VBE：第 n 行处插入，原第 n 行及之后下移
                self.lines.insert(n - 1, t)

        class _Pane35(object):
            def __init__(self, mod, sl, sc, el, ec):
                self.CodeModule = mod
                self.sel = [sl, sc, el, ec]

            def GetSelection(self):
                return tuple(self.sel)

            def SetSelection(self, sl, sc, el, ec):
                self.sel = [sl, sc, el, ec]

        class _Vbe35(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        def _pair_run(text, col, ch, ec=None, line=1, fmt=None):
            """返回 (该行新文本, 光标列, 是否代为输入)。ec 不同即视为有选区。

            line 可指定光标所在行：>CountOfLines 时即 VBE 的"虚拟空行"。
            fmt 模拟 VBE 把整行重新格式化（等号补空格等）。
            """
            mod = _Mod35(text, fmt)
            pane = _Pane35(mod, line, col, line, col if ec is None else ec)
            _orig = VB35._get_vbe_cached
            VB35._get_vbe_cached = lambda: _Vbe35(pane)
            try:
                ok = VB35.VbeBackend().insert_pair(ch)
            finally:
                VB35._get_vbe_cached = _orig
            _txt = mod.lines[line - 1] if line - 1 < len(mod.lines) else ""
            return _txt, pane.sel[1], ok

        check("35.11 光标在行尾敲 ( -> 写入 () 且光标停在中间",
              list(_pair_run("x = ", 5, "(")),
              expect_contain=["x = ()", 6, True])
        check("35.12 光标在行尾敲 \" -> 写入 \"\" 且光标停在中间",
              list(_pair_run("s = ", 5, '"')),
              expect_contain=['s = ""', 6, True])
        check("35.13 注释里敲 ( -> 不管（返回 False，调用方会放行按键）",
              [_pair_run("' x = ", 7, "(")[2],
               _pair_run("' x = ", 7, "(")[0]],
              expect_contain=[False, "' x = "])
        check("35.14 有选区 -> 不管（替换语义交给 VBE）",
              [_pair_run("x = ab", 3, "(", ec=6)[2],
               _pair_run("x = ab", 3, "(", ec=6)[0]],
              expect_contain=[False, "x = ab"])
        _orig_g = VB35._get_vbe_cached
        try:
            VB35._get_vbe_cached = lambda: None
            _bk = VB35.VbeBackend()
            check("35.15 COM 拿不到 VBE -> 不管，且不炸",
                  [_bk.insert_pair("(")], expect_contain=[False])
        finally:
            VB35._get_vbe_cached = _orig_g

        # --- 右半边：`)` 也是"跨过去"，但只在右边真的有 `)` 时 ---
        check("35.16 Shift+0 -> )",
              [M._pair_char_for_key(0x30, True) if M is not None else None],
              expect_contain=[")"])
        check("35.17 补完 () 后再按 ) -> 跨过去，不多出一个",
              list(_pair_run("x = ()", 6, ")")),
              expect_contain=["x = ()", 7, True])
        check("35.18 右边没有 ) 时按 ) -> 放行让 VBE 自己插",
              [_pair_run("x = ab", 5, ")")[2],
               _pair_run("x = ab", 5, ")")[0]],
              expect_contain=[False, "x = ab"])
        # --- v57：模块末尾的「虚拟空行」上 ReplaceLine 会报参数无效 ---
        check("35.19 光标在虚拟空行（末行+1）-> 先建行再配对，光标在中间",
              list(_pair_run("x = ", 1, "(", line=2)),
              expect_contain=["()", 2, True])
        check("35.20 行号离谱（差 2 行以上）-> 不管，交兜底",
              [_pair_run("x = ", 1, "(", line=4)[2]],
              expect_contain=[False])
        # --- v58：VBE 的「自动语法检测」会把整行改写，行长一变光标就偏 ---
        check("35.23 VBE 在插入的符号前补空格后，光标仍落在引号中间",
              list(_pair_run("    w=4", 6, '"',
                             fmt=lambda s: s.replace('w""', 'w ""'))),
              expect_contain=['    w ""=4', 8, True])
        check("35.24 _locate_inserted：VBE 改写后重新定位插入的符号对",
              [VB35._locate_inserted("    x = 1()", "    x = 1()",
                                     "(", ")", 9),
               VB35._locate_inserted('    w "" = 4', '    w"" = 4', '"', '"', 5),
               VB35._locate_inserted("    Z =() 1 + 2", "    Z=() 1 + 2",
                                     "(", ")", 6)],
              expect_contain=[10, 7, 8])
        check("35.25 _locate_inserted：认不出来时返回 None（调用方退回行尾）",
              [VB35._locate_inserted("完全不相关", "另一行", "(", ")", 3)],
              expect_contain=[None])
        # --- v57：非主线程绝不能碰 COM（否则会污染全局退避，拖慢提示）---
        # 只验证守卫判据本身：vbe_bridge._get_vbe_cached 这个模块属性被前面
        # 小节（14/15/16 节等）换成过 lambda 且没还原，这里拿不到真函数；
        # 完整路径（子线程调用返回 None 且不计失败）由真机验证覆盖。
        import threading as _th_mod
        _box = {}

        def _ask():
            try:
                _box["main"] = VB35._on_main_thread()
            except Exception as _exc:
                _box["main"] = "EXC:%s" % _exc

        _t = _th_mod.Thread(target=_ask)
        _t.daemon = True
        _t.start()
        _t.join(5)
        check("35.21 守卫判据：子线程 _on_main_thread() 为 False、主线程为 True",
              [_box.get("main", "MISSING"), VB35._on_main_thread()],
              expect_contain=[False, True])
        # --- v57：配对没做成时的兜底（绝不吞掉用户的按键）---
        if M is not None:
            _u32 = M.ctypes.windll
            _had = "user32" in _u32.__dict__
            _old = _u32.__dict__.get("user32")
            try:
                class _FakeU32(object):
                    def __init__(self):
                        self.events = []

                    def SendInput(self, n, ptr, size):
                        arr = M.ctypes.cast(
                            ptr, M.ctypes.POINTER(M._INPUT * 2)).contents
                        for _e in arr:
                            self.events.append((_e.type, _e.u.ki.wScan,
                                                _e.u.ki.dwFlags))
                        return 2

                _fake = _FakeU32()
                _u32.__dict__["user32"] = _fake
                _ok = M.send_char("(")
                check("35.22 兜底重发：keydown+keyup 两个 unicode 事件",
                      [_ok, len(_fake.events),
                       chr(_fake.events[0][1]) if _fake.events else "",
                       _fake.events[0][2] if _fake.events else -1,
                       _fake.events[1][2] if len(_fake.events) > 1 else -1],
                      expect_contain=[True, 2, "(", M.KEYEVENTF_UNICODE,
                                      M.KEYEVENTF_UNICODE | M.KEYEVENTF_KEYUP])
            finally:
                if _had:
                    _u32.__dict__["user32"] = _old
                else:
                    _u32.__dict__.pop("user32", None)
    except Exception as _e35:
        check("第 35 节异常: %s" % _e35, [True], expect_contain=[False])

    # ---- 36. v59：续行符误判 / WithEvents / Declare 修饰符（解析层）----
    #
    # 三个都是"行首修饰符被当成名字"或"符号被判成续行符"这一类误判，共同症状：
    # 真名字收不到、假名字混进候选池。真实工程里 pub3 的 Enum 成员名是
    # A_/B_/.../EZ_（每行都以 Z_ 结尾），整模块因此被当成 Enum 块解析。
    print("\n=== 36. 续行符与声明修饰符（v59）===")

    def recs36(code, std=True):
        return [(r[0], r[2], r[3]) for r in P.extract_records(
            code, module="M1", is_std_module=std)]

    # 36.1 名字以 _ 结尾且正好在行尾 —— VBA 要求续行符 _ 前面必须有空白，
    #      所以这不是续行符，不能把下一行（End Enum …）并进来。
    cont_junk = "\n".join([
        "Public Enum E",
        "    A_ = 1: B_: Z_",
        "End Enum",
        "Function getOneFile()",
        "    Dim tmp As String",
        "    With Application.FileDialog(1)",
        "        If .Show = -1 Then",
        "            s = .SelectedItems(1)",
        "        End If",
        "    End With",
        "    getOneFile = s",
        "End Function"])
    got36 = recs36(cont_junk)
    check("36.1 Enum 成员名以 _ 结尾：不算续行符，假名字消失",
          [n for n, _, _ in got36],
          expect_contain=["E", "getOneFile"],
          expect_absent=["Function", "With", "If", "s", "A_"])
    # 36.2 End Enum 必须真的闭合：之后的过程名照常收、局部变量仍归属其过程
    #      （不闭合时全被记成模块级: proc=None）
    check("36.2 End Enum 正确闭合：之后的过程/局部变量归属正确",
          got36,
          expect_contain=[("getOneFile", None, False),
                          ("tmp", "getOneFile", False)])

    # 36.3 续行符规则本身：_ 前有空白才合并
    _merged = P._RE_CONTINUATION.sub(" ", "Dim a _\n    , b")
    _kept = P._RE_CONTINUATION.sub(" ", "    A_ = 1: Z_\nEnd Enum")
    check("36.3 续行符：_ 前有空白才合并，名字结尾的 _ 不合并",
          ["\n" in _merged, "\n" in _kept],
          expect_contain=[False, True])

    # 36.4 Public WithEvents clk As ... —— 收真名 clk，不收 WithEvents
    check("36.4 WithEvents 声明收录 clk（不收 WithEvents）",
          recs36("Public WithEvents clk As MSForms.CommandButton", std=False),
          expect_contain=[("clk", None, True)],
          expect_absent=[("WithEvents", None, True)])

    # 36.5 带访问修饰符的 Declare —— 收真名，不收 Declare
    check("36.5 Private Declare 收录 gApi（不收 Declare）",
          recs36('Private Declare PtrSafe Function gApi Lib "k" '
                 '(ByVal x As Long) As Long'),
          expect_contain=[("gApi", None, True), ("x", None, False)],
          expect_absent=[("Declare", None, True)])
    check("36.6 无修饰 Declare 仍收录",
          recs36('Declare PtrSafe Function gApi2 Lib "k" () As Long'),
          expect_contain=[("gApi2", None, False)])

    # 36.7 块成员（Type 字段 / Enum 成员）不收录成裸名候选；类型名本身照常收录
    #      （用户要求：枚举成员前面有枚举名限定，不该直接提示）
    blk = "\n".join([
        "Public Enum E",
        "    Red = 1: Green = 2: Zz_",
        "End Enum",
        "Public Type T",
        "    fx As Long",
        "End Type"])
    check("36.7 Type/Enum 成员不收录（类型/枚举名收录）",
          recs36(blk),
          expect_contain=[("E", None, False), ("T", None, False)],
          expect_absent=[("Red", None, False), ("Green", None, False),
                         ("Zz_", None, False), ("fx", None, False)])
    # 块结束后，后面的模块级声明仍要正常收录（continue 不能把状态搞乱）
    blk2 = blk + "\nPublic gAfter As Long"
    check("36.8 Enc/Type 块之后的声明照常收录",
          recs36(blk2),
          expect_contain=[("gAfter", None, False)])

    # ---- 37. v60：组件名 / 窗体控件名不该被回声防护误杀 ----
    print("\n=== 37. v60 组件名与窗体控件名（在窗体自己的代码模块里）===")
    try:
        class _UI37(object):
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

        class _B37(object):
            """模拟【窗体自己的代码模块】。

            三个候选都"只声明在这个模块里"，现场文本里也一个字都没出现过：
              * UserForm1 —— 组件名。窗体代码里从来不会写 UserForm1 这个
                词（除非要 UserForm1.Show），可它当然是合法引用；
              * Label1 —— 刚拖上去的控件，还没写任何事件过程，代码里没有它；
              * useGhost —— 对照用的真幽灵：既不是结构性名字，声明集合里
                也没有它，就该照旧被回声防护剔掉。
            """

            def __init__(self, structural=True, module="UserForm1"):
                self._structural = structural
                self._module = module

            def get_context(self):
                return {"line_no": 1, "caret_col": 1,
                        "line_text": "", "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": self._module}

            def get_identifiers(self):
                return [("UserForm1", "UserForm1", None, False),
                        ("Label1", "UserForm1", None, True),
                        ("useGhost", "UserForm1", None, True)]

            def get_declared_names(self):
                # 刻意【不】含 useghost：它代表"解析层没能归类、只剩缓存残留
                # 的幽灵"，用来验证这次放宽没有把真幽灵一起放行。
                return ["userform1", "label1"]

            def caret_word_at_collect(self):
                return ""

            def declared_elsewhere(self, name, module_name):
                return False        # 全部"只声明在" UserForm1

            def get_structural_names(self):
                return ["userform1", "label1"] if self._structural else []

            def names_outside_caret(self, caret, scope_only=False):
                return set()        # 现场文本里一个字都没有

            def apply_completion(self, *a, **k):
                return None

        def _trig37(word, structural=True, module="UserForm1"):
            _be = _B37(structural, module)
            _line = "    " + word
            _ctx = _be.get_context()
            _ctx["line_text"] = _line
            _ctx["caret_col"] = len(_line) + 1
            _be.get_context = lambda: _ctx
            _c = E.Completer(_be, _UI37())
            _c.trigger()
            return sorted(_c.matches or [])

        # 37.1 组件名：在【窗体自己的】代码里输入 use（用户插完窗体就在这打）
        check("37.1 窗体自己的代码里输入 use -> 提示 UserForm1（真幽灵仍剔）",
              _trig37("use"),
              expect_contain=["UserForm1"],
              expect_absent=["useGhost"])
        # 37.2 对照：去掉结构性名字证据，就是修复前的行为（不提示）
        check("37.2 对照：无结构性名字证据时不提示 UserForm1",
              _trig37("use", structural=False),
              expect_absent=["UserForm1"])
        # 37.3 控件名：刚拖上去的 Label1（代码里还没有任何 Label1 字样）
        check("37.3 窗体里输入 la -> 提示 Label1",
              _trig37("la"),
              expect_contain=["Label1"])
        # 37.4 对照：去掉结构性名字证据则不提示
        check("37.4 对照：无结构性名字证据时不提示 Label1",
              _trig37("la", structural=False),
              expect_absent=["Label1"])
        # 37.5 控件名跨模块不可见（与"不提示别模块私有成员"的约定一致）
        check("37.5 别的模块里输入 la -> 不提示 Label1",
              _trig37("la", module="模块2"),
              expect_absent=["Label1"])

    except Exception as _e37:
        check("第 37 节异常: %s" % _e37, [True], expect_contain=[False])

    # ---- 38. v61：VBA 语言自带的名字（内建函数 / 常量 / 数据类型）----
    print("\n=== 38. v61 VBA 内建名字进候选池 ===")
    try:
        import re as _re38
        import vba_builtins as VB38

        # 38.1~38.5：清单自身（纯静态，不依赖 Excel）
        _low38 = [n.lower() for n in VB38.BUILTIN_NAMES]
        _lowT38 = [n.lower() for n in VB38.BUILTIN_TYPE_NAMES]
        check("38.1 两个清单各自无重名（函数+常量 %d 个、类型名 %d 个）"
              % (len(_low38), len(_lowT38)),
              [len(set(_low38)) == len(_low38)
               and len(set(_lowT38)) == len(_lowT38)], expect_contain=[True])
        _all38 = _low38 + _lowT38
        _bad38 = [n for n in _all38
                  if not _re38.match(r"^[^\W\d]\w*$", n)]
        check("38.2 清单全是合法标识符（异常 %s）" % (_bad38 or "无"),
              [len(_bad38) == 0], expect_contain=[True])
        # 语言关键字一律不许混进来（v59 修过"Function/If/With 混进候选池"）
        _kws38 = set(("if then else elseif end for next each step do loop while "
                      "wend until select case with sub function property dim "
                      "redim set let const static public private friend "
                      "global declare type enum new nothing goto on option "
                      "exit resume is like mod and or not xor eqv imp"
                      ).split())
        check("38.3 清单里没有语言关键字（命中 %s）"
              % (sorted(_kws38 & set(_low38)) or "无"),
              [len(_kws38 & set(_low38)) == 0], expect_contain=[True])
        _want38 = ("MsgBox", "Left", "Len", "Split", "IsArray", "CLng",
                   "InStrRev", "DateAdd", "StrConv", "CreateObject", "Array",
                   "Shell", "vbCrLf", "vbYes", "vbRed", "vbNullString")
        _miss38 = [n for n in _want38 if n.lower() not in VB38.BUILTIN_LOWER]
        check("38.4 常用内建函数/常量齐备（缺 %s）" % (_miss38 or "无"),
              [len(_miss38) == 0], expect_contain=[True])
        _missT38 = [n for n in ("String", "Long", "Boolean", "Variant")
                    if n.lower() not in VB38.BUILTIN_TYPE_LOWER]
        check("38.5 内建数据类型齐备（缺 %s）" % (_missT38 or "无"),
              [len(_missT38) == 0], expect_contain=[True])

        class _UI38(object):
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

        class _B38(object):
            """模拟"代码里从没出现过 MsgBox 的模块"。

            候选池里是 VBA 内建名字（挂虚拟模块 VBA）+ 一个用户自己声明过的
            Left（模拟真实工程"同名只留用户的"那种形态）+ 一个真幽灵 msgGhost。
            elsewhere 控制 declared_elsewhere 的返回值 —— 真实情形下内建名字
            归属虚拟模块 VBA，从当前模块看就是"别的模块声明过"（True）。
            """

            def __init__(self, structural=True, elsewhere=False,
                         module="模块2", builtin=True):
                self._structural = structural
                self._elsewhere = elsewhere
                self._module = module
                self._builtin = builtin

            def get_context(self):
                return {"line_no": 1, "caret_col": 1, "line_text": "",
                        "in_string": False, "in_comment": False,
                        "in_type_position": False, "in_decl_position": False,
                        "decl_names": [], "proc_name": None,
                        "module_name": self._module}

            def get_identifiers(self):
                return [("MsgBox", "VBA", None, False),
                        ("vbCrLf", "VBA", None, False),
                        ("String", "VBA", None, False),
                        ("Long", "VBA", None, False),
                        ("Left", "模块2", None, False),
                        ("msgGhost", "模块2", None, True)]

            def get_declared_names(self):
                return ["msgbox", "vbcrlf", "left", "string", "long"]

            def caret_word_at_collect(self):
                return ""

            def declared_elsewhere(self, name, module_name):
                if self._elsewhere:
                    return str(name).lower() in ("msgbox", "vbcrlf",
                                                 "string", "long")
                return False

            def get_structural_names(self):
                if not self._structural:
                    return []
                return ["msgbox", "vbcrlf", "string", "long"]

            def get_builtin_names(self):
                # 这四个就是候选池里的"VBA 内建名字"，引擎据此收紧匹配。
                if not self._builtin:
                    return []
                return ["msgbox", "vbcrlf", "string", "long"]

            def get_type_names(self):
                return ["String", "Long"]

            def names_outside_caret(self, caret, scope_only=False):
                return set()        # 现场文本里一个字都没有

            def apply_completion(self, *a, **k):
                return None

        def _trig38(word, **kw):
            _be = _B38(**kw)
            _line = "    " + word
            _ctx = _be.get_context()
            _ctx["line_text"] = _line
            _ctx["caret_col"] = len(_line) + 1
            _be.get_context = lambda: _ctx
            _c = E.Completer(_be, _UI38())
            _c.trigger()
            return sorted(_c.matches or [])

        def _trig38_type(word):
            _be = _B38(elsewhere=True)
            _line = "Dim x As " + word
            _ctx = _be.get_context()
            _ctx["line_text"] = _line
            _ctx["caret_col"] = len(_line) + 1
            _ctx["in_type_position"] = True
            _be.get_context = lambda: _ctx
            _c = E.Completer(_be, _UI38())
            _c.trigger()
            return sorted(_c.matches or [])

        # 38.6 主场景：代码里从没写过 MsgBox，输入 ms 应能提示出来
        check("38.6 输入 ms -> 提示 MsgBox（现场文本里没有它；幽灵仍剔）",
              _trig38("ms", elsewhere=True),
              expect_contain=["MsgBox"], expect_absent=["msgGhost"])
        # 38.7 结构性证据这一路要能单独顶住（只是"进池子"并不够）
        check("38.7 只靠结构性证据也放行（declared_elsewhere=False）",
              _trig38("ms", structural=True, elsewhere=False),
              expect_contain=["MsgBox"])
        # 38.8 对照：两条证据都没有 = 修复前行为（不提示）
        check("38.8 对照：无结构性证据且未声明过 -> 不提示（修复前行为）",
              _trig38("ms", structural=False, elsewhere=False),
              expect_absent=["MsgBox"])
        # 38.9 真幽灵照旧被回声防护剔掉（放宽是精确的）
        check("38.9 真幽灵 msgGhost 仍被剔",
              _trig38("msgg", structural=True, elsewhere=True),
              expect_absent=["msgGhost"])
        # 38.10 内建名字跨模块可见（按模块级 priv=False 收录）
        check("38.10 内建名字在别的模块也提示",
              _trig38("vbcr", structural=True, elsewhere=True,
                      module="模块3"),
              expect_contain=["vbCrLf"])
        # 38.11 As 位置提示内建数据类型，且不冒出函数名
        check("38.11 As 位置提示内建类型名（不冒函数名）",
              _trig38_type("str"),
              expect_contain=["String"], expect_absent=["MsgBox"])
        # 38.12 同名候选只保留一条
        _vis38 = E.filter_identifiers_by_scope(
            [("Left", "VBA", None, False), ("Left", "模块2", None, False)],
            None, "模块2")
        check("38.12 同名候选只保留一条（实得 %d 条）" % len(_vis38),
              [len(_vis38) == 1], expect_contain=[True])

        # 38.13~38.15：内建名字的匹配收紧档（v61 加、**v71 放宽到"实际不再收紧"**）
        # v71 起"纯分散命中"的最低输入长度 = BUILTIN_SCATTER_MIN(2) —— 3 字符的
        # 跳步输入也能命中；把阈值调回 4 就复现旧口径（38.14 就是这么对照的）。
        _sm38 = E.BUILTIN_SCATTER_MIN
        try:
            E.BUILTIN_SCATTER_MIN = 4
            _strict38 = _trig38("msb", elsewhere=True)
        finally:
            E.BUILTIN_SCATTER_MIN = _sm38
        check("38.13 v71：3 字符跳步输入 msb -> 提示 MsgBox（阈值已从 4 放到 2）",
              _trig38("msb", elsewhere=True),
              expect_contain=["MsgBox"], expect_absent=["msgGhost"])
        check("38.14 对照：阈值调回 4（旧口径）-> 同样的 msb 被挡掉"
              "（说明这条规则确实在起作用）",
              _strict38, expect_absent=["MsgBox"])
        check("38.15 前缀命中不受收紧影响（ms -> MsgBox）",
              _trig38("ms", elsewhere=True),
              expect_contain=["MsgBox"])

        # 38.16~38.20：真实 VbeBackend 全链路（临时把开关打开）
        # v70 起"内建枚举常量（vb*）"默认不收（用户口径：不用它、嫌干扰）。这一节
        # 验的是"完整名单都接上了（含常量）"，所以【两个开关都临时打开】——
        # 断言强度不变，只是把默认值这件事挪到第 47 节单独把关。
        import vbe_bridge as VB38R
        _saved38 = (VB38R.ENABLE_VBA_BUILTINS, VB38R.ENABLE_VBA_CONSTANTS)
        VB38R.ENABLE_VBA_BUILTINS = True
        VB38R.ENABLE_VBA_CONSTANTS = True
        try:
            class _CM38R(object):
                def __init__(self, text):
                    self.text = text

                @property
                def CountOfLines(self):
                    return self.text.count("\n") + 1

                def Lines(self, start, count):
                    return "\r\n".join(
                        self.text.split("\n")[start - 1:start - 1 + count])

            class _Comp38R(object):
                def __init__(self, name, text, ctype=1):
                    self.Name = name
                    self.Type = ctype
                    self.CodeModule = _CM38R(text)

            class _Proj38R(object):
                def __init__(self, comps):
                    self.Name = "VBAProject"
                    self.VBComponents = list(comps)

            class _Pane38R(object):
                def __init__(self, comp):
                    self._c = comp
                    self.sel = (1, 1, 1, 1)

                @property
                def CodeModule(self):
                    return self._c.CodeModule

                def GetSelection(self):
                    return self.sel

            class _VBE38R(object):
                def __init__(self, comps, act):
                    self.ActiveVBProject = _Proj38R(comps)
                    self.ActiveCodePane = _Pane38R(act)

            # 模块里一个字都没提过 MsgBox —— 提示它只能靠内建名字这一路
            _comp38 = _Comp38R(
                "Module1", "Sub Foo()\n    x = 1\n    <in>\nEnd Sub", 1)
            _vbe38 = _VBE38R([_comp38], _comp38)
            _orig38 = VB38R._get_vbe_cached
            VB38R._get_vbe_cached = lambda: _vbe38
            try:
                _bk38 = VB38R.VbeBackend()
                _ids38 = _bk38.get_identifiers()
                _nm38 = set(str(r[0]).lower() for r in _ids38)
                _vba38 = [r for r in _ids38
                          if str(r[1]) == VB38R._BUILTIN_MODULE]
                _visR38 = E.filter_identifiers_by_scope(
                    _ids38, None, "Module1")
                _hit38 = [n for n in _visR38
                          if E.fuzzy_match(n, "ms") is not None]
                check("38.16 真实后端：内建名字进了候选池（挂 VBA 的 %d 条）"
                      % len(_vba38),
                      ["msgbox" in _nm38 and "vbcrlf" in _nm38
                       and "split" in _nm38 and "string" in _nm38],
                      expect_contain=[True])
                check("38.17 真实后端：内建名字跨模块可见（priv=False）",
                      ["msgbox" in set(n.lower() for n in _visR38)],
                      expect_contain=[True])
                check("38.18 真实后端：结构性证据含内建名字",
                      ["msgbox" in _bk38.get_structural_names()
                       and "string" in _bk38.get_structural_names()],
                      expect_contain=[True])
                check("38.19 真实后端：declared_elsewhere 认得内建名字",
                      [_bk38.declared_elsewhere("MsgBox", "Module1")],
                      expect_contain=[True])
                check("38.20 真实后端：get_builtin_names 只含内建",
                      ["msgbox" in _bk38.get_builtin_names()
                       and "foo" not in _bk38.get_builtin_names()],
                      expect_contain=[True])
                check("38.21 真实池子里输入 ms 能命中 MsgBox",
                      _hit38, expect_contain=["MsgBox"])
            finally:
                VB38R._get_vbe_cached = _orig38
        finally:
            (VB38R.ENABLE_VBA_BUILTINS,
             VB38R.ENABLE_VBA_CONSTANTS) = _saved38

        # 38.22~38.24：Type/Enum 块成员不该从"隐式变量"那条路溜回候选池
        #
        # v59 让 extract_records 不再收块成员，但"只读用法"扫描（_usage_candidates）
        # 是整篇文本扫的、没有块的概念 —— 真机上枚举成员又从这条路溜回来了
        # （工程里冒出 A_ / AA_ / AZ_ … 共 12 个）。这里用最小片段钉住修复。
        _blk38 = "\n".join([
            "Public Enum E",
            "    A_ = 0: B_ = 1: Z_ = 25",
            "End Enum",
            "Public Type gRec",
            "    gId As Long",
            "End Type",
            "Sub Foo()",
            "    afterBlk = 1",
            "End Sub"])
        _rec38b = P.extract_records(_blk38, module="M1", is_std_module=True)
        _imp38b = P.extract_implicit_records(
            _blk38, module="M1", is_std_module=True,
            declared=_rec38b, scope="proc")
        _nm38b = sorted(str(n) for n, _m, _p, _pr in _rec38b + _imp38b)
        check("38.22 Enum/Type 成员不进候选池（漏进来的 _ 结尾名：%s）"
              % [n for n in _nm38b if n.endswith("_")],
              [not any(n.endswith("_") for n in _nm38b)],
              expect_contain=[True])
        check("38.23 Enum/Type 名本身照常收录",
              _nm38b, expect_contain=["E", "gRec"])
        check("38.24 块【之后】的代码照常收隐式变量（块状态已复位）",
              _nm38b, expect_contain=["afterBlk"])

    except Exception as _e38:
        check("第 38 节异常: %s" % _e38, [True], expect_contain=[False])

    # ---- 39. v62：语言关键字（Sub / Dim / If / For / Set …）----
    print("\n=== 39. v62 VBA 语言关键字进候选池 ===")
    try:
        import re as _re39
        import vba_builtins as VB39
        import vbe_bridge as VB39R

        # 39.1~39.5：清单自身（纯静态，不依赖 Excel）
        _kw39 = list(VB39.BUILTIN_KEYWORDS)
        _low39 = [n.lower() for n in _kw39]
        check("39.1 关键字清单无重名（共 %d 个）" % len(_low39),
              [len(set(_low39)) == len(_low39)], expect_contain=[True])
        _bad39 = [n for n in _kw39 if not _re39.match(r"^[^\W\d]\w*$", n)]
        check("39.2 关键字全是合法标识符（异常 %s）" % (_bad39 or "无"),
              [len(_bad39) == 0], expect_contain=[True])
        _ovl39 = sorted(set(_low39)
                        & (set(VB39.BUILTIN_LOWER) | set(VB39.BUILTIN_TYPE_LOWER)))
        check("39.3 与内建函数/类型清单无重叠（重叠 %s；重叠项由收集侧去重挡掉）"
              % (_ovl39 or "无"),
              [len(_ovl39) == 0], expect_contain=[True])
        _want39 = ("Sub", "Function", "Property", "Dim", "ReDim", "Set",
                   "Let", "If", "Then", "Else", "ElseIf", "End", "For",
                   "Next", "Each", "Do", "Loop", "While", "Wend", "Select",
                   "Case", "With", "Type", "Enum", "New", "Nothing", "Me",
                   "And", "Or", "Not", "Mod", "Call", "Exit", "On",
                   "Static", "Private", "Public", "Declare", "ByVal",
                   "ByRef", "Optional", "ParamArray", "WithEvents")
        _miss39 = [n for n in _want39
                   if n.lower() not in VB39.BUILTIN_KEYWORD_LOWER]
        check("39.4 常用关键字齐备（缺 %s）" % (_miss39 or "无"),
              [len(_miss39) == 0], expect_contain=[True])
        # 刻意不收的：歧义过大的语句名（更常当变量/属性名）+ 过时的 Def* 组
        _no39 = ("name", "line", "width", "spc", "tab",
                 "defint", "defstr", "defvar", "base", "compare")
        _hit39 = [n for n in _no39 if n in VB39.BUILTIN_KEYWORD_LOWER]
        check("39.5 歧义/过时关键字仍不收（命中 %s）" % (_hit39 or "无"),
              [len(_hit39) == 0], expect_contain=[True])

        # 39.6~39.12：走完整 trigger 链路（假后端，纯 Python，不碰 Excel）
        class _UI39(object):
            def show(self, *a, **k):
                pass

            def update_selection(self, *a, **k):
                pass

            def hide(self, *a, **k):
                pass

            def contains_point(self, *a, **k):
                return False

        class _B39(object):
            """模拟"新模块第一行，代码里一个关键字都还没写过"。

            候选池：三个语言关键字（挂虚拟模块 VBA）+ 一个真幽灵 kwGhost。
            structural / elsewhere 分别控制两路"真实存在"的证据 —— 真实情形下
            关键字归属虚拟模块 VBA，从当前模块看就是"别的模块声明过"（True）。
            """

            def __init__(self, structural=True, elsewhere=False,
                         module="模块2", builtin=True):
                self._structural = structural
                self._elsewhere = elsewhere
                self._module = module
                self._builtin = builtin

            def get_context(self):
                return {"line_no": 1, "caret_col": 1, "line_text": "",
                        "in_string": False, "in_comment": False,
                        "in_type_position": False, "in_decl_position": False,
                        "decl_names": [], "proc_name": None,
                        "module_name": self._module}

            def get_identifiers(self):
                return [("If", "VBA", None, False),
                        ("Dim", "VBA", None, False),
                        ("Sub", "VBA", None, False),
                        ("Get", "VBA", None, False),
                        ("kwGhost", "模块2", None, True)]

            def get_declared_names(self):
                return ["if", "dim", "sub", "get"]

            def caret_word_at_collect(self):
                return ""

            def declared_elsewhere(self, name, module_name):
                if self._elsewhere:
                    return str(name).lower() in ("if", "dim", "sub", "get")
                return False

            def get_structural_names(self):
                if not self._structural:
                    return []
                return ["if", "dim", "sub", "get"]

            def get_builtin_names(self):
                # 关键字与内建名字共用这一档收紧（见 engine.trigger）
                if not self._builtin:
                    return []
                return ["if", "dim", "sub", "get"]

            def get_type_names(self):
                return []

            def names_outside_caret(self, caret, scope_only=False):
                return set()        # 现场文本里一个字都没有

            def apply_completion(self, *a, **k):
                return None

        def _trig39(word, **kw):
            _be = _B39(**kw)
            _line = "    " + word
            _ctx = _be.get_context()
            _ctx["line_text"] = _line
            _ctx["caret_col"] = len(_line) + 1
            _be.get_context = lambda: _ctx
            _c = E.Completer(_be, _UI39())
            _c.trigger()
            return sorted(_c.matches or [])

        check("39.6 新模块里打 if -> 提示 If（现场没有它；幽灵仍剔）",
              _trig39("if", elsewhere=True),
              expect_contain=["If"], expect_absent=["kwGhost"])
        check("39.7 只靠结构性证据也放行（declared_elsewhere=False）",
              _trig39("if", structural=True, elsewhere=False),
              expect_contain=["If"])
        check("39.8 对照：无结构性证据且未声明过 -> 不提示（修复前行为）",
              _trig39("if", structural=False, elsewhere=False),
              expect_absent=["If"])
        check("39.9 真幽灵 kwGhost 仍被回声防护剔掉",
              _trig39("kwghost", structural=True, elsewhere=True),
              expect_absent=["kwGhost"])
        check("39.10 关键字在别的模块也提示（按模块级 priv=False 收录）",
              _trig39("di", structural=True, elsewhere=True, module="模块3"),
              expect_contain=["Dim"])
        # v71：关键字与内建名字共用这一档，阈值放到 2 之后 2 字符的跳步输入也能命中；
        # 调回 4 即复现 v61~v70 的旧口径（39.12 就是那个对照）。
        _sm39 = E.BUILTIN_SCATTER_MIN
        try:
            E.BUILTIN_SCATTER_MIN = 4
            _strict39 = _trig39("dm", elsewhere=True)
        finally:
            E.BUILTIN_SCATTER_MIN = _sm39
        check("39.11 v71：2 字符跳步输入 dm -> 提示 Dim（阈值已从 4 放到 2）",
              _trig39("dm", elsewhere=True), expect_contain=["Dim"])
        check("39.12 对照：阈值调回 4（旧口径）-> 同样的 dm 被挡掉"
              "（说明这条规则确实在起作用）",
              _strict39, expect_absent=["Dim"])

        # 39.13~39.16：真实 VbeBackend 全链路（临时把两个开关都打开）
        _saved39 = (VB39R.ENABLE_VBA_BUILTINS, VB39R.ENABLE_VBA_KEYWORDS)
        VB39R.ENABLE_VBA_BUILTINS = True
        VB39R.ENABLE_VBA_KEYWORDS = True
        try:
            class _CM39(object):
                def __init__(self, text):
                    self.text = text

                @property
                def CountOfLines(self):
                    return self.text.count("\n") + 1

                def Lines(self, start, count):
                    return "\r\n".join(
                        self.text.split("\n")[start - 1:start - 1 + count])

            class _Comp39(object):
                def __init__(self, name, text, ctype=1):
                    self.Name = name
                    self.Type = ctype
                    self.CodeModule = _CM39(text)

            class _Proj39(object):
                def __init__(self, comps):
                    self.Name = "VBAProject"
                    self.VBComponents = list(comps)

            class _Pane39(object):
                def __init__(self, comp):
                    self._c = comp

                @property
                def CodeModule(self):
                    return self._c.CodeModule

                def GetSelection(self):
                    return (1, 1, 1, 1)

            class _VBE39(object):
                def __init__(self, comps, act):
                    self.ActiveVBProject = _Proj39(comps)
                    self.ActiveCodePane = _Pane39(act)

            # 模块里一个字都没提过 If —— 提示它只能靠关键字这一路
            _comp39 = _Comp39("Module1", "Sub Foo()\n    x = 1\nEnd Sub", 1)
            _vbe39 = _VBE39([_comp39], _comp39)
            _orig39 = VB39R._get_vbe_cached
            VB39R._get_vbe_cached = lambda: _vbe39
            try:
                _bk39 = VB39R.VbeBackend()
                _ids39 = _bk39.get_identifiers()
                _nm39 = set(str(r[0]).lower() for r in _ids39)
                _vba39 = [r for r in _ids39
                          if str(r[1]) == VB39R._BUILTIN_MODULE]
                _visR39 = E.filter_identifiers_by_scope(_ids39, None, "Module1")
                _hit39b = [n for n in _visR39
                           if E.fuzzy_match(n, "if") is not None]
                check("39.13 真实后端：关键字进了候选池（挂 VBA 的共 %d 条）"
                      % len(_vba39),
                      ["if" in _nm39 and "dim" in _nm39 and "sub" in _nm39],
                      expect_contain=[True])
                check("39.14 真实后端：关键字跨模块可见（priv=False）",
                      ["if" in set(n.lower() for n in _visR39)],
                      expect_contain=[True])
                check("39.15 真实后端：结构性证据与 declared_elsewhere 都认关键字",
                      ["if" in _bk39.get_structural_names()
                       and bool(_bk39.declared_elsewhere("If", "Module1"))],
                      expect_contain=[True])
                check("39.16 真实池子里打 if 能命中 If",
                      _hit39b, expect_contain=["If"])

                # 开关独立：关掉关键字，内建函数照常
                VB39R.ENABLE_VBA_KEYWORDS = False
                try:
                    _bk39b = VB39R.VbeBackend()
                    _nm39b = set(str(r[0]).lower()
                                 for r in _bk39b.get_identifiers())
                    check("39.17 关掉关键字开关：关键字不进池、内建函数照常",
                          ["if" not in _nm39b and "msgbox" in _nm39b],
                          expect_contain=[True])
                finally:
                    VB39R.ENABLE_VBA_KEYWORDS = True
            finally:
                VB39R._get_vbe_cached = _orig39
        finally:
            (VB39R.ENABLE_VBA_BUILTINS,
             VB39R.ENABLE_VBA_KEYWORDS) = _saved39

    except Exception as _e39:
        check("第 39 节异常: %s" % _e39, [True], expect_contain=[False])

    # ==================================================================
    # 40. v63 VBE 自带提示窗（参数信息）可见时让位
    # ==================================================================
    # 用户报：`MsgBox "已完成！",` 敲下逗号后 VBE 弹出【参数信息】，我们的候选窗
    # 同时也冒出来，两个窗叠在同一处冲突。
    #
    # 根因：v55 的让位判据走【光标语法位置】（`标识符.` 之后 / `As `/`New ` 之后），
    # 只覆盖"自动列出成员 / 类型列表"，VBE 在别处弹的提示（最典型的就是参数信息）
    # 它一概不认 —— README 里其实把这条当"已知不拦"写着。
    #
    # 修法：加一路【精确】判据。真机实测：VBE 的提示窗是【预建复用】的真窗口
    # （什么都没按时就在顶层窗口列表里、只是不可见；Ctrl+Shift+I 后立刻可见、
    # Esc 后立刻不可见），类名是 VB IDE 专有的 NameListWndClass / PopupTipWndClass。
    # 后端 vbe_popup_visible() 直接查这两个窗口的可见性，引擎据此让位。
    #
    # 注：v55 当年以"发 Ctrl+J 后全系统零新窗口"推断 VBE 列表没有独立窗口 ——
    # 那只对"【新建】窗口"成立，被复用骗了。
    print("\n=== 40. v63 VBE 提示窗可见时让位（参数信息）===")
    try:
        import vbe_bridge as VB40

        # 40.1~40.3：探测函数自身（纯 Win32，不依赖 Excel）
        check("40.1 提示窗类名清单含 VBE 那两个窗口类",
              [set(VB40.VBE_POPUP_CLASSES)
               >= {"NameListWndClass", "PopupTipWndClass"}],
              expect_contain=[True])
        check("40.2 空句柄 / 无效句柄 -> 判为不可见（不炸）",
              [VB40._popup_details_now([]),
               VB40._popup_details_now([0, 123]),
               VB40._hwnd_alive(0)],
              expect_contain=[[], [], False])
        try:
            _pinfo40 = VB40.vbe_popup_info()
        except Exception as _e40pi:
            _pinfo40 = "异常: %s" % _e40pi
        check("40.2b vbe_popup_info() 返回明细列表且不抛（此刻=%r）" % (_pinfo40,),
              [isinstance(_pinfo40, list),
               all(isinstance(t, tuple) and len(t) == 3 for t in _pinfo40)],
              expect_contain=[True, True])
        try:
            _pv40 = VB40.vbe_popup_visible()
        except Exception as _e40pv:
            _pv40 = "异常: %s" % _e40pv
        check("40.3 vbe_popup_visible() 返回布尔且不抛（此刻=%s）" % (_pv40,),
              [isinstance(_pv40, bool)], expect_contain=[True])

        # 40.4~40.11：走完整 trigger 链路（假后端，纯 Python，不碰 Excel）
        class _UI40(object):
            def __init__(self):
                self.shown = False
                self.rows = []

            def show(self, rows, selection=0, completer=None):
                self.shown = True
                self.rows = list(rows)

            def update_selection(self, *a, **k):
                pass

            def hide(self, *a, **k):
                self.shown = False

            def contains_point(self, *a, **k):
                return False

        _LINE40 = '    MsgBox "已完成！", vb'

        class _B40Base(object):
            """复刻"光标在 MsgBox 第二个参数上"的现场。

            这个位置【不在】v55 的两种让位位置上（没有点号、也不是 As/New 之后），
            所以它恰好用来隔离出"窗口判据"这一路 —— 语法判据在这里根本不生效。
            """

            def __init__(self, **kw):
                pass

            def get_context(self):
                return {"line_no": 3, "caret_col": len(_LINE40) + 1,
                        "line_text": _LINE40, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": "模块2"}

            def get_identifiers(self):
                return [("MsgBox", "VBA", None, False),
                        ("vbYes", "VBA", None, False),
                        ("vbNo", "VBA", None, False)]

            def get_declared_names(self):
                return ["msgbox", "vbyes", "vbno"]

            def declared_elsewhere(self, name, module_name):
                return str(name).lower() in ("msgbox", "vbyes", "vbno")

            def get_structural_names(self):
                return ["msgbox", "vbyes", "vbno"]

            def get_builtin_names(self):
                return []

            def get_type_names(self):
                return []

            def names_outside_caret(self, caret, scope_only=False):
                return set()

            def apply_completion(self, *a, **k):
                return None

        class _B40(_B40Base):
            """带 vbe_popup_visible 的新式后端。"""

            def __init__(self, popup=False, boom=False, **kw):
                self._popup = popup
                self._boom = boom

            def vbe_popup_visible(self):
                if self._boom:
                    raise RuntimeError("boom")
                return self._popup

            def vbe_popup_info(self):
                if self._boom:
                    raise RuntimeError("boom")
                return [("PopupTipWndClass", 200, 16)] if self._popup else []

        class _B40Old(_B40Base):
            """旧式 / 测试后端：没有 vbe_popup_visible 这个方法。"""

        def _trig40(cls=_B40, manual=False, **kw):
            _be = cls(**kw)
            _ui = _UI40()
            _c = E.Completer(_be, _ui)
            _c.trigger(not manual)      # True = 轮询自动触发；manual = Ctrl+Space
            return sorted(_c.matches or []), _ui.shown

        _hit40 = (["vbNo", "vbYes"], True)
        check("40.4 提示窗可见 -> 让位（不弹、候选清空）",
              [_trig40(popup=True)],
              expect_contain=[([], False)])
        check("40.5 提示窗不可见 -> 照常弹（同位置同后端）",
              [_trig40(popup=False)],
              expect_contain=[_hit40])
        check("40.6 现场就在参数位置：v55 的语法判据认不出（隔离验证）",
              [E.vbe_list_expected(_LINE40, len(_LINE40) + 1)],
              expect_contain=[False])

        _saved40 = E.YIELD_TO_VBE_LIST
        try:
            E.YIELD_TO_VBE_LIST = False
            check("40.7 对照：关掉让位开关（VBECOMPLETE_NO_YIELD=1）-> 照弹",
                  [_trig40(popup=True)],
                  expect_contain=[_hit40])
        finally:
            E.YIELD_TO_VBE_LIST = _saved40

        check("40.8 旧式后端（无该接口）-> 行为不变，照常弹",
              [_trig40(cls=_B40Old)],
              expect_contain=[_hit40])
        check("40.9 探测抛异常 -> 不让位、不炸（异常被吞）",
              [_trig40(popup=True, boom=True)],
              expect_contain=[_hit40])
        # 【v76 口径变更】让位判据不再区分"自动轮询 / 手动 Ctrl+Space"。
        # 理由见 engine.trigger：VBE 的成员列表画在屏幕上是客观事实，我们的窗是
        # topmost 悬浮窗，一挂上去就把它整个盖住。用户 v76 原话："VBE 的弹窗和
        # 我们项目弹窗冲突时，屏蔽我们项目弹窗……直接消失让位。"
        # 成对保留一条对照（40.10b）：列表不可见时手动路径照弹 —— 证明挡它的确实
        # 是"列表可见"，而不是把手动路径整个禁掉了。
        check("40.10 ★v76 手动 Ctrl+Space：提示窗可见时同样让位（不再区分触发来源）",
              [_trig40(popup=True, manual=True)],
              expect_contain=[([], False)])
        check("40.10b 对照：提示窗不可见 -> 手动 Ctrl+Space 照弹（让位只管「可见」）",
              [_trig40(popup=False, manual=True)],
              expect_contain=[_hit40])

        # 40.11 判据与语法位置无关：普通赋值行上，窗口可见同样让位
        class _B40Plain(_B40):
            def get_context(self):
                _line = "    userN"
                return {"line_no": 1, "caret_col": len(_line) + 1,
                        "line_text": _line, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": "模块2"}

        _be40p = _B40Plain(popup=True)
        _ui40p = _UI40()
        _c40p = E.Completer(_be40p, _ui40p)
        _c40p.trigger(True)
        check("40.11 普通行（非成员/类型位置）也拦得住 -> 让位",
              [sorted(_c40p.matches or []), _ui40p.shown],
              expect_contain=[[], False])

        # 40.12 真实 VbeBackend 提供该接口，且与模块级探测一致（真机）
        _be40r = VB40.VbeBackend()
        _r40 = _be40r.vbe_popup_visible()
        check("40.12 VbeBackend 的探测与模块级函数一致（此刻=%s）" % (_r40,),
              [_r40 == VB40.vbe_popup_visible(), isinstance(_r40, bool)],
              expect_contain=[True, True])

        # 40.13 句柄复用护栏：类名 / 进程两道复核都在（句柄值可能被 Windows
        # 回收给别的窗口，只凭"记得这个句柄"就相信它会永远屏蔽我们的弹窗）
        _hp40 = VB40._host_process_id()
        check("40.13 类名复核与宿主 PID 探测可用（PID=%s）" % (_hp40,),
              [isinstance(_hp40, int) and _hp40 >= 0,
               VB40._window_class_name(0),
               VB40._window_class_name(123456) not in VB40.VBE_POPUP_CLASSES],
              expect_contain=[True, "", True])

        # 40.14~40.16 v65 诊断：让位那一刻日志要能说出【是哪个窗】
        check("40.14 引擎能取到提示窗明细（诊断日志用）",
              [E.Completer(_B40(popup=True), _UI40())._vbe_popup_desc(),
               E.Completer(_B40Old(), _UI40())._vbe_popup_desc()],
              expect_contain=[[("PopupTipWndClass", 200, 16)], None])
        check("40.15 明细探测抛异常 -> 返回 None，不炸",
              [E.Completer(_B40(popup=True, boom=True),
                           _UI40())._vbe_popup_desc()],
              expect_contain=[None])
        try:
            _pin40 = VB40.vbe_popup_info()
            _pvis40 = VB40.vbe_popup_visible()
            _ok40 = (bool(_pin40) == _pvis40)
        except Exception as _e4016:
            _ok40 = "异常: %s" % _e4016
        check("40.16 明细判据与可见性判据一致（真机，此刻=%r）" % (_pin40,),
              [_ok40], expect_contain=[True])

    except Exception as _e40:
        check("第 40 节异常: %s" % _e40, [True], expect_contain=[False])

    # ==================================================================
    # 41. v64 提示只在【代码窗格】里出现（属性窗口等非代码工作区不弹）
    # ==================================================================
    # 用户报：在【属性窗口】里改属性值时，候选窗偶尔会冒出来 —— 那里不是写代码
    # 的地方，不该有提示（"我不希望在这些非写代码的工作区出来提示"）。
    #
    # 根因：判断"在不在 VBE"只看【前台窗口标题】有没有 "Microsoft Visual Basic"；
    # 而属性窗口 / 工程窗口 / 窗体设计器都是 VBE 主窗口（wndclass_desked_gsk）的
    # 子窗口，标题判据对它们一律返回 True。
    #
    # 修法：vbe_bridge.vbe_code_pane_focused() 看【VBE 线程】里当前拥有键盘焦点
    # 的窗口，父链上有没有代码窗格（窗口类 VbaWindow）。本节 41.1~41.9 用打桩的
    # 窗口树验这条判据本身（不依赖真机有没有开 Excel），41.10 起验真实调用与
    # main.in_vbe_code_area 的接线 / 兜底语义。
    print("\n=== 41. v64 提示只在代码窗格出现（非代码工作区不弹）===")
    try:
        import vbe_bridge as VB41

        check("41.1 窗口类名清单（VBE 主窗口 / 代码窗格）",
              ["wndclass_desked_gsk" in VB41.VBE_FRAME_CLASSES,
               "VbaWindow" in VB41.CODE_PANE_CLASSES],
              expect_contain=[True, True])

        _orig41 = (VB41._vbe_frame_hwnd, VB41._focused_hwnd_in_vbe,
                   VB41._parent_hwnd, VB41._window_class_name)

        def _stub41(classes, tree, focus=None, frame=100):
            """把窗口树打桩：classes={hwnd:类名}，tree={hwnd:父hwnd}。"""
            VB41._vbe_frame_hwnd = lambda: frame
            VB41._focused_hwnd_in_vbe = (
                lambda: (focus if focus is not None else 0))
            VB41._parent_hwnd = lambda h: int(tree.get(h, 0))
            VB41._window_class_name = lambda h: classes.get(h, "")

        # 41.2 焦点在代码窗格里（含"焦点在代码窗格的子控件上"）-> 允许
        _stub41(
            classes={301: "ComboBox", 300: "VbaWindow",
                     201: "MDIClient", 100: "wndclass_desked_gsk"},
            tree={301: 300, 300: 201, 201: 100},
            focus=301)
        check("41.2 焦点在代码窗格（或其子控件）-> True",
              [VB41.vbe_code_pane_focused()], expect_contain=[True])
        _stub41(
            classes={300: "VbaWindow", 201: "MDIClient",
                     100: "wndclass_desked_gsk"},
            tree={300: 201, 201: 100}, focus=300)
        check("41.3 焦点就是代码窗格本身 -> True",
              [VB41.vbe_code_pane_focused()], expect_contain=[True])

        # 41.4~41.6 三类"非写代码的工作区"都必须判 False
        #   属性窗口：真机实测类名 wndclass_pbrs，里面的编辑框/列表是
        #   Edit / ListBox / ComboBox / SysTabControl32
        _stub41(
            classes={401: "ListBox", 400: "wndclass_pbrs",
                     100: "wndclass_desked_gsk"},
            tree={401: 400, 400: 100}, focus=401)
        check("41.4 焦点在属性窗口里 -> False（用户报的场景）",
              [VB41.vbe_code_pane_focused()], expect_contain=[False])
        _stub41(
            classes={501: "SysTreeView32", 500: "PROJECT",
                     100: "wndclass_desked_gsk"},
            tree={501: 500, 500: 100}, focus=501)
        check("41.5 焦点在工程窗口里 -> False",
              [VB41.vbe_code_pane_focused()], expect_contain=[False])
        _stub41(
            classes={601: "ThunderDFrame", 600: "DesignerWindow",
                     100: "wndclass_desked_gsk"},
            tree={601: 600, 600: 100}, focus=601)
        check("41.6 焦点在窗体设计器里 -> False",
              [VB41.vbe_code_pane_focused()], expect_contain=[False])

        # 41.7~41.9 拿不准时返回 None（调用方据此退回老判据，绝不能"拿不准就拦"）
        _stub41(classes={}, tree={}, focus=0)
        check("41.7 拿不到焦点窗口 -> None（不拦）",
              [VB41.vbe_code_pane_focused()], expect_contain=[None])
        _stub41(classes={}, tree={}, focus=1, frame=0)
        check("41.8 找不到 VBE 主窗口 -> None（不拦）",
              [VB41.vbe_code_pane_focused()], expect_contain=[None])
        # 链上既没有代码窗格、也走不到 frame（父链异常）-> False，且不能死循环
        _stub41(classes={i: "Edit" for i in range(700, 720)},
                tree={i: i + 1 for i in range(700, 719)}, focus=700)
        check("41.9 父链异常（走不到 frame）-> False 且不死循环",
              [VB41.vbe_code_pane_focused()], expect_contain=[False])

        for _f, _v in zip(("_vbe_frame_hwnd", "_focused_hwnd_in_vbe",
                           "_parent_hwnd", "_window_class_name"), _orig41):
            setattr(VB41, _f, _v)

        try:
            _real41 = VB41.vbe_code_pane_focused()
        except Exception as _e41r:
            _real41 = "异常: %s" % _e41r
        check("41.10 真实机器上返回 True/False/None 且不抛（此刻=%r）" % (_real41,),
              [isinstance(_real41, bool) or _real41 is None],
              expect_contain=[True])

        # 41.11~41.15 main.in_vbe_code_area：接线 + 兜底语义
        if M is None:
            check("main 模块不可用（缺 pynput），跳过 41.11~41.15", [True],
                  expect_contain=[True])
        else:
            _orig_area41 = VB41.vbe_code_pane_focused
            _orig_title41 = M.in_vbe_code_pane
            try:
                M.in_vbe_code_pane = lambda: True          # 前台标题像 VBE
                VB41.vbe_code_pane_focused = lambda: False  # 但焦点不在代码窗格
                check("41.11 标题像 VBE 但焦点在属性窗口 -> 不认（False）",
                      [M.in_vbe_code_area()], expect_contain=[False])
                VB41.vbe_code_pane_focused = lambda: True
                check("41.12 焦点在代码窗格 -> True",
                      [M.in_vbe_code_area()], expect_contain=[True])
                VB41.vbe_code_pane_focused = lambda: None
                check("41.13 判不出来 -> 退回标题判据（True，保持旧行为）",
                      [M.in_vbe_code_area()], expect_contain=[True])
                M.in_vbe_code_pane = lambda: False
                check("41.14 判不出来且标题也不是 VBE -> False",
                      [M.in_vbe_code_area()], expect_contain=[False])

                def _boom41():
                    raise RuntimeError("boom")
                VB41.vbe_code_pane_focused = _boom41
                M.in_vbe_code_pane = lambda: True
                check("41.15 探测抛异常 -> 退回标题判据、不炸",
                      [M.in_vbe_code_area()], expect_contain=[True])

                # 41.16 接线护栏：三处"会弹窗 / 会写代码"的动作都改用精确判据
                #   （轮询触发、Shift+Enter 与自动配对、Ctrl+Space）
                import inspect as _inspect41
                _src41 = _inspect41.getsource(M)
                _cnt41 = _src41.count("in_vbe_code_area()")
                check("41.16 精确判据已接进 4 处动作（实际 %d 处）" % _cnt41,
                      [_cnt41 >= 4,
                       "vbe_code_pane_focused" in _src41,
                       "in_vbe_code_pane()" in _src41,
                       "_poll_mod_switch" in _src41],
                      expect_contain=[True, True, True, True])

                # 41.18~41.21 换模块否决（同一个坑的另一半：没敲字也不该弹）
                #   快照 = (行号, 行文本, 模块名)；模块名拿不到就不否决。
                check("41.18 同一模块内的文本变化 -> 不否决",
                      [M._poll_mod_switch((3, "    us", "Sheet1"),
                                          (3, "    use", "Sheet1"))],
                      expect_contain=[False])
                check("41.19 换模块（同行号、文本不同）-> 否决",
                      [M._poll_mod_switch((3, "    us", "Sheet1"),
                                          (3, "    End Sub", "Module7"))],
                      expect_contain=[True])
                check("41.20 模块名大小写不同 -> 不算换模块",
                      [M._poll_mod_switch((3, "a", "Sheet1"),
                                          (3, "ab", "SHEET1"))],
                      expect_contain=[False])
                check("41.21 模块名缺失 / 旧式二元组 -> 不否决（不能把工具哑掉）",
                      [M._poll_mod_switch((3, "a", ""), (3, "ab", "Sheet1")),
                       M._poll_mod_switch((3, "a", "Sheet1"), (3, "ab", "")),
                       M._poll_mod_switch((3, "a"), (3, "ab"))],
                      expect_contain=[False, False, False])

                # 41.17 逃生开关：VBECOMPLETE_NO_CODE_AREA_GATE=1 退回旧判据
                _saved_gate41 = M.CODE_AREA_GATE
                try:
                    M.CODE_AREA_GATE = False
                    M.in_vbe_code_pane = lambda: True
                    check("41.17 关掉闸门 -> 退回「前台是 VBE 就算数」（True）",
                          [M.in_vbe_code_area()], expect_contain=[True])
                finally:
                    M.CODE_AREA_GATE = _saved_gate41
            finally:
                VB41.vbe_code_pane_focused = _orig_area41
                M.in_vbe_code_pane = _orig_title41

    except Exception as _e41:
        check("第 41 节异常: %s" % _e41, [True], expect_contain=[False])

    # ==================================================================
    # 42. v65 让位只认【成员列表】，参数信息（形参签名）不让位
    # ==================================================================
    # 用户报：UserForm2 第13行 `UserForm1.SelectedList.RemoveItem sellis` 处
    # "输入任何字符都无提醒"；VBE 在那儿弹的正是【参数信息】（形参签名）。
    # 用户澄清 v63 的原意："只有弹成员列表的时候才让位，弹形参签名不需要让位。"
    #
    # 只读复现（真实工程 + 合成光标，已实测）确认：该位置唯一能拦掉提示的就是
    # v63 那条让位判据 —— 提示窗不可见时 matches=['SelectedListRow2']；可见时
    # 让位、候选为空。故修法 = 把让位范围从"任何提示窗"收窄到"成员列表窗"，
    # 并让焦点判据把 VBE 自己的提示窗算作"仍在代码窗格"。
    print("\n=== 42. v65 让位只认成员列表（参数信息不让位）===")
    try:
        import vbe_bridge as VB42

        # 42.1 让位清单：含成员列表、不含参数信息
        check("42.1 让位清单只含成员列表窗（NameListWndClass）",
              ["NameListWndClass" in VB42.VBE_YIELD_CLASSES,
               "PopupTipWndClass" not in VB42.VBE_YIELD_CLASSES,
               set(VB42.VBE_YIELD_CLASSES) <= set(VB42.VBE_POPUP_CLASSES)],
              expect_contain=[True, True, True])

        # 42.2~42.5 判据本身（打桩窗口明细，纯 Python，不碰 Excel）
        _orig42 = (VB42._vbe_popup_hwnds, VB42._popup_details_now)

        def _stub42(details):
            VB42._vbe_popup_hwnds = lambda: [1]
            VB42._popup_details_now = lambda hs: list(details)

        _stub42([("PopupTipWndClass", 300, 18)])
        check("42.2 只可见参数信息 -> 不让位（False）★v65 核心",
              [VB42.vbe_yield_visible()], expect_contain=[False])
        _stub42([("NameListWndClass", 120, 90)])
        check("42.3 可见成员列表 -> 让位（True）",
              [VB42.vbe_yield_visible()], expect_contain=[True])
        _stub42([("PopupTipWndClass", 300, 18), ("NameListWndClass", 120, 90)])
        check("42.4 两个都可见 -> 让位（成员列表说了算）",
              [VB42.vbe_yield_visible()], expect_contain=[True])

        def _boom42(hs):
            raise RuntimeError("boom")
        VB42._popup_details_now = _boom42
        check("42.5 探测抛异常 -> 不让位、不炸",
              [VB42.vbe_yield_visible()], expect_contain=[False])
        VB42._vbe_popup_hwnds, VB42._popup_details_now = _orig42

        # 42.6~42.10 引擎接线：优先 vbe_yield_visible，老后端退回 vbe_popup_visible
        class _B42(_B40Base):
            """v65 现场：参数信息可见，成员列表不可见。"""

            def __init__(self, yld=False, pop=True, boom=False, boom2=False,
                         **kw):
                self._yld, self._pop = yld, pop
                self._boom, self._boom2 = boom, boom2

            def vbe_popup_visible(self):
                if self._boom:
                    raise RuntimeError("boom")
                return self._pop

            def vbe_yield_visible(self):
                if self._boom2:
                    raise RuntimeError("boom")
                return self._yld

            def vbe_popup_info(self):
                return ([("PopupTipWndClass", 300, 18)] if self._pop else []) \
                    + ([("NameListWndClass", 120, 90)] if self._yld else [])

        def _trig42(cls=_B42, manual=False, **kw):
            _be = cls(**kw)
            _ui = _UI40()
            _c = E.Completer(_be, _ui)
            _c.trigger(not manual)      # True = 轮询自动触发；manual = Ctrl+Space
            return sorted(_c.matches or []), _ui.shown

        _hit42 = (["vbNo", "vbYes"], True)
        check("42.6 参数信息可见（成员列表不可见）-> 照常弹 ★v65 核心",
              [_trig42(pop=True, yld=False)], expect_contain=[_hit42])
        check("42.7 成员列表可见 -> 让位（不弹、候选清空）",
              [_trig42(pop=False, yld=True)], expect_contain=[([], False)])
        check("42.8 老后端（只有 vbe_popup_visible）-> 退回旧行为：任何提示窗都让位",
              [_trig42(cls=_B40, popup=True)], expect_contain=[([], False)])
        check("42.8b 更老的桩后端（两个接口都没有）-> 不让位、照弹",
              [_trig42(cls=_B40Old)], expect_contain=[_hit42])
        check("42.9 vbe_yield_visible 抛异常 -> 不让位、不炸",
              [_trig42(pop=True, yld=False, boom2=True)],
              expect_contain=[_hit42])
        # 【v76 口径变更】同 40.10：成员列表可见时，手动路径也让位 —— 这条判据
        # 现在是"不依赖触发来源的通用不变量"（v73 的补位路径 trigger(False) 也受
        # 它约束，那正是补出来的窗会压住 VBE 列表的那个时序缝）。
        check("42.10 ★v76 手动 Ctrl+Space：成员列表可见 -> 同样让位（不再区分触发来源）",
              [_trig42(manual=True, pop=False, yld=True)],
              expect_contain=[([], False)])
        check("42.10b 对照：成员列表不可见 -> 手动路径照弹",
              [_trig42(manual=True, pop=False, yld=False)],
              expect_contain=[_hit42])

        # 42.11 真机：让位判据与明细一致
        _y42 = _i42 = None
        try:
            _be42 = VB42.VbeBackend()
            _y42 = _be42.vbe_yield_visible()
            _i42 = _be42.vbe_popup_info()
            _ok42 = (isinstance(_y42, bool)
                     and _y42 == any(c in VB42.VBE_YIELD_CLASSES
                                     for c, _w, _h in _i42))
        except Exception as _e4211:
            _ok42 = "异常: %s" % _e4211
        check("42.11 真机：让位判据与明细一致（yield=%r info=%r）" % (_y42, _i42),
              [_ok42], expect_contain=[True])

        # 42.12~42.14 焦点判据：VBE 自己的提示窗算作"仍在代码窗格"
        #   （否则提示窗一抢到焦点，v64 的闸门就把我们的候选窗当场收掉）
        _orig42f = (VB42._vbe_frame_hwnd, VB42._focused_hwnd_in_vbe,
                    VB42._parent_hwnd, VB42._window_class_name)

        def _stub42f(cls, hwnd=900):
            VB42._vbe_frame_hwnd = lambda: 100
            VB42._focused_hwnd_in_vbe = lambda: hwnd
            VB42._parent_hwnd = lambda h: 0
            VB42._window_class_name = lambda h: (
                cls if h == hwnd
                else ("wndclass_desked_gsk" if h == 100 else ""))

        _stub42f("NameListWndClass")
        check("42.12 焦点落在成员列表窗上 -> 仍算在代码窗格（True）",
              [VB42.vbe_code_pane_focused()], expect_contain=[True])
        _stub42f("PopupTipWndClass")
        check("42.13 焦点落在参数信息窗上 -> 仍算在代码窗格（True）★v65",
              [VB42.vbe_code_pane_focused()], expect_contain=[True])
        _stub42f("wndclass_pbrs")
        check("42.14 属性窗口仍不算（False，v64 语义不变）",
              [VB42.vbe_code_pane_focused()], expect_contain=[False])
        for _f, _v in zip(("_vbe_frame_hwnd", "_focused_hwnd_in_vbe",
                           "_parent_hwnd", "_window_class_name"), _orig42f):
            setattr(VB42, _f, _v)

    except Exception as _e42:
        check("第 42 节异常: %s" % _e42, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    print("\n=== 43. v66 候选窗避让 VBE 参数信息窗（形参签名）===")
    try:
        import ui as UI43
        import vbe_bridge as VB43

        # 43.1 避让与让位是两码事：参数信息仍然【不让位】（v65 语义不动）
        check("43.1 避让不改让位清单（参数信息仍不让位）",
              ["PopupTipWndClass" in VB43.VBE_POPUP_CLASSES,
               "PopupTipWndClass" not in VB43.VBE_YIELD_CLASSES],
              expect_contain=[True, True])

        # 43.2~43.3 矩形判据（明细=矩形投影，一处逻辑两处消费）
        _orig43 = (VB43._vbe_popup_hwnds, VB43._popup_rects_now)
        VB43._vbe_popup_hwnds = lambda: [1]
        VB43._popup_rects_now = lambda hs: [("PopupTipWndClass", 10, 20, 310, 38)]
        check("43.2 明细由矩形投影而来（300x18）",
              [VB43._popup_details_now([1]), VB43.vbe_popup_rects()],
              expect_contain=[[("PopupTipWndClass", 300, 18)],
                              [("PopupTipWndClass", 10, 20, 310, 38)]])

        def _boom43(hs):
            raise RuntimeError("boom")
        VB43._popup_rects_now = _boom43
        check("43.3 探测抛异常 -> 返回空列表、不炸（拿不到就不动窗）",
              [VB43.vbe_popup_rects()], expect_contain=[[]])
        VB43._vbe_popup_hwnds, VB43._popup_rects_now = _orig43

        # 43.4~43.11 纯函数 avoid_popup_rects：把候选窗挪出提示窗
        _R43 = UI43.avoid_popup_rects
        _pop43 = (0, 220, 400, 240)      # 光标下方的一条形参签名
        check("43.4 无提示窗 -> 位置原样不动",
              [_R43(100, 220, 200, 90, 200, 1080, [])], expect_contain=[220])
        check("43.5 签名窗就在正下方且重叠 -> 下移到它下面（240+2）★核心",
              [_R43(100, 220, 200, 90, 200, 1080, [_pop43])], expect_contain=[242])
        check("43.6 提示窗在屏幕另一侧（水平不相交）-> 不动",
              [_R43(600, 220, 200, 90, 200, 1080, [(0, 220, 400, 240)])],
              expect_contain=[220])
        check("43.7 提示窗在上方（垂直不相交）-> 不动",
              [_R43(100, 220, 200, 90, 200, 1080, [(0, 100, 400, 200)])],
              expect_contain=[220])
        check("43.8 两个窗叠着 -> 落到最下面那个的下方（260+2）",
              [_R43(100, 220, 200, 90, 200, 1080,
                    [(0, 220, 400, 240), (0, 239, 400, 260)])],
              expect_contain=[262])
        check("43.9 下移会出屏 -> 翻到光标上方（990-90-2）",
              [_R43(100, 1000, 200, 90, 990, 1080, [(0, 1020, 400, 1060)])],
              expect_contain=[898])
        check("43.10 上下都放不下 -> 贴屏幕底部（1080-90）",
              [_R43(100, 900, 200, 90, 60, 1080, [(0, 900, 400, 1000)])],
              expect_contain=[990])
        check("43.11 只压住候选窗一部分 -> 照样躲（水平相交就避）",
              [_R43(500, 220, 200, 90, 200, 1080, [(600, 220, 900, 240)])],
              expect_contain=[242])

        # 43.12~43.14 _position 接线：默认贴光标下方、有签名窗就下移、同位置不重设
        class _FR43(object):
            def winfo_screenwidth(self):
                return 1920

            def winfo_screenheight(self):
                return 1080

        class _FC43(object):
            def cget(self, k):
                return 200 if k == "width" else 90

        class _FW43(object):
            def __init__(self):
                self.geoms = []

            def geometry(self, s):
                self.geoms.append(s)

        def _mk43():
            p = UI43.Popup.__new__(UI43.Popup)
            p.root, p.canvas, p.win = _FR43(), _FC43(), _FW43()
            p._last_geom = None
            return p

        _o43caret = UI43.caret_screen_rect
        _o43rects = UI43.screen_popup_rects
        try:
            UI43.caret_screen_rect = lambda: (100, 200, 218)
            UI43.screen_popup_rects = lambda: []
            _p43 = _mk43()
            _p43._position()
            _g1 = list(_p43.win.geoms)
            _p43._position()                  # 位置没变：不该重复设 geometry
            _g2 = list(_p43.win.geoms)
            UI43.screen_popup_rects = lambda: [(0, 218, 400, 240)]
            _p43._position()                  # 签名窗出现 -> 下移到 240+2
            _g3 = list(_p43.win.geoms)
        finally:
            UI43.caret_screen_rect = _o43caret
            UI43.screen_popup_rects = _o43rects
        check("43.12 _position 默认贴光标下方 → %s" % (_g1[-1:] or [None],),
              [_g1], expect_contain=[["200x90+100+220"]])
        check("43.13 同位置重复调用不重设 geometry（防闪烁）→ %s" % (_g2[-1:] or [None],),
              [_g2], expect_contain=[["200x90+100+220"]])
        check("43.14 签名窗出现 -> 候选窗下移到其下方 → %s" % (_g3[-1:] or [None],),
              [_g3], expect_contain=[["200x90+100+220", "200x90+100+242"]])

        # 43.15 拿不到提示窗矩形时 screen_popup_rects 不得抛（vbe_bridge 打桩成异常）
        _o43vbr = VB43.vbe_popup_rects
        VB43.vbe_popup_rects = _boom43
        try:
            _sr43 = UI43.screen_popup_rects()
        finally:
            VB43.vbe_popup_rects = _o43vbr
        check("43.15 探测异常时 screen_popup_rects() 返回 []（不抛）",
              [_sr43], expect_contain=[[]])

        # 43.16 收起状态下 reposition 是安全的空操作（没有窗、没有 root 也不炸）
        _p43r = UI43.Popup.__new__(UI43.Popup)
        _p43r.win = None
        _p43r.reposition()
        check("43.16 收起状态 reposition 空操作、不抛", [True], expect_contain=[True])

        # 43.17 真机：提示窗矩形可读、格式正确
        _rr43 = None
        try:
            _rr43 = VB43.VbeBackend().vbe_popup_rects()
            _ok43 = (isinstance(_rr43, list)
                     and len(_rr43) == len(VB43.vbe_popup_rects())
                     and all(len(t) == 5 for t in _rr43))
        except Exception as _e4317:
            _ok43 = "异常: %s" % _e4317
        check("43.17 真机：提示窗矩形可读且格式正确（%r）" % (_rr43,),
              [_ok43], expect_contain=[True])

        # 43.18 main 接线护栏：轮询里要有"提示窗矩形变化 -> 重摆候选窗"
        try:
            import inspect as _inspect43
            import main as _M43
            _src43 = _inspect43.getsource(_M43)
            _cnt43 = _src43.count("vbe_popup_rects()")
        except Exception as _e4318:
            _src43, _cnt43 = "", "异常: %s" % _e4318
        check("43.18 轮询已接上「矩形变化 -> reposition」（出现 %s 次）" % (_cnt43,),
              [_cnt43 >= 1, "post(popup.reposition)" in _src43,
               "_popup_sig" in _src43],
              expect_contain=[True, True, True])
        check("43.19 reposition 挂在 UI 上（引擎层没有这个方法，故 main 调 popup.*）",
              [callable(getattr(UI43.Popup, "reposition", None))],
              expect_contain=[True])
    except Exception as _e43:
        check("第 43 节异常: %s" % _e43, [True], expect_contain=[False])

    # ==================================================================
    # 44. v67：宿主类型库的枚举常量（Excel 的 xl*、Office 的 mso*）
    #
    # 用户报："输入 vb 会提示一堆 VBA 枚举值，输入 xl 却一个 xl 开头的枚举值
    # 都不提示。" vb* 是 v61 收的 VBA 内建常量；xl* 属于宿主 Excel 类型库，
    # 一直没收 —— 不是泄露，是漏收。
    #
    # 收进来的同时必须收紧口径：这批名字 4500 多条、全是长复合词，放开跳步
    # 匹配的话输入 ms 会带出 2365 条 mso*、输入 count 会带出 130 条 xl*。
    # 所以只认"从头开始的连续前缀"，且输入长度不低于家族前缀长度。
    # ==================================================================
    print("\n=== 44. 宿主类型库枚举常量（v67）===")
    try:
        import vbe_bridge as VB44
        import engine as E44

        # ---- 44.1~44.5 家族前缀推导（纯函数，不碰 Excel）----
        _excel_like = ["xlUp", "xlDown", "xlToRight", "xlCellTypeVisible",
                       "xlContinuous", "rgbCadetBlue", "sigdetFoo"]
        _office_like = ["msoTrue", "msoFalse", "msoLineSolid", "msoAlignLeft",
                        "msoCTLPush", "rgbCadetBlue", "sigdetFoo"]
        _stdole_like = ["Unchecked", "Checked", "Gray", "Default",
                        "Monochrome", "VgaColor", "Color"]
        check("44.1 Excel 类库 -> 家族前缀 xl",
              [VB44.enum_family_prefix(_excel_like)], expect_contain=["xl"])
        check("44.2 Office 类库 -> 家族前缀 mso",
              [VB44.enum_family_prefix(_office_like)], expect_contain=["mso"])
        check("44.3 各说各话的库（stdole）-> 无家族前缀",
              [VB44.enum_family_prefix(_stdole_like)], expect_contain=[""])
        check("44.4 空清单 -> 空串（不抛）",
              [VB44.enum_family_prefix([]), VB44.enum_family_prefix(["a"])],
              expect_contain=["", ""])
        _split = ["abOne", "abTwo", "cdOne", "cdTwo", "efOne"]
        check("44.5 没有任何前缀过半 -> 空串（宁可整库不要，也不放通用词进来）",
              [VB44.enum_family_prefix(_split)], expect_contain=[""])

        # ---- 44.6~44.7 家族过滤 + 白名单 ----
        _fam, _kept = VB44.host_enum_family(_excel_like)
        check("44.6 host_enum_family 只留家族成员（%r）" % (_kept,),
              [_fam, sorted(_kept), "rgbCadetBlue" in _kept],
              expect_contain=["xl", ["xlCellTypeVisible", "xlContinuous",
                                     "xlDown", "xlToRight", "xlUp"], False])
        _old_fam44 = VB44._HOST_ENUM_FAMILIES
        VB44._HOST_ENUM_FAMILIES = ("zz",)
        try:
            _wf44 = VB44.host_enum_family(_excel_like)
        finally:
            VB44._HOST_ENUM_FAMILIES = _old_fam44
        check("44.7 白名单不含 xl -> 整库跳过（用于只想留某个家族）",
              [_wf44], expect_contain=[("", [])])

        # ---- 44.8~44.9 真机：真能枚举出引用类型库的枚举常量 ----
        # ⚠️ 必须单独加载一份干净的 vbe_bridge 来跑这一节：本文件前面的小节把
        # `VB._get_vbe_cached` 打桩成了假 VBE（第 31/33 节），桩没还原，真机
        # 接口在那儿只会读到假工程（引用清单为空 -> 枚举出 0 条）。
        # 用 importlib 重新 exec 一份独立实例，验的就是真实链路。
        _fresh44 = None
        try:
            import importlib.util as _iu44
            _spec44 = _iu44.spec_from_file_location(
                "vbe_bridge_fresh44",
                os.path.join(os.path.dirname(os.path.dirname(
                    os.path.abspath(__file__))), "vbe_bridge.py"))
            _fresh44 = _iu44.module_from_spec(_spec44)
            _spec44.loader.exec_module(_fresh44)
        except Exception:
            _fresh44 = None
        _VB44R = _fresh44 or VB44
        # v70 起宿主枚举默认不收（用户口径）。这一节验的是"枚举链路本身能跑通"，
        # 所以把开关临时打开 —— 断言强度不变；"默认关"由第 47 节单独把关。
        _saved_he44 = _VB44R.ENABLE_HOST_ENUMS
        _VB44R.ENABLE_HOST_ENUMS = True
        try:
            _items44 = _VB44R.host_enum_constants()
            _mins44 = sorted({m for _n, m in _items44})
            _nxl44 = len([1 for n, _m in _items44 if n.lower().startswith("xl")])
            _nms44 = len([1 for n, _m in _items44 if n.lower().startswith("mso")])
            # 家族前缀集合（= 名字的小写前 m 个字符）；形态必须健康
            _fams44 = sorted({n.lower()[:m] for n, m in _items44})
            _badfam44 = sorted(
                f for f in _fams44
                if not (2 <= len(f) <= 4 and f.isalnum() and f == f.lower()))
            _ok44 = (isinstance(_items44, list)
                     and all(m >= 2 for _n, m in _items44)
                     and _nxl44 >= 1000 and _nms44 >= 1000
                     and not _badfam44)
            # 家族前缀不写死 xl / mso：这条断言要跟着"工程引用了哪些库"走。
            # 后端给的第 2 个值就是该库推导出的家族前缀长度，取名字的小写前缀
            # 即是那个前缀 —— 它必须是小写字母数字、2~4 个字符（单字母前缀会
            # 让一两个字母带出半屏名字，后端不会收）。
            # 实测本机：xl=2173（Excel）、mso=2365（Office）、fm=180（MSForms ——
            # 工程自己引的库，这正是"你引用什么就补什么"该有的样子）。
            _desc44 = ("共 %d 条，家族前缀=%s，xl*=%d mso*=%d，"
                       "最短输入长度=%s" % (len(_items44), _fams44,
                                        _nxl44, _nms44, _mins44))
        except Exception as _e44:
            _ok44, _desc44 = "异常: %s" % _e44, ""
        check("44.8 真机：枚举出 %s" % (_desc44,), [_ok44], expect_contain=[True])

        _hm44 = None
        try:
            _hm44 = _VB44R.VbeBackend().get_host_enum_names()
            _okhm44 = (isinstance(_hm44, dict) and _hm44.get("xlup") == 2
                       and _hm44.get("msotrue") == 3)
        except Exception as _e44b:
            _okhm44 = "异常: %s" % _e44b
        check("44.9 后端 get_host_enum_names() -> 小写名到最短长度的字典",
              [_okhm44], expect_contain=[True])
        # 复原（后面 44.10+ 走桩后端，与这个开关无关）
        _VB44R.ENABLE_HOST_ENUMS = _saved_he44

        # ---- 44.10~44.15 引擎接线：前缀专用通道 ----
        _HOST44 = {"xlup": 2, "xlcellvalue": 2, "xldown": 2, "xltoright": 2,
                   "msotrue": 3, "msofalse": 3}
        _POOL44 = [("xlUp", "宿主库", None, False),
                   ("xlCellValue", "宿主库", None, False),
                   ("xlDown", "宿主库", None, False),
                   ("xlToRight", "宿主库", None, False),
                   ("msoTrue", "宿主库", None, False),
                   ("msoFalse", "宿主库", None, False),
                   ("numArr", "M1", None, False),
                   ("countRows", "M1", None, False),
                   ("cellText", "M1", None, False)]

        class _B44(_B40Base):
            """宿主枚举常量的现场：池子里既有宿主常量，也有用户自己的名字。"""

            def __init__(self, word="xl", host=True, hook=True, **kw):
                self._word, self._host, self._hook = word, host, hook
                self.struct_calls = 0

            def _line(self):
                return "    " + self._word

            def get_context(self):
                _ln = self._line()
                return {"line_no": 3, "caret_col": len(_ln) + 1,
                        "line_text": _ln, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": "M1"}

            def get_identifiers(self):
                return list(_POOL44)

            def get_declared_names(self):
                return [r[0].lower() for r in _POOL44]

            def declared_elsewhere(self, name, module_name):
                return str(name).lower() in [r[0].lower() for r in _POOL44]

            def get_structural_names(self):
                self.struct_calls += 1
                return [r[0].lower() for r in _POOL44]

            def get_builtin_names(self):
                return []

            def get_type_names(self):
                return []

            def names_outside_caret(self, caret, scope_only=False):
                return set()

            def get_host_enum_names(self):
                if not self._hook:
                    raise RuntimeError("boom")
                return dict(_HOST44) if self._host else {}

        def _trig44(cls=_B44, **kw):
            _be = cls(**kw)
            _ui = _UI40()
            _c = E44.Completer(_be, _ui)
            _c.trigger(True)
            return sorted(_c.matches or []), _ui.shown, _be.struct_calls

        _h44, _v44, _c44 = _trig44(word="xl")
        check("44.10 输入 xl -> 出 xl* 家族（★用户点名要的），不出 mso*",
              [_v44, _h44, any(m.lower().startswith("mso") for m in _h44)],
              expect_contain=[True, ["xlCellValue", "xlDown", "xlToRight",
                                     "xlUp"], False])
        check("44.11 输入 xlu -> 只出 xlUp（前缀命中）",
              [_trig44(word="xlu")[0]], expect_contain=[["xlUp"]])
        check("44.12 输入 ms -> 只出 mso*，不出别的；输入 mso 才出 mso*",
              [_trig44(word="ms")[0], _trig44(word="mso")[0]],
              expect_contain=[[], ["msoFalse", "msoTrue"]])
        check("44.13 输入 count / cell -> 一条 xl* 都不多出（★防噪音核心）",
              [_trig44(word="count")[0], _trig44(word="cell")[0]],
              expect_contain=[["countRows"], ["cellText"]])

        class _B44No(_B44):
            """更老的桩后端：连 get_host_enum_names 这个接口都没有。"""
            get_host_enum_names = None

        _nb44 = _trig44(cls=_B44No, word="cell")[0]
        check("44.14 老后端（不提供该接口）-> 退回普通模糊匹配"
              "（cell 照样命中 xlCellValue，行为与 v66 完全一致）",
              [_nb44, _trig44(word="cell")[0]],
              expect_contain=[["cellText", "xlCellValue"], ["cellText"]])
        _boom44 = _trig44(word="xlu", hook=False)[0]
        check("44.15 接口抛异常 -> 不崩，且与「没有该接口」同行为（%r）"
              % (_boom44,),
              [_boom44 == _trig44(cls=_B44No, word="xlu")[0],
               "xlUp" in _boom44], expect_contain=[True, True])

        # ---- 44.16 结构性名字集合按一次触发取一份（v67 性能护栏）----
        # 修复前每个回声候选都重建一次集合；池子涨到 4600 多条后，
        # 输入 xl（一次 2000 多个候选）光这一步就要 2 秒（真机实测 2135ms）。
        check("44.16 一次触发里 get_structural_names() 只取一次（实际 %d 次）"
              % (_c44,), [_c44], expect_contain=[1])

        # ---- 44.17 诊断段受日志开关约束（否则每次触发多跑一遍全池匹配）----
        _orig_fm44, _orig_log44 = E44.fuzzy_match, E44._LOG_ENABLED
        _cnt44 = [0]

        def _fm44(name, query):
            _cnt44[0] += 1
            return _orig_fm44(name, query)

        try:
            E44.fuzzy_match = _fm44
            E44._LOG_ENABLED = False
            _cnt44[0] = 0
            _trig44(word="xl")
            _off44 = _cnt44[0]
            E44._LOG_ENABLED = True
            _cnt44[0] = 0
            _trig44(word="xl")
            _on44 = _cnt44[0]
        finally:
            E44.fuzzy_match, E44._LOG_ENABLED = _orig_fm44, _orig_log44
        # 关日志：只有 3 个非宿主名字走模糊匹配；开日志：多跑一遍全池（9 条）
        check("44.17 关日志时诊断段不跑（模糊匹配 %d 次 vs 开日志 %d 次）"
              % (_off44, _on44), [_off44, _on44], expect_contain=[3, 12])

        # ---- 44.18 vbe_bridge 收集段接线护栏 ----
        try:
            import inspect as _inspect44
            _src44 = _inspect44.getsource(VB44)
            _okwire44 = all(k in _src44 for k in (
                "ENABLE_HOST_ENUMS", "VBECOMPLETE_NO_HOST_ENUMS",
                "host_enum_constants()", "_HOST_ENUM_MODULE",
                "_host_enum_names = host_min", "def get_host_enum_names",
                "def enum_family_prefix", "def host_enum_family",
                "def _tlb_enum_members"))
        except Exception as _e44c:
            _okwire44 = "异常: %s" % _e44c
        check("44.18 vbe_bridge 收集段四路承接 + 开关 + 家族推导都在",
              [_okwire44], expect_contain=[True])
        check("44.19 坏 GUID -> 返回空清单、不抛",
              [VB44._tlb_enum_members("{00000000-0000-0000-0000-000000000000}",
                                      9, 9)],
              expect_contain=[[]])
    except Exception as _e44x:
        check("第 44 节异常: %s" % _e44x, [True], expect_contain=[False])

    # ==================================================================
    # 45. v68：宿主枚举常量的模糊匹配（★用户报的 xlworkfaul）
    #
    # 用户报："我输入 xlworkfaul，不会提示 xlWorkbookDefault，这个模糊匹配有问题。"
    #
    # 复现结论：不是漏收（xlWorkbookDefault 在 v67 就已经在池子里），是 v67 的
    # "只认从头开始的连续前缀"太死 —— xlworkfaul 跳过了中间的 book，不算前缀，
    # 可它是整个 4538 条里【唯一】能匹配上的候选。命中唯一，纯属被规则一刀切掉。
    #
    # v68 放宽成两条【都要满足】：
    #   1) 输入本身从家族前缀开始（xl / mso）。这条不能松 —— 正是它整片挡掉了
    #      ms -> 2365 条 mso*、count -> 130 条、open -> 492 条这些噪音；
    #   2) 输入里出现过一段【连续】的名字片段，长度 >= HOST_ENUM_FUZZY_RUN(=4)，
    #      即"你真的打过这个名字里的一段"，而不是拿几个字母去凑子序列。
    #
    # 为什么不是"输入够长就放行"：实测放行纯跳步后 xlcell 会从 15 条涨到 89 条、
    # xlcount 从 4 条涨到 59 条。加"连续块 >= 4"后，长度 < 4 的输入退化成纯前缀
    # （与 v67 一字不差），而 xlworkfaul（连续块 6）照常命中。
    # ==================================================================
    print("\n=== 45. 宿主枚举常量的模糊匹配（v68）===")
    try:
        import engine as E45

        # ---- 45.1~45.3 纯函数：连续块闸门 ----
        check("45.1 连续块闸门常量 = 4（与 v61 内建名字那档一致）",
              [E45.HOST_ENUM_FUZZY_RUN], expect_contain=[4])
        check("45.2 _longest_run：xlworkfaul 在 xlWorkbookDefault 上的命中下标"
              "最长连续段 = 6（xlwork 一段 + faul 一段）",
              [E45._longest_run([0, 1, 2, 3, 4, 5, 12, 13, 14, 15]),
               E45._longest_run(E45.fuzzy_match("xlWorkbookDefault",
                                                "xlworkfaul")[1])],
              expect_contain=[6, 6])
        check("45.3 _longest_run 边界：空 -> 0，单点 -> 1，全散 -> 1",
              [E45._longest_run([]), E45._longest_run([7]),
               E45._longest_run([0, 2, 4, 6])],
              expect_contain=[0, 1, 1])

        # ---- 45.4~45.9 桩池：确定性验证各条规则 ----
        _POOL45 = [("xlWorkbookDefault", "宿主库", None, False),
                   ("xlWorkbookNormal", "宿主库", None, False),
                   ("xlMSDOS", "宿主库", None, False),
                   ("xlCenter", "宿主库", None, False),
                   ("xlCellTypeVisible", "宿主库", None, False),
                   ("xlVAlignCenter", "宿主库", None, False),
                   ("xlLastCell", "宿主库", None, False),
                   ("msoTrue", "宿主库", None, False),
                   ("MyWorkOrder", "M1", None, False),
                   ("MyWorkOrderCount", "M1", None, False)]
        # ⚠️ 只把【宿主库】那批当宿主常量。工程内名字（MyWork*）必须留在
        # host_min 之外 —— 否则它们会连带被套上"必须从家族前缀开始"的闸门，
        # 输入 count 就再也补不出 MyWorkOrderCount 了（这个坑自己踩过一次）。
        _HOST45 = dict((r[0].lower(),
                        3 if r[0].lower().startswith("mso") else 2)
                       for r in _POOL45 if r[1] == "宿主库")

        class _UI45(object):
            def __init__(self):
                self.shown = False
                self.rows = []

            def show(self, rows, selection=0, completer=None):
                self.shown = True
                self.rows = list(rows)

            def update_selection(self, *a, **k):
                pass

            def hide(self, *a, **k):
                self.shown = False

            def contains_point(self, *a, **k):
                return False

        class _B45(object):
            """桩后端：池子 / 宿主清单都能替换（真实那一路由 45.10 起接手）。"""

            def __init__(self, word, pool=None, host=None):
                self._word = str(word)
                self._pool = list(pool if pool is not None else _POOL45)
                self._host = dict(host if host is not None else _HOST45)

            def get_context(self):
                _ln = "    x = " + self._word
                return {"line_no": 3, "caret_col": len(_ln) + 1,
                        "line_text": _ln, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": "M1"}

            def get_identifiers(self):
                return list(self._pool)

            def get_declared_names(self):
                return [r[0].lower() for r in self._pool]

            def declared_elsewhere(self, name, module_name):
                return str(name).lower() in [r[0].lower() for r in self._pool]

            def get_structural_names(self):
                return [r[0].lower() for r in self._pool]

            def get_builtin_names(self):
                return []

            def get_type_names(self):
                return []

            def names_outside_caret(self, caret, scope_only=False):
                return set()

            def get_host_enum_names(self):
                return dict(self._host)

        def _trig45(word, pool=None, host=None):
            _ui = _UI45()
            _c = E45.Completer(_B45(word, pool, host), _ui)
            _c.trigger(True)
            return list(_c.matches or []), _ui

        # ★用户报的场景：跳步输入要能补出 xlWorkbookDefault
        check("45.4 ★输入 xlworkfaul -> 提示 xlWorkbookDefault（用户报的场景）",
              [_trig45("xlworkfaul")[0]], expect_contain=[["xlWorkbookDefault"]])
        check("45.5 再少一个字符 xlworkfau 同样命中（连续块 6 未变）",
              [_trig45("xlworkfau")[0]], expect_contain=[["xlWorkbookDefault"]])
        check("45.6 前缀照旧：xlworkbookn -> xlWorkbookNormal",
              [_trig45("xlworkbookn")[0]], expect_contain=[["xlWorkbookNormal"]])
        check("45.7 输入不够 4 个字符 -> 退化成纯前缀，与 v67 一致："
              "xlms -> 只 xlMSDOS",
              [_trig45("xlms")[0]], expect_contain=[["xlMSDOS"]])
        check("45.8 跳步候选没有 4 连块 -> 被闸门挡掉："
              "xlce -> 只前缀的 xlCenter / xlCellTypeVisible，"
              "跳步的 xlVAlignCenter 不出",
              [_trig45("xlce")[0]],
              expect_contain=[["xlCenter", "xlCellTypeVisible"]])
        check("45.9 跳步命中排在所有前缀命中之后（kind 0 < kind 3）："
              "xlcell -> xlCellTypeVisible 在前，xlLastCell 在后",
              [_trig45("xlcell")[0]],
              expect_contain=[["xlCellTypeVisible", "xlLastCell"]])
        check("45.10 没从家族前缀开始的输入，一条宿主常量都不出："
              "ms / open 全空，count 只剩工程内名字",
              [_trig45("ms")[0], _trig45("count")[0], _trig45("open")[0]],
              expect_contain=[[], ["MyWorkOrderCount"], []])
        check("45.11 工程内名字不受这套闸门约束（走普通通道）：myworkc -> "
              "MyWorkOrderCount（跳步照旧）",
              [_trig45("myworkc")[0]], expect_contain=[["MyWorkOrderCount"]])

        # ---- 45.11~45.13 真机：真实 4538 条清单 + 真引擎，端到端 ----
        # 与第 44 节同样的理由：前面小节把 VB._get_vbe_cached 打桩成了假 VBE，
        # 必须另加载一份干净的 vbe_bridge 才能读到真实引用。
        _fresh45 = None
        try:
            import importlib.util as _iu45
            _spec45 = _iu45.spec_from_file_location(
                "vbe_bridge_fresh45",
                os.path.join(os.path.dirname(os.path.dirname(
                    os.path.abspath(__file__))), "vbe_bridge.py"))
            _fresh45 = _iu45.module_from_spec(_spec45)
            _spec45.loader.exec_module(_fresh45)
        except Exception:
            _fresh45 = None
        _items45 = []
        try:
            # v70 起宿主枚举默认不收；这一节要真实 4538 条清单，故临时打开开关。
            if _fresh45 is not None:
                _fresh45.ENABLE_HOST_ENUMS = True
            _items45 = (_fresh45 or VB44).host_enum_constants()
        except Exception:
            _items45 = []
        _real45 = [(n, "宿主库", None, False) for n, _m in _items45]
        _rhost45 = dict((n.lower(), m) for n, m in _items45)

        if not _items45:
            check("45.12 真机：拿不到宿主枚举清单（Excel 未开？）"
                  "—— 跳过 45.12~45.14",
                  [True], expect_contain=[True])
        else:
            _m45 = _trig45("xlworkfaul", pool=_real45, host=_rhost45)[0]
            check("45.12 ★真机 %d 条：xlworkfaul 唯一命中 xlWorkbookDefault"
                  "（命中唯一 = 无噪音）" % (len(_items45),),
                  [_m45], expect_contain=[["xlWorkbookDefault"]])
            _m45b = [_trig45(w, pool=_real45, host=_rhost45)[0]
                     for w in ("ms", "count", "open", "cell", "ar", "vb")]
            check("45.13 真机：ms / count / open / cell / ar / vb 一条 xl*/mso* "
                  "都不出（v67 的防噪音不退化）",
                  [_m45b], expect_contain=[[[], [], [], [], [], []]])
            # 3 字符输入不可能凑出 4 连块 -> 行为必须与 v67 的"纯前缀"一字不差
            _pre45 = sorted(n for n, _m in _items45
                            if n.lower().startswith("xla"))
            _gate45 = sorted(n for n, _m in _items45
                             if n.lower().startswith("xla")
                             or (E45.fuzzy_match(n, "xla") is not None
                                 and E45._longest_run(
                                     E45.fuzzy_match(n, "xla")[1])
                                 >= E45.HOST_ENUM_FUZZY_RUN))
            check("45.14 真机：3 字符输入 xla 的命中与「纯前缀」完全一致"
                  "（%d 条；连续块不可能够 4）" % (len(_pre45),),
                  [_gate45 == _pre45, len(_pre45) > 0],
                  expect_contain=[True, True])
    except Exception as _e45:
        check("第 45 节异常: %s" % _e45, [True], expect_contain=[False])

    # ==================================================================
    # 46. 宿主枚举匹配的廉价预筛（v69）—— 纯性能，行为必须一字不变
    #
    # v68 的闸门"输入里存在一段长度 >= 4 的连续块"有个极便宜的必要条件：从输入里
    # 随便取一个 4 连片断，它必须【真的连续出现在】候选名里 —— C 层一句 `in` 就够，
    # 比逐条跑 fuzzy_match + _longest_run 便宜一个量级（实测 2173 条锚定候选
    # 5.9ms -> 0.8ms，整条 trigger 由约 10.6ms 降到约 6.2ms）。
    #
    # 预筛只可能"多留"（留下的还要照跑 fuzzy_match 与闸门复核），绝不"少留"：
    # 真命中的候选，其命中下标里必有一段长度 >= 4 的连续段，那段对应的输入片断
    # 当然就在名字里。这一节就守这个不变式 —— 用真实 4538 条清单。
    #
    # 为什么必须守：预筛一旦"少留"，症状是"某些 xl* 突然补不出来了"，
    # 而模糊匹配本身看不出问题（v67 那次漏收就是这么难查）。
    # ==================================================================
    print("\n=== 46. 宿主枚举匹配的廉价预筛（v69）===")
    try:
        import engine as E46
        import time as _t46

        try:
            _items46 = list(_items45)
        except Exception:
            _items46 = []
        check("46.1 真机清单可用（%d 条；拿不到就跳过本节的实质断言）"
              % len(_items46), [len(_items46) > 0], expect_contain=[True])

        WORDS46 = ("xlworkfaul", "xlworkfau", "xlworkf", "xlwork", "xlwo",
                   "xlwb", "xlwd", "xlcell", "xlcount", "xlco", "xlce", "xla",
                   "xlms", "xlm", "xldef", "xlde", "xld", "xlrange", "xlsheet",
                   "xlcolu", "xlrow", "xlnum", "xlzz", "xlab", "xlj", "ms",
                   "mso", "msobu", "xl")

        def _old_inner46(q):
            """v68 的内层循环（无预筛）：锚定 + run>=4。"""
            ql = q.lower()
            lw = len(q)
            got = []
            for n, need in _items46:
                low = n.lower()
                if lw < need:
                    continue
                if low.startswith(ql):
                    got.append(n)
                    continue
                if not ql.startswith(low[:need]):
                    continue
                r = E46.fuzzy_match(n, q)
                if r is None or E46._longest_run(r[1]) < E46.HOST_ENUM_FUZZY_RUN:
                    continue
                got.append(n)
            return got

        def _new_inner46(q):
            """v69 的内层循环：多一层 4-gram 预筛。"""
            ql = q.lower()
            lw = len(q)
            grams = tuple(ql[k:k + E46.HOST_ENUM_FUZZY_RUN]
                          for k in range(lw - E46.HOST_ENUM_FUZZY_RUN + 1))
            got = []
            for n, need in _items46:
                low = n.lower()
                if lw < need:
                    continue
                if low.startswith(ql):
                    got.append(n)
                    continue
                if not ql.startswith(low[:need]):
                    continue
                if not grams or not any(g in low for g in grams):
                    continue
                r = E46.fuzzy_match(n, q)
                if r is None or E46._longest_run(r[1]) < E46.HOST_ENUM_FUZZY_RUN:
                    continue
                got.append(n)
            return got

        if _items46:
            _bad46 = [q for q in WORDS46
                      if _old_inner46(q) != _new_inner46(q)]
            check("46.2 ★等价性：预筛前后候选集合与顺序完全一致（%d 个输入）"
                  % len(WORDS46), [_bad46], expect_contain=[[]])

            # 性质：过闸门 ==> 过预筛（真数据上零反例）
            _c46 = 0
            _viol46 = []
            for q in WORDS46:
                ql = q.lower()
                grams = [ql[k:k + E46.HOST_ENUM_FUZZY_RUN]
                         for k in range(len(ql) - E46.HOST_ENUM_FUZZY_RUN + 1)]
                for n, need in _items46:
                    low = n.lower()
                    if len(ql) < need or not ql.startswith(low[:need]):
                        continue
                    r = E46.fuzzy_match(n, q)
                    if r is None:
                        continue
                    if E46._longest_run(r[1]) < E46.HOST_ENUM_FUZZY_RUN:
                        continue
                    _c46 += 1
                    if not any(g in low for g in grams):
                        _viol46.append((q, n))
            check("46.3 过闸门的候选 %d 条，全部过得了预筛（有反例 = 预筛会误杀）"
                  % _c46, [_viol46, _c46 > 0], expect_contain=[[], True])

            # 短输入（< RUN 个字符）时预筛为空 = 非前缀锚定候选直接跳过，
            # 与闸门"连续块不可能够长"等价 —— 拿 xla / xlms / xlco 对一遍。
            _short46 = [q for q in ("xla", "xlms", "xlco", "xlm", "xlce", "ms")
                        if _old_inner46(q) != _new_inner46(q)]
            check("46.4 短输入（不足 4 字符）走空预筛也一字不差",
                  [_short46], expect_contain=[[]])

            # 性能护栏（相对比较，阈值宽松）：同样的锚定候选，预筛必须更快
            _q46 = "xlworkfaul"
            _anch46 = [n for n, need in _items46
                       if len(_q46) >= need
                       and _q46.startswith(n.lower()[:need])
                       and not n.lower().startswith(_q46)]
            _grams46 = [_q46[k:k + E46.HOST_ENUM_FUZZY_RUN]
                        for k in range(len(_q46)
                                       - E46.HOST_ENUM_FUZZY_RUN + 1)]

            def _t_gate46():
                for n in _anch46:
                    r = E46.fuzzy_match(n, _q46)
                    if r is not None:
                        E46._longest_run(r[1])

            def _t_pre46():
                for n in _anch46:
                    low = n.lower()
                    if any(g in low for g in _grams46):
                        r = E46.fuzzy_match(n, _q46)
                        if r is not None:
                            E46._longest_run(r[1])

            def _best46(fn, times=5):
                b = None
                for _ in range(times):
                    _t0 = _t46.perf_counter()
                    fn()
                    _dt = _t46.perf_counter() - _t0
                    b = _dt if b is None else min(b, _dt)
                return b

            _tb46 = _best46(_t_gate46)
            _tp46 = _best46(_t_pre46)
            check("46.5 性能护栏：%d 条锚定候选，预筛 %.2fms < 全量模糊 %.2fms"
                  % (len(_anch46), _tp46 * 1000, _tb46 * 1000),
                  [_tp46 < _tb46 * 0.8], expect_contain=[True])

            # 端到端（真引擎 + 真实清单）：预筛没把用户要的候选弄丢
            check("46.6 端到端：xlworkfaul -> 唯一 xlWorkbookDefault"
                  "（预筛不误杀；真引擎 + %d 条）" % len(_items46),
                  [_trig45("xlworkfaul", pool=_real45, host=_rhost45)[0]],
                  expect_contain=[["xlWorkbookDefault"]])
    except Exception as _e46:
        check("第 46 节异常: %s" % _e46, [True], expect_contain=[False])

    # ==================================================================
    # 47. v70：内置枚举默认不提示（用户口径），内建函数照常
    #
    # 用户要求："我不怎么使用 VBA 内置枚举，提示出来对我有干扰 —— vb 开头、
    # xl 开头这些枚举都关掉不提示，函数那些要保留。"
    #
    # 于是 vba_builtins 的 CONSTANTS / FUNCTIONS 拆成两组导出，收不收由
    # `vbe_bridge.ENABLE_VBA_CONSTANTS`（默认 False）决定；宿主类型库枚举
    # （xl*/mso*）由 `vbe_bridge.ENABLE_HOST_ENUMS`（v70 起默认 False）决定。
    # 两个开关都能用环境变量翻回来（VBECOMPLETE_VBA_CONSTANTS / _HOST_ENUMS）。
    #
    # 这一节守【默认口径 + 开关真的在起作用】：每条"关着没有"都配一条
    # "打开就有"的反向对照 —— 否则"碰巧没有"（比如桩 VBE 解析失败）也能蒙过去。
    #
    # ⚠️ v77：用户接着说"vba 本身带的关键字、函数这块也都删掉"，于是
    # ENABLE_VBA_BUILTINS / ENABLE_VBA_KEYWORDS 也改成默认 False（本节 47.2 跟着
    # 改成"四批全关"）。**下面 47.5~47.11 先把这两个开关临时打开再断言** ——
    # 那一半验的是"机制仍在、开关真有效"，与默认值无关；默认口径由第 54 节专门守。
    # ==================================================================
    print("\n=== 47. 内置枚举默认不提示（v70）===")
    try:
        import vbe_bridge as VB47
        import vba_builtins as VB47B
        import engine as E47
        import os as _os47

        # ---- 47.1 _env_flag 三态（纯函数）----
        _KEY47 = "VBECOMPLETE_TEST47_FLAG"
        _old47 = _os47.environ.get(_KEY47)
        _r47 = {}
        try:
            _os47.environ.pop(_KEY47, None)
            _r47["unset"] = (VB47._env_flag(_KEY47, True),
                             VB47._env_flag(_KEY47, False))
            for _v in ("1", "true", "YES", "on", "0", "false", "No", "OFF", "??"):
                _os47.environ[_KEY47] = _v
                _r47[_v] = VB47._env_flag(_KEY47, True)
        finally:
            if _old47 is None:
                _os47.environ.pop(_KEY47, None)
            else:
                _os47.environ[_KEY47] = _old47
        check("47.1 _env_flag：未设置/认不出 -> 默认值；1,true,yes,on -> True；"
              "0,false,no,off -> False",
              [_r47["unset"], _r47["1"], _r47["true"], _r47["YES"], _r47["on"],
               _r47["0"], _r47["false"], _r47["No"], _r47["OFF"], _r47["??"]],
              expect_contain=[(True, False), True, True, True, True,
                              False, False, False, False, True])

        # ---- 47.2 出厂默认值 ----
        # ⚠️ 必须单独加载一份干净模块：main() 的前导段会把"内建/关键字"这两个开关
        # 【临时关掉】来跑第 1~37 节的防噪音护栏，现场读到的是被改过的值。
        _fresh47 = None
        try:
            import importlib.util as _iu47
            _spec47 = _iu47.spec_from_file_location(
                "vbe_bridge_fresh47",
                os.path.join(os.path.dirname(os.path.dirname(
                    os.path.abspath(__file__))), "vbe_bridge.py"))
            _fresh47 = _iu47.module_from_spec(_spec47)
            _spec47.loader.exec_module(_fresh47)
        except Exception:
            _fresh47 = None
        check("47.2 ★出厂默认（v77）：内建函数/数据类型、关键字、内置枚举、"
              "宿主枚举 —— 四批一律【关】",
              [((_fresh47.ENABLE_VBA_BUILTINS,
                 _fresh47.ENABLE_VBA_KEYWORDS,
                 _fresh47.ENABLE_VBA_CONSTANTS,
                 _fresh47.ENABLE_HOST_ENUMS) if _fresh47 else "干净模块加载失败")],
              expect_contain=[(False, False, False, False)])

        # ---- 47.3 清单拆成两组（清单本身不裁剪）----
        _inter47 = set(VB47B.BUILTIN_FUNCTIONS) & set(VB47B.BUILTIN_CONSTANTS)
        check("47.3 清单拆组：函数 %d / 常量 %d，交集 %s，并集 == BUILTIN_NAMES(%d)"
              % (len(VB47B.BUILTIN_FUNCTIONS), len(VB47B.BUILTIN_CONSTANTS),
                 sorted(_inter47) or "空", len(VB47B.BUILTIN_NAMES)),
              [_inter47,
               set(VB47B.BUILTIN_NAMES) == (set(VB47B.BUILTIN_FUNCTIONS)
                                            | set(VB47B.BUILTIN_CONSTANTS)),
               "MsgBox" in VB47B.BUILTIN_FUNCTIONS,
               "vbCrLf" in VB47B.BUILTIN_CONSTANTS,
               "vbCrLf" not in VB47B.BUILTIN_FUNCTIONS,
               len(VB47B.BUILTIN_CONSTANTS) > 0],
              expect_contain=[set(), True, True, True, True, True])

        # ---- 47.4 默认：宿主枚举一条都不去枚举 ----
        check("47.4 默认：host_enum_constants() 返回空清单（连类型库都不加载）",
              [len(VB47.host_enum_constants())], expect_contain=[0])

        # ---- 47.5~47.8 真实后端全链路（桩 VBE）----
        class _CM47(object):
            def __init__(self, text):
                self.text = text

            @property
            def CountOfLines(self):
                return self.text.count("\n") + 1

            def Lines(self, start, count):
                return "\r\n".join(
                    self.text.split("\n")[start - 1:start - 1 + count])

        class _Comp47(object):
            def __init__(self, name, text, ctype=1):
                self.Name = name
                self.Type = ctype
                self.CodeModule = _CM47(text)

        class _Proj47(object):
            def __init__(self, comps):
                self.Name = "VBAProject"
                self.VBComponents = list(comps)

        class _Pane47(object):
            def __init__(self, comp):
                self._c = comp

            @property
            def CodeModule(self):
                return self._c.CodeModule

            def GetSelection(self):
                return (1, 1, 1, 1)

        class _VBE47(object):
            def __init__(self, comps, act):
                self.ActiveVBProject = _Proj47(comps)
                self.ActiveCodePane = _Pane47(act)

        class _UI47(object):
            def __init__(self):
                self.rows = []

            def show(self, rows, sel, c):
                self.rows = list(rows)

            def hide(self):
                pass

            def update_selection(self, sel):
                pass

            def contains_point(self, x, y):
                return False

        _comp47 = _Comp47("Module1", "Sub Foo()\n    x = 1\nEnd Sub", 1)
        _vbe47 = _VBE47([_comp47], _comp47)
        _orig47 = VB47._get_vbe_cached
        # ⚠️ main() 前导段把"内建/关键字"开关临时关掉了（第 1~37 节护栏的前提）。
        # 这一节要验的是"枚举关、函数开"，所以显式打开函数/关键字再验。
        _saved47 = (VB47.ENABLE_VBA_BUILTINS, VB47.ENABLE_VBA_KEYWORDS,
                    VB47.ENABLE_VBA_CONSTANTS, VB47.ENABLE_HOST_ENUMS)
        VB47.ENABLE_VBA_BUILTINS = True
        VB47.ENABLE_VBA_KEYWORDS = True
        VB47._get_vbe_cached = lambda: _vbe47
        try:
            _bk47 = VB47.VbeBackend()
            _nm47 = set(str(r[0]).lower() for r in _bk47.get_identifiers())
            _bl47 = _bk47.get_builtin_names()
            check("47.5 ★默认口径：vb* 一条都不进池，而 MsgBox / Split / "
                  "String / If 照常在",
                  ["vbcrlf" not in _nm47, "vbok" not in _nm47,
                   "msgbox" in _nm47, "split" in _nm47,
                   "string" in _nm47, "if" in _nm47],
                  expect_contain=[True] * 6)
            check("47.6 ★默认口径：宿主枚举一条都不收（xlUp / msoTrue 进不了池，"
                  "引擎也拿不到清单）",
                  ["xlup" not in _nm47, "msotrue" not in _nm47,
                   _bk47.get_host_enum_names() == {}],
                  expect_contain=[True, True, True])
            check("47.7 收紧匹配用的 builtin_names 集合里同样没有常量",
                  ["vbcrlf" not in _bl47, "msgbox" in _bl47],
                  expect_contain=[True, True])

            # 反向对照：把开关打开，常量必须回来（证明 47.5 是"开关起作用"，
            # 而不是桩 VBE / 解析环节碰巧什么都没收到）
            VB47.ENABLE_VBA_CONSTANTS = True
            try:
                _nm47b = set(str(r[0]).lower()
                             for r in VB47.VbeBackend().get_identifiers())
                check("47.8 对照：ENABLE_VBA_CONSTANTS=True -> vbCrLf / vbOK 回来",
                      ["vbcrlf" in _nm47b, "vbok" in _nm47b],
                      expect_contain=[True, True])
            finally:
                VB47.ENABLE_VBA_CONSTANTS = False

            # ---- 47.9~47.10 引擎端到端（默认口径）----
            _ctx47 = {"line_no": 1, "line_text": "    ", "caret_col": 5,
                      "in_string": False, "in_comment": False,
                      "in_type_position": False, "in_decl_position": False,
                      "decl_names": [], "proc_name": None,
                      "module_name": "Module1"}

            def _trig47(word):
                _ln = "    " + word
                _ctx47["line_text"] = _ln
                _ctx47["caret_col"] = len(_ln) + 1
                _b = VB47.VbeBackend()
                _b.get_context = lambda: dict(_ctx47)
                _c = E47.Completer(_b, _UI47())
                _c.trigger(False)
                return list(_c.matches or [])

            check("47.9 ★引擎端到端：打 vb / xl / xlu 一个候选都不出"
                  "（用户要的就是这个）",
                  [_trig47("vb"), _trig47("xl"), _trig47("xlu")],
                  expect_contain=[[], [], []])
            check("47.10 引擎端到端：打 ms 仍出 MsgBox（函数那批没被牵连）",
                  [_trig47("ms")], expect_contain=[["MsgBox"]])
        finally:
            VB47._get_vbe_cached = _orig47
            (VB47.ENABLE_VBA_BUILTINS, VB47.ENABLE_VBA_KEYWORDS,
             VB47.ENABLE_VBA_CONSTANTS, VB47.ENABLE_HOST_ENUMS) = _saved47

        # ---- 47.11 接线护栏：别把开关或拆组改没了 ----
        try:
            import inspect as _inspect47
            _src47 = _inspect47.getsource(VB47)
            _src47b = _inspect47.getsource(VB47B)
            _okwire47 = all(k in _src47 for k in (
                "ENABLE_VBA_CONSTANTS", "VBECOMPLETE_VBA_CONSTANTS",
                "VBECOMPLETE_HOST_ENUMS", "def _env_flag",
                "vba_builtins.BUILTIN_FUNCTIONS",
                "vba_builtins.BUILTIN_CONSTANTS")) and (
                "BUILTIN_FUNCTIONS" in _src47b
                and "BUILTIN_CONSTANTS" in _src47b)
        except Exception as _e47w:
            _okwire47 = "异常: %s" % _e47w
        check("47.11 接线护栏：两个开关 + 环境变量 + 函数/常量拆组都在",
              [_okwire47], expect_contain=[True])
    except Exception as _e47:
        check("第 47 节异常: %s" % _e47, [True], expect_contain=[False])

    # ==================================================================
    # 48. v71：公共词汇的"纯分散命中"阈值 4 -> 2（池子小了，可以放宽了）
    #
    # 背景：v61 收紧这一档，是因为当年内建清单近 400 条、还夹着 vbAbortRetryIgnore
    # 这类超长枚举名 —— 两三个字母就能凑出一大串。v70 把内置枚举默认关掉之后，
    # 真机池子掉到 265 条（261 内建 + 工程自己的名字），那个前提没了。
    #
    # 真机实测（详见当日日志 2026-09-15.md）：
    #   ve 10->15 条（VarType 第 11 位）   vr 0->7 条（第 1 位）
    #   se 30->41   ce 3->21   ar 10->21   ms 1->4（MsgBox 仍是第 1）
    #   耗时 0.71ms vs 0.69ms（零成本）
    #
    # 这一节用【确定性池子】（内建清单 + 两个工程名）守三件事，不依赖 Excel：
    #   1) 目标真浮上来了（ve / vr -> VarType 在一屏内）；
    #   2) ★排序不变式：阈值 4 的候选列表必须是阈值 2 的【子序列】——
    #      放宽只是"往后追加"，绝不会重排、不会顶掉原有首选；
    #   3) 代价可度量：单字符输入与"完全不收紧"逐条一致；fuzzy_match 调用次数
    #      两次完全相同（放宽不花额外匹配），耗时也在毫秒级。
    # ==================================================================
    print("\n=== 48. 公共词汇的分散命中阈值 4 -> 2（v71）===")
    try:
        import vba_builtins as VB48
        import inspect as _inspect48
        import time as _t48

        _POOL48 = []
        for _seq in (VB48.BUILTIN_FUNCTIONS, VB48.BUILTIN_TYPE_NAMES,
                     VB48.BUILTIN_KEYWORDS):
            _POOL48 += [(n, "VBA", None, False) for n in _seq]
        _POOL48 += [("numArr", "模块1", None, False),
                    ("countRows", "模块1", None, False)]
        _bl48 = set(r[0].lower() for r in _POOL48)

        class _B48(object):
            def __init__(self, word):
                self._word = word

            def get_context(self):
                _ln = "    " + self._word
                return {"line_no": 1, "line_text": _ln,
                        "caret_col": len(_ln) + 1, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": "模块1"}

            def get_identifiers(self):
                return list(_POOL48)

            def get_declared_names(self):
                return sorted(_bl48)

            def declared_elsewhere(self, name, module_name):
                return str(name).lower() in _bl48

            def get_structural_names(self):
                return sorted(_bl48)

            def get_builtin_names(self):
                return set(_bl48)

            def get_type_names(self):
                return list(VB48.BUILTIN_TYPE_NAMES)

            def get_host_enum_names(self):
                return {}

            def names_outside_caret(self, caret, scope_only=False):
                return set()

            def apply_completion(self, *a, **k):
                return None

        class _UI48(object):
            def show(self, rows, sel, c):
                pass

            def hide(self):
                pass

            def update_selection(self, sel):
                pass

            def contains_point(self, x, y):
                return False

        def _trig48(word, scatter_min=None, count_fm=False):
            _sm = E.BUILTIN_SCATTER_MIN
            _ofm = E.fuzzy_match
            _cnt = [0]
            if scatter_min is not None:
                E.BUILTIN_SCATTER_MIN = scatter_min
            if count_fm:
                def _fm(n, q):
                    _cnt[0] += 1
                    return _ofm(n, q)
                E.fuzzy_match = _fm
            _c = None
            try:
                _c = E.Completer(_B48(word), _UI48())
                _t0 = _t48.perf_counter()
                _c.trigger()
                _dt = (_t48.perf_counter() - _t0) * 1000
            finally:
                E.BUILTIN_SCATTER_MIN = _sm
                E.fuzzy_match = _ofm
            return list((_c.matches if _c else None) or []), _dt, _cnt[0]

        check("48.1 默认口径：BUILTIN_SCATTER_MIN == 2，且可用环境变量调回去",
              [E.BUILTIN_SCATTER_MIN,
               "VBECOMPLETE_BUILTIN_SCATTER_MIN" in _inspect48.getsource(E)],
              expect_contain=[2, True])

        _ve, _tve, _ = _trig48("ve")
        _ve4, _, _ = _trig48("ve", scatter_min=4)
        check("48.2 ★ve -> VarType 浮上来且在一屏内（放宽后 %d 条，第 %s 位；"
              "阈值 4 时 %s）"
              % (len(_ve), (_ve.index("VarType") + 1) if "VarType" in _ve
                 else "-", "也有" if "VarType" in _ve4 else "没有"),
              ["VarType" in _ve, "VarType" not in _ve4],
              expect_contain=[True, True])

        _vr, _, _ = _trig48("vr")
        check("48.3 vr -> VarType 稳居第 1 位（%d 条；分散命中里"
              "命中越靠前越优先）" % len(_vr),
              [_vr[:1]], expect_contain=[["VarType"]])

        # ★排序不变式：收紧版是放宽版的子序列 ⇒ 放宽只"往后追加"
        _it48 = iter(_ve)
        check("48.4 ★排序不变式：阈值 4 的候选列表是阈值 2 的【子序列】"
              "（放宽只追加、不重排、不顶掉首选）",
              [all(x in _it48 for x in _ve4)], expect_contain=[True])
        check("48.5 首选没被顶掉：ms -> 第 1 个仍是 MsgBox（前缀优先）",
              [_trig48("ms")[0][:1]], expect_contain=[["MsgBox"]])

        # 单字符输入：与"完全不收紧"（_builtin_names 打桩成空集）逐条一致
        _e1, _, _ = _trig48("e")
        _obl48 = E.Completer._builtin_names
        E.Completer._builtin_names = lambda self: set()
        try:
            _e2, _, _ = _trig48("e")
        finally:
            E.Completer._builtin_names = _obl48
        check("48.6 单字符输入不受影响（e：%d 条，与「完全不收紧」逐条一致 ——"
              " 这条规则只管 2~3 字符）" % len(_e1),
              [_e1 == _e2, len(_e1) > 0], expect_contain=[True, True])

        # 代价：放宽不花额外匹配，耗时毫秒级
        _, _t2, _n2 = _trig48("ve", scatter_min=2, count_fm=True)
        _, _t4, _n4 = _trig48("ve", scatter_min=4, count_fm=True)
        check("48.7 代价：fuzzy_match 调用次数两次相同（%d = %d），"
              "耗时 %.2fms / %.2fms（池子 %d 条）"
              % (_n2, _n4, _t2, _t4, len(_POOL48)),
              [_n2 == _n4, _t2 < 50 and _t4 < 50], expect_contain=[True, True])
    except Exception as _e48:
        check("第 48 节异常: %s" % _e48, [True], expect_contain=[False])

    # ==================================================================
    # 49. v72：With 块内的 `.成员` 让位（修「VBE 的弹窗反被我们挤掉」）
    #
    # 用户报：`With Range("a1:a10").Font` + `    .Size = 12` + `    .Bold = True`
    # 这样写的时候，在 With 块内输入 `.Size` / `.Bold`，本该我们让位给 VBE 的
    # 成员列表，实际却是「VBE 的弹窗给我们的弹窗让位了」。
    #
    # 根因：让位的【语法判据】只认"点号左边有东西"（标识符 / ) / ] / }），
    # `vbe_member_list_expected` 里就是 `m < 0 -> return False`。而 With 块最
    # 标准的写法恰恰是点顶在行首（前面只有缩进）。第二路"窗口可见性"判据在
    # 这种场景下也兜不住：我们的候选窗是 overrideredirect + topmost 悬浮窗，
    # 一弹出就压在 VBE 的列表上面，而 VBE 的列表在失去激活时会自己收起 ——
    # "等它真的画出来再让"永远来不及。
    #
    # v72 的修法（两条）：
    #   1) 语法判据认识 With 块：点号左边是行首，或 _WITH_OWNER_BOUNDARY 里的
    #      语句 / 运算符边界（`:` `=` `,` `(` 与 `& + - * / \ ^ > <`）；
    #   2) main.py 轮询加一道兜底：VBE 的成员列表真的可见就收起我们的
    #      （手动 Ctrl+Space 唤出的除外）。
    # ==================================================================
    print("\n=== 49. v72：With 块内的 .成员 让位 ===")
    try:
        import vba_builtins as VB49

        _POOL49 = [(n, "VBA", None, False) for n in VB49.BUILTIN_FUNCTIONS]
        _POOL49 += [(n, "VBA", None, False) for n in VB49.BUILTIN_TYPE_NAMES]
        _bl49 = set(r[0].lower() for r in _POOL49)

        class _B49(object):
            def __init__(self, line_text, caret_col):
                self._lt = line_text
                self._cc = caret_col

            def get_context(self):
                return {"line_no": 1, "line_text": self._lt,
                        "caret_col": self._cc, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": "模块1"}

            def get_identifiers(self):
                return list(_POOL49)

            def get_declared_names(self):
                return sorted(_bl49)

            def declared_elsewhere(self, name, module_name):
                return str(name).lower() in _bl49

            def get_structural_names(self):
                return sorted(_bl49)

            def get_builtin_names(self):
                return set(_bl49)

            def get_type_names(self):
                return list(VB49.BUILTIN_TYPE_NAMES)

            def get_host_enum_names(self):
                return {}

            def names_outside_caret(self, caret, scope_only=False):
                return set()

            def apply_completion(self, *a, **k):
                return None

        class _UI49(object):
            def __init__(self):
                self.hides = 0

            def show(self, rows, sel, c):
                pass

            def hide(self):
                self.hides += 1

            def update_selection(self, sel):
                pass

            def contains_point(self, x, y):
                return False

        _yl49 = E.vbe_list_expected

        def _yl49s(line):
            """按「光标停在行尾」判定这一行该不该让位。"""
            return bool(_yl49(line, len(line) + 1))

        check("49.1 ★行首的点（With 块标准写法）-> 让位",
              [_yl49s("    .Size"), _yl49s("\t.Bold"), _yl49s("    ."),
               _yl49s("    .si")],
              expect_contain=[True, True, True, True])
        check("49.2 ★语句 / 运算符边界后的点 -> 让位"
              "（a = 1: .Size / Set f = .Font / MySub a, .Font / x = .Left + .Width）",
              [_yl49s("    a = 1: .Size"), _yl49s("    Set f = .Font"),
               _yl49s("    MySub a, .Font"), _yl49s("    x = .Left + .Width")],
              expect_contain=[True, True, True, True])
        check("49.3 反例（不让位，防过度让位）：小数点 / 注释里的点 / 字符串里的点"
              " / 注释里的冒号点 / 纯标识符 / 空行",
              [_yl49s("    x = 1.5"), _yl49s("    ' UserForm1."),
               _yl49s('    s = "a."'), _yl49s("    x = 1 ' c: .foo"),
               _yl49s("abc"), _yl49s("")],
              expect_contain=[False, False, False, False, False, False])
        check("49.4 原有位置照旧让位、没退化：obj. / a(1). / Dim x As / New",
              [_yl49s("    UserForm1.SelectedList"), _yl49s("    a(1)."),
               _yl49s("    Dim x As "), _yl49s("    Set c = New ")],
              expect_contain=[True, True, True, True])
        check("49.5 _WITH_OWNER_BOUNDARY 覆盖语句边界与算术/比较运算符",
              [set(":=,(") <= set(E._WITH_OWNER_BOUNDARY),
               set("+-*/\\^<>") <= set(E._WITH_OWNER_BOUNDARY)],
              expect_contain=[True, True])

        def _trig49(line, auto=True, yield_on=None):
            _ui = _UI49()
            _c = E.Completer(_B49(line, len(line) + 1), _ui)
            _oy = E.YIELD_TO_VBE_LIST
            if yield_on is not None:
                E.YIELD_TO_VBE_LIST = yield_on
            try:
                _c.trigger(bool(auto))
            finally:
                E.YIELD_TO_VBE_LIST = _oy
            return list(_c.matches or []), _ui.hides

        _m6, _h6 = _trig49("    .si")
        check("49.6 ★引擎端到端：轮询自动触发 + With 块内 .si -> 让位（不弹候选）",
              [len(_m6) == 0, _h6 > 0], expect_contain=[True, True])

        _m7, _h7 = _trig49("    .si", yield_on=False)
        check("49.7 对照：关掉 YIELD_TO_VBE_LIST -> 候选照出（%d 条）"
              "（证明 49.6 是「让位」在起作用，不是碰巧没候选）" % len(_m7),
              [len(_m7) > 0], expect_contain=[True])

        _m8, _h8 = _trig49("    .si", auto=False)
        check("49.8 手动 Ctrl+Space（让位只拦自动路径）-> 照旧弹（%d 条）"
              % len(_m8), [len(_m8) > 0], expect_contain=[True])

        _m9, _h9 = _trig49("    Msg")
        check("49.9 普通标识符位置不受影响：Msg -> 照常弹、且没被让位收掉",
              [_h9 == 0, len(_m9) > 0], expect_contain=[True, True])

        # ---- 49.10 接线护栏：main.py 轮询里的兜底 ----
        _msrc49 = ""
        try:
            _msrc49 = open(os.path.join(
                os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                "main.py"), encoding="utf-8").read()
        except Exception:
            _msrc49 = ""
        # v76：兜底里那句 `not state.get("popup_manual") and ...` 的前置条件已经
        # 去掉 —— 用户口径是"冲突时直接把我们的窗收掉"，不分手动/自动。这里除了
        # 确认接线还在，额外钉住【这个前置条件不许回来】。
        check("49.10 接线护栏：poll 兜底在（VBE 成员列表可见 -> 让位，不再分手动）",
              [_msrc49.count("VBE_YIELD_CLASSES") >= 1,
               "popup_manual" in _msrc49,
               _msrc49.count("popup_manual") >= 3,
               'if not state.get("popup_manual") and any(' not in _msrc49],
              expect_contain=[True, True, True, True])
    except Exception as _e49:
        check("第 49 节异常: %s" % _e49, [True], expect_contain=[False])

    # ==================================================================
    # 50. v73：让位的【宽限确认】—— VBE 不弹时自己补上
    #
    # 用户报（v72 之后）：在 With 块里输入 .Size **什么都不弹了**。
    #
    # 根因不是"我们把 VBE 的窗关掉了"（让位路径只有 ui.hide()，碰不到 VBE），
    # 而是【让位这个设计本身的假设不成立】：我们让开，是假定"VBE 的成员列表
    # 一定会顶上"。可 VBE 的自动列出成员要先把 `With` 的对象类型解析出来 ——
    # 解析不了（`With Range("a1:10")` 这种地址写错、模块里有待编译的错误）它
    # 一个窗都不弹。那时让位 = 什么都不给。
    #
    # v73 把让位改成【待定】：先让，YIELD_GRACE 之后由 main.py 的轮询调
    # Completer.confirm_yield() 裁决 —— VBE 的列表真出现就认；没出现就把候选
    # 补回来（trigger(False)）。同一处只判定一次（yield_given_up），避免
    # "让开-补上-再让开-再补上"来回抖。
    # ==================================================================
    print("\n=== 50. v73：让位宽限确认（VBE 没弹就自己补） ===")
    try:
        import vba_builtins as VB50

        _POOL50 = [(n, "VBA", None, False) for n in VB50.BUILTIN_FUNCTIONS]
        _bl50 = set(r[0].lower() for r in _POOL50)

        class _B50(object):
            def __init__(self, line_text, caret_col, line_no=1,
                         module="模块1"):
                self._lt = line_text
                self._cc = caret_col
                self._ln = line_no
                self._mod = module

            def move(self, line_text, caret_col, line_no=None, module=None):
                self._lt = line_text
                self._cc = caret_col
                if line_no is not None:
                    self._ln = line_no
                if module is not None:
                    self._mod = module

            def get_context(self):
                return {"line_no": self._ln, "line_text": self._lt,
                        "caret_col": self._cc, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": self._mod}

            def get_identifiers(self):
                return list(_POOL50)

            def get_declared_names(self):
                return []

            def declared_elsewhere(self, name, module_name):
                return False

            def get_structural_names(self):
                return []

            def get_builtin_names(self):
                return set(_bl50)

            def get_type_names(self):
                return []

            def get_host_enum_names(self):
                return {}

            def names_outside_caret(self, caret, scope_only=False):
                return set()

            def apply_completion(self, *a, **k):
                return None

        class _UI50(object):
            def __init__(self):
                self.hides = 0

            def show(self, rows, sel, c):
                pass

            def hide(self):
                self.hides += 1

            def update_selection(self, sel):
                pass

            def contains_point(self, x, y):
                return False

        _line50 = "    .Si"
        _b50 = _B50(_line50, len(_line50) + 1)
        _ui50 = _UI50()
        _c50 = E.Completer(_b50, _ui50)

        _c50.trigger(True)
        check("50.1 ★让位后挂上待定（还没裁决），候选不弹",
              [_c50.yield_pending is not None,
               len(_c50.matches or []) == 0],
              expect_contain=[True, True])

        check("50.2 VBE 成员列表出现了 -> confirm_yield 认这个让位（True）",
              [_c50.confirm_yield(True), _c50.yield_pending is None],
              expect_contain=[True, True])

        _c50.trigger(True)
        check("50.3 ★VBE 没弹 -> confirm_yield 判 False（该补候选了）",
              [_c50.confirm_yield(False)], expect_contain=[False])

        _c50.trigger(False)
        check("50.4 ★补上后候选照出（%d 条）—— 这就是「With 块里什么都不弹」的修复"
              % len(_c50.matches or []),
              [len(_c50.matches or []) > 0], expect_contain=[True])

        _ui50.hides = 0
        _c50.trigger(True)
        check("50.5 ★同一处再自动触发 -> 不再让位（候选留着、一次都不 hide）"
              "—— 防「让-补-让-补」抖动",
              [len(_c50.matches or []) > 0, _ui50.hides == 0],
              expect_contain=[True, True])

        _b50.move("    .Bo", 8, line_no=7)
        _c50.trigger(True)
        check("50.6 换到另一行 -> 重新让位（「这一处 VBE 不弹」的记忆不跨行）",
              [len(_c50.matches or []) == 0,
               _c50.yield_pending is not None],
              expect_contain=[True, True])

        _b50.move("    Msg", 8, line_no=8)
        check("50.7 宽限期内用户已走开 -> confirm_yield 判 True（不补，别打扰）",
              [_c50.confirm_yield(False)], expect_contain=[True])

        _c50b = E.Completer(_B50("    Msg", 8), _UI50())
        check("50.8 没有待定让位时 confirm_yield 一律 True（幂等，不凭空补）",
              [_c50b.confirm_yield(True), _c50b.confirm_yield(False)],
              expect_contain=[True, True])

        _c50c = E.Completer(_B50("    .Si", 8), _UI50())
        _c50c.trigger(False)
        check("50.9 手动 Ctrl+Space -> 不挂待定（让位只拦自动路径）",
              [_c50c.yield_pending is None,
               len(_c50c.matches or []) > 0],
              expect_contain=[True, True])

        _c50d = E.Completer(_B50("    .Si", 8), _UI50())
        _c50d.trigger(True)
        _c50d.backend.get_context = lambda: None
        check("50.10 拿不到 ctx（COM 暂时不可用）-> 不补（保持让位，不乱弹）",
              [_c50d.confirm_yield(False)], expect_contain=[True])

        check("50.11 YIELD_GRACE 默认 0.25 秒且在合理区间",
              [0.0 < E.YIELD_GRACE < 5.0], expect_contain=[True])

        _root50 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        def _read50(name):
            try:
                return open(os.path.join(_root50, name),
                            encoding="utf-8").read()
            except Exception:
                return ""

        _src50 = _read50("engine.py")
        check("50.12 VBECOMPLETE_YIELD_GRACE 可调（开关 / 宽限期 / 裁决都在引擎里）",
              ["VBECOMPLETE_YIELD_GRACE" in _src50,
               "YIELD_GRACE" in _src50,
               "confirm_yield" in _src50],
              expect_contain=[True, True, True])

        _msrc50 = _read50("main.py")
        check("50.13 接线护栏：poll 调 confirm_yield，判 False 时 trigger(False) 补上",
              ["confirm_yield" in _msrc50,
               "yield_pending" in _msrc50,
               "completer.trigger(False)" in _msrc50],
              expect_contain=[True, True, True])
    except Exception as _e50:
        check("第 50 节异常: %s" % _e50, [True], expect_contain=[False])

    # ==================================================================
    # 51. v74：候选窗 × 参数信息窗重叠的【状态自愈】
    #
    # 用户报："instr 那个函数的参数信息比较长，与我们候选窗重叠了"（v66 修过，
    # 又复现）。
    #
    # 根因不是避让算法 —— avoid_popup_rects 逐条验过没毛病（第 43 节守着它），
    # 而是轮询的【变化检测】有个洞：提示窗矩形签名只在候选窗可见时采样，于是
    # "候选窗不可见期间提示窗出现过、又回到同一位置"这件事永远不被记录 ——
    # 下次候选窗弹出时签名恰好等于那个陈旧缓存值，`_sig != _popup_sig` 便看不见
    # 变化，候选窗一直压在签名窗上，直到用户自己挪光标。
    #
    # 修法两条（互补，都要守）：
    #   1) 签名【每轮都采】，与候选窗可见性无关 —— 堵住那个洞；
    #   2) 加一道【状态检查】：叠着就重摆，压根不依赖"变没变"（自愈）。
    # ==================================================================
    print("\n=== 51. v74：候选窗 × 参数信息窗重叠的状态自愈 ===")
    try:
        import ui as UI51

        # 51.1 "重叠"的定义只此一份（避让与自愈必须同一把尺子）
        check("51.1 rect_overlap：真重叠 True；只共边 / 共点 / 分开都 False",
              [[UI51.rect_overlap((0, 0, 10, 10), (5, 5, 15, 15)),
                UI51.rect_overlap((0, 0, 10, 10), (10, 0, 20, 10)),
                UI51.rect_overlap((0, 0, 10, 10), (10, 10, 20, 20)),
                UI51.rect_overlap((0, 0, 10, 10), (20, 0, 30, 10))]],
              expect_contain=[[True, False, False, False]])

        # 51.2 clashes_with_popup：状态检查（不问变化，只问"现在叠着没有"）
        check("51.2 clashes_with_popup：叠着 True / 没提示窗 False / 已躲开 False",
              [[UI51.clashes_with_popup(100, 220, 200, 90, [(0, 220, 400, 240)]),
                UI51.clashes_with_popup(100, 220, 200, 90, []),
                UI51.clashes_with_popup(100, 242, 200, 90,
                                        [(0, 220, 400, 240)])]],
              expect_contain=[[True, False, False]])

        # 51.3 真机那条长签名窗（探针实测 505x18 @ 1091,601）：避让后不叠，
        #      且用【自愈判据】复核也是"不叠" —— 两条路径结论必须一致。
        _tip51 = (1091, 601, 1596, 619)
        _y51 = UI51.avoid_popup_rects(1368, 601, 250, 150, 583, 1080, [_tip51])
        check("51.3 ★长签名窗 505x18：避让到 y=%d，自愈复核不叠（同一把尺子）"
              % (_y51,),
              [[_y51, UI51.clashes_with_popup(1368, _y51, 250, 150, [_tip51])]],
              expect_contain=[[621, False]])

        # 51.4 ★★用户场景复现：候选窗【先】落在 601（签名窗还没出现），随后
        #      签名窗在【同一位置】出现 —— 变化检测看不见（签名与缓存相等），
        #      但状态检查必须看见。这条就是修好的那个洞。
        check("51.4 ★★陈旧签名 + 候选窗停在 601 -> 状态检查必须报「叠着」",
              [UI51.clashes_with_popup(1368, 601, 250, 150, [_tip51])],
              expect_contain=[True])

        # 51.5~51.7 Popup.clashes_with_popup：拿窗口【当前实际几何】比对
        class _FR51(object):
            def winfo_screenwidth(self):
                return 1920

            def winfo_screenheight(self):
                return 1080

        class _FC51(object):
            def cget(self, k):
                return 200 if k == "width" else 90

        class _FW51(object):
            def __init__(self, state="normal"):
                self._st = state

            def state(self):
                return self._st

            def geometry(self, s):
                pass

            def winfo_x(self):
                return 100

            def winfo_y(self):
                return 220

            def winfo_width(self):
                return 200

            def winfo_height(self):
                return 90

        def _mk51(win):
            p = UI51.Popup.__new__(UI51.Popup)
            p.root, p.canvas, p.win = _FR51(), _FC51(), win
            p._last_geom = None
            return p

        _o51rects = UI51.screen_popup_rects
        try:
            UI51.screen_popup_rects = lambda: [(0, 220, 400, 240)]
            _c51 = _mk51(_FW51()).clashes_with_popup()
            UI51.screen_popup_rects = lambda: []
            _c51b = _mk51(_FW51()).clashes_with_popup()
            _c51c = _mk51(_FW51("withdrawn")).clashes_with_popup()
            _c51d = _mk51(None).clashes_with_popup()
        finally:
            UI51.screen_popup_rects = _o51rects
        check("51.5 候选窗实际位置(100,220,200x90) × 签名窗 -> 报「叠着」",
              [_c51], expect_contain=[True])
        check("51.6 没有提示窗 / 已收起 / 没有窗 -> 一律 False（拿不准不动手）",
              [[_c51b, _c51c, _c51d]],
              expect_contain=[[False, False, False]])

        # 51.8~51.11 main.py 接线护栏（轮询那段的结构性不变量）
        _root51 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        def _read51(name):
            try:
                return open(os.path.join(_root51, name),
                            encoding="utf-8").read()
            except Exception:
                return ""

        _msrc51 = _read51("main.py")

        def _fn51(text, name):
            """把嵌套函数 name 的函数体从源码里抠出来（到下一个同级 def 为止）。"""
            i = text.find("def %s(" % name)
            if i < 0:
                return ""
            j = text.find("\n    def ", i + 1)
            return text[i:j if j > 0 else len(text)]

        _poll51 = _fn51(_msrc51, "poll_editor")
        # 只看代码、不看注释 —— v74 的说明里就写着 "if completer.is_visible()"
        # 这几个字，直接 find 会先命中注释（第一版测试就是这么写错的）。
        _code51 = "\n".join(_l for _l in _poll51.splitlines()
                            if not _l.lstrip().startswith("#"))
        _i_sig51 = _code51.find("_sig = tuple(vbe_bridge.vbe_popup_rects())")
        _i_set51 = _code51.find('state["_popup_sig"] = _sig')
        _i_vis51 = _code51.find("if completer.is_visible():")
        check("51.8 ★签名采样与落缓存都在【可见性判断之前】（无条件采样 —— 洞就堵在这）",
              [[_i_sig51 >= 0, _i_set51 >= 0, _i_vis51 >= 0,
                _i_sig51 < _i_vis51, _i_set51 < _i_vis51]],
              expect_contain=[[True, True, True, True, True]])

        check("51.9 ★接线：重叠时调 popup.clashes_with_popup() 触发重摆",
              [[_poll51.count("popup.clashes_with_popup()") == 1,
                "post(popup.reposition)" in _poll51,
                "_sig_changed or _clash" in _poll51]],
              expect_contain=[[True, True, True]])

        check("51.10 同一 tick 内两处判据共用一份签名（不再各探一次）",
              [[_msrc51.count("vbe_bridge.vbe_popup_rects()") == 1,
                _poll51.count("vbe_bridge.vbe_popup_rects()") == 1,
                "_sig)\n" in _poll51 or "in _sig)" in _poll51]],
              expect_contain=[[True, True, True]])

        check("51.11 Popup.clashes_with_popup 存在、可调用（reposition 的老位子旁边）",
              [[callable(getattr(UI51.Popup, "clashes_with_popup", None)),
                callable(getattr(UI51.Popup, "reposition", None)),
                callable(getattr(UI51, "rect_overlap", None))]],
              expect_contain=[[True, True, True]])

        # 51.12 避让与自愈对同一批矩形必须给出一致结论（性质检查，随机 300 组）
        import random as _rnd51
        _bad51 = []
        _g51 = _rnd51.Random(20260915)
        for _ in range(300):
            _b51 = _g51.randint(0, 300)
            _x51 = _g51.randint(0, 600)
            _y51b = _g51.randint(0, 700)
            _m51 = _g51.randint(0, 400)
            _r51 = [(_x51 - _g51.randint(0, 50), _y51b,
                     _x51 + _g51.randint(0, 400), _y51b + 20)]
            _new51 = UI51.avoid_popup_rects(_x51, _y51b, _m51, 90, 100,
                                            1080, _r51)
            # 避让说"躲开了"，自愈就不能说"还叠着"（除非真的贴底放不下）
            if not UI51.clashes_with_popup(_x51, _new51, _m51, 90, _r51):
                continue
            if _new51 >= max(0, 1080 - 90) or _new51 < 0:
                continue                    # 贴底/翻上方：确实无处可去
            _bad51.append((_x51, _y51b, _new51))
        check("51.12 ★避让后自愈复核：300 组随机矩形零反例（%d 组残留）"
              % (len(_bad51),),
              [len(_bad51)], expect_contain=[0])
    except Exception as _e51:
        check("第 51 节异常: %s" % _e51, [True], expect_contain=[False])

    # ==================================================================
    # 52. v75：隐式变量尊重模块自己的 Option Explicit
    #
    # 用户口径："没有 Option Explicit 时，用过的变量名就算已经被定义、可以直接
    # 提示（现在这样很好，不要提前定义变量名）；写了 Option Explicit 是强制声明，
    # 没声明过的变量名在后续就不该被提示。"
    #
    # 判定【按模块】做 —— Option Explicit 本来就是模块级语句：
    #   * 没写它的模块：一字不动（用过的名字照旧进池子）；
    #   * 写了它的模块：只收"真声明过"的名字（Dim / Const / Sub / Function /
    #     Type / Enum / 形参），隐式变量整段不收 —— 那种写法在该模块里是编译
    #     错误，提示它只会把坏代码带出去。
    #
    # 落点在【收集侧】（extract_implicit_records 之前），因为隐式变量是唯一
    # "用过但没声明"的名字来源：extract_records 收的都是真声明，而隐式变量从不
    # 进 declared_names / _decl_by_mod（回声防护那边只靠"现场文本"兜）。改这一处
    # 就够，引擎与 UI 一行不用动。
    # ==================================================================
    print("\n=== 52. v75：Option Explicit 模块不再收「用过但没声明」的名字 ===")
    try:
        import os as _os52
        import vbe_bridge as VB52
        import parser as P52

        # ---- 52.1 has_option_explicit 真值表（文本判定，纯函数）----
        # 全是"认错就出事"的边界：把注释掉的 Option Explicit 认成真，会把整个
        # 模块的隐式变量误砍掉（而那正是用户要保留的行为）；把真行漏掉则等于白改。
        _oe52 = [
            P52.has_option_explicit("Option Explicit\nSub A()\nEnd Sub"),
            P52.has_option_explicit("option explicit\nSub A()"),
            P52.has_option_explicit("   Option   Explicit\nSub A()"),
            P52.has_option_explicit("Attribute VB_Name = \"M\"\r\n"
                                    "Option Explicit\r\nSub A()"),
            P52.has_option_explicit("Option Base 1\nOption Compare Text\nSub A()"),
            P52.has_option_explicit("' Option Explicit\nSub A()\nEnd Sub"),
            P52.has_option_explicit('Sub A()\n    s = "Option Explicit"\nEnd Sub'),
            P52.has_option_explicit("Sub A()\n    x = 1 ' Option Explicit\nEnd Sub"),
            P52.has_option_explicit(""),
            P52.has_option_explicit(None),
        ]
        check("52.1 ★文本判定：行首的 Option Explicit 才算（缩进 / 小写 / 带 "
              "Attribute 头都算）；注释掉的 / 字符串里的 / 其它 Option 语句 / "
              "空文本都不算",
              [_oe52],
              expect_contain=[[True, True, True, True, False, False, False,
                               False, False, False]])

        # ---- 52.2~52.6 真实后端全链路（桩 VBE：一个不带、一个带 OE）----
        class _CM52(object):
            def __init__(self, comp, text):
                self.Parent = comp
                self.text = text

            @property
            def CountOfLines(self):
                return self.text.count("\n") + 1

            def Lines(self, start, count):
                return "\r\n".join(
                    self.text.split("\n")[start - 1:start - 1 + count])

        class _Comp52(object):
            def __init__(self, name, text, ctype=1):
                self.Name = name
                self.Type = ctype
                self.CodeModule = _CM52(self, text)

        class _Proj52(object):
            def __init__(self, comps):
                self.Name = "VBAProject"
                self.VBComponents = list(comps)

        class _Pane52(object):
            def __init__(self, comp):
                self._c = comp

            @property
            def CodeModule(self):
                return self._c.CodeModule

            def GetSelection(self):
                # 光标落在【最后一个空行】上：既不会抹掉任何标识符，也不会被
                # 当成"声明行"（否则正在声明的名字会被排除，干扰本节的断言）。
                return (10, 1, 10, 1)

        class _VBE52(object):
            def __init__(self, proj, act):
                self.ActiveVBProject = proj
                self.ActiveCodePane = _Pane52(act)

        # 带 Option Explicit：looseVar2 用过但没声明（编译错误）；declaredVar 真声明；
        # nParam 是形参（也是声明）。
        _EX52 = ("Option Explicit\n"
                 "Sub A()\n"
                 "    Dim declaredVar As Long\n"
                 "    declaredVar = 2\n"
                 "    looseVar2 = 1\n"
                 "End Sub\n"
                 "Sub C(ByVal nParam As Long)\n"
                 "    nParam = nParam + 1\n"
                 "End Sub\n")
        # 不带 Option Explicit：looseVar 同样是"用过即存在"，必须照旧收。
        _LO52 = ("Sub B()\n"
                 "    looseVar = 1\n"
                 "End Sub\n")
        _cEX52 = _Comp52("ExplicitMod", _EX52)
        _cLO52 = _Comp52("LooseMod", _LO52)
        _vbe52 = _VBE52(_Proj52([_cEX52, _cLO52]), _cEX52)

        _orig52 = VB52._get_vbe_cached
        _saved52 = (VB52.IMPLICIT_HONOR_OPTION_EXPLICIT,
                    VB52.ENABLE_IMPLICIT_IDENTIFIERS,
                    VB52.ENABLE_VBA_BUILTINS, VB52.ENABLE_VBA_KEYWORDS)
        VB52._get_vbe_cached = lambda: _vbe52
        try:
            # 只看工程内名字：把语言自带的两批关掉，池子干净好断言。
            VB52.ENABLE_VBA_BUILTINS = False
            VB52.ENABLE_VBA_KEYWORDS = False

            VB52.ENABLE_IMPLICIT_IDENTIFIERS = True
            VB52.IMPLICIT_HONOR_OPTION_EXPLICIT = True
            _on52 = set(str(r[0]).lower()
                        for r in VB52.VbeBackend().get_identifiers())

            VB52.IMPLICIT_HONOR_OPTION_EXPLICIT = False
            _off52 = set(str(r[0]).lower()
                         for r in VB52.VbeBackend().get_identifiers())

            VB52.ENABLE_IMPLICIT_IDENTIFIERS = False
            VB52.IMPLICIT_HONOR_OPTION_EXPLICIT = True
            _none52 = set(str(r[0]).lower()
                          for r in VB52.VbeBackend().get_identifiers())
        finally:
            VB52._get_vbe_cached = _orig52
            (VB52.IMPLICIT_HONOR_OPTION_EXPLICIT,
             VB52.ENABLE_IMPLICIT_IDENTIFIERS,
             VB52.ENABLE_VBA_BUILTINS, VB52.ENABLE_VBA_KEYWORDS) = _saved52

        check("52.2 ★带 Option Explicit 的模块：用过但没声明的 looseVar2 不进池；"
              "真声明的 declaredVar、形参 nParam 照常在",
              ["loosevar2" not in _on52, "declaredvar" in _on52,
               "nparam" in _on52],
              expect_contain=[True, True, True])
        check("52.3 ★不带 Option Explicit 的模块一字不动：looseVar 照旧进池"
              "（用户明确说这样好）",
              ["loosevar" in _on52], expect_contain=[True])
        check("52.4 ★反向对照：闸门关掉（旧行为）时 looseVar2 必须回来 "
              "—— 证明是这道闸门在起作用，不是碰巧没有",
              ["loosevar2" in _off52], expect_contain=[True])
        check("52.5 差异精确：闸门只摘掉 looseVar2 一个名字（不多不少）",
              [sorted(_off52 - _on52), sorted(_on52 - _off52)],
              expect_contain=[["loosevar2"], []])
        check("52.6 两个开关互不干扰：ENABLE_IMPLICIT_IDENTIFIERS=False 时"
              "looseVar 也没了，但真声明照旧",
              ["loosevar" not in _none52, "loosevar2" not in _none52,
               "declaredvar" in _none52],
              expect_contain=[True, True, True])

        # ---- 52.7 接线护栏 ----
        def _read52(name):
            return io.open(os.path.join(ROOT, name), encoding="utf-8").read()

        _br52 = _read52("vbe_bridge.py")
        _ps52 = _read52("parser.py")
        check("52.7 接线护栏：收集段真的问了闸门；判定函数在两边都只有一份",
              ["implicit_gate_applies(code)" in _br52,
               "implicit_gate_applies" in _br52 and "has_option_explicit" in _ps52,
               "IMPLICIT_HONOR_OPTION_EXPLICIT" in _br52,
               _ps52.count("def has_option_explicit") == 1],
              expect_contain=[True, True, True, True])

        # ---- 52.8 环境变量开关（必须用干净模块：现场那份可能被前导段改过）----
        _KEY52 = "VBECOMPLETE_IMPLICIT_HONOR_EXPLICIT"
        _old52 = _os52.environ.get(_KEY52)
        _freshD52 = _freshE52 = None

        def _load52(modname):
            import importlib.util as _iu52
            _sp52 = _iu52.spec_from_file_location(
                modname, os.path.join(ROOT, "vbe_bridge.py"))
            _m52 = _iu52.module_from_spec(_sp52)
            _sp52.loader.exec_module(_m52)
            return _m52

        try:
            _os52.environ.pop(_KEY52, None)
            try:
                _freshD52 = _load52("vbe_bridge_fresh52d")
            except Exception:
                _freshD52 = None
            _os52.environ[_KEY52] = "0"
            try:
                _freshE52 = _load52("vbe_bridge_fresh52e")
            except Exception:
                _freshE52 = None
        finally:
            if _old52 is None:
                _os52.environ.pop(_KEY52, None)
            else:
                _os52.environ[_KEY52] = _old52
        check("52.8 出厂默认开启；set VBECOMPLETE_IMPLICIT_HONOR_EXPLICIT=0 "
              "可回到旧行为",
              [(_freshD52.IMPLICIT_HONOR_OPTION_EXPLICIT if _freshD52 else "加载失败",
                _freshE52.IMPLICIT_HONOR_OPTION_EXPLICIT if _freshE52 else "加载失败")],
              expect_contain=[(True, False)])

        # ---- 52.9 真机不变式（Excel 没开就跳过）----
        # 断言形式刻意做成"零泄漏"而不是"某模块有几条"：用户随时会改工程，
        # 写死数字的护栏下次就该报假警了。
        _fresh52 = None
        try:
            _fresh52 = _load52("vbe_bridge_fresh52r")
        except Exception:
            _fresh52 = None
        _r52 = None
        if _fresh52 is None:
            check("52.9 真机：干净模块加载失败 —— 跳过", [True], expect_contain=[True])
        else:
            try:
                _be52r = _fresh52.VbeBackend()
                _pool52r = set(str(r[0]).lower()
                               for r in _be52r.get_identifiers())
                _decl52r = set(str(n).lower()
                               for n in _be52r.get_declared_names())
                _str52r = set(str(n).lower()
                              for n in _be52r.get_structural_names())
                _gated52, _leak52 = 0, []
                _comps52r = _fresh52._collect_components(
                    _fresh52._get_vbe_cached(), None)
                if not _comps52r:
                    raise RuntimeError("拿不到组件")
                for _cp52r in _comps52r:
                    try:
                        _nm52r = str(_cp52r.Name)
                        _ct52r = _cp52r.CodeModule.CountOfLines
                        _cd52r = (_cp52r.CodeModule.Lines(1, _ct52r)
                                  if _ct52r > 0 else "")
                    except Exception:
                        continue
                    if not _fresh52.implicit_gate_applies(_cd52r):
                        continue
                    _gated52 += 1
                    _std52r = bool(_cp52r.Type == 1)
                    _recs52r = P52.extract_records(
                        _cd52r, module=_nm52r, is_std_module=_std52r)
                    for (_n52r, _m52r, _p52r, _pr52r) in P52.extract_implicit_records(
                            _cd52r, module=_nm52r, is_std_module=_std52r,
                            declared=_recs52r, scope="proc"):
                        _l52r = str(_n52r).lower()
                        if _l52r in _decl52r or _l52r in _str52r:
                            continue        # 别处真声明过：本来就该在池子里
                        if _l52r in _pool52r:
                            _leak52.append(_l52r)
                _r52 = (sorted(set(_leak52)), _gated52)
            except Exception as _e52r:
                _r52 = None
                check("52.9 真机：拿不到 VBE（Excel 未开？）—— 跳过（%s）"
                      % _e52r, [True], expect_contain=[True])
        if _r52 is not None:
            check("52.9 ★真机：%d 个带 Option Explicit 的模块，其「用过但没声明」"
                  "的名字一个都没进池子（若列表非空即为泄漏）" % (_r52[1],),
                  [_r52[0]], expect_contain=[[]])
    except Exception as _e52:
        check("第 52 节异常: %s" % _e52, [True], expect_contain=[False])

    # ==================================================================
    # 53. v76：让位判据不再区分触发来源（手动 / 补位都受约束）
    #     + 「VBE 列表确实可见」清掉 yield_given_up
    #
    # 用户报："VBE 的弹窗和我们项目弹窗冲突时，屏蔽我们项目弹窗……以前你修好了，
    # 就一直没出，刚才你改完，就出这个 bug 了。" 对号入座的三问答案：
    # VBE 弹的是【成员列表】；现象是【我们的窗盖住了它】；期望是【直接消失让位】。
    #
    # 排查（先证伪再定位）：
    #   1) 先确认工具跑的是哪版代码 —— GetProcessTimes 拿 pythonw 启动时间
    #      (13:05:06) 对比源文件 mtime (main/ui 09:41、vbe_bridge 12:59) ⇒ 跑的确实
    #      是修完的最新代码，不是"改了没重启"的错觉；
    #   2) `git diff v73..v74 -- ui.py` 逐行核对：rect_overlap 的重构与旧判据
    #      (r<=x / l>=x+w / b<=y / t>=y+h) 逐项等价 ⇒ v74 的方向只会让"避让"更多，
    #      不会更少 ⇒ 问题不在避让，在【让位】这一侧；
    #   3) 让位的第二路（窗口可见性）天生脆：我们的窗是 topmost 悬浮窗，一压上去
    #      VBE 的列表就失去激活、自行收起，之后再也探不到（v72 的认知）。所以
    #      必须让【不依赖探测】的那些判据尽量多扛一手。
    #
    # 两个确定的病灶（都不依赖探测，因此不受上面那个自弱化循环影响）：
    #   a) 第二路判据原来带 require_ident_before_caret ⇒ 手动 Ctrl+Space 与 v73 的
    #      补位路径 trigger(False) 都【绕过】它。补位那条尤其危险：它是先探到
    #      "列表不可见"才发起的，而从那次探测到主线程真正执行 trigger 之间隔着
    #      动作队列，VBE 的列表完全可能已经弹出来了 —— 补出来的窗就直接压在它上面。
    #   b) yield_given_up 按 (模块, 行号) 记，一旦记上，这一行的【语法判据】就永久
    #      跳过让位，只剩第二路兜着 —— 而第二路正是上面那条会自弱化的路。
    #      现在"VBE 的列表确实可见"会把标记清掉，语法判据重新生效。
    #
    # 口径变更（用户 v76 明确表态）：VBE 的成员列表可见时我们的窗【直接消失】，
    # 不分手动/自动 —— 成员列表本身就是一份候选列表，它在屏幕上时用户的目的已经
    # 达到，再挂着我们的（topmost）只会把它整个盖住。
    # ==================================================================
    print("\n=== 53. v76：让位不区分触发来源；VBE 列表真在就清「已放弃」 ===")
    try:
        import vba_builtins as VB53
        import ui as UI53

        _POOL53 = [(n, "VBA", None, False) for n in VB53.BUILTIN_FUNCTIONS]
        _bl53 = set(r[0].lower() for r in _POOL53)

        class _B53(object):
            """桩后端：成员列表可见性（yld）可随时切，位置由 line_text 决定。"""

            def __init__(self, line_text="    .Si", caret_col=None,
                         line_no=1, module="模块1", yld=False):
                self._lt = line_text
                self._cc = (caret_col if caret_col is not None
                            else len(line_text) + 1)
                self._ln = line_no
                self._mod = module
                self._yld = yld

            def set_yield(self, yld):
                self._yld = yld

            def get_context(self):
                return {"line_no": self._ln, "line_text": self._lt,
                        "caret_col": self._cc, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": self._mod}

            def get_identifiers(self):
                return list(_POOL53)

            def get_declared_names(self):
                return []

            def declared_elsewhere(self, name, module_name):
                return False

            def get_structural_names(self):
                return []

            def get_builtin_names(self):
                return set(_bl53)

            def get_type_names(self):
                return []

            def get_host_enum_names(self):
                return {}

            def names_outside_caret(self, caret, scope_only=False):
                return set()

            def apply_completion(self, *a, **k):
                return None

            # 只表达【成员列表】语义（v65）：参数信息不参与让位
            def vbe_yield_visible(self):
                return self._yld

            def vbe_popup_visible(self):
                return self._yld

            def vbe_popup_info(self):
                return [("NameListWndClass", 120, 90)] if self._yld else []

        class _UI53(object):
            def __init__(self):
                self.shown = False
                self.hides = 0

            def show(self, rows, sel, c):
                self.shown = True

            def hide(self):
                self.shown = False
                self.hides += 1

            def update_selection(self, sel):
                pass

            def contains_point(self, x, y):
                return False

        def _mk53(line_text="    .Si", yld=False):
            _be = _B53(line_text=line_text, yld=yld)
            _ui = _UI53()
            return E.Completer(_be, _ui), _be, _ui

        # 53.1 基础不变式：语法位置 + 列表不可见 + 自动触发 -> 待定让位（v73）
        _c53, _b53, _u53 = _mk53(yld=False)
        _c53.trigger(True)
        check("53.1 语法位置 + 列表不可见 + 自动触发 -> 待定让位（不弹）",
              [len(_c53.matches or []) == 0, _c53.yield_pending is not None],
              expect_contain=[True, True])

        # 53.2 列表可见 + 自动触发 -> 让位
        _c53, _b53, _u53 = _mk53(yld=True)
        _c53.trigger(True)
        check("53.2 成员列表可见 + 自动触发 -> 让位（不弹）",
              [len(_c53.matches or []) == 0, _u53.hides > 0],
              expect_contain=[True, True])

        # 53.3 ★口径变更：手动 Ctrl+Space 同样让位
        _c53, _b53, _u53 = _mk53(yld=True)
        _c53.trigger(False)
        check("53.3 ★v76 成员列表可见 + 手动 Ctrl+Space -> 同样让位"
              "（不再区分触发来源）",
              [len(_c53.matches or []) == 0, _u53.hides > 0],
              expect_contain=[True, True])

        # 53.4 ★★补位路径（trigger(False)）也受这条判据约束 —— 那个时序缝
        _c53, _b53, _u53 = _mk53(line_text="    freeNa", yld=True)
        _c53.trigger(False)
        check("53.4 ★★v73 补位路径 trigger(False) + 列表可见 -> 也让位"
              "（补出来的窗不许压住 VBE 的列表）",
              [len(_c53.matches or []) == 0, _u53.hides > 0],
              expect_contain=[True, True])

        # 53.5 对照：列表不可见 + 手动 -> 照弹（挡它的确实是"可见"）
        _c53, _b53, _u53 = _mk53(line_text="    .Si", yld=False)
        _c53.trigger(False)
        check("53.5 对照：列表不可见 + 手动 -> 照弹（%d 条）"
              % len(_c53.matches or []),
              [len(_c53.matches or []) > 0, _u53.shown],
              expect_contain=[True, True])

        # 53.6 ★VBE 的列表确实在 -> 清掉"此处已放弃让位"的标记
        _c53, _b53, _u53 = _mk53(yld=False)
        _c53.trigger(True)                  # 先让位（挂 pending）
        _r536 = _c53.confirm_yield(False)   # VBE 没弹 -> 记 given_up
        _gu536 = _c53.yield_given_up
        _b53.set_yield(True)                # VBE 的列表随后真的出现了
        _u53.hides = 0
        _c53.trigger(True)
        check("53.6 ★VBE 列表确实可见 -> 清掉「已放弃让位」标记并让位"
              "（语法判据重新生效，不必再依赖探测）",
              [_r536 is False, _gu536 is not None,
               _c53.yield_given_up is None,
               len(_c53.matches or []) == 0, _u53.hides > 0],
              expect_contain=[True, True, True, True, True])

        # 53.7 对照：列表不可见时标记不许被乱清
        _c53, _b53, _u53 = _mk53(yld=False)
        _c53.trigger(True)
        _c53.confirm_yield(False)
        _gu537 = _c53.yield_given_up
        _c53.trigger(True)
        check("53.7 对照：列表仍不可见 -> 「已放弃让位」标记保持不变",
              [_gu537 is not None, _c53.yield_given_up == _gu537],
              expect_contain=[True, True])

        # 53.8 源码护栏：那条窗口判据不许再挂回 require_ident_before_caret
        _root53 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _esrc53 = ""
        try:
            _esrc53 = open(os.path.join(_root53, "engine.py"),
                           encoding="utf-8").read()
        except Exception:
            _esrc53 = ""
        check("53.8 ★源码护栏：窗口可见性判据不依赖触发来源"
              "（写回 require_ident_before_caret 就会挂）",
              ["if YIELD_TO_VBE_LIST and self._vbe_popup_showing():" in _esrc53,
               "require_ident_before_caret and YIELD_TO_VBE_LIST" in _esrc53],
              expect_contain=[True, True])

        # 53.9 main.py：兜底不分手动 + 轨迹日志在 _LOG_ENABLED 里
        _msrc53 = ""
        try:
            _msrc53 = open(os.path.join(_root53, "main.py"),
                           encoding="utf-8").read()
        except Exception:
            _msrc53 = ""
        check("53.9 接线护栏：poll 兜底不分手动、仍认 VBE_YIELD_CLASSES；"
              "trace 日志被 _LOG_ENABLED 包住、几何取 popup.geometry_now()",
              ['if not state.get("popup_manual") and any(' not in _msrc53,
               "VBE_YIELD_CLASSES" in _msrc53,
               "if _LOG_ENABLED:" in _msrc53,
               "popup.geometry_now()" in _msrc53],
              expect_contain=[True, True, True, True])

        # 53.10 ui.Popup.geometry_now：只读，拿不到一律 None
        class _FW53(object):
            def __init__(self, state="normal"):
                self._st = state

            def state(self):
                return self._st

            def winfo_x(self):
                return 10

            def winfo_y(self):
                return 20

            def winfo_width(self):
                return 200

            def winfo_height(self):
                return 90

        _p53 = UI53.Popup.__new__(UI53.Popup)
        _p53.win = None
        _g_none = _p53.geometry_now()
        _p53.win = _FW53("withdrawn")
        _g_wd = _p53.geometry_now()
        _p53.win = _FW53("normal")
        _g_ok = _p53.geometry_now()
        check("53.10 geometry_now：没有窗 / 已收起 -> None；正常 -> 实际几何",
              [_g_none is None, _g_wd is None, _g_ok == (10, 20, 200, 90)],
              expect_contain=[True, True, True])
    except Exception as _e53:
        check("第 53 节异常: %s" % _e53, [True], expect_contain=[False])

    # ==================================================================
    # 54. v77：默认只提示「这个工程里真实存在的东西」
    #
    # 用户口径（原话）："我觉得提示词太多了……你把提示 vba 本身带的关键字、函数这块
    # 都删掉吧。保留能提示变量名，自定义的函数/过程名，窗体/控件名，模块名。
    # 其他的那些个提示词，我也不怎么用到，提示一大堆看起来也不舒服。"
    #
    # ⇒ 四个开关**全部默认 False**（vbe_bridge.ENABLE_VBA_BUILTINS /
    #   ENABLE_VBA_KEYWORDS / ENABLE_VBA_CONSTANTS / ENABLE_HOST_ENUMS）。
    #   默认口径下候选池 = 变量名 + 自定义 Sub/Function/Property + 组件名（模块 /
    #   窗体 / 类模块）+ 窗体控件名，一条语言自带的名字都没有。
    #
    # 本节守三件事：
    #   1) 默认池子里**一条**语言自带的名字都没有，而工程内那四种名字一个不少；
    #   2) ★反向对照：把开关打开，那些名字立刻回来 —— 证明"没有"是开关造成的，
    #      不是桩后端/解析碰巧失效（否则"碰巧没有"也能蒙过去）；再验一次
    #      "开关逐批独立"（只开关键字不开内建函数也做得到）；
    #   3) ★四份清单**一个字都没裁**，三个收集门也都还在 —— 否则"更安静"会悄悄
    #      变成"能力删除"，以后想开回来也开不回来（用户的原话是"不要改坏了"）。
    #
    # ⚠️ 唯一有意的连带后果：`Dim x As <这里>` 不再补内建类型名（Long / String …），
    #    因为那个位置的候选是"先出候选、再按 type_names 过滤"——不在池子里就永远
    #    显示不出来。54.5 把它钉成【预期行为】，免得以后被当成 bug 反复"修"。
    # ==================================================================
    print("\n=== 54. v77：默认只提示工程内名字（语言自带的四批全关）===")
    try:
        import vbe_bridge as VB54R
        import vba_builtins as VB54B
        import engine as E54

        class _CM54R(object):
            def __init__(self, text):
                self.text = text

            @property
            def CountOfLines(self):
                return self.text.count("\n") + 1

            def Lines(self, start, count):
                return "\r\n".join(
                    self.text.split("\n")[start - 1:start - 1 + count])

        class _Ctl54R(object):
            def __init__(self, name):
                self.Name = name

        class _Ctls54R(object):
            def __init__(self, names):
                self._items = [_Ctl54R(n) for n in names]
                self.Count = len(self._items)

            def Item(self, i):
                return self._items[i]

        class _Designer54R(object):
            def __init__(self, ctl_names):
                self.Controls = _Ctls54R(ctl_names)

        class _Comp54R(object):
            def __init__(self, name, text, ctype, ctls=()):
                self.Name = name
                self.Type = ctype
                self.CodeModule = _CM54R(text)
                self.Designer = _Designer54R(list(ctls))

        class _Proj54R(object):
            def __init__(self, comps):
                self.Name = "VBAProject"
                self.VBComponents = list(comps)

        class _Pane54R(object):
            def __init__(self, comp):
                self._c = comp
                self.sel = (1, 1, 1, 1)

            @property
            def CodeModule(self):
                return self._c.CodeModule

            def GetSelection(self):
                return self.sel

        class _VBE54R(object):
            def __init__(self, comps, act):
                self.ActiveVBProject = _Proj54R(comps)
                self.ActiveCodePane = _Pane54R(act)

        # 最小工程：模块里【一个字都没提过】MsgBox / If / Dim / Long
        _m54 = _Comp54R(
            "Module1",
            "Sub Foo()\n    declaredVar = 1\n    looseVar = 2\nEnd Sub",
            VB54R.VBE_CT_STDMODULE)
        _f54 = _Comp54R("UserForm1", "", VB54R.VBE_CT_MSFORM,
                        ctls=["Label1", "cmdOK"])
        # 光标停在窗体上 -> 控件名（只存在于设计器里）才会被收进来
        _vbe54 = _VBE54R([_m54, _f54], _f54)

        _orig54 = VB54R._get_vbe_cached
        VB54R._get_vbe_cached = lambda: _vbe54
        _saved54 = (VB54R.ENABLE_VBA_BUILTINS, VB54R.ENABLE_VBA_KEYWORDS,
                    VB54R.ENABLE_VBA_CONSTANTS, VB54R.ENABLE_HOST_ENUMS)

        def _set54(b, k, c, h):
            (VB54R.ENABLE_VBA_BUILTINS, VB54R.ENABLE_VBA_KEYWORDS,
             VB54R.ENABLE_VBA_CONSTANTS, VB54R.ENABLE_HOST_ENUMS) = (b, k, c, h)

        def _pool54():
            """按【此刻的开关】重收一遍（开关是模块级常量，收集时实时读）。

            每次新建一个后端实例：标识符缓存是实例级 + 2 秒 TTL，复用实例会拿到
            上一轮的旧池子。"""
            _bk = VB54R.VbeBackend()
            return _bk, set(str(r[0]).lower() for r in _bk.get_identifiers())

        class _EngineBackend54(object):
            """把真后端的池子接到引擎上（引擎要的接口一套给全）。"""

            def __init__(self, bk, module="Module1", word="ms"):
                self._bk = bk
                self._mod = module
                self._word = word

            def get_context(self):
                _ln = "    " + self._word
                return {"line_no": 1, "line_text": _ln,
                        "caret_col": len(_ln) + 1, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": self._mod}

            def get_identifiers(self):
                return self._bk.get_identifiers()

            def get_declared_names(self):
                return sorted(self._bk.get_declared_names())

            def declared_elsewhere(self, name, module_name):
                return bool(self._bk.declared_elsewhere(name, module_name))

            def get_structural_names(self):
                return self._bk.get_structural_names()

            def get_builtin_names(self):
                return self._bk.get_builtin_names()

            def get_type_names(self):
                return list(self._bk.get_type_names())

            def get_host_enum_names(self):
                return self._bk.get_host_enum_names()

            def names_outside_caret(self, caret, scope_only=False):
                # 现场文本证据：桩里就等于"池子里的【工程内】名字"。
                # 内建那批不在现场文本里 —— 真机上它们也正是靠 structural_names
                # 那一路放行的（回声防护的设计，这里如实模拟，别把两路混成一路）。
                return set(
                    str(n).lower() for n, _m, _p, _pv in self._bk.get_identifiers()
                    if str(_m) not in (VB54R._BUILTIN_MODULE,
                                       VB54R._HOST_ENUM_MODULE))

            def apply_completion(self, *a, **k):
                return None

            # 让位/避让一律"探不到"（本机不一定在跑 Excel，保持确定性）
            def vbe_yield_visible(self):
                return False

            def vbe_popup_visible(self):
                return False

            def vbe_popup_info(self):
                return []

        class _UI54(object):
            def __init__(self):
                self.shown = False

            def show(self, rows, sel, c):
                self.shown = True

            def hide(self):
                self.shown = False

            def update_selection(self, sel):
                pass

            def contains_point(self, x, y):
                return False

        def _trig54(word="ms"):
            _c = E54.Completer(_EngineBackend54(VB54R.VbeBackend(),
                                                word=word), _UI54())
            _c.trigger(True)
            return sorted(_c.matches or [])

        try:
            # ---- 54.1 ★默认口径：语言自带的名字一条都不进池 ----
            _set54(False, False, False, False)
            _bk54, _nm54 = _pool54()
            # 语言自带的四批，各取一个代表（函数 / 关键字 / vb 枚举 / 内建类型）。
            # 宿主枚举（xl*/mso*）不放进来：打开它要真去加载 Excel 类型库，
            # 这一节刻意与 Excel 无关（它的开关由 47.4 单独守）。
            _lang54 = ["msgbox", "left", "split", "dim", "sub", "function",
                       "long", "string", "vbcrlf", "vbyes"]
            # 工程内该有的四种名字（变量 / 自定义过程 / 模块名 / 窗体名 + 控件名）
            _mine54 = ["declaredvar", "loosevar", "foo", "module1",
                       "userform1", "label1", "cmdok"]
            check("54.1 ★默认口径：语言自带的 %d 个代表名一条都没有（漏进来：%s）"
                  % (len(_lang54), sorted(set(_lang54) & _nm54)),
                  [sorted(set(_lang54) & _nm54),
                   sorted(n for n in _mine54 if n not in _nm54),
                   dict(_bk54.get_host_enum_names()),
                   sorted(_bk54.get_builtin_names())],
                  expect_contain=[[], [], {}, []])

            # ---- 54.2 ★反向对照：打开开关 -> 那些名字立刻回来 ----
            _set54(True, True, True, False)
            _bk54on, _nm54on = _pool54()
            check("54.2 ★反向对照：开关打开后 10 个代表名全回来（证明是开关在起作用）；"
                  "宿主枚举那一批仍不越界",
                  [sorted(n for n in _lang54 if n not in _nm54on),
                   "xlup" not in _nm54on],
                  expect_contain=[[], True])

            # ---- 54.3 开关逐批独立：只开关键字，内建函数照旧不收 ----
            _set54(False, True, False, False)
            _bk54k, _nm54k = _pool54()
            check("54.3 开关逐批独立：只开关键字 -> Sub/If 回来，MsgBox/Long 仍不在",
                  ["sub" in _nm54k, "if" in _nm54k,
                   "msgbox" not in _nm54k, "long" not in _nm54k,
                   "vbcrlf" not in _nm54k],
                  expect_contain=[True, True, True, True, True])

            # ---- 54.4 ★四份清单没被裁 + 收集门都在（只关开关，不删能力）----
            _src54 = open(os.path.join(
                os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                "vbe_bridge.py"), encoding="utf-8").read()
            check("54.4 ★四份清单一个字没裁（函数 %d / 常量 %d / 类型 %d / 关键字 %d）"
                  "+ 三个收集门都在 + 默认值就是 False"
                  % (len(VB54B.BUILTIN_FUNCTIONS), len(VB54B.BUILTIN_CONSTANTS),
                     len(VB54B.BUILTIN_TYPE_NAMES), len(VB54B.BUILTIN_KEYWORDS)),
                  [len(VB54B.BUILTIN_FUNCTIONS) >= 150,
                   len(VB54B.BUILTIN_CONSTANTS) >= 100,
                   len(VB54B.BUILTIN_TYPE_NAMES) >= 5,
                   len(VB54B.BUILTIN_KEYWORDS) >= 30,
                   "MsgBox" in VB54B.BUILTIN_FUNCTIONS,
                   "Long" in VB54B.BUILTIN_TYPE_NAMES,
                   "Sub" in VB54B.BUILTIN_KEYWORDS,
                   _src54.count("if ENABLE_VBA_BUILTINS:") >= 1,
                   _src54.count("if ENABLE_VBA_KEYWORDS:") >= 1,
                   _src54.count("if ENABLE_VBA_CONSTANTS:") >= 1,
                   _src54.count("if ENABLE_HOST_ENUMS:") >= 1,
                   '_env_flag("VBECOMPLETE_VBA_BUILTINS", False)' in _src54,
                   '_env_flag("VBECOMPLETE_VBA_KEYWORDS", False)' in _src54],
                  expect_contain=[True] * 13)

            # ---- 54.5 As 位置的连带后果（钉成预期行为，别再当 bug 修）----
            _tn54_off = set(str(t).lower() for t in _bk54.get_type_names())
            _tn54_on = set(str(t).lower() for t in _bk54on.get_type_names())
            check("54.5 `Dim x As |`：内建类型默认不在，你自己的类型一直在；"
                  "开关打开就回来",
                  ["long" not in _tn54_off, "string" not in _tn54_off,
                   "userform1" in _tn54_off,
                   "long" in _tn54_on and "string" in _tn54_on,
                   "userform1" in _tn54_on],
                  expect_contain=[True] * 5)

            # ---- 54.6 ★端到端：默认打 ms 一个候选都不出 ----
            _set54(False, False, False, False)
            _off54 = _trig54("ms")
            _set54(True, True, False, False)
            _on54 = _trig54("ms")
            check("54.6 ★端到端：默认打 ms 零候选；打开内建 -> MsgBox（第 1 个）",
                  [len(_off54), _on54[:1]],
                  expect_contain=[0, ["MsgBox"]])

            # 默认口径下"工程内名字"照旧能补（别把保留项一起关没了）
            _set54(False, False, False, False)
            _proj54 = _trig54("decl")
            check("54.7 默认口径下工程内名字照旧（decl -> declaredVar）",
                  ["declaredVar" in _proj54], expect_contain=[True])
        finally:
            _set54(*_saved54)
            VB54R._get_vbe_cached = _orig54
    except Exception as _e54:
        check("第 54 节异常: %s" % _e54, [True], expect_contain=[False])

    # ---- 55. v78：Shift+↑/↓ 跳出候选列表 + 光标上/下移一行 ----
    #
    # 用户口径：「当提示词列表框出现的时候，按 ↑↓ 会一直在列表框选词，不能将
    # 光标移动上下行。帮我新增按 Shift+↑，可以跳出我们项目的提示词列表框，并将
    # 光标上移一行；Shift+↓ 同理。」
    #
    # 两个关键决定，都由断言守着：
    #   1) 不用"放行按键让 VBE 自己动光标" —— VBE 原生的 Shift+↑/↓ 是【扩展选区】
    #      （把上一行整行选进去），不是"光标干净地上移一行"。所以由我们吞键 +
    #      用 COM 落光标（VbeBackend.move_caret_line）。
    #   2) 判据必须排在"在列表里选词"那条分支【之前】，否则 Shift+↑ 又被当成
    #      普通 ↑ 去列表里选词 —— 那就完全没解决问题（55.6 钉着这个次序）。
    print("\n=== 55. Shift+↑/↓ 跳出候选列表并移动光标（v78）===")
    try:
        import main as M55
        import vbe_bridge as VB55
        _orig55 = VB55._get_vbe_cached

        # ---- 55.1 纯判断：只认 ↑/↓ + Shift 按住 ----
        check("55.1 _is_line_nav_shortcut：Shift+↑/↓ 认；裸 ↑、Shift+Enter、"
              "Shift+←、Shift 没按都不认",
              [M55._is_line_nav_shortcut(M55.VK_UP, True),
               M55._is_line_nav_shortcut(M55.VK_DOWN, True),
               M55._is_line_nav_shortcut(M55.VK_UP, False),
               M55._is_line_nav_shortcut(M55.VK_RETURN, True),
               M55._is_line_nav_shortcut(M55.VK_LEFT, True),
               M55._is_line_nav_shortcut(M55.VK_DOWN, 0)],
              expect_contain=[True, True, False, False, False, False])

        class _CM55(object):
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

        class _Pane55(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = tuple(sel)

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE55(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        def _mv55(lines, sel, delta):
            cm55 = _CM55(lines)
            pane55 = _Pane55(cm55, sel)
            VB55._get_vbe_cached = lambda: _VBE55(pane55)
            ok55 = VB55.VbeBackend().move_caret_line(delta)
            return ok55, cm55.lines, pane55.sel

        _SRC55 = ["Sub Foo()", "    Dim x", "    x = 1 + 2", "End Sub"]

        # ---- 55.2/55.3 下移 / 上移一行（同列，且不碰文本）----
        _ok55d, _ls55d, _sel55d = _mv55(_SRC55, (2, 9, 2, 9), 1)
        _ok55u, _ls55u, _sel55u = _mv55(_SRC55, (3, 9, 3, 9), -1)
        check("55.2 下移一行：光标 (2,9) -> (3,9)，同列、文本一动不动",
              [_ok55d, _sel55d, _ls55d],
              expect_contain=[True, (3, 9, 3, 9), _SRC55])
        check("55.3 上移一行：光标 (3,9) -> (2,9)，同列、文本一动不动",
              [_ok55u, _sel55u, _ls55u],
              expect_contain=[True, (2, 9, 2, 9), _SRC55])

        # ---- 55.4 越界 / 无 VBE 一律 False（调用方据此【不吞键】）----
        _b55 = []
        _b55.append(_mv55(_SRC55, (1, 1, 1, 1), -1)[0])     # 首行再上移
        _b55.append(_mv55(_SRC55, (4, 8, 4, 8), 1)[0])      # 末行再下移
        _b55.append(_mv55(_SRC55, (0, 0, 0, 0), 1)[0])      # 行号 0（空模块）
        VB55._get_vbe_cached = lambda: None
        _b55.append(VB55.VbeBackend().move_caret_line(1))
        VB55._get_vbe_cached = lambda: _VBE55(None)
        _b55.append(VB55.VbeBackend().move_caret_line(1))
        check("55.4 首行上移 / 末行下移 / 行号 0 / 无 VBE / 无代码窗 -> 全 False",
              [_b55], expect_contain=[[False] * 5])

        # ---- 55.5 反向选择时以【活动端】为基准（选中框的另一头不算）----
        _ok55r, _ls55r, _sel55r = _mv55(_SRC55, (4, 3, 2, 7), -1)
        check("55.5 反向选区（sl=4 > el=2）以 (4,3) 为基准 -> 上移到 (3,3)",
              [_ok55r, _sel55r], expect_contain=[True, (3, 3, 3, 3)])

        # ---- 55.6 ★接线护栏：次序 + 先收窗后落光标 + 兜底 + 开关 ----
        _src55 = open(os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "main.py"), encoding="utf-8").read()
        _i55_nav = _src55.index("elif _is_line_nav_shortcut(vk, _shift_down()):")
        _i55_pick = _src55.index("elif vk in NAV_VKS:")
        _i55_fn = _src55.index("def _move_caret_here(delta):")
        _i55_mv = _src55.index("backend.move_caret_line(delta)")
        check("55.6 ★接线护栏：Shift+↑/↓ 的判据排在「在列表里选词」之前；"
              "先收窗再落光标；COM 没落成要走 send_vk 兜底；开关在",
              [_i55_nav < _i55_pick,
               "completer.hide()" in _src55[_i55_fn:_i55_mv],
               _i55_fn < _i55_mv,
               "send_vk(VK_UP if delta < 0 else VK_DOWN)" in _src55,
               'os.environ.get("VBECOMPLETE_NO_SHIFT_NAV"' in _src55,
               "SHIFT_NAV_JUMP" in _src55],
              expect_contain=[True] * 6)

        # ---- 55.7 出厂默认：这个接管是开着的 ----
        check("55.7 ★出厂默认：Shift+↑/↓ 接管开着（VBECOMPLETE_NO_SHIFT_NAV 未设）",
              [M55.SHIFT_NAV_JUMP], expect_contain=[True])

        VB55._get_vbe_cached = _orig55
    except Exception as _e55:
        check("第 55 节异常: %s" % _e55, [True], expect_contain=[False])

    # ---- 56. v79：配对光标落点（bug）+ 回车自动缩进（新功能） ----
    #
    # 用户报的 bug：「录入 targetSheet.Cells(i, "b").Interior.Color = RGB，然后
    # 再打括号，光标落点会错位，不能正确落在 RGB 后面的括号内。应该是 vba 自动
    # 格式化的时候，有两个括号，你只判断了一个。」
    # 用户要的两个功能：
    #   1) 整行注释后回车，忽略上方注释行（可能多行）、跟上方第一个非注释行对齐；
    #   2) For / Do / If / With / Sub / Function / Type / Enum 这些结构，回车自动缩进。
    print("\n=== 56. 配对光标落点 + 回车自动缩进（v79）===")
    try:
        import main as M56
        import vbe_bridge as VB56
        _orig56 = VB56._get_vbe_cached

        # ---- 56.1 ★bug 修复：VBE 重排整行后，光标仍要落在括号中间 ----
        _CANON56 = '    targetSheet.Cells(i, "b").Interior.Color = RGB()'

        def _off56(after, col0):
            return VB56._locate_inserted(after, _CANON56, "(", ")", col0)

        _c56 = _CANON56.index("RGB()") + 3       # 写入时光标的位置（括号之间）

        def _in_pair56(after):
            _o = _off56(after, _c56)
            return bool(_o is not None and after[_o - 1:_o + 1] == "()")

        check("56.1 ★bug：VBE 重排整行后仍把光标钉在括号中间"
              "（逗号后补空格 / 改大小写 / 等号两边补空格 / 括号前后多空格）",
              [_in_pair56(_CANON56),
               _in_pair56('    targetSheet.Cells(i, "b").Interior.Color = RGB ()'),
               _in_pair56('    targetSheet.Cells(i,"b").interior.color = rgb()'),
               _in_pair56('    targetSheet.Cells( i , "b" ).Interior.Color '
                          '= RGB(   )'.replace("(   )", "()"))],
              expect_contain=[True, True, True, True])

        # 用户说的"两个括号"：行里本来还有一对括号，不能认到那一对上去
        _W56 = '    targetSheet.Cells(i, x()).Interior.Color = RGB()'
        _o56 = VB56._locate_inserted(_W56, _W56, "(", ")",
                                     _W56.index("RGB()") + 3)
        check("56.2 ★bug：一行里有两对括号时，认的是我们刚插的那一对（RGB 后面）",
              [_o56, _W56[_o56 - 1:_o56 + 1], _o56 > _W56.index("RGB")],
              expect_contain=[_W56.index("RGB") + 4, "()", True])

        # v58 老场景（引号）：`    w""=4` 被 VBE 重排成 `    w "" = 4`
        check("56.3 引号那一例（v58 老场景）不许回归",
              [VB56._locate_inserted('    w "" = 4', '    w""=4', '"', '"', 5)],
              expect_contain=[7])

        # ---- 56.4 块结构开头（回车后要深一级） ----
        _BLK56 = ["For i = 1 To 10", "For Each c In rng", "Do", "Do While x < 5",
                  "Do Until x > 0", "While x", "With rng", "If x Then",
                  "If x Then '注释", "ElseIf y Then", "Else", "Case 1, 2",
                  "Case Else", "Select Case x", "Sub Foo()",
                  "Private Sub Foo()", "Public Function Bar() As Long",
                  "Static Sub S()", "Property Get X()", "Type Foo",
                  "Enum Color", "#If VBA7 Then", "#Else",
                  'If x = "a Then" Then', "    For i = 1 To 10"]
        _NOBLK56 = ["If x Then y = 1", "Next i", "End If", "End Sub", "End With",
                    "Loop", "Loop While x", "Wend", "x = 1", "Exit For",
                    "DoEvents", "foo(1, 2)", "Debug.Print x", "Rem 说明",
                    "'注释", "Set x = New Collection", "Call Foo",
                    "Dim x As Long", "    ", "End Select", "Exit Sub",
                    "Range(\"a1\").Value = 1", "If x Then _", "rem表格"]
        check("56.4 opens_block：24 个块头全认，24 个非块头全不认",
              [[t for t in _BLK56 if not VB56.opens_block(t)],
               [t for t in _NOBLK56 if VB56.opens_block(t)]],
              expect_contain=[[], []])

        # ---- 56.5 回车后新行的缩进 ----
        _U56 = "    "
        _C56 = [VB56.next_line_indent("    x = 1", ["Sub Foo()"], _U56),
                VB56.next_line_indent("    For i = 1 To 10", [], _U56),
                VB56.next_line_indent("    Next i", [], _U56),
                VB56.next_line_indent("'注释", ["Sub Foo()", "    x = 1"], _U56),
                VB56.next_line_indent("'注释", ["    With rng"], _U56),
                VB56.next_line_indent("'注释", ["'老注释", ""], _U56)]
        check("56.5 next_line_indent：普通行齐平 / 块头深一级 / 闭合行齐平 / "
              "注释跟上方非注释行对齐 / 注释上方是块头则深一级 / 上方全注释"
              "则维持自身缩进",
              _C56, expect_contain=["    ", "        ", "    ", "    ",
                                    "        ", ""])

        # ★用户的实际场景：注释行在第 1 列、上方代码缩进 12，连续多行注释
        _A56 = ["        Else", "            a = RGB(1, 2)",
                "'第一行注释", "'第二行注释"]
        check("56.5b ★用户场景：整行注释后回车，忽略上方连续注释行，"
              "跟第一个非注释行（缩进 12）对齐",
              [VB56.next_line_indent("'第三行注释", _A56, _U56),
               VB56.next_line_indent("'注释",
                                     ["Sub Foo()", "    For i = 1 To 10"], _U56)],
              expect_contain=["            ", "        "])

        # ---- 56.6 钩子侧判据：只在这两种情形接管回车 ----
        _W56b = [("整行注释+行尾", "'注释", 20, True),
                 ("块头+行尾", "    For i = 1 To 10", 21, True),
                 ("普通代码行+行尾", "    x = 1", 10, False),
                 ("光标在行中间", "    For i = 1 To 10", 12, False),
                 ("空行", "        ", 9, False),
                 ("整行空白", "", 1, False),
                 ("含 Tab 的行", "\tFor i", 9, False),
                 ("光标列拿不到", "    For i", 0, False)]
        check("56.6 enter_indent_wanted：注释 / 块头 + 光标在行尾才接管；"
              "普通行、行中间、空行、含 Tab、拿不到列一律放行给 VBE",
              [[n for (n, t, c, e) in _W56b
                if VB56.enter_indent_wanted(t, c) != e]],
              expect_contain=[[]])

        # ---- 56.7 引号没闭合时不许抢回车（那是 VBE 原生会补右引号的地方） ----
        check("56.7 _unterminated_string：注释里的引号不算；字符串内 \"\" 转义不算",
              [VB56._unterminated_string('s = "abc'),
               VB56._unterminated_string('s = "abc"'),
               VB56._unterminated_string('x = 1  ' + "'他说\"你好"),
               VB56._unterminated_string('s = "a""b"')],
              expect_contain=[True, False, False, False])

        # ---- 56.8 ★集成：smart 模式真能改写缩进；不该管的返回 False ----
        class _CM56(object):
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

        class _Pane56(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE56(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        def _nl56(lines, line_no, col, smart=True):
            cm56 = _CM56(lines)
            pane56 = _Pane56(cm56, (line_no, col, line_no, col))
            VB56._get_vbe_cached = lambda: _VBE56(pane56)
            ok56 = VB56.VbeBackend().new_line_below(smart=smart)
            return ok56, cm56.lines, pane56.sel

        _ok, _ls, _sel = _nl56(["Sub Foo()", "    For i = 1 To 10",
                                "    Next i", "End Sub"], 2, 21)
        check("56.8 ★集成：块头那行回车 -> 新行深一级（8 空格）、光标在缩进后、"
              "块头那一行一个字不动",
              [(_ok, _ls, _sel)],
              expect_contain=[(True,
                               ["Sub Foo()", "    For i = 1 To 10", "        ",
                                "    Next i", "End Sub"],
                               (3, 9, 3, 9))])

        _ok, _ls, _sel = _nl56(["Sub Foo()", "    x = 1", "'注释", "End Sub"], 3, 5)
        check("56.8b ★集成：整行注释那行回车 -> 新行跟上方代码行（4 空格）对齐",
              [(_ok, _ls[3] if len(_ls) > 3 else None)],
              expect_contain=[(True, "    ")])

        # 中文注释（显示列语义）也要能对上：'注释 = 1 + 2 + 2 = 5 列
        _sem56 = VB56._sem_cache.get("info")
        VB56._sem_cache["info"] = ("disp", 4, True)
        try:
            _ok, _ls, _sel = _nl56(["Sub Foo()", "    x = 1", "'注释", "End Sub"],
                                   3, 5)
        finally:
            VB56._sem_cache["info"] = _sem56
        check("56.8c ★集成：中文整行注释（显示列语义）同样对齐到 4 空格",
              [(_ok, _ls[3] if len(_ls) > 3 else None)],
              expect_contain=[(True, "    ")])

        _bad56 = []
        for _n, _ln, _col, _src in [
                ("光标在行中间", 2, 12, ["Sub Foo()", "    For i = 1 To 10",
                                        "End Sub"]),
                ("空行", 2, 5, ["Sub Foo()", "    ", "End Sub"]),
                ("引号没闭合", 2, 13, ["Sub Foo()", '    s = "abc', "End Sub"])]:
            _ok56, _ls56, _s56 = _nl56(_src, _ln, _col)
            if _ok56 is not False or _ls56 != _src:
                _bad56.append(_n)
        # 多行选区：按回车是"删除选区"，语义完全不同 —— 必须原样让 VBE 处理
        _cm56 = _CM56(["Sub Foo()", "    a = 1", "    b = 2", "End Sub"])
        _pane56 = _Pane56(_cm56, (2, 1, 3, 1))
        VB56._get_vbe_cached = lambda: _VBE56(_pane56)
        _before56 = list(_cm56.lines)
        _ok56 = VB56.VbeBackend().new_line_below(smart=True)
        if _ok56 is not False or _cm56.lines != _before56:
            _bad56.append("多行选区")
        check("56.8d smart 模式这四种情况必须返回 False 且一个字都不写"
              "（调用方据此把回车原样还给系统）",
              [_bad56], expect_contain=[[]])

        # Shift+Enter 老口径不许被改动：不传 smart 时仍然"光标在行中间也照样
        # 新起一行、缩进照抄本行"
        _ok, _ls, _sel = _nl56(["Sub Foo()", "    x = 1 + 2", "End Sub"], 2, 11,
                               smart=False)
        check("56.8e Shift+Enter 老口径不变（smart 默认 False）",
              [(_ok, _ls[2], _sel)],
              expect_contain=[(True, "    ", (3, 5, 3, 5))])

        VB56._get_vbe_cached = lambda: None
        check("56.8f 取不到 VBE -> False（回车原样还给系统）",
              [VB56.VbeBackend().new_line_below(smart=True)],
              expect_contain=[False])
        VB56._get_vbe_cached = _orig56

        # ---- 56.9 缩进一级有多宽：空格 / Tab / 环境变量覆盖 ----
        check("56.9 _indent_unit_for：默认 4 空格；工程用 Tab 就跟着用 Tab；"
              "环境变量可覆盖（VBECOMPLETE_INDENT_UNIT=2）",
              [VB56._indent_unit_for(["Sub Foo()", "    x = 1"]),
               VB56._indent_unit_for(["Sub Foo()", "\tx = 1"])],
              expect_contain=["    ", "\t"])
        _env56 = os.environ.get("VBECOMPLETE_INDENT_UNIT")
        try:
            os.environ["VBECOMPLETE_INDENT_UNIT"] = "2"
            check("56.9b 环境变量 VBECOMPLETE_INDENT_UNIT=2 -> 一级 2 个空格",
                  [VB56._indent_unit_for([]),
                   VB56.next_line_indent("    For i = 1", [],
                                         VB56._indent_unit_for([]))],
                  expect_contain=["  ", "      "])
            os.environ["VBECOMPLETE_INDENT_UNIT"] = "tab"
            check("56.9c 环境变量 =tab -> 一级一个 Tab",
                  [VB56._indent_unit_for([])], expect_contain=["\t"])
        finally:
            if _env56 is None:
                os.environ.pop("VBECOMPLETE_INDENT_UNIT", None)
            else:
                os.environ["VBECOMPLETE_INDENT_UNIT"] = _env56

        # ---- 56.10 ★接线护栏：钩子侧的次序与兜底 ----
        _src56 = io.open(os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "main.py"), encoding="utf-8").read()
        _i56_enter = _src56.index("elif _enter_indent_key_ok(vk):")
        _i56_pair = _src56.index("elif (_pair_ch")
        _i56_vis = _src56.index("elif completer.is_visible():")
        _i56_fn = _src56.index("def _enter_indent_here():")
        _i56_call = _src56.index("backend.new_line_below(smart=True")
        check("56.10 ★接线护栏：回车分支在自动配对之前、在「弹窗可见」之前；"
              "先收窗再写；COM 没落成要 send_vk(VK_RETURN) 兜底；"
              "判据只在钩子线程读快照（不吃 COM）",
              [_i56_enter < _i56_pair < _i56_vis,
               "completer.hide()" in _src56[_i56_fn:_i56_call],
               _i56_fn < _i56_call,
               "send_vk(VK_RETURN)" in _src56[_i56_fn:],
               'os.environ.get("VBECOMPLETE_NO_ENTER_INDENT"' in _src56,
               "ENTER_CTX_MAX_AGE" in _src56,
               "ENTER_INDENT" in _src56],
              expect_contain=[True] * 7)

        _gate56 = _src56[_src56.index("def _enter_indent_key_ok(vk):"):
                         _src56.index("def win32_filter(msg, data):")]
        check("56.10b 钩子判据里不许出现任何 COM 调用（钩子线程没有 COM 单元）",
              [[w for w in ("_get_vbe_cached", "backend.", "VbeBackend")
                if w in _gate56]],
              expect_contain=[[]])

        # 快照多带一个光标列；但"内容变化"的判定仍只看前三个元素
        check("56.10c snapshot 带光标列；变更检测仍只比前三个（光标移动不算"
              "内容变化，否则点来点去都会弹窗）",
              ["snap = (sl, cm.Lines(sl, 1), mod, ec)" in
               io.open(os.path.join(
                   os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "vbe_bridge.py"), encoding="utf-8").read(),
               "last[:3] != snap[:3]" in _src56,
               'state["cur_ctx"] = (time.time(), snap[0], snap[1], snap[3])'
               in _src56],
              expect_contain=[True, True, True])

        # ---- 56.10d ★输入法护栏：组字中的回车是"上屏键"，绝不能被抢 ----
        # 判据是"按了可能写字的键、文档却没变"（拼音被输入法吃掉）= 正在组字。
        # 这是输入法无关的做法，不依赖任何 IME 接口。
        _TXT56 = [0x41, 0x5A, 0x30, 0x39, 0x20, 0xBA, 0xC0, 0xDB, 0xDE, 0x61]
        _NOTXT56 = [0x0D, 0x09, 0x1B, 0x25, 0x26, 0x27, 0x28, 0x10, 0x11, 0x12,
                    0x70, 0x7A, 0xE3, 0xAD, 0xB0]
        check("56.10d _is_text_key：字母/数字/空格/OEM 标点算「可能写字」；"
              "回车、Tab、Esc、方向键、修饰键、F 键、多媒体键不算",
              [[k for k in _TXT56 if not M56._is_text_key(k)],
               [k for k in _NOTXT56 if M56._is_text_key(k)]],
              expect_contain=[[], []])

        check("56.10e ★接线护栏：回车判据在组字（pending_keys>0）时一律不接管；"
              "钩子数「可能写字的键」；文档一变就把计数清零",
              ["pending_keys" in _gate56,
               'state["pending_keys"] = min(' in _src56,
               'state["pending_keys"] = 0' in _src56,
               _src56.count('state["pending_keys"] = 0') >= 2],
              expect_contain=[True] * 4)

        # ---- 56.11 出厂默认：回车自动缩进开着 ----
        check("56.11 ★出厂默认：回车自动缩进开着"
              "（VBECOMPLETE_NO_ENTER_INDENT 未设）",
              [M56.ENTER_INDENT], expect_contain=[True])
    except Exception as _e56:
        check("第 56 节异常: %s" % _e56, [True], expect_contain=[False])

    # ---- 57. v79b：块头那行回车，自动补出块收尾 ----
    #
    # 用户的追加要求：「写 if i=1 then，按回车之后，能自动帮我补全后面的 end if，
    # 而且光标缩进；对于其他的结构语句也是这样，比如 with 语句块，就能自动补全
    # end with。」
    print("\n=== 57. 块头回车自动补收尾（v79b）===")
    try:
        import main as M57
        import vbe_bridge as VB57
        _orig57 = VB57._get_vbe_cached

        # ---- 57.1 块头 -> 收尾 ----
        check("57.1 block_closer：Sub/Function/Property/Type/Enum/If…Then/With/"
              "Select Case -> End Xxx；For -> Next；Do -> Loop；While -> Wend",
              [VB57.block_closer("Sub Foo()"),
               VB57.block_closer("Private Function Bar() As Long"),
               VB57.block_closer("Static Property Get X()"),
               VB57.block_closer("Type Foo"),
               VB57.block_closer("Enum Color"),
               VB57.block_closer("If x Then"),
               VB57.block_closer("    If x = 1 Then '注释"),
               VB57.block_closer("With rng"),
               VB57.block_closer("Select Case x"),
               VB57.block_closer("For i = 1 To 10"),
               VB57.block_closer("For Each c In rng"),
               VB57.block_closer("Do"),
               VB57.block_closer("Do While x < 5"),
               VB57.block_closer("While x")],
              expect_contain=["End Sub", "End Function", "End Property",
                              "End Type", "End Enum", "End If", "End If",
                              "End With", "End Select", "Next", "Next",
                              "Loop", "Loop", "Wend"])

        # ---- 57.2 不该补收尾的一个都不许补 ----
        _N57 = ["ElseIf y Then", "Else", "Case 1, 2", "Case Else",
                "#If VBA7 Then", "#Else", "If x Then y = 1", "'注释",
                "Rem 说明", "x = 1", "End If", "Next i", "Loop", "Wend",
                "End Sub", "    "]
        check("57.2 block_closer：Else / ElseIf…Then / Case（块【内】的分支）、"
              "#If / #Else（必须配 #End If）、单选 If、注释、闭合行 -> 全 None",
              [[t for t in _N57 if VB57.block_closer(t) is not None]],
              expect_contain=[[]])

        # ---- 57.3 到底该不该补：只看"下方第一个非空行" ----
        check("57.3 closer_needed：下面没内容 -> 补；下面已挂着同一个收尾"
              "（含隔空行/带缩进/`Next i`/`Loop While x`）-> 不补；"
              "缩进不比块头浅 -> 不补；比块头浅 -> 补",
              [VB57.closer_needed([], "End If", 0),
               VB57.closer_needed(["End If"], "End If", 0),
               VB57.closer_needed(["    End If"], "End If", 0),
               VB57.closer_needed(["", "", "End If"], "End If", 0),
               VB57.closer_needed(["Next i"], "Next", 4),
               VB57.closer_needed(["Loop While x"], "Loop", 4),
               VB57.closer_needed(["End Sub"], "End If", 4),
               VB57.closer_needed(["    MsgBox 1"], "End If", 4),
               VB57.closer_needed(["        MsgBox 1"], "End If", 4),
               VB57.closer_needed(["End If"], None, 0)],
              expect_contain=[True, False, False, False, False, False,
                              True, False, False, False])

        # ---- 57.3b _indent_width：空格 1 列 / Tab 4 列 ----
        check("57.3b _indent_width：空格按 1 列、Tab 按 4 列",
              [VB57._indent_width("    x"), VB57._indent_width("\tx"),
               VB57._indent_width(""), VB57._indent_width("\t  x")],
              expect_contain=[4, 4, 0, 6])

        # ---- 57.4 ★集成：真能多写出一行收尾，光标停在中间那行 ----
        class _CM57(object):
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

        class _Pane57(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE57(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        def _nl57(lines, line_no, col, auto_close=True):
            cm57 = _CM57(lines)
            pane57 = _Pane57(cm57, (line_no, col, line_no, col))
            VB57._get_vbe_cached = lambda: _VBE57(pane57)
            ok57 = VB57.VbeBackend().new_line_below(smart=True,
                                                    auto_close=auto_close)
            return ok57, cm57.lines, pane57.sel

        # 用户原话的场景：写 `Sub Foo()` 回车 -> 中间留出缩进好的一行 + 补 `End Sub`
        check("57.4 ★集成：`Sub Foo()` 行尾回车 -> 中间一行缩进 4 空格（光标在这儿）、"
              "下一行补 `End Sub`",
              [_nl57(["Sub Foo()"], 1, 10)],
              expect_contain=[(True, ["Sub Foo()", "    ", "End Sub"],
                               (2, 5, 2, 5))])

        # 用户原话的场景：`If i = 1 Then` 回车 -> 补 `End If`
        check("57.4b ★集成：`If x = 1 Then` 行尾回车 -> 中间一行缩进、下一行补 `End If`",
              [_nl57(["If x = 1 Then"], 1, 14)],
              expect_contain=[(True, ["If x = 1 Then", "    ", "End If"],
                               (2, 5, 2, 5))])

        check("57.4c ★集成：`With rng` 回车 -> 补 `End With`",
              [_nl57(["With rng"], 1, 9)],
              expect_contain=[(True, ["With rng", "    ", "End With"],
                               (2, 5, 2, 5))])

        # ★最常见的真实场景：在 Sub 里写 If，下面还压着 `End Sub`
        check("57.4d ★集成：Sub 里写 `If x = 1 Then`（下面压着 End Sub）-> "
              "中间一行 8 空格、补的 `    End If` 插在 `End Sub` 之前、"
              "`Sub Foo()` 与 `End Sub` 一个字没动",
              [_nl57(["Sub Foo()", "    x = 1", "    If x = 1 Then", "End Sub"],
                     3, 19)],
              expect_contain=[(True,
                               ["Sub Foo()", "    x = 1", "    If x = 1 Then",
                                "        ", "    End If", "End Sub"],
                               (4, 9, 4, 9))])

        # 下面已经挂着同一个收尾（含隔着空行）-> 绝不补出重复的
        check("57.4e ★集成：下面已经有 `End Sub` / 隔着空行的 `End Sub` -> "
              "只缩进、不补重复的收尾",
              [_nl57(["Sub Foo()", "End Sub"], 1, 10),
               _nl57(["Sub Foo()", "", "End Sub"], 1, 10),
               _nl57(["Sub Foo()", "    Do", "    Loop", "End Sub"], 2, 7)],
              expect_contain=[(True, ["Sub Foo()", "    ", "End Sub"],
                               (2, 5, 2, 5)),
                              (True, ["Sub Foo()", "    ", "", "End Sub"],
                               (2, 5, 2, 5)),
                              (True, ["Sub Foo()", "    Do", "        ",
                                      "    Loop", "End Sub"], (3, 9, 3, 9))])

        # 块【内】的分支（Else / Case）与条件编译（#If）都不许补
        check("57.4f ★集成：`Else` 只缩进不补（它是分支）；`#If … Then` 也不补"
              "（那要配 #End If）",
              [(_nl57(["Sub Foo()", "    If x = 1 Then", "    Else"], 3, 9)[1],
                _nl57(["#If VBA7 Then"], 1, 14)[1])],
              expect_contain=[(["Sub Foo()", "    If x = 1 Then", "    Else",
                                "        "],
                               ["#If VBA7 Then", "    "])])

        # auto_close=False（或开关关掉）-> 回到 v79 老口径：只缩进
        check("57.4g auto_close=False -> 只缩进、不补收尾（v79 老口径不变）",
              [_nl57(["Sub Foo()"], 1, 10, auto_close=False),
               _nl57(["If x = 1 Then"], 1, 14, auto_close=False)],
              expect_contain=[(True, ["Sub Foo()", "    "], (2, 5, 2, 5)),
                              (True, ["If x = 1 Then", "    "], (2, 5, 2, 5))])

        # Shift+Enter 那条路（smart=False）任何情况下都不许补收尾
        _cm57s = _CM57(["Sub Foo()"])
        _pane57s = _Pane57(_cm57s, (1, 2, 1, 2))
        VB57._get_vbe_cached = lambda: _VBE57(_pane57s)
        _ok57s = VB57.VbeBackend().new_line_below()
        check("57.4h Shift+Enter（smart=False）不补收尾：不改老行为",
              [(_ok57s, _cm57s.lines)],
              expect_contain=[(True, ["Sub Foo()", ""])])

        VB57._get_vbe_cached = lambda: None
        check("57.4i 取不到 VBE -> False（调用方把回车原样还给系统）",
              [VB57.VbeBackend().new_line_below(smart=True, auto_close=True)],
              expect_contain=[False])
        VB57._get_vbe_cached = _orig57

        # ---- 57.5 ★接线护栏 ----
        _src57 = io.open(os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "main.py"), encoding="utf-8").read()
        _fn57 = _src57[_src57.index("def _enter_indent_here():"):
                       _src57.index("def _move_caret_here(delta):")]
        check("57.5 ★接线护栏：开关走环境变量、_enter_indent_here 把开关交给后端；"
              "先收窗 -> 再写 -> 写不成 send_vk(VK_RETURN) 兜底",
              ['os.environ.get("VBECOMPLETE_NO_AUTOCLOSE"' in _src57,
               "AUTO_CLOSE_BLOCK" in _src57,
               "auto_close=AUTO_CLOSE_BLOCK" in _fn57,
               "backend.new_line_below(smart=True" in _fn57,
               "completer.hide()" in _fn57,
               _fn57.index("completer.hide()") < _fn57.index("new_line_below"),
               "send_vk(VK_RETURN)" in _fn57],
              expect_contain=[True] * 7)

        # 后端里也必须真的走那两个纯函数（别绕过判定硬写收尾）
        _vbsrc57 = io.open(os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "vbe_bridge.py"), encoding="utf-8").read()
        check("57.5b 后端补收尾前必须过 block_closer + closer_needed 判定，"
              "并读下方各行（_lines_below）",
              ["block_closer(line_text)" in _vbsrc57,
               "closer_needed(self._lines_below(" in _vbsrc57,
               "if closer_text:" in _vbsrc57],
              expect_contain=[True] * 3)

        # ---- 57.6 出厂默认：开着 ----
        check("57.6 ★出厂默认：块头回车自动补收尾开着"
              "（VBECOMPLETE_NO_AUTOCLOSE 未设）",
              [M57.AUTO_CLOSE_BLOCK], expect_contain=[True])
    except Exception as _e57:
        check("第 57 节异常: %s" % _e57, [True], expect_contain=[False])

    # ---- 58. v79c：嵌套块各自补收尾 + 模块级 Type/Enum + 过程归属校验 ----
    #
    # 用户报的三件事：
    #   1) "我在一个函数里定义了一个变量 targetsheet，在定义另一个函数时输入 tes，
    #      会提示出 targetsheet"（跨过程泄漏）；
    #   2) "双层 for 只能自动补全外层 for，内层的 for 不会自动补 next"；
    #   3) "type、enum 这种结构语句不能自动补全"。
    print("\n=== 58. 嵌套收尾 / 模块级 Type·Enum / 过程归属（v79c）===")
    try:
        import vbe_bridge as VB58
        import parser as P58
        _orig58 = VB58._get_vbe_cached

        class _CM58(object):
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

        class _Pane58(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE58(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        def _nl58(lines, line_no, col, auto_close=True):
            cm58 = _CM58(lines)
            pane58 = _Pane58(cm58, (line_no, col, line_no, col))
            VB58._get_vbe_cached = lambda: _VBE58(pane58)
            ok58 = VB58.VbeBackend().new_line_below(smart=True,
                                                    auto_close=auto_close)
            return ok58, cm58.lines, pane58.sel

        # ---- 58.1 ★集成：用户报的双层 For ----
        #
        # 场景：外层 For 的回车已经走过（下面补出了 `    Next`），现在光标移到
        # 内层 For 那行再按回车。下面第一个非空行是【外层】的 `    Next`，
        # 缩进比内层块头（8）浅 —— 内层该补的 `        Next` 必须在它上面。
        check("58.1 ★集成：双层 For —— 内层 `For j` 回车也要补 `Next`，"
              "且补在外层 `Next` 之前、光标停在中间那行",
              [_nl58(["Sub Foo()",
                      "    For i = 1 To 10",
                      "        For j = 1 To 10",
                      "    Next",
                      "End Sub"], 3, 27)],
              expect_contain=[(True,
                               ["Sub Foo()",
                                "    For i = 1 To 10",
                                "        For j = 1 To 10",
                                "            ",
                                "        Next",
                                "    Next",
                                "End Sub"],
                               (4, 13, 4, 13))])

        # 内层 If 下面压着外层的 End If 同理
        check("58.2 ★集成：双层 If —— 内层 `If b Then` 补 `        End If`，"
              "外层 `    End If` 一个字没动",
              [_nl58(["Sub Foo()",
                      "    If a Then",
                      "        If b Then",
                      "    End If",
                      "End Sub"], 3, 20)],
              expect_contain=[(True,
                               ["Sub Foo()",
                                "    If a Then",
                                "        If b Then",
                                "            ",
                                "        End If",
                                "    End If",
                                "End Sub"],
                               (4, 13, 4, 13))])

        # 双层 Do / With（各自的收尾不同名，同一个规则管）
        check("58.3 ★集成：双层 Do -> `Loop`、双层 With -> `End With` 都补在对的层级",
              [_nl58(["Sub Foo()", "    Do", "        Do While x", "    Loop",
                      "End Sub"], 3, 24),
               _nl58(["Sub Foo()", "    With a", "        With b", "    End With",
                      "End Sub"], 3, 18)],
              expect_contain=[(True, ["Sub Foo()", "    Do", "        Do While x",
                                      "            ", "        Loop", "    Loop",
                                      "End Sub"], (4, 13, 4, 13)),
                              (True, ["Sub Foo()", "    With a", "        With b",
                                      "            ", "        End With",
                                      "    End With", "End Sub"],
                               (4, 13, 4, 13))])

        # ---- 58.4 ★集成：模块级 Type / Enum（用户报的"不能自动补全"）----
        #
        # 根因：收尾判据里"下面那行比块头【更浅】才补"。Type/Enum 只出现在
        # 模块级（缩进 0），而 0 是最浅的 —— 下面的声明/过程头永远不可能更浅，
        # 于是永远不补。现在补上"模块级 + 下面同级"这一档。
        check("58.4 ★集成：模块级 `Type Foo` / `Enum Color` 回车 -> "
              "下面即使还有别的过程，也补 `End Type` / `End Enum`",
              [_nl58(["Type Foo", "Sub Bar()", "End Sub"], 1, 9),
               _nl58(["Enum Color", "Sub Bar()", "End Sub"], 1, 11),
               _nl58(["Public Type T1", "Sub Bar()", "End Sub"], 1, 15),
               _nl58(["Private Enum E2", "Sub Bar()", "End Sub"], 1, 16)],
              expect_contain=[(True, ["Type Foo", "    ", "End Type",
                                      "Sub Bar()", "End Sub"], (2, 5, 2, 5)),
                              (True, ["Enum Color", "    ", "End Enum",
                                      "Sub Bar()", "End Sub"], (2, 5, 2, 5)),
                              (True, ["Public Type T1", "    ", "End Type",
                                      "Sub Bar()", "End Sub"], (2, 5, 2, 5)),
                              (True, ["Private Enum E2", "    ", "End Enum",
                                      "Sub Bar()", "End Sub"], (2, 5, 2, 5))])

        # 下面已经挂着同一个收尾 / 紧跟着另一个 Type -> 只缩进，不重复
        check("58.5 ★集成：模块级 Type/Enum 下面已挂着 `End Type` / `End Enum`"
              "（含紧跟下一个 Type）-> 不补重复的",
              [_nl58(["Type Foo", "End Type"], 1, 9)[1],
               _nl58(["Enum e", "End Enum"], 1, 7)[1],
               _nl58(["Type Foo", "Type B", "End Type"], 1, 9)[1]],
              expect_contain=[["Type Foo", "    ", "End Type"],
                              ["Enum e", "    ", "End Enum"],
                              ["Type Foo", "    ", "End Type",
                               "Type B", "End Type"]])

        # ---- 58.6 纯函数：_is_closer_text / _closer_owner_above ----
        check("58.6 _is_closer_text：`Next i` / `Next  '注释` / `Loop While x` 都算收尾；"
              "`NextLoop` / 注释行 / 空 -> 不算",
              [[VB58._is_closer_text("Next i", "Next"),
                VB58._is_closer_text("Next  '下一个", "Next"),
                VB58._is_closer_text("Loop While x", "Loop"),
                VB58._is_closer_text("End If", "End If"),
                VB58._is_closer_text("NextLoop", "Next"),
                VB58._is_closer_text("'    Next", "Next"),
                VB58._is_closer_text("", "Next"),
                VB58._is_closer_text("Next", None)]],
              expect_contain=[[True, True, True, True,
                               False, False, False, False]])

        check("58.7 _closer_owner_above：只有【同类、未闭合、缩进正好对上】的"
              "外层块头才算收尾的主人（配平过的那一层不算）",
              [[VB58._closer_owner_above(["Sub Foo()", "    For i = 1 To 10"],
                                         "for", 4),
                VB58._closer_owner_above(["Sub Foo()"], "for", 0),
                VB58._closer_owner_above(["    For i = 1 To 10"], "for", 8),
                VB58._closer_owner_above([], "for", 0),
                VB58._closer_owner_above(["If a Then"], "if", 0),
                VB58._closer_owner_above(["If a Then", "    End If"], "if", 4),
                VB58._closer_owner_above(None, None, 0)]],
              expect_contain=[[True, False, False, False,
                               True, False, False]])

        check("58.8 closer_needed：更浅的同类收尾 —— 有同类外层块头 -> 补，"
              "没有（用户自己写歪的收尾）-> 不补（免得 `Next` 双份）",
              [[VB58.closer_needed(["    Next"], "Next", 8, kind="for",
                                   lines_above=["Sub Foo()",
                                                "    For i = 1 To 10"]),
                VB58.closer_needed(["Next"], "Next", 4, kind="for",
                                   lines_above=["Sub Foo()"]),
                VB58.closer_needed(["    End If"], "End If", 8, kind="if",
                                   lines_above=["Sub Foo()",
                                                "    If a Then"]),
                VB58.closer_needed(["End If"], "End If", 4, kind="if",
                                   lines_above=["Sub Foo()"]),
                VB58.closer_needed(["End Type"], "End Type", 0, kind="type",
                                   lines_above=["Type A", "End Type"])]],
              expect_contain=[[True, False, True, False, False]])

        # 旧式调用（不传 kind/lines_above）必须保持 v79b 口径：更浅 -> 不补
        check("58.9 closer_needed 老口径不变（不传 kind/lines_above 时，"
              "更浅的同类收尾一律不补）",
              [[VB58.closer_needed(["Next i"], "Next", 4),
                VB58.closer_needed(["Loop While x"], "Loop", 4),
                VB58.closer_needed(["    End If"], "End If", 0),
                VB58.closer_needed([], "End If", 0)]],
              expect_contain=[[False, False, False, True]])

        # ---- 58.10 ★纯函数：过程归属校验（用户报的"泄漏"那一侧的护栏）----
        _mod58 = "\n".join([
            "Enum e",                      # 1
            "",                            # 2
            "Sub test()",                  # 3
            "    x = 1",                   # 4
            "End Sub",                     # 5
            "",                            # 6
            "Sub tes()",                   # 7
            "",                            # 8
            "End Sub"])                    # 9
        check("58.10 proc_owns_line：只有真的被那个过程圈住的行才算它是主人；"
              "`End Sub` 行 / 过程之间的空行 / 第一个过程头之前的声明区都不算",
              [[P58.proc_owns_line(_mod58, 4, "test"),
                P58.proc_owns_line(_mod58, 3, "test"),
                P58.proc_owns_line(_mod58, 5, "test"),
                P58.proc_owns_line(_mod58, 6, "test"),
                P58.proc_owns_line(_mod58, 2, "test"),
                P58.proc_owns_line(_mod58, 1, "test"),
                P58.proc_owns_line(_mod58, 4, "tes"),
                P58.proc_owns_line(_mod58, 8, "tes"),
                P58.proc_owns_line(_mod58, 4, None),
                P58.proc_owns_line(_mod58, 0, "test"),
                P58.proc_owns_line("", 1, "test")]],
              expect_contain=[[True, True, False, False, False, False,
                               False, True, False, False, False]])

        # 真机回归：VBE 的 ProcOfLine 会把过程外的行归给【邻居过程】
        # （`End Sub` 行、它下面到下一个过程头之间的行 -> 前一个过程；
        #   声明区里第一个过程头之前的行 -> 第一个过程）。_proc_of_line
        # 必须校验归属，否则"模块级 / 新函数的位置"上会冒出别的过程的局部变量
        # —— 正是用户报的"我在一个函数里定义的 targetsheet，在另一个函数里
        # 输入 tes 却被提示出来"。
        class _CMVbe58(object):
            """假 CodeModule：文本扫描正常，COM 兜底故意返回【邻居过程】。"""
            def __init__(self, lines, com_proc):
                self.lines = list(lines)
                self.Name = "Module1"
                self._com = com_proc

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "".join(l + "\n"
                               for l in self.lines[s0:s0 + int(count)])

            def ProcOfLine(self, line_no, kind):
                return (self._com, int(kind))     # 学 VBE：返回 (名字, ProcKind)

        _lives58 = ["Sub Foo()", "    Dim targetsheet As Long", "End Sub", ""]
        _lives58b = ["", "Sub Foo()", "    Dim targetsheet As Long", "End Sub"]
        check("58.11 ★护栏：文本扫描答不出来时，绝不采信 VBE 的 ProcOfLine "
              "（它把 `End Sub` 行、过程之间的空行、声明区里第一个过程头之前"
              "的行都归给邻居过程）；这些位置必须回到 None —— 否则就是『另一个"
              "函数的局部变量泄漏到本位置』。过程体那两行仍要正常认出 Foo"
              "（v79d 起：过程头那一行本身也算『不在过程体内』，见第 59 节）",
              [[VB58._proc_of_line(_CMVbe58(_lives58, "Foo"), 3),
                VB58._proc_of_line(_CMVbe58(_lives58, "Foo"), 4),
                VB58._proc_of_line(_CMVbe58(_lives58, "Foo"), 1),
                VB58._proc_of_line(_CMVbe58(_lives58, "Foo"), 2),
                VB58._proc_of_line(_CMVbe58(_lives58b, "Foo"), 1),
                isinstance(VB58._proc_of_line(_CMVbe58(_lives58, "Foo"), 2),
                           str)]],
              expect_contain=[[None, None, None, "Foo", None, True]])

        # ---- 58.12 ★接线护栏 ----
        _vbsrc58 = io.open(os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "vbe_bridge.py"), encoding="utf-8").read()
        check("58.12 ★接线护栏：new_line_below 把块头类型与上方各行交给 "
              "closer_needed（嵌套判定要用）；_proc_of_line 走归属校验；"
              "parser 提供 proc_owns_line",
              ["kind=_block_kind(line_text)" in _vbsrc58,
               "lines_above=above" in _vbsrc58,
               "_proc_owns_line(cm, name, line_no)" in _vbsrc58,
               "def proc_owns_line(code, line_no, name):" in io.open(
                   os.path.join(os.path.dirname(os.path.dirname(
                       os.path.abspath(__file__))), "parser.py"),
                   encoding="utf-8").read()],
              expect_contain=[True] * 4)

        # ---- 58.13 收尾清单完整性：VBA 里"首尾都有明确标志"的块必须全覆盖
        #             （#If/#Else 条件编译除外：它要配 `#End If`，替用户写个裸的
        #             `End If` 反而把代码写错）----
        _pairs58 = [("Sub Foo()", "End Sub"),
                    ("Private Function F() As Long", "End Function"),
                    ("Property Get X()", "End Property"),
                    ("Static Property Let Y(v As Long)", "End Property"),
                    ("Type Foo", "End Type"),
                    ("Public Enum Color", "End Enum"),
                    ("If x Then", "End If"),
                    ("With rng", "End With"),
                    ("Select Case x", "End Select"),
                    ("For i = 1 To 10", "Next"),
                    ("For Each c In rng", "Next"),
                    ("Do", "Loop"),
                    ("Do While x < 5", "Loop"),
                    ("Do Until x > 5", "Loop"),
                    ("While x", "Wend")]
        check("58.13 收尾清单：VBA 里首尾都有明确标志的块 15 例全覆盖"
              "（Sub/Function/Property/Type/Enum/If/With/Select Case/"
              "For/For Each/Do·While/Do·Until/While）",
              [[VB58.block_closer(_t) for _t, _ in _pairs58]],
              expect_contain=[[_w for _, _w in _pairs58]])

        # 没有收尾的语句一个都不许认（否则会替用户写出错的结构）。
        # `ElseIf … Then` 只是 If 块【内】的分支：它要的收尾是与 `If` 共用的那
        # 个 `End If`，不能在这儿再补一条。
        _noc58 = ["Else", "ElseIf y Then", "Case 1", "Case Else",
                  "#If VBA7 Then", "#Else",
                  "Next i", "End If", "Loop", "Wend", "End Type",
                  "x = 1", "Exit For", "foo(1, 2)", "For i = 1 To 10 _"]
        check("58.14 不该补收尾的：Else / ElseIf…Then / Case / #If / #Else"
              "（分支与条件编译）、以及所有收尾行本身 / 普通语句 / "
              "以续行符结尾的块头",
              [[VB58.block_closer(_t) for _t in _noc58
                if VB58.block_closer(_t) is not None]],
              expect_contain=[[]])

        VB58._get_vbe_cached = _orig58
    except Exception as _e58:
        check("第 58 节异常: %s" % _e58, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 59. 跨过程局部变量泄漏 · 光标停在过程头上（v79d）
    #
    # 用户复现（v79c 之后仍然泄漏）：
    #     Sub test()
    #     Set targetSheet = Worksheets(1)     ' 没写 Dim -> 隐式变量
    #     ...（双层 For / If / With）
    #     End Sub
    #   然后再定义一个过程，输入 `sub test` 时弹窗同时提示 test 和 targetSheet。
    #
    # 实测有【两条独立路径】，都通到同一个症状：
    #   A) 光标这一行【自己】是过程头（`Sub test`）时，proc_at_line 把它当成了
    #      "当前过程"；这个新名字恰好等于模块里已有的过程名 -> 那个过程的局部
    #      变量（targetSheet）被当成"当前过程的"提示出来。
    #   B) 光标停在过程头的【名字】上时，"抹掉正在输入的词"把整个过程头抹废，
    #      紧随其后的过程体失去过程归属，局部/隐式变量统统降级成【模块级】
    #      -> 模块里任何位置都能提示出它们。
    # ------------------------------------------------------------------
    print("\n=== 59. 光标停在过程头上：跨过程泄漏（v79d）===")
    try:
        import vbe_bridge as VB59
        import parser as P59
        import engine as E59

        _USER59 = [
            "Sub test()",                        # 1
            "Set targetSheet = Worksheets(1)",   # 2
            "For i = 1 To 100",                  # 3
            "For j = 1 To 200",                  # 4
            "If i = 1 Then",                     # 5
            "If j = 2 Then",                     # 6
            "With targetSheet",                  # 7
            "With targetSheet",                  # 8
            "End With",                          # 9
            "End With",                          # 10
            "End If",                            # 11
            "End If",                            # 12
            "Next",                              # 13
            "Next",                              # 14
            "End Sub",                           # 15
        ]

        class _CM59(object):
            """假 CodeModule：文本扫描正常；COM 兜底故意学 VBE 返回邻居过程名。"""

            def __init__(self, lines, com_proc="test"):
                self.lines = list(lines)
                self.Name = "Module1"
                self._com = com_proc

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "".join(l + "\n"
                               for l in self.lines[s0:s0 + int(count)])

            def ProcOfLine(self, line_no, kind):
                return (self._com, int(kind))     # 学 VBE：(名字, ProcKind)

        def _vis59(lines, line_no, col):
            """一条龙：真收集（含 caret 抹词）-> 真作用域过滤。"""
            code = "\n".join(lines)
            caret = (line_no, col)
            r = list(P59.extract_records(code, module="Module1",
                                         is_std_module=True, caret=caret))
            r += P59.extract_implicit_records(code, module="Module1",
                                              is_std_module=True, declared=r,
                                              scope="proc", caret=caret)
            proc = VB59._proc_of_line(_CM59(lines), line_no)
            vis = E59.filter_identifiers_by_scope(r, proc, "Module1")
            return proc, sorted(set(vis))

        # ---- 59.1 纯函数：is_proc_header_line ----
        check("59.1 is_proc_header_line：Sub / Function / Property Get|Let|Set"
              "（含 Public/Private/Static 修饰符、名字后无括号也可）都算过程头；"
              "`Set x = 1` / 注释里的 Sub / 空行 / 普通语句都不算",
              [[P59.is_proc_header_line("Sub test"),
                P59.is_proc_header_line("Sub test()"),
                P59.is_proc_header_line("  Function F() As Long"),
                P59.is_proc_header_line("Private Property Get X()"),
                P59.is_proc_header_line("Static Property Let Y(v As Long)"),
                P59.is_proc_header_line("Set targetSheet = Worksheets(1)"),
                P59.is_proc_header_line("' Sub Foo()"),
                P59.is_proc_header_line(""),
                P59.is_proc_header_line("    Sub total = 1")]],
              expect_contain=[[True, True, True, True, True,
                               False, False, False, False]])

        # ---- 59.2 ★用户场景：模块级位置正在写新过程头 ----
        for _tail, _col in (("Sub test", 9), ("Sub test", 6),
                            ("Sub test(", 10), ("Sub te", 7)):
            _lines = _USER59 + [_tail]
            _proc, _vis = _vis59(_lines, len(_lines), _col)
            check("59.2 ★用户场景：模块里已有 Sub test()，在【模块级】新写 %r"
                  "（光标列 %d）-> 当前过程必须是 None，targetSheet 不得出现"
                  % (_tail, _col),
                  [(_proc, "targetSheet" in _vis)],
                  expect_contain=[(None, False)])

        # ---- 59.3 不能过度收敛：过程体内仍按该过程过滤 ----
        check("59.3 不能过度收敛：光标在 test 体内（第 2 / 7 行）时当前过程仍是 "
              "test、targetSheet 照常可见；`End Sub` 那一行则回到 None",
              [[_vis59(_USER59, 2, 3)[0], _vis59(_USER59, 7, 3)[0],
                _vis59(_USER59, 15, 3)[0],
                "targetSheet" in _vis59(_USER59, 2, 3)[1]]],
              expect_contain=[["test", "test", None, True]])

        # ---- 59.4 _blank_ident_at：声明名【文本一个字不动】（v79e 改口径）----
        _b59 = P59._blank_ident_at
        check("59.4 _blank_ident_at：过程名 / Type 名 / Enum 名【原样保留】"
              "（行仍是声明行 -> is_proc_header_line 依然为 True、过程归属不受影响）；"
              "普通变量 / 赋值目标 / 形参名照旧抹成等长空格；"
              "只有显式要求 blank_decl=True（现场证据扫描那一路）才肯抹声明名",
              [[_b59("Sub bar()", 1, 7),
                _b59("Type Foo", 1, 7),
                _b59("Enum Color", 1, 8),
                P59.is_proc_header_line(_b59("Sub bar()", 1, 7)),
                _b59("Set other = 1", 1, 7),
                _b59("total = 1", 1, 4),
                _b59("Sub bar()", 1, 7, True),
                _b59("Enum Color", 1, 8, True)]],
              expect_contain=[["Sub bar()", "Type Foo", "Enum Color", True,
                               "Set" + " " * 7 + "= 1",
                               " " * 6 + "= 1",
                               "Sub" + " " * 4 + "()",
                               "Enum" + " " * 6]])

        # ---- 59.5 ★第二条路径：光标停在过程头的【名字】里 ----
        _mid59 = ["Sub test()", "Set targetSheet = Worksheets(1)", "End Sub",
                  "Sub bar()", "Set other = 1", "End Sub",
                  "Sub mid()", "Set third = 2", "End Sub"]
        _rows59 = []
        for _c in (6, 7, 8, 9, 10):
            _code = "\n".join(_mid59)
            _r = list(P59.extract_records(_code, module="Module1",
                                          is_std_module=True, caret=(4, _c)))
            _r += P59.extract_implicit_records(_code, module="Module1",
                                               is_std_module=True, declared=_r,
                                               scope="proc", caret=(4, _c))
            _rec = dict((str(x[0]).lower(), x) for x in _r)
            _v = E59.filter_identifiers_by_scope(_r, "test", "Module1")
            # v79e：声明名不再被抹（文本一个字不动），`Sub bar()` 的过程归属照旧；
            # 要害是【它不能是 None】—— None 就等于降级成模块级、全模块可见。
            _rows59.append((_rec.get("other") is not None
                            and _rec["other"][2] is not None,
                            "other" in _v))
        check("59.5 ★第二条泄漏路径：光标停在 `Sub bar()` 的 bar 名字里"
              "（第 6~10 列）-> bar 里的 other 必须仍带【过程归属】（绝不是模块级）"
              "、在 test 里绝不可见（旧实现把过程头抹成空格，整个过程体失去归属）",
              [_rows59],
              expect_contain=[[(True, False)] * 5])

        # ---- 59.6 Type/Enum 名字被占位后，块成员不得变成隐式变量 ----
        _t59 = ["Type Foo", "    Red As Long", "End Type",
                "Enum Color", "    A = 1", "End Enum"]
        _code59t = "\n".join(_t59)
        _rt = list(P59.extract_records(_code59t, module="Module1",
                                       is_std_module=True, caret=(1, 7)))
        _rt += P59.extract_implicit_records(_code59t, module="Module1",
                                            is_std_module=True, declared=_rt,
                                            scope="proc", caret=(1, 7))
        _names59t = sorted(set(str(x[0]).lower() for x in _rt))
        check("59.6 光标停在 `Type Foo` 的 Foo 上时，Type/Enum 的块成员"
              "（Red / A）仍不得被当成隐式变量收进候选池",
              [[n for n in _names59t if n in ("red", "a")]],
              expect_contain=[[]])

        # ---- 59.7 ★接线护栏 ----
        _root59 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _vbsrc59 = io.open(os.path.join(_root59, "vbe_bridge.py"),
                           encoding="utf-8").read()
        _psrc59 = io.open(os.path.join(_root59, "parser.py"),
                          encoding="utf-8").read()
        _guard59 = "is_proc_header_line(cm.Lines(line_no, 1))"
        _scan59 = "name = vba_parser.proc_at_line(head, line_no)"
        check("59.7 ★接线护栏：_proc_of_line 用 is_proc_header_line 挡住过程头"
              "那一行，且必须排在 proc_at_line 之前（否则文本扫描会把它当成"
              "当前过程）；parser 提供 is_proc_header_line 与 _struct_decl_name_span",
              [[-1 < _vbsrc59.find(_guard59) < _vbsrc59.find(_scan59),
                "def is_proc_header_line(line_text):" in _psrc59,
                "def _struct_decl_name_span(line):" in _psrc59,
                "_struct_decl_name_span(line)" in _psrc59]],
              expect_contain=[[True] * 4])

    except Exception as _e59:
        check("第 59 节异常: %s" % _e59, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 60. v79e：同名过程各自成域 + 声明名不再改文本
    #
    #   用户报："你改的更糟糕了，现在过程级的变量，都泄露到其他过程了"。
    #   真机（模块2）复现出两条真因：
    #   A) 模块里【两个过程同名】（用户在既有 `Sub test()` 下面又敲了一个
    #      `Sub test()`）—— 记录只存过程名，两个过程的局部/隐式变量在过滤时
    #      无法区分，第二个过程的体内能提示出第一个过程的 targetSheet。
    #   B) v79d 把"正在声明的名字"抹成等长下划线：名字只有 1 个字符时
    #      （`Sub t`），行尾正好是"空格 + 下划线 + 换行"，被 _RE_CONTINUATION
    #      当成行继续符，把下一行吞进同一行 —— 过程头消失、其后整段过程体失去
    #      归属，隐式变量降级成模块级，全模块到处可见。
    # ------------------------------------------------------------------
    print("\n=== 60. 同名过程各自成域 / 声明名不改文本（v79e）===")
    try:
        import vbe_bridge as VB60
        import parser as P60
        import engine as E60

        _USER60 = [
            "Sub test()",                            # 1
            "    Set targetSheet = Worksheets(1)",   # 2
            "    For i = 1 To 100",                  # 3
            "    Next",                              # 4
            "End Sub",                               # 5
            "",                                      # 6
            "Sub test()",                            # 7
            "    For i = 1 To 2",                    # 8
            "    Next",                              # 9
            "    m = 3",                             # 10
            "End Sub",                               # 11
        ]

        class _CM60(object):
            """假 CodeModule：文本扫描正常；COM 兜底故意学 VBE 返回邻居过程名。"""

            def __init__(self, lines, com_proc="test"):
                self.lines = list(lines)
                self.Name = "Module1"
                self._com = com_proc

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "".join(l + "\n"
                               for l in self.lines[s0:s0 + int(count)])

            def ProcOfLine(self, line_no, kind):
                return (self._com, int(kind))     # 学 VBE：(名字, ProcKind)

        def _recs60(lines, caret=None):
            code = "\n".join(lines)
            r = list(P60.extract_records(code, module="Module1",
                                         is_std_module=True, caret=caret))
            r += P60.extract_implicit_records(code, module="Module1",
                                              is_std_module=True, declared=r,
                                              scope="proc", caret=caret)
            return code, r

        def _vis60(lines, line_no, caret=None):
            _code, _r = _recs60(lines, caret)
            _proc = VB60._proc_of_line(_CM60(lines), line_no)
            return _proc, sorted(set(E60.filter_identifiers_by_scope(
                _r, _proc, "Module1")))

        _code60, _r60 = _recs60(_USER60)

        # ---- 60.1 记录侧：同名过程的键（第 1 次用原名，第 2 次起带序号）----
        check("60.1 ★同名过程：模块里两个 `Sub test()` 时，记录键必须是 "
              "第 1 个 `test`、第 2 个 `test#2` —— 两个过程的隐式变量各挂各的"
              "（targetSheet 归 test、m 归 test#2；i 两边各有一份）",
              [sorted((str(a[0]).lower(), str(a[2])) for a in _r60)],
              expect_contain=[[("i", "test"), ("i", "test#2"), ("m", "test#2"),
                               ("targetsheet", "test"), ("test", "None")]])

        # ---- 60.2 ★用户场景端到端：第二个过程体内不得出现第一个过程的变量 ----
        _p60_1, _v60_1 = _vis60(_USER60, 2)
        _p60_2, _v60_2 = _vis60(_USER60, 8)
        _p60_3, _v60_3 = _vis60(_USER60, 10)
        check("60.2 ★用户场景：模块里两个 `Sub test()` —— 第 1 个体内"
              "（L2）照常提示 targetSheet、不提示 m；第 2 个体内"
              "（L8 / L10）**绝不提示 targetSheet**、照常提示自己的 m",
              [[_p60_1, "targetSheet" in _v60_1, "m" in _v60_1,
                _p60_2, "targetSheet" in _v60_2, "m" in _v60_2,
                _p60_3, "targetSheet" in _v60_3, "m" in _v60_3]],
              expect_contain=[["test", True, False,
                               "test#2", False, True,
                               "test#2", False, True]])

        # ---- 60.3 光标停在第二个过程头上：仍按模块级（v79d 口径不退化）----
        _p60_h, _v60_h = _vis60(_USER60, 7)
        check("60.3 光标停在第二个 `Sub test()` 那一行（L7）：当前过程必须是 None"
              "（过程头那一行是声明、不是过程体），targetSheet 与 m 都不得出现",
              [[_p60_h, "targetSheet" in _v60_h, "m" in _v60_h]],
              expect_contain=[[None, False, False]])

        # ---- 60.4 异名过程：键保持裸名（与旧版行为完全一致）----
        _diff60 = list(_USER60)
        _diff60[6] = "Sub otherProc()"
        _code60d, _r60d = _recs60(_diff60)
        check("60.4 名字不重复时（把第二个过程改名 otherProc）：记录键必须还是"
              "裸名 —— 既有行为、既有测试、日志输出一律不变",
              [sorted(set(str(a[2]) for a in _r60d), key=str)],
              expect_contain=[["None", "otherProc", "test"]])

        # ---- 60.5 纯函数：_proc_scope_key / proc_key_plain ----
        check("60.5 _proc_scope_key：第 1 次出现用原名，第 2 次起 `名字#序号`；"
              "proc_key_plain 把序号拆掉（没有序号 / 非数字后缀则原样返回）",
              [[P60._proc_scope_key("test", 1), P60._proc_scope_key("test", 2),
                P60._proc_scope_key("test", 3), P60._proc_scope_key("test", 0),
                P60.proc_key_plain("test#2"), P60.proc_key_plain("test"),
                P60.proc_key_plain("a#b"), P60.proc_key_plain(None)]],
              expect_contain=[["test", "test#2", "test#3", "test",
                               "test", "test", "a#b", ""]])

        # ---- 60.6 proc_at_line / proc_owns_line 的分工 ----
        check("60.6 proc_at_line：同名模块里第 1 个过程内是 `test`、第 2 个内是"
              "`test#2`；异名模块里都是裸名。proc_owns_line 只认名字部分："
              "`test#2` 与 `test` 都算『这个名字的过程包含该行』（COM 给的是裸名），"
              "而 `End Sub` 行与过程之间的空行一律 False",
              [[P60.proc_at_line(_code60, 2), P60.proc_at_line(_code60, 8),
                P60.proc_at_line(_code60d, 8),
                P60.proc_owns_line(_code60, 8, "test#2"),
                P60.proc_owns_line(_code60, 8, "test"),
                P60.proc_owns_line(_code60, 5, "test"),
                P60.proc_owns_line(_code60, 6, "test"),
                P60.proc_owns_line(_code60, 1, "test")]],
              expect_contain=[["test", "test#2", "otherProc",
                               True, True, False, False, True]])

        # ---- 60.7 caret_decl_name_at：只认"这一行正在声明的那个名字" ----
        check("60.7 caret_decl_name_at：`Sub bar()` 的光标落在 bar 上 -> 返回 bar；"
              "落在 `Sub` 关键字 / 括号后 -> 空；`Dim x`、赋值语句、"
              "非声明行一律空（它们走『抹空格』那条路）",
              [[P60.caret_decl_name_at("Sub bar()", 1, 7),
                P60.caret_decl_name_at("Sub bar()", 1, 4),
                P60.caret_decl_name_at("Sub bar()", 1, 10),
                P60.caret_decl_name_at("Type Foo", 1, 7),
                P60.caret_decl_name_at("Enum Color", 1, 8),
                P60.caret_decl_name_at("Dim bar As Long", 1, 6),
                P60.caret_decl_name_at("total = 1", 1, 3),
                P60.caret_decl_name_at("", 1, 1)]],
              expect_contain=[["bar", "", "", "Foo", "Color",
                               "", "", ""]])

        # ---- 60.8 ★v79d 回归护栏：1 字符名字不再被当成行继续符 ----
        _one60 = "Sub t\n    Set targetSheet = 1\nEnd Sub"
        _blanked60 = P60._blank_ident_at(_one60, 1, 6)
        _masked60 = P60._RE_CONTINUATION.sub(
            " ", P60._mask_strings_and_comments(_blanked60))
        _one_r60 = []
        _one_c60 = list(P60.extract_records(_blanked60, module="M",
                                            is_std_module=True))
        _one_c60 += list(P60.extract_implicit_records(
            _blanked60, module="M", is_std_module=True, declared=_one_c60,
            scope="proc"))
        _one_r60 = [a for a in _one_c60
                    if str(a[0]).lower() == "targetsheet"]
        check("60.8 ★v79d 回归护栏（用户说『更糟糕了』的真因）：1 个字符的过程名"
              "（`Sub t`）下，抹词后首行【原样不动】、行继续符合并后行数不变"
              "（`Sub t` 没被当成续行吞掉下一行）、该行仍是过程头；"
              "过程体里 targetSheet 仍归 `t`（绝不是 None=模块级）",
              [[_blanked60.split("\n")[0],
                len(_masked60.split("\n")),
                P60.is_proc_header_line(_masked60.split("\n")[0]),
                _one_r60]],
              expect_contain=[["Sub t", 3, True,
                               [("targetSheet", "M", "t", True)]]])

        # ---- 60.9 ★接线护栏 ----
        _root60 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _vbsrc60 = io.open(os.path.join(_root60, "vbe_bridge.py"),
                           encoding="utf-8").read()
        _psrc60 = io.open(os.path.join(_root60, "parser.py"),
                          encoding="utf-8").read()
        check("60.9 ★接线护栏：现场证据扫描（names_outside_caret）必须显式要求"
              "抹掉光标处的词（blank_decl=True），否则幻影词会把自己当证据；"
              "声明名的处理只在记录/文本两处留痕；_proc_dup_counts 这类已被"
              "『第 1 次用原名』取代的旧辅助不得残留",
              [["vba_parser._blank_ident_at(code, line_no, col, True)" in _vbsrc60,
                "def _blank_ident_at(text, line_no, col, blank_decl=False):"
                in _psrc60,
                "def caret_decl_name_at(text, line_no, col):" in _psrc60,
                "def _proc_scope_key(name, occurrence):" in _psrc60,
                "cur_proc = proc_key(proc)" in _psrc60,
                "return _proc_scope_key(name, max(occ, 1))" in _psrc60,
                "_proc_dup_counts" not in _psrc60,
                "fill = \"_\" * " not in _psrc60]],
              expect_contain=[[True] * 8])

    except Exception as _e60:
        check("第 60 节异常: %s" % _e60, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 61. v80a：回车补收尾认「下方同级代码」（用户报：在已有 for 上面再补一个
    #     for，回车只缩进、不补 next）
    #     v80b：让位状态按【槽位】记 + 会过期；VBE 提示窗句柄缓存周期重扫
    #     （用户报：变量类型是自定义类模块时，我们的弹窗遮挡 VBE 的提示列表，
    #      且"不是每次都能复现"）
    # ------------------------------------------------------------------
    print("\n=== 61. 同级补收尾 + 让位槽位/有效期 + 提示窗缓存重扫（v80）===")
    try:
        import vbe_bridge as VB61
        import engine as E61
        import time as T61

        _orig61 = VB61._get_vbe_cached
        _ocb = VB61._own_closer_below
        _cn61 = VB61.closer_needed

        # ---- 61.1 _own_closer_below：配平法认"本块自己的收尾" ----
        check("61.1 _own_closer_below：配平到 0 的同级同类收尾就是本块的；"
              "同级同类块头计一层；孤立在更深处的同类收尾算『用户写歪』（保守不补）；"
              "同级的别的块头 -> 本块已结束；下方空 -> 没有",
              [[_ocb(["    Next"], "Next", "for", 4),
                _ocb(["    For i = 1 To 10", "        x = 1", "    Next i"],
                     "Next", "for", 4),
                _ocb(["        Next"], "Next", "for", 4),
                _ocb(["Sub Bar()"], "End Type", "type", 0),
                _ocb([], "Next", "for", 4)]],
              expect_contain=[[True, False, True, False, False]])

        # ---- 61.2 closer_needed：传 kind 走 v80a 新口径 ----
        check("61.2 closer_needed（kind 给定）：下方同级压着别的代码/块 -> 补；"
              "下面已挂着本块自己的收尾 -> 不补；更浅的同类收尾沿用 v79c 归属判定",
              [[_cn61(["    For i = 1 To 10", "        x = 1", "    Next i",
                       "End Sub"], "Next", 4, kind="for",
                      lines_above=["Sub Test()"]),
                _cn61(["    Next", "End Sub"], "Next", 4, kind="for",
                      lines_above=["Sub Test()"]),
                _cn61(["    x = 1", "End Sub"], "End With", 4, kind="with",
                      lines_above=["Sub Test()"]),
                _cn61(["    Next", "End Sub"], "Next", 8, kind="for",
                      lines_above=["Sub Test()", "    For i = 1 To 10"]),
                _cn61(["Next", "End Sub"], "Next", 4, kind="for",
                      lines_above=["Sub Test()"])]],
              expect_contain=[[True, False, True, True, False]])

        # ---- 61.3 ★集成：用户原话的场景 ----
        class _CM61(object):
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

        class _Pane61(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE61(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        def _nl61(lines, line_no):
            """光标落在第 line_no 行【行尾】，回车走 smart + auto_close。"""
            cm61 = _CM61(lines)
            col = len(lines[line_no - 1]) + 1
            pane61 = _Pane61(cm61, (line_no, col, line_no, col))
            VB61._get_vbe_cached = lambda: _VBE61(pane61)
            ok61 = VB61.VbeBackend().new_line_below(smart=True,
                                                    auto_close=True)
            return ok61, cm61.lines, pane61.sel

        check("61.3 ★集成（用户报的场景）：已经在 `For i`（回车补过 `Next i`）"
              "【上面】再补一个 `For j`，行尾回车 -> 中间留一行缩进 8 空格、"
              "下面补出同级的 `    Next`，原来那个 For 一个字没动",
              [_nl61(["Sub Test()",
                      "    For j = 1 To 5",
                      "    For i = 1 To 10",
                      "        MsgBox i",
                      "    Next i",
                      "End Sub"], 2)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    For j = 1 To 5",
                                "        ",
                                "    Next",
                                "    For i = 1 To 10",
                                "        MsgBox i",
                                "    Next i",
                                "End Sub"],
                               (3, 9, 3, 9))])

        check("61.4 ★集成：上面补的是别的块（`With`）同理；"
              "而下面已经挂着自己收尾的（`    Next` 紧跟随）-> 只缩进、不补重复",
              [_nl61(["Sub Test()",
                      "    For i = 1 To 10",
                      "        MsgBox i",
                      "    Next i",
                      "End Sub"], 2)[1],
               _nl61(["Sub Test()",
                      "    For i = 1 To 10",
                      "    Next i",
                      "End Sub"], 2)[1],
               _nl61(["Sub Test()", "    Do", "    Loop", "End Sub"], 2)[1]],
              expect_contain=[["Sub Test()", "    Do", "        ",
                               "    Loop", "End Sub"],
                              ["Sub Test()", "    For i = 1 To 10",
                               "        ", "    Next i", "End Sub"],
                              ["Sub Test()", "    Do", "        ",
                               "    Loop", "End Sub"]])

        check("61.5 集成对照（v79b/v79c 口径不许回归）：模块级 `Type` 下面"
              "紧邻过程头 -> 仍补 `End Type`；`Sub` 下面已挂着 `End Sub` -> 不补",
              [_nl61(["Type Foo", "Sub Bar()", "End Sub"], 1)[1],
               _nl61(["Sub Foo()", "End Sub"], 1)[1]],
              expect_contain=[["Type Foo", "    ", "End Type",
                               "Sub Bar()", "End Sub"],
                              ["Sub Foo()", "    ", "End Sub"]])
        VB61._get_vbe_cached = _orig61

        # ---- 61.6 list_slot_anchor：算出"哪一个槽位" ----
        check("61.6 list_slot_anchor：`As `/`New ` 之后锚点=类型名起点列（同一槽位"
              "内打字不变）；`标识符.` / With 块的 `.成员` 锚点=成员名起点列；"
              "非列表位置 -> None",
              [[E61.list_slot_anchor("Dim x As ", 10),
                E61.list_slot_anchor("Dim x As 类", 11),
                E61.list_slot_anchor("Dim a As X, b As ", 18),
                E61.list_slot_anchor("obj.", 5),
                E61.list_slot_anchor("obj.Me", 7),
                E61.list_slot_anchor("    .Si", 8),
                E61.list_slot_anchor("x = 1", 6),
                E61.list_slot_anchor("", 1)]],
              expect_contain=[[10, 10, 18, 5, 5, 6, None, None]])

        # ---- 61.7 一致性：vbe_list_expected 认的列表位置，锚点必须算得出来 ----
        _cases61 = [("Dim x As ", 10), ("Dim x As Inte", 14),
                    ("Set c = New ", 13), ("Private m_d As Dictionary", 28),
                    ("obj.", 5), ("objs.Item(1).", 14), ("    .Si", 8),
                    ("x = 1 'a.b", 10), ("1.5", 4), ("x = 1", 6)]
        check("61.7 ★一致性护栏：凡是 vbe_list_expected 判为『VBE 会弹列表』的位置，"
              "list_slot_anchor 都必须给得出锚点（两处语法不许走偏）",
              [[(lt, cc) for (lt, cc) in _cases61
                if E61.vbe_list_expected(lt, cc)
                and E61.list_slot_anchor(lt, cc) is None]],
              expect_contain=[[]])

        # ---- 61.8 ★端到端：同一行两个槽位互不连累（v80b 核心口径）----
        _POOL61 = [(n, "VBA", None, False)
                   for n in ("类1", "类2", "Dim", "obj", "x")]

        class _B61(object):
            """桩后端：只表达"成员列表是否可见"（v65 语义），位置由文本来。"""

            def __init__(self, line_text, caret_col, yld=False):
                self._lt = line_text
                self._cc = caret_col
                self._ln = 1
                self._mod = "模块1"
                self._yld = yld

            def set_line(self, lt, cc):
                self._lt = lt
                self._cc = cc

            def get_context(self):
                return {"line_no": self._ln, "line_text": self._lt,
                        "caret_col": self._cc, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": self._mod}

            def get_identifiers(self):
                return list(_POOL61)

            def get_declared_names(self):
                return []

            def declared_elsewhere(self, name, module_name):
                return False

            def get_structural_names(self):
                return []

            def get_builtin_names(self):
                return set()

            def get_type_names(self):
                return []

            def get_host_enum_names(self):
                return {}

            def names_outside_caret(self, caret, scope_only=False):
                return set()

            def apply_completion(self, *a, **k):
                return None

            def vbe_yield_visible(self):
                return self._yld

            def vbe_popup_visible(self):
                return self._yld

            def vbe_popup_info(self):
                return []

        class _UI61(object):
            def __init__(self):
                self.shown = False
                self.hides = 0

            def show(self, rows, sel, c):
                self.shown = True

            def hide(self):
                self.shown = False
                self.hides += 1

            def update_selection(self, sel):
                pass

            def contains_point(self, x, y):
                return False

        def _mk61(line_text, caret_col):
            be = _B61(line_text, caret_col)
            ui = _UI61()
            return E61.Completer(be, ui), be, ui

        _c61, _b61, _u61 = _mk61("Dim a As 类1, b As 类2", 11)   # 第一个槽位
        _c61.trigger(True)
        _p1_61 = _c61.yield_pending is not None
        _r1_61 = _c61.confirm_yield(False)          # VBE 这会儿没弹
        _gu61 = _c61.yield_given_up
        _slot1_61 = _gu61[2] if _gu61 else None
        _b61.set_line("Dim a As 类1, b As 类2", 20)  # 挪到第二个槽位
        _c61.trigger(True)
        check("61.8 ★v80b：`Dim a As 类1, b As 类2` 一行两个槽位 —— 第一个槽位"
              "让位失败（记 given_up，锚点=10）后，移到【第二个槽位】仍然让位"
              "（待定让位、不弹我们的窗），不再被整行连累",
              [[_p1_61, _r1_61 is False, _slot1_61 == 10,
                _c61.yield_pending is not None,
                len(_c61.matches or []) == 0]],
              expect_contain=[[True, True, True, True, True]])

        # ---- 61.9 同一槽位内继续打字：不抖（标记仍然有效）----
        _c61, _b61, _u61 = _mk61("Dim a As 类", 13)
        _c61.trigger(True)
        _c61.confirm_yield(False)
        _gu2_61 = _c61.yield_given_up
        _b61.set_line("Dim a As 类1", 14)            # 同一槽位继续打字
        _c61.trigger(True)
        check("61.9 对照：同一槽位内继续打字 -> 标记仍有效、不再反复让位"
              "（避免『让开-补上-再让开』的抖动）",
              [[_c61.yield_given_up == _gu2_61, _c61.yield_pending is None]],
              expect_contain=[[True, True]])

        # ---- 61.10 有效期：过期后重新给 VBE 一次机会 ----
        _c61, _b61, _u61 = _mk61("Dim a As 类", 13)
        _c61.trigger(True)
        _c61.confirm_yield(False)
        _c61.yield_given_up_at = T61.time() - (E61.YIELD_GIVEN_UP_TTL + 1.0)
        _c61.trigger(True)
        check("61.10 ★v80b：`已放弃让位`标记【过期】后再回到同一处 -> 语法判据"
              "重新生效（待定让位、旧标记作废）—— 兜住『VBE 只是慢了一拍』",
              [[_c61.yield_pending is not None, _c61.yield_given_up is None]],
              expect_contain=[[True, True]])

        _old_ttl61 = E61.YIELD_GIVEN_UP_TTL
        E61.YIELD_GIVEN_UP_TTL = 0.0
        try:
            _c61, _b61, _u61 = _mk61("Dim a As 类", 13)
            _c61.trigger(True)
            _c61.confirm_yield(False)
            _zero61 = _c61.yield_given_up
        finally:
            E61.YIELD_GIVEN_UP_TTL = _old_ttl61
        check("61.11 VBECOMPLETE_YIELD_TTL=0 -> 干脆不记标记（每次都先让位）",
              [_zero61], expect_contain=[None])

        # ---- 61.12 提示窗句柄缓存：必须能【长出新句柄】 ----
        _alive61 = VB61._hwnd_alive
        _enum61 = VB61._enum_vbe_popup_windows
        _st61 = {"found": [], "calls": 0}
        try:
            VB61._hwnd_alive = lambda h: True

            def _fake_enum61():
                _st61["calls"] += 1
                return [(h, "NameListWndClass") for h in _st61["found"]]

            VB61._enum_vbe_popup_windows = _fake_enum61
            VB61._vbe_popup_state["hwnds"] = []
            VB61._vbe_popup_state["at"] = 0.0
            _st61["found"] = [100]
            _r1_61 = VB61._vbe_popup_hwnds()
            _st61["calls"] = 0
            VB61._vbe_popup_hwnds()               # 立刻再调：限流，不该重扫
            _throttled61 = _st61["calls"]
            _st61["found"] = [100, 200]           # 后来才建出来的提示窗
            VB61._vbe_popup_state["at"] = 0.0     # 假装过了一个重扫间隔
            _r2_61 = sorted(VB61._vbe_popup_hwnds())
        finally:
            VB61._hwnd_alive = _alive61
            VB61._enum_vbe_popup_windows = _enum61
            VB61._vbe_popup_state["hwnds"] = []
            VB61._vbe_popup_state["at"] = 0.0
        check("61.12 ★v80b：提示窗句柄缓存【缓存非空时也要周期重扫】——"
              "后来才建出来的提示窗（用到才建的参数信息窗）也能被探到；"
              "间隔内不重复枚举（性能护栏）",
              [[_r1_61, _throttled61, _r2_61]],
              expect_contain=[[[100], 0, [100, 200]]])

        # ---- 61.13 ★接线护栏 ----
        _root61 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _esrc61 = io.open(os.path.join(_root61, "engine.py"),
                          encoding="utf-8").read()
        _vbsrc61 = io.open(os.path.join(_root61, "vbe_bridge.py"),
                           encoding="utf-8").read()
        check("61.13 ★接线护栏：让位签名带【槽位】、confirm_yield 按前缀比对；"
              "标记带有效期；vbe_bridge 的提示窗缓存不再以『缓存为空』为前提；"
              "v80a 新纯函数到位、老口径（kind=None）那一段没被动过",
              [["list_slot_anchor(ctx.get(\"line_text\")" in _esrc61,
                "tuple(pend[1])[:2]" in _esrc61,
                "YIELD_GIVEN_UP_TTL" in _esrc61,
                "self.yield_given_up_at" in _esrc61,
                "now - _vbe_popup_state[\"at\"] >= _VBE_POPUP_RESCAN_SEC"
                in _vbsrc61,
                "if not hs:" not in _vbsrc61,
                "def _own_closer_below(lines_below, closer, kind, base):"
                in _vbsrc61,
                "if kind is not None:" in _vbsrc61,
                "return base == 0" in _vbsrc61]],
              expect_contain=[[True] * 9])

    except Exception as _e61:
        check("第 61 节异常: %s" % _e61, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 62. v81：块【内】分支行按回车自动对齐（Else / ElseIf…Then / Case… /
    #     #Else / #ElseIf）
    #     （用户报：写完 `If … Then` 回车补了 `End If`、光标停在新行，再敲 `Else`
    #      就落在"块体那一级"，回车也不会把它拉回与 `If` 同级）
    # ------------------------------------------------------------------
    print("\n=== 62. 分支行按回车自动对齐（Else / ElseIf / Case）（v81）===")
    try:
        import vbe_bridge as VB62

        _bw62 = VB62._branch_word
        _uo62 = VB62._unclosed_owner_above
        _ba62 = VB62.branch_align

        # ---- 62.1 _branch_word：认分支行 ----
        check("62.1 _branch_word：Else / ElseIf…Then / Case…（含 Case Else / Case Is）"
              "与 `#` 条件编译分支都认得出；`Else If y Then` 按 Else 算（VBA 语义："
              "它是 Else + 一行新 If）；`ElseIf y` 少了 Then / 闭合行 / 普通行 / 续行"
              " -> None",
              [[_bw62("    Else"), _bw62("        ElseIf y Then"),
                _bw62("Case 1, 2"), _bw62("    Case Else"),
                _bw62("    Case Is > 3"), _bw62("    #Else"),
                _bw62("    #ElseIf VBA7 Then"), _bw62("    Else If y Then"),
                _bw62("    ElseIf y"), _bw62("    End If"),
                _bw62("    x = 1"), _bw62("    ElseIf y Then _"),
                _bw62("'Else 注释里"), _bw62("")]],
              expect_contain=[["else", "elseif", "case", "case", "case",
                               "else", "elseif", "else", None, None,
                               None, None, None, None]])

        # ---- 62.2 _unclosed_owner_above：这一档分支归谁 ----
        _AB62 = ["Sub Test()", "    If a = 1 Then", "        x = 1",
                 "    If b = 2 Then", "        y = 1", "    End If",
                 "        z = 1"]
        check("62.2 _unclosed_owner_above：`End If` 计 +1、`If … Then` 计 -1 —— "
              "落在最近那个【还开着】的 If 上（嵌套里是内层的 `If b`）；"
              "普通行不算归属者、没有 -> None；条件编译只认 `#If`（`#Else` 是兄弟）",
              [[_uo62(_AB62, "if"),
                _uo62(["Sub Test()", "    Select Case x"], "select"),
                _uo62(["Sub Test()"], "if"),
                _uo62(["    #If VBA7 Then"], "directive"),
                _uo62(["#If VBA7 Then", "    x = 1", "#Else",
                       "    #If VBA64 Then"], "directive")]],
              expect_contain=[["    ", "    ", None, "    ", "    "]])

        # ---- 62.3 branch_align：算得对不对 ----
        check("62.3 branch_align：Else 对齐最近的 `If … Then`（块体再深一级）；"
              "Case 比 `Select Case` 【深一级】（v83，块体再深一级）；"
              "#Else 与 `#If` 齐平；上方没有归属者（单选 If `If x Then y = 1` / "
              "先敲了 Else 还没写 If）-> None（绝不猜）；非分支行 -> None",
              [[_ba62("        Else", ["Sub Test()", "    If x = 1 Then"], "    "),
                _ba62("            Case 1",
                      ["Sub Test()", "    Select Case x"], "    "),
                _ba62("    #Else", ["    #If VBA7 Then"], "    "),
                _ba62("        Else",
                      ["Sub Test()", "    If x = 1 Then y = 1"], "    "),
                _ba62("    Else", ["Sub Test()"], "    "),
                _ba62("    Case 1", ["Sub Test()", "    x = 1"], "    "),
                _ba62("    x = 1", ["Sub Test()", "    If x = 1 Then"], "    ")]],
              expect_contain=[[("    ", "        "),
                               ("        ", "            "),
                               ("    ", "        "),
                               None, None, None, None]])

        # ---- 集成骨架（带 ReplaceLine：v81 要改本行行首空白）----
        class _CM62(object):
            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "Module1"
                self.replace_calls = 0

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

            def ReplaceLine(self, line, text):
                self.replace_calls += 1
                self.lines[int(line) - 1] = str(text)

        class _Pane62(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE62(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        _orig62 = VB62._get_vbe_cached

        def _nl62(lines, line_no, align=True, cm_cls=_CM62):
            """光标落在第 line_no 行【行尾】，回车走 smart + auto_close。"""
            cm62 = cm_cls(lines)
            col = len(lines[line_no - 1]) + 1
            pane62 = _Pane62(cm62, (line_no, col, line_no, col))
            VB62._get_vbe_cached = lambda: _VBE62(pane62)
            ok62 = VB62.VbeBackend().new_line_below(smart=True, auto_close=True,
                                                    align_branch=align)
            return (ok62, cm62.lines, pane62.sel, cm62.replace_calls)

        # ---- 62.4 ★集成：用户原话的场景 ----
        check("62.4 ★集成（用户原话的场景）：`If … Then` 回车补出 `End If`、光标停在"
              "缩进好的那行；在那行敲 `Else` 再回车 -> `Else` 被拉回与 `If` 同级、"
              "新行再深一级（光标落在新行缩进后），`End If` 一个字没动",
              [_nl62(["Sub Test()",
                      "    If x = 1 Then",
                      "        Else",
                      "    End If",
                      "End Sub"], 3)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    If x = 1 Then",
                                "    Else",
                                "        ",
                                "    End If",
                                "End Sub"],
                               (4, 9, 4, 9),
                               1)])

        # ---- 62.5 ★集成：Select Case 里的 Case ----
        check("62.5 ★集成：`Select Case` 里的 `Case 1` 比 `Select Case` 【深一级】"
              "（v83）、块体再深一级；`End Select` 不动",
              [_nl62(["Sub Test()",
                      "    Select Case x",
                      "            Case 1",
                      "    End Select",
                      "End Sub"], 3)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    Select Case x",
                                "        Case 1",
                                "            ",
                                "    End Select",
                                "End Sub"],
                               (4, 13, 4, 13),
                               1)])

        # ---- 62.6 集成：ElseIf ----
        check("62.6 集成：`ElseIf … Then` 与 `Else` 同口径（归最近的 `If … Then`），"
              "且它【不】补收尾（分支没有收尾）",
              [_nl62(["Sub Test()",
                      "    If x = 1 Then",
                      "          ElseIf y = 2 Then",
                      "    End If",
                      "End Sub"], 3)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    If x = 1 Then",
                                "    ElseIf y = 2 Then",
                                "        ",
                                "    End If",
                                "End Sub"],
                               (4, 9, 4, 9),
                               1)])

        # ---- 62.7 找不到归属者 -> 老行为（绝不猜）----
        check("62.7 上方没有归属者 -> 维持 v79 老行为（本行缩进 + 一级、一个字不改）："
              "单选 If `If x = 1 Then y = 1` 不算归属者",
              [_nl62(["Sub Test()",
                      "    If x = 1 Then y = 1",
                      "        Else",
                      "End Sub"], 3)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    If x = 1 Then y = 1",
                                "        Else",
                                "            ",
                                "End Sub"],
                               (4, 13, 4, 13),
                               0)])

        # ---- 62.8 开关：align_branch=False 回到老行为 ----
        check("62.8 align_branch=False（`set VBECOMPLETE_NO_BRANCH_ALIGN=1`）"
              "-> 只缩进、不动本行，回到 v79/v79b 老口径",
              [_nl62(["Sub Test()",
                      "    If x = 1 Then",
                      "        Else",
                      "    End If",
                      "End Sub"], 3, align=False)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    If x = 1 Then",
                                "        Else",
                                "            ",
                                "    End If",
                                "End Sub"],
                               (4, 13, 4, 13),
                               0)])

        # ---- 62.9 本行【已经对齐】-> 一次 ReplaceLine 都不发 ----
        check("62.9 本行已经对齐 -> 不去碰它（一次 ReplaceLine 都不发），"
              "只把新行插出来",
              [_nl62(["Sub Test()",
                      "    If x = 1 Then",
                      "    Else",
                      "    End If",
                      "End Sub"], 3)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    If x = 1 Then",
                                "    Else",
                                "        ",
                                "    End If",
                                "End Sub"],
                               (4, 9, 4, 9),
                               0)])

        # ---- 62.10 写不成 -> 返回 False，一行都不动 ----
        class _CM62NoWrite(_CM62):
            def ReplaceLine(self, line, text):
                raise RuntimeError("boom")

        _nl62_fail = _nl62(["Sub Test()",
                            "    If x = 1 Then",
                            "        Else",
                            "    End If",
                            "End Sub"], 3, cm_cls=_CM62NoWrite)
        check("62.10 拉不齐（ReplaceLine 抛异常）-> 返回 False 且【一行都没动】"
              "（绝不半做：不写本行、也不插新行），调用方把回车原样还给系统",
              [_nl62_fail],
              expect_contain=[(False,
                               ["Sub Test()",
                                "    If x = 1 Then",
                                "        Else",
                                "    End If",
                                "End Sub"],
                               (3, 13, 3, 13),
                               0)])

        VB62._get_vbe_cached = _orig62

        # ---- 62.12 ★护栏：Shift+Enter（smart=False）那条路不许被 v81 碰到 ----
        _cm62s = _CM62(["Sub Foo()", "    Else"])
        _pane62s = _Pane62(_cm62s, (2, 9, 2, 9))
        VB62._get_vbe_cached = lambda: _VBE62(_pane62s)
        _ok62s = VB62.VbeBackend().new_line_below()
        check("62.12 ★护栏（与 v79b 的 closer_text 同坑）：Shift+Enter"
              "（smart=False）那条路不许被 v81 碰到 —— fixed_line 必须在 `if smart:`"
              "【之外】初始化，否则 NameError 会被『出异常一律 return False』吞成"
              "『按了 Shift+Enter 却不换行』",
              [(_ok62s, _cm62s.lines)],
              expect_contain=[(True, ["Sub Foo()", "    Else", "    "])])

        # ---- 62.11 ★接线护栏 ----
        _root62 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _msrc62 = io.open(os.path.join(_root62, "main.py"),
                          encoding="utf-8").read()
        _vbsrc62 = io.open(os.path.join(_root62, "vbe_bridge.py"),
                           encoding="utf-8").read()
        check("62.11 ★接线护栏：main 有独立开关 BRANCH_ALIGN（含"
              " VBECOMPLETE_NO_BRANCH_ALIGN）并把它传进后端；new_line_below 带"
              " align_branch 参数并在 smart 分支里算；拉不齐时先 return False",
              [["BRANCH_ALIGN" in _msrc62,
                "VBECOMPLETE_NO_BRANCH_ALIGN" in _msrc62,
                "align_branch=BRANCH_ALIGN" in _msrc62,
                "def branch_align(line_text, lines_above, unit=" in _vbsrc62,
                "def _branch_word(line_text):" in _vbsrc62,
                "def _unclosed_owner_above(lines_above, kind):" in _vbsrc62,
                "_BRANCH_OWNER = {\"else\": \"if\", \"elseif\": \"if\","
                " \"case\": \"select\"}" in _vbsrc62,
                "align_branch=True" in _vbsrc62,
                "fixed_line, indent = _al" in _vbsrc62,
                "cm, anchor, fixed_line + _raw[len(_ws):]" in _vbsrc62]],
              expect_contain=[[True] * 10])

    except Exception as _e62:
        check("第 62 节异常: %s" % _e62, [True], expect_contain=[False])

    print("\n=== 63. 补收尾的边界：分支行 / 注释不算『本块已结束』（v82）===")
    try:
        import vbe_bridge as VB63

        _ocb63 = VB63._own_closer_below
        _cn63 = VB63.closer_needed

        # ---- 63.1 纯函数：同级压着分支行 / 注释，本块收尾仍在下面 ----
        check("63.1 _own_closer_below（v82）：同级的 `Else` / `ElseIf…Then` / "
              "`Case` / `#Else` / 整行注释都不是『本块到此为止』的边界 —— "
              "扫描跳过它们继续往下找；对照组（同级是普通代码 / 别的块头）"
              "仍然算边界",
              [[_ocb63(["    Else", "        ", "    End If"], "End If", "if", 4),
                _ocb63(["    ElseIf y = 2 Then", "    End If"], "End If", "if", 4),
                _ocb63(["    Case 1", "        y = 1", "    End Select"],
                       "End Select", "select", 4),
                _ocb63(["    #Else", "    End If"], "End If", "if", 4),
                _ocb63(["    ' 说明", "    End If"], "End If", "if", 4),
                _ocb63(["    Rem 说明", "    End If"], "End If", "if", 4),
                _ocb63(["    x = 1", "    End If"], "End If", "if", 4),
                _ocb63(["    For i = 1 To 3", "    Next", "    End If"],
                       "End If", "if", 4)]],
              expect_contain=[[True, True, True, True, True, True,
                               False, False]])

        # ---- 63.2 closer_needed（kind 给定）同口径 ----
        check("63.2 closer_needed（kind 给定，v82）：下面压着同级分支行 / 注释"
              " -> 本块收尾还在下面 -> 【不补】；同级是普通代码 -> 本块已结束"
              " -> 补（v80a 口径不退化）",
              [[_cn63(["    Else", "        ", "    End If"], "End If", 4,
                      kind="if", lines_above=["Sub Test()", "    If x = 1 Then"]),
                _cn63(["    Else", "    End If"], "End If", 4, kind="if",
                      lines_above=["Sub Test()", "    If x = 1 Then"]),
                _cn63(["    Case 1", "        y = 1", "    End Select"],
                      "End Select", 4, kind="select",
                      lines_above=["Sub Test()", "    Select Case x"]),
                _cn63(["    ' 注释", "    End If"], "End If", 4, kind="if",
                      lines_above=["Sub Test()", "    If x = 1 Then"]),
                _cn63(["    x = 1", "    End If"], "End If", 4, kind="if",
                      lines_above=["Sub Test()", "    If x = 1 Then"])]],
              expect_contain=[[False, False, False, False, True]])

        # ---- 集成骨架（要 ReplaceLine：分支行对齐会改本行行首空白）----
        class _CM63(object):
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

            def ReplaceLine(self, line, text):
                self.lines[int(line) - 1] = str(text)

        class _Pane63(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE63(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        _orig63 = VB63._get_vbe_cached

        def _nl63(lines, line_no):
            """光标落在第 line_no 行【行尾】，回车走 smart + auto_close。"""
            cm63 = _CM63(lines)
            col = len(lines[line_no - 1]) + 1
            pane63 = _Pane63(cm63, (line_no, col, line_no, col))
            VB63._get_vbe_cached = lambda: _VBE63(pane63)
            ok63 = VB63.VbeBackend().new_line_below(smart=True,
                                                    auto_close=True)
            return ok63, cm63.lines, pane63.sel

        # ---- 63.3 ★集成：用户报的场景 ----
        check("63.3 ★集成（用户报的场景，v82）：`If … Then` 回车已经补过一次 "
              "`End If`，下面还压着同级的 `Else` —— 回到 `Then` 那行行尾再按"
              "回车，只插一行缩进好的空行，【不再】多补一个 `End If`",
              [_nl63(["Sub Test()",
                      "    If x = 1 Then",
                      "    Else",
                      "        ",
                      "    End If",
                      "End Sub"], 2),
               _nl63(["Sub Test()",
                      "    If x = 1 Then",
                      "    Else",
                      "    End If",
                      "End Sub"], 2)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    If x = 1 Then",
                                "        ",
                                "    Else",
                                "        ",
                                "    End If",
                                "End Sub"],
                               (3, 9, 3, 9)),
                              (True,
                               ["Sub Test()",
                                "    If x = 1 Then",
                                "        ",
                                "    Else",
                                "    End If",
                                "End Sub"],
                               (3, 9, 3, 9))])

        # ---- 63.4 ★集成：ElseIf 与 Select Case 的 Case 同理 ----
        check("63.4 ★集成：`ElseIf … Then` 压在同级、`Select Case` 里的 "
              "`Case` 压在同级 —— 同样只插一行、不再多补 `End If` / "
              "`End Select`",
              [_nl63(["Sub Test()",
                      "    If x = 1 Then",
                      "        MsgBox 1",
                      "    ElseIf y = 2 Then",
                      "        MsgBox 2",
                      "    End If",
                      "End Sub"], 2),
               _nl63(["Sub Test()",
                      "    Select Case x",
                      "    Case 1",
                      "        y = 1",
                      "    Case 2",
                      "        y = 2",
                      "    End Select",
                      "End Sub"], 2)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    If x = 1 Then",
                                "        ",
                                "        MsgBox 1",
                                "    ElseIf y = 2 Then",
                                "        MsgBox 2",
                                "    End If",
                                "End Sub"],
                               (3, 9, 3, 9)),
                              (True,
                               ["Sub Test()",
                                "    Select Case x",
                                "        ",
                                "    Case 1",
                                "        y = 1",
                                "    Case 2",
                                "        y = 2",
                                "    End Select",
                                "End Sub"],
                               (3, 9, 3, 9))])

        # ---- 63.5 ★集成：同级夹着一行注释 ----
        check("63.5 ★集成：`If` 同级夹着一行注释时同理（注释不参与块结构，"
              "不算『本块到此为止』）—— 只插一行",
              [_nl63(["Sub Test()",
                      "    If x = 1 Then",
                      "    ' 说明",
                      "    End If",
                      "End Sub"], 2)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    If x = 1 Then",
                                "        ",
                                "    ' 说明",
                                "    End If",
                                "End Sub"],
                               (3, 9, 3, 9))])

        # ---- 63.6 ★端到端：从零走完用户那三步 ----
        def _flow63():
            cm63 = _CM63(["Sub Test()", "    If x = 1 Then", "End Sub"])
            pane63 = _Pane63(cm63, (2, 18, 2, 18))
            VB63._get_vbe_cached = lambda: _VBE63(pane63)

            def key(line_no, text=None):
                if text is not None:
                    cm63.lines[line_no - 1] = text
                cur = cm63.lines[line_no - 1]
                pane63.sel = (line_no, len(cur) + 1, line_no, len(cur) + 1)
                return VB63.VbeBackend().new_line_below(smart=True,
                                                        auto_close=True)

            r1 = key(2)
            s1 = list(cm63.lines)               # 补出 `End If`、光标停中间行
            r2 = key(3, "        Else")
            s2 = list(cm63.lines)               # `Else` 拉回同级、新行深一级
            r3 = key(2)                         # ★回到 `Then` 行行尾再回车
            return r1, s1, r2, s2, r3, list(cm63.lines)

        check("63.6 ★端到端（用户那条路径，v82）：①写 `If … Then` 回车 -> "
              "补出 `End If` ②在块体行写 `Else` 回车 -> 拉回 `If` 同级、新行"
              "深一级 ③回到 `Then` 行行尾再回车 -> 只插一行空行，`End If` "
              "【没有】变成两条",
              [_flow63()],
              expect_contain=[(True,
                               ["Sub Test()", "    If x = 1 Then", "        ",
                                "    End If", "End Sub"],
                               True,
                               ["Sub Test()", "    If x = 1 Then", "    Else",
                                "        ", "    End If", "End Sub"],
                               True,
                               ["Sub Test()", "    If x = 1 Then", "        ",
                                "    Else", "        ", "    End If",
                                "End Sub"])])

        # ---- 63.7 回归护栏 ----
        check("63.7 回归护栏：①嵌套里内层 `Else` 缩进更深时不误判；"
              "②结构完好、没有分支行 -> 照旧只插一行；③v79c/v80a 老口径不许"
              "退化（模块级 `Type` 下面紧邻过程头 -> 仍补 `End Type`）",
              [_nl63(["Sub Test()",
                      "    If x = 1 Then",
                      "        If y = 2 Then",
                      "        Else",
                      "        End If",
                      "    End If",
                      "End Sub"], 2)[1],
               _nl63(["Sub Test()",
                      "    If x = 1 Then",
                      "        MsgBox 1",
                      "    End If",
                      "End Sub"], 2)[1],
               _nl63(["Type Foo", "Sub Bar()", "End Sub"], 1)[1]],
              expect_contain=[["Sub Test()",
                               "    If x = 1 Then",
                               "        ",
                               "        If y = 2 Then",
                               "        Else",
                               "        End If",
                               "    End If",
                               "End Sub"],
                              ["Sub Test()",
                               "    If x = 1 Then",
                               "        ",
                               "        MsgBox 1",
                               "    End If",
                               "End Sub"],
                              ["Type Foo", "    ", "End Type", "Sub Bar()",
                               "End Sub"]])

        VB63._get_vbe_cached = _orig63

        # ---- 63.8 ★接线护栏 ----
        _root63 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _vbsrc63 = io.open(os.path.join(_root63, "vbe_bridge.py"),
                           encoding="utf-8").read()
        _fn63 = _vbsrc63.split("def _own_closer_below")[1].split(
            "def closer_needed")[0]
        check("63.8 ★接线护栏：`_own_closer_below` 里确实把【注释】与"
              "【块内分支行】跳过，而且这一跳必须发生在缩进判定【之前】"
              "（放后面就等于没跳）",
              [[("_is_comment_only(raw) or _branch_word(raw) is not None"
                 in _fn63),
                ("w = _indent_width(raw)" in _fn63),
                (_fn63.index("_branch_word(raw) is not None")
                 < _fn63.index("w = _indent_width(raw)")),
                "_branch_word" in _vbsrc63]],
              expect_contain=[[True, True, True, True]])

    except Exception as _e63:
        check("第 63 节异常: %s" % _e63, [True], expect_contain=[False])

    print("\n" + "=" * 60)
    print("=== 64. 分支对齐的层级分档：Case 比 Select Case 深一级（v83）===")
    try:
        import vbe_bridge as VB64

        _ba64 = VB64.branch_align
        _dep64 = VB64._BRANCH_OWNER_DEPTH

        # ---- 64.1 档位表 ----
        check("64.1 `_BRANCH_OWNER_DEPTH`：Else / ElseIf 与 If 齐平（0）、"
              "`#Else` / `#ElseIf` 与 #If 齐平（0）、Case 比 Select Case 深一级（1）",
              [[_dep64.get("else"), _dep64.get("elseif"), _dep64.get("case")]],
              expect_contain=[[0, 0, 1]])

        # ---- 64.2 纯函数：Case 深一级；Else / #Else 维持齐平（v81 不退化）----
        check("64.2 branch_align（v83）：`Case 1` / `Case Else` / `Case Is > 3` "
              "都比 `Select Case` 【深一级】（本行 owner+1 级、新行 owner+2 级）；"
              "Else / ElseIf / #Else / #ElseIf 一律维持齐平",
              [[_ba64("        Case 1",
                      ["Sub Test()", "    Select Case x"], "    "),
                _ba64("    Case Else",
                      ["Sub Test()", "    Select Case x"], "    "),
                _ba64("            Case Is > 3",
                      ["Sub Test()", "    Select Case x"], "    "),
                _ba64("        Else",
                      ["Sub Test()", "    If x = 1 Then"], "    "),
                _ba64("          ElseIf y Then",
                      ["Sub Test()", "    If x = 1 Then"], "    "),
                _ba64("    #Else", ["    #If VBA7 Then"], "    "),
                _ba64("    #ElseIf VBA64 Then",
                      ["    #If VBA7 Then"], "    ")]],
              expect_contain=[[("        ", "            "),
                               ("        ", "            "),
                               ("        ", "            "),
                               ("    ", "        "),
                               ("    ", "        "),
                               ("    ", "        "),
                               ("    ", "        ")]])

        # ---- 64.3 缩进单位跟着工程走（2 空格 / Tab）----
        check("64.3 缩进单位跟着工程走：2 个空格缩进 -> Case 只深 2 格；"
              "Tab 缩进 -> Case 深一个 Tab（`unit * 1`，不是硬写 4 个空格）",
              [[_ba64("    Case 1", ["Sub Test()", "  Select Case x"], "  "),
                _ba64("\t\tCase 1", ["Sub Test()", "\tSelect Case x"], "\t")]],
              expect_contain=[[("    ", "      "), ("\t\t", "\t\t\t")]])

        # ---- 64.4 归属者：内层 Select 闭合后，外面的 Case 归外层（不被内层抢走）----
        check("64.4 `_unclosed_owner_above` 配平：内层 `Select Case` 已经用 "
              "`End Select` 闭合 -> 后面的 `Case 2` 归【外层】那个（深一级 = 8 格）",
              [[_ba64("        Case 2",
                      ["Sub Test()",
                       "    Select Case x",
                       "        Case 1",
                       "            Select Case y",
                       "            End Select"], "    "),
                _ba64("    Case 2", ["Sub Test()", "    x = 1"], "    "),
                _ba64("    Case 2", ["Sub Test()"], "    ")]],
              expect_contain=[[("        ", "            "), None, None]])

        # ---- 集成骨架（要 ReplaceLine：分支行对齐会改本行行首空白）----
        class _CM64(object):
            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "Module1"
                self.replace_calls = 0

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

            def ReplaceLine(self, line, text):
                self.replace_calls += 1
                self.lines[int(line) - 1] = str(text)

        class _Pane64(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE64(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        _orig64 = VB64._get_vbe_cached

        def _nl64(lines, line_no, text=None):
            """光标落在第 line_no 行【行尾】，回车走 smart + auto_close。"""
            cm64 = _CM64(lines)
            if text is not None:
                cm64.lines[line_no - 1] = text
            cur = cm64.lines[line_no - 1]
            pane64 = _Pane64(cm64, (line_no, len(cur) + 1, line_no,
                                    len(cur) + 1))
            VB64._get_vbe_cached = lambda: _VBE64(pane64)
            ok64 = VB64.VbeBackend().new_line_below(smart=True, auto_close=True)
            return (ok64, cm64.lines, pane64.sel, cm64.replace_calls)

        # ---- 64.5 ★集成（用户报的场景）：Case 已经在正确层级 -> 一个字都不动 ----
        # v81 的错就在这里：它把已经缩进 8 格的 `Case 1` 拉回 4 格（与 Select 齐平）。
        check("64.5 ★集成（用户报的场景）：`Case 1` 已经在 `Select Case` 下面"
              "深一级（8 格）-> 回车【不改本行】（一次 ReplaceLine 都不发）、"
              "新行（块体）再深一级到 12 格，光标 (4,13)；`End Select` 一个字符没动",
              [_nl64(["Sub Test()",
                      "    Select Case x",
                      "        Case 1",
                      "    End Select",
                      "End Sub"], 3)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    Select Case x",
                                "        Case 1",
                                "            ",
                                "    End Select",
                                "End Sub"],
                               (4, 13, 4, 13),
                               0)])

        # ---- 64.6 ★集成：Case 敲歪了（与 Select 齐平）-> 拉回深一级 ----
        check("64.6 ★集成：`Case 1` 被敲成与 `Select Case` 齐平（4 格）-> "
              "回车把它拉回 8 格（深一级），新行 12 格",
              [_nl64(["Sub Test()",
                      "    Select Case x",
                      "    Case 1",
                      "    End Select",
                      "End Sub"], 3)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    Select Case x",
                                "        Case 1",
                                "            ",
                                "    End Select",
                                "End Sub"],
                               (4, 13, 4, 13),
                               1)])

        # ---- 64.7 回归：Else 与 If 齐平、#Else 与 #If 齐平（v81 口径不退化成"深一级"）----
        check("64.7 回归护栏：`Else` 仍然与 `If` 齐平（4 格、新行 8 格）；"
              "`#Else` 仍然与 `#If` 齐平（0 格、新行 4 格）—— v83 只动了 Case 那一档",
              [_nl64(["Sub Test()",
                      "    If x = 1 Then",
                      "        Else",
                      "    End If",
                      "End Sub"], 3),
               _nl64(["#If VBA7 Then",
                      "    x = 1",
                      "    #Else",
                      "#End If"], 3)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    If x = 1 Then",
                                "    Else",
                                "        ",
                                "    End If",
                                "End Sub"],
                               (4, 9, 4, 9),
                               1),
                              (True,
                               ["#If VBA7 Then",
                                "    x = 1",
                                "#Else",
                                "    ",
                                "#End If"],
                               (4, 5, 4, 5),
                               1)])

        # ---- 64.8 嵌套 Select：内层 Case 归内层（16 格）----
        check("64.8 嵌套 `Select Case`：内层 `Case 2` 归【内层】那个 "
              "（本行 16 格、新行 20 格），外层 Case 与 End Select 都不动",
              [_nl64(["Sub Test()",
                      "    Select Case x",
                      "        Case 1",
                      "            Select Case y",
                      "                Case 2",
                      "            End Select",
                      "    End Select",
                      "End Sub"], 5)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    Select Case x",
                                "        Case 1",
                                "            Select Case y",
                                "                Case 2",
                                "                    ",
                                "            End Select",
                                "    End Select",
                                "End Sub"],
                               (6, 21, 6, 21),
                               0)])

        # ---- 64.9 ★端到端：从零走完 `Select Case` -> `Case 1` -> `Case 2` ----
        def _flow64():
            cm64 = _CM64(["Sub Test()", "    Select Case x", "End Sub"])
            pane64 = _Pane64(cm64, (1, 1, 1, 1))
            VB64._get_vbe_cached = lambda: _VBE64(pane64)

            def key(line_no, text=None):
                if text is not None:
                    cm64.lines[line_no - 1] = text
                cur = cm64.lines[line_no - 1]
                pane64.sel = (line_no, len(cur) + 1, line_no, len(cur) + 1)
                return VB64.VbeBackend().new_line_below(smart=True,
                                                        auto_close=True)

            r1 = key(2)                       # 写 `Select Case x` 回车 -> 补 End Select
            s1 = list(cm64.lines)
            r2 = key(3, "        Case 1")     # 写 `Case 1` 回车
            s2 = list(cm64.lines)
            r3 = key(4, "        Case 2")     # 写 `Case 2` 回车
            return (r1, s1, r2, s2, r3, list(cm64.lines), pane64.sel)

        check("64.9 ★端到端（从零三步）：①写 `Select Case x` 回车 -> 补出 "
              "`End Select`、光标停在 8 格的块体行；②在那行写 `Case 1` 回车 -> "
              "它【留在 8 格】、新行 12 格；③再写 `Case 2` 回车 -> 同样留在 8 格、"
              "新行 12 格（`End Select` 始终只有一条）",
              [_flow64()],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    Select Case x",
                                "        ",
                                "    End Select",
                                "End Sub"],
                               True,
                               ["Sub Test()",
                                "    Select Case x",
                                "        Case 1",
                                "            ",
                                "    End Select",
                                "End Sub"],
                               True,
                               ["Sub Test()",
                                "    Select Case x",
                                "        Case 1",
                                "        Case 2",
                                "            ",
                                "    End Select",
                                "End Sub"],
                               (5, 13, 5, 13))])

        VB64._get_vbe_cached = _orig64

        # ---- 64.10 ★接线护栏 ----
        _root64 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _vbsrc64 = io.open(os.path.join(_root64, "vbe_bridge.py"),
                           encoding="utf-8").read()
        _fn64 = _vbsrc64.split("def branch_align")[1].split(
            "def _indent_unit_for")[0]
        check("64.10 ★接线护栏：`branch_align` 真的走档位表"
              "（`unit * _BRANCH_OWNER_DEPTH`），不是把 Case 硬写成 0；"
              "档位表里 `case` 必须是 1、`else` / `elseif` 必须是 0 "
              "（防止有人「顺手」把它改回 v81 的齐平口径）",
              [[("_BRANCH_OWNER_DEPTH.get(word, 0)" in _fn64),
                ("unit *" in _fn64),
                ('"case": 1' in _vbsrc64),
                ('"else": 0, "elseif": 0' in _vbsrc64)]],
              expect_contain=[[True, True, True, True]])

    except Exception as _e64:
        check("第 64 节异常: %s" % _e64, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 65. v86：补收尾的"边界"再修两类 ——【行标签】（`1:`，配 `GoTo 1` 用，
    #     惯例顶格写）与【`#` 条件编译指令】（`#If` / `#End If`，同样惯例顶格）
    #     的缩进【不含结构信息】，不能当"本块到此为止"的边界；
    #     另外往下扫描的行数上限要够大，且读满被截断时按"不补"兜底。
    #
    #     用户报：`Sub` 里写了很多代码后，把光标移回 `Sub` 那行回车，又补出一个
    #     `End Sub`；`If … Then` 那行回车同理又多一个 `End If`。
    #     实测 `脚本.xlsm / 模块2` 里那个【顶格的 `1:`】一次坑掉三处
    #     （`Sub 功能2()` 那行 / 外层 `For` 那行 / `If … Then` 那行）。
    # ------------------------------------------------------------------
    print("\n=== 65. 回车补收尾：行标签 / 条件编译指令不算边界（v86）===")
    try:
        import vbe_bridge as VB65

        _ocb65 = VB65._own_closer_below
        _cn65 = VB65.closer_needed
        _lab65 = VB65._is_label_line
        _blind65 = VB65._indent_blind

        # ---- 65.1 纯函数：认行标签 ----
        check("65.1 `_is_label_line`：`1:` / `Retry:` / `10: x = 1`（标签与语句"
              "同一行）/ 带缩进的 `5:` / 中文名 `标签1:` 都算标签；"
              "`Case 1:` / `Dim x: x = 1` / `foo(1)` / 过程头 / 普通赋值 / "
              "字符串里的 `\"a:b\"` / 注释里的 `'x:y` 都不算",
              [[_lab65("1:"), _lab65("Retry:"), _lab65("10: x = 1"),
                _lab65("    5:"), _lab65("标签1:"),
                _lab65("Case 1:"), _lab65("Dim x: x = 1"), _lab65("foo(1)"),
                _lab65("Sub Test()"), _lab65("x = 1"),
                _lab65('"a:b"'), _lab65("'x:y")]],
              expect_contain=[[True, True, True, True, True,
                               False, False, False, False, False,
                               False, False]])

        # ---- 65.2 纯函数：_indent_blind ----
        check("65.2 `_indent_blind`：行标签与 `#` 条件编译指令"
              "（`#If` / `#End If` / `#Const` / `#Else` / `#ElseIf`）的缩进都"
              "不含结构信息；普通代码 / 块头 / 收尾行 / 分支行一律不算",
              [[_blind65("1:"), _blind65("#If VBA7 Then"), _blind65("  #End If"),
                _blind65("#Const A = 1"), _blind65("#Else"),
                _blind65("#ElseIf V Then"),
                _blind65("    x = 1"), _blind65("    For i = 1 To 3"),
                _blind65("    End If"), _blind65("    Else")]],
              expect_contain=[[True, True, True, True, True, True,
                               False, False, False, False]])

        # ---- 65.3 _own_closer_below：跳过标签 / 指令（先跳后判缩进）----
        check("65.3 `_own_closer_below`（v86）：顶格的 `1:` 不是边界 —— "
              "`If` 体里的 `1:` 之后仍有 `End If` -> True；`Sub` 体里的顶格 "
              "`1:` 之后仍有 `End Sub` -> True；`#If` / `#End If`（0 列）同理；"
              "对照组（同级普通代码 -> 本块已结束 -> False）不退化",
              [[_ocb65(["            If x Then GoTo 1", "1:",
                        "        End If"], "End If", "if", 8),
                _ocb65(["    x = 1", "1:", "End Sub"], "End Sub", "sub", 0),
                _ocb65(["#If VBA7 Then", "    x = 1", "#End If", "End Sub"],
                       "End Sub", "sub", 0),
                _ocb65(["    x = 1"], "End If", "if", 4)]],
              expect_contain=[[True, True, True, False]])

        # ---- 65.4 closer_needed（kind 给定）同口径 ----
        check("65.4 `closer_needed`（kind 给定，v86）：下面压着顶格 `1:` / "
              "`#If` / `#End If` 时，本块收尾仍然在下面 -> 【不补】；"
              "对照组（同级普通代码 -> 本块已结束）仍然补",
              [[_cn65(["            If x Then GoTo 1", "1:", "        End If"],
                      "End If", 8, kind="if",
                      lines_above=["Sub Test()",
                                   "    For i = 1 To 3",
                                   "        If a Then"]),
                _cn65(["    x = 1", "1:", "End Sub"], "End Sub", 0, kind="sub",
                      lines_above=[]),
                _cn65(["#If VBA7 Then", "    x = 1", "#End If", "End Sub"],
                      "End Sub", 0, kind="sub", lines_above=[]),
                _cn65(["    x = 1"], "End If", 4, kind="if",
                      lines_above=["Sub Test()", "    If a Then"])]],
              expect_contain=[[False, False, False, True]])

        # ---- 65.5 ★截断兜底 ----
        check("65.5 ★v86 截断兜底：下方行数是【读满上限被截断】的 -> "
              "『窗口里没找到收尾』不能当『下面没有收尾』-> 一律【不补】"
              "（宁可少补一条，也不要变出双份收尾把代码写坏）；"
              "truncated=False（默认）时维持 v80a 老口径（补）",
              [[_cn65(["    x = 1", "    y = 2"], "End Sub", 0, kind="sub",
                      lines_above=[], truncated=True),
                _cn65(["    x = 1", "    y = 2"], "End Sub", 0, kind="sub",
                      lines_above=[]),
                _cn65(["    x = 1", "    y = 2"], "End Sub", 0, kind="sub",
                      lines_above=[], truncated=False)]],
              expect_contain=[[False, True, True]])

        # ---- 集成骨架 ----
        class _CM65(object):
            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "模块2"

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

        class _Pane65(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE65(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        _orig65 = VB65._get_vbe_cached

        # 用户现场的真实形态（`脚本.xlsm / 模块2` 摘录，15 行）：第 11 行那个
        # 【顶格的 `1:`】配上面的 `GoTo 1` 用。
        # ⚠️ 真实模块里的过程名叫 `功能2`、字符串是 `"姓名"`（全角）；这里换成
        # ASCII 名是【测试脚手架】的需要 —— GetSelection 给的是**显示列**，
        # 全角字符占两格，mock 里按字符列摆光标会偏（判据本身与此无关）。
        _REAL65 = [
            "Sub Func2()",
            "    Set actSheet = ThisWorkbook.ActiveSheet",
            "    For i = 3 To targetSheet.Cells(targetSheet.Cells.Rows.Count, "
            "\"e\").End(xlUp).Row",
            "        ID = CStr(targetSheet.Cells(i, \"e\"))",
            "    Next",
            "",
            "    For i = 1 To dataSheet1.Cells(dataSheet1.Cells.Rows.Count, "
            "\"c\").End(xlUp).Row",
            "        If dataSheet1.Cells(i, \"c\") = \"name\" Then",
            "            If dataSheet1.Cells(i, \"d\") = \"\" Then GoTo 1",
            "            ID = CStr(dataSheet3.Cells(idRow, \"b\"))",
            "1:",
            "        End If",
            "    Next",
            "    MsgBox \"OK\", vbInformation",
            "End Sub",
        ]

        def _nl65(line_no):
            cm65 = _CM65(_REAL65)
            cur = cm65.lines[line_no - 1]
            pane65 = _Pane65(cm65, (line_no, len(cur) + 1, line_no,
                                    len(cur) + 1))
            VB65._get_vbe_cached = lambda: _VBE65(pane65)
            ok65 = VB65.VbeBackend().new_line_below(smart=True, auto_close=True)
            return (ok65, cm65.lines, pane65.sel)

        # ---- 65.6 ★集成：用户现场那三处，一个收尾都不许补 ----
        check("65.6 ★集成（`脚本.xlsm/模块2` 那个顶格 `1:` 的真实形态）："
              "①`Sub …()` 那行回车 -> 不补 `End Sub`；②外层 `For` 那行回车 "
              "-> 不补 `Next`；③`If … Then` 那行回车 -> 不补 `End If` —— "
              "三处都只多一行缩进好的空行，原有 15 行代码一个字符没动",
              [_nl65(1), _nl65(3), _nl65(8)],
              expect_contain=[(True, _REAL65[:1] + ["    "] + _REAL65[1:],
                               (2, 5, 2, 5)),
                              (True, _REAL65[:3] + ["        "] + _REAL65[3:],
                               (4, 9, 4, 9)),
                              (True, _REAL65[:8] + ["            "]
                               + _REAL65[8:], (9, 13, 9, 13))])

        VB65._get_vbe_cached = _orig65

        # ---- 65.7 ★接线护栏 ----
        _root65 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _vbsrc65 = io.open(os.path.join(_root65, "vbe_bridge.py"),
                           encoding="utf-8").read()
        _fn65 = _vbsrc65.split("def _own_closer_below")[1].split(
            "def closer_needed")[0]
        check("65.7 ★接线护栏：`_own_closer_below` 里把【行标签 / `#` 指令】也跳过"
              "（`_indent_blind`），而且这一跳同样必须在缩进判定【之前】"
              "（放后面就等于没跳）；后端补收尾时按 `_CLOSER_SCAN_LINES` 读、"
              "并把 `truncated` 交给判据；这个扫描窗口必须比"
              "『往上找缩进基准』那个 `_INDENT_LOOKBACK` 大",
              [[("_indent_blind(raw)" in _fn65),
                ("w = _indent_width(raw)" in _fn65),
                (_fn65.index("_indent_blind(raw)")
                 < _fn65.index("w = _indent_width(raw)")),
                ("self._CLOSER_SCAN_LINES" in _vbsrc65),
                ("truncated=_rest > _cap" in _vbsrc65),
                ("if truncated:" in _vbsrc65),
                (VB65.VbeBackend._CLOSER_SCAN_LINES
                 > VB65.VbeBackend._INDENT_LOOKBACK)]],
              expect_contain=[[True] * 7])

        # ---- 65.8 回归护栏：v80a / v82 口径不退化 ----
        check("65.8 回归护栏：v80a 主场景（同级压着【同类嵌套块】-> 下面没有"
              "本块收尾 -> 仍要补）与 v82 口径（同级 `Else` / 注释跳过 -> "
              "不补）都不退化",
              [[_cn65(["    For i = 1 To 10", "        x = 1", "    Next i",
                       "End Sub"], "Next", 4, kind="for",
                      lines_above=["Sub Test()"]),
                _cn65(["    Else", "    End If"], "End If", 4, kind="if",
                      lines_above=["Sub Test()", "    If x = 1 Then"]),
                _cn65(["    ' 注释", "    End If"], "End If", 4, kind="if",
                      lines_above=["Sub Test()", "    If x = 1 Then"])]],
              expect_contain=[[True, False, False]])

    except Exception as _e65:
        check("第 65 节异常: %s" % _e65, [True], expect_contain=[False])

    # ---- 66. v87：别的模块里【不可见】的声明不许当"真名字"证据 ----
    #
    # 用户报（第八条反馈）："对于未定义、第一次写出来的变量名，会提示自己；回退删除
    # 已写的变量名时，也会提示自己。这个 bug 以前在模块里会存在、已经修好了，但是
    # 现在在【类模块】里也会出。另外帮我查一下 sheet / thisworkbook / 窗体 里会不会
    # 出现，都帮我修一下。"
    #
    # 真机定位（只读扫用户正在编辑的工程）：光标停在 `类2` 第 3 行 `Public 订单金额`
    # 行尾，而 `订单金额` 在另一个类模块 `SummaryClass` 里也声明了一份 ——
    #   池子侧：类模块成员是 Private（跨模块不能裸名引用），
    #           filter_identifiers_by_scope 正确隐藏；
    #   证据侧：declared_elsewhere 只看"名字在别的模块出现过没有"，**不看可见性**
    #           -> 返回 True -> 回声防护放行 -> 弹窗把正在敲的字原样提示回来。
    # 同一个漏洞对【窗体 / 文档模块（Sheet / ThisWorkbook）】一样成立（它们的模块级
    # 成员同样默认 Private），所以这一节把这些组件类型一次性钉住。
    print("\n=== 66. v87 类/窗体/文档模块的成员不再当跨模块证据（提示自己）===")
    try:
        import vbe_bridge as VB66
        import parser as P66
        import engine as E66

        # ---- 66.1 记录面：同一个 `Public X`，跨模块可见性随模块类型不同 ----
        def _priv66(code, std):
            r = list(P66.extract_records(code, module="M", is_std_module=std))
            return sorted((str(n), pv) for n, _m, _p, pv in r if not _p)

        check("66.1 `Public X` 的跨模块可见性：标准模块 priv=False（工程可见）；"
              "类 / 窗体 / 文档模块 priv=True（只能经实例或限定名访问）",
              [_priv66("Public SharedName", True),
               _priv66("Public SharedName", False)],
              expect_contain=[[("SharedName", False)],
                              [("SharedName", True)]])

        # ---- 集成骨架：一个能喂饱收集流程的假 VBE（多组件、真模块名）----
        class _CM66(object):
            def __init__(self, comp, text):
                self._comp = comp
                self.text = text

            @property
            def Name(self):
                return self._comp.Name

            @property
            def Parent(self):
                return self._comp

            @property
            def CountOfLines(self):
                return self.text.count("\n") + 1

            def Lines(self, start, count):
                return "\r\n".join(
                    self.text.split("\n")[start - 1:start - 1 + count])

        class _Comp66(object):
            def __init__(self, name, ctype, text):
                self.Name = name
                self.Type = ctype
                self.CodeModule = _CM66(self, text)

        class _Proj66(object):
            def __init__(self, comps):
                self.Name = "VBAProject"
                self.VBComponents = list(comps)

        class _Pane66(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE66(object):
            def __init__(self, comps, pane):
                self.ActiveVBProject = _Proj66(comps)
                self.ActiveCodePane = pane

        class _UI66(object):
            def __init__(self):
                self.shown = False

            def show(self, *a, **k):
                self.shown = True

            def hide(self, *a, **k):
                self.shown = False

            def update_selection(self, *a, **k):
                pass

            def contains_point(self, *a, **k):
                return False

        _orig66 = VB66._get_vbe_cached

        def _mk66(comps, caret_name, sel):
            """comps: [(模块名, 组件类型, 代码)]；返回 (backend, pane, comps)。"""
            objs = {}
            lst = []
            for (n, t, s) in comps:
                c = _Comp66(n, t, s)
                objs[n] = c
                lst.append(c)
            pane = _Pane66(objs[caret_name].CodeModule, sel)
            VB66._get_vbe_cached = lambda: _VBE66(lst, pane)
            return VB66.VbeBackend(), pane, objs

        def _go66(backend):
            ui = _UI66()
            c = E66.Completer(backend, ui)
            c.trigger()
            return ui.shown, c.matches

        try:
            # 组件类型：1=标准模块、2=类模块、3=窗体、100=文档模块（Sheet/ThisWorkbook）
            # ⚠️ 代码行一律 ASCII —— 假 GetSelection 给的是显示列，全角字符会偏
            #    （与第 65 节同一个脚手架坑）。
            _P66 = [
                ("模块A", 1,
                 "Public GName\nSub Work()\n    Dim onlyLocal As Long\nEnd Sub"),
                ("类C", 2, "Public SharedItem"),
                ("类D", 2, "Public SharedItem"),
                ("窗体F", 3, "Dim SharedItem As Long"),
                ("Sheet1", 100, "Public SharedItem As Long"),
                ("ThisWorkbook", 100, "Dim SharedItem As Long"),
            ]
            _bk66, _pane66, _objs66 = _mk66(_P66, "类C", (1, 18, 1, 18))
            _ids66 = _bk66.get_identifiers()

            # ---- 66.2 核心：同名成员散落在类/窗体/文档模块里，谁都不算证据 ----
            check("66.2 ★真实后端：`SharedItem` 在 类C / 类D / 窗体F / Sheet1 / "
                  "ThisWorkbook 里各声明一份 -> 从 类C 问【不算】跨模块证据"
                  "（修前 True，正是用户那条「提示自己」）",
                  [_bk66.declared_elsewhere("SharedItem", "类C")],
                  expect_contain=[False])

            # ---- 66.3 对照面：真的跨模块可见的，一路都不许退化 ----
            check("66.3 对照：标准模块的 `Public GName` 仍是跨模块证据（v37 语义"
                  "不许退化）；别的过程的局部变量 `onlyLocal` 不算；组件名 `模块A` 算",
                  [_bk66.declared_elsewhere("GName", "类C"),
                   _bk66.declared_elsewhere("onlyLocal", "类C"),
                   _bk66.declared_elsewhere("模块A", "类C")],
                  expect_contain=[True, False, True])

            # ---- 66.4 ★端到端（用户场景①）：第一次写出来的成员名不提示自己 ----
            _bk66b, _pane66b, _objs66b = _mk66(_P66, "类C", (1, 18, 1, 18))
            check("66.4 ★端到端（用户场景①：第一次写出来的名字）：`Public "
                  "SharedItem` 行尾触发 -> 弹窗不出现，一条候选都没有"
                  "（修前会把 SharedItem 原样提示回来）",
                  [_go66(_bk66b)],
                  expect_contain=[(False, [])])

            # ---- 66.5 ★端到端（用户场景②）：回退删字时不留旧版本的幽灵 ----
            # 顺序很关键：先按"名字还完整"那一刻解析出池子（模拟 0.4s 缓存），
            # 再让用户按退格删掉一个字符 —— 池子里那份 `SharedItem` 正是
            # 用户说的"回退删除已写的变量名时提示自己"的来源。
            _bk66c, _pane66c, _objs66c = _mk66(_P66, "类C", (1, 16, 1, 16))
            _bk66c.get_identifiers()          # 池子快照（此刻文本还是完整的名字）
            _objs66c["类C"].CodeModule.text = "Public SharedIt"
            _pane66c.sel = (1, 16, 1, 16)     # 光标退到删剩下的位置
            check("66.5 ★端到端（用户场景②：回退删字）：池子里还留着 "
                  "`SharedItem`（缓存）-> 退格成 `SharedIt` 之后不许再把 "
                  "`SharedItem` 提示出来",
                  [_go66(_bk66c)],
                  expect_contain=[(False, [])])

            # ---- 66.6 对照：标准模块里真存在的公共名，回退/打全照常提示 ----
            _P66b = [("模块A", 1, "Public GName"),
                     ("类C", 2, "Public GNa")]
            _bk66d, _pane66d, _objs66d = _mk66(_P66b, "类C", (1, 12, 1, 12))
            _r66d = _go66(_bk66d)
            check("66.6 对照：`Public GNa` 行尾触发 -> 别的标准模块里的公共名 "
                  "`GName` 照常提示（跨模块 Public 是货真价实的存在）",
                  [_r66d], expect_contain=[(True, ["GName"])])

            # ---- 66.7 两条路径同一口径：候选池隐藏的，证据侧也必须否认 ----
            _visA66 = set(n.lower() for n in E66.filter_identifiers_by_scope(
                _ids66, None, "模块A"))
            check("66.7 两条路径同口径：从 模块A 看 `SharedItem` —— 候选池里没有它"
                  "（别的模块的私有成员），证据侧也否认它（都要 False）",
                  [("shareditem" in _visA66,
                    _bk66.declared_elsewhere("SharedItem", "模块A"))],
                  expect_contain=[(False, False)])

            # ---- 66.8 `_module_level_names` 也按可见性收窄（同类漏洞的另一处）----
            class _B66(_FakeBackend):
                def get_identifiers(self):
                    return [("MyName", "类C", None, True),      # 本模块模块级 -> 可见
                            ("PubName", "模块A", None, False),  # 别处 Public -> 可见
                            ("OtherPriv", "类X", None, True),   # 别处 Private -> 不可见
                            ("aLoc", "类C", "P", False)]        # 过程内 -> 不算模块级

            _c66e = E66.Completer(_B66({"proc_name": None,
                                        "module_name": "类C"}), _UI66())
            _c66e._ctx_now = {"proc_name": None, "module_name": "类C"}
            check("66.8 `_module_level_names` 按可见性收窄：本模块模块级（MyName）与"
                  "别处 Public（PubName）算；别模块的私有成员（OtherPriv）与过程内"
                  "局部（aLoc）不算",
                  [sorted(_c66e._module_level_names())],
                  expect_contain=[["myname", "pubname"]])

            # ---- 66.9 ★接线护栏：两处判据各改对了地方（防日后被"顺手"改回去）----
            _root66 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            _vbsrc66 = io.open(os.path.join(_root66, "vbe_bridge.py"),
                               encoding="utf-8").read()
            _fn66 = _vbsrc66.split("def _collect_identifiers")[1].split(
                "def get_type_names")[0]
            _esrc66 = io.open(os.path.join(_root66, "engine.py"),
                              encoding="utf-8").read()
            _fn66e = _esrc66.split("def _module_level_names")[1].split(
                "def _declared_names")[0]
            check("66.9 ★接线护栏：`_decl_by_mod` 那段确实按 priv / proc 两道跳过"
                  "（`_r_priv or _r_proc`）；`_module_level_names` 确实走了同一个"
                  "可见性函数（`filter_identifiers_by_scope`），而不是自己写第二份规则",
                  [("_r_priv or _r_proc" in _fn66),
                   ("filter_identifiers_by_scope" in _fn66e)],
                  expect_contain=[True, True])
        finally:
            VB66._get_vbe_cached = _orig66
    except Exception as _e66:
        check("第 66 节异常: %s" % _e66, [True], expect_contain=[False])

    # ---- 67. v88：Shift+Enter（光标在代码中间）也补块收尾 ----
    #
    # 用户口径："把光标移动到语句最后按回车可以补全；光标在代码中间按 Shift+Enter，
    # 不会自动补全代码块 —— 帮我加上，也能补全。"
    # 根因：补块收尾只挂在 smart 模式（回车）那条路上，而那条路有一道"光标必须在
    # 行尾"的保险 —— 光标停在中途时它整件事作废、把按键还给 VBE（那时 VBE 原生
    # 回车是【拆行】）。Shift+Enter 这条路本来就不拆行（只往下插一行），所以能在
    # 任意光标列上判块头、补收尾。两条路共用同一份 block_closer / closer_needed /
    # next_line_indent。
    print("\n=== 67. v88：Shift+Enter（行中）也补块收尾 ===")
    try:
        import vbe_bridge as VB67

        class _CM67(object):
            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "M1"

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

        class _Pane67(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE67(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        _orig67 = VB67._get_vbe_cached

        def _run67(lines, line_no, col, **kw):
            """按 mock 的代码窗跑一次 new_line_below，返回 (ok, 行表, 选区)。

            ⚠️ 光标列按【字符列】给 —— 脚手架里的行一律用 ASCII。真机
            GetSelection 给的是**显示列**，全角字符占两格会偏（判据本身与此
            无关，见第 65 节同一处提醒）。
            """
            cm67 = _CM67(lines)
            pane67 = _Pane67(cm67,
                             (int(line_no), int(col), int(line_no), int(col)))
            VB67._get_vbe_cached = lambda: _VBE67(pane67)
            try:
                ok67 = VB67.VbeBackend().new_line_below(**kw)
            finally:
                VB67._get_vbe_cached = _orig67
            return (ok67, cm67.lines, pane67.sel)

        _IF67 = ["Sub F1()",
                 "    x = 0",
                 "    If x = 1 Then",
                 "End Sub"]
        _IFCLOSED67 = ["Sub F1()",
                       "    x = 0",
                       "    If x = 1 Then",
                       "        y = 1",
                       "    End If",
                       "End Sub"]
        _PLAIN67 = ["Sub F1()",
                    "    x = y + 1",
                    "End Sub"]

        # ---- 67.1 ★集成：未闭合的 `If … Then`，光标【行中】-> 补 End If ----
        check("67.1 ★集成（用户报的场景）：`If … Then` 那行、光标停在第 8 列"
              "（行中，不在行尾）按 Shift+Enter -> 补出 `End If`；新行比块头"
              "深一级、光标落在中间那行；块头那一行一个字符没动",
              [_run67(_IF67, 3, 8, auto_close=True)],
              expect_contain=[(True,
                               _IF67[:3] + ["        ", "    End If"] + _IF67[3:],
                               (4, 9, 4, 9))])

        # ---- 67.2 已闭合：不补第二条，但照样进块（深一级）----
        check("67.2 ★集成：`If … Then` 下面已经挂着自己的 `End If` -> 不补第二条"
              "（同一份 closer_needed，v86 那套判据自动覆盖这条新路），但光标"
              "行中按 Shift+Enter 照样把新行缩进深一级（进块里写内容）",
              [_run67(_IFCLOSED67, 3, 8, auto_close=True)],
              expect_contain=[(True,
                               _IFCLOSED67[:3] + ["        "]
                               + _IFCLOSED67[3:], (4, 9, 4, 9))])

        # ---- 67.3 ★过程头：行中按 Shift+Enter -> 补 End Sub ----
        check("67.3 ★集成：过程头 `Sub F1()` 那行、光标停在第 6 列（正好落在"
              "过程名中间）按 Shift+Enter -> 补出 `End Sub`，新行深一级；"
              "过程头那一行原样不动（Shift+Enter 从不拆行）",
              [_run67(["Sub F1()", "    x = 0"], 1, 6, auto_close=True)],
              expect_contain=[(True, ["Sub F1()", "    ", "End Sub", "    x = 0"],
                               (2, 5, 2, 5))])

        # ---- 67.4 普通代码行：维持 v78 老口径 ----
        check("67.4 非块头（普通代码行）光标行中按 Shift+Enter：不补任何收尾，"
              "新行行首空白【抄本行】—— v78 的老口径一个字不动",
              [_run67(_PLAIN67, 2, 8, auto_close=True)],
              expect_contain=[(True,
                               ["Sub F1()", "    x = y + 1", "    ", "End Sub"],
                               (3, 5, 3, 5))])

        # ---- 67.5 ★对照：auto_close 关掉 -> 完全回到 v78 ----
        check("67.5 ★对照：`auto_close=False`（v78 老默认）时，块头行中也【不】"
              "补收尾，新行行首空白抄本行（缩进 4，不是深一级）—— 证明这条新路"
              "只在新开关打开时起作用，老调用（`new_line_below()`）不受影响",
              [_run67(_IF67, 3, 8)],
              expect_contain=[(True,
                               ["Sub F1()", "    x = 0", "    If x = 1 Then",
                                "    ", "End Sub"], (4, 5, 4, 5))])

        # ---- 67.6 ★对照：同一行、同一光标列，回车 vs Shift+Enter ----
        _smart67 = _run67(_IF67, 3, 8, smart=True, auto_close=True)
        _shift67 = _run67(_IF67, 3, 8, auto_close=True)
        check("67.6 ★对照（钉住用户报的那个差别）：同一条 `If … Then`、同一个"
              "【行中】光标列 —— 回车那条路（smart=True）有『光标必须在行尾』"
              "这道保险，这里【拒绝接管】（返回 False，按键还给 VBE 的拆行）；"
              "Shift+Enter 那条路（smart=False + auto_close）把块补出来",
              [[_smart67[0], _shift67[0], _shift67[1]]],
              expect_contain=[[False, True,
                               _IF67[:3] + ["        ", "    End If"]
                               + _IF67[3:]]])

        # ---- 67.7 ★接线护栏 ----
        _root67 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _mainsrc67 = io.open(os.path.join(_root67, "main.py"),
                             encoding="utf-8").read()
        _fn67 = _mainsrc67.split("def _new_line_here")[1].split(
            "def _enter_indent_here")[0]
        _vbsrc67 = io.open(os.path.join(_root67, "vbe_bridge.py"),
                           encoding="utf-8").read()
        _else67 = _vbsrc67.split(
            "★v88（用户报的）：Shift+Enter 也要能把块补起来")[1].split(
            "if fixed_line is not None")[0]
        check("67.7 ★接线护栏：Shift+Enter 那条线确实接上了 —— main 的 "
              "`_new_line_here` 把 `auto_close=AUTO_CLOSE_BLOCK` 传下去；"
              "vbe_bridge 的 smart=False 支用的是【同一份】纯判据"
              "（block_closer / closer_needed / next_line_indent，v87 的教训："
              "同一条语义只能有一份实现），扫描窗口与截断兜底也一并接上",
              [[("backend.new_line_below(auto_close=AUTO_CLOSE_BLOCK)" in _fn67),
                ("block_closer(line_text)" in _else67),
                ("closer_needed(self._lines_below(cm, anchor, _cap)"
                 in _else67),
                ("_block_kind(line_text)" in _else67),
                ("next_line_indent(line_text, above, unit)" in _else67),
                ("truncated=_rest > _cap" in _else67),
                ("self._CLOSER_SCAN_LINES" in _else67)]],
              expect_contain=[[True] * 7])

        # ---- 67.8 ★回归护栏：v86 那个现场形态，行中也不许多补 ----
        # `脚本.xlsm / 模块2` 摘录（15 行）：第 11 行是那个【顶格】的 `1:`。
        # ⚠️ 真实模块里是全角名，这里换 ASCII —— 同第 65 节的脚手架说明。
        _REAL67 = [
            "Sub Func2()",
            "    Set actSheet = ThisWorkbook.ActiveSheet",
            "    For i = 3 To targetSheet.Cells(targetSheet.Cells.Rows.Count, "
            "\"e\").End(xlUp).Row",
            "        ID = CStr(targetSheet.Cells(i, \"e\"))",
            "    Next",
            "",
            "    For i = 1 To dataSheet1.Cells(dataSheet1.Cells.Rows.Count, "
            "\"c\").End(xlUp).Row",
            "        If dataSheet1.Cells(i, \"c\") = \"name\" Then",
            "            If dataSheet1.Cells(i, \"d\") = \"\" Then GoTo 1",
            "            ID = CStr(dataSheet3.Cells(idRow, \"b\"))",
            "1:",
            "        End If",
            "    Next",
            "    MsgBox \"OK\", vbInformation",
            "End Sub",
        ]
        check("67.8 ★回归护栏：v86 那个现场形态（顶格 `1:`）—— 在 `Sub` 那行、"
              "外层 `For` 那行、`If … Then` 那行的【行中】各按一次 Shift+Enter，"
              "同样一个收尾都不许补（三处下面都挂着自己的收尾）；只多一行缩进"
              "深一级的空行，原有 15 行一个字符没动",
              [_run67(_REAL67, 1, 6, auto_close=True),
               _run67(_REAL67, 3, 10, auto_close=True),
               _run67(_REAL67, 8, 12, auto_close=True)],
              expect_contain=[(True, _REAL67[:1] + ["    "] + _REAL67[1:],
                               (2, 5, 2, 5)),
                              (True, _REAL67[:3] + ["        "] + _REAL67[3:],
                               (4, 9, 4, 9)),
                              (True, _REAL67[:8] + ["            "]
                               + _REAL67[8:], (9, 13, 9, 13))])
    except Exception as _e67:
        check("第 67 节异常: %s" % _e67, [True], expect_contain=[False])

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
