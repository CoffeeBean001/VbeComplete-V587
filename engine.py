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
  apply_completion(line_no, word_start_col, caret_col, completion)
"""

import re
import time

from log import log as _log

# 标识符（含中文）：首字符为字母/下划线，其后可跟字母/数字/下划线。
# [^\W\d] = 非“非单词字符”且非数字 = 字母或下划线（Unicode 感知，含中文）。
_IDENT = re.compile(r"[^\W\d]\w*$")

# 确认补全后的"静默期"（秒）：这段时间内忽略新的触发请求。
# 目的：Tab 确认后，可能还有排队中的 trigger（来自刚才那次按键），
# 若不屏蔽，弹窗会在收起后立刻又冒出来。
SUPPRESS_AFTER_ACCEPT = 0.35


def extract_word_before(line_text, caret_col):
    """返回 (word, start_col)，均为 1-based；无则返回 (None, None)。"""
    prefix = line_text[:caret_col - 1] if caret_col and caret_col > 1 else ""
    m = _IDENT.search(prefix)
    if not m:
        return None, None
    return m.group(0), m.start() + 1


def _ident_char_before_caret(ctx):
    """光标前一个字符是否属于标识符（字母/数字/下划线，含中文）。

    用于"编辑器内容变化"触发的路径：若刚输入的是空格/标点/换行，
    说明用户并没有在拼标识符，应当收起弹窗（保持与按空格收起一致的行为）。
    """
    caret = ctx.get("caret_col") or 0
    if caret <= 1:
        return True   # 行首：交给后续取词逻辑（取不到词自然会收起）
    text = ctx.get("line_text") or ""
    i = caret - 2
    if i < 0 or i >= len(text):
        return True
    ch = text[i]
    return ch.isalnum() or ch == "_"


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
    """
    pos = []
    i = 0
    for ch in query:
        best = None
        best_key = None
        for j in range(i, len(low)):
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


class Completer:
    def __init__(self, backend, ui):
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

    def _name_really_exists(self, name, ctx=None):
        """这个名字在工程里是【真实存在】的，还是只是光标处的回声？

        两条互为补充的证据，命中任一即认为真实存在：

        1. `get_declared_names()` 里有它 —— 全工程范围内被 Dim/Const/Sub/
           Function/Type/Enum 真正声明过。跨模块的 Public 名字靠这一条。
        2. `name_exists_outside_caret()` 为真 —— 把光标处的词抹掉后，它在
           当前模块别处还出现。隐式变量（用到即存在）、以及解析层没能归类
           的声明，靠这一条兜住。

        只用 1 会漏判隐式变量（正是 v35"打全 arr 列表消失"的成因）；
        只用 2 会漏判跨模块的 Public 名字（`name_exists_outside_caret`
        刻意只扫当前模块以省开销）。所以两条都要。
        """
        low = str(name).lower()
        if not low:
            return False
        try:
            if low in self._declared_names():
                return True
        except Exception:
            pass
        if ctx is None:
            ctx = self.ctx or {}
        return _name_exists_outside_caret(self.backend, name,
                                          ctx.get("line_no"),
                                          ctx.get("caret_col"))

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
        try:
            return set(str(i).lower() for i in self.backend.get_identifiers())
        except Exception:
            return set()

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
        if require_ident_before_caret and not _ident_char_before_caret(ctx):
            self.hide()
            return
        word, start_col = extract_word_before(ctx["line_text"], ctx["caret_col"])
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
        for i in visible_ids:
            r = fuzzy_match(i, word)
            if r is not None:
                scored.append((r[0], i, r[1]))
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
            exact = word.lower()
            if any(m.lower() == exact for m in matches) \
                    and not self._name_really_exists(word, ctx):
                matches = [m for m in matches if m.lower() != exact]
        # 注意：这里【不要】再加"唯一候选恰好等于所输词就收起"的收尾规则。
        #
        # 那条规则（v30 为修"回退键提示自己"引入）误伤了正常场景：变量 arr 已定义，
        # 输入 a / ar 都提示 arr，把 arr 打全反倒把列表关掉了 —— 用户明确要求
        # "输入完整，提示保留"，否则没法 Tab 确认、也看不出自己写对了没有。
        #
        # 它原本要防的"提示自己"，现在由上面两条更精确的规则覆盖：
        #   1) 词不是工程里真实声明的名字（回退出来的 numA）-> 剔除它本身；
        #   2) 声明行上正在起的新名字 -> 剔除其完整拼法。
        # 真实存在的名字打全了就该继续提示，与打前缀行为一致。
        # 诊断：把"原始记录里有哪些同名命中、各属于哪个模块/过程"一并记下，
        # 这样日志能直接区分是"压根没收录"还是"收录了但被作用域过滤掉"。
        _log("trigger: line=%s col=%s word=%r proc=%r module=%r visible=%d matches=%r"
             % (ctx.get("line_no"), ctx.get("caret_col"), word,
                ctx.get("proc_name"), ctx.get("module_name"),
                len(visible_ids), matches))
        try:
            hits = [r for r in self.backend.get_identifiers()
                    if fuzzy_match(str(r[0]), word) is not None]
            _log("  raw hits for %r: %s" % (word, hits[:40]))
            _log("  hit positions: %s" % ({k: v for k, v in
                                           list(self.match_hits.items())[:10]},))
        except Exception:
            pass
        if not matches:
            self.hide()
            return
        self.matches = matches
        # 默认选中：若当前词恰好等于某候选（已打全），优先选中它，
        # 按 Tab 即可直接确认当前已输入的完整名字；否则选中列表首项。
        exact = word.lower()
        sel = 0
        for idx, m in enumerate(matches):
            if m.lower() == exact:
                sel = idx
                break
        self.selected = sel
        self.ctx = {
            "line_no": ctx["line_no"],
            "caret_col": ctx["caret_col"],
            "word": word,
            "word_start_col": start_col,
        }
        self.visible = True
        self.shown_at = time.time()
        self.word_is_complete = word.lower() in set(i.lower() for i in visible_ids)
        self.ui.show(self.matches, self.selected, self)

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
        self.ui.update_selection(self.selected)

    def accept(self):
        if not self.visible or not self.matches:
            return
        chosen = self.matches[self.selected % len(self.matches)]
        ctx = self.ctx
        try:
            self.backend.apply_completion(
                ctx["line_no"], ctx["word_start_col"], ctx["caret_col"], chosen
            )
        except Exception:
            # 插入失败也要保证状态复位，否则弹窗会卡住不再触发
            pass
        finally:
            # 进入静默期：后续排队中的 trigger 不会再把它弹回来
            self._suppress_until = time.time() + SUPPRESS_AFTER_ACCEPT
            self.hide()

    def pick(self, index):
        """按序号直接选中并确认（数字键 1..9 选词）。

        index 为 0-based。越界（列表没那么多项）返回 False 且不改变状态，
        调用方据此决定"该按键是否要被吞掉"——越界时应放行，让用户正常输入数字。
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
        try:
            inside = self.ui.contains_point(px, py)
        except Exception:
            inside = False
        if not inside:
            self.hide()
