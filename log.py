# -*- coding: utf-8 -*-
"""诊断日志的空实现（占位）。

这一块原来是一套完整的可选日志（`run_debug.bat` 设 `VBECOMPLETE_LOG=1`
写 `VbeComplete.log`，另有 `show_log.bat` 看日志）。项目收尾时按用户口径
"只保留核心功能"，把那套排障设施与 `run_debug.bat` / `show_log.bat`
一起删掉了 —— 它默认完全关闭、平时零开销，但对日常使用不是必需品。

**为什么留这个空壳、而不是把散在 3 个文件里约 60 处 `_log(...)` 调用
和几段 `if _LOG_ENABLED:` 诊断块逐个挖掉**：

1. 调用点是纯副作用（拼字符串 → 丢弃），挖掉它们对行为**零**影响，可
   那是 60 处机械改动 —— 每处都可能连带改错缩进或吞掉半条语句，收益却
   是零（本项目的判断标准一直是"用户接受后能编译吗、行为有没有变"）。
2. 空实现保持 `log()` 直接 return、`LOG_ENABLED` 恒为 False，
   ⇒ 那几段 `if _LOG_ENABLED:` 诊断块（含 `popup.geometry_now()` 之类
   的额外只读 COM 调用）**自动整段跳过**，不会给主循环添开销。
3. 接口留着，将来真要排障，把日志实现加回本文件即可，其余文件一行不动。

⚠️ 这个模块**不可删**：engine.py / vbe_bridge.py / main.py 都在
    import 它，删掉会直接 ImportError 起不来。
"""

# 恒为 False：主循环里 `if _LOG_ENABLED:` 包着的诊断块整段不执行。
#
# ★v101：**改成真的可开关**（原先恒为 False，等于把排障设施彻底封死）。
# 设 `VBECOMPLETE_LOG=1` 才落盘，默认仍是零开销的空实现 ——
# 日常使用行为与开销与从前**完全一致**（log() 第一行就 return）。
#
# 为什么必须留这个后门：v100 那次"Call 后按 Tab 不补括号"我连修两次都没修好，
# 根因是**每一层判据在离线环境里都显示正常**，而真机跑的是另一条路
# （`get_proc_names` 拿的是 vbe_bridge 内部那份缓存，不是 mock）。
# 没有运行时数据就只能猜，猜错就要用户再等一轮 —— 有了它，
# 下一次一行 `set VBECOMPLETE_LOG=1` 就能把"走到哪一步"直接读出来。
import os as _os

LOG_ENABLED = (_os.environ.get("VBECOMPLETE_LOG", "0").strip() == "1")
_LOG_PATH = _os.environ.get("VBECOMPLETE_LOG_FILE", "").strip() or \
    r"D:\codes\VbeComplete-V587\VbeComplete.log"


def log(*args, **kwargs):
    """写一行诊断日志到 `_LOG_PATH`（仅 VBECOMPLETE_LOG=1 时真正做事）。"""
    if not LOG_ENABLED:
        return                      # 默认路径：立刻返回，零开销
    try:
        line = " ".join(str(a) for a in args)
        if kwargs:
            line += " " + repr(kwargs)
        with open(_LOG_PATH, "a", encoding="utf-8", errors="replace") as f:
            f.write(line[:1200] + "\n")
    except Exception:
        pass                        # 日志失败绝不影响主功能


def log_boot():
    """记录启动信息（同样只在 VBECOMPLETE_LOG=1 时写）。"""
    log("boot: 日志已开启 ->", _LOG_PATH)