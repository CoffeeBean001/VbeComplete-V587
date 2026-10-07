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
LOG_ENABLED = False


def log(*_args, **_kwargs):
    """空实现：什么都不做（参数原样吞掉，调用点无需判断）。"""


def log_boot():
    """空实现：原先在这里记录启动信息。"""