# -*- coding: utf-8 -*-
"""可选的诊断日志。

默认**完全关闭**，日常使用不产生任何文件、没有任何开销。

需要排查"某个变量为什么不提示"时，运行 `run_debug.bat`（它会设置
环境变量 VBECOMPLETE_LOG=1），日志写到程序目录下的 `VbeComplete.log`，
内容涵盖：收集到的模块、光标位置、光标所属过程、候选命中的原始记录
（含作用域）——足以区分"压根没收录"还是"收录了但被作用域过滤掉"。

不需要时直接删除本文件与 `log(...) / LOG_ENABLED` 调用即可，核心逻辑不依赖它。
"""

import os
import time

LOG_ENABLED = os.environ.get("VBECOMPLETE_LOG", "0").strip() == "1"
LOG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "VbeComplete.log")
_MAX_BYTES = 4 * 1024 * 1024
_START = time.time()


def log_boot():
    """启动时记录「实际运行的目录」与各源文件时间。

    排查"改了文件却没生效"时这是最有用的一行：一眼就能看出程序跑的是
    哪个目录、里面的 .py 是不是刚替换过的新版本。
    """
    if not LOG_ENABLED:
        return
    here = os.path.dirname(os.path.abspath(__file__))
    log("boot: VBECOMPLETE_LOG 已开启，日志=%s" % LOG_PATH)
    log("boot: 运行目录=%s" % here)
    for name in ("main.py", "parser.py", "engine.py", "vbe_bridge.py",
                 "ui.py", "log.py"):
        path = os.path.join(here, name)
        try:
            log("  %-14s mtime=%s size=%d"
                % (name,
                   time.strftime("%Y-%m-%d %H:%M:%S",
                                 time.localtime(os.path.getmtime(path))),
                   os.path.getsize(path)))
        except Exception:
            log("  %-14s <缺失>" % name)


def log(msg):
    """写一行日志；未开启时立即返回（近乎零开销）。"""
    if not LOG_ENABLED:
        return
    try:
        try:
            if os.path.exists(LOG_PATH) and os.path.getsize(LOG_PATH) > _MAX_BYTES:
                os.remove(LOG_PATH)
        except Exception:
            pass
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write("[%8.2f] %s\n" % (time.time() - _START, msg))
    except Exception:
        pass
