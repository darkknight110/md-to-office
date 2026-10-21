"""md-to-office：Markdown → 中文 Word/PDF，以及公文模板引擎。"""

__version__ = "1.0.0"

import sys as _sys  # noqa: E402


def utf8_stdio() -> None:
    """Windows 控制台默认不是 UTF-8，输出 ✓/✗ 会 UnicodeEncodeError。

    这不是"显示难看"，是**直接崩**。所有 CLI 入口先调它；失败就退化为 replace，不阻断。
    """
    for stream in (_sys.stdout, _sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:                      # noqa: BLE001
            pass
