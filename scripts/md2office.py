#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""兼容入口：真正的实现在 md_to_office/md2office.py（同时支持 pip 安装后的 `md2office` 命令）。

保留这个文件是因为文档、技能说明与既有脚本都按 `scripts/md2office.py` 引用；
删掉会让所有旧用法失效 —— 兼容层比“统一路径”更值钱。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from md_to_office.md2office import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
