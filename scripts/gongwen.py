#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""兼容入口：实现在 md_to_office/gongwen.py（pip 安装后也可直接调用包内模块）。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from md_to_office.gongwen import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
