#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""打包一致性：版本号、入口命令、包内模块必须与 pyproject 同步。

为什么值得测：版本号写两份（__init__ 与 pyproject）是典型会腐烂的地方；
入口命令写错（模块路径拼错）则是"装得上、跑不起来"。这两类都只能靠断言钉住。
"""
import os
import re
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def read(path):
    with open(os.path.join(ROOT, path), encoding="utf-8") as fh:
        return fh.read()


class TestPackaging(unittest.TestCase):
    def test_版本号与_pyproject_一致(self):
        pyproject = read("pyproject.toml")
        version = re.search(r'^version\s*=\s*"([^"]+)"', pyproject, re.M).group(1)
        from md_to_office import __version__
        self.assertEqual(__version__, version)

    def test_入口命令指向真实可调用对象(self):
        pyproject = read("pyproject.toml")
        scripts = re.findall(r'^([\w-]+)\s*=\s*"([\w.]+):(\w+)"', pyproject, re.M)
        self.assertGreaterEqual(len(scripts), 3)
        import importlib
        for _name, module, func in scripts:
            mod = importlib.import_module(module)
            self.assertTrue(callable(getattr(mod, func)), f"{module}:{func} 不可调用")

    def test_包内模块齐全(self):
        import md_to_office
        for mod in ("md2office", "md2docx", "diagram", "archify_bridge", "gongwen"):
            importlib = __import__("importlib")
            importlib.import_module(f"md_to_office.{mod}")


if __name__ == "__main__":
    unittest.main()
