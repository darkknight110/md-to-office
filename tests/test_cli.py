#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""CLI 契约测试：版本、覆盖保护。

为什么单测这两件：
  · `--version` 缺了，用户报 bug 时第一句永远是"你装的哪版"；
  · 静默覆盖交付物，等于把上一版悄悄删掉。
"""
import hashlib
import os
import subprocess
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SCRIPT = os.path.join(ROOT, "scripts", "md2office.py")


def run(*args):
    return subprocess.run([sys.executable, SCRIPT, *args], capture_output=True, text=True)


class TestCliContract(unittest.TestCase):
    def test_doctor_不转换只体检(self):
        r = run("--doctor")
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn("python-docx", r.stdout)
        self.assertIn("LibreOffice", r.stdout)

    def test_json_输出可解析(self):
        import json
        with tempfile.TemporaryDirectory() as td:
            src = os.path.join(td, "a.md")
            open(src, "w", encoding="utf-8").write("# 标题\n\n正文。\n")
            out = os.path.join(td, "out")
            r = run("--input", src, "--output", out, "--format", "docx", "--json")
            self.assertEqual(r.returncode, 0, r.stderr)
            payload = json.loads(r.stdout)
            self.assertEqual(payload["counts"]["docx"], 1)
            self.assertEqual(payload["files"][0]["status"], "ok")
            self.assertEqual(payload["files"][0]["docx"], "a.docx")

    def test_version(self):
        r = run("--version")
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertRegex(r.stdout.strip(), r"^md-to-office \d+\.\d+\.\d+")

    def test_no_overwrite_keeps_existing_file(self):
        with tempfile.TemporaryDirectory() as td:
            src = os.path.join(td, "a.md")
            open(src, "w", encoding="utf-8").write("# 标题\n\n正文。\n")
            out = os.path.join(td, "out")
            first = run("--input", src, "--output", out, "--format", "docx")
            self.assertEqual(first.returncode, 0, first.stderr)
            target = os.path.join(out, "a.docx")
            before = hashlib.md5(open(target, "rb").read()).hexdigest()

            second = run("--input", src, "--output", out, "--format", "docx", "--no-overwrite")
            self.assertEqual(second.returncode, 0, second.stderr)
            self.assertIn("跳过", second.stdout)
            after = hashlib.md5(open(target, "rb").read()).hexdigest()
            self.assertEqual(before, after, "加了 --no-overwrite 却改动了文件")

    def test_overwrite_warns(self):
        with tempfile.TemporaryDirectory() as td:
            src = os.path.join(td, "a.md")
            open(src, "w", encoding="utf-8").write("# 标题\n\n正文。\n")
            out = os.path.join(td, "out")
            run("--input", src, "--output", out, "--format", "docx")
            again = run("--input", src, "--output", out, "--format", "docx")
            self.assertIn("覆盖已有文件", again.stdout)


if __name__ == "__main__":
    unittest.main()
