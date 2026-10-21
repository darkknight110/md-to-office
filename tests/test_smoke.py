#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""md-to-office 冒烟测试。

只依赖标准库 unittest 与被测代码自身的依赖（python-docx / Pillow），
不需要 pytest：

    python -m unittest discover -s tests -v
    # 或
    python tests/test_smoke.py

覆盖四条主线：
  1. Markdown → DOCX 真的能生成
  2. ASCII 字符画的「命中则替换 / 未命中则保留」两条分支
  3. 数据驱动架构图真的能渲染出 PNG
  4. PDF 后端的通用回退链（soffice → libreoffice-kit）与报错信息
"""
from __future__ import annotations

import os
import shutil
import stat
import subprocess
import sys
import tempfile
import unittest
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SCRIPTS = os.path.join(ROOT, "scripts")
sys.path.insert(0, ROOT)

from md_to_office import md2office  # noqa: E402
from md_to_office import diagram as diagram_mod  # noqa: E402
from md_to_office import archify_bridge  # noqa: E402

# 生成架构图需要中文字体；最小化 Linux 镜像默认没有（CI 里要先装 fonts-noto-cjk）。
# 没有字体时跳过相关用例，而不是让整套测试红 —— 否则"环境缺字体"会被误读成"代码坏了"。
def _has_cjk_font() -> bool:
    try:
        diagram_mod.find_font("heading")
        diagram_mod.find_font("body")
        return True
    except Exception:
        return False


HAS_CJK_FONT = _has_cjk_font()
NO_FONT_REASON = "本机无中文字体（Linux 请装 fonts-noto-cjk）"

PY = sys.executable
MD2OFFICE = os.path.join(SCRIPTS, "md2office.py")

SAMPLE_MD = """# 测试文档

## 一、架构

```
┌───────────────┐
│   顶层平台     │
└───────────────┘
        ▼
┌───────────────┐
│  核心业务系统  │
└───────────────┘
```

## 二、表格

| 项 | 说明 |
| --- | --- |
| 甲 | 第一条 |
"""

SAMPLE_SPEC = {
    "diagrams": [{
        "out": "diagrams/arch.png",
        "top": {"title": "顶层平台", "sub": "驾驶舱"},
        "main": {"title": "核心业务系统", "badge": "本方案",
                 "modules": [["模块一", "模块二", "模块三"]]},
        "bottoms": [{"title": "设备接入", "sub": "传感器"}],
    }],
    "blocks": [{"contains": ["顶层平台", "核心业务系统"],
                "image": "diagrams/arch.png"}],
}


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="md2office-test-")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.indir = os.path.join(self.tmp, "in")
        os.makedirs(self.indir)
        self.md = os.path.join(self.indir, "doc.md")
        with open(self.md, "w", encoding="utf-8") as fh:
            fh.write(SAMPLE_MD)

    def run_cli(self, *args):
        return subprocess.run([PY, MD2OFFICE, *args],
                              capture_output=True, text=True)


class TestDocx(Base):
    def test_docx_is_generated(self):
        out = os.path.join(self.tmp, "out")
        r = self.run_cli("--input", self.indir, "--output", out, "--format", "docx")
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        docx = os.path.join(out, "doc.docx")
        self.assertTrue(os.path.exists(docx), "未生成 docx")
        self.assertGreater(os.path.getsize(docx), 1000)
        # 必须是合法 OOXML 包
        with zipfile.ZipFile(docx) as z:
            self.assertIn("word/document.xml", z.namelist())

    def test_docx_only_needs_no_pdf_backend(self):
        """--format docx 不应依赖任何 PDF 后端。"""
        out = os.path.join(self.tmp, "out")
        env = dict(os.environ)
        env["LIBREOFFICE_PATH"] = "/nonexistent/soffice"
        r = subprocess.run(
            [PY, MD2OFFICE, "--input", self.indir, "--output", out, "--format", "docx"],
            capture_output=True, text=True, env=env)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)


class TestAsciiBlockReplacement(Base):
    @unittest.skipUnless(HAS_CJK_FONT, NO_FONT_REASON)
    def test_hit_replaces_block_with_image(self):
        spec_path = os.path.join(self.indir, "spec.json")
        import json
        with open(spec_path, "w", encoding="utf-8") as fh:
            json.dump(SAMPLE_SPEC, fh, ensure_ascii=False)

        out = os.path.join(self.tmp, "out")
        r = self.run_cli("--input", self.indir, "--output", out,
                         "--format", "docx", "--spec", spec_path)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertTrue(os.path.exists(os.path.join(out, "diagrams", "arch.png")),
                        "架构图 PNG 未生成")
        # 预处理后的 md 应含图片引用、且原字符画块已消失
        candidates = [os.path.join(self.tmp, p) for p in
                      ("prepared/doc.md", "out/prepared/doc.md")]
        prepared = next((c for c in candidates if os.path.exists(c)), None)
        if prepared:
            body = open(prepared, encoding="utf-8").read()
            self.assertIn("![架构图]", body)
            self.assertNotIn("┌───────────────┐", body)

    def test_miss_keeps_block_verbatim(self):
        """关键词未全命中时必须保留原文，绝不误替换。"""
        lines = ["┌───┐", "│ x │", "└───┘"]
        kind, _ = md2office.classify_block(
            lines, [{"contains": ["不存在的关键词A", "不存在的关键词B"],
                     "image": "x.png"}], False)
        self.assertEqual(kind, "keep")

    def test_hit_requires_all_keywords(self):
        lines = ["┌───┐", "│ 顶层平台 │", "└───┘"]
        rules = [{"contains": ["顶层平台", "核心业务系统"], "image": "x.png"}]
        kind, _ = md2office.classify_block(lines, rules, False)
        self.assertEqual(kind, "keep", "只命中部分关键词时不应替换")

    def test_no_box_chars_is_never_replaced(self):
        lines = ["just text", "no box drawing"]
        kind, _ = md2office.classify_block(
            lines, [{"contains": ["text"], "image": "x.png"}], False)
        self.assertEqual(kind, "keep")


class TestDiagram(Base):
    @unittest.skipUnless(HAS_CJK_FONT, NO_FONT_REASON)
    def test_diagram_renders_png(self):
        outdir = os.path.join(self.tmp, "diagrams")
        os.makedirs(outdir)
        spec = {"diagrams": [{"out": "arch.png",
                              "top": {"title": "顶层平台", "sub": "驾驶舱"},
                              "main": {"title": "核心业务系统", "badge": "本方案",
                                       "modules": [["模块一", "模块二", "模块三"]]},
                              "bottoms": [{"title": "设备接入", "sub": "传感器"}]}]}
        results = diagram_mod.generate(spec, outdir)
        png = os.path.join(outdir, "arch.png")
        self.assertTrue(os.path.exists(png), "未生成 PNG")
        self.assertGreater(os.path.getsize(png), 1000)
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0][0], png)
        with open(png, "rb") as fh:
            self.assertEqual(fh.read(8), b"\x89PNG\r\n\x1a\n")


class TestPdfBackend(unittest.TestCase):
    """PDF 后端回退链：通用 soffice 优先，libreoffice-kit 兜底。"""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="md2office-pdf-")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.docx = os.path.join(self.tmp, "in.docx")
        with open(self.docx, "wb") as fh:
            fh.write(b"PK\x03\x04stub")

    def _fake_soffice(self):
        """造一个假 soffice：按 --outdir 产出 <stem>.pdf。"""
        path = os.path.join(self.tmp, "fake-soffice")
        with open(path, "w", encoding="utf-8") as fh:
            fh.write("#!/bin/sh\n"
                     "outdir='.'\n"
                     "while [ $# -gt 0 ]; do\n"
                     "  case \"$1\" in\n"
                     "    --outdir) outdir=\"$2\"; shift 2;;\n"
                     "    *) src=\"$1\"; shift;;\n"
                     "  esac\n"
                     "done\n"
                     "base=$(basename \"$src\" .docx)\n"
                     "printf '%%PDF-1.4 fake' > \"$outdir/$base.pdf\"\n")
        os.chmod(path, os.stat(path).st_mode | stat.S_IEXEC | stat.S_IXGRP)
        return path

    def test_soffice_backend_is_used_when_available(self):
        fake = self._fake_soffice()
        pdf = os.path.join(self.tmp, "out", "in.pdf")
        os.makedirs(os.path.dirname(pdf))
        old = os.environ.copy()
        try:
            os.environ["LIBREOFFICE_PATH"] = fake
            result = md2office.docx_to_pdf(self.docx, pdf)
        finally:
            os.environ.clear()
            os.environ.update(old)
        self.assertEqual(result, pdf)
        self.assertTrue(os.path.exists(pdf))
        with open(pdf, "rb") as fh:
            self.assertTrue(fh.read().startswith(b"%PDF"))

    def test_explicit_missing_soffice_raises(self):
        old = os.environ.copy()
        try:
            os.environ["LIBREOFFICE_PATH"] = "/definitely/not/here/soffice"
            with self.assertRaises(RuntimeError) as ctx:
                md2office.find_soffice()
        finally:
            os.environ.clear()
            os.environ.update(old)
        self.assertIn("不存在", str(ctx.exception))

    def test_error_lists_both_options_when_no_backend(self):
        """两个后端都没有时，报错必须同时给出 soffice 与 kit 两条路。"""
        old_s, old_k = md2office.find_soffice, md2office.find_libreoffice_kit
        md2office.find_soffice = lambda: None
        md2office.find_libreoffice_kit = lambda: (None, None)
        try:
            with self.assertRaises(RuntimeError) as ctx:
                md2office.docx_to_pdf(self.docx, os.path.join(self.tmp, "x.pdf"))
        finally:
            md2office.find_soffice, md2office.find_libreoffice_kit = old_s, old_k
        msg = str(ctx.exception)
        self.assertIn("soffice", msg)
        self.assertIn("LIBREOFFICE_KIT_NODE", msg)
        self.assertIn("--format docx", msg)


def _archify_available():
    """archify 是可选依赖，装在才会跑那两条端到端用例。"""
    try:
        return archify_bridge.find_archify() is not None
    except RuntimeError:
        return False


ARCHIFY_AVAILABLE = _archify_available()


class TestArchifyBridge(unittest.TestCase):
    """archify 是可选依赖：没装时这些用例应被跳过，而不是失败。"""

    ARCHIFY_SPEC = {
        "schema_version": 2,
        "diagram_type": "workflow",
        "meta": {"title": "订单处理流程", "locale": "zh",
                 "quality_profile": "showcase", "output": "x.html"},
        "lanes": [{"id": "customer", "label": "顾客"},
                  {"id": "platform", "label": "平台"}],
        "nodes": [
            {"id": "order", "lane": "customer", "col": 0, "type": "external",
             "label": "下单", "sublabel": "小程序"},
            {"id": "pay", "lane": "platform", "col": 1, "type": "backend",
             "label": "支付", "sublabel": "网关"},
        ],
        "edges": [{"id": "e1", "from": "order", "to": "pay",
                   "label": "确认", "variant": "emphasis"}],
    }

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="md2office-arch-")
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def test_missing_archify_raises_with_install_hint(self):
        old = archify_bridge.find_archify
        archify_bridge.find_archify = lambda: None
        try:
            with self.assertRaises(RuntimeError) as ctx:
                archify_bridge.render(self.ARCHIFY_SPEC,
                                      os.path.join(self.tmp, "x.png"))
        finally:
            archify_bridge.find_archify = old
        self.assertIn("archify", str(ctx.exception))
        self.assertIn("ARCHIFY_BIN", str(ctx.exception))

    def test_png_size_reader(self):
        png = os.path.join(self.tmp, "t.png")
        with open(png, "wb") as fh:
            fh.write(b"\x89PNG\r\n\x1a\n" + b"\x00\x00\x00\rIHDR"
                     + (800).to_bytes(4, "big") + (600).to_bytes(4, "big"))
        self.assertEqual(archify_bridge._png_size(png), (800, 600))

    @unittest.skipUnless(ARCHIFY_AVAILABLE, "未安装 archify（可选依赖，跳过）")
    def test_archify_renders_png_end_to_end(self):
        out = os.path.join(self.tmp, "flow.png")
        path, w, h = archify_bridge.render(self.ARCHIFY_SPEC, out)
        self.assertEqual(path, out)
        self.assertTrue(os.path.exists(out))
        self.assertGreater(os.path.getsize(out), 1000)
        self.assertGreater(w, 0)
        self.assertGreater(h, 0)
        with open(out, "rb") as fh:
            self.assertEqual(fh.read(8), b"\x89PNG\r\n\x1a\n")


if __name__ == "__main__":
    unittest.main(verbosity=2)
