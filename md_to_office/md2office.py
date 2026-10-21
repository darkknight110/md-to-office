#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""md2office —— 统一入口：Markdown → Word / PDF（可选架构图生成与字符画替换）

    python md2office.py --input <文件或目录> --output <目录> --format docx|pdf|both

常用参数
    --format       docx / pdf / both（默认 both）
    --only GLOB    只处理匹配的文件（可重复）
    --spec FILE    规格文件（架构图 + 字符画替换规则）；不给则不做预处理
    --no-prep      强制跳过预处理
    --keep-intermediate  pdf-only 时保留中间 .docx
    --font-body / --font-heading / --font-code / --size-body / --size-table
    --page a4|letter
    --no-footer    不输出页码

规格文件（可选）示例：
{
  "diagrams": [ { "out": "diagrams/arch.png", "top": {...}, "main": {...}, "bottoms": [...] } ],
  "blocks":   [ { "contains": ["关键词A","关键词B"], "image": "diagrams/arch.png" } ],
  "defence_numbered_blocks": true
}
"""
from __future__ import annotations

import argparse
import fnmatch
import glob
import json
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.dont_write_bytecode = True        # 技能目录可能只读，避免生成 __pycache__
from . import diagram as diagram_mod      # noqa: E402
from . import md2docx as engine           # noqa: E402

BOX_CHARS = set("┌│└├┐┘┤─▼▲┬┴┼")


# ─────────────────────── LibreOffice Kit 定位 ───────────────────────

def find_libreoffice_kit():
    """返回 (node, cli)；找不到返回 (None, None)。"""
    node = os.environ.get("LIBREOFFICE_KIT_NODE")
    cli = os.environ.get("LIBREOFFICE_KIT_CLI")
    if node and cli and os.path.exists(node) and os.path.exists(cli):
        return node, cli

    # DSH 随附 Node
    if not node:
        for pat in (os.path.expanduser("~/.dsh/dsh-runtimes/*/dependencies/node/bin/node"),
                    "/Applications/DeepSeek Harness.app/Contents/Resources/runtime/"
                    "primary-runtime/dependencies/node/bin/node"):
            hit = sorted(glob.glob(pat))
            if hit:
                node = hit[-1]
                break
    # LibreOffice Kit CLI
    if not cli:
        for pat in (
            "/Applications/DeepSeek Harness.app/Contents/Resources/app.asar.unpacked/"
            "dsh/node_modules/@deepseek-ai/libreoffice-kit/lib/cli.js",
            os.path.expanduser("~/.dsh/**/libreoffice-kit/lib/cli.js"),
        ):
            hit = sorted(glob.glob(pat, recursive=True))
            if hit:
                cli = hit[0]
                break
    if node and cli and os.path.exists(node) and os.path.exists(cli):
        return node, cli
    return None, None


def find_soffice():
    """返回系统 LibreOffice 的 soffice 可执行文件路径；找不到返回 None。

    显式指定优先：环境变量 LIBREOFFICE_PATH 或 SOFFICE_PATH（可执行文件路径，
    或 PATH 上的命令名）。显式指定了却找不到会直接报错，避免静默回退到别的后端。
    """
    explicit = os.environ.get("LIBREOFFICE_PATH") or os.environ.get("SOFFICE_PATH")
    if explicit:
        if os.path.isfile(explicit):
            return explicit
        found = shutil.which(explicit)
        if found:
            return found
        raise RuntimeError(
            "LIBREOFFICE_PATH / SOFFICE_PATH 指向的可执行文件不存在：" + explicit)

    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found

    for c in (
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
        "/opt/homebrew/bin/soffice",
        "/snap/bin/libreoffice",
        "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
        "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
    ):
        if os.path.isfile(c):
            return c
    return None


def _pdf_via_soffice(soffice, docx_path, pdf_path):
    """用 LibreOffice 的 soffice --headless 转 PDF（通用后端，无第三方依赖）。"""
    outdir = os.path.dirname(os.path.abspath(pdf_path))
    os.makedirs(outdir, exist_ok=True)
    # 用一次性 profile，避免污染/抢占用户正在使用的 LibreOffice 实例
    with tempfile.TemporaryDirectory(prefix="md2office-lo-") as profile:
        r = subprocess.run(
            [soffice, "--headless", "--norestore", "--invisible", "--nolockcheck",
             "-env:UserInstallation=file://" + profile,
             "--convert-to", "pdf", "--outdir", outdir, docx_path],
            capture_output=True, text=True)
    if r.returncode != 0 or not os.path.exists(pdf_path):
        raise RuntimeError(
            "soffice 返回 %s：%s" % (r.returncode, (r.stderr or r.stdout).strip()[:400]))
    return pdf_path


def docx_to_pdf(docx_path, pdf_path):
    """DOCX → PDF。依次尝试通用 soffice 与 DSH libreoffice-kit 两个后端。

    通用 LibreOffice 优先（任何机器装了都能用）；DSH 随附的 libreoffice-kit
    作为回退，保证在 DSH 环境内仍然开箱可用。
    """
    errors = []

    soffice = find_soffice()
    if soffice:
        try:
            return _pdf_via_soffice(soffice, docx_path, pdf_path)
        except Exception as exc:                      # noqa: BLE001
            errors.append("soffice(%s)：%s" % (soffice, exc))

    node, cli = find_libreoffice_kit()
    if node and cli:
        os.makedirs(os.path.dirname(os.path.abspath(pdf_path)), exist_ok=True)
        r = subprocess.run([node, cli, "convert", "--input", docx_path, "--output", pdf_path],
                           capture_output=True, text=True)
        if r.returncode == 0 and os.path.exists(pdf_path):
            return pdf_path
        errors.append("libreoffice-kit：%s" % (r.stderr or r.stdout).strip()[:400])

    raise RuntimeError(
        "找不到可用的 PDF 转换后端。任选一种即可：\n"
        "  1) 安装 LibreOffice（会自动探测 soffice），"
        "或用 LIBREOFFICE_PATH 指向 soffice 可执行文件；\n"
        "  2) 设置 LIBREOFFICE_KIT_NODE 与 LIBREOFFICE_KIT_CLI（DSH 随附的 libreoffice-kit）。\n"
        "  仅需 Word 时请用 --format docx，无需任何 PDF 后端。"
        + ("\n尝试记录：\n  - " + "\n  - ".join(errors) if errors else ""))


# ─────────────────────── 预处理（字符画 → 图片 / 去围栏） ───────────────────────

def classify_block(block_lines, blocks_rules, defence_numbered):
    body = "\n".join(block_lines)
    if any(c in BOX_CHARS for c in body):
        for rule in blocks_rules:
            keys = rule.get("contains") or []
            if keys and all(k in body for k in keys):
                return "img", rule["image"]
        if blocks_rules:
            # 有规则但都没命中：保留原样，避免误替换
            return "keep", None
        return "keep", None
    if defence_numbered and len(block_lines) >= 3 and \
            __import__("re").match(r"^\s*1[.、]\s", block_lines[0]):
        return "list", None
    return "keep", None


def prepare(md_path, out_md, spec, images_abs_dir):
    """把 md 复制到 out_md，按规格替换字符画块 / 去围栏编号流程。"""
    blocks_rules = spec.get("blocks") or []
    defence = bool(spec.get("defence_numbered_blocks"))
    lines = open(md_path, encoding="utf-8").read().replace("\r\n", "\n").split("\n")
    out, i, n = [], 0, len(lines)
    stats = {"img": 0, "list": 0, "keep": 0}
    while i < n:
        if lines[i].strip().startswith("```"):
            j, buf = i + 1, []
            while j < n and not lines[j].strip().startswith("```"):
                buf.append(lines[j])
                j += 1
            kind, arg = classify_block(buf, blocks_rules, defence)
            stats[kind] += 1
            if kind == "img":
                img_abs = os.path.join(images_abs_dir, arg)
                out += ["", f"![架构图]({img_abs})", ""]
            elif kind == "list":
                out += [""] + buf + [""]
            else:
                out += lines[i:j + 1]
            i = j + 1
        else:
            out.append(lines[i])
            i += 1
    os.makedirs(os.path.dirname(os.path.abspath(out_md)), exist_ok=True)
    open(out_md, "w", encoding="utf-8").write("\n".join(out))
    return stats


# ─────────────────────── 主流程 ───────────────────────

def collect_md(input_path, only):
    if os.path.isfile(input_path):
        return [input_path], os.path.dirname(os.path.abspath(input_path))
    root = os.path.abspath(input_path)
    files = []
    for dirpath, _dirs, names in os.walk(root):
        for nm in sorted(names):
            if not nm.endswith(".md"):
                continue
            if only and not any(fnmatch.fnmatch(nm, g) for g in only):
                continue
            files.append(os.path.join(dirpath, nm))
    return sorted(files), root


def main(argv=None):
    ap = argparse.ArgumentParser(description="Markdown → Word / PDF 统一转换")
    ap.add_argument("--input", required=True, help="输入的 .md 文件或目录（递归）")
    ap.add_argument("--output", required=True, help="输出目录")
    ap.add_argument("--format", choices=["docx", "pdf", "both"], default="both")
    ap.add_argument("--only", action="append", default=[], help="文件名通配，可重复")
    ap.add_argument("--spec", help="规格 JSON（架构图 + 字符画替换规则）")
    ap.add_argument("--no-prep", action="store_true", help="跳过预处理")
    ap.add_argument("--keep-intermediate", action="store_true",
                    help="pdf-only 时保留中间 .docx")
    ap.add_argument("--keep-prepared", action="store_true", help="保留预处理后的 md")
    ap.add_argument("--font-body", help="正文中文字体（默认 宋体）")
    ap.add_argument("--font-heading", help="标题中文字体（默认 微软雅黑）")
    ap.add_argument("--font-code", help="代码字体（默认 Consolas）")
    ap.add_argument("--size-body", type=float)
    ap.add_argument("--size-table", type=float)
    ap.add_argument("--page", choices=["a4", "letter"], default="a4")
    ap.add_argument("--no-footer", action="store_true")
    a = ap.parse_args(argv)

    cfg = engine.Config()
    cfg.apply_overrides(
        body_ea=a.font_body, heading_ea=a.font_heading, code_font=a.font_code,
        body_size=a.size_body, table_size=a.size_table)
    if a.page == "letter":
        cfg.page_width, cfg.page_height = 21.59, 27.94
    if a.no_footer:
        cfg.page_number_footer = False

    out_root = os.path.abspath(a.output)
    os.makedirs(out_root, exist_ok=True)

    # 规格：未指定则尝试自动发现
    spec_path = a.spec
    if not spec_path and not a.no_prep:
        base = (os.path.abspath(a.input) if os.path.isdir(a.input)
                else os.path.dirname(os.path.abspath(a.input)))
        # 依次在「输入目录」与「其上一级」查找——便于把规格放在项目根、而按子目录转换
        for d in (base, os.path.dirname(base)):
            for cand in ("md2office.spec.json", "office.spec.json"):
                p = os.path.join(d, cand)
                if os.path.exists(p):
                    spec_path = p
                    break
            if spec_path:
                break
    spec = json.load(open(spec_path, encoding="utf-8")) if spec_path else {}

    # 1) 生成架构图（若有）
    if spec.get("diagrams") and not a.no_prep:
        print("生成架构图：")
        for item in spec["diagrams"]:
            out_rel = item.get("out") or f"{item.get('name', 'diagram')}.png"
            if str(item.get("tool", "")).lower() == "archify":
                # 可选后端：时序图/ER 图/数据流等复杂图型，交给 archify
                try:
                    from . import archify_bridge
                    path, w, h = archify_bridge.render(
                        item["archify"], os.path.join(out_root, out_rel),
                        item.get("quality", "showcase"))
                    print(f"  ✓ {path}  {w}×{h}  (archify)")
                except Exception as exc:                  # noqa: BLE001
                    print(f"  ✗ {out_rel}  archify 渲染失败：{exc}", file=sys.stderr)
                    return 1
            else:
                # 默认后端：内置 Pillow 渲染器（零外部依赖）
                for path, w, h in diagram_mod.generate({"diagrams": [item]}, out_root):
                    print(f"  ✓ {path}  {w}×{h}")

    files, in_root = collect_md(a.input, a.only)
    if not files:
        print("没有找到匹配的 Markdown 文件", file=sys.stderr)
        return 1

    tmp = tempfile.mkdtemp(prefix="md2office-")
    prepared_dir = os.path.join(tmp, "prepared")
    images_abs_dir = out_root
    total = {"docx": 0, "pdf": 0}

    try:
        for src in files:
            rel = os.path.relpath(src, in_root if os.path.isdir(a.input) else
                                  os.path.dirname(src))
            stem = os.path.splitext(os.path.basename(rel))[0]
            sub = os.path.dirname(rel)
            dest_dir = os.path.join(out_root, sub) if sub else out_root

            # 2) 预处理
            work_md = src
            if spec and not a.no_prep:
                tgt = os.path.join(prepared_dir, rel)
                st = prepare(src, tgt, spec, images_abs_dir)
                work_md = tgt
                if st["img"] or st["list"]:
                    print(f"  预处理 {os.path.basename(src)}: "
                          f"替换图片 {st['img']} 处、去围栏 {st['list']} 处")

            # 3) DOCX
            docx_path = os.path.join(dest_dir, stem + ".docx")
            if a.format == "pdf" and not a.keep_intermediate:
                tmp_docx = os.path.join(tmp, "docx", sub, stem + ".docx")
                engine.convert(work_md, tmp_docx, cfg)
                docx_for_pdf = tmp_docx
                print(f"  · {rel} → (中间 DOCX)")
            else:
                engine.convert(work_md, docx_path, cfg)
                docx_for_pdf = docx_path
                total["docx"] += 1
                print(f"  ✓ {os.path.join(sub, stem + '.docx')}")

            # 4) PDF
            if a.format in ("pdf", "both"):
                pdf_path = os.path.join(dest_dir, stem + ".pdf")
                docx_to_pdf(docx_for_pdf, pdf_path)
                total["pdf"] += 1
                print(f"  ✓ {os.path.join(sub, stem + '.pdf')}")

        if a.keep_prepared and os.path.isdir(prepared_dir):
            shutil.copytree(prepared_dir, os.path.join(out_root, "_prepared"),
                            dirs_exist_ok=True)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    print(f"\n完成：格式={a.format}  DOCX {total['docx']} 个，PDF {total['pdf']} 个")
    print(f"输出目录：{out_root}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
