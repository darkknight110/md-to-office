#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""archify 桥接：把 archify 渲染出的图接进 md-to-office 的交付流程。

内置的 `diagram.py` 是零依赖的 Pillow 渲染器，负责常规「分层架构图」；
但遇到时序图、ER 图、数据流、状态机这类更复杂的图，`archify` 更合适
（10 种图型，自带真实浏览器校验，输出可交互 HTML）。

本桥接做三件事，全部通过子进程调用 archify CLI，**不 import 它的代码**：
    1. finalize   —— 渲染 + 四道门禁（validate / deliver / check / browser-check）
    2. visual-check --out-dir —— 导出 PNG
    3. 把浅色主题的 PNG 复制到规格里指定的 `out` 路径，供 md2docx 嵌入

archify 是**可选依赖**。没装时本模块抛出带安装指引的 RuntimeError，
不会影响 `diagram.py` 那条默认路径。

定位顺序：环境变量 ARCHIFY_BIN → PATH 上的 archify → 常见技能目录。
"""
from __future__ import annotations

import glob
import json
import os
import shutil
import subprocess
import sys
import tempfile
from typing import Optional

# 与「渲染时无需联网」无关，仅用于定位 CLI
_SKILL_GLOBS = (
    "~/.dsh/skills/archify/bin/archify.mjs",
    "~/.claude/skills/archify/bin/archify.mjs",
    "~/.agents/skills/archify/bin/archify.mjs",
    "~/.codex/skills/archify/bin/archify.mjs",
    "./.dsh/skills/archify/bin/archify.mjs",
    "./.agents/skills/archify/bin/archify.mjs",
)

INSTALL_HINT = (
    "找不到 archify。安装方式任选其一：\n"
    "  1) npx skills add tt-a1i/archify -g\n"
    "  2) 把 archify 技能目录放到 ~/.dsh/skills/archify/（或 ~/.claude/skills/ 等）\n"
    "  3) export ARCHIFY_BIN=/绝对路径/bin/archify.mjs\n"
    "也可以不用 archify：把规格里的 \"tool\": \"archify\" 去掉，即回退到内置 Pillow 渲染器。"
)


def find_archify() -> Optional[list]:
    """返回可直接执行的 argv 前缀，如 ['node', '/path/bin/archify.mjs']；找不到返回 None。"""
    explicit = os.environ.get("ARCHIFY_BIN")
    if explicit:
        if not os.path.exists(explicit):
            raise RuntimeError("ARCHIFY_BIN 指向的文件不存在：" + explicit)
        return _wrap(explicit)

    onpath = shutil.which("archify")
    if onpath:
        return [onpath]

    for pat in _SKILL_GLOBS:
        for hit in sorted(glob.glob(os.path.expanduser(pat))):
            return _wrap(hit)
    return None


def _wrap(script: str) -> list:
    """把 .mjs 包成 `node <script>`；可执行文件则直接调用。"""
    if script.endswith((".mjs", ".js")):
        node = shutil.which("node") or sys.executable
        return [node, script]
    return [script]


def _env() -> dict:
    """渲染子进程环境。

    ARCHIFY_CHROME_NO_SANDBOX=1 是 archify 官方支持的开关：在容器/受限沙箱里
    Chrome 起不来时必须打开。若用户已显式设置则不覆盖。
    """
    env = dict(os.environ)
    env.setdefault("ARCHIFY_CHROME_NO_SANDBOX", "1")
    return env


def render(archify_spec: dict, out_path: str, quality: str = "showcase",
           timeout: int = 300) -> tuple:
    """渲染一张 archify 图并导出 PNG 到 out_path。

    archify_spec: 完整的 archify JSON-IR（schema_version / diagram_type / meta / ...），
                  其中 meta.output 会被本函数重写为临时 HTML 路径。
    返回 (out_path, width, height)。
    """
    argv = find_archify()
    if not argv:
        raise RuntimeError(INSTALL_HINT)

    out_path = os.path.abspath(out_path)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="md2office-archify-") as work:
        # archify 要求 meta.output 是「相对当前工作目录」的路径，
        # 因此所有子进程都以 work 为 cwd，并一律传相对路径。
        spec = json.loads(json.dumps(archify_spec))          # 深拷贝，不改调用方对象
        meta = spec.setdefault("meta", {})
        meta["output"] = "diagram.html"
        with open(os.path.join(work, "candidate.json"), "w", encoding="utf-8") as fh:
            json.dump(spec, fh, ensure_ascii=False, indent=1)

        dtype = spec.get("diagram_type")
        if not dtype:
            raise RuntimeError("archify 规格缺少 diagram_type")

        # 1) finalize：渲染 + 四道门禁
        r = subprocess.run(
            [*argv, "finalize", dtype, "candidate.json", "diagram.html",
             "--quality", quality, "--json"],
            capture_output=True, text=True, env=_env(), cwd=work, timeout=timeout)
        if r.returncode != 0:
            raise RuntimeError(_explain(r))

        # 2) visual-check：导出 PNG（含浅色主题）
        r2 = subprocess.run(
            [*argv, "visual-check", "diagram.html", "--out-dir", "shots", "--json"],
            capture_output=True, text=True, env=_env(), cwd=work, timeout=timeout)
        if r2.returncode != 0:
            raise RuntimeError(_explain(r2, stage="visual-check"))

        shot_dir = os.path.join(work, "shots")
        png = _pick_light_png(shot_dir)
        if not png:
            listing = ", ".join(sorted(os.listdir(shot_dir))) \
                if os.path.isdir(shot_dir) else "（无目录）"
            raise RuntimeError("archify 未产出 PNG。visual-check 输出目录内容：" + listing)

        shutil.copyfile(png, out_path)
        size = _png_size(out_path)
        return out_path, size[0], size[1]


def _pick_light_png(shot_dir: str) -> Optional[str]:
    """优先取尺寸最大的浅色主题 PNG（便于嵌入 Word）。"""
    if not os.path.isdir(shot_dir):
        return None
    pngs = glob.glob(os.path.join(shot_dir, "*.png"))
    light = [p for p in pngs if ".light." in os.path.basename(p)]
    pool = light or pngs
    if not pool:
        return None
    return max(pool, key=lambda p: (os.path.getsize(p), p))


def _png_size(path: str) -> tuple:
    """从 PNG 头部读宽高（IHDR），无需 Pillow。"""
    import struct
    with open(path, "rb") as fh:
        head = fh.read(24)
    if len(head) >= 24 and head[:8] == b"\x89PNG\r\n\x1a\n":
        return struct.unpack(">II", head[16:24])
    return (0, 0)


def _explain(result, stage: str = "finalize") -> str:
    """把 archify 的 JSON 回执变成一句人话，保留原始输出兜底。"""
    raw = (result.stdout or "").strip() or (result.stderr or "").strip()
    try:
        d = json.loads(result.stdout)
    except Exception:                                        # noqa: BLE001
        return "archify %s 失败（exit %s）：%s" % (stage, result.returncode, raw[:600])

    gates = d.get("gates") or {}
    failed = d.get("failedStage") or [k for k, v in gates.items() if v != "pass"]
    msgs = [str(g.get("message")) for g in (d.get("diagnostics") or [])][:3]
    hint = ""
    if any("browser" in str(g.get("code", "")) for g in (d.get("diagnostics") or [])):
        hint = ("\n提示：browser-check 失败通常是 Chrome 起不来。"
                "本桥接已默认设 ARCHIFY_CHROME_NO_SANDBOX=1；"
                "若仍失败，检查本机是否装了 Chrome/Chromium。")
    return "archify %s 失败：gate=%s，失败阶段=%s%s%s" % (
        stage, gates, failed, ("；" + " / ".join(msgs)) if msgs else "", hint)


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser(description="用 archify 渲染一张图并导出 PNG")
    ap.add_argument("--spec", required=True, help="archify JSON-IR 规格文件")
    ap.add_argument("--out", required=True, help="导出的 PNG 路径")
    ap.add_argument("--quality", default="showcase", choices=["standard", "showcase"])
    a = ap.parse_args(argv)
    with open(a.spec, encoding="utf-8") as fh:
        spec = json.load(fh)
    path, w, h = render(spec, a.out, a.quality)
    print("  ✓ %s  %s×%s" % (path, w, h))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
