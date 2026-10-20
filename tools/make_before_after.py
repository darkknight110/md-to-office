#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 docs/before-after.png：同一份架构信息，转换前（ASCII 字符画）与转换后（真实分层框图）。

为什么值得单独留一个脚本：这张图是本项目最有说服力的**视觉证据**。
它必须能被复跑，而不是存在于某个人的截图文件夹里 —— 否则半年后没人知道它怎么来的。

用法：
    python tools/make_before_after.py [--out docs/before-after.png]

依赖：Pillow（与 diagram.py 相同），无其它第三方依赖。
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import tempfile

from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
DIAGRAM = os.path.join(ROOT, "scripts", "diagram.py")

# 典型的"手写字符画"：多数中文技术文档里就是这么画架构图的。
# 这里用代码生成，保证「中文占 2 格」的宽度算对 —— 这正是它在等宽环境里能对齐的原因，
# 也正是它一进 Word（比例字体 + 自动换行）就散架的原因。
BOX_W = 42


def _is_cjk(ch):
    return "\u2e80" <= ch <= "\u9fff" or "\uff00" <= ch <= "\uffef" or "\u3000" <= ch <= "\u303f"


def _dw(text):
    return sum(2 if _is_cjk(ch) else 1 for ch in text)


def _center(text, width):
    pad = max(0, width - _dw(text))
    left = pad // 2
    return " " * left + text + " " * (pad - left)


def _box(rows, width=BOX_W):
    inner = width - 2
    out = ["┌" + "─" * inner + "┐"]
    out += ["│" + _center(r, inner) + "│" for r in rows]
    out.append("└" + "─" * inner + "┘")
    return out


def build_ascii():
    mid = 1 + (BOX_W - 2) // 2  # 框体中轴所在列
    top = _box(["小程序前端（四种场）", "职场对话 · 落笔改写 · 学堂 · 江湖"])
    core = _box(["成语内核（一套引擎）", "知识库 │ 生产管线 │ 复盘 │ 掌握度", "限流   │ AI 兜底  │ 埋点 │ 会员"])
    # 在框底开一个向下的接点
    top[-1] = top[-1][:mid] + "┬" + top[-1][mid + 1:]
    core[-1] = core[-1][:mid] + "┬" + core[-1][mid + 1:]

    lines = []
    lines += top
    lines.append(" " * mid + "│  HTTPS / JSON")
    lines.append(" " * mid + "▼")
    lines += core
    lines.append(" " * mid + "│")

    # 四路分叉：一条横线，四个接点，四个向下箭头
    slots = [5, 16, 27, 38]
    span = [" "] * (BOX_W + 8)
    for col in range(slots[0], slots[-1] + 1):
        span[col] = "─"
    for s_ in slots:
        span[s_] = "┬"
    span[mid] = "┼"
    lines.append("".join(span).rstrip())

    arrow = [" "] * (BOX_W + 8)
    for s_ in slots:
        arrow[s_] = "▼"
    lines.append("".join(arrow).rstrip())

    row1 = [" "] * (BOX_W + 10)
    row2 = [" "] * (BOX_W + 10)
    pairs = [("AI 能力层", "出题/对话"), ("微信云开发", "云函数/DB"),
             ("合规与安全", "内容安全"), ("变现", "会员/广告")]
    for s_, (a, b) in zip(slots, pairs):
        for i, ch in enumerate(a):
            row1[s_ + i] = ch
        for i, ch in enumerate(b):
            row2[s_ + i] = ch
    lines.append("".join(row1).rstrip())
    lines.append("".join(row2).rstrip())
    return "\n".join(lines)


ASCII_ART = build_ascii()

SPEC = {
    "out": "architecture.png",
    "top": {
        "title": "小程序前端（四种场 + 一层温故）",
        "sub": "职场对话陪练 · 落笔改写 · 学堂词表 · 江湖收集 · 温故连连看",
    },
    "main": {
        "title": "成语内核（一套引擎）",
        "badge": "本方案",
        "modules": [
            ["成语知识库", "场景生产管线", "复盘调度", "用户与掌握度"],
            ["配额与限流", "AI 降级兜底", "埋点与看板", "计费与会员"],
        ],
    },
    "bottoms": [
        {"title": "AI 能力层", "sub": "出题 / 对话 / 语音 / 图像"},
        {"title": "微信云开发", "sub": "云函数 · 数据库 · 存储"},
        {"title": "合规与安全", "sub": "内容安全 · AIGC 标识"},
        {"title": "变现", "sub": "会员 · 课程 · 广告 · ToB"},
    ],
    "options": {"canvas_width": 980},
}

CJK_FONTS = [
    "/System/Library/Fonts/STHeiti Medium.ttc",
    "/System/Library/Fonts/PingFang.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "C:/Windows/Fonts/msyh.ttc",
]
MONO_FONTS = [
    "/System/Library/Fonts/Menlo.ttc",
    "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
    "C:/Windows/Fonts/consola.ttf",
]

PANEL_W, PAD, HEAD_H, GAP = 760, 28, 104, 40
BG = (247, 248, 250)
INK = (28, 32, 38)
MUTED = (110, 118, 128)
RED = (192, 57, 43)
GREEN = (16, 122, 66)


def load(fonts, size):
    for p in fonts:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except OSError:
                continue
    return ImageFont.load_default()


def render_ascii(path, mono, cjk):
    """按"半角 1 格 / 全角 2 格"逐字排格绘制。

    不能整行丢给一个字体：Menlo 没有中文字形（会渲染成豆腐块），
    STHeiti 又不是等宽（框线会错位）。真实文档里之所以能对齐，
    靠的正是"中文占 2 格"这条排版规则 —— 这里如实模拟它。
    """
    lines = ASCII_ART.rstrip("\n").split("\n")
    cell = int(mono.getlength("M"))
    line_h = int(cell * 2.2)
    cols = max(sum(2 if _is_cjk(c) else 1 for c in l) for l in lines)
    w = cols * cell + PAD * 2
    h = line_h * len(lines) + PAD * 2
    img = Image.new("RGB", (w, h), (255, 255, 255))
    d = ImageDraw.Draw(img)
    for i, line in enumerate(lines):
        col = 0
        for ch in line:
            x = PAD + col * cell
            y = PAD + i * line_h
            if _is_cjk(ch):
                bbox = d.textbbox((0, 0), ch, font=cjk)
                cw = bbox[2] - bbox[0]
                d.text((x + (2 * cell - cw) / 2, y), ch, font=cjk, fill=INK)
                col += 2
            else:
                d.text((x, y), ch, font=mono, fill=INK)
                col += 1
    return img


def render_diagram(workdir):
    spec_path = os.path.join(workdir, "spec.json")
    # diagram.py 的正典格式是 {"diagrams": [...]}，单个图也要包一层
    with open(spec_path, "w", encoding="utf-8") as fh:
        json.dump({"diagrams": [SPEC]}, fh, ensure_ascii=False, indent=2)
    out_dir = os.path.join(workdir, "out")
    os.makedirs(out_dir, exist_ok=True)
    subprocess.run(
        [sys.executable, DIAGRAM, "--spec", spec_path, "--out-dir", out_dir],
        check=True,
        stdout=subprocess.DEVNULL,
    )
    return Image.open(os.path.join(out_dir, SPEC["out"])).convert("RGB")


def label(d, x, y, text, color, font):
    d.text((x, y), text, font=font, fill=color)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(ROOT, "docs", "before-after.png"))
    args = ap.parse_args()

    h_font = load(CJK_FONTS, 26)
    small = load(CJK_FONTS, 17)
    mono = load(MONO_FONTS, 15)
    cjk_mono = load(CJK_FONTS, 15)

    with tempfile.TemporaryDirectory() as td:
        left = render_ascii(os.path.join(td, "ascii.txt"), mono, cjk_mono)
        right = render_diagram(td)

        panels = [left, right]
        heights = []
        for p in panels:
            scale = min(1.0, (PANEL_W - PAD * 2) / p.width)
            heights.append(int(p.height * scale))
        body_h = max(heights)
        W = PANEL_W * 2 + GAP + PAD * 2
        H = HEAD_H + body_h + PAD * 2 + 28

        canvas = Image.new("RGB", (W, H), BG)
        d = ImageDraw.Draw(canvas)

        label(d, PAD, 22, "转换前：ASCII 字符画", RED, h_font)
        label(d, PAD, 58, "等宽字体下能对齐；换到 Word / PDF 的比例字体 + 中文占 2 格，必然错位", MUTED, small)
        label(d, PAD + PANEL_W + GAP, 22, "转换后：按 md2office.spec.json 生成的分层框图", GREEN, h_font)
        label(d, PAD + PANEL_W + GAP, 58, "真实 PNG（由 scripts/diagram.py 渲染），替换稿内字符画", MUTED, small)

        for idx, p in enumerate(panels):
            x = PAD + idx * (PANEL_W + GAP)
            y = HEAD_H + (body_h - heights[idx]) // 2
            scale = heights[idx] / p.height
            p2 = p.resize((int(p.width * scale), heights[idx]), Image.LANCZOS)
            border = RED if idx == 0 else GREEN
            d.rectangle([x - 3, y - 3, x + p2.width + 2, y + p2.height + 2], outline=border, width=2)
            canvas.paste(p2, (x, y))

        os.makedirs(os.path.dirname(args.out), exist_ok=True)
        canvas.save(args.out, optimize=True)
        print(f"{args.out}  {canvas.width}x{canvas.height}  {os.path.getsize(args.out)//1024}KB")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
