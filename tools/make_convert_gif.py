#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成"ASCII 字符画 → 真框图"的演示动画（GIF + 故事板 PNG）。

证据来源：
  · 左栏的字符画就是方案里的原文；
  · 右栏"转换后"的图由 `scripts/diagram.py` 按 `examples/demo/convert.spec.json` **真的渲染**出来，
    不是画上去的示意图。

清晰度与 devflow 演示同一套策略：2 倍超采样 + LANCZOS 缩回、少而大、低抖动调色板；
另出 PNG 故事板，README 首屏用它，点开才是动图。

用法：
    python3 tools/make_convert_gif.py --out docs/demo-convert.gif
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import tempfile

from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

LW, LH = 1100, 680
SCALE = 1.2
SS = 2
K = SCALE * SS
FPS = 12

BG = (243, 245, 248)
INK = (28, 34, 42)
DIM = (110, 120, 132)
RED = (192, 57, 43)
GREEN = (22, 122, 66)
BLUE = (37, 99, 235)
CARD = (255, 255, 255)
BORDER = (214, 220, 228)
CODE_BG = (250, 251, 253)

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

BOX_W = 32


def _is_cjk(ch):
    return "\u2e80" <= ch <= "\u9fff" or "\uff00" <= ch <= "\uffef" or "\u3000" <= ch <= "\u303f"


def _cw(text):
    return sum(2 if _is_cjk(ch) else 1 for ch in text)


def _center(text, width):
    pad = max(0, width - _cw(text))
    left = pad // 2
    return " " * left + text + " " * (pad - left)


def _box(rows, width=BOX_W):
    inner = width - 2
    out = ["┌" + "─" * inner + "┐"]
    out += ["│" + _center(r, inner) + "│" for r in rows]
    out.append("└" + "─" * inner + "┘")
    return out


def build_ascii():
    mid = 1 + (BOX_W - 2) // 2
    top = _box(["小程序前端（四种场）", "职场对话 · 落笔改写 · 学堂"])
    core = _box(["成语内核（一套引擎）", "知识库 │ 生产管线 │ 复盘", "限流   │ AI 兜底  │ 埋点"])
    top[-1] = top[-1][:mid] + "┬" + top[-1][mid + 1:]
    core[-1] = core[-1][:mid] + "┬" + core[-1][mid + 1:]
    lines = top + [" " * mid + "│ HTTPS", " " * mid + "▼"] + core + [" " * mid + "│"]
    slots = [4, 12, 20, 26]
    span = [" "] * (BOX_W + 6)
    for c in range(slots[0], slots[-1] + 1):
        span[c] = "─"
    for sl in slots:
        span[sl] = "┬"
    span[mid] = "┼"
    lines.append("".join(span).rstrip())
    arrow = [" "] * (BOX_W + 6)
    for sl in slots:
        arrow[sl] = "▼"
    lines.append("".join(arrow).rstrip())
    row = [" "] * (BOX_W + 8)
    labels = [("AI 能力层", 4), ("微信云开发", 11), ("合规安全", 21), ("变现", 30)]
    for text, pos in labels:
        for i, ch in enumerate(text):
            row[pos + i] = ch
    lines.append("".join(row).rstrip())
    return "\n".join(lines)


ASCII_ART = build_ascii()

PHASES = [
    ("① 稿子里是 ASCII 字符画：等宽字体下能对齐", 16),
    ("② 转成 Word：比例字体 + 中文占 2 格 —— 必然错位", 34),
    ("③ md2office 按 spec 生成真图，替换掉字符画", 30),
    ("④ 交付物里是一张真图，不再是字体碰运气", 22),
]


def load(fonts, size):
    for p in fonts:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, int(size * K))
            except OSError:
                continue
    return ImageFont.load_default()


def is_cjk(ch):
    return "\u2e80" <= ch <= "\u9fff" or "\uff00" <= ch <= "\uffef" or "\u3000" <= ch <= "\u303f"


def grid_text(d, text, x, y, mono, cjk, fill):
    cell = mono.getlength("M")
    col = 0.0
    for ch in text:
        if is_cjk(ch):
            w = cjk.getlength(ch)
            d.text(((x * K) + col * cell + (2 * cell - w) / 2, (y - 1) * K), ch, font=cjk, fill=fill)
            col += 2
        else:
            d.text(((x * K) + col * cell, y * K), ch, font=mono, fill=fill)
            col += 1


def card(d, x, y, w, h, fill=CARD, outline=BORDER):
    d.rounded_rectangle([x * K, y * K, (x + w) * K, (y + h) * K], radius=12 * K,
                        fill=fill, outline=outline, width=max(1, int(K)))


def render_diagram(workdir: str) -> Image.Image:
    spec = os.path.join(ROOT, "examples", "demo", "convert.spec.json")
    out_dir = os.path.join(workdir, "diagram")
    os.makedirs(out_dir, exist_ok=True)
    subprocess.run([sys.executable, os.path.join(ROOT, "scripts", "diagram.py"),
                    "--spec", spec, "--out-dir", out_dir], check=True, stdout=subprocess.DEVNULL)
    img = Image.open(os.path.join(out_dir, "arch.png")).convert("RGB")
    # 裁掉渲染留白：否则贴进页面后图很小、字看不清
    from PIL import ImageChops
    bg = Image.new("RGB", img.size, (255, 255, 255))
    bbox = ImageChops.difference(img, bg).getbbox()
    return img.crop(bbox) if bbox else img


def broken_lines(font, max_w):
    """把同一段字符画按**比例字体**排版并折行 —— 就是 Word 里看到的样子。"""
    out = []
    for line in ASCII_ART.rstrip("\n").split("\n"):
        cur = ""
        for ch in line:
            if font.getlength(cur + ch) > max_w:
                out.append(cur)
                cur = ch
            else:
                cur += ch
        out.append(cur)
    return out


FONTS = {}


def frame(phase_idx, progress, diagram):
    """progress ∈ [0,1]：本阶段内部的推进度（用于逐行出现 / 擦除 / 淡入）。"""
    img = Image.new("RGB", (int(LW * K), int(LH * K)), BG)
    d = ImageDraw.Draw(img)
    f = FONTS

    d.text((40 * K, 24 * K), "md-to-office · ASCII 图 → 真框图", font=f["h1"], fill=INK)
    d.text((40 * K, 58 * K), "左：方案原文　右：导出后的 Word 页面", font=f["small"], fill=DIM)

    # 两栏
    card(d, 40, 96, 470, 420)
    card(d, 590, 96, 470, 420)

    grid_text(d, "方案.md", 58, 110, f["label_mono"], f["label"], DIM)
    grid_text(d, "导出/方案.docx", 608, 110, f["label_mono"], f["label"], DIM)

    # 左：ASCII 原样（等宽）
    y = 142
    for line in ASCII_ART.rstrip("\n").split("\n"):
        grid_text(d, line, 58, y, f["mono"], f["cjk"], INK)
        y += 20

    # 右：先坏后好
    box = [608, 142, 1044, 500]
    if phase_idx == 0:
        d.text((620 * K, 150 * K), "（等待转换…）", font=f["small"], fill=(170, 178, 188))
    elif phase_idx == 1:
        # 比例字体 + 折行：错位
        lines = broken_lines(f["prop"], 420)
        show = max(1, int(len(lines) * progress))
        yy = 142
        for line in lines[:show]:
            d.text((608 * K, yy * K), line, font=f["prop"], fill=INK)
            yy += 21
        if yy > 470:
            break_mark = True  # noqa: F841  （视觉上溢出即说明问题）
    else:
        # 真图：擦除式出现
        w = 430
        h = int(diagram.height * w / diagram.width)
        small = diagram.resize((int(w * K), int(h * K)), Image.LANCZOS)
        top = 142
        cut = int(h * K * min(1.0, progress))
        region = small.crop((0, 0, small.width, max(1, cut)))
        img.paste(region, (int(608 * K), int(top * K)))
        d.rectangle([int((608 + 2) * K), int((top + h + 6) * K), int((608 + 418) * K), int((top + h + 10) * K)], fill=BORDER)

    # 阶段说明 + 进度
    caption, _ = PHASES[min(phase_idx, len(PHASES) - 1)]
    color = RED if phase_idx == 2 else (GREEN if phase_idx >= 3 else INK)
    d.text((40 * K, 540 * K), caption, font=f["caption"], fill=color)

    # 阶段点
    x = 40
    for i, (cap, _n) in enumerate(PHASES):
        fill = GREEN if i < phase_idx else (BLUE if i == phase_idx else (210, 216, 224))
        d.ellipse([x * K, 584 * K, (x + 9) * K, (588 + 5) * K], fill=fill)
        x += 22
    grid_text(d, "由 scripts/diagram.py 真实渲染；复跑：python3 tools/make_convert_gif.py",
              40, 610, f["tiny_mono"], f["tiny"], (140, 150, 160))
    return img.resize((int(LW * SCALE), int(LH * SCALE)), Image.LANCZOS)


def storyboard(frames, out_png):
    picks = [frames[15], frames[49], frames[79], frames[-1]]
    w, h = picks[0].size
    gap = int(12 * SCALE)
    canvas = Image.new("RGB", (w * 2 + gap, h * 2 + gap), (231, 235, 240))
    for i, im in enumerate(picks):
        canvas.paste(im, ((i % 2) * (w + gap), (i // 2) * (h + gap)))
    canvas.save(out_png, optimize=True)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    FONTS.update({
        "h1": load(CJK_FONTS, 24),
        "small": load(CJK_FONTS, 14),
        "label": load(CJK_FONTS, 14),
        "label_mono": load(MONO_FONTS, 14),
        "mono": load(MONO_FONTS, 12),
        "cjk": load(CJK_FONTS, 12),
        "prop": load(CJK_FONTS, 14),
        "caption": load(CJK_FONTS, 17),
        "tiny": load(CJK_FONTS, 11),
        "tiny_mono": load(MONO_FONTS, 11),
    })

    with tempfile.TemporaryDirectory() as td:
        diagram = render_diagram(td)
        frames = []
        for idx, (_cap, n) in enumerate(PHASES):
            for i in range(n):
                frames.append(frame(idx, (i + 1) / n, diagram))

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    base = frames[-1].quantize(colors=96, method=Image.MEDIANCUT)
    q = [f.quantize(palette=base, dither=Image.NONE) for f in frames]
    q[0].save(args.out, save_all=True, append_images=q[1:], duration=int(1000 / FPS), loop=0,
              optimize=True, disposal=2)
    poster = os.path.splitext(args.out)[0] + ".png"
    frames[-1].save(poster, optimize=True)
    storyboard(frames, os.path.splitext(args.out)[0] + "-story.png")
    print(f"{args.out}  {len(frames)} 帧  {os.path.getsize(args.out)/1048576:.2f}MB  {frames[0].size}")
    print(poster)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
