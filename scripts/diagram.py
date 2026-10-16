#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""分层架构图生成器（数据驱动）

读 JSON 规格，输出 PNG。用于替代 Markdown 里的 ASCII 字符画——字符画在
Word/PDF 中会因字体替换而错位。

规格示例见 spec.example.json：

    python diagram.py --spec spec.json [--out-dir DIR]

单个图规格字段：
    out        输出 PNG 路径（相对 --out-dir）
    top        {title, sub}            顶部平台层
    main       {title, badge, modules} 主体层；modules 为二维数组（每行一组模块）
    bottoms    [{title, sub}]          底部对接对象
    options    可选：canvas_width / 各字号 / 配色
"""
from __future__ import annotations

import argparse
import json
import os
import sys

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:  # pragma: no cover
    print("需要 Pillow：请使用 DSH 随附 Python（含 Pillow）", file=sys.stderr)
    raise

# 中文字体候选（macOS / Linux / Windows）
FONT_CANDIDATES = {
    "heading": [
        "/System/Library/Fonts/STHeiti Medium.ttc",
        "/System/Library/Fonts/Hiragino Sans GB.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        "C:/Windows/Fonts/msyhbd.ttc",
        "C:/Windows/Fonts/msyh.ttc",
    ],
    "body": [
        "/System/Library/Fonts/Hiragino Sans GB.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simhei.ttf",
    ],
}

DEFAULTS = {
    "canvas_width": 1960,
    "title_size": 40,
    "sub_size": 30,
    "mod_size": 32,
    "stop_size": 30,
    "stop_sub_size": 24,
    "badge_size": 26,
    "l1_height": 150,
    "gap1": 70,
    "main_title_height": 92,
    "mod_height": 104,
    "mod_gap": 22,
    "mod_pad": 34,
    "gap2": 70,
    "stop_height": 150,
    "margin": 60,
    "l1_indent": 150,
}

COLORS = {
    "title_bg": (31, 56, 100),
    "title_fg": (255, 255, 255),
    "l1_fill": (237, 242, 250),
    "l1_bord": (142, 170, 219),
    "mod_fill": (222, 234, 246),
    "mod_bord": (46, 84, 150),
    "stop_fill": (255, 242, 224),
    "stop_bord": (197, 137, 60),
    "arrow": (46, 84, 150),
    "text": (32, 32, 32),
    "sub": (90, 100, 115),
    "badge": (224, 122, 39),
}


def find_font(kind: str) -> str:
    for p in FONT_CANDIDATES[kind]:
        if os.path.exists(p):
            return p
    raise FileNotFoundError(
        "找不到中文字体。请在 FONT_CANDIDATES 中补充本机字体路径，"
        "或用 --font 指定一个 .ttf/.ttc 文件。")


class Renderer:
    def __init__(self, spec: dict, font_dir_override: str | None = None):
        o = spec.get("options", {})
        self.o = {**DEFAULTS, **o}
        self.c = {**COLORS, **{k: tuple(v) for k, v in (o.get("colors") or {}).items()}}
        heading = font_dir_override or find_font("heading")
        body = font_dir_override or find_font("body")
        self.f_heading = heading
        self.f_body = body
        self._cache: dict = {}

    def F(self, path, size):
        key = (path, size)
        if key not in self._cache:
            self._cache[key] = ImageFont.truetype(path, size)
        return self._cache[key]

    # ── 绘制原语 ──

    def ctext(self, d, cx, cy, s, font, fill):
        bb = d.textbbox((0, 0), s, font=font)
        d.text((cx - (bb[2] - bb[0]) / 2 - bb[0], cy - (bb[3] - bb[1]) / 2 - bb[1]),
               s, font=font, fill=fill)

    def wrap(self, d, s, font, max_w):
        lines, cur = [], ""
        for ch in s:
            test = cur + ch
            bb = d.textbbox((0, 0), test, font=font)
            if bb[2] - bb[0] > max_w and cur:
                lines.append(cur)
                cur = ch
            else:
                cur = test
        if cur:
            lines.append(cur)
        return lines

    def ctext_block(self, d, cx, y_top, lines, font, fill, gap=6):
        h = d.textbbox((0, 0), "字", font=font)
        lh = (h[3] - h[1]) + gap
        y = y_top
        for ln in lines:
            self.ctext(d, cx, y + lh / 2, ln, font, fill)
            y += lh
        return y - y_top

    def rbox(self, d, x, y, w, h, fill, outline, radius=14, width=3):
        d.rounded_rectangle([x, y, x + w, y + h], radius=radius,
                            fill=fill, outline=outline, width=width)

    def arrow_down(self, d, x, y1, y2, width=5):
        d.line([x, y1, x, y2], fill=self.c["arrow"], width=width)
        d.polygon([(x - 13, y2 - 18), (x + 13, y2 - 18), (x, y2)], fill=self.c["arrow"])

    def badge(self, d, x, y, s, font):
        bb = d.textbbox((0, 0), s, font=font)
        w, h = bb[2] - bb[0] + 36, bb[3] - bb[1] + 22
        d.rounded_rectangle([x, y, x + w, y + h], radius=h / 2, fill=self.c["badge"])
        self.ctext(d, x + w / 2, y + h / 2, s, font, (255, 255, 255))

    # ── 主流程 ──

    def render(self, spec: dict, out_path: str):
        o = self.o
        M = o["margin"]
        cw = o["canvas_width"]
        top = spec.get("top") or {}
        main = spec.get("main") or {}
        mods = main.get("modules") or []
        bottoms = spec.get("bottoms") or []

        l1_h, gap1 = o["l1_height"], o["gap1"]
        mb_title, mod_h, mod_gap, mod_pad = (o["main_title_height"], o["mod_height"],
                                             o["mod_gap"], o["mod_pad"])
        gap2, stop_h = o["gap2"], o["stop_height"]

        l2_h = mb_title + len(mods) * (mod_h + mod_gap) + mod_pad
        H = M + l1_h + gap1 + l2_h + gap2 + stop_h + M
        img = Image.new("RGB", (cw, H), "white")
        d = ImageDraw.Draw(img)

        f_title = self.F(self.f_heading, o["title_size"])
        f_sub = self.F(self.f_body, o["sub_size"])
        f_mod = self.F(self.f_body, o["mod_size"])
        f_stop = self.F(self.f_body, o["stop_size"])
        f_stopsub = self.F(self.f_body, o["stop_sub_size"])
        f_badge = self.F(self.f_body, o["badge_size"])

        # 顶部平台层
        x1, y1 = M + o["l1_indent"], M
        w1 = cw - 2 * (M + o["l1_indent"])
        self.rbox(d, x1, y1, w1, l1_h, self.c["l1_fill"], self.c["l1_bord"])
        if top.get("title"):
            has_sub = bool(top.get("sub"))
            self.ctext(d, x1 + w1 / 2, y1 + (l1_h * 0.36 if has_sub else l1_h / 2),
                       top["title"], f_title, self.c["title_bg"])
            if has_sub:
                sl = self.wrap(d, top["sub"], f_sub, w1 - 60)
                self.ctext_block(d, x1 + w1 / 2, y1 + l1_h * 0.62, sl, f_sub, self.c["sub"])

        cx = cw / 2
        self.arrow_down(d, cx, y1 + l1_h, y1 + l1_h + gap1)

        # 主体层
        y2, x2 = y1 + l1_h + gap1, M
        w2 = cw - 2 * M
        self.rbox(d, x2, y2, w2, l2_h, (250, 252, 255), self.c["mod_bord"],
                  radius=18, width=4)
        d.rounded_rectangle([x2, y2, x2 + w2, y2 + mb_title], radius=18,
                            fill=self.c["title_bg"])
        d.rectangle([x2, y2 + mb_title - 20, x2 + w2, y2 + mb_title],
                    fill=self.c["title_bg"])
        if main.get("title"):
            self.ctext(d, x2 + w2 / 2, y2 + mb_title / 2, main["title"],
                       f_title, self.c["title_fg"])
        if main.get("badge"):
            self.badge(d, x2 + w2 - 210, y2 + 22, main["badge"], f_badge)

        for r, row in enumerate(mods):
            n = max(len(row), 1)
            total_gap = mod_gap * (n - 1)
            bw = (w2 - 2 * mod_pad - total_gap) / n
            by = y2 + mb_title + mod_pad / 2 + r * (mod_h + mod_gap)
            for c, name in enumerate(row):
                bx = x2 + mod_pad + c * (bw + mod_gap)
                self.rbox(d, bx, by, bw, mod_h, self.c["mod_fill"],
                          self.c["mod_bord"], radius=10, width=3)
                ml = self.wrap(d, name, f_mod, bw - 20)
                mh = d.textbbox((0, 0), "字", font=f_mod)[3] + 6
                self.ctext_block(d, bx + bw / 2, by + (mod_h - len(ml) * mh) / 2,
                                 ml, f_mod, self.c["text"])

        # 底部对接对象
        y3 = y2 + l2_h
        ns = max(len(bottoms), 1)
        span = w2 - 2 * mod_pad
        step = span / ns
        for i in range(ns):
            self.arrow_down(d, x2 + mod_pad + step * (i + 0.5), y3, y3 + gap2, width=4)

        y4 = y3 + gap2
        for i, b in enumerate(bottoms):
            bx = x2 + mod_pad + step * i + 6
            bw = step - 12
            self.rbox(d, bx, y4, bw, stop_h, self.c["stop_fill"],
                      self.c["stop_bord"], radius=10, width=3)
            nl = self.wrap(d, b.get("title", ""), f_stop, bw - 24)
            hh = d.textbbox((0, 0), "字", font=f_stop)[3] + 6
            if b.get("sub"):
                sl = self.wrap(d, b["sub"], f_stopsub, bw - 24)
                sh = d.textbbox((0, 0), "字", font=f_stopsub)[3] + 6
                total = len(nl) * hh + len(sl) * sh + 8
                y0 = y4 + (stop_h - total) / 2
                used = self.ctext_block(d, bx + bw / 2, y0, nl, f_stop, self.c["text"])
                self.ctext_block(d, bx + bw / 2, y0 + used + 8, sl, f_stopsub,
                                 self.c["sub"])
            else:
                self.ctext_block(d, bx + bw / 2,
                                 y4 + (stop_h - len(nl) * hh) / 2, nl, f_stop,
                                 self.c["text"])

        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        img.save(out_path, dpi=(220, 220))
        return out_path, img.width, img.height


def generate(spec: dict, out_dir: str, font_override: str | None = None) -> list:
    """按规格生成全部图，返回 [(path, w, h), ...]"""
    r = Renderer(spec, font_override)
    results = []
    for item in spec.get("diagrams", []):
        out = os.path.join(out_dir, item.get("out") or f"{item.get('name','diagram')}.png")
        results.append(r.render(item, out))
    return results


def main(argv=None):
    ap = argparse.ArgumentParser(description="分层架构图生成器（数据驱动）")
    ap.add_argument("--spec", required=True, help="JSON 规格文件")
    ap.add_argument("--out-dir", default=".", help="输出根目录（规格中的 out 相对此目录）")
    ap.add_argument("--font", help="强制使用指定字体文件")
    a = ap.parse_args(argv)
    spec = json.load(open(a.spec, encoding="utf-8"))
    for path, w, h in generate(spec, a.out_dir, a.font):
        print(f"  ✓ {path}  {w}×{h}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
