#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""中文文档排版量测：从 PDF / DOCX 里量出**可核对的参数**，而不是凭印象说"看着还行"。

量什么（每份文档）：
  · 页面尺寸（mm）
  · 页边距（mm，按正文文本包围盒估计）
  · 正文字号（pt，按字符数加权后取众数）
  · 行距（pt，同一文本块内基线间距的中位数）
  · 每页行数、每行字数（中文字符计 2 列宽折算）
  · 标题字号（pt，正文之上最大的两档）

为什么单独做成脚本：基线要能被复跑。别人 clone 之后用同一批文件跑一遍，
应当得到同一组数；「规范」不是写死的结论，而是一条可重复的测量流水线。

用法：
    python tools/measure_typography.py --out baseline.json a.pdf b.docx dir/
    python tools/measure_typography.py --json-only a.pdf

依赖：pdfplumber（仅 PDF）、python-docx（仅 DOCX）。两者都缺时只跳过对应格式。
"""
from __future__ import annotations

import argparse
import json
import os
import statistics
import sys
from collections import Counter
from dataclasses import dataclass, field, asdict
from typing import Iterable

MM_PER_PT = 25.4 / 72.0


@dataclass
class DocStat:
    path: str
    kind: str
    pages: int = 0
    page_size_mm: tuple | None = None
    margins_mm: dict = field(default_factory=dict)
    body_pt: float | None = None
    line_pitch_pt: float | None = None
    chars_per_line: float | None = None
    lines_per_page: float | None = None
    heading_pt: list = field(default_factory=list)
    samples: int = 0
    warnings: list = field(default_factory=list)


def _is_cjk(ch: str) -> bool:
    return "\u3000" <= ch <= "\u9fff" or "\uff00" <= ch <= "\uffef"


def _disp_width(text: str) -> int:
    return sum(2 if _is_cjk(ch) else 1 for ch in text)


def measure_pdf(path: str) -> DocStat:
    st = DocStat(path=path, kind="pdf")
    try:
        import pdfplumber
    except ImportError:
        st.warnings.append("未安装 pdfplumber，跳过 PDF 量测")
        return st

    sizes = Counter()
    pitches = []
    margins = {"top": [], "bottom": [], "left": [], "right": []}
    per_page_lines = []
    per_page_widths = []
    heading_sizes = Counter()

    with pdfplumber.open(path) as pdf:
        st.pages = len(pdf.pages)
        for page in pdf.pages:
            st.page_size_mm = (round(page.width * MM_PER_PT, 1), round(page.height * MM_PER_PT, 1))
            chars = page.chars or []
            if not chars:
                continue
            # 页边距：字符包围盒到纸张边缘的距离
            x0 = min(c["x0"] for c in chars) * MM_PER_PT
            x1 = max(c["x1"] for c in chars) * MM_PER_PT
            # pdfplumber 的 top/bottom 都是"距页面顶端的距离"：
            # 上边距 = 最小的 top；下边距 = 页高 - 最大的 bottom（原先写反了，量出来的是对侧白边）
            top = min(c["top"] for c in chars) * MM_PER_PT
            bottom = (page.height - max(c["bottom"] for c in chars)) * MM_PER_PT
            margins["left"].append(x0)
            margins["right"].append(page.width * MM_PER_PT - x1)
            margins["top"].append(top)
            margins["bottom"].append(bottom)

            # 正文候选：字号按字符数加权（标题字少，会被自然压下去）
            for c in chars:
                sizes[round(c["size"], 1)] += 1

            # 行分组：同一视觉行的字符 bottom 会差零点几 pt，按 1pt 归并
            lines = {}
            for c in chars:
                key = round(c["bottom"])
                lines.setdefault(key, []).append(c)
            keys = sorted(lines)
            # 行距：只取"整行宽"的相邻行（表格单元格、页码会污染间距）
            widths = {}
            for k, row in lines.items():
                row_sorted = sorted(row, key=lambda c: c["x0"])
                widths[k] = _disp_width("".join(c["text"] for c in row_sorted))
            max_w = max(widths.values()) if widths else 0
            body_keys = [k for k in keys if widths.get(k, 0) >= max_w * 0.6]
            deltas = [body_keys[i + 1] - body_keys[i] for i in range(len(body_keys) - 1)]
            deltas = [d for d in deltas if 6 <= d <= 60]
            if deltas:
                pitches.append(statistics.median(deltas))
            per_page_lines.append(len(body_keys))
            # 每页取"整行宽"的中位数（正文行），避免短行/表格把均值拉低
            body_widths = [widths[k] for k in body_keys if widths.get(k, 0) > 0]
            if body_widths:
                per_page_widths.append(statistics.median(body_widths))
            st.samples += len(chars)

            # 标题字号：比正文大一档以上的常见字号
            for c in chars:
                if c["size"] >= 12:
                    heading_sizes[round(c["size"], 1)] += 1

    if sizes:
        st.body_pt = sizes.most_common(1)[0][0]
        if heading_sizes:
            st.heading_pt = sorted({s for s in heading_sizes if s > st.body_pt}, reverse=True)[:3]
    if pitches:
        st.line_pitch_pt = round(statistics.median(pitches), 1)
    per_page_lines = [n for n in per_page_lines if n > 3]
    if per_page_lines:
        st.lines_per_page = round(statistics.median(per_page_lines), 1)
    widths = [w for w in per_page_widths if w > 20]
    if widths:
        st.chars_per_line = round(statistics.median(widths) / 2, 1)  # 折算为中文字数（1 字 = 2 列）
    st.margins_mm = {k: round(statistics.median(v), 1) for k, v in margins.items() if v}
    return st


def measure_docx(path: str) -> DocStat:
    st = DocStat(path=path, kind="docx")
    try:
        import docx
        from docx.shared import Length  # noqa: F401  仅用于类型说明
    except ImportError:
        st.warnings.append("未安装 python-docx，跳过 DOCX 量测")
        return st

    d = docx.Document(path)
    if d.sections:
        sec = d.sections[0]
        st.page_size_mm = (round(sec.page_width.mm, 1), round(sec.page_height.mm, 1))
        st.margins_mm = {
            "top": round(sec.top_margin.mm, 1),
            "bottom": round(sec.bottom_margin.mm, 1),
            "left": round(sec.left_margin.mm, 1),
            "right": round(sec.right_margin.mm, 1),
        }
    sizes = Counter()
    heading_sizes = Counter()

    def size_of(run, para):
        """字号可能来自 4 个地方：run 直设 / 字符样式 / 段落样式 / Normal。

        踩过的坑：只读 run.font.size 时，**用段落样式排版的模板**全部读不到字号
        （产物被报成 None）—— 而那恰恰是最规范的一类模板。
        """
        for cand in (
            run.font.size,
            getattr(getattr(run, "style", None), "font", None) and run.style.font.size,
            getattr(getattr(para, "style", None), "font", None) and para.style.font.size,
        ):
            if cand:
                return round(cand.pt, 1)
        try:
            normal = d.styles["Normal"]
            if normal.font.size:
                return round(normal.font.size.pt, 1)
        except KeyError:
            pass
        return None

    for p in d.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        for r in p.runs:
            if not r.text.strip():
                continue
            pt = size_of(r, p)
            if pt is None:
                continue
            sizes[pt] += len(r.text)
            if p.style and p.style.name.lower().startswith("heading"):
                heading_sizes[pt] += 1
    if sizes:
        st.body_pt = sizes.most_common(1)[0][0]
        st.heading_pt = sorted({s for s in heading_sizes if s > st.body_pt}, reverse=True)[:3]
    st.pages = 0
    st.warnings.append("DOCX 无版面，行距/每页行数取 Normal 样式声明值")
    try:
        normal = d.styles["Normal"]
        ls = normal.paragraph_format.line_spacing
        if isinstance(ls, float):
            st.line_pitch_pt = round(ls * (st.body_pt or 10.5) * 1.2, 1)
        elif ls is not None:
            st.line_pitch_pt = round(ls.pt, 1)
    except Exception:
        pass
    return st


def collect_paths(inputs: Iterable[str]) -> list[str]:
    out = []
    for item in inputs:
        if os.path.isdir(item):
            for root, _dirs, files in os.walk(item):
                for f in files:
                    if f.lower().endswith((".pdf", ".docx")) and not f.startswith("~$"):
                        out.append(os.path.join(root, f))
        elif os.path.isfile(item):
            out.append(item)
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="量测中文文档排版参数")
    ap.add_argument("inputs", nargs="+", help="PDF/DOCX 文件或目录")
    ap.add_argument("--out", help="输出 JSON 路径")
    ap.add_argument("--json-only", action="store_true", help="只打 JSON，不打人类可读摘要")
    args = ap.parse_args()

    paths = collect_paths(args.inputs)
    if not paths:
        print("没有找到 .pdf / .docx 文件", file=sys.stderr)
        return 2

    stats = []
    for p in paths:
        st = measure_pdf(p) if p.lower().endswith(".pdf") else measure_docx(p)
        stats.append(st)

    usable = [s for s in stats if s.body_pt]
    agg = {}
    if usable:
        def med(attr):
            vals = [getattr(s, attr) for s in usable if getattr(s, attr)]
            return round(statistics.median(vals), 1) if vals else None
        agg = {
            "documents": len(usable),
            "body_pt_median": med("body_pt"),
            "line_pitch_pt_median": med("line_pitch_pt"),
            "chars_per_line_median": med("chars_per_line"),
            "lines_per_page_median": med("lines_per_page"),
        }
        for side in ("top", "bottom", "left", "right"):
            vals = [s.margins_mm[side] for s in usable if side in s.margins_mm]
            if vals:
                agg.setdefault("margins_mm_median", {})[side] = round(statistics.median(vals), 1)

    payload = {"aggregate": agg, "documents": [asdict(s) for s in stats]}
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as fh:
            fh.write(text + "\n")
    if not args.json_only:
        print(f"量测 {len(stats)} 份，其中 {len(usable)} 份有可读文本")
        for k, v in agg.items():
            if isinstance(v, dict):
                print(f"  {k}: " + ", ".join(f"{kk}={vv}" for kk, vv in v.items()))
            else:
                print(f"  {k}: {v}")
        for s in stats:
            if s.warnings:
                print(f"  ! {os.path.basename(s.path)}: {'; '.join(s.warnings)}")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
