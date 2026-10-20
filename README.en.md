# md-to-office — Markdown → Chinese Word / PDF

Convert Markdown into properly typeset Chinese **Word (`.docx`)** and/or **PDF**, and — optionally — replace the
ASCII box diagrams inside the manuscript with **real layered PNG diagrams** generated from a JSON spec.

![before / after](docs/before-after.png)

<sub>Left: what the manuscript looks like (it only lines up in a monospace font). Right: the real PNG that lands in
the document, rendered by `scripts/diagram.py` from `md2office.spec.json`. Regenerate with `python3 tools/make_before_after.py`.</sub>

Two ways to use it: **as a CLI** (any project), or **as an Agent Skill** (Claude Code / Codex / DSH …) so your
agent calls it automatically. Built for Chinese technical proposals, reports and client deliverables.

> Word-only output needs **no external binary** — just `python-docx` and `Pillow`.

---

## Install in one line

```bash
git clone https://github.com/darkknight110/md-to-office.git ~/.md-to-office \
  && python3 -m pip install -q -r ~/.md-to-office/requirements.txt \
  && python3 ~/.md-to-office/scripts/md2office.py --help
```

Then:

```bash
python3 ~/.md-to-office/scripts/md2office.py --input ./docs --output ./out --format docx
```

`--format` accepts `docx`, `pdf` or `both`. PDF requires LibreOffice (`soffice`) on `PATH`.

---

## Why it exists

Chinese Word documents have a failure mode that plain Markdown converters ignore: **the ASCII architecture diagram**.

```
┌──────────────────────────┐
│      Operations platform │
└───────────┬──────────────┘
```

It aligns only in a monospace editor. Word uses proportional fonts, and CJK glyphs occupy two columns — so the
diagram collapses into a ragged mess the moment anyone opens the file. Most tools leave it broken, or silently
reflow it into an unreadable table.

This tool takes the diagram out of the text layer entirely: you describe the layers in `md2office.spec.json`, it
renders a real PNG (`scripts/diagram.py`) and **substitutes it for the ASCII block** in the output. Nothing to fix
by hand; the diagram is an asset, not a font accident.

---

## API surface (stable enough to depend on)

| Entry point | Purpose |
|---|---|
| `scripts/md2office.py` | Unified driver: batch input dir, format selection, diagram substitution |
| `scripts/md2docx.py` | DOCX engine: styles, tables, code blocks, page numbers, TOC |
| `scripts/diagram.py` | JSON spec → layered PNG |
| `scripts/archify_bridge.py` | Optional: hand complex figures to the `archify` skill |

`python3 -m unittest discover -s tests -v` runs the smoke suite (DOCX generation, diagram hit/miss branches,
PDF fallback chain).

---

## Typography is measured, not guessed

The defaults (10.5 pt body, 1.15 line spacing, 22/22/20/20 mm margins) are not taste — they are the measured values
of **52 real deliverables** produced with this tool, all identical. The gap to the Chinese government document
standard (GB/T 9704-2012: 16 pt, 28 pt line pitch, 37/35/28/26 mm) is quantified in
[docs/typography-baseline.md](docs/typography-baseline.md), with machine-readable data in
[docs/typography-baseline.json](docs/typography-baseline.json).

Re-measure your own corpus:

```bash
python3 tools/measure_typography.py --out my-baseline.json ./my-docs
```

---

## Known limits

1. **PDF needs LibreOffice.** Without it, `--format pdf` fails with an explicit message (it never silently emits a
   Word file with a `.pdf` name).
2. **CJK fonts are assumed.** The default style asks for 宋体 / 微软雅黑 / Consolas; on systems without them the
   renderer substitutes, and metrics may shift.
3. **The diagram spec is declarative.** It draws layered boxes, not arbitrary graphs — for sequence/ER/state
   diagrams, use the `archify` bridge.
4. **English typography is not tuned.** Defaults target Chinese business documents; English-only manuscripts will
   look denser than a native English template.

---

## Provenance

The code was written by an AI coding assistant (DeepSeek v4.1 flash) under this project's specification, with the
maintainer defining requirements, providing context, reviewing and verifying. See [NOTICE.md](NOTICE.md) for the
full statement, including what the maintainer does and does not claim. MIT licensed.
