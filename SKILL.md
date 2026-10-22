---
name: md-to-office
description: 把 Markdown 转换为排版规范的中文 Word（.docx）与／或 PDF，可在转换时自动生成分层架构图并替换稿内的 ASCII 字符画。当用户要求"导出 Word""导出 PDF""生成可发送的文档成品""转成 docx／pdf"，或交付物是 .docx／.pdf 文件时使用。可只出 Word、只出 PDF，或两者都出。执行转换前先加载本技能。
whenToUse: 需要把 Markdown 方案／报告／文档交付为 Word 或 PDF；需要为文档生成分层架构图；需要把稿内用字符画的架构图换成真正的图。
---

# Markdown → Word / PDF 交付

把 Markdown 转成**排版规范的中文 Word 与／或 PDF**，并可选地生成架构图、替换稿内的 ASCII 字符画。

**先问清输出格式。** 若用户未指明要 Word、PDF 还是两者都要，**用 `ask_user_question` 问一次**，不要默认全出。

## 快速开始

```bash
PY="<带 python-docx 与 Pillow 的 python>"   # DSH 环境用 load_workspace_dependencies 返回的
SK="<本技能目录>/scripts"

# 两者都出（默认）
"$PY" "$SK/md2office.py" --input ./文档 --output ./导出 --format both

# 只出 Word（最快，不走 PDF 转换）
"$PY" "$SK/md2office.py" --input ./文档 --output ./导出 --format docx

# 只出 PDF（内部先生成临时 DOCX，转完只保留 PDF）
"$PY" "$SK/md2office.py" --input ./文档 --output ./导出 --format pdf

# 只转某几个文件
"$PY" "$SK/md2office.py" --input ./文档 --output ./导出 --format pdf --only "07_*.md"

# 单个文件
"$PY" "$SK/md2office.py" --input ./方案.md --output ./导出 --format docx
```

> 已经用 `pipx install md-to-office` 装过？直接用 `md2office --input … --output … --format docx`，
> 并可先跑 `md2office --doctor` 体检环境（依赖 / 中文字体 / LibreOffice）。

**Python 需要在带 `python-docx` 与 `Pillow` 的环境里跑。** 在 DSH 中，用
`load_workspace_dependencies` 返回的 python 最省事（它已带这两个库）；在任何其他环境里，
`pip install -r requirements.txt` 后用自己的解释器即可。**只出 Word 时不需要任何外部二进制。**

## 参数

| 参数 | 说明 | 默认 |
|---|---|---|
| `--input` | `.md` 文件或目录（目录会**递归**找 `*.md`） | 必填 |
| `--output` | 输出目录，按输入的相对结构建子目录 | 必填 |
| `--format` | `docx` / `pdf` / `both` | `both` |
| `--only` | 文件名通配，可重复（如 `--only "01_*.md"`） | 全部 |
| `--spec` | 规格 JSON（架构图 + 字符画替换规则） | 自动找 `md2office.spec.json` |
| `--no-prep` | 跳过预处理 | 否 |
| `--keep-intermediate` | `pdf` 模式下保留中间 `.docx` | 否 |
| `--font-body` / `--font-heading` / `--font-code` | 字体 | 宋体 / 微软雅黑 / Consolas |
| `--size-body` / `--size-table` | 字号 pt | 10.5 / 9 |
| `--page` | `a4` / `letter` | `a4` |
| `--no-footer` | 不输出页码 | — |

## 架构图（可选）

Markdown 里用字符画（`┌─┐│└┘`）画的架构图，在 Word/PDF 里会因字体替换而**错位成一团**。用规格文件生成真正的图并自动替换。

写一个 `md2office.spec.json` 放在被转换目录下（或 `--spec` 指定）：

```json
{
  "diagrams": [
    {
      "out": "diagrams/arch.png",
      "top":   { "title": "顶层平台", "sub": "副标题" },
      "main":  { "title": "核心系统", "badge": "本方案",
                 "modules": [["模块1","模块2","模块3"], ["模块4","模块5","模块6"]] },
      "bottoms": [ { "title": "对接对象A", "sub": "说明" },
                   { "title": "对接对象B" } ]
    }
  ],
  "blocks": [
    { "contains": ["顶层平台", "核心系统"], "image": "diagrams/arch.png" }
  ],
  "defence_numbered_blocks": true
}
```

- `blocks[].contains` 里的关键词**全部命中**才替换；有规则但都没命中时**保留原样**，不误替换
- **替换只对含制表符画线（`┌─┐│└┘`）的代码块生效**，普通代码块不会被动
- 生成的图存在 `<输出目录>/<out>`，并在预处理阶段以绝对路径插入
- 也可单独跑图：`python diagram.py --spec spec.json --out-dir ./导出`
- `defence_numbered_blocks` 会把代码块里的编号流程（1. 2. 3.）去围栏、变成正常列表

### 复杂图型：可选接 archify

内置渲染器只画分层框图。**时序图 / ER 图 / 数据流 / 状态机**这类图，给该条目加
`"tool": "archify"` 并用 `archify` 字段放 archify 的 JSON-IR 即可（`"quality"` 可选，
默认 `showcase`）：

```json
{ "out": "diagrams/flow.png", "tool": "archify",
  "archify": { "schema_version": 2, "diagram_type": "workflow",
               "meta": { "title": "订单流程" },
               "lanes": [ { "id": "c", "label": "顾客" }, { "id": "p", "label": "平台" } ],
               "nodes": [ { "id": "o", "lane": "c", "col": 0, "type": "external", "label": "下单" },
                          { "id": "y", "lane": "p", "col": 1, "type": "backend",  "label": "支付" } ],
               "edges": [ { "from": "o", "to": "y", "label": "确认" } ] } }
```

**archify 是可选依赖**，未安装时会报错并给出安装指引，不影响内置渲染器。
安装：`npx skills add tt-a1i/archify -g`，或用 `ARCHIFY_BIN` 指定 CLI 路径。
桥接脚本：`scripts/archify_bridge.py`。

## 支持的 Markdown 语法

标题 1–4 级 · 表格（表头跨页重复、行不跨页断开）· 无序／有序列表 · 引用块 · 代码块 · 独占一行的图片 · `**粗体**` · `` `行内代码` `` · `[链接](url)` · `---` 分隔线。

**不支持**：嵌套列表的多级样式、`**粗体**` 与斜体叠加、单元格内转义竖线 `\|`、目录（TOC）自动生成、页眉、脚注、公式、HTML。表格列宽为等分；列多且文本长时会被挤，可接受但不完美。

## 注意事项

- **PDF 是从 DOCX 转出来的**，链路必然为 `Markdown → DOCX → PDF`；环境里没有 pandoc，LibreOffice 也不认 Markdown。`--format pdf` 会生成临时 DOCX，转完只保留 PDF。
- **PDF 需要任一个 LibreOffice 后端**，自动按顺序选择：① 系统 LibreOffice（`LIBREOFFICE_PATH`／`SOFFICE_PATH`，或 PATH 上的 `soffice`，或常见安装路径）；② DSH 随附的 libreoffice-kit（自动探测，或 `LIBREOFFICE_KIT_NODE` + `LIBREOFFICE_KIT_CLI`）。两个都没有时会报错并列出这两条路。
- **只出 Word 就不要碰 PDF 后端**：用 `--format docx`，零外部依赖。
- 生成 PDF 的机器上若缺 `宋体`／`微软雅黑`／`Consolas`，PDF 会以替换字体渲染（中文正常，字形略有差异）。要严格一致，需在生成 PDF 的机器上安装这些字体。
- 表格行宽等分；`--font-*` 只影响 Word 中的字体名，PDF 是否命中取决于渲染机装没装该字体。
- **本仓库不分发任何字体**，架构图渲染读取系统已装字体。
- 交付前**建议**（非必需）用 `office-docx` 技能的 `check_office.py` 校验 DOCX 结构；没装该技能不影响本技能运行：

```bash
"$PY" "<office-docx 技能的 scripts>/check_office.py" out.docx --out checks.json
```
