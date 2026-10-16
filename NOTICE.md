# 来源与声明 / Provenance and Notices

本文件补充说明本项目代码的来源、作者身份，以及第三方依赖与字体情况。
**许可证正文见 [LICENSE](LICENSE)（标准 MIT，未修改）。**

---

## 一、来源与作者身份 / Provenance and authorship

**本项目的代码由 AI 编码助手（DeepSeek v4.1 flash）在维护者提供的本项目具体场景下生成，
由维护者确定需求、提供上下文、审阅、修改并验证。**

它不是从任何既有开源项目复制的：维护者已确认代码来源为上述 AI 辅助生成；
审计亦未发现与既有同类项目（如各类 `md2docx`、Markdown→Word 转换器、
ASCII 图生成器）存在代码雷同。

旁证：代码高度贴合本项目场景，而非通用库形态——默认字体直接取中文商务文档的
`宋体`/`微软雅黑`/`Consolas`（`scripts/md2docx.py:38-40`），字符画识别用一组特化的
制表符集合 `BOX_CHARS`（`scripts/md2office.py:43`），还有为答辩场景设计的
`defence_numbered_blocks`（`scripts/md2office.py:177`）。全部 import 仅标准库 +
`python-docx` + `Pillow`，没有搬运任何模块结构。

### 这对许可意味着什么，请如实理解

1. 维护者对本项目的主张，限于其**可主张的部分**——需求定义、架构选择、规格文件格式设计、
   排版参数取舍、测试用例，以及全部人工审阅与修改。
2. 在部分司法辖区（如美国版权局现行立场），**纯机器生成的内容可能不构成可受版权保护的
   作品**；中国已有判例承认在人类有实质性智力投入时 AI 生成内容可受保护，但标准仍在演进。
   因此 MIT 许可证的授予，其效力**不超过许可方实际拥有的权利范围**。
3. 无论上述著作权状况如何，**MIT 授予你的使用、修改、分发权利依然完整有效**——
   授权不因上游是否可受版权保护而失效；你无需为使用本项目承担额外义务。
4. 使用本项目请自行评估合规性。本文件不构成法律意见。

---

## 二、第三方依赖声明 / Third-party notices

本项目**不打包任何第三方源码或二进制**，仅在运行时依赖以下库：

| 依赖 | 许可证 | 用途 |
| --- | --- | --- |
| [python-docx](https://github.com/python-openxml/python-docx) | MIT | 生成 .docx |
| [Pillow](https://python-pillow.org/) | MIT-CMU (HPND) | 渲染架构图 PNG |
| LibreOffice（可选） | MPL-2.0 | 仅通过命令行进程调用 `soffice` 转 PDF，不链接、不复制其代码 |

`scripts/archify_bridge.py` 通过子进程调用可选的外部工具
[archify](https://github.com/tt-a1i/archify)（MIT），不包含也不修改其代码。

---

## 三、字体声明 / Fonts

**本仓库不包含、不分发任何字体文件。**

架构图渲染时按绝对路径调用**操作系统已安装**的字体：

| 平台 | 字体 |
| --- | --- |
| macOS | STHeiti / Hiragino Sans GB |
| Linux | Noto Sans CJK / 文泉驿 |
| Windows | 微软雅黑 / 黑体 |

这些字体的授权由使用者与其系统提供方决定，与本项目无关。若需在容器等无字体环境中运行，
请自行安装合规字体，或通过 `--font-*` 参数指定字体文件。
