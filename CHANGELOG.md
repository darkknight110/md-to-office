# 更新记录

## 1.0.0

首个可安装版本（`pipx install git+…`）。

- `md2office`：Markdown → 中文 Word/PDF；按 spec 生成真框图并替换稿内 ASCII 字符画
- `gongwen`：公文模板引擎（`template add/show` → `render` → `check`），默认参数对齐 GB/T 9704-2012
- CLI 契约：`--version`、`--doctor`（依赖/字体/LibreOffice 体检）、`--json`、`--no-overwrite`
- `gongwen --strict`：未知字段与残留占位符一律报错，不生成文件
- 字体检测：常见中文字体按系统字体文件判断（macOS 仿宋直报缺失，不再显示 `？`）
- 排版基线：52 份真实交付物实测 + 国标对照，脚本可复跑
- 测试 36 项；CI 覆盖 Python 3.9/3.11/3.12/3.13 + PDF 路径 + 每日定时

### 1.0.1

- Windows：CLI 入口统一把 stdout/stderr 切到 UTF-8；此前打印 `✓/✗` 会 `UnicodeEncodeError` 直接崩
- CI 增加 Windows 测试（无中文字体的用例自动跳过）
