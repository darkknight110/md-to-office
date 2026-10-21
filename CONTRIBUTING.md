# 参与贡献

Markdown → 中文 Word/PDF，含真框图与公文模板。MIT 许可，欢迎 issue 和 PR。

## 本地跑起来

```bash
git clone https://github.com/darkknight110/md-to-office.git
cd md-to-office
pipx install git+https://github.com/darkknight110/md-to-office.git
```

## 提交前必须跑

```bash
python -m unittest discover -s tests -v
```

CI 会跑同样的命令；本地绿了再提 PR。

## 约定

- **先有证据，再下结论**：新增检查项要能给出可复核的证据，不靠"感觉"。
- **失败要出声**：不确定就报错或警告，绝不静默放行。
- **诚实边界**：做不到的能力写进文档，不写成"已支持"。
- 改排版默认值前先看 `docs/typography-baseline.json`——默认值不是偏好，是 52 份交付物的实测。

## 提交信息

`<type>(<scope>): <摘要>`，不超过 72 字；一个提交只做一件事。
