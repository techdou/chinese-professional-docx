# chinese-professional-docx

中文/中英混排 Word（DOCX/DOTX）制作、排版、审计与修复的 Agent Skill。默认中国正式文档通行的黑白学术排版：正文首行缩进 2 字符、宋体 + Times New Roman 混排、Word 原生三级标题编号、浅灰代码底纹、三线表，并提供 GB/T 7714—2025 参考文献著录与图表公式交叉引用。

[English](README.en.md)

## 功能特性

- **正文排版**：首行缩进 2 字符（`w:firstLineChars="200"`），宋体 12 pt / Times New Roman、1.5 倍行距、A4 · 2.5 cm 页边距。
- **标题编号**：四套中文范式（`一、／（一）／1.`、`1／1.1／1.1.1`、`第1章／1.1／1.1.1`、`第一章／一、／（一）`），全部为 Word 原生多级列表，支持目录自动更新。
- **参考文献**：GB/T 7714—2025 结构化著录（J/M/C/D/R/S/EB/DS/PP 九类核心），`[[CITE:id]]` 上标方括号、连续区间 `[1-3]`、引文页码、著者—出版年制可选。
- **图表公式**：按章编号（图 1-1 / 表 1.1 / 式（1-1）），真实 `SEQ`/`REF`/`PAGEREF` 域 + 书签，插图/附表目录（Word `TOC` 字段）。
- **长表与续表**：自然分页重复表头（`repeat-header`）或显式分段续表（`explicit-chunks`），续表标题用原表 `REF` 联动。
- **审计与修复**：`audit` / `fix-indent` / `format-code` / `format-tables` / `format-headings` / `audit-bibliography` / `audit-references` 等保守修复命令，全部新路径输出、不覆盖原件。
- **质量保障**：逐页渲染预览（LibreOffice / PyMuPDF / Word COM 三后端）、审计警告分类、[互操作验收矩阵](docs/compatibility-matrix.md)、35 项回归测试。

## 环境要求

- Python 3.10+
- 核心：`python-docx`、`lxml`
- 渲染（可选）：LibreOffice 或 PyMuPDF，或宿主 Word（`--backend word` 需 Windows + pywin32）
- 测试（可选）：`jsonschema`、`PyYAML`

## 快速开始

```bash
# 安装依赖（核心 + 渲染 + 测试）
python -m pip install -r requirements-dev.txt

# 环境探测
python scripts/doctor.py

# 新建文档（brief → 生成 → 审计渲染一条龙）
python scripts/docx_ops.py validate-brief assets/examples/brief.json
python scripts/docx_ops.py new --brief assets/examples/brief.json --out 正式文档.docx
python scripts/docx_ops.py pipeline 正式文档.docx --render-dir qa --summary

# 旧稿排版修复（先审计，再修，再复查）
python scripts/docx_ops.py audit 旧稿.docx --details --report audit.json
python scripts/docx_ops.py fix-indent 旧稿.docx --out 修复稿.docx

# 标题编号诊断与保守迁移
python scripts/docx_ops.py audit-headings 旧稿.docx --scheme chinese-formal
python scripts/docx_ops.py format-headings 旧稿.docx --scheme chinese-formal --out 编号稿.docx

# 代码底纹 / 三线表显式修复
python scripts/docx_ops.py format-code 旧稿.docx --out 代码底纹稿.docx
python scripts/docx_ops.py format-tables 旧稿.docx --all --style three-line --out 表格稿.docx

# GB/T 7714—2025 文献审计
python scripts/docx_ops.py audit-bibliography assets/examples/brief-gbt7714-2025.json
```

完整 CLI 与退出码见 [references/scripts.md](references/scripts.md)；审计判定与警告分类见 [references/qa.md](references/qa.md)。

## 目录结构

```
chinese-professional-docx/
├── SKILL.md                 # Skill 主入口（路线决策表 + 不变约束 + 执行入口）
├── scripts/                 # docx_ops.py 主 CLI 与辅助脚本
├── references/              # 按需加载的参考文档（编号范式、7714 契约、审计规则等）
├── assets/
│   ├── templates/           # base.docx / base.dotx 通用母版
│   ├── examples/            # 场景示例与 brief 样例
│   ├── schemas/             # brief JSON Schema
│   └── style-profile.json   # 样式参数
├── evals/                   # 触发用例与评测集
├── tests/                   # 单元/回归测试（unittest）
├── tools/                   # merge_full_original.py 完整版合并工具
├── docs/                    # 兼容性矩阵（开发计划类文档不入库）
├── CHANGELOG.md             # 版本历史
└── README.en.md             # 英文说明
```

## 测试

```bash
python -m unittest discover -s tests -v    # 35 项回归
python scripts/docx_ops.py audit assets/templates/base.docx --summary   # 母版自检应 0 错误
```

## 字体说明

本仓库**不包含任何字体二进制**（Windows 附带字体受许可限制，禁止再分发）。本地使用时自行准备宋体、黑体、楷体、仿宋、Times New Roman 等系统字体：预览渲染从系统字体解析，交付前可用 `scripts/embed_fonts.py`（Word COM 后端）将字体嵌入文档。`tools/merge_full_original.py` 可将本源码与本地完整版 ZIP 合并（只写新文件、保留原包内容并校验字体哈希）。

## 版本

当前版本 **v1.7.0**（template-version 1.3），变更历史见 [CHANGELOG.md](CHANGELOG.md)。`metadata.version` 随 API/行为变更升级；`template-version` 仅随母版二进制变更。

## 边界声明

- 默认排版是通用范式，**不构成任何院校的官方学位论文合规**；指定学校/赛事模板优先，正式提交稿必须逐项核对官方模板。
- 自动审计不能核实原始文献事实；7714 著录的元数据必须经真实来源核验。
- LibreOffice 渲染通过不等同于 Word/WPS 原生验证；正式验收以真实宿主逐页检查为准。
- 未附开源许可证；字体与文档交付边界见 [references/sources.md](references/sources.md)。
