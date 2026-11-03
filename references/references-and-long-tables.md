# 引文、图表、公式与长表：语义优先、字段联动

## 规范依据及适用边界

- 默认采用中文正式文档广泛使用的黑白学术排版：**顺序编码引用、正文上标、图表/公式自动编号、三线表**。
- 部分高校研究生学位论文规范的实际要求：顺序编码制文内方括号上标；著者—出版年制可选；图、表、公式按章编号，格式如“图1.1 / 图1-1”；需跨页续排的表格另页“续表2.1”并重复表头。
- 现行国家标准是 [GB/T 7714—2025](https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=C6CE52E55AC09B9C79A20AEA77CEDD14)，已于 2026-07-01 实施。此 Skill 支持 **9 种核心类型的结构化 2025 格式化**，同时接受已审校的 `bibliography[].text` 兼容旧稿；无法由工具自动确认的著录事实须人工核对，详见 [GB/T 7714—2025 数据契约](gbt7714-2025.md)。
- 用户或学校模板若要求著者—出版年、非上标引用、其他章序号/分隔符，应采用其明确规范。当前引擎支持 `numeric-superscript`、`numeric-inline` 和独立的 `author-year`；著者—出版年制要求真实的 authors/year 元数据，书目可编辑，最终按学校规范复核。

## 推荐 JSON brief 契约

```json
{
  "title": "技术论文示例",
  "document_type": "thesis-chapter",
  "citation_mode": "numeric-superscript",
  "caption_numbering": "chapter",
  "caption_separator": "-",
  "bibliography": [
    {"id":"refA","text":"示例作者. 示例文章[J]. 示例期刊, 2025."}
  ],
  "blocks": [
    {"type":"toc"},
    {"type":"list-of-figures"},
    {"type":"list-of-tables"},
    {"type":"heading","level":1,"text":"绪论"},
    {"type":"paragraph","text":"已有研究[[CITE:refA]]，结构见[[FIG:structure]]，结果见[[TABLE:ablation]]。"},
    {"type":"figure","id":"structure","caption":"模型结构图","placeholder":"替换为真实图片"},
    {"type":"table","id":"ablation","caption":"消融实验结果","headers":["模型","指标"],
     "rows":[["A","0.80"],["B","0.81"],["C","0.83"]],
     "continuation":{"mode":"explicit-chunks","rows_per_part":2,"page_break_before":true},
     "note":"注：仅演示排版。"},
    {"type":"equation","id":"loss","expression":"L = a + b"},
    {"type":"paragraph","text":"式[[EQ:loss]]见上方，参见[[PAGE:structure]]页。"}
  ]
}
```

`[[CITE:a;b;c]]` 对连续编号生成形如 `[1-3]` 的引用区间，区间端点保留独立 REF 域；`[[CITE:a@23-25]]` 可指定引用页码（方括号外）。**Word 只能刷新 REF 数字，不会重新判断区间是否仍连续或自动排列书目**：章节插入、引用顺序变化后应重新由 brief 生成，并审核引用区间。先后顺序以正文首次引用为准。

`[[FIG:id]]` → `图1-1`；`[[TABLE:id]]` → `表1-1`；`[[EQ:id]]` → `式（1-1）`；`[[PAGE:id]]` → 页码 REF。上述均依赖 DOCX 内实际 bookmark + `REF` / `PAGEREF` 字段，**不是普通文本数字替换**。

## 长表两种模式

1. `repeat-header`（默认）：一张物理 Word 表；第一行设置 `w:tblHeader` 自动重复，行设置 `w:cantSplit`，表格结构不拆分。**Word 没有自动在每个自然分页处插入“续表”段落的原生字段**，这种模式不声称满足学校对逐页续表标题的全部要求。
2. `explicit-chunks`：按 `rows_per_part` 分成多个可编辑的 Word 表，每部分重复表头；第二部分起在指定新页输出 `续表` + `REF` 原表编号，不额外创建 `SEQ CPTable`。只计入一个表号，目录不应重复计数；脚本会把“注”放在最后一部分下面。

**分页并非按页高智能识别。** `rows_per_part` 的值必须根据实际列宽、字体、正文页边距和渲染结果确定；文稿应逐页检查是否存在大片空白、续表标题孤立或表格边线断裂。带合并单元格、嵌套表/表内公式的既有长表应在 Word/WPS 中保留原结构并人工检查，不允许自动拆碎。

## Word 原生字段与编号机制

- 每个编号一级标题插入隐藏的 `SEQ CPChapter`，同时重置 `SEQ CPFigure`、`SEQ CPTable`、`SEQ CPEquation` 计数器；该方案经 LibreOffice 重新计算后能正确生成第二章的 `2-1`，避免 `SEQ ... \\s 1` 在部分渲染器下失效。
- 原始图、表和公式题注创建 `SEQ` 字段，并用仅包裹编号的书签；正文引用插入 `REF <bookmark> \\h`，页码引用使用 `PAGEREF`。
- 续表使用原表序号 `REF`，不使用新的 `SEQ`，因此书签稳定、图表目录不会因续表额外编号。
- 图表目录使用 `TOC \\c "CPFigure"` / `TOC \\c "CPTable"`，属于可更新的目录域；完成排版后在宿主更新。
- 摘要、目录、参考文献、致谢等特殊章节不混入正文多级编号。

## 常用 CLI

```bash
python scripts/docx_ops.py audit-bibliography assets/examples/brief-gbt7714-2025.json
python scripts/docx_ops.py validate-brief assets/examples/brief-references.json
python scripts/docx_ops.py new --brief assets/examples/brief-references.json \
  --out academic.docx --manifest-out refs.json
python scripts/docx_ops.py audit-references academic.docx --strict-review
python scripts/docx_ops.py audit-long-tables academic.docx --strict-review
python scripts/qa_cross_compat.py academic.docx --roundtrip-libreoffice \
  --report qa/interop.json

# 修改旧稿：需要显式标记以及已有书签清单，不猜“表2-1”是什么
python scripts/docx_ops.py link-references draft.docx --manifest refs.json \
  --out draft-linked.docx
```

## 刷新 / 最终验收

在 Word 或 WPS 文字中打开文档，`Ctrl + A` → `F9` 更新域；必要时单独选中表格执行 `F9`，目录选择“更新整个目录”。检查引用序号、分隔符、上标、公式排版、页码、表题/图题位置，检查断开的书签（例如“错误！未找到引用源”）。

采用低侵入改稿：旧稿默认只审计；`link-references` 仅替换显式 `[[...]]` 标记，必须指定 manifest，且始终生成新文件。复杂域、跨 run 标记、标记缺失目标则停止，不静默猜测或重建原始稿。交付前必须在所使用的**实际** Word/WPS 版本上验证；LibreOffice 通过不等于 Word/WPS 原生通过。
