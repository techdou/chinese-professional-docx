# 审计判定机制速查（对照 scripts/docx_ops.py 实现）

调用 `audit` 前先读这份，不需要再翻源码。行号以 docx_ops.py 当前版本为准（`classify` ≈ L193、`body_ids` ≈ L163、`body_issues` ≈ L304、`MANUAL_LIST` ≈ L36）。

## 1. 段落角色是怎么判定的（classify，按顺序短路）

| 优先级 | 条件 | 判定角色 | 含义 |
| --- | --- | --- | --- |
| 1 | 祖先含 `w:tc` | table | 表格内段落，不做正文缩进检查 |
| 2 | 祖先含 `txbxContent` | textbox | 文本框；**不产生 manual-review 警告**，也不参与正文审计 |
| 3 | 祖先含 `ins/del/moveFrom/moveTo` | tracked | 修订中 |
| 4 | 祖先含 `sdt` | content-control | 内容控件 |
| 5 | 有效 `outlineLvl` < 9 | heading | 标题（含 Heading 样式的大多数情况） |
| 6 | 有效 `numPr/numId` ≠ 0 | list | 列表 |
| 7 | 样式链任一层的样式键以 `heading/toc/list` 开头，或命中 special_style_keys | special | 题注/页眉页脚/Title 等特殊样式，**豁免正文缩进检查** |
| 8 | 段内含 `w:drawing` 或 `w:pict` | mixed-drawing | 含图段落（即使样式是 Normal） |
| 9 | 含 `oMath/oMathPara` | equation | 公式 |
| 10 | 文本为空 | empty-body / empty | 在正文样式集合内为 empty-body |
| 11 | 文本以 `•●▪■`、`- `/`+ `、`1、`/`一、`、`1.`、`(一)` 等开头（MANUAL_LIST 正则） | ambiguous-manual-list | 疑似手工列表 |
| 12 | 段落样式在正文样式集合内 | body | **正文，执行缩进检查** |
| 13 | 其余 | unclassified | 未知角色，不猜 |

关键推论：**角色 12 的判定只看样式，不看内容**。封面大字、日期行只要用 Normal 就会被判 body 并要求缩进——它们必须换用第 7 行的 special 样式（见 §4）。

## 2. "正文样式集合"怎么算（body_ids）

- 基础 = style-profile.json 的 `body_style_keys`（Normal、正文、Body Text 1-3、紧凑正文、正式正文（首行缩进）等）∪ 文档默认样式。
- 扩展 = 上述样式的**命名派生**样式也属于正文。
- 排除 = 样式链上任一层是标题/TOC/列表/special 键，或有 outline/编号的（structural），即使 basedOn 正文样式也移出正文集合。

第三方文档常见数字样式 ID（`23/33/a1`）本身不是问题——问题是你不知道它基于谁。audit 的 `--details` 会列出每段 style_id/名称/角色，先跑一次再决定是否显式 `--body-style "样式名"`。

## 3. 正文缩进的双重要求（body_issues）

对每个 body 段落及其样式检查同一套规则：

| 检查 | 条件 | 级别 |
| --- | --- | --- |
| first-line-chars | 有效 `w:ind/@firstLineChars` 必须 = `200` | error（style 级与段落级各查一遍，**只改段落会漏掉 style 级报错**） |
| hanging-conflict | `hanging/hangingChars` ≠ 0 | error |
| fallback-indent | `w:ind/@firstLine` 必须 = `round(字号半点 × 20)` twips | warning（12pt → 480；11pt → 440，字号改了要同步改后备值） |

## 4. 高频修复路径（判定规则的应用）

| 症状 | 根因 | 修法 |
| --- | --- | --- |
| 封面/日期/大字段落报 first-line-chars | 用 Normal 写封面 → 判 body | 封面元素改用 `Title`/`Subtitle`/`文档信息` 等 special 键样式；自定义封面样式应 basedOn Title（而非 Normal）或直接命名进 special 键 |
| `style:23 style:33 …` 报错 | 第三方模板数字 ID 的 Body Text 系样式没设样式级缩进 | 在样式级（不是段落级）补 `firstLineChars=200`+匹配后备值；或 fix-indent |
| 图片段落 warning: manual-review (mixed-drawing) | 含图段落按设计不判正文 | 结构性豁免，不修；见 qa.md 警告分类表 |
| 参考文献段落 warning: manual-review (unclassified) | 独立样式**不是** special 键命名（如自建"参考文献"样式），判 unclassified | 结构性豁免（悬挂缩进正是预期形态）。注意：若样式名恰为 special 键（如"参考资料"）则判 special，连警告都没有——两种情况都不修，也不要改回 Normal |
| fallback-indent 反复出现 | 样式字号 ≠ 12pt 但后备值写死 480 | 按字号算后备值，或统一正文系样式为 12pt |

## 5. generic 附加检查（scope=generic 时，作用于默认样式）

Normal 样式应满足：`sz=24`（12pt）、`eastAsia=宋体`、`ascii/hAnsi=Times New Roman`、`line=360 & lineRule=auto`（1.5 倍）、`jc=both`（两端对齐）；不满足记 warning。Heading1/2/3 期望 31/27/24 半点（15.5/13.5/12pt），偏差记 heading-size warning。用户指定模板的页面/字号不同不算违规（见 style-guide 规则优先级）。

已知限制：§5 这些 Normal/标题检查的期望值目前**硬编码在 audit 实现里**，修改 style-profile.json 不会改变它们（profile 当前只驱动部分检查，如标题尺寸表）。要让检查可配置需改代码——改动时同步更新本表与 style-guide.md。
