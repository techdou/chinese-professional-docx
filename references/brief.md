# JSON brief 契约

用于从用户内容新建通用 Word，不用于复刻学校/赛事固定模板或编辑复杂原稿。先执行 `docx_ops.py validate-brief BRIEF.json`，验证不写 Word 文件；通过后才 `new --brief BRIEF.json --out NEW.docx`。

顶层：必填非空 `title`；可选字符串 `subtitle`、`info`，以及最多 5000 项的 `blocks`。不接受未知键。

块类型：`paragraph`、`heading`（level=1–3）、`bullets`/`numbered`（非空字符串 items）、`note`、`code`、`reference`、`table`（headers、rows，可选 caption）、`pagebreak`。文本块的 `text` 必须是字符串；表格每行的列数必须等于表头列数。表格 `style` 可选：`three-line`（默认）、`grid`（特殊复杂表）、`none`（无框线）；不提供彩色表头/斑马纹。

普通正文 text 的换行会形成真实段落，不允许行首手工空格或 Tab；代码保留缩进。编号列表每组重新从 1 开始。表格题注置于表上方、使用 SEQ 域，首行重复表头；长表跨页“续表”需在终稿人工核验。`pagebreak` 仅为明确需要的分页使用。

机器 Schema 位于 `assets/schemas/brief.schema.json`，内容例子位于 `assets/examples/brief.json`。JSON Schema 覆盖结构及类型；表格等宽约束与字符串语义由 validate-brief 最终核验。不臆造未支持的块；figure/equation 已在 v1.6 支持（image_path/placeholder、纯文本 expression），复杂 MathML、OMML、脚注对象交给宿主 DOCX 工具。

模板复制与文档生成要区分：`new --scene ...` 复制示例，含占位内容；`new --brief ...` 使用真实已确认内容。缺失成果/预算应标明待补，不能沿用示例中的状态假装完成。


## v1.6 引文与学术对象契约

`bibliography` 由唯一 ASCII `id` 和已核实的原样 `text` 组成（不要手工加 `[n]`）；`citation_mode` 默认 `numeric-superscript`，可选 `numeric-inline`；`caption_numbering` 为 `auto/global/chapter`，`caption_separator` 为 `-` 或 `.`。

正文 `[[CITE:refA;refB]]`、`[[FIG:figureA]]`、`[[TABLE:tableA]]`、`[[EQ:loss]]`、`[[PAGE:figureA]]` 创建可更新 REF/PAGEREF 字段；不支持用题名猜测引用目标。

`figure` 必须有 `id`、`caption` 及 `image_path`/`placeholder` 二选一；`equation` 必须有 `id`、纯文本 `expression`；`table` 可有 `id`、`caption`、`note` 和 `continuation.mode`。`explicit-chunks` 时必须声明 `id/caption/continuation.rows_per_part`。

`list-of-figures` / `list-of-tables` 是可更新图表目录块。正式参考实例：[assets/examples/brief-references.json](../assets/examples/brief-references.json)；完整说明：[references-and-long-tables.md](references-and-long-tables.md)。


## v1.7：现行 GB/T 7714—2025 的文献数据引擎

`reference_standard` 默认 `gbt7714-2025`；学校/期刊明确要求可以设置 `institution-custom`，旧稿可设 `gbt7714-2015-compat`（这两种情况下仅接受已人工核验的 `text` 条目，不套 2025 格式）。

`bibliography` 每项可为原样 `{"id":"...","text":"..."}`，也可为包含 `type/title/authors/year/doi/url/...` 的结构化记录。结构化 2025 核心支持 J/M/C/D/R/S/EB/DS/PP 九类；对于其他文献类型用已校验 `text`，不得直接标记为已自动著录。

`citation_mode`：`numeric-superscript` / `numeric-inline` / `author-year`。`[[CITE:id;id]]` 支持同处多文献，连续编号在初次生成时组成 `1-3` 区间；`[[CITE:id@23-25]]` 支持单文献引文页码。作者—年份制要求结构化作者与年份、同年多文献 `a/b/c`，采用隐藏标签书签+可更新 `REF`，与数字制分离。

命令：`python scripts/docx_ops.py audit-bibliography brief.json --strict-review`；完整规则见 [gbt7714-2025.md](gbt7714-2025.md)。中文正式文档采用此基准，并非学校/期刊的强制唯一形式。
