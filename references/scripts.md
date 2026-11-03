# 脚本接口

## 环境与路径

运行 `python scripts/doctor.py` 做只读预检：Python 3.10+、python-docx、lxml、母版存在性及可选渲染/字体探测。`--require-render` 要求本包 LibreOffice + PyMuPDF 后端；无此参数时，缺渲染器不阻止本地结构操作，仍需宿主完成视觉检查。预检不联网、不安装软件、不写文件，不等于客户端已加载。

下列路径以 Skill 根目录为基准。实际调用通过 SKILL.md 的路径定位脚本，并使用用户工作区的真实输入/输出绝对路径。已有 .venv 优先；不要假定 /work 存在或必须切换当前目录。

## docx_ops.py

```bash
python scripts/docx_ops.py validate-brief 内容.json
python scripts/docx_ops.py audit-headings 原稿.docx --scheme chinese-formal
python scripts/docx_ops.py format-headings 原稿.docx --scheme chinese-formal --out 多级标题版.docx
python scripts/docx_ops.py new --brief 内容.json --out 新报告.docx
python scripts/docx_ops.py new --scene handout --out 新讲义.dotx
python scripts/docx_ops.py audit 原稿.docx --scope indent --details --report 审计.json
python scripts/docx_ops.py fix-indent 原稿.docx --out 排版版.docx
python scripts/docx_ops.py audit 排版版.docx --scope indent --details --strict-review
python scripts/docx_ops.py format-code 旧稿.docx --out 代码底纹修复.docx
python scripts/docx_ops.py format-tables 旧稿.docx --all --style three-line --out 三线表修复.docx
python scripts/docx_ops.py convert 新报告.docx --out 报告母版.dotx
```

`new --brief` 还接受顶层 `document_type`（general/formal-report/teaching/project-report/research-report/technical-report/thesis-natural/thesis-chapter/thesis-humanities）、`numbering`（auto/none/chinese-formal/academic-decimal/academic-chapter/humanities），heading block 可设 `"numbered": false`，可插入 `{"type":"toc"}` 生成可更新的 Word 目录域。标题文字在启用自动编号时不得含手动序号。详见 [heading-numbering.md](heading-numbering.md)。

`new` 的 scene 为 generic/competition/handout/report/components；无 brief 时复制示例，保留占位内容。brief 与非 generic scene 互斥。validate-brief 只验证内容接口，不创建文档。JSON Schema 及例子在 assets 中；行宽等语义由脚本最终验证。

`audit` 默认 scope=generic，检查正文缩进及部分通用字体/页面参数。scope=indent 跳过通用字号、字体、行距和页边距比较，适合只修缩进或不覆盖指定模板。`--details` 增加段落序号、样式 ID/名称、角色和有效缩进，不复制正文文本。`--strict-review` 在没有错误但仍有警告时退出 3。`--summary` 只向 stdout 打印一行式结论 JSON（file/status/errors 数/warnings 数/warning_codes 分布/checks/review_required），全量报告仍可用 `--report` 落盘——给 Agent 消费或 CI 判定用，人读全量时不用。

`pipeline` = audit + 可选渲染一条龙：`docx_ops.py pipeline 稿.docx --render-dir 空目录 --dpi 130 --backend libreoffice [--summary]`。audit 语义与退出码同上；渲染经 subprocess 调 render_preview.py（`--render-dir` 必须为空或不存在），渲染失败时 JSON 的 render 字段带 error 且退出 2。退出码优先级：audit 有错误 1 > 渲染失败 2 > strict-review 3 > 0。

判定规则（样式→角色、缩进双重要求、高频修复路径）见 [audit-rules.md](audit-rules.md)；警告哪些可修、哪些按设计存在见 [qa.md](qa.md) 警告分类表。

审计 JSON 保留 status、errors、warnings、paragraph_roles 等原字段，schema_version=2.0 增加 scope、review_required、分项 checks；details 时增加 paragraphs。调用方应读取 errors/warnings/coverage，不应只比对一个 pass 字符串。

`fix-indent` 只修已识别普通正文，不做全文重写。Normal/Body Text 系列与确认的子样式属于正文；未知 Normal 子样式不能直接猜测。确认后可显式加 `--body-style "样式名或ID"`；只在确认行首填充时加 `--strip-manual-padding`。清理到超链接、域或内容控件即停止，不跨过去删除后面的空白。

`audit-headings` 只读核查 Heading 1–3 的自动编号范式、层级跳跃与手动前缀；`format-headings` 只对没有其他编号的三级标题使用原生多级编号，默认跳过已有手工编号/不同编号体系/特殊标题。`--convert-manual` 是明确授权时才可使用的保守转换，复杂标题自动跳过，原件不覆盖。

`format-code` 只处理段落样式已明确标为“代码段”/"Code Block"/"CodeBlock" 的代码（附加已有自定义样式用 `--style`）。修复段落底纹并去掉彩色左边框，不基于文本猜代码，不改普通正文。`format-tables` 默认只处理已标为 `AcademicThreeLine` 的表格；对其他旧表必须显式 `--all`，且可用 `--style three-line|grid|none`，不会改动单元格文本、合并关系和图片。两者都保留原文件并输出新副本，保护/修订/签名内容拒绝改动。

`convert` 修改真实 DOCX/DOTX 内容类型，不是改扩展名。所有输出，包括审计报告，要求新路径，不静默覆盖原件或旧结果。只支持未加密常规 DOCX/DOTX；拒绝宏、恶意 ZIP 路径、DTD/实体和循环样式。检测到保护、签名或修订的内容，不自动修复/转换。

退出码：0=操作完成，普通审计仍可能带警告；1=审计有缩进/结构错误；2=输入、环境或安全问题；3=严格审计仍需复核。原来的未加 strict-review 的调用方式仍兼容。

## 预览

```bash
python scripts/render_preview.py 成稿.docx --out-dir qa-新目录 --dpi 130
python scripts/render_preview.py 成稿.docx --out-dir qa-noto --fallback-cjk
python scripts/render_preview.py 成稿.docx --out-dir qa-word --backend word
```

输出目录须为空或不存在，产生逐页 PNG、内部 preview.pdf 和 render.json。渲染使用临时副本与隔离 LibreOffice 配置；程序只记录“等待逐页查看”，不宣称视觉通过。不是安全沙箱，不在正常桌面环境运行不受信任的 Office 文档。

`--backend word`（Windows）经 COM 用宿主桌面 Word 导出 PDF 再转 PNG：保真度最高，能正确呈现嵌入字体；原文档以只读方式打开、不修改。该后端需要 pywin32 与桌面 Word，与 `--fallback-cjk` 互斥；默认后端仍是 libreoffice。注意 `--timeout` 仅对 libreoffice 后端生效，COM 调用本身无超时保护（卡死需手工结束 WINWORD 进程）。

缺宋体/黑体时，fallback-cjk 只在预览副本映射到已安装的 Noto Serif/Sans CJK SC，正式文件不改。原字体在最终设备的 Word/WPS 尚需核验。

## 字体嵌入

```bash
python scripts/embed_fonts.py 交付.docx            # 就地嵌入保存
python scripts/embed_fonts.py 交付.docx --out 交付-嵌字体.docx
```

Word COM 后端（Windows）：`EmbedTrueTypeFonts=True` 且不排除系统字体，保存后自验包内 `word/fonts/` 非空并在 JSON 里列出嵌入字体清单。文件显著增大属预期。嵌入让缺字体机器正常**显示**，不转移字体许可，不分发字体文件本身。非 Windows 宿主在 Word 内手动嵌入。

## doctor 探测项

除原有的 Python 版本、python-docx/lxml/PyMuPDF、LibreOffice、fontconfig 字体探测外，还报告：`word_com`（pywin32 + 桌面 Word 激活探测，供 `--backend word` 与字体嵌入用）和 `assets_fonts`（skill 自带字体资产清单）。均为只读探测。


## v1.6 可更新引用与续表

```bash
python scripts/docx_ops.py new --brief assets/examples/brief-references.json --out paper.docx --manifest-out refs.json
python scripts/docx_ops.py audit-references paper.docx --strict-review
python scripts/docx_ops.py audit-long-tables paper.docx --strict-review
python scripts/docx_ops.py link-references existing.docx --manifest refs.json --out existing-linked.docx
python scripts/qa_cross_compat.py paper.docx --roundtrip-libreoffice --report qa-interop.json
```

`link-references` 仅能绑定已存在的 target 书签；不可把普通文字“图1-1”猜测为目标。插图目录、附表目录和续表规则见 [references-and-long-tables.md](references-and-long-tables.md)。


## v1.7：引用标准、结构化书目与著者—出版年

```bash
python scripts/docx_ops.py audit-bibliography assets/examples/brief-gbt7714-2025.json --strict-review
python scripts/docx_ops.py new --brief assets/examples/brief-gbt7714-2025.json --out paper.docx --manifest-out refs.json
python scripts/docx_ops.py audit-references paper.docx --strict-review
python scripts/docx_ops.py link-references edited.docx --manifest refs.json --out edited-linked.docx
```

- `reference_standard`：`gbt7714-2025`（默认）、`gbt7714-2015-compat`、`institution-custom`。后两者仅支持 `text` 形式兼容旧标准或机构版式。
- `citation_mode`：`numeric-superscript`（默认）、`numeric-inline`、`author-year`，`link-references` 默认 `auto` 从 manifest 读取，防止数字/作者年份混用。
- 新增 `[[CITE:key@12-16]]` 单文献引文页码；连续引用 `[[CITE:k1;k2;k3]]` 按初始顺序组成范围，需要在重排正文后重新生成分组。
- `audit-bibliography` 输出条目级 `errors`/`warnings`，无错误不等于实际出处已核实。未支持的结构化类型请使用已审核 `text`。
- 新文档 `core_properties.keywords` 保存 citation_mode 与 reference_standard 以供 Word 后续审计识别。
- 详细规则见 [gbt7714-2025.md](gbt7714-2025.md)。
