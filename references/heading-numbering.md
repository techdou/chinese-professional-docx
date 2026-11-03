# 标题层级与原生多级编号

## 适用目标

中文正式文档一般使用清晰稳定的三级标题体系。编号是一种 **可调整的默认排版范式**，不是未经用户确认的学校规定。首次创建时根据 `document_type` 选策略；已有稿、学校模板只读审查优先，不自动替换原始编号。

| Profile | 一级（Heading 1） | 二级（Heading 2） | 三级（Heading 3） | 常见场景 |
|---|---|---|---|---|
| `chinese-formal` | 一、 | （一） | 1. | 工作报告、课程讲义、竞赛材料、一般中文正式文件 |
| `academic-decimal` | 1 | 1.1 | 1.1.1 | 自然科学论文、科研与技术报告 |
| `academic-chapter` | 第1章 | 1.1 | 1.1.1 | 章节式技术论文、专著报告 |
| `humanities` | 第一章 | 一、 | （一） | 人文社科类学位论文 |
| `none` | 无编号 | 无编号 | 无编号 | 用户指定不编号，或沿用既有正式模板 |

公开发布的高校《研究生学位论文写作格式规范》“正文—章、节标题”部分**建议**自然科学采用 1/1.2/1.2.1、人文社科采用 第一章/一/（一）。这里使用其中的编号范式，不将学校关于章标题居中、换页、字号、页码等单独要求扩展为全部文档的统一强制标准。

## 执行机制

1. **优先级**：学校/机构正式模板的已有要求 → 用户此次明确排版要求 → 文档用途对应的推荐方案 → `general` 默认 `chinese-formal`。
2. **语义优先**：一级、二级、三级标题使用 Word 的 `Heading 1`/`Heading 2`/`Heading 3`（含 `outlineLvl`）；**不要**只通过字体大小猜测标题，更不要在正文标题前硬编码序号。
3. **原生多级列表**：`word/numbering.xml` 添加独立的 `abstractNum`/`num`，`w:lvl` 与 Heading1–3 关联，段落保留 `numPr`/`ilvl`/`numId`，二、三级在前级变化时重启；支持 Word 原生修改、重新编号及目录域。
4. **正文不受影响**：只对标题段落使用编号，图表题注、正文缩进、代码/表格正文不共用标题样式。
5. **特殊标题不计入正文序号**：摘要、目录、Abstract、参考文献、致谢、后记、附录等按文档需要设为 `numbered:false`，或使用明确的不编号标题样式。**不要**把所有 Heading 1 自动加“第一章”。
6. **自动目录**：新建 brief 可插入 `{"type":"toc"}` 生成 Word 原生 TOC 域；初次打开 Word/WPS 后更新域即可获得页码。布局引擎可能不会替自动目录刷新最终页码，必须复核。用户未要求目录则不强行插入。

## 新建 brief 示例

```json
{
  "title": "中文专业文档",
  "document_type": "research-report",
  "numbering": "auto",
  "blocks": [
    {"type": "heading", "level": 1, "text": "研究方法"},
    {"type": "heading", "level": 2, "text": "特征编码"},
    {"type": "heading", "level": 3, "text": "多模态特征融合"},
    {"type": "heading", "level": 1, "text": "参考文献", "numbered": false}
  ]
}
```

`document_type` 主要选择“默认方案”，不是强制章节内容。`numbering` 可以直接指定 `chinese-formal`、`academic-decimal`、`academic-chapter`、`humanities` 或 `none`，覆盖 `auto`。所有新建自动编号标题的 `text` 应仅包含标题文字；已有 `一、`、`1.1` 等手动序号时请删掉序号，或在该标题 `numbered:false`。

## 对现有 DOCX 保守修复

```bash
python scripts/docx_ops.py audit-headings 已有稿.docx --scheme chinese-formal
python scripts/docx_ops.py format-headings 已有稿.docx --scheme chinese-formal --out 标题调整版.docx
# 确认手工编号是可删除的前缀后，才显式开启安全转换
python scripts/docx_ops.py format-headings 已有稿.docx --scheme chinese-formal --convert-manual --out 转换编号版.docx
```

默认不处理原有不同编号体系或带手写前缀的标题，不批量替换普通段落；`--convert-manual` 仅作用于结构简单的 Heading 1–3 段落，只删除可识别的前缀，保留正文各 run 的加粗/斜体，不碰超链接、域、批注与图像等复杂混排标题。无法安全转换会跳过并报告，绝不静默清空标题。输出必须是新路径；二次执行应保持内容不变。

> 注意：文档中 Heading 1–3 的自定义属性、内嵌目录域或来自学校的特殊分节可能需要人工验证；脚本仅声明实现了 OOXML 原生编号及基本层级检查，不宣称已通过所有版本 Word/WPS 兼容性或学校学位论文正式审查。
