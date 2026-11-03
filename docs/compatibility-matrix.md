# Word / WPS / LibreOffice 互操作验证矩阵（v1.6）

> 目的：验证真实 DOCX 结构、更新域以及视觉排版。**不能将 LibreOffice 通过等同于微软 Word/WPS 原生验证。**

| 验收项 | OOXML / Python | LibreOffice（本机） | Microsoft Word 桌面版 | WPS 文字桌面版 |
| --- | --- | --- | --- | --- |
| 输入 DOCX、字段与 ZIP 完整性 | 自动 | 自动 | 人工检查 | 人工检查 |
| 两章 `SEQ` 图/表/公式重置为 `1-1`、`2-1` | 自动缓存/域结构 | 更新后渲染验证 | **待宿主验证** | **待宿主验证** |
| 上标 `[1,2]` / 反复引用 | 结构和样式自动 | 页面检查 | 待宿主验证 | 待宿主验证 |
| `REF`、`PAGEREF`、书签跳转 | 完整性审计 | 字段往返/显示 | 待宿主验证 | 待宿主验证 |
| 新插图/删除目标后的编号刷新 | 可制造测试夹具 | 可进行往返 | 待宿主操作 | 待宿主操作 |
| `TOC` / 插图目录 / 附表目录更新 | 域结构自动 | 可能存在刷新差异 | 待宿主操作 | 待宿主操作 |
| 长表重复表头、`续表`、最后一段注释 | OOXML 自动 | 分页渲染 | 待宿主验证 | 待宿主验证 |
| 文档字体/行距/封面 | 原始资产与 XML 审计 | 可能替代字体 | **待本地字体环境验证** | **待本地字体环境验证** |
| 修改保存再次打开，无域/图片丢失 | 原件保真测试 | 往返保存结构 | 待宿主验证 | 待宿主验证 |

## 原生桌面宿主验收步骤（需要真实安装的软件）

1. 使用 `assets/examples/brief-references.json` 生成 DOCX；用 Microsoft Word for Windows 打开，确认标题层级、字体、图片题注、表格边线和段落首行缩进。
2. `Ctrl+A`、`F9`；选中表格再 `F9`；目录选择“更新整个目录”；确认 `图1-1` / `表1-1` / `式（1-1）` 与第二章 `2-1`，没有“错误！未找到引用源”。
3. 在第一章图1-1前新增一张新图和题注，全部更新后，原图与其正文 `REF` 应变为图1-2；表/公式同理。**不可只观察缓存文本**。
4. 临时删除某一目标题注书签并更新，确认错误可被 `audit-references` 发现；恢复原文或撤销后重新验收。
5. 插入/删除引文并调换顺序，执行 `audit-references`；如果出现 `citation-order-mismatch`，人工核对并重排书目，而不是盲目重编号。
6. 把显式续表分为两页及三页，检查“续表”出现在每个后续表的表题前，表头与三线表完整，表注仅最后一段出现。
7. 在 WPS 文字中执行同一轮更新，另存 DOCX，使用 `qa_cross_compat.py` 对 WPS 存档审计 `SEQ/REF`、书签、表格结构；逐页对比 Word 版。
8. 更新 Word/WPS 后使用 PDF 导出作为视觉冻结副本；归档宿主版本号、字体版本、页面截图、差异和确认人。

## 自动化执行

```bash
python scripts/qa_cross_compat.py file.docx --roundtrip-libreoffice --report qa-interop.json
python scripts/docx_ops.py audit-references file.docx --strict-review
python scripts/docx_ops.py audit-long-tables file.docx --strict-review
python scripts/render_preview.py file.docx --out-dir qa-render
```

`qa_cross_compat.py` 的 `microsoft_word.status` 和 `wps_writer.status` 应保持 `not-tested`，直至在对应软件中完成上述步骤并附带真实复核记录。不能由 CI、PDF 或 LibreOffice 对这两项自动宣称通过。

权威参考：[Microsoft Word—交叉引用](https://support.microsoft.com/en-us/word/create-a-cross-reference)、[Microsoft Word—更新字段](https://support.microsoft.com/zh-cn/word/update-fields)、[WPS 交叉引用常见问题](https://plus.wps.cn/blog/p120323.html)。
