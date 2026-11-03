# 验证与交付

## 结构与内容

检查每个被识别普通正文段落的有效 firstLineChars=200；检查继承、直接覆盖、hanging 冲突和与字号匹配的 twip 后备值。代码/标题/列表/表格/题注/公式/页眉页脚/引用不套正文缩进。段落 role 不明时要读取并确认，不能因为 audit 没有 errors 就声称覆盖全部正文。

只修缩进使用 scope=indent；全文通用排版才比较通用样式。用户指定模板的页面或字号不同不自动判违规。审计可只读完成，不能为了证明工作量无必要地改写文件。

仅排版需对比前后文本、关键数字、公式、链接、引用、媒体与重要域；清理空白也需限定真实段首。文件需通过 ZIP、OOXML 内容类型及打开检查。DOTX 做真实转换/新建测试，不以扩展名代替。

完成稿应替换占位文本；母版有意保留的起稿提示无需删除。待补内容应明确，不捏造成果或引用。

## 黑白学术样式专项核验

- 对照规范：官方学校模板优先；通用生成检查正文黑色、纸张白色，代码样式段落底纹 `F5F5F5` 且无彩色侧线；正文不受代码格式影响。
- 表格：顶线、表头线、底线三条黑色水平规则；无竖线、无斑马纹；表题在上，表头重复。长表跨页续表文本需人工处理，不能因为重复了表头就声称完整自动续表。
- 不对用户原稿或学校模板做隐式全局变更。专门的 `format-tables --all` 需要有明确格式变更意图。

## 逐页视觉

渲染最终修改版本，打开每页图像：确认行首缩进、字体缺字、标题孤行、表格宽度/跨页、页码/题注、留白与裁切。每次版式修改后重新渲染。不能用“PNG已生成”代替逐页查看。

缺字体时原稿仍保留指定字体；可用 preview 副本 fallback-cjk，并声明替代预览。无渲染能力时写“结构检查完成，视觉未核验”，不能报告全量通过。不得打包字体。

## 分开记录

| 维度 | 允许状态 |
| --- | --- |
| 结构与缩进 | 通过 / 有错误 / 部分覆盖 |
| 内容保真 | 已比较 / 有差异需解释 / 未比较 |
| 视觉 | 逐页已复核 / 待复核 / 渲染失败 |
| 字体 | 期望字体实测 / 替代字体 / 未检查 |
| 客户端加载 | 真实宿主已测 / 仅安装复制 / 未测试 |
| 学校或赛事合规 | 已逐项比对的具体范围 / 未评价 |

审计 JSON 与退出码只是自动化证据的一部分。strict-review 的退出 3 要解决警告或披露范围，不能忽略再写“全部完成”。

## 警告分类：哪些要修、哪些按设计存在

audit 的 code 分四个性质，不要试图把警告计数清零：

| 类别 | code | 处理 |
| --- | --- | --- |
| **可修复错误**（计入 errors，退出 1） | `first-line-chars`、`hanging-conflict` | 按语义修正样式或段落（见 [audit-rules.md](audit-rules.md) §4） |
| **可修复警告** | `fallback-indent`、`normal-font-size`、`chinese-font`、`latin-font`、`line-spacing`、`alignment`、`heading-size` | 同上，按语义修 |
| **按设计存在或仅需披露** | `manual-review` 且 role 为 `mixed-drawing`（含图段落）、`unclassified`（非 special 键命名的独立样式段）、`equation`、`content-control`；`page-margins`、`page-size`（用户模板可有意不同，实现注明 may be intentional）；`external-relationships`（包级披露，不跟踪） | **不修文档**；在交付说明中披露豁免范围。把参考文献改回 Normal 来消警告是倒退 |
| **真待复核** | `manual-review` 且 role 为 `ambiguous-manual-list`（疑似手工编号列表）；`manual-padding`（行首空格/Tab）；`soft-break`（Shift+Enter 模拟分段）；`edit-blocked`（包带保护/签名，拒绝修改） | 打开段落/包确认意图：转真列表 / 清填充（需授权 `--strip-manual-padding`）/ 真实分段 / 放弃修改 |

`textbox` 角色不产生 manual-review 警告（文本框段落不参与正文审计，见 audit-rules.md §1 注记）。"按设计存在"类警告在 strict-review 下也会触发退出 3——正确动作是披露范围，不是反复修改文档。

## 产物治理

优先交 DOCX；需要模板再交 DOTX，批量母版维护可打包。PNG/PDF、审计报告与测试原稿默认属于内部 QA，保留在 .local 或当前会话检查目录，不提交 Git。仓库仅放 Skill 源码、经过确认的通用资产、用户使用指南和测试；不放凭据、本地软件配置或个人科研原稿。字体治理例外：`assets/fonts/` 为所有者授权的本机字体资产（见 SKILL.md“字体资产”节），仅在本机用于预览注册与文档嵌入，不出现在任何交付物中；若 skill 仓库公开分发，字体目录按许可要求另行处理。


## v1.6 引用与长表质检

使用 `audit-references` 判定 dangling REF、重复 bookmark、没有上标等问题；用 `audit-long-tables` 检查表头重复、续表显式分段和 missing linked captions；用 `qa_cross_compat.py` 验证 LibreOffice 往返，并在 Windows Word/WPS 中人工更新全部字段及视觉检查。必须提供真实宿主验收记录方可声称跨平台通过。执行清单：[兼容性矩阵](../docs/compatibility-matrix.md)。
