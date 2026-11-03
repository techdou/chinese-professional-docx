# 规范来源与边界

以下为本次维护核对的一手来源，访问于 2026-10-14。官方文档会更新；以下是格式及调用原则依据，不是对某一客户端未来版本的保证。

- Agent Skills Specification: https://agentskills.io/specification 。约束 SKILL.md、name/description、可选字段、目录与渐进加载。description 1–1024 字符；compatibility 若提供为 1–500 字符；metadata 使用字符串键值。规范链接的 skills-ref 是格式参考校验器，不是模型匹配评测器。
- OpenAI skills documentation: https://developers.openai.com/codex/skills/ （当前重定向到 https://learn.chatgpt.com/docs/build-skills）。说明显式/隐式调用、描述前置任务关键词、可选 agents/openai.yaml 与本地加载范围。插件分发和本机目录安装不是同一安装范围。
- Anthropic authoring guide: https://platform.claude.com/docs/en/agents-and-tools/agent-skills/best-practices 。用于简明指令、按需加载、明确脚本接口、评测迭代及边界案例。
- Claude Code skills: https://code.claude.com/docs/en/skills 。用于客户端发现与用户级目录说明。宿主扩展字段不添加到通用 frontmatter 以避免不必要耦合。

学术排印的参照层级：公开发布的高校研究生院《研究生学位论文写作格式规范》及其自然科学、人文社科官方模板附件（已实读核对；原文件不入库、不列具体院校），以及国家标准 GB/T 7713.1—2025 的现行状态 https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=36C05B5738C54D42B4B262525320B52F （2026-02-01 实施）。核到的明确内容：自然科学章节建议使用 1/1.2/1.2.1、人文社科建议第一章/一/（一）；表的编排一般使用三线表，上下线 1.5 磅、表头线 1 磅，必要时可加辅助线；章标题另起一页、三级标题字号等为部分院校的特定要求。这里只参考其专业范式，**不声称通用 Skill 自动满足学校正式学位论文所有规定**。代码浅灰底纹属于本 Skill 自选设计，不是引述的学校强制条款。

本项目使用上面共同支持的核心格式，不采用必须依赖特定宿主的动态 shell 插值、隐藏命令或免审批工具授权。agents/openai.yaml 和评测夹具是可选适配/工程资产，不声称为跨客户端必需的官方字段。

Word 字体、字号、页面、缩进来自内置通用母版 v1.1，资产摘要记录在 assets/provenance.json。它借鉴学术文档风格，但不是任何院校官方学位论文模板；本次核验了公开规范来源和现行国标状态，并只更新通用母版的视觉样式（代码/表格/黑白装饰）；原有页面、正文缩进、字体与字号参数保留；新增原生标题编号范式与可选安全修复。正式学校/赛事文档应逐项读取其指定材料。

未附带或授权分发任何字体文件、学校标识或正式声明。公开发布需由权利人决定；私人文档操作不自动授权 GitHub 发布、云端同步或本机软件安装。


## v1.6 实读与标准更新

- 公开高校《研究生学位论文写作格式规范》及附件，规范了分章图表、续表表头重复及顺序编码上标。仅作为通用范式的高质量来源，不强加校方特有格式。
- [国家标准信息公开服务平台：GB/T 7714—2025](https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=C6CE52E55AC09B9C79A20AEA77CEDD14)，2026-07-01 实施，代替 2015 版。
- [Microsoft Word：创建交叉引用](https://support.microsoft.com/en-us/word/create-a-cross-reference)；[Microsoft Word：更新字段](https://support.microsoft.com/zh-cn/word/update-fields)；[WPS 交叉引用](https://plus.wps.cn/blog/p120323.html)。
