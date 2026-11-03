# 匹配与执行评测

本地单元测试不等于 Agent 触发率。这里分三层：规范/资源静态检查；真实文件生成与修复测试；独立客户端/模型的实际触发及任务执行。各层结果独立记录。

`evals/trigger-cases.json` 包含 40 个带理由的正反例，含上下文、省略称谓、硬负例、指定模板、多产物和中英指令。它们是测试集，不是已完成的模型结果。`evals/evals.json` 提供 8 个执行任务；缺少模板的任务应询问原件，不应生成文件。

源码工具 `tools/evaluate_triggers.py prepare --out .local/blind-cases.json` 导出不含 should_trigger、reason、expected_route 的盲测输入。让独立宿主在技能已安装可发现的会话运行；初次匹配只能见元数据，不能预先塞入 SKILL.md 全文或预期标签。正负任务使用相同模型、客户端设置及其他技能，记录版本、重复次数和真实 trace。

每个观测为 JSONL：case_id、mode（host_invocation 或 metadata_probe）、host、model、skill_sha256、selected_skills（数组）、trace_file（相对观测文件路径）。可选 route。host_invocation 记录实际技能加载/调用；metadata_probe 仅说明模型面对元数据的选择，不能当作客户端真的调用。

`python tools/evaluate_triggers.py score --observations .local/observations.jsonl` 检查 case_id、证据文件、重复项和技能摘要，按 host/model/mode/skill 摘要分别统计 TP/FP/TN/FN、准确率、精确率、召回率及测试覆盖率。没有日志不生成成功率；证据文件存在只是可追溯条件，不等同于已验证其内容真实性。

A/B：对 v1.0.1 与本版使用同一批盲测输入、同样的其他技能和模型；旧描述保存在 evals/baselines/v1.0.1.json。先运行各自实际宿主，再比较结果。测试集中存在的反例不能全部改成描述中的逐字匹配词。

真实文件评测：分别检查产物存在、角色分类、两字符缩进、内容保真、模板类型、逐页视觉与字体状态。自动预览生成不等于视觉通过。测试日志、原稿和内部 Review 放 .local/，不提交到 Git。
