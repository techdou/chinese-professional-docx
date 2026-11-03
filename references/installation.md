# 安装与宿主发现

安装对象是包含 `SKILL.md`、`scripts/`、`references/`、`assets/` 的完整目录。独立 Skill ZIP 解压后可直接复制；源码仓库的安装源是 `skills/chinese-professional-docx/`。

## 本地宿主

Codex 用户级：`~/.agents/skills/chinese-professional-docx/`；项目级：`<项目>/.agents/skills/chinese-professional-docx/`。

Claude Code 用户级：`~/.claude/skills/chinese-professional-docx/`；项目级：`<项目>/.claude/skills/chinese-professional-docx/`。

不要同时放置多个重复名称的副本。已有不同内容时先比较，明确授权后备份替换。源码仓库可运行 `bash setup.command`，或使用 `python3 tools/install_skill.py --target codex --install-deps`。升级使用 `--replace`，安装器会先备份。依赖安装需要网络且只在明确授权后执行。

需要无交互预检时使用 `--dry-run`；不把计划输出当成已执行。本机安装结果也不等于 ChatGPT 账号或插件永久安装。

## 依赖与路径

Python 3.10+；生成/修复依赖 `scripts/requirements.txt`；预览另需 LibreOffice 和 `scripts/requirements-render.txt`。执行前可运行 `scripts/doctor.py`，它只读、不安装、不联网。

安装器的 `--install-deps` 把依赖放在安装目录 `.venv/`。Agent 优先使用该解释器，不能仅因全局 Python 缺少库便判定整套 Skill 不可用。脚本及资产路径相对实际 Skill 文件位置解析，不依赖用户当前目录。

## 确认能否调用

在目标宿主的技能列表确认名称和路径。Codex 可通过 `/skills` 或 `$chinese-professional-docx` 显式选择；Claude Code 独立安装用 `/chinese-professional-docx`，插件形式遵循其命名空间规则。未发现时按宿主文档重新加载或重启，而不是无证据声称安装失败/成功。

确认加载后，运行一个无敏感数据的实际文档任务，查看是否读取入口、选择正确分流、执行检查并交付文件。自然语言匹配需使用真实宿主日志评估；40 条触发用例本身不代表成功率。

`agents/openai.yaml` 与 `.claude-plugin/plugin.json` 是可选宿主扩展，不是通用 Skills 格式的强制文件，也不会自动授予操作权限。
