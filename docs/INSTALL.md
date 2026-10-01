# 安装 dut-thesis-formatter

这是一个遵循 `SKILL.md` + YAML frontmatter 约定的 Agent Skill，可用于
Claude Code、DeepSeek Harness（DSH）以及其它支持该约定的 Agent 工具。

**唯一依赖**：`python-docx`

```bash
pip install python-docx      # 或：python -m pip install python-docx
```

> 若你的 Agent 跑在虚拟环境里，请装到**它使用的那个解释器**，
> 而不是系统 Python。

---

## 方式一：让 Agent 自己装（推荐）

把下面整段贴给 Agent 即可，它会自己找技能目录、克隆、装依赖、验证：

```text
请帮我安装一个 Agent Skill：

仓库：https://github.com/chen-xing512/DUT-thesis-formatter.git
技能名：dut-thesis-formatter

要求：
1. 找到本 Agent 工具的技能目录（skills 目录）。若无法确定，
   先把探测到的候选路径列给我确认，不要自己猜着写。
   常见约定见下：
     - Claude Code      ~/.claude/skills/
     - DeepSeek Harness $DSH_HOME/skills/（默认 ~/.dsh/skills/）
     - 项目级           <项目根>/.claude/skills/  或  <项目根>/.dsh/skills/
2. 把仓库克隆到技能目录，目录名必须是 dut-thesis-formatter
   （必须与 SKILL.md 里 frontmatter 的 name 一致，否则可能加载不到）。
   已存在同名目录时先告诉我，不要直接覆盖或删除。
3. 确认技能目录里有 SKILL.md，且 frontmatter 的 name / description 完整。
4. 检查 python-docx 是否已安装；缺的话装到本 Agent 实际使用的解释器里。
5. 用这份最小代码自检，并报告输出：
       from docx import Document
       from dut_thesis import reformat, audit, print_report
6. 最后告诉我：技能装到了哪个绝对路径、依赖是否就绪、
   以及是否需要重启会话才能生效。
```

### 为什么这样写

几个容易踩的点，提示词里已经显式约束了：

- **技能名要对齐 frontmatter**：克隆出来的目录名如果和 `SKILL.md` 里的
  `name:` 不一致，有些工具会加载不到或名字错位；
- **别让 Agent 猜目录**：不同工具约定不同、有些还会读环境变量
  （例如 DSH 的 `$DSH_HOME`），所以要求它"不确定就先问"；
- **别让 Agent 直接覆盖**：避免把已有的同名技能删掉；
- **装完要自检**：确认 `SKILL.md` 与 `dut_thesis.py` 真的能被导入。

---

## 方式二：自己克隆（通用）

先确认你的工具用的是哪个技能目录：

| 工具 | 用户级 | 项目级 |
|---|---|---|
| Claude Code | `~/.claude/skills/` | `<项目>/.claude/skills/` |
| DeepSeek Harness | `$DSH_HOME/skills/`（默认 `~/.dsh/skills/`） | `<项目>/.dsh/skills/` |
| 其它工具 | 查该工具的 skills / extensions 配置 | 同上 |

```bash
# 用户级（所有项目可用）——把 <技能目录> 换成上表里对应的一列
git clone https://github.com/chen-xing512/DUT-thesis-formatter.git \
  <技能目录>/dut-thesis-formatter

# 例：Claude Code 用户级
git clone https://github.com/chen-xing512/DUT-thesis-formatter.git \
  ~/.claude/skills/dut-thesis-formatter

# 例：DSH 项目级
git clone https://github.com/chen-xing512/DUT-thesis-formatter.git \
  .dsh/skills/dut-thesis-formatter
```

装完确认一下：

```bash
ls <技能目录>/dut-thesis-formatter/SKILL.md
```

## 方式三：打包成 `.skill` 再导入

适合不支持直接读目录、只接受压缩包导入的工具，或需要离线分发时：

```bash
git clone https://github.com/chen-xing512/DUT-thesis-formatter.git
cd dut-thesis-formatter
python scripts/package_skill.py
# 生成 dist/dut-thesis-formatter.skill
```

然后把 `.skill` 文件交给你的工具导入。仓库打 `v*` 标签时，
GitHub Actions 会自动跑测试、打包并发布到 Releases，也可以直接从
[Releases](https://github.com/chen-xing512/DUT-thesis-formatter/releases) 下载。

---

## 安装后怎么用

**不需要手动点名技能**——触发条件写在 `SKILL.md` 的 `description` 里，
Agent 会根据你的意图自动匹配。给 Agent 一份 `.docx` 并说明意图即可，例如：

- 「把这份论文草稿按大连理工大学硕士学位论文格式规范排版」
- 「严格按研究生院的学位论文模板修改订正这份 docx，页眉页码目录都要对」
- 「帮我检查这篇论文哪里不符合大工硕士论文格式要求」

也可以直接用命令行（不经过 Agent）：

```bash
python scripts/cli.py reformat 论文.docx -o 论文_DUT.docx --title "论文中文题目"
python scripts/cli.py audit 论文_DUT.docx
python scripts/cli.py template 骨架.docx --title "题目" --author 张三
```

## 卸载

删掉技能目录即可：

```bash
rm -rf <技能目录>/dut-thesis-formatter
```
