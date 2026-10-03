---
name: dut-thesis-formatter
description: >
  把中文论文/报告/学位论文草稿的 .docx 严格改写为《大连理工大学硕士学位论文格式规范》
  （DUT 硕士论文格式），也可从零生成符合规范的论文骨架。处理内容包括：A4 页边距与分节、
  封面、中英文摘要与关键词、目录/图目录/表目录域、章-节-小三 级标题、正文首行缩进与行距、
  三线表、图题表题（中英文双行、分章编号）、公式居中+行末编号、参考文献（GB/T 7714，
  按引用顺序重排、删未引用条目、右上角标）、附录/致谢/攻读学位期间成果、奇偶页页眉、
  罗马+阿拉伯双套页码。当用户提供一份 .docx 并要求按大连理工/大工/DUT 论文格式排版、
  套用学位论文模板、修改订正论文格式，或要求按学校格式规范整理时触发。即使用户没有明说
  「格式模板」，只要涉及把中文文档改成大连理工大学硕士学位论文格式、调整字体字号行距
  页码页眉目录、把表格改成三线表、规范参考文献格式，也应使用此技能。
license: MIT
---

# dut-thesis-formatter — 大连理工大学硕士学位论文格式改写

把一份内容杂乱的中文 `.docx`（论文草稿、学院模板、拼凑的文档）**就地改写**成严格符合
《大连理工大学硕士学位论文格式规范》的成品；也支持从零组装新论文。

核心是**复用现成脚本**（`scripts/dut_thesis.py`），而不是每次重写 python-docx 的
分节/页码/三线表/`w:eastAsia` 字体那套样板代码——那是最容易出错、最费 token 的部分。

## 安装（若提示词要求你安装本技能）

本技能遵循 `SKILL.md` + YAML frontmatter 约定，适用于 Claude Code、DSH 及其它
支持该约定的 Agent 工具。用户让你安装时按下面做，**不要凭猜测写路径**：

1. **先确定技能目录**，拿不准就把候选路径列给用户确认。
   常见约定：
   - Claude Code：`~/.claude/skills/`（用户级）、`<项目>/.claude/skills/`（项目级）
   - Codex：`$CODEX_HOME/skills/`（默认 `~/.codex/skills/`）
   - DeepSeek Harness：`$DSH_HOME/skills/`（默认 `~/.dsh/skills/`）、`<项目>/.dsh/skills/`
   - 其它工具：查它的 skills / extensions 配置

   注意：本仓库的 `SKILL.md` 在**仓库根目录**，Codex 自带的 `skill-installer`
   （`--path` 只接受子目录）装不了，Codex 请直接用 `git clone`。
2. **克隆时目录名必须是 `dut-thesis-formatter`**，与 `SKILL.md` 里
   frontmatter 的 `name:` 严格一致，否则有的工具加载不到。已存在同名目录时先问用户。
3. **装依赖**到本 Agent 实际使用的解释器：`pip install python-docx`
4. **自检**并报告输出：

   ```python
   from docx import Document
   from dut_thesis import reformat, audit, print_report
   ```

仓库：<https://github.com/chen-xing512/DUT-thesis-formatter.git>
完整安装指南：[`docs/INSTALL.md`](docs/INSTALL.md)

## 前置依赖

```bash
pip install python-docx
```

本技能**不需要 matplotlib**——它的职责是「格式合规」，不是「数据可视化」。
（若用户另外要求把表格数据做成图，再单独装 matplotlib，规范里图表只是插图，
图的质量要求见 `references/dut-format-rules.md` 第九节。）

## 权威依据（冲突时优先级从高到低）

1. `references/dut-format-rules.md` — **本技能硬编码的规则表，唯一事实来源**
2. 用户手上提供的 `大连理工大学硕士学位论文格式规范.docx`（研究生院，2026-04-15 版）
3. `DUTLatexSS/DUT-thesis-grd.cls`（LaTeX 模板，仅供旁证；其左边距 2.7 cm、章标题三号
   与 docx 说明不一致，**以 docx 为准**）

开始动手前**必须先读** `references/dut-format-rules.md`。

## 工作流

### Step 0：先判断任务类型

| 情况 | 做法 |
|---|---|
| 用户给一份已有的 `.docx`，要求改成 DUT 格式 | **`reformat` 模式** → Step 1–4 |
| 用户要新建一篇 DUT 论文 | **`template` 模式** → `scripts/cli.py template`，再写内容 |
| 用户只想检查格式问题 | **`audit` 模式** → `scripts/cli.py audit`，输出问题清单 |

### Step 1：读输入，摸清结构

```python
from docx import Document
doc = Document('论文.docx')
print(len(doc.paragraphs), len(doc.tables), len(doc.sections))
for p in doc.paragraphs[:60]:
    print(repr(p.style.name), p.text[:80])
```

要留意：

- **纯文本框排版的 docx**：`python-docx` 读不到正文，需要解压 `word/document.xml`
  用 `zipfile` + 遍历 `w:t` 取文本。但**格式规则本身已硬编码**，无需再解析模板。
- **表格是粘贴的图片**：规范要求「表要用 WORD 绘制，不要粘贴」，发现后要提示用户。
- **公式是图片或纯文本**：规范要求用 Word 自带公式编辑器（OMML），图片公式要提示重做。

### Step 2：一键改写

```python
from dut_thesis import reformat, print_report, enable_update_fields, audit
from docx import Document

doc = Document('论文.docx')
stats = reformat(doc, title_cn='基于XX的YY研究')   # ← 中文题目用于偶数页页眉
print_report(stats)

# 若用户还需要目录/页码自动更新
enable_update_fields(doc)
doc.save('论文_DUT.docx')

print_report(audit(Document('论文_DUT.docx')))     # ← 交付前必做
```

`reformat()` 会自动完成：

1. 页面设置：A4、上 3.5 / 下 2.5 / 左 2.5 / 右 2.5 cm、页眉 2.5 cm、页脚 2.0 cm
2. 建立全部官方段落样式（`摘要题目` / `图名中文` / `参考文献正文` / `公式` …）
3. **逐段分类**（`classify_paragraph()`）后套用对应格式：
   - 章标题 → 黑体三号居左，段后 1 行，`outlineLvl=0`，**段前分页**
   - 节标题 → 黑体四号居左，段前 0.5 行，`outlineLvl=1`
   - 小节标题 → 黑体小四居左，段前 0.5 行，`outlineLvl=2`
   - 居中章级标题（摘要/ABSTRACT/目录/参考文献/附录/致谢/成果）→ 小三黑体加粗居中，段后 1 行
   - 正文 → 宋体小四，两端对齐，**1.25 倍行距**，首行缩进 2 字符，`snapToGrid=0`
   - 图题表题 → 宋体五号居中，表题在表上方、图题在图下方
   - 参考文献条目 → 宋体五号，悬挂缩进，1.25 倍行距
4. 所有表格 → **三线表**，无竖线，上线/下线 1.5 磅、表头线 1 磅，表内五号居中
5. 分节页码：前置罗马数字、正文阿拉伯数字从 1 起、封面无页码
6. 奇偶页页眉：奇数页「大连理工大学硕士学位论文」，偶数页论文中文题目，宋体五号居中
7. 参考文献：按**正文首次引用顺序**重排、删除未引用条目
8. 源文档若有「目录」标题但没有目录域，自动在其后补一个 `TOC` 域
   （只新增域，不删你已有的静态目录条目）

### Step 3：补内容（不要只做格式）

格式改完后，这些**规范硬性要求**必须一并检查并补齐，否则论文送审仍会被打回：

- [ ] **摘要限一页**，且**不得出现插图**；摘要后空一行再列「关键词：」（3–5 个，`；` 分隔，末尾无标点）
- [ ] **英文摘要**与中文摘要对应，`Key Words：` 加粗，`;` 分隔
- [ ] **参考文献 ≥ 50 篇，英文 ≥ 60%，近五年 ≥ 20%**；作者超 3 位只列前三位 + `et al`/`等`
- [ ] **正文引用用右上角标** `[1]`、`[2-4]`（`add_superscript_citation()`），未引用条目已删
- [ ] **每章另有引言性文字 + 本章小结**
- [ ] **附录 / 致谢 / 攻读硕士学位期间科研项目及科研成果**齐全（致谢不可省略）
- [ ] 量与单位用**法定符号**（GB 3100~3102-93），无废弃单位；数字用阿拉伯数码
- [ ] 外文斜体/正体、标量正体、向量黑体符合第八节说明
- [ ] 目录首页须落在**奇数页**（`section.start_type = ODD_PAGE`）

### Step 4：校验（交付前必做）

```python
from dut_thesis import audit, print_report
print_report(audit(Document('论文_DUT.docx')))
```

`audit()` 会检查：页面参数、`evenAndOddHeaders`、`updateFields`、TOC/PAGE 域是否存在、
标题层级是否齐全、标题 `outlineLvl`、正文中文字体与首行缩进与网格对齐、表格是否三线表。

**不要只信「保存成功」**，一定要重新打开产物读回来验证。命令行等价做法：

```bash
python scripts/cli.py audit 论文_DUT.docx    # 有问题时退出码为 1
```

## 参考文件

- **`references/dut-format-rules.md`** — 完整格式规范（页面/字号/标题/图表/公式/参考文献/
  页眉页脚/附录）。**开始 Step 2 前必读。**
- **`scripts/dut_thesis.py`** — python-docx 工具库：字体（含 `w:eastAsia`）、段落、
  三线表、公式制表位、TOC 域、奇偶页页眉、分节页码、封面、段落分类、改写、校验。
- **`scripts/cli.py`** — 命令行入口：`reformat` / `audit` / `template`。
- **`scripts/tests/test_smoke.py`** — 端到端冒烟测试（生成 → 改写 → 校验）。
- **`scripts/make_demo.py`** — 造内容完整的示例论文（`docs/showcase.docx`）。
- **`scripts/make_demo_real.py`** — 把渲染出的 PDF 合成 `docs/demo.png`（需 matplotlib + pymupdf）。
- **`scripts/convert_with_word.sh`** — 调本机 Word 把 docx 转 PDF（须在沙箱外运行）。

## 质量检查清单（输出前自检）

- [ ] 页面：A4，上 3.5 / 下 2.5 / 左 2.5 / 右 2.5 cm，页眉 2.5 cm，页脚 2.0 cm
- [ ] 封面无页码；前置部分罗马数字 I、II…；正文阿拉伯数字从 1 起
- [ ] `w:evenAndOddHeaders` 已写；奇数页页眉校名，偶数页页眉论文中文题目，宋体五号居中
- [ ] 页脚页码底部居中、Times New Roman 小五，形如 `- 7 -`
- [ ] 章标题黑体三号居左 + 段后 1 行 + 每章另起一页；节黑体四号；小节黑体小四
- [ ] **三级标头都不加粗**（规范只说「黑体」这个字体名，官方模板也没有 `w:b`）
- [ ] 标头的字号/字重/字体/颜色**有效格式**符合规范——
      即已排除「Word 样式面板被改」与「run 上残留直接格式」两种情况
- [ ] 正文宋体小四、两端对齐、1.25 倍行距、首行缩进 2 字符（`firstLineChars="200"`）、
      **取消对齐到网格**（`w:snapToGrid val="0"`）
- [ ] 所有中文字体都写了 `w:eastAsia`（只设 `run.font.name` 是无效的）
- [ ] **所有文字颜色是黑色**（`w:color="000000"`）；标题没有继承到主题蓝
- [ ] 三线表：无竖线、上线/下线 1.5 磅、表头线 1 磅、表内五号居中、1.25 倍行距、无首行缩进
- [ ] 表题在表**上方**（中英文两行，宋体/TNR 五号居中），图题在图**下方**（同样两行）
- [ ] 图、表、公式均**分章编号**（`图 1.1` / `表 3.2` / `(3.1)`），编号与文字间留 1 空格
- [ ] 公式居中、编号在行末右对齐（居中制表位 + 右对齐制表位）
- [ ] 参考文献标题小三黑体居中；条目五号宋体、悬挂缩进、按引用顺序、无未引用条目
- [ ] 正文引用为**右上角标** `[n]`
- [ ] TOC 域 `TOC \o "1-3" \h \z \u` + `w:updateFields="true"`；「目录」标题自身不列入目录
- [ ] 附录 / 致谢 / 攻读学位期间成果齐全，各自小三黑体居中标题 + 小四正文
- [ ] 已用 `audit()` 重新打开产物校验，而非仅凭「保存成功」

## 示例触发语

- 「把这份论文草稿按大连理工大学硕士学位论文格式规范排版：[路径.docx]」
- 「严格按研究生院的学位论文模板修改订正这份 docx，页眉页码目录都要对」
- 「帮我检查这篇论文哪里不符合大工硕士论文格式要求」
- 「新建一份大连理工大学硕士论文的 Word 骨架，题目是……」

## 三层都要管：样式 → 字符 → 段落

**只「套样式名」是不够的**——Word 里样式只是继承基线，下面两种改动都会盖过它：

1. 用户在 Word「样式」面板里改了 Heading 1/2/3（字号、加粗、间距）→
   所有同类标题一起变形；
2. 从别处粘来的标题自带 run 级直接格式（`w:sz` / `w:b` / `w:spacing`）→ 优先于样式。

`reformat()` 与 `add_heading()` 通过 `enforce_heading_format()` 三层处理：
样式层把规范值写全、字符层清掉一切直接格式（只留字体名）、段落层删掉间距对齐覆盖。
`audit()` 会逐个标题核对**有效格式**并列出不符项。

规范值集中在 `HEADING_SPEC`（字号/加粗/段前后行数/行距），样式、段落、校验三处共用一份，
避免各处各写一套而对不上。

## 两个只有真实渲染才看得出来的坑

改完格式后**务必渲染成 PDF 目视核对**（见 `docs/RENDERING.md`）。以下两类问题
单测和 XML 检查都发现不了，但纸面上一眼就能看出：

1. **标题变蓝**。python-docx 内置模板的 Heading 1/2/3 自带主题蓝
   （`365F91` / `4F81BD`），规范要求标题是「黑体」（字体 + 黑色）。
   `apply_style_set()` 已经把所有样式的 `w:color` 钉成 `000000`；
   `audit()` 也会报「N 处文字不是黑色」。若你自己新建样式，记得调
   `strip_style_color(style)`。
2. **页眉页脚串节**。`doc.add_section()` 会把前一节的
   `headerReference`/`footerReference` 复制给新节（两节共用部件），
   加上 python-docx `is_linked_to_previous` 的反直觉语义，会出现
   「封面长出页眉」「偶数页页码消失」「每节页码都从 1 重新开始」。
   用 `add_section(doc, kind='front'|'body'|'cover')` 即可正确处理。

## 已知边界

- 公式若原本是**图片或纯文本**，本技能不会凭空生成 OMML——会保持内容并把行内格式改对，
  同时提示用户在 Word 公式编辑器里重做。`add_equation(omml_xml=...)` 支持传入 OMML。
- 官方模板文件里的**表头下线实际是 0.5 磅**（`w:sz="4"`），而规范正文明确写 **1 磅**。
  本技能按规范正文的 1 磅执行（`w:sz="8"`）；需要贴合模板文件时传 `header_sz=4`。
- 封面中「大连理工大学印刷厂统一制作」的封皮不在电子稿范围内，`add_cover()` 生成的是
  封一（论文题目 + 字段块 + 页脚校名）。
- 「创新点」等学院自加章节规范里未规定，保持用户原有层级即可。
