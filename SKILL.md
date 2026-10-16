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
9. **合并被拆开的关键词行**：源文档若把「关键词：」题头单独占一段、内容另起一段，
   先把两段并成一段（中英文各一次），再做格式；规范要求题头与关键词同一行
10. **`toc N` 目录条目一律跳过不改写**（文字来自 TOC 域、包在 `w:hyperlink` 里，
    改写会把文字写两遍并让条目变成 Heading），只把 `toc N` 样式统一成小四宋体
11. **收尾按文字兜底钉正前置标题**：摘要 / ABSTRACT / 目录 / 参考文献 / 致谢 / 附录 /
    成果标题不论原来套了什么样式、写成什么变体（`References`、`Acknowledgements`、
    `谢辞`…），最后都强制居中小三黑体，并清掉标题上的列表编号
12. 参考文献条目**左对齐**（避免长 URL 把上一行字距拉爆），网址/邮箱/带点数字
    在补标点空格时不被拆断
13. **封面（封一）单独处理、主循环整块跳过**：封面标题保持「华文细黑 24/22/16 pt
    居中」（绝不按正文套成小四宋体 + 首行缩进），字段块只保留值单元格下边框这条
    **填写线**（绝不转成三线表），封面页脚保留「大连理工大学」校名（`style='none'`
    会连校名一起清掉，不能用在封面）
14. **插图段落一律单倍行距**（`has_drawing()` → `CH_FIGURE`）：只要段落里有
    `w:drawing`/`w:pict`，行距就不能是固定值（`exact`）——否则嵌入型图片会被
    行框裁掉一半（Word 里选中图片按 Ctrl+1 恢复的就是这个）；主循环之后再跑一遍
    `unclip_figures()` 兜底

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
- [ ] **封面**：三行标题是「华文细黑 24 / 22 / 16 pt 加粗居中」（**不是**宋体小四带首行缩进）；
      字段块是**填写横线**（值单元格下边框 0.5 磅）**不是三线表**；封面无页眉、无页码，
      页脚有「大连理工大学」校名
- [ ] **插图段落是单倍行距、不是固定行距**（`w:lineRule` 不能是 `exact`），
      否则嵌入型图片会被行高裁掉一半（`audit()` 会报「含图的段落用了固定行距」）
- [ ] 表题在表**上方**（中英文两行，宋体/TNR 五号居中），图题在图**下方**（同样两行）
- [ ] 图、表、公式均**分章编号**（`图 1.1` / `表 3.2` / `(3.1)`），编号与文字间留 1 空格
- [ ] 公式居中、编号在行末右对齐（居中制表位 + 右对齐制表位）
- [ ] 参考文献标题小三黑体**居中**；条目五号宋体、悬挂缩进 170、**左对齐**、按引用顺序、无未引用条目
- [ ] 条目里的网址 / 邮箱 / DOI **没有被补空格拆断**（`www.gov.cn` 不能变成 `www. gov. cn`）
- [ ] 摘要后「关键词：」题头与关键词在**同一段同一行**（中英文各一处）
- [ ] 致谢标题小三黑体**居中**（不能是 Heading 1 左对齐）
- [ ] 目录条目仍是 `toc N` 样式、文字**没有翻倍**、没有被套成 Heading 混进目录
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

## 只有真实渲染才看得出来的坑

改完格式后**务必渲染成 PDF 目视核对**（见 `docs/RENDERING.md`）。以下问题
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
3. **`w:pPr` / `w:rPr` 子元素顺序**。两者在 OOXML 里都是 `sequence`，
   python-docx 自己新增元素时按 `CT_PPr._tag_seq` 插入，手工
   `ppr.append(...)` 会写出不合规顺序（不同软件容忍度不一致）。统一走
   `set_keep_with_next()` / `set_page_break_before()` / `_set_ind()` /
   `set_outline()` 这些封装，它们内部会调 `_fix_order()` 归一化；
   `audit()` 也会报「w:pPr 子元素顺序不符合 OOXML schema」。
4. ⚠ **先确认导出的确实是 PDF**。`convert_with_word.sh` 的 `save as` 若把格式写成数字
   （如 15，实际是 wdFormatXMLTemplateMacroEnabled），Word 会生成一个**扩展名是 .pdf、
   内容却是 Word 模板**的文件；PDF 阅读库还能把它渲染出来，看着像回事，但页宽不是 A4、
   字体全是兜底字体、**所有分页符都失效**，据此"目视核对"会得出错误结论。
   必须写枚举名 `file format format PDF`，并用 `file 出.pdf` 确认显示 `PDF document`。
   脚本第 [3/3] 步已内建该校验。详见 `docs/RENDERING.md`。

## 回归清单：七个「只在真实文档上才暴露」的 bug

这七类问题都出在**真实论文**上（草稿的段落结构远比示例复杂），单测与 XML 抽查都不易发现，
但纸面上一眼就能看出。修完后都补了回归测试（`scripts/tests/test_smoke.py`），改动前先跑一遍。

| 症状 | 根因 | 现在的做法 |
|---|---|---|
| 「关键词：」独占一行，关键词另起一行 | 草稿把题头与内容写成**两段**，旧代码只处理「同段」的情形 | `merge_keyword_lines()`：改写前先把题头段与紧跟的内容段并成一段（中间空段一并清掉） |
| 参考文献 / 致谢标题**没有居中** | 这些标题在真实文档里已套 `Heading 1`（左对齐），或写成 `References`、`Acknowledgements`、`谢辞` 等变体、还带列表编号；只按样式 + 精确文字分类会漏判 | `_front_title_key()` 放宽识别；`enforce_front_titles()` 在收尾时**按文字再扫一遍**强制居中并清 `w:numPr` |
| 参考文献条目字距炸开（`国 办 发 （ 2023 ） 3 号`）、网址被写成 `www. gov. cn` | ① 条目用了两端对齐，长 URL 把上一行拉爆；② 补标点空格的规则伸进了 URL 内部 | 条目改**左对齐**（`REF_ALIGN`）；`format_reference_entry()` 先把 URL / 邮箱 / 带点数字「藏」起来，且句点两侧都是西文字母数字时不补空格 |
| 目录条目文字翻倍（`1 引言` → `1 引言 11 引言 1`）并被套成 `Heading 1` | 目录条目文字在 `w:hyperlink` 里，而 `p.runs` **不包含**超链接内的 run；「清空 runs 再 add_run(p.text)」等于把文字写两遍 | 目录条目判为 `CH_TOC` **一律跳过**；其余段落改用 `para_runs()`（含超链接）/ `replace_text()`；`toc N` 样式统一改字号，而不是改条目段落 |
| 参考文献重排后，正文的 `[6]` 指到了另一篇文献上 | `order_references()` 按引用顺序重排并重编号，却**没有同步改正文的标注**。真实论文的首次引用顺序可能是 `1,2,3,4,5,16,6,…`，重排后第 6 位换了文献，正文仍写 `[6]` | 重排时算出「旧编号 → 新编号」，用 `_remap_citations()` 逐 run 改写正文标注（保留上标格式与 `[2-4]` 合并写法）；`audit()` 新增「正文引用了不存在的编号」检查 |
| **封面被改烂**：标题变成小四宋体 + 首行缩进，字段块变成三线表（填写横线消失、多出上下两条粗黑线），页脚校名被清空 | ① 批量改写把封面段落当正文处理；② 所有表格一律走 `_normalize_table()`（三线表）；③ 封面节用 `set_page_number_footer(style='none')`，而 `'none'` 的语义是「清空页脚」 | `cover_region()` 以**字段表为锚**划出封面 → 主循环整块跳过；`normalize_cover()` 按规范 24/22/16 pt 居中钉正；`is_field_table()` 认字段块并只修边框；封面页脚改用 `_ensure_cover_footer()`（保留/补校名、去掉 PAGE 域） |
| **插图被隐藏一半**：只含图的段落被当空行设成固定行距 25 磅，图片只剩一条边 | 旧 `classify_paragraph()` 只看文字——`p.text` 为空就当空行，而插图段落恰恰没有文字；空行分支用的是 `w:lineRule="exact"` 固定行高，行框装不下图片就把图裁了。（Word 里选中图片按 Ctrl+1 改单倍行距即可恢复，就是这个原因） | `has_drawing()` 识别含图段落 → 新增 `CH_FIGURE` 分类，统一给**单倍行距**、不缩进；收尾再用 `unclip_figures()` 兜底（连「图与文字同段」一起覆盖）；`audit()` 新增「含图的段落用了固定行距」检查 |

> 顺带删掉了 `dut_thesis.py` 里一份**同名旧实现**（`strip_char_overrides` /
> `_effective_char_fmt` / `enforce_heading_format` / `heading_format_problems` 各写了两遍，
> 后定义者才生效）。同名函数并存时极易改错那一份——`grep` 到的第一处往往不是真正生效的那处。

## 已知边界

- 公式若原本是**图片或纯文本**，本技能不会凭空生成 OMML——会保持内容并把行内格式改对，
  同时提示用户在 Word 公式编辑器里重做。`add_equation(omml_xml=...)` 支持传入 OMML。
- 官方模板文件里的**表头下线实际是 0.5 磅**（`w:sz="4"`），而规范正文明确写 **1 磅**。
  本技能按规范正文的 1 磅执行（`w:sz="8"`）；需要贴合模板文件时传 `header_sz=4`。
- 封面中「大连理工大学印刷厂统一制作」的封皮不在电子稿范围内，`add_cover()` 生成的是
  封一（论文题目 + 字段块 + 页脚校名）。
- 「创新点」等学院自加章节规范里未规定，保持用户原有层级即可。
