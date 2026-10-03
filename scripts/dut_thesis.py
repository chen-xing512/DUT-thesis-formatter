# -*- coding: utf-8 -*-
"""dut_thesis.py — 大连理工大学硕士学位论文格式 排版工具库

把一份内容杂乱的中文 `.docx`（论文草稿 / 学院模板 / 拼凑的文档）**就地改写**为
符合《大连理工大学硕士学位论文格式规范》的成品；也支持从零组装新论文。

设计要点
--------
* **样式优先**：优先套用官方模板里的段落样式（`摘要题目` / `图名中文` / `参考文献正文` …），
  只有在样式不存在时才逐段硬设字体——这样产出的文档在 Word 里仍然「可编辑、可统一修改」。
* **中文字体必须写 `w:eastAsia`**：python-docx 的 `run.font.name` 只设置西文，
  这是 90% 的「中文没变成黑体」问题的根因。本库统一由 `set_run()` 处理。
* **缩进用字符单位**：`w:firstLineChars="200"`，不按磅硬算，中文才不会跑偏。
* **三线表 + 公式制表位 + TOC 域 + 奇偶页页眉 + 分节页码**：官方规范里最容易漏、
  手写最费 token 的样板代码，全部封装在下面的函数里。

依赖：`python-docx`（`pip install python-docx`）

用法速览
--------
    from dut_thesis import *

    # A) 改写既有文档
    doc = Document('我的论文.docx')
    report = reformat(doc)          # 一键套用全部 DUT 规则
    print_report(report)
    enable_update_fields(doc)       # 让 Word 打开时自动更新目录/页码
    doc.save('我的论文_DUT.docx')

    # B) 从零新建
    doc = new_doc()
    add_cover(doc, title_cn='...', title_en='...', fields=[('作者姓名', '张三'), ...])
    doc.save('out.docx')

    # C) 校验
    doc = Document('out.docx')
    print_report(audit(doc))
"""
from __future__ import annotations

import copy
import re
from typing import Iterable, Sequence

from docx import Document
from docx.document import Document as _Doc
from docx.enum.section import WD_SECTION, WD_SECTION_START
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph

__all__ = [
    # 常量
    'FONT_SIZE', 'CN_FONT', 'EN_FONT', 'HEI', 'SONG', 'PAGE_W_CM', 'TEXT_W_CM',
    # 底层
    'set_run', 'fmt_para', 'first_line_chars', 'set_outline', 'set_snap_to_grid',
    'set_keep_with_next', 'page_break_before',
    # 段落级
    'add_para', 'add_body', 'add_heading', 'add_caption', 'add_blank',
    'add_figure', 'add_equation', 'add_superscript_citation',
    # 表格
    'three_line_table', 'merge_cells_vertical', 'clear_cell',
    'add_field_table', 'set_table_borders', 'set_table_grid',
    # 结构
    'new_doc', 'ensure_style', 'apply_style_set', 'add_cover', 'add_page_field',
    'set_pgnum', 'set_header_pair', 'add_toc_field', 'add_list_of_field',
    'enable_update_fields', 'add_section',
    # 改写 / 校验
    'classify_paragraph', 'reformat', 'audit', 'print_report',
    # 段落分类常量
    'CH_LEVEL1', 'CH_LEVEL2', 'CH_LEVEL3', 'CH_FRONT', 'CH_CAPTION_FIG',
    'CH_CAPTION_TAB', 'CH_REF', 'CH_BODY', 'CH_BODY_EN', 'CH_KEYWORDS', 'CH_EMPTY',
    'is_latin_prose', 'strip_style_color',
    # 其它
    'clear_cell', 'add_plain_heading', 'set_page_number_footer', 'set_keep_with_next',
    'set_keep_lines', 'add_list_of_field', 'PT_TITLE', 'PT_ABSTRACT', 'PT_SECTION',
    'PT_BODY', 'PT_CAPTION', 'PT_PAGENUM', 'LINE_BODY', 'LINE_HEADING',
    # 文本
    'norm_ref_spacing', 'detect_ref_type', 'format_reference_entry', 'order_references',
    'find_reference_blocks', 'ensure_toc_field',
]


# --------------------------------------------------------------------------
# 常量
# --------------------------------------------------------------------------

#: 中文字号 → 磅值
FONT_SIZE = {
    '初号': 42, '小初': 36, '一号': 26, '小一': 24,
    '二号': 22, '小二': 18, '三号': 16, '小三': 15,
    '四号': 14, '小四': 12, '五号': 10.5, '小五': 9,
    '六号': 7.5, '小六': 6.5,
}

CN_FONT = '宋体'          # 中文正文
HEI = '黑体'               # 中文标题
SONG = '宋体'
EN_FONT = 'Times New Roman'   # 英文与数字

PT_TITLE = FONT_SIZE['三号']     # 16  章标题
PT_ABSTRACT = FONT_SIZE['小三']  # 15  摘要/目录/参考文献/附录/致谢 标题
PT_SECTION = FONT_SIZE['四号']   # 14  节标题
PT_BODY = FONT_SIZE['小四']      # 12  正文
PT_CAPTION = FONT_SIZE['五号']   # 10.5 图题/表题/表内/参考文献
PT_PAGENUM = FONT_SIZE['小五']   # 9   页码

LINE_BODY = 1.25        # 正文多倍行距
LINE_HEADING = 1.5      # 标题行距

#: 三级标题的**权威规格**——样式层、段落层、校验层都从这里取，避免三处各写一份而对不上。
#:
#: 取值依据是官方模板 `大连理工大学硕士学位论文格式规范.docx` 的样式定义
#: （直接读它的 ``w:rPr`` / ``w:pPr`` 得到），要点：
#:
#: * **不加粗**。规范文字只说「黑体」——那是**字体名**（SimHei），
#:   官方模板的 Heading 1/2/3 都**没有** ``w:b``（只有 ``w:bCs``，那是复杂文种用的）。
#:   黑体本身笔画已经够重，再加 ``w:b`` 会变成合成粗体，和模板不一致。
#: * ``sz`` 单位是半磅，所以三号 16pt → ``32``；Heading 3 官方**不写 sz**，
#:   直接继承 Normal 的 24（小四 12pt）。
#: * 行距 ``line=360 lineRule=auto``（360/240 = 1.5 倍），段前/段后用 ``*Lines``
#:   表达「0.5 行 / 1 行」的行单位。
HEADING_SPEC = {
    #     字号        加粗    段前(行)  段后(行)  行距
    1: dict(size=PT_TITLE,   bold=False, before_lines=0.0, after_lines=1.0, line=LINE_HEADING),
    2: dict(size=PT_SECTION, bold=False, before_lines=0.5, after_lines=0.0, line=LINE_HEADING),
    3: dict(size=PT_BODY,    bold=False, before_lines=0.5, after_lines=0.0, line=LINE_HEADING),
}

#: 版心宽度（cm）—— A4 21cm − 左右各 2.5cm
PAGE_W_CM = 21.0
TEXT_W_CM = 16.0

#: 官方模板自定义样式的 styleId → (名称, 说明)，用于 ensure_style
_OFFICIAL_STYLES = {
    'affe': '摘要题目', 'ABSTRACT0': 'ABSTRACT', 'affc': '关键词', 'affd': '关键词题头',
    'afff': '英文摘要正文', 'afff0': '英文关键词', 'Keywords': 'Keywords',
    'afff1': '目录', 'afff2': '图名中文', 'afff3': '图名英文', 'afff4': '参考文献标题',
    'afff5': '参考文献正文', 'afff6': '授权说明正文', 'afff7': '作者简介',
    'afff9': 'Caption', 'ad': '附录', 'ae': '致谢', '公式': '公式',
    'a5': 'Footer', 'a4': 'Header', 'a': 'Normal',
}


# --------------------------------------------------------------------------
# 底层：字体 / 段落 / 缩进
# --------------------------------------------------------------------------

def set_run(run, cn: str = SONG, en: str = EN_FONT, size: float = PT_BODY,
            bold: bool = False, italic: bool = False, underline: bool = False,
            color=None, superscript: bool = False):
    """设置 run 的字体——**含中文字体**（python-docx 不暴露的那部分）。

    cn 走 ``w:eastAsia``，en 走 ``w:ascii`` / ``w:hAnsi``。
    """
    run.font.name = en
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = underline
    # python-docx 里 font.bold=False 会**删除** <w:b>，而 Word 的语义是
    # 「标签存在即为真」——删掉就退回继承。所以显式写 val="0" 把字重钉死。
    _rpr = run._element.get_or_add_rPr()
    for _tag in ('w:b', 'w:i', 'w:u'):
        for _el in _rpr.findall(qn(_tag)):
            _rpr.remove(_el)
    for _tag, _flag in (('w:b', bold), ('w:i', italic)):
        _el = OxmlElement(_tag)
        _el.set(qn('w:val'), '1' if _flag else '0')
        _rpr.append(_el)
    _u = OxmlElement('w:u')
    _u.set(qn('w:val'), 'single' if underline else 'none')
    _rpr.append(_u)
    if superscript:
        run.font.superscript = True
    if color is not None:
        run.font.color.rgb = RGBColor(*color)
    rpr = run._element.get_or_add_rPr()
    rf = rpr.get_or_add_rFonts()
    rf.set(qn('w:ascii'), en)
    rf.set(qn('w:hAnsi'), en)
    rf.set(qn('w:eastAsia'), cn)
    rf.set(qn('w:cs'), en)
    return run


def fmt_para(p, align=None, line: float = LINE_BODY, before: float = 0,
             after: float = 0, exact: float | None = None,
             first_line_chars: int | None = None,
             left_indent_pt: float | None = None,
             hanging_pt: float | None = None,
             snap_to_grid: bool | None = None):
    """设置段落对齐 / 行距 / 段前后 / 缩进。行距单位是「倍」，``exact`` 为固定磅值。"""
    pf = p.paragraph_format
    if align is not None:
        pf.alignment = align
    if exact is not None:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(exact)
    else:
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.line_spacing = line
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if first_line_chars is not None:
        _set_ind(p, firstLineChars=first_line_chars * 100,
                 firstLine=first_line_chars * 240)
    if left_indent_pt is not None or hanging_pt is not None:
        kw = {}
        if left_indent_pt is not None:
            kw['left'] = int(left_indent_pt * 20)
        if hanging_pt is not None:
            kw['hanging'] = int(hanging_pt * 20)
            kw['hangingChars'] = int(hanging_pt * 20)
        _set_ind(p, **kw)
    if snap_to_grid is not None:
        set_snap_to_grid(p, snap_to_grid)
    return p


def _set_ind(p, **attrs):
    """写 ``w:ind``。属性名用 OOXML 原名（firstLineChars / hanging / left …），值为字符串。"""
    ppr = p._p.get_or_add_pPr()
    ind = ppr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        # w:ind 必须排在 pPr 的合法位置：插到 numPr/jc 之后、rPr 之前
        rpr = ppr.find(qn('w:rPr'))
        if rpr is not None:
            rpr.addprevious(ind)
        else:
            ppr.append(ind)
    for k, v in attrs.items():
        ind.set(qn('w:' + k), str(v))
    return ind


def first_line_chars(p, chars: int = 2):
    """首行缩进 N 个汉字字符（用 ``w:firstLineChars``，比按磅换算准）。"""
    _set_ind(p, firstLineChars=chars * 100, firstLine=chars * 240)
    return p


def set_snap_to_grid(p, on: bool = False):
    """取消/启用「对齐到网格」。正文必须关闭（``w:snapToGrid val="0"``），否则行距会被网格拉偏。"""
    ppr = p._p.get_or_add_pPr()
    el = ppr.find(qn('w:snapToGrid'))
    if el is None:
        el = OxmlElement('w:snapToGrid')
        ppr.insert(0, el)
    el.set(qn('w:val'), '1' if on else '0')
    return p


def set_outline(p, level: int):
    """outline level：0=章 1=节 2=小节，9=正文级（排除出目录）。"""
    ppr = p._p.get_or_add_pPr()
    for old in ppr.findall(qn('w:outlineLvl')):
        ppr.remove(old)
    ol = OxmlElement('w:outlineLvl')
    ol.set(qn('w:val'), str(level))
    ppr.append(ol)
    return p


def set_keep_with_next(p, on: bool = True):
    """标题与图表题：与下段同页。"""
    ppr = p._p.get_or_add_pPr()
    tag = 'w:keepNext'
    for old in ppr.findall(qn(tag)):
        ppr.remove(old)
    if on:
        ppr.append(OxmlElement(tag))
    return p


def set_keep_lines(p, on: bool = True):
    """图表题：段中不分页。"""
    ppr = p._p.get_or_add_pPr()
    tag = 'w:keepLines'
    for old in ppr.findall(qn(tag)):
        ppr.remove(old)
    if on:
        ppr.append(OxmlElement(tag))
    return p


def set_page_break_before(p, on: bool = True):
    """段前分页——「每章另起一页」用这个，比插空行稳。"""
    ppr = p._p.get_or_add_pPr()
    for old in ppr.findall(qn('w:pageBreakBefore')):
        ppr.remove(old)
    if on:
        el = OxmlElement('w:pageBreakBefore')
        ppr.append(el)
    return p


# 向后兼容别名
page_break_before = set_page_break_before


# --------------------------------------------------------------------------
# 段落级构件
# --------------------------------------------------------------------------

def add_para(doc, text: str = '', cn: str = SONG, en: str = EN_FONT,
             size: float = PT_BODY, bold: bool = False, align=None,
             line: float = LINE_BODY, before: float = 0, after: float = 0,
             indent_chars: int | None = None, snap_to_grid: bool | None = None,
             style: str | None = None):
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    if text:
        set_run(p.add_run(text), cn=cn, en=en, size=size, bold=bold)
    fmt_para(p, align=align, line=line, before=before, after=after,
             snap_to_grid=snap_to_grid)
    if indent_chars is not None:
        first_line_chars(p, indent_chars)
    return p


def add_body(doc, text: str, before: float = 0, after: float = 0,
             bold_lead: str | None = None, align=None, indent: bool = True):
    """正文段：宋体小四，两端对齐，1.25 倍行距，首行缩进 2 字符，取消网格对齐。"""
    p = doc.add_paragraph(style='Normal')
    if bold_lead:
        set_run(p.add_run(bold_lead), cn=HEI, en=EN_FONT, size=PT_BODY, bold=True)
    set_run(p.add_run(text), cn=SONG, en=EN_FONT, size=PT_BODY)
    fmt_para(p, align=align or WD_ALIGN_PARAGRAPH.JUSTIFY, line=LINE_BODY,
             before=before, after=after, snap_to_grid=False)
    if indent:
        first_line_chars(p, 2)
    else:
        _set_ind(p, firstLineChars=0, firstLine=0)
    return p


def add_heading(doc, text: str, level: int, toc: bool = True, page_break: bool = False):
    """章 / 节 / 小节标题。格式取自 :data:`HEADING_SPEC`（依据官方模板）：

    * level 1 → 黑体三号 **不加粗** 居左，段后 1 行，outline 0（章）
    * level 2 → 黑体四号 **不加粗** 居左，段前 0.5 行，outline 1（节）
    * level 3 → 黑体小四 **不加粗** 居左，段前 0.5 行，outline 2（小节）
    * ``toc=False`` → outline 9（不进目录）

    注意「黑体」是**字体名**（SimHei），规范并未要求再加粗；
    官方模板的 Heading 1/2/3 都没有 ``w:b``。多加深粗会变成合成粗体，与模板不一致。
    """
    if level not in (1, 2, 3):
        raise ValueError('level 必须是 1/2/3')
    spec = HEADING_SPEC[level]
    p = doc.add_paragraph(style=f'Heading {level}')
    p.add_run(text)
    set_keep_with_next(p, True)
    if page_break:
        set_page_break_before(p, True)
    # 间距/对齐/缩进/大纲交给样式，段落层不再覆盖；并校验样式本身
    enforce_heading_format(doc, level, p)
    if not toc:
        set_outline(p, 9)      # 目录标题不进目录，这是样式里没有的意图
    return p


def add_plain_heading(doc, text: str, size: float = PT_ABSTRACT, cn: str = HEI,
                      en: str = EN_FONT, after: float = 12, center: bool = True,
                      toc: bool = False, style: str | None = None):
    """居中章级标题（摘要 / ABSTRACT / 目录 / 参考文献 / 附录 / 致谢 / 攻读学位期间成果）。

    规范：小三 15pt 黑体加粗居中，1.5 倍行距，段后 1 行。
    ``toc=True`` 时设 outline 0，使其进入目录（参考文献 / 附录 / 致谢需要）。
    """
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    set_run(p.add_run(text), cn=cn, en=en, size=size, bold=True)
    fmt_para(p, align=WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT,
             line=LINE_HEADING, before=0, after=after, snap_to_grid=False)
    _set_ind(p, left=0, firstLineChars=0, firstLine=0)
    set_outline(p, 0 if toc else 9)
    set_keep_with_next(p, True)
    return p


def add_caption(doc, text_cn: str, text_en: str | None = None,
                style_cn: str = 'afff2', style_en: str = 'afff3', size: float = PT_CAPTION):
    """图题 / 表题：中文一行（宋体五号居中）+ 英文一行（TNR 五号居中）。

    规范要求中英文**两行**，编号与文字间留 1 个空格（如 ``图 1.1 系统结构图``）。
    """
    p1 = doc.add_paragraph()
    set_run(p1.add_run(text_cn), cn=SONG, en=EN_FONT, size=size)
    fmt_para(p1, align=WD_ALIGN_PARAGRAPH.CENTER, line=1.0, before=3, after=0,
             snap_to_grid=False)
    _set_ind(p1, firstLineChars=0, firstLine=0)
    set_keep_lines(p1, True)
    if text_en:
        p2 = doc.add_paragraph()
        set_run(p2.add_run(text_en), cn=EN_FONT, en=EN_FONT, size=size)
        fmt_para(p2, align=WD_ALIGN_PARAGRAPH.CENTER, line=1.0, before=0, after=3,
                 snap_to_grid=False)
        _set_ind(p2, firstLineChars=0, firstLine=0)
        set_keep_lines(p2, True)
    return p1


def add_blank(doc, size: float = PT_BODY):
    """空行——固定行距 25 磅，避免被 1.25 倍行距放大。"""
    p = doc.add_paragraph()
    fmt_para(p, line=1.0, exact=25, before=0, after=0, snap_to_grid=False)
    _set_ind(p, firstLineChars=0, firstLine=0)
    return p


def add_figure(doc, path, caption_cn: str, caption_en: str | None = None,
               width_cm: float = 12.0, before: bool = True, after: bool = False):
    """插图 + 图题（图题在图**下方**）。

    规范：图居中、与下文留一空行、嵌入型版式、单图不超过一页。
    ``width_cm`` 默认 12，上限 ``TEXT_W_CM``（16.0）。
    """
    if width_cm > TEXT_W_CM:
        raise ValueError(f'图宽 {width_cm}cm 超过版心宽 {TEXT_W_CM}cm')
    if before:
        add_blank(doc)
    p = doc.add_paragraph()
    fmt_para(p, align=WD_ALIGN_PARAGRAPH.CENTER, line=1.0, before=0, after=0,
             snap_to_grid=False)
    _set_ind(p, firstLineChars=0, firstLine=0)
    set_keep_with_next(p, True)
    p.add_run().add_picture(path, width=Cm(width_cm))
    add_caption(doc, caption_cn, caption_en)
    if after:
        add_blank(doc)
    return p


def add_equation(doc, omml_xml: str | None = None, number: str = '',
                 text: str = '', text_width_cm: float = TEXT_W_CM):
    """公式行：公式**居中**、编号**行末右对齐**，段前段后 6 磅。

    用「居中制表位 + 右对齐制表位」实现（等价于官方模板的公式样式做法）。

    :param omml_xml: 可选，`<m:oMath>…</m:oMath>` 的 XML 字符串（Word 原生公式）。
    :param text: 纯文本公式内容（没有 OMML 时的降级方案，建议用 Cambria Math）。
    :param number: 编号文本，如 ``(3.1)``。
    """
    p = doc.add_paragraph()
    fmt_para(p, align=WD_ALIGN_PARAGRAPH.LEFT, line=1.0, before=6, after=6,
             snap_to_grid=False)
    _set_ind(p, firstLineChars=0, firstLine=0, left=0)
    pf = p.paragraph_format
    pf.tab_stops.add_tab_stop(Cm(text_width_cm / 2.0), WD_TAB_ALIGNMENT.CENTER)
    pf.tab_stops.add_tab_stop(Cm(text_width_cm), WD_TAB_ALIGNMENT.RIGHT)
    p.add_run('\t')
    if omml_xml:
        from lxml import etree
        p._p.append(etree.fromstring(omml_xml))
    elif text:
        set_run(p.add_run(text), cn='Cambria Math', en='Cambria Math',
                size=PT_BODY, italic=False)
    if number:
        p.add_run('\t')
        set_run(p.add_run(number), cn=SONG, en=EN_FONT, size=PT_BODY)
    return p


_CITE_RE = re.compile(r'(\[[0-9]+(?:[,\-–][0-9]+)*\])')


def add_superscript_citation(p, text: str):
    """把段内 ``[1]`` / ``[2-4]`` 这类引用标记转成**右上角标**。

    规范：正文引用处用方括号 + 阿拉伯数字，**右上角标**标注。
    """
    for chunk in _CITE_RE.split(text):
        if not chunk:
            continue
        if _CITE_RE.fullmatch(chunk):
            set_run(p.add_run(chunk), cn=SONG, en=EN_FONT, size=PT_BODY,
                    superscript=True)
        else:
            set_run(p.add_run(chunk), cn=SONG, en=EN_FONT, size=PT_BODY)
    return p


# --------------------------------------------------------------------------
# 三线表
# --------------------------------------------------------------------------

def _set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
    """给单元格设边框，``sz`` 单位是 1/8 磅（1.5pt → 12，1pt → 8，0.5pt → 4）。"""
    tcPr = cell._tc.get_or_add_tcPr()
    tb = tcPr.find(qn('w:tcBorders'))
    if tb is None:
        tb = OxmlElement('w:tcBorders')
        tcPr.append(tb)
    for edge, val in (('top', top), ('bottom', bottom), ('left', left), ('right', right)):
        if val is None:
            continue
        el = tb.find(qn('w:' + edge))
        if el is None:
            el = OxmlElement('w:' + edge)
            tb.append(el)
        if val == 0:
            el.set(qn('w:val'), 'none')
        else:
            el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), str(val))
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), '000000')


def set_table_grid(table, widths_cm: Sequence[float]):
    """按 cm 设置表格的**固定列宽**（tblW + tblGrid + 每格 tcW）。

    只设 ``cell.width`` 是不够的：python-docx 建的表默认等宽网格，
    Word 以 ``w:tblGrid`` 为准，于是列宽不按你写的走——
    标签列被压窄后 ``w:jc="distribute"`` 会把文字挤到换行。
    """
    total = int(round(sum(widths_cm) * 567))
    tblPr = table._tbl.tblPr
    for tag in ('w:tblW', 'w:tblLayout'):
        for old in tblPr.findall(qn(tag)):
            tblPr.remove(old)
    tw = OxmlElement('w:tblW')
    tw.set(qn('w:w'), str(total))
    tw.set(qn('w:type'), 'dxa')
    tblPr.append(tw)
    lay = OxmlElement('w:tblLayout')
    lay.set(qn('w:type'), 'fixed')
    tblPr.append(lay)

    grid = table._tbl.find(qn('w:tblGrid'))
    if grid is not None:
        table._tbl.remove(grid)
    grid = OxmlElement('w:tblGrid')
    for w in widths_cm:
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), str(int(round(w * 567))))
        grid.append(gc)
    # tblGrid 必须排在 tblPr 之后
    tblPr.addnext(grid)

    for row in table.rows:
        for i, cell in enumerate(row.cells):
            if i >= len(widths_cm):
                break
            cell.width = Cm(widths_cm[i])
    return table


def set_table_borders(table, spec: dict | None = None, none: bool = False):
    """设置**表级**边框。``none=True`` 时清成无框（隐藏边框表格）。

    ``spec`` 形如 ``{'top': 12, 'bottom': 12, 'insideH': 0}``，值单位是 1/8 磅；
    0 表示该边为无框。
    """
    tblPr = table._tbl.tblPr
    for old in tblPr.findall(qn('w:tblBorders')):
        tblPr.remove(old)
    borders = OxmlElement('w:tblBorders')
    wanted = {k: 0 for k in ('top', 'left', 'bottom', 'right',
                             'insideH', 'insideV')} if none else dict(spec or {})
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement('w:' + edge)
        sz = wanted.get(edge, 0)
        if sz:
            el.set(qn('w:val'), 'single')
            el.set(qn('w:sz'), str(sz))
        else:
            el.set(qn('w:val'), 'none')
            el.set(qn('w:sz'), '0')
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), 'auto')
        borders.append(el)
    tblPr.append(borders)
    return table


def add_field_table(doc, rows, label_w_cm: float = 2.7, colon_w_cm: float = 0.6,
                    value_w_cm: float = 7.4, seps: Sequence[str] = ('、',),
                    size: float = None, underline_values: bool = True,
                    label_cn: str = SONG, value_cn: str = SONG,
                    value_align=None, line_height_pt: float = 16.0,
                    line_gap_pt: float = 6.0):
    """**隐藏边框的三列表格**——封面 / 报告首页的字段块：``标签 | 冒号 | 值``。

    三列各自独立成**单元格**，所以打开 Word 的"显示框线"或给表格加边框时，
    能清楚看到是三格；如果不这样拆，冒号会和标签同格（视觉上分不开）。

    :param label_w_cm: 标签列宽。**这一列决定分散对齐的字距**，是调外观的关键：
        分散对齐会把标签撑满整列，列越宽字距越大。

        * 2.7 cm → 字距紧凑（≈2 pt），接近"正常书写"的样子；
        * 3.6 cm → 字距明显拉开（≈23 pt），像手工排的封面。

        实测各标签天然宽度（15 pt）：``学号`` 2 字 1.06 cm、
        ``指导教师``/``答辩日期`` 4 字 2.12 cm、``学科、专业`` 5 字 2.65 cm。
        所以 2.7 cm 刚好略宽于最长标签，字距最小。
    :param colon_w_cm: 冒号列宽（全角冒号约 0.53 cm，故 0.6 cm 够）。
    :param value_w_cm: 值列宽，**决定填写线长度**（线 = 值单元格下边框）。
        须不小于最长值宽度，否则折行（15 pt「信息与通信工程」8 字 ≈ 4.23 cm）。
    :param seps: 需要补分隔符的标签，默认 ``('、',)`` → 「学科、专业」。
    :param value_align: 值的对齐，默认居中。
    :param line_height_pt: 固定行高（磅）。
    :param line_gap_pt: 固定行高之上再追加的段后空隙（磅），微调线与字的距离。
    """
    size = size if size is not None else FONT_SIZE['小三']
    value_align = value_align or WD_ALIGN_PARAGRAPH.CENTER

    t = doc.add_table(rows=0, cols=3)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    set_table_borders(t, none=True)          # 隐藏边框
    tblPr = t._tbl.tblPr
    mar = OxmlElement('w:tblCellMar')
    # bottom 必须为 0，否则下边框会被边距推离文字；左右留白要小，
    # 否则冒号列两边会各多出 0.1 cm，把标签与冒号推开
    for side, w in (('left', 10), ('right', 10), ('top', 20), ('bottom', 0)):
        el = OxmlElement('w:' + side)
        el.set(qn('w:w'), str(w))
        el.set(qn('w:type'), 'dxa')
        mar.append(el)
    tblPr.append(mar)

    def _prep(cell, align, cn, txt, after=0.0):
        clear_cell(cell)
        para = cell.paragraphs[0]
        fmt_para(para, align=align, line=1.0, before=0, after=after,
                 snap_to_grid=False)
        # 固定行高——否则 Word 撑高行框，下边框会掉到下一行旁边
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        para.paragraph_format.line_spacing = Pt(line_height_pt)
        _set_ind(para, firstLineChars=0, firstLine=0, left=0)
        if txt:
            set_run(para.add_run(txt), cn=cn, en=EN_FONT, size=size)
        return para

    for label, value in rows:
        r = t.add_row()
        # ① 标签列：分散对齐撑满本列（不含冒号，拉伸范围纯粹）
        _prep(r.cells[0], WD_ALIGN_PARAGRAPH.DISTRIBUTE, label_cn,
              _label_with_sep(label, seps))
        # ② 冒号列：**独立单元格**。右对齐，紧贴值列
        _prep(r.cells[1], WD_ALIGN_PARAGRAPH.RIGHT, label_cn, '：')
        # ③ 值列：内容 + 该单元格**自己的下边框**当填写线
        _prep(r.cells[2], value_align, value_cn, f'{value}',
              after=(line_gap_pt if underline_values else 0))
        if underline_values:
            _set_cell_borders(r.cells[2], top=0, left=0, right=0, bottom=4)
        else:
            _set_cell_borders(r.cells[2], top=0, left=0, right=0, bottom=0)
        for c in (r.cells[0], r.cells[1]):
            _set_cell_borders(c, top=0, left=0, right=0, bottom=0)

    set_table_grid(t, [label_w_cm, colon_w_cm, value_w_cm])
    return t


def _label_with_sep(label: str, seps: Sequence[str]):
    """给需要分隔符的标签补上分隔符（「学科、专业」「答辩日期」）。

    标签**不该**由调用方自己在中间塞空格——那是旧写法。
    """
    lab = (label or '').strip()
    for sep in seps:
        plain = sep.replace('　', '')
        if plain and plain in lab:
            return sep.join(lab.split(plain))
    return lab


def clear_cell(cell):
    """清空单元格：删除多余段落、清掉所有 run。

    不要用 ``cell.text = ''``——那会留下一个**空 run**，把后面的字号设置读乱。
    """
    for p in list(cell.paragraphs[1:]):
        p._p.getparent().remove(p._p)
    p = cell.paragraphs[0]
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    return p


def _clear_table_borders(table):
    tblPr = table._tbl.tblPr
    for old in tblPr.findall(qn('w:tblBorders')):
        tblPr.remove(old)
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement('w:' + edge)
        el.set(qn('w:val'), 'none')
        el.set(qn('w:sz'), '0')
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), 'auto')
        borders.append(el)
    tblPr.append(borders)


def three_line_table(doc, header: Sequence[str], rows: Iterable[Sequence],
                     col_widths: Sequence[float] | None = None,
                     top_sz: int = 12, header_sz: int = 8, bottom_sz: int = 12,
                     font_size: float = PT_CAPTION, header_cn: str = SONG,
                     header_bold: bool = True, align_center: bool = True):
    """DUT 三线表（**无竖线**）。

    * 上线 / 下线：1.5 磅 → ``w:sz="12"``
    * 表头下线：1 磅 → ``w:sz="8"``
    * 表内：中文宋体 / 英文 Times New Roman，五号，居中，1.25 倍行距，无首行缩进

    :param col_widths: 各列宽度（cm）。
    :param header_sz: 表头下线粗细（1/8 磅单位）。改 0.5 磅传 4。
    """
    rows = list(rows)
    ncol = len(header)
    t = doc.add_table(rows=1 + len(rows), cols=ncol)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = True
    _clear_table_borders(t)

    def _fill(cell, val, bold=False, cn=SONG):
        clear_cell(cell)
        p = cell.paragraphs[0]
        fmt_para(p, align=WD_ALIGN_PARAGRAPH.CENTER if align_center else WD_ALIGN_PARAGRAPH.LEFT,
                 line=1.25, before=1, after=1, snap_to_grid=False)
        _set_ind(p, firstLineChars=0, firstLine=0)
        set_run(p.add_run(str(val)), cn=cn, en=EN_FONT, size=font_size, bold=bold)

    for j, h in enumerate(header):
        c = t.rows[0].cells[j]
        _fill(c, h, bold=header_bold, cn=header_cn)
        _set_cell_borders(c, top=top_sz, bottom=header_sz, left=0, right=0)

    for i, row in enumerate(rows):
        last = (i == len(rows) - 1)
        for j, val in enumerate(row):
            c = t.rows[i + 1].cells[j]
            _fill(c, val)
            _set_cell_borders(c, left=0, right=0,
                              bottom=bottom_sz if last else 0)

    if col_widths:
        for j, w in enumerate(col_widths):
            for row in t.rows:
                row.cells[j].width = Cm(w)
    return t


def merge_cells_vertical(table, col: int, r0: int, r1: int, keep: str | None = None):
    """纵向合并 ``col`` 列第 r0..r1 行（含）。默认保留 r0 内容。"""
    c0 = table.rows[r0].cells[col]
    if keep is not None:
        clear_cell(c0)
        c0.paragraphs[0].add_run(keep)
    tcPr0 = c0._tc.get_or_add_tcPr()
    vm0 = OxmlElement('w:vMerge')
    vm0.set(qn('w:val'), 'restart')
    tcPr0.append(vm0)
    for r in range(r0 + 1, r1 + 1):
        c = table.rows[r].cells[col]
        clear_cell(c)
        tcPr = c._tc.get_or_add_tcPr()
        tcPr.append(OxmlElement('w:vMerge'))
    return table


# --------------------------------------------------------------------------
# 文档级：样式、页面、分节、页眉页脚
# --------------------------------------------------------------------------

def strip_style_color(style, rgb=(0, 0, 0)):
    """把段落样式里的字体颜色**钉成黑色**（不设置颜色会继承主题的蓝色）。

    为什么必须显式处理：python-docx 内置模板的 Heading 1/2/3 自带
    ``w:color val="365F91"/"4F81BD"``（主题 accent 蓝）。DUT 规范要求各级标题
    是**黑体**——指黑体字（SimHei）且**颜色为黑**。不清掉这个蓝色，
    生成的论文在 Word 里标题就是蓝的，打印出来更是明显不对。

    官方模板的做法是**根本不写 w:color**（从而继承正文的自动黑）。
    这里两者兼顾：先删掉继承来的颜色定义，再显式写 000000——
    显式写比依赖继承更稳，也不会被主题色影响。
    """
    rpr = style.element.get_or_add_rPr()
    for old in rpr.findall(qn('w:color')):
        rpr.remove(old)
    el = OxmlElement('w:color')
    el.set(qn('w:val'), '%02X%02X%02X' % rgb)
    rpr.append(el)
    return style


def _style_rpr(style):
    return style.element.get_or_add_rPr()


#: 会**盖过样式**的字符级属性。命名样式只是继承基线，run 上只要有这些就直接生效。
_OVERRIDE_RPR_TAGS = (
    'w:sz', 'w:szCs',            # 字号：最大嫌疑
    'w:b', 'w:bCs', 'w:i', 'w:iCs',
    'w:rFonts', 'w:color',       # 字体 / 颜色
    'w:spacing', 'w:w', 'w:kern', 'w:position',   # 字距 / 缩放 / 字偶距 / 上下移
    'w:u', 'w:strike', 'w:dstrike', 'w:shd', 'w:highlight', 'w:emboss',
    'w:imprint', 'w:outline', 'w:vanish', 'w:smallCaps', 'w:caps', 'w:effect',
)

#: 上面那些里**允许**留在 run 上的。``w:rFonts`` 必须逐个 run 写：
#: 中西文分属两个字体槽（``w:eastAsia`` / ``w:ascii``），样式层给了也可能被
#: run 的旧字体覆盖，而它不属于「会改变字号字重的坏覆盖」，故白名单放行。
_OVERRIDE_ALLOWED = ('w:rFonts',)


def strip_char_overrides(para, keep_bold: bool | None = None):
    """剥掉段落内所有 run 的**字符级直接格式**，让样式成为唯一来源。

    命名样式只是「继承基线」：run 上任何直接格式都优先于样式。用户从别处
    粘一段标题进来、或手动加粗/改字号，样式就形同虚设。所以套样式之后
    必须把这些覆盖清掉，渲染才真正听规范的。

    :param keep_bold: ``None`` 时保留现有粗体设置（只清字号字体等）；
        给 ``True``/``False`` 则连 ``w:b`` 一起定型。
    """
    for r in para.runs:
        rpr = r._element.find(qn('w:rPr'))
        if rpr is None:
            continue
        for tag in _OVERRIDE_RPR_TAGS:
            if tag in ('w:b', 'w:bCs') and keep_bold is not None:
                continue
            for el in rpr.findall(qn(tag)):
                rpr.remove(el)
        if keep_bold is not None:
            for tag in ('w:b', 'w:bCs'):
                for el in rpr.findall(qn(tag)):
                    rpr.remove(el)
                if keep_bold:
                    rpr.append(OxmlElement(tag))
    return para


def _effective_char_fmt(para):
    """返回段落首个 run 的**有效**字符格式 ``(size_pt, bold, eastAsia, color)``。

    先看 run 的直接格式，没有则回落到它的段落样式，再看样式的 basedOn 链。
    """
    def _from_rpr(rpr):
        if rpr is None:
            return {}
        out = {}
        sz = rpr.find(qn('w:sz'))
        if sz is not None:
            try:
                out['size'] = int(sz.get(qn('w:val'))) / 2.0
            except (TypeError, ValueError):
                pass
        _b = rpr.find(qn('w:b'))
        if _b is not None:
            _v = (_b.get(qn('w:val')) or '1').lower()
            out['bold'] = _v not in ('0', 'false', 'off')
        # 没有 w:b 时不写 bold，留给继承链决定
        rf = rpr.find(qn('w:rFonts'))
        if rf is not None and rf.get(qn('w:eastAsia')):
            out['eastAsia'] = rf.get(qn('w:eastAsia'))
        c = rpr.find(qn('w:color'))
        if c is not None:
            out['color'] = (c.get(qn('w:val')) or '').upper()
        return out

    run = para.runs[0] if para.runs else None
    direct = _from_rpr(run._element.find(qn('w:rPr'))) if run is not None else {}

    st = para.style
    chain = []
    while st is not None:
        chain.append(st)
        try:
            st = st.base_style
        except Exception:
            st = None
        if len(chain) > 10:
            break
    merged = {}
    for style in reversed(chain):          # 从最根上的 Normal 往下盖
        merged.update(_from_rpr(style.element.find(qn('w:rPr'))))
    merged.update(direct)                  # 直接格式优先级最高
    return (merged.get('size'), merged.get('bold'),
            merged.get('eastAsia'), merged.get('color'))


def enforce_heading_format(doc, level: int, para=None):
    """确保标题**渲染出来**就是规范的样子，而不只是「套了样式」。

    三步走，对应三种会被用户改坏的地方：

    1. **样式层**——调 :func:`ensure_heading_style`，纠正 Word 样式面板里
       被改过的字号/加粗/间距（改样式会让所有同类标题一起变形）；
    2. **字符层**——:func:`strip_char_overrides` 清掉 run 上的直接格式
       （字号、加粗、字体、字距……），并把粗体按规范定型；
    3. **段落层**——删掉 ``w:spacing`` / ``w:jc`` / ``w:ind`` / ``w:snapToGrid``
       / ``w:outlineLvl`` 的段落级覆盖，让间距对齐回到样式；
       只保留 ``w:pageBreakBefore``（章前分页）这类样式里没有的意图。

    ``para=None`` 时只修样式（用于建文档阶段），随后新建的段落自然会正确。
    """
    spec = HEADING_SPEC[level]
    ensure_heading_style(doc, level)

    if para is None:
        return None

    # --- 样式套用（若当前不是该 Heading 样式，先套上）---
    want_style = f'Heading {level}'
    if para.style.name != want_style:
        try:
            para.style = doc.styles[want_style]
        except KeyError:
            pass
    # --- 字符层：**清掉一切直接格式**，只把中/西文字体名写回去 ---
    #
    # 为什么字号/字重/颜色不写在 run 上：三者样式层已经有了（见 ensure_heading_style），
    # run 上再写一份就是「直接格式压过样式」——用户之后在 Word 样式面板里改字号
    # 会发现改不动，因为 run 那份更优先。让样式成为唯一来源，才既符合规范又受管制。
    strip_char_overrides(para, keep_bold=bool(spec['bold']))
    for r in para.runs:
        _rpr = r._element.get_or_add_rPr()
        for _tag in ('w:sz', 'w:szCs', 'w:b', 'w:bCs', 'w:i', 'w:iCs',
                     'w:color', 'w:u', 'w:spacing', 'w:w', 'w:kern', 'w:position'):
            for _el in _rpr.findall(qn(_tag)):
                _rpr.remove(_el)
        _rf = _rpr.get_or_add_rFonts()
        _rf.set(qn('w:ascii'), EN_FONT)
        _rf.set(qn('w:hAnsi'), EN_FONT)
        _rf.set(qn('w:eastAsia'), HEI)
        _rf.set(qn('w:cs'), EN_FONT)
    # --- 段落层：让间距/对齐/缩进回到样式 ---
    ppr = para._p.get_or_add_pPr()
    for tag in ('w:spacing', 'w:jc', 'w:ind', 'w:snapToGrid', 'w:outlineLvl'):
        for el in ppr.findall(qn(tag)):
            ppr.remove(el)
    return para


def heading_format_problems(doc) -> list:
    """检查所有标题的**有效格式**是否符合 :data:`HEADING_SPEC`。

    只看结果、不看过程：无论问题出在样式被改、还是 run 上有覆盖，
    只要最终字号/加粗/字体/颜色/对齐/大纲级别不对，就报出来。
    """
    problems = []
    for i, para in enumerate(doc.paragraphs):
        sn = para.style.name if para.style is not None else ''
        if not (sn.startswith('Heading') and sn.split()[-1].isdigit()):
            continue
        level = int(sn.split()[-1])
        if level not in HEADING_SPEC or not para.text.strip():
            continue
        spec = HEADING_SPEC[level]
        size, bold, ea, color = _effective_char_fmt(para)
        label = f'第{i}段 {para.text[:14]!r}'
        if size is None or abs(size - spec['size']) > 0.01:
            problems.append(f'{label} 字号 {size}，应为 {spec["size"]}pt')
        if bool(bold) != bool(spec['bold']):
            problems.append(f'{label} 加粗 {bold}，应为 {spec["bold"]}（规范只要求黑体字体，不加粗）')
        if ea != HEI:
            problems.append(f'{label} 中文字体 {ea!r}，应为 {HEI!r}')
        if color not in (None, '000000', 'AUTO'):
            problems.append(f'{label} 颜色 #{color}，应为黑色')
        ppr = para._p.find(qn('w:pPr'))
        if ppr is not None:
            jc = ppr.find(qn('w:jc'))
            if jc is not None and jc.get(qn('w:val')) != 'left':
                problems.append(f'{label} 段落对齐 {jc.get(qn("w:val"))}，应为 left')
            ol = ppr.find(qn('w:outlineLvl'))
            if ol is not None and ol.get(qn('w:val')) != str(level - 1):
                problems.append(f'{label} 大纲级别 {ol.get(qn("w:val"))}，应为 {level - 1}')
        # 段落级的字符覆盖（样式之外的直接格式）
        for r in para.runs:
            rpr = r._element.find(qn('w:rPr'))
            if rpr is None:
                continue
            leftover = [t.split(':')[1] for t in _OVERRIDE_RPR_TAGS
                        if t not in _OVERRIDE_ALLOWED and rpr.find(qn(t)) is not None]
            if leftover:
                problems.append(f'{label} run 上残留直接格式 {leftover}（会盖过样式）')
                break
    return problems


def ensure_heading_style(doc, level: int):
    """把 ``Heading <level>`` 样式**修正**到 :data:`HEADING_SPEC`，返回是否改动过。

    为什么不能只在建文档时设一次：Word 的「样式」面板里，用户（或从别处粘进来的
    内容）随时能改掉样式本身的字号/加粗/间距；一旦样式被改，
    **所有套用该样式的标题会一起变形**，而段落层看不出来。

    实现上**一次性把规范值全部写全**，不做「按需局部修正」——
    因为 ``ensure_style`` 形如 ``size=PT_BODY, line=LINE_BODY`` 都带默认值，
    只传其中一两个参数去「局部修」时，**没传的字段会被默默重置成默认值**。
    曾经因此出现「修字体时把刚修好的字号打回 12pt」的连环 bug。
    """
    spec = HEADING_SPEC[level]
    name = f'Heading {level}'
    st = doc.styles[name]

    # 先确保它是个段落样式（名字可能被同名 Character 样式占着）
    if st.type != WD_STYLE_TYPE.PARAGRAPH:
        st.name = f'{name} (字符样式)'
        st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)

    # 基础属性一次性写全
    ensure_style(doc, name, based_on='Normal', cn=HEI, en=EN_FONT,
                 size=spec['size'], bold=spec['bold'], align=WD_ALIGN_PARAGRAPH.LEFT,
                 line=spec['line'], line_rule='auto',
                 before_lines=spec['before_lines'], after_lines=spec['after_lines'],
                 first_line_chars_=0, outline=level - 1)

    # ensure_style 只会写 w:b=True，不会主动删；这里按规范显式定型字重
    rpr = st.element.get_or_add_rPr()
    for tag in ('w:b', 'w:bCs', 'w:i', 'w:iCs'):
        for el in rpr.findall(qn(tag)):
            rpr.remove(el)
    if spec['bold']:
        rpr.append(OxmlElement('w:b'))
    # 颜色钉成黑色（python-docx 内置模板的标题自带主题蓝）
    strip_style_color(st, (0, 0, 0))

    # 章标题前的分页是**样式里没有的意图**，不能由样式给（否则所有 Heading 1
    # 都强制分页，附录/参考文献那种居中标题也会被带上）。这里清掉任何残留。
    ppr = st.element.get_or_add_pPr()
    for el in ppr.findall(qn('w:pageBreakBefore')):
        ppr.remove(el)
    return changed_style(st, level)


def _style_signature(st, level: int):
    """取样式的「规格指纹」，用于判断是否需要修正。"""
    spec = HEADING_SPEC[level]
    rpr = st.element.find(qn('w:rPr'))
    ppr = st.element.find(qn('w:pPr'))
    def _f(parent, tag):
        return None if parent is None else parent.find(qn(tag))

    rf, sz, b, c = _f(rpr, 'w:rFonts'), _f(rpr, 'w:sz'), _f(rpr, 'w:b'), _f(rpr, 'w:color')
    sp, ol, jc = _f(ppr, 'w:spacing'), _f(ppr, 'w:outlineLvl'), _f(ppr, 'w:jc')
    # 字重归一化：没有 w:b（继承）与 w:b val="0"（显式关闭）在「是否加粗」上等价，
    # 都算不加粗。只有真的要求加粗时才算 True——否则「缺省」会被误报成不符。
    bold_val = False
    if b is not None:
        bold_val = (b.get(qn('w:val')) or '1').lower() not in ('0', 'false', 'off')
    return (
        int(sz.get(qn('w:val'))) / 2.0 if sz is not None else None,
        bold_val,
        rf.get(qn('w:eastAsia')) if rf is not None else None,
        (c.get(qn('w:val')) or '').upper() if c is not None else None,
        sp.get(qn('w:line')) if sp is not None else None,
        sp.get(qn('w:beforeLines')) if sp is not None else None,
        sp.get(qn('w:afterLines')) if sp is not None else None,
        jc.get(qn('w:val')) if jc is not None else None,
        ol.get(qn('w:val')) if ol is not None else None,
    )


def _wanted_signature(level: int):
    spec = HEADING_SPEC[level]
    return (
        float(spec['size']),
        bool(spec['bold']),
        HEI,
        '000000',
        str(int(spec['line'] * 240)),
        str(int(spec['before_lines'] * 100)),
        str(int(spec['after_lines'] * 100)),
        'left',
        str(level - 1),
    )


def changed_style(st, level: int) -> bool:
    """样式是否与规范不符（不符就该修）。"""
    try:
        return _style_signature(st, level) != _wanted_signature(level)
    except Exception:
        return True


#: 会**盖过样式**的字符级属性。命名样式只是继承基线，run 上只要有这些就直接生效。
_OVERRIDE_RPR_TAGS = (
    'w:sz', 'w:szCs',            # 字号：最大嫌疑
    'w:b', 'w:bCs', 'w:i', 'w:iCs',
    'w:rFonts', 'w:color',       # 字体 / 颜色
    'w:spacing', 'w:w', 'w:kern', 'w:position',   # 字距 / 缩放 / 字偶距 / 上下移
    'w:u', 'w:strike', 'w:dstrike', 'w:shd', 'w:highlight', 'w:emboss',
    'w:imprint', 'w:outline', 'w:vanish', 'w:smallCaps', 'w:caps', 'w:effect',
)

#: 上面那些里**允许**留在 run 上的。``w:rFonts`` 必须逐个 run 写：
#: 中西文分属两个字体槽（``w:eastAsia`` / ``w:ascii``），样式层给了也可能被
#: run 的旧字体覆盖，而它不属于「会改变字号字重的坏覆盖」，故白名单放行。
_OVERRIDE_ALLOWED = ('w:rFonts',)


def strip_char_overrides(para, keep_bold: bool | None = None):
    """剥掉段落内所有 run 的**字符级直接格式**，让样式成为唯一来源。

    命名样式只是「继承基线」：run 上任何直接格式都优先于样式。用户从别处
    粘一段标题进来、或手动加粗/改字号，样式就形同虚设。所以套样式之后
    必须把这些覆盖清掉，渲染才真正听规范的。

    :param keep_bold: ``None`` 时保留现有粗体设置（只清字号字体等）；
        给 ``True``/``False`` 则连 ``w:b`` 一起定型。
    """
    for r in para.runs:
        rpr = r._element.find(qn('w:rPr'))
        if rpr is None:
            continue
        for tag in _OVERRIDE_RPR_TAGS:
            if tag in ('w:b', 'w:bCs') and keep_bold is not None:
                continue
            for el in rpr.findall(qn(tag)):
                rpr.remove(el)
        if keep_bold is not None:
            for tag in ('w:b', 'w:bCs'):
                for el in rpr.findall(qn(tag)):
                    rpr.remove(el)
                if keep_bold:
                    rpr.append(OxmlElement(tag))
    return para


def _effective_char_fmt(para):
    """返回段落首个 run 的**有效**字符格式 ``(size_pt, bold, eastAsia, color)``。

    先看 run 的直接格式，没有则回落到它的段落样式，再看样式的 basedOn 链。
    """
    def _from_rpr(rpr):
        if rpr is None:
            return {}
        out = {}
        sz = rpr.find(qn('w:sz'))
        if sz is not None:
            try:
                out['size'] = int(sz.get(qn('w:val'))) / 2.0
            except (TypeError, ValueError):
                pass
        _b = rpr.find(qn('w:b'))
        if _b is not None:
            _v = (_b.get(qn('w:val')) or '1').lower()
            out['bold'] = _v not in ('0', 'false', 'off')
        # 没有 w:b 时不写 bold，留给继承链决定
        rf = rpr.find(qn('w:rFonts'))
        if rf is not None and rf.get(qn('w:eastAsia')):
            out['eastAsia'] = rf.get(qn('w:eastAsia'))
        c = rpr.find(qn('w:color'))
        if c is not None:
            out['color'] = (c.get(qn('w:val')) or '').upper()
        return out

    run = para.runs[0] if para.runs else None
    direct = _from_rpr(run._element.find(qn('w:rPr'))) if run is not None else {}

    st = para.style
    chain = []
    while st is not None:
        chain.append(st)
        try:
            st = st.base_style
        except Exception:
            st = None
        if len(chain) > 10:
            break
    merged = {}
    for style in reversed(chain):          # 从最根上的 Normal 往下盖
        merged.update(_from_rpr(style.element.find(qn('w:rPr'))))
    merged.update(direct)                  # 直接格式优先级最高
    return (merged.get('size'), merged.get('bold'),
            merged.get('eastAsia'), merged.get('color'))


def enforce_heading_format(doc, level: int, para=None):
    """确保标题**渲染出来**就是规范的样子，而不只是「套了样式」。

    三步走，对应三种会被用户改坏的地方：

    1. **样式层**——调 :func:`ensure_heading_style`，纠正 Word 样式面板里
       被改过的字号/加粗/间距（改样式会让所有同类标题一起变形）；
    2. **字符层**——:func:`strip_char_overrides` 清掉 run 上的直接格式
       （字号、加粗、字体、字距……），并把粗体按规范定型；
    3. **段落层**——删掉 ``w:spacing`` / ``w:jc`` / ``w:ind`` / ``w:snapToGrid``
       / ``w:outlineLvl`` 的段落级覆盖，让间距对齐回到样式；
       只保留 ``w:pageBreakBefore``（章前分页）这类样式里没有的意图。

    ``para=None`` 时只修样式（用于建文档阶段），随后新建的段落自然会正确。
    """
    spec = HEADING_SPEC[level]
    ensure_heading_style(doc, level)

    if para is None:
        return None

    # --- 样式套用（若当前不是该 Heading 样式，先套上）---
    want_style = f'Heading {level}'
    if para.style.name != want_style:
        try:
            para.style = doc.styles[want_style]
        except KeyError:
            pass
    # --- 字符层：**清掉一切直接格式**，只把中/西文字体名写回去 ---
    #
    # 为什么字号/字重/颜色不写在 run 上：三者样式层已经有了（见 ensure_heading_style），
    # run 上再写一份就是「直接格式压过样式」——用户之后在 Word 样式面板里改字号
    # 会发现改不动，因为 run 那份更优先。让样式成为唯一来源，才既符合规范又受管制。
    strip_char_overrides(para, keep_bold=bool(spec['bold']))
    for r in para.runs:
        _rpr = r._element.get_or_add_rPr()
        for _tag in ('w:sz', 'w:szCs', 'w:b', 'w:bCs', 'w:i', 'w:iCs',
                     'w:color', 'w:u', 'w:spacing', 'w:w', 'w:kern', 'w:position'):
            for _el in _rpr.findall(qn(_tag)):
                _rpr.remove(_el)
        _rf = _rpr.get_or_add_rFonts()
        _rf.set(qn('w:ascii'), EN_FONT)
        _rf.set(qn('w:hAnsi'), EN_FONT)
        _rf.set(qn('w:eastAsia'), HEI)
        _rf.set(qn('w:cs'), EN_FONT)
    # --- 段落层：让间距/对齐/缩进回到样式 ---
    ppr = para._p.get_or_add_pPr()
    for tag in ('w:spacing', 'w:jc', 'w:ind', 'w:snapToGrid', 'w:outlineLvl'):
        for el in ppr.findall(qn(tag)):
            ppr.remove(el)
    return para


def heading_format_problems(doc) -> list:
    """检查所有标题的**有效格式**是否符合 :data:`HEADING_SPEC`。

    只看结果、不看过程：无论问题出在样式被改、还是 run 上有覆盖，
    只要最终字号/加粗/字体/颜色/对齐/大纲级别不对，就报出来。
    """
    problems = []
    for i, para in enumerate(doc.paragraphs):
        sn = para.style.name if para.style is not None else ''
        if not (sn.startswith('Heading') and sn.split()[-1].isdigit()):
            continue
        level = int(sn.split()[-1])
        if level not in HEADING_SPEC or not para.text.strip():
            continue
        spec = HEADING_SPEC[level]
        size, bold, ea, color = _effective_char_fmt(para)
        label = f'第{i}段 {para.text[:14]!r}'
        if size is None or abs(size - spec['size']) > 0.01:
            problems.append(f'{label} 字号 {size}，应为 {spec["size"]}pt')
        if bool(bold) != bool(spec['bold']):
            problems.append(f'{label} 加粗 {bold}，应为 {spec["bold"]}（规范只要求黑体字体，不加粗）')
        if ea != HEI:
            problems.append(f'{label} 中文字体 {ea!r}，应为 {HEI!r}')
        if color not in (None, '000000', 'AUTO'):
            problems.append(f'{label} 颜色 #{color}，应为黑色')
        ppr = para._p.find(qn('w:pPr'))
        if ppr is not None:
            jc = ppr.find(qn('w:jc'))
            if jc is not None and jc.get(qn('w:val')) != 'left':
                problems.append(f'{label} 段落对齐 {jc.get(qn("w:val"))}，应为 left')
            ol = ppr.find(qn('w:outlineLvl'))
            if ol is not None and ol.get(qn('w:val')) != str(level - 1):
                problems.append(f'{label} 大纲级别 {ol.get(qn("w:val"))}，应为 {level - 1}')
        # 段落级的字符覆盖（样式之外的直接格式）
        for r in para.runs:
            rpr = r._element.find(qn('w:rPr'))
            if rpr is None:
                continue
            leftover = [t.split(':')[1] for t in _OVERRIDE_RPR_TAGS
                        if t not in _OVERRIDE_ALLOWED and rpr.find(qn(t)) is not None]
            if leftover:
                problems.append(f'{label} run 上残留直接格式 {leftover}（会盖过样式）')
                break
    return problems


def ensure_style(doc, name: str, style_id: str | None = None,
                 based_on: str = 'Normal', cn: str = SONG, en: str = EN_FONT,
                 size: float = PT_BODY, bold: bool = False, align=None,
                 line: float = LINE_BODY, line_rule: str = 'auto',
                 before: float = 0, after: float = 0,
                 before_lines: float | None = None, after_lines: float | None = None,
                 first_line_chars_: int = 0, outline: int | None = None,
                 left: int = 0, hanging: int | None = None):
    """确保文档中存在 ``name`` 段落样式（不存在则创建），并设置其属性。

    改写既有文档时，把段落 ``p.style = doc.styles[name]`` 即可，
    这样在 Word 的样式面板里仍然可以整体调整。

    :param line_rule: ``'auto'``（Word 默认，配合 ``w:line`` 表示倍数）或 ``'exact'``。
    :param before_lines: 用**行**为单位的段前距（1 行 = 1.0），写 ``w:beforeLines``。
        规范里的「段后 1 行」「段前 0.5 行」必须用行单位——按磅写死只对某一种
        字号成立，标题字号一变就不是「1 行」了。给了它就忽略 ``before``。
    """
    try:
        st = doc.styles[name]
    except KeyError:
        st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        if style_id:
            st.element.set(qn('w:styleId'), style_id)

    # python-docx 里段落样式与字符样式**共用名字空间**。官方模板中存在
    # 「关键词 Char」这类字符样式，直接取用会在 st.paragraph_format 上抛
    # AttributeError。对策：名字被字符样式占用时先改名让位，再建同名段落样式。
    if st.type != WD_STYLE_TYPE.PARAGRAPH:
        st.name = f'{name} (字符样式)'
        st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        if style_id:
            st.element.set(qn('w:styleId'), style_id)

    try:
        st.base_style = doc.styles[based_on]
    except KeyError:
        pass
    st.font.name = en
    st.font.size = Pt(size)
    # 同 set_run：显式写 w:b wal="0/1"，否则 False 会被 python-docx 当作「删除」，
    # 结果样式没有 w:b（=不加粗）而检测却按「继承」判断，极易误判。
    _srpr = st.element.get_or_add_rPr()
    for _tag in ('w:b', 'w:bCs'):
        for _el in _srpr.findall(qn(_tag)):
            _srpr.remove(_el)
    for _tag in ('w:b', 'w:bCs'):
        _el = OxmlElement(_tag)
        _el.set(qn('w:val'), '1' if bold else '0')
        _srpr.append(_el)
    # 注意：不要再碰 st.font.bold —— 它会把刚写好的 w:b 删掉
    rf = _style_rpr(st).get_or_add_rFonts()
    rf.set(qn('w:ascii'), en)
    rf.set(qn('w:hAnsi'), en)
    rf.set(qn('w:eastAsia'), cn)
    # 清掉从 python-docx 内置模板继承来的主题蓝，钉成黑色
    strip_style_color(st, (0, 0, 0))
    pf = st.paragraph_format
    if align is not None:
        pf.alignment = align
    pf.line_spacing = line
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    # 行单位的段前/段后（w:beforeLines / w:afterLines，单位 1/100 行）
    _ppr0 = st.element.get_or_add_pPr()
    _sp0 = _ppr0.find(qn('w:spacing'))
    if _sp0 is None:
        _sp0 = OxmlElement('w:spacing')
        _ppr0.append(_sp0)
    _sp0.set(qn('w:lineRule'), line_rule)
    _sp0.set(qn('w:line'), str(int(round(line * 240))))
    if before_lines is not None:
        _sp0.set(qn('w:beforeLines'), str(int(round(before_lines * 100))))
        _sp0.set(qn('w:before'), str(int(round(before_lines * 12 * 20))))
    if after_lines is not None:
        _sp0.set(qn('w:afterLines'), str(int(round(after_lines * 100))))
        _sp0.set(qn('w:after'), str(int(round(after_lines * 12 * 20))))
    # 段落样式里的 ind / outline 走 XML
    ppr = st.element.get_or_add_pPr()
    ind = ppr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        ppr.append(ind)
    ind.set(qn('w:firstLineChars'), str(first_line_chars_ * 100))
    ind.set(qn('w:firstLine'), str(first_line_chars_ * 240))
    ind.set(qn('w:left'), str(left))
    if hanging is not None:
        ind.set(qn('w:hangingChars'), str(hanging))
        ind.set(qn('w:hanging'), str(hanging))
    if outline is not None:
        for old in ppr.findall(qn('w:outlineLvl')):
            ppr.remove(old)
        ol = OxmlElement('w:outlineLvl')
        ol.set(qn('w:val'), str(outline))
        ppr.append(ol)
    return st


def apply_style_set(doc):
    """把 DUT 官方模板的自定义段落样式全部建好（幂等）。

    返回已就绪的样式名列表。改写既有文档前先调用它，后面就能按名套样式。
    """
    A = WD_ALIGN_PARAGRAPH
    made = []
    specs = [
        # name, kwargs
        ('Normal', dict(based_on='Normal', cn=SONG, size=PT_BODY, align=A.JUSTIFY,
                        line=LINE_BODY, first_line_chars_=2)),
        ('摘要题目', dict(cn=HEI, size=PT_ABSTRACT, bold=True, align=A.CENTER,
                        line=LINE_HEADING, after=12)),
        ('ABSTRACT', dict(cn=EN_FONT, size=PT_ABSTRACT, bold=True, align=A.CENTER,
                          line=LINE_HEADING, after=12)),
        ('目录', dict(cn=HEI, size=PT_ABSTRACT, bold=True, align=A.CENTER,
                     line=LINE_HEADING, after=12)),
        ('关键词', dict(cn='仿宋_GB2312', size=PT_BODY, align=A.JUSTIFY,
                       line=LINE_BODY, first_line_chars_=0)),
        ('关键词题头', dict(cn=HEI, size=PT_BODY, align=A.JUSTIFY, line=LINE_BODY,
                          first_line_chars_=0)),
        ('英文摘要正文', dict(cn=EN_FONT, en=EN_FONT, size=PT_BODY, align=A.JUSTIFY,
                            line=LINE_BODY, first_line_chars_=2)),
        ('英文关键词', dict(cn=EN_FONT, size=PT_BODY, align=A.JUSTIFY, line=LINE_BODY,
                          first_line_chars_=0)),
        ('Keywords', dict(cn=EN_FONT, size=PT_BODY, bold=True, align=A.JUSTIFY,
                          line=LINE_BODY, first_line_chars_=0)),
        ('图名中文', dict(cn=SONG, size=PT_CAPTION, align=A.CENTER, line=1.0)),
        ('图名英文', dict(cn=EN_FONT, en=EN_FONT, size=PT_CAPTION, align=A.CENTER, line=1.0)),
        ('Caption', dict(cn=SONG, size=PT_CAPTION, align=A.CENTER, line=1.0)),
        ('参考文献标题', dict(cn=HEI, size=PT_ABSTRACT, bold=True, align=A.CENTER,
                           line=LINE_HEADING, after=12, outline=0)),
        ('参考文献正文', dict(cn=SONG, size=PT_CAPTION, align=A.JUSTIFY,
                           line=LINE_BODY, first_line_chars_=0, left=170, hanging=170)),
        ('公式', dict(cn=SONG, size=PT_BODY, align=A.LEFT, line=1.0, before=6, after=6,
                     first_line_chars_=0)),
        ('附录', dict(cn=HEI, size=PT_ABSTRACT, bold=True, align=A.CENTER,
                     line=LINE_HEADING, after=12)),
        ('致谢', dict(cn=HEI, size=PT_ABSTRACT, bold=True, align=A.CENTER,
                     line=LINE_HEADING, after=12)),
        ('作者简介', dict(cn=HEI, size=PT_ABSTRACT, bold=True, align=A.CENTER,
                       line=LINE_HEADING, after=12)),
        ('授权说明正文', dict(cn=SONG, size=PT_BODY, align=A.JUSTIFY, line=LINE_BODY,
                          first_line_chars_=2)),
        ('关键词 Char', dict(cn=SONG, size=PT_BODY)),
    ]
    for name, kw in specs:
        ensure_style(doc, name, **kw)
        made.append(name)

    # 三级标题样式——逐项按 HEADING_SPEC 落，含「不加粗」与「行单位」间距
    for lvl in (1, 2, 3):
        spec = HEADING_SPEC[lvl]
        name = f'Heading {lvl}'
        ensure_style(doc, name, based_on='Normal', cn=HEI, size=spec['size'],
                     bold=spec['bold'], align=A.LEFT, line=spec['line'],
                     line_rule='auto', before_lines=spec['before_lines'],
                     after_lines=spec['after_lines'],
                     first_line_chars_=0, outline=lvl - 1)
        # ensure_style 只写 w:b=True，不会主动删；这里显式清掉残留的 w:b
        rpr = doc.styles[name].element.get_or_add_rPr()
        for tag in ('w:b', 'w:bCs'):
            for el in rpr.findall(qn(tag)):
                rpr.remove(el)
        made.append(name)
    return made


def _set_page(sec, top=3.5, bottom=2.5, left=2.5, right=2.5,
              header=2.5, footer=2.0, width=PAGE_W_CM, height=29.7):
    sec.page_width = Cm(width)
    sec.page_height = Cm(height)
    sec.top_margin = Cm(top)
    sec.bottom_margin = Cm(bottom)
    sec.left_margin = Cm(left)
    sec.right_margin = Cm(right)
    sec.header_distance = Cm(header)
    sec.footer_distance = Cm(footer)
    sec.gutter = Cm(0)
    return sec


def new_doc(odd_page_start: bool = True) -> _Doc:
    """新建已套好 DUT 页面设置 + 全部样式的文档。

    A4，上 3.5 / 下 2.5 / 左 2.5 / 右 2.5 cm，页眉 2.5 cm，页脚 2.0 cm。
    """
    doc = Document()
    apply_style_set(doc)
    _set_page(doc.sections[0])
    if odd_page_start:
        doc.sections[0].start_type = WD_SECTION_START.ODD_PAGE
    _enable_even_odd_headers(doc)
    normal = doc.styles['Normal']
    normal.font.size = Pt(PT_BODY)
    return doc


def add_section(doc, start_type=WD_SECTION_START.NEW_PAGE, kind: str | None = None,
                title_cn: str | None = None, restart: bool | None = None):
    """新增一节并继承 DUT 页面设置。

    :param kind: 该节在论文中的角色，决定页眉/页脚/页码。**强烈建议传**，因为
        新节默认会「链接到前一节」——封面节的页脚是校名，若不显式断开链接，
        后面每一页都会印上封面的校名页脚，且不会有页码。取值：

        * ``'cover'`` — 封面：无页眉、页脚为校名（无页码）
        * ``'front'`` — 摘要 / ABSTRACT / 目录：页眉校名，页脚 ``- I -``（罗马数字从 I 起）
        * ``'body'``  — 正文起：页眉校名，页脚 ``- 1 -``（阿拉伯数字从 1 起）
        * ``None``    — 只设页面，页眉页脚自行处理

    :param restart: 是否**重置**页码为 1。默认 ``None`` = 按 kind 自动判断：
        ``front`` 只在第一次调用时重置（后续摘要/ABSTRACT/目录节应继续编号），
        ``body`` 同理只在第一次调用时重置。显式传 True/False 可覆盖。
        留空会导致每节都从 I / 1 重新开始——这是分节页码最常见的坑。
    """
    sec = doc.add_section(start_type)
    _set_page(sec)
    # **关键**：python-docx 的 add_section() 会把前一节的 headerReference /
    # footerReference 原样复制给新节——两节于是指向**同一个**页眉/页脚部件。
    # 之后改新节的页眉，会把前一节的一起改掉（封面就莫名长出页眉）。
    # 这里先显式断开链接，让新节拥有自己的部件。
    _unlink_headfoot(sec)
    if kind == 'cover':
        set_page_number_footer(sec, fmt='decimal', start=1, style='none')
        set_header_pair(sec, odd_text='', even_text='')
        return sec

    if kind in ('front', 'body'):
        if restart is None:
            restart = not _section_kind_seen(doc, kind)
        fmt = 'upperRoman' if kind == 'front' else 'decimal'
        set_page_number_footer(sec, fmt=fmt, start=(1 if restart else None),
                               style='dash')
        set_header_pair(sec, odd_text='大连理工大学硕士学位论文',
                        even_text=title_cn or '学位论文题目')
    return sec


def _unlink_headfoot(sec):
    """让该节拥有**自己的、空的**页眉/页脚部件，而不是与前一节共用。

    这里连踩三个坑，注释留全，免得以后再踩：

    1. ``doc.add_section()`` 会把前一节的 ``headerReference`` / ``footerReference``
       （rId）复制给新节，**不新建部件**——两节共用同一个部件，
       于是改新节的页眉会把前一节的一起改掉（封面就莫名长出页眉）。
    2. python-docx 的 ``is_linked_to_previous`` 是「按需创建」语义：
       ``_has_definition`` 只看 reference 在不在；复制来的 reference 已存在，
       直接赋 ``False`` 是 **no-op**。
    3. 更阴的是 ``_get_or_add_definition()``：本节点「无定义」时会**往上找**
       前一节的定义并复用其部件。所以仅仅删掉 reference 还不够——
       必须真正为本节新建一个部件，否则 ``.header.element`` 写的是前一节的内容。

    可靠做法：**先赋 True（删掉继承来的定义）再赋 False（新建本节专属空部件）**。
    两次赋值状态一定发生变化，因此必然触达 ``_drop_definition`` / ``_add_definition``。
    """
    for attr in ('header', 'even_page_header', 'first_page_header',
                 'footer', 'even_page_footer', 'first_page_footer'):
        proxy = getattr(sec, attr, None)
        if proxy is None:
            continue
        try:
            proxy.is_linked_to_previous = True    # 删掉继承来的定义
            proxy.is_linked_to_previous = False   # 新建本节专属的空部件
        except Exception:
            pass
    return sec


def _section_kind_seen(doc, kind: str) -> bool:
    """判断 ``doc`` 中是否已经出现过同类节（用于决定要不要重置页码）。

    直接检查**已存在各节的页码格式**，不把状态挂在 Document 对象上——
    这样即使节是被别的代码加的，也能正确判断。
    """
    want = 'upperRoman' if kind == 'front' else 'decimal'
    for sec in doc.sections:
        pg = sec._sectPr.find(qn('w:pgNumType'))
        if pg is None:
            continue
        if pg.get(qn('w:fmt')) == want and pg.get(qn('w:start')) is not None:
            return True
    return False


def _enable_even_odd_headers(doc):
    """写 ``w:evenAndOddHeaders``——奇偶页页眉不同的前提。"""
    st = doc.settings.element
    if st.find(qn('w:evenAndOddHeaders')) is None:
        st.append(OxmlElement('w:evenAndOddHeaders'))
    return st


def add_page_field(p, fmt: str = 'PAGE', cn: str = SONG, en: str = EN_FONT,
                   size: float = PT_PAGENUM):
    """在段落中插入域（``PAGE`` / ``NUMPAGES`` / ``STYLEREF`` …）。"""
    r = p.add_run()
    f1 = OxmlElement('w:fldChar')
    f1.set(qn('w:fldCharType'), 'begin')
    it = OxmlElement('w:instrText')
    it.set(qn('xml:space'), 'preserve')
    it.text = f' {fmt} '
    f2 = OxmlElement('w:fldChar')
    f2.set(qn('w:fldCharType'), 'end')
    r._r.append(f1)
    r._r.append(it)
    r._r.append(f2)
    set_run(r, cn=cn, en=en, size=size)
    return r


def set_pgnum(section, fmt: str = 'decimal', start: int | None = 1):
    """设置本节页码格式。``fmt``：``decimal`` / ``upperRoman`` / ``lowerRoman`` / ``none``。"""
    sectPr = section._sectPr
    pg = sectPr.find(qn('w:pgNumType'))
    if pg is None:
        pg = OxmlElement('w:pgNumType')
        sectPr.append(pg)
    pg.set(qn('w:fmt'), fmt)
    if start is not None:
        pg.set(qn('w:start'), str(start))
    else:
        pg.attrib.pop(qn('w:start'), None)
    return section


def _footer_paragraphs(sec, which='footer'):
    part = {'footer': sec.footer, 'header': sec.header,
            'even_footer': sec.even_page_footer, 'even_header': sec.even_page_header,
            'first_footer': sec.first_page_footer, 'first_header': sec.first_page_header}[which]
    return part


def _fill_footer_part(part, style: str, linked: bool | None = None):
    """往一个页脚部件里写页码（或清空）。

    ``linked=None``（默认）表示**保持现状**——已经有定义就不碰它。
    绝不能无脑赋 ``False``：python-docx 里赋 ``True`` 是「删掉本节的页脚定义」，
    而赋 ``False`` 只在「本来就没有定义」时才会新建。
    对已经有定义的部件赋错值，会把页脚/页眉直接删掉。
    """
    if linked is not None:
        part.is_linked_to_previous = linked
    elif not part._has_definition:
        part.is_linked_to_previous = False
    p = part.paragraphs[0] if part.paragraphs else part.add_paragraph()
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    fmt_para(p, align=WD_ALIGN_PARAGRAPH.CENTER, line=1.0, before=0, after=0,
             snap_to_grid=False)
    _set_ind(p, firstLineChars=0, firstLine=0)
    if style == 'none':
        return p
    if style == 'dash':
        set_run(p.add_run('- '), cn=SONG, en=EN_FONT, size=PT_PAGENUM)
    add_page_field(p)
    if style == 'dash':
        set_run(p.add_run(' -'), cn=SONG, en=EN_FONT, size=PT_PAGENUM)
    return p


def set_page_number_footer(sec, fmt: str = 'decimal', start: int | None = 1,
                           style: str = 'dash', which: str = 'all',
                           linked: bool | None = None):
    """页脚页码。``style``：

    * ``dash`` → ``- 7 -``（DUT 模板样式）
    * ``plain`` → ``7``
    * ``none`` → 清空页脚

    ``fmt`` / ``start`` 会写到该节的 ``w:pgNumType``。

    :param which: ``'all'``（默认）会同时设置**首页 / 奇数页 / 偶数页**三种页脚。

        **务必用 ``all``**：文档开启 ``w:evenAndOddHeaders`` 后，偶数页使用
        ``w:footerReference type="even"`` 指向的部件。若只设 ``footer``，
        偶数页会「链接到前一节」而套用**封面**的校名页脚，页码会在偶数页上凭空消失。
    """
    set_pgnum(sec, fmt, start)
    if which == 'all':
        parts = [sec.footer, sec.even_page_footer, sec.first_page_footer]
    elif which == 'footer':
        parts = [sec.footer]
    else:
        parts = [_footer_paragraphs(sec, which)]
    out = None
    for part in parts:
        out = _fill_footer_part(part, style, linked=linked)
    return out


def set_header_pair(sec, odd_text: str = '大连理工大学硕士学位论文',
                    even_text: str | None = None, cn: str = SONG,
                    en: str = EN_FONT, size: float = PT_CAPTION,
                    link_odd: bool | None = None, link_even: bool | None = None,
                    which: str = 'all'):
    """奇偶页页眉。

    * 奇数页 → ``odd_text``（默认「大连理工大学硕士学位论文」，宋体五号居中）
    * 偶数页 → ``even_text``（默认论文中文题目）

    需先调用 ``_enable_even_odd_headers``（``new_doc()`` 已自动调用）。

    :param which: ``'all'``（默认）会同时设置**首页 / 奇数页 / 偶数页**三种页眉。
        与 :func:`set_page_number_footer` 同理——只设 ``header`` 时，偶数页页眉
        会「链接到前一节」而套用封面的空页眉或前节的页眉。

    :param link_odd: ``None``（默认）保持现状。**不要传 True**——python-docx 里
        赋 ``True`` 是「删掉本节的页眉定义、改为继承前一节」，会把刚建好的页眉删掉。
    """
    def _fill(part, text, linked):
        # linked=None → 保持现状（有定义就保留），见 _fill_footer_part 的说明
        if linked is not None:
            part.is_linked_to_previous = linked
        elif not part._has_definition:
            part.is_linked_to_previous = False
        p = part.paragraphs[0] if part.paragraphs else part.add_paragraph()
        for r in list(p.runs):
            r._r.getparent().remove(r._r)
        fmt_para(p, align=WD_ALIGN_PARAGRAPH.CENTER, line=1.0, before=0, after=0,
                 snap_to_grid=False)
        _set_ind(p, firstLineChars=0, firstLine=0)
        if text:
            set_run(p.add_run(text), cn=cn, en=en, size=size)

    _fill(sec.header, odd_text, link_odd)
    if even_text is not None:
        _fill(sec.even_page_header, even_text, link_even)
    if which == 'all':
        # 首页页眉与奇数页一致；否则首页会继承前一节的页眉
        _fill(sec.first_page_header, odd_text, link_odd)
    return sec


def add_toc_field(doc, levels: str = '1-3', placeholder: str | None = None):
    """插入自动目录域 ``TOC \\o "1-3" \\h \\z \\u``。

    配合 ``enable_update_fields()``，Word 打开时会自动生成/更新。
    """
    placeholder = placeholder or '（在 Word 中右键此处选择“更新域”生成目录）'
    p = doc.add_paragraph()
    fmt_para(p, align=WD_ALIGN_PARAGRAPH.LEFT, line=LINE_BODY, before=0, after=0,
             snap_to_grid=False)
    _set_ind(p, firstLineChars=0, firstLine=0)
    r = p.add_run()
    f1 = OxmlElement('w:fldChar')
    f1.set(qn('w:fldCharType'), 'begin')
    it = OxmlElement('w:instrText')
    it.set(qn('xml:space'), 'preserve')
    it.text = f' TOC \\o "{levels}" \\h \\z \\u '
    f2 = OxmlElement('w:fldChar')
    f2.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t')
    t.text = placeholder
    f3 = OxmlElement('w:fldChar')
    f3.set(qn('w:fldCharType'), 'end')
    for el in (f1, it, f2, t, f3):
        r._r.append(el)
    set_run(r, cn=SONG, en=EN_FONT, size=PT_BODY)
    return p


def add_list_of_field(doc, kind: str = 'figure', placeholder: str | None = None):
    """图目录 / 表目录域。

    :param kind: ``figure`` → ``TOC \\h \\z \\c "图"``；``table`` → ``\\c "表"``；
                 ``equation`` → ``\\c "公式"``。
    """
    label = {'figure': '图', 'table': '表', 'equation': '公式'}[kind]
    placeholder = placeholder or f'（在 Word 中右键“更新域”生成{label}目录）'
    p = doc.add_paragraph()
    fmt_para(p, align=WD_ALIGN_PARAGRAPH.LEFT, line=LINE_BODY, snap_to_grid=False)
    _set_ind(p, firstLineChars=0, firstLine=0)
    r = p.add_run()
    f1 = OxmlElement('w:fldChar')
    f1.set(qn('w:fldCharType'), 'begin')
    it = OxmlElement('w:instrText')
    it.set(qn('xml:space'), 'preserve')
    it.text = f' TOC \\h \\z \\c "{label}" '
    f2 = OxmlElement('w:fldChar')
    f2.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t')
    t.text = placeholder
    f3 = OxmlElement('w:fldChar')
    f3.set(qn('w:fldCharType'), 'end')
    for el in (f1, it, f2, t, f3):
        r._r.append(el)
    set_run(r, cn=SONG, en=EN_FONT, size=PT_BODY)
    return p


def enable_update_fields(doc, on: bool = True):
    """写 ``w:updateFields val="true"``——Word 打开时自动更新目录与页码。"""
    st = doc.settings.element
    for old in st.findall(qn('w:updateFields')):
        st.remove(old)
    if on:
        el = OxmlElement('w:updateFields')
        el.set(qn('w:val'), 'true')
        st.append(el)
    return st


# --------------------------------------------------------------------------
# 封面
# --------------------------------------------------------------------------

def add_cover(doc, title_cn: str, title_en: str = '', fields: Sequence[tuple] = (),
              school_cn: str = '大连理工大学', school_en: str = 'Dalian University of Technology'):
    """DUT 封面页。

    :param fields: ``[(标签, 值), ...]``，如 ``[('作者姓名', '张三'), ('学号', '12345678')]``
                   标签走黑体/华文细黑，值加下划线。
    封面所在节**不编页码**。
    """
    A = WD_ALIGN_PARAGRAPH
    for _ in range(2):
        add_blank(doc)
    p = doc.add_paragraph()
    set_run(p.add_run('硕 士 学 位 论 文'), cn='华文细黑', en='STXihei',
            size=FONT_SIZE['小二'] + 6, bold=True)
    fmt_para(p, align=A.CENTER, line=1.5, before=0, after=24, snap_to_grid=False)
    _set_ind(p, firstLineChars=0, firstLine=0)

    p = doc.add_paragraph()
    set_run(p.add_run(title_cn), cn='华文细黑', en='STXihei',
            size=FONT_SIZE['二号'], bold=True)
    fmt_para(p, align=A.CENTER, line=1.5, before=0, after=12, snap_to_grid=False)
    _set_ind(p, firstLineChars=0, firstLine=0)

    if title_en:
        p = doc.add_paragraph()
        set_run(p.add_run(title_en), cn=EN_FONT, en=EN_FONT,
                size=FONT_SIZE['三号'], bold=True)
        fmt_para(p, align=A.CENTER, line=1.5, before=0, after=36, snap_to_grid=False)
        _set_ind(p, firstLineChars=0, firstLine=0)

    if fields:
        # 字段块用**隐藏边框表格**：两列固定列宽，标签右对齐、值居中，
        # 值下方是单元格下边框（等宽横线）。
        # 用段落 + 文字下划线的话，横线长度会随内容变化，永远对不齐。
        add_blank(doc)
        # 不要设 row_h_cm：行高下限（w:trHeight）会盖过固定行高，
        # 把下边框整体推低到下一行旁边。行距交给 line_height_pt 控制。
        add_field_table(doc, list(fields), size=FONT_SIZE['小三'])

    # 封面页脚：校名（华文行楷 + 英文校名）
    sec = doc.sections[-1]
    # 规范：「封一、封二不编入页码」，封面页脚只有校名，**无页眉**。
    # 必须显式断开与后一节的链接，否则 Word 会让封面继承正文页眉。
    for which in ('header', 'even_header', 'first_header'):
        try:
            _footer_paragraphs(sec, which).is_linked_to_previous = False
        except Exception:
            pass
    for which, text, cn, size in (
            ('footer', school_cn, '华文行楷', FONT_SIZE['小二']),
            ('even_footer', school_cn, '华文行楷', FONT_SIZE['小二'])):
        part = _footer_paragraphs(sec, which)
        part.is_linked_to_previous = False
        pf = part.paragraphs[0] if part.paragraphs else part.add_paragraph()
        for r in list(pf.runs):
            r._r.getparent().remove(r._r)
        fmt_para(pf, align=A.CENTER, line=1.0, snap_to_grid=False)
        _set_ind(pf, firstLineChars=0, firstLine=0)
        set_run(pf.add_run(text), cn=cn, en=cn, size=size)
        pf2 = part.add_paragraph()
        fmt_para(pf2, align=A.CENTER, line=1.0, snap_to_grid=False)
        _set_ind(pf2, firstLineChars=0, firstLine=0)
        set_run(pf2.add_run(school_en), cn=EN_FONT, en=EN_FONT,
                size=FONT_SIZE['小四'])
    return doc


# --------------------------------------------------------------------------
# 段落分类（改写既有文档的核心）
# --------------------------------------------------------------------------

CH_LEVEL1 = 'ch1'      # 「1 绪论」类章标题
CH_LEVEL2 = 'ch2'      # 「1.1 xxx」节标题
CH_LEVEL3 = 'ch3'      # 「1.1.1 xxx」小节标题
CH_FRONT = 'front'     # 摘要 / ABSTRACT / 目录 / 参考文献 / 附录 / 致谢 等居中章级标题
CH_CAPTION_FIG = 'capf'   # 图题
CH_CAPTION_TAB = 'capt'   # 表题
CH_REF = 'ref'         # 参考文献条目
CH_BODY = 'body'       # 正文
CH_EMPTY = 'empty'     # 空段
CH_BODY_EN = 'body_en'  # 英文正文（英文摘要等）：Times New Roman，不套宋体
CH_KEYWORDS = 'kw'      # 「关键词：」「Key Words：」行

_CJK_RE = re.compile(r'[\u4e00-\u9fff]')
_RE_KEYWORDS = re.compile(r'^\s*(关键词\s*[:：]|Key\s*Words\s*[:：]|关键字\s*[:：])', re.I)


def is_latin_prose(text: str, threshold: float = 0.35) -> bool:
    """判断是否「英文正文」——中日韩字符占比低于阈值即算英文。

    英文摘要正文用的是 Times New Roman，**不是宋体**，也不该按中文正文的
    「宋体 + 首行缩进 2 字符」去检查；不区分的话审计会把英文摘要全部误报。
    """
    t = (text or '').strip()
    if not t:
        return False
    cjk = len(_CJK_RE.findall(t))
    return cjk / max(len(t), 1) < threshold


def _is_prose(text: str, min_len: int = 25) -> bool:
    """判断一段文字是否算「真正的正文」（而不是标题 / 占位小标签）。

    规则：去掉空白后长度 >= ``min_len``；且不像编号标题、图表题、参考文献条目。
    """
    t = _WS_RE.sub('', text or '')
    if len(t) < min_len:
        return False
    if _RE_REFLIST.match(text or ''):
        return False
    if _RE_FIGCAP.match(text or '') or _RE_TABCAP.match(text or ''):
        return False
    if _RE_CH3.match(text or '') or _RE_CH2.match(text or ''):
        return False
    if _norm_title(text) in _FRONT_TITLES:
        return False
    return True


#: 章标题：「第1章 xxx」「第一章 xxx」「1 绪论」「1. 绪论」
_RE_CH1 = re.compile(r'^\s*(?:第\s*[一二三四五六七八九十百零\d]+\s*[章篇]|[0-9]{1,2})[\s、.．]+\S')
#: 节标题：「1.1 xxx」「1.1. xxx」「1.1xxx」——不能误吞「1.1.1」
_RE_CH2 = re.compile(r'^\s*\d{1,2}\.\d{1,2}(?:\s*[、.．]\s*|\s+)\S')
#: 小节标题：「1.1.1 xxx」「1.1.1xxx」
_RE_CH3 = re.compile(r'^\s*\d{1,2}\.\d{1,2}\.\d{1,2}(?:\s*[、.．]\s*|\s+)\S')
_RE_REFLIST = re.compile(r'^\s*\[\d+(?:[,\-–]\d+)*\]\s*\S')
_RE_FIGCAP = re.compile(r'^\s*(?:图|Fig\.?|Figure|附录[-－]图|App\.\s*Fig\.?)\s*\d')
_RE_TABCAP = re.compile(r'^\s*(?:表|Tab\.?|Table|附录[-－]表|App\.\s*Tab\.?)\s*\d')

#: 「居中章级标题」的识别表——键是去空格后的标题文本
_WS_RE = re.compile(r'[\s\u3000]+')   # \s 在部分实现里不含全角空格 U+3000


def _norm_title(text):
    '''标题归一化：去掉所有空白（含全角空格），用于查表匹配。'''
    return _WS_RE.sub('', text or '')


_FRONT_TITLES = {
    '摘要': ('摘要题目', '摘　　要'),
    'ABSTRACT': ('ABSTRACT', 'ABSTRACT'),
    '目录': ('目录', '目　　录'),
    'TABLEOFCONTENTS': ('目录', 'TABLE OF CONTENTS'),
    '图目录': ('摘要题目', '图目录'),
    '表目录': ('摘要题目', '表目录'),
    '主要符号表': ('摘要题目', '主要符号表'),
    '符号表': ('摘要题目', '主要符号表'),
    '参考文献': ('参考文献标题', '参 考 文 献'),
    'REFERENCES': ('参考文献标题', 'References'),
    '致谢': ('致谢', '致　　谢'),
    'ACKNOWLEDGEMENTS': ('致谢', 'Acknowledgements'),
    '攻读硕士学位期间科研项目及科研成果': ('作者简介', '攻读硕士学位期间科研项目及科研成果'),
    '攻读学位期间科研成果': ('作者简介', '攻读硕士学位期间科研项目及科研成果'),
    '作者简介': ('作者简介', '作者简介'),
}
_RE_APPENDIX_TITLE = re.compile(r'^\s*附\s*录\s*[A-ZＡ-Ｚ一二三四五六七八九十]?\s*[：:、.．]?\s*(.*)$')


def classify_paragraph(p, doc=None) -> str:
    """判断一个段落在 DUT 体系里属于哪一类。返回上面的 ``CH_*`` 常量之一。"""
    text = (p.text or '').strip()
    if not text:
        return CH_EMPTY
    style_name = p.style.name if p.style is not None else ''

    # 已经带官方样式名 → 直接采信
    if style_name in ('摘要题目', 'ABSTRACT', '目录', '参考文献标题', '附录', '致谢',
                      '作者简介', '图名中文', '图名英文', 'Caption',
                      '图名中文 Char', '图名英文 Char'):
        if style_name in ('图名中文', '图名英文', 'Caption'):
            return CH_CAPTION_FIG if _RE_FIGCAP.match(text) else (
                CH_CAPTION_TAB if _RE_TABCAP.match(text) else CH_CAPTION_FIG)
        return CH_FRONT
    if style_name == '参考文献正文':
        return CH_REF

    if _RE_FIGCAP.match(text):
        return CH_CAPTION_FIG
    if _RE_TABCAP.match(text):
        return CH_CAPTION_TAB
    if len(text) <= 60:
        if _RE_CH3.match(text):
            return CH_LEVEL3
        if _RE_CH2.match(text):
            return CH_LEVEL2
    if text.startswith('图') or text.startswith('表'):
        # 「图 1.1 xxx」不带数字的兜底已由正则覆盖；这里防误判
        pass

    key = _norm_title(text)
    if key in _FRONT_TITLES:
        return CH_FRONT
    if _RE_APPENDIX_TITLE.match(text) and len(text) < 40:
        return CH_FRONT

    if _RE_KEYWORDS.match(text):
        return CH_KEYWORDS
    if _RE_REFLIST.match(text):
        return CH_REF
    if _RE_CH1.match(text) and len(text) < 60:
        return CH_LEVEL1
    if style_name.startswith('Heading'):
        lvl = int(style_name.split()[-1]) if style_name.split()[-1].isdigit() else 1
        return {1: CH_LEVEL1, 2: CH_LEVEL2}.get(lvl, CH_LEVEL3)
    if is_latin_prose(text):
        return CH_BODY_EN
    return CH_BODY


# --------------------------------------------------------------------------
# 参考文献
# --------------------------------------------------------------------------

_REF_PATTERNS = {
    'J': re.compile(r'\[J\]', re.I),
    'M': re.compile(r'\[M\]', re.I),
    'C': re.compile(r'\[C\]', re.I),
    'D': re.compile(r'\[D\]', re.I),
    'R': re.compile(r'\[R\]', re.I),
    'N': re.compile(r'\[N\]', re.I),
    'S': re.compile(r'\[S\]', re.I),
    'P': re.compile(r'\[P\]', re.I),
    'EB': re.compile(r'\[(?:EB|OL|DB|CP|J/OL|M/OL|EB/OL)\]', re.I),
}


def detect_ref_type(entry: str) -> str:
    """从条目文本里识别 GB/T 7714 文献类型标志，识别不出返回 ``'?'``。"""
    for code, pat in _REF_PATTERNS.items():
        if pat.search(entry):
            return code
    return '?'


def format_reference_entry(entry: str, index: int | None = None) -> str:
    """规范化参考文献条目的序号与空格。

    * 统一成 ``[n] `` 前缀
    * 中文标点后补一个空格，英文标点保持
    * 类型标志统一成大写带方括号
    * 修正 ``et al`` / ``等`` 前多余的空格
    """
    s = re.sub(r'^\s*\[\s*(\d+)\s*\]\s*', '', entry.strip())
    s = re.sub(r'^\s*(\d+)\s*[.、．]\s*', '', s)
    s = re.sub(r'\s+', ' ', s)
    # 类型标志规范化
    s = re.sub(r'\[\s*([a-zA-Z]{1,3})(\s*/\s*([a-zA-Z]{2}))?\s*\]',
               lambda m: '[' + m.group(1).upper() + (('/' + m.group(3).upper()) if m.group(3) else '') + ']',
               s)
    # 中文标点后补空格（但不在数字与单位之间乱加）
    s = re.sub(r'([。；：，、])\s*', r'\1 ', s)
    s = re.sub(r'\.\s*', '. ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    s = re.sub(r'\s+([,.;:])', r'\1', s)
    s = s.replace('et al.', 'et al').replace('ET AL', 'et al')
    n = index if index is not None else 1
    return f'[{n}] {s}'


_REF_HEADINGS = {'参考文献', 'REFERENCES', 'References'}


def find_reference_blocks(paras: Sequence) -> list[tuple[int, int]]:
    """定位所有参考文献块，返回 ``[(标题下标, 块结束下标), ...]``。

    判定方式：找到「参考文献」标题后，**向后收集连续**的条目段（``CH_REF``）与空行；
    遇到第一个既非条目也非空行的段落即认为块结束。

    这样比「一直吃到下一个标题」稳健得多——后者会把标题识别错误放大成
    「整块引用丢失」，而且遇到章内编号（如「2 结论」）时会直接截断。
    """
    blocks: list[tuple[int, int]] = []
    for i, p in enumerate(paras):
        key = _norm_title(p.text)
        if key not in _REF_HEADINGS:
            continue
        j = i + 1
        last = i
        while j < len(paras):
            kind = classify_paragraph(paras[j])
            if kind == CH_REF:
                last = j
                j += 1
            elif kind == CH_EMPTY:
                j += 1
            else:
                break
        blocks.append((i, last))
    return blocks


def order_references(doc, ref_start_heading: str = '参考文献') -> dict:
    """把参考文献条目按**正文首次引用顺序**重排，并删除未被引用的条目。

    支持文档中存在多个参考文献块；引用顺序取**第一个参考文献块之前**的正文。

    :returns: ``{'kept': n, 'removed': [str, ...], 'reordered': bool}``
    """
    paras = list(doc.paragraphs)
    blocks = find_reference_blocks(paras)
    if not blocks:
        return {'kept': 0, 'removed': [], 'reordered': False,
                'error': '未找到参考文献标题'}

    if _norm_title(ref_start_heading) not in _REF_HEADINGS:
        _REF_HEADINGS.add(_norm_title(ref_start_heading))
    body_end = min(b[0] for b in blocks)          # 正文截止到第一个参考文献块之前
    ref_paras: list = []
    for _title_i, block_end in blocks:
        for p in paras[_title_i + 1:block_end + 1]:
            if classify_paragraph(p) == CH_REF:
                ref_paras.append(p)

    # 全文引用顺序（取第一个参考文献块之前的正文）
    cited_order: list[int] = []
    for p in paras[:body_end]:
        for m in _CITE_RE.findall(p.text or ''):
            for part in m.strip('[]').split(','):
                part = part.strip()
                if '-' in part or '–' in part:
                    bits = re.split(r'[-–]', part, maxsplit=1)
                    if len(bits) != 2:
                        continue
                    a, b = bits
                    try:
                        cited_order.extend(range(int(a), int(b) + 1))
                    except ValueError:
                        continue
                elif part.isdigit():
                    cited_order.append(int(part))
    seen, order = set(), []
    for n in cited_order:
        if n not in seen:
            seen.add(n)
            order.append(n)

    entries: dict[int, str] = {}
    for p in ref_paras:
        m = re.match(r'^\s*\[\s*(\d+)\s*\]', p.text or '')
        if m:
            entries[int(m.group(1))] = p.text.strip()[m.end():].strip()

    removed = [f'[{n}] {t}' for n, t in entries.items() if n not in seen]
    # 只在确实抓到引用信息时才按引用顺序重排，否则保留原有编号顺序
    reordered = bool(order)
    if reordered:
        final = [n for n in order if n in entries] + [n for n in sorted(entries) if n not in seen]
    else:
        final = sorted(entries)

    # 在第一个块的标题后重建条目，并删除其余块里的旧条目
    for p in ref_paras:
        p._p.getparent().remove(p._p)

    prev = paras[blocks[0][0]]
    for new_i, n in enumerate([x for x in final if (not reordered) or x in seen], start=1):
        p = doc.add_paragraph()
        set_run(p.add_run(format_reference_entry(entries[n], new_i)),
                cn=SONG, en=EN_FONT, size=PT_CAPTION)
        fmt_para(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, line=LINE_BODY, before=0,
                 after=0, snap_to_grid=False, left_indent_pt=8.5, hanging_pt=8.5)
        prev._p.addnext(p._p)
        prev = p
    return {'kept': len([x for x in final if (not reordered) or x in seen]),
            'removed': removed, 'reordered': reordered}


# --------------------------------------------------------------------------
# 一键改写
# --------------------------------------------------------------------------

def reformat(doc, mode: str = 'reformat', add_front_matter: bool = False,
             title_cn: str | None = None, strict_refs: bool = True) -> dict:
    """把整份文档改写为 DUT 格式。**就地修改** ``doc``。

    :param mode: ``reformat``（只改格式，不动结构）/ ``rebuild``（重建前置部分与分节）。
    :param add_front_matter: 是否补齐「摘要 / ABSTRACT / 目录」等前置页占位。
    :param title_cn: 论文中文题目（用于**偶数页页眉**）。
    :param strict_refs: 是否重排参考文献并删除未引用条目。

    :returns: 变更统计，交给 :func:`print_report` 打印。
    """
    stats = {'total': 0, 'styled': 0, 'headings': 0, 'captions': 0,
             'refs': 0, 'page_breaks': 0, 'by_kind': {}}
    apply_style_set(doc)
    _set_page(doc.sections[0])
    _enable_even_odd_headers(doc)

    A = WD_ALIGN_PARAGRAPH
    paras = list(doc.paragraphs)
    first_h1_seen = False

    for p in paras:
        stats['total'] += 1
        kind = classify_paragraph(p)
        stats['by_kind'][kind] = stats['by_kind'].get(kind, 0) + 1

        if kind == CH_EMPTY:
            fmt_para(p, line=1.0, exact=25, before=0, after=0, snap_to_grid=False)
            _set_ind(p, firstLineChars=0, firstLine=0)
            continue

        if kind == CH_LEVEL1:
            text = re.sub(r'\s+', ' ', p.text.strip())     # 必须在删 run 之前取
            is_chapter = bool(_RE_CH1.match(text))
            for r in list(p.runs):
                r._r.getparent().remove(r._r)
            spec = HEADING_SPEC[1]
            set_run(p.add_run(text), cn=HEI, en=EN_FONT,
                    size=spec['size'], bold=spec['bold'])
            # 清掉 run 上的残留直接格式 + 段落级覆盖，并纠正样式本身
            enforce_heading_format(doc, 1, p)
            set_keep_with_next(p, True)
            if is_chapter:
                set_page_break_before(p, True)
                stats['page_breaks'] += 1
            first_h1_seen = True
            stats['headings'] += 1
            stats['styled'] += 1

        elif kind == CH_LEVEL2:
            text = re.sub(r'\s+', ' ', p.text.strip())
            for r in list(p.runs):
                r._r.getparent().remove(r._r)
            spec = HEADING_SPEC[2]
            set_run(p.add_run(text), cn=HEI, en=EN_FONT,
                    size=spec['size'], bold=spec['bold'])
            enforce_heading_format(doc, 2, p)
            set_keep_with_next(p, True)
            stats['headings'] += 1
            stats['styled'] += 1

        elif kind == CH_LEVEL3:
            text = re.sub(r'\s+', ' ', p.text.strip())
            for r in list(p.runs):
                r._r.getparent().remove(r._r)
            spec = HEADING_SPEC[3]
            set_run(p.add_run(text), cn=HEI, en=EN_FONT,
                    size=spec['size'], bold=spec['bold'])
            enforce_heading_format(doc, 3, p)
            set_keep_with_next(p, True)
            stats['headings'] += 1
            stats['styled'] += 1

        elif kind == CH_FRONT:
            raw = p.text.strip()
            key = _norm_title(raw)
            style_name, canonical = _FRONT_TITLES.get(key, ('摘要题目', raw))
            for r in list(p.runs):
                r._r.getparent().remove(r._r)
            cn = EN_FONT if key in ('ABSTRACT', 'REFERENCES', 'TABLEOFCONTENTS',
                                    'ACKNOWLEDGEMENTS') else HEI
            set_run(p.add_run(canonical), cn=cn, en=EN_FONT, size=PT_ABSTRACT,
                    bold=True)
            try:
                p.style = doc.styles[style_name]
            except KeyError:
                pass
            fmt_para(p, align=A.CENTER, line=LINE_HEADING, before=0, after=12,
                     snap_to_grid=False)
            _set_ind(p, left=0, firstLineChars=0, firstLine=0)
            set_outline(p, 0 if style_name in ('参考文献标题', '附录', '致谢', '作者简介', 'ABSTRACT') else 9)
            set_keep_with_next(p, True)
            if style_name in ('参考文献标题', '附录', '致谢', '作者简介'):
                set_page_break_before(p, True)
                stats['page_breaks'] += 1
            stats['styled'] += 1

        elif kind in (CH_CAPTION_FIG, CH_CAPTION_TAB):
            is_fig = kind == CH_CAPTION_FIG
            txt = re.sub(r'\s+', ' ', p.text.strip())
            txt = re.sub(r'^(图|表|Fig\.?|Tab\.?)\s*', lambda m: m.group(1) + ' ', txt)
            for r in list(p.runs):
                r._r.getparent().remove(r._r)
            set_run(p.add_run(txt), cn=SONG if is_fig else SONG, en=EN_FONT,
                    size=PT_CAPTION)
            fmt_para(p, align=A.CENTER, line=1.0, before=3 if is_fig else 6,
                     after=6 if is_fig else 3, snap_to_grid=False)
            _set_ind(p, firstLineChars=0, firstLine=0)
            set_keep_lines(p, True)
            if not is_fig:
                set_keep_with_next(p, True)   # 表题与表同页
            stats['captions'] += 1
            stats['styled'] += 1

        elif kind == CH_REF:
            text = re.sub(r'\s+', ' ', p.text.strip())
            for r in list(p.runs):
                r._r.getparent().remove(r._r)
            set_run(p.add_run(text), cn=SONG,
                    en=EN_FONT, size=PT_CAPTION)
            try:
                p.style = doc.styles['参考文献正文']
            except KeyError:
                pass
            fmt_para(p, align=A.JUSTIFY, line=LINE_BODY, before=0, after=0,
                     snap_to_grid=False, left_indent_pt=8.5, hanging_pt=8.5)
            stats['refs'] += 1
            stats['styled'] += 1

        elif kind == CH_KEYWORDS:
            # 「关键词：」黑体小四 + 内容宋体小四；「Key Words：」加粗 + TNR
            raw = re.sub(r'\s+', ' ', p.text.strip())
            m = _RE_KEYWORDS.match(raw)
            head, rest = (m.group(1), raw[m.end():]) if m else (raw, '')
            is_en = bool(re.match(r'\s*(Key\s*Words|Keywords)', head, re.I))
            for r in list(p.runs):
                r._r.getparent().remove(r._r)
            set_run(p.add_run(head), cn=(EN_FONT if is_en else HEI),
                    en=EN_FONT, size=PT_BODY, bold=is_en)
            if rest:
                set_run(p.add_run(rest), cn=(EN_FONT if is_en else SONG),
                        en=EN_FONT, size=PT_BODY)
            fmt_para(p, align=A.JUSTIFY, line=LINE_BODY, before=0, after=0,
                     snap_to_grid=False)
            _set_ind(p, firstLineChars=0, firstLine=0)
            stats['styled'] += 1

        elif kind == CH_BODY_EN:
            # 英文正文：Times New Roman 小四，首行缩进 2 字符，不套宋体
            for r in list(p.runs):
                set_run(r, cn=EN_FONT, en=EN_FONT, size=PT_BODY)
            fmt_para(p, align=A.JUSTIFY, line=LINE_BODY, before=0, after=0,
                     snap_to_grid=False)
            first_line_chars(p, 2)
            stats['styled'] += 1

        else:  # CH_BODY
            for r in list(p.runs):
                set_run(r, cn=SONG, en=EN_FONT, size=PT_BODY)
            fmt_para(p, align=A.JUSTIFY, line=LINE_BODY, before=0, after=0,
                     snap_to_grid=False)
            first_line_chars(p, 2)
            stats['styled'] += 1

    # 三线表
    for t in doc.tables:
        _normalize_table(t)

    # 分节页码与页眉
    _apply_sections_and_numbering(doc, title_cn=title_cn)

    if strict_refs:
        try:
            stats['references'] = order_references(doc)
        except Exception as exc:  # 引用解析失败不应中断整个改写
            stats['references'] = {'error': str(exc)}

    if add_front_matter:
        stats['front_matter'] = _stub_front_matter(doc)

    # 目录域：源文档若已有「目录」标题但没有 TOC 域，就在这里补一个
    try:
        stats['toc_field_added'] = ensure_toc_field(doc)
    except Exception as exc:
        stats['toc_field_added'] = f'error: {exc}'
    return stats


def _normalize_table(t: Table):
    """把任意表格转成 DUT 三线表：去竖线、上线/下线 1.5pt、表头线 1pt。"""
    _clear_table_borders(t)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    nrow = len(t.rows)
    for i, row in enumerate(t.rows):
        for cell in row.cells:
            for p in cell.paragraphs:
                fmt_para(p, align=WD_ALIGN_PARAGRAPH.CENTER, line=1.25,
                         before=1, after=1, snap_to_grid=False)
                _set_ind(p, firstLineChars=0, firstLine=0)
                for r in p.runs:
                    set_run(r, cn=SONG, en=EN_FONT, size=PT_CAPTION)
            top = 12 if i == 0 else None
            bottom = 8 if i == 0 else (12 if i == nrow - 1 else 0)
            _set_cell_borders(cell, top=top, bottom=bottom, left=0, right=0)
    return t


def _apply_sections_and_numbering(doc, title_cn: str | None = None):
    """为所有节套页面设置、页码、奇偶页页眉。

    规范要求**封面不编页码、前置部分罗马数字、正文阿拉伯数字从 1 起**，
    这至少需要 3 节。源文档往往只有 1 节，因此这里在缺失时自动补出分节点：

    * 第 1 节（原节）：封面 → 无页码
    * 第 2 节（新建）：摘要 / ABSTRACT / 目录 → 罗马数字从 I 起
    * 第 3 节（新建）：正文起 → 阿拉伯数字从 1 起
    """
    if len(doc.sections) == 1:
        for _ in range(2):
            add_section(doc, WD_SECTION_START.NEW_PAGE)

    first_body_seen = False
    for i, sec in enumerate(doc.sections):
        _set_page(sec)
        if i == 0:
            # 封面：无页码、**无页眉**（规范：封一、封二不编入页码）
            set_page_number_footer(sec, fmt='decimal', start=1, style='none')
            set_header_pair(sec, odd_text='', even_text='')
            continue
        if i == 1:
            # 前置部分：罗马数字从 I 起，页脚 - I -
            set_page_number_footer(sec, fmt='upperRoman', start=1, style='dash')
        elif i == 2:
            # 第二个前置节（ABSTRACT / 目录）：继续罗马数字，**不重置**
            set_page_number_footer(sec, fmt='upperRoman', start=None, style='dash')
        elif i == 3 and not first_body_seen:
            # 正文第一个节：阿拉伯数字从 1 起，页脚 - 1 -
            set_page_number_footer(sec, fmt='decimal', start=1, style='dash')
            first_body_seen = True
        else:
            # 其余正文节：继续编号，**不重置**
            set_page_number_footer(sec, fmt='decimal', start=None, style='dash')
        set_header_pair(sec, odd_text='大连理工大学硕士学位论文',
                        even_text=title_cn or '学位论文题目')
    _enable_even_odd_headers(doc)
    return doc


def ensure_toc_field(doc, levels: str = '1-3') -> bool:
    """若「目　　录」标题之后还没有目录域，就在其后插入一个 ``TOC`` 域。

    只**新增域**，不删除源文档里已有的静态目录条目——删除会造成不可逆的内容丢失。
    用户打开 Word 后按提示「更新域」即可让目录变成带页码的正式目录，
    已有的静态条目可以自行删掉。

    :returns: 是否插入了新域。
    """
    if 'TOC \\o' in doc.element.body.xml:
        return False
    for p in list(doc.paragraphs):
        if _norm_title(p.text) != '目录':
            continue
        field_p = add_toc_field(doc, levels=levels)
        p._p.addnext(field_p._p)
        return True
    return False


def _stub_front_matter(doc) -> dict:
    """在正文之前补齐「摘要 / ABSTRACT / 目录」占位页（只加结构，不编内容）。"""
    made = {}
    first = doc.paragraphs[0] if doc.paragraphs else None
    for label in ('摘　　要', 'ABSTRACT', '目　　录'):
        p = add_plain_heading(doc, label,
                              cn=EN_FONT if label == 'ABSTRACT' else HEI)
        if first is not None:
            first._p.addprevious(p._p)
        made[label] = True
        if label == '目　　录':
            fp = add_toc_field(doc)
            p._p.addnext(fp._p)
    return made


# --------------------------------------------------------------------------
# 校验
# --------------------------------------------------------------------------

def audit(doc) -> dict:
    """校验产出的 docx 是否符合 DUT 规范，返回问题清单。**交付前必做。**"""
    issues: list[str] = []     # 违规项（必须修）
    warnings: list[str] = []   # 提示项（视论文情况而定）
    info = {}

    # 页面
    secs = list(doc.sections)
    info['sections'] = len(secs)
    s0 = secs[0]
    expect = {'top_margin': 3.5, 'bottom_margin': 2.5, 'left_margin': 2.5,
              'right_margin': 2.5, 'header_distance': 2.5, 'footer_distance': 2.0}
    for attr, want in expect.items():
        got = round(getattr(s0, attr).cm, 2)
        if abs(got - want) > 0.06:
            issues.append(f'页面 {attr} = {got} cm，应为 {want} cm')
    info['page'] = {k: round(getattr(s0, k).cm, 2) for k in expect}

    # 奇偶页页眉开关
    if doc.settings.element.find(qn('w:evenAndOddHeaders')) is None:
        issues.append('缺少 w:evenAndOddHeaders：奇偶页页眉无法不同')
    info['even_odd_headers'] = doc.settings.element.find(qn('w:evenAndOddHeaders')) is not None

    # 自动更新域
    upd = doc.settings.element.find(qn('w:updateFields'))
    info['update_fields'] = (upd is not None and upd.get(qn('w:val')) == 'true')
    if not info['update_fields']:
        issues.append('未写 w:updateFields：Word 打开时目录/页码不会自动更新')

    # 目录域（正文 part）与页码域（页脚 part）——两处都在，才算完整
    body_xml = doc.element.body.xml
    info['toc_field'] = 'TOC \\o' in body_xml
    if not info['toc_field']:
        warnings.append('未找到自动目录域 TOC \\o "1-3"（用 add_toc_field() 插入；'
                        '论文确实不需要目录时可忽略）')
    info['page_field'] = False
    for sec in doc.sections:
        for part in (sec.footer, sec.even_page_footer, sec.first_page_footer):
            try:
                if 'PAGE' in part._element.xml:
                    info['page_field'] = True
            except Exception:
                continue
    if not info['page_field']:
        issues.append('所有节的页脚都没有 PAGE 域——页码不会显示')

    # 奇偶页页眉内容抽查
    for i, sec in enumerate(doc.sections):
        try:
            odd = sec.header.paragraphs[0].text.strip() if sec.header.paragraphs else ''
        except Exception:
            continue
        if odd and '大连理工大学' not in odd:
            issues.append(f'第 {i + 1} 节奇数页页眉是 {odd!r}，应为「大连理工大学硕士学位论文」')

    # 标题层级
    lv = {1: 0, 2: 0, 3: 0}
    for p in doc.paragraphs:
        sn = p.style.name if p.style is not None else ''
        if sn.startswith('Heading'):
            tail = sn.split()[-1]
            if tail.isdigit() and int(tail) in lv:
                lv[int(tail)] += 1
    info['headings'] = lv
    if lv[1] == 0:
        issues.append('没有任何「章标题」（Heading 1）——目录会抓不到内容')

    # 标题的 outline level：段落层或**样式层**任一处有即可。
    # 官方模板把 outlineLvl 放在样式上（heading 1/2/3 的 pPr 里），
    # 所以只查段落层会误报。
    missing_outline = 0
    for p in doc.paragraphs:
        sn = p.style.name if p.style is not None else ''
        if not sn.startswith('Heading') or not p.text.strip():
            continue
        ppr = p._p.find(qn('w:pPr'))
        in_para = ppr is not None and ppr.find(qn('w:outlineLvl')) is not None
        in_style = False
        try:
            sppr = p.style.element.find(qn('w:pPr'))
            in_style = sppr is not None and sppr.find(qn('w:outlineLvl')) is not None
        except Exception:
            pass
        if not (in_para or in_style):
            missing_outline += 1
    if missing_outline:
        issues.append(f'{missing_outline} 个标题缺少 outlineLvl（段落层或样式层都没有，目录会抓不到）')

    # 标题的**有效字符格式**：字号/字重/字体/颜色/对齐/大纲级别是否符合规范。
    # 这是应对「Word 样式面板被改过」与「run 上残留直接格式」的关键检查——
    # 只看有没有套样式名是不够的。
    head_problems = heading_format_problems(doc)
    info['heading_format_problems'] = head_problems
    if head_problems:
        issues.append(f'{len(head_problems)} 处标题格式不符规范：'
                      + '；'.join(head_problems[:3]))

    # 正文字体与缩进抽查——只看「真正的陈述性段落」。
    # 空白骨架里的占位小标题（如「摘　　要」下还没写内容）不应被算作正文，
    # 因此用 _is_prose() 做二次过滤，避免误报。
    body_bad_font, body_bad_indent, body_grid, body_checked = 0, 0, 0, 0
    for p in doc.paragraphs:
        # 只检查**中文**陈述性正文：英文摘要用 Times New Roman，
        # 「关键词：」行按规范本就不缩进，都不能按中文正文的规则去要求。
        if classify_paragraph(p) != CH_BODY or not _is_prose(p.text):
            continue
        body_checked += 1
        ppr = p._p.find(qn('w:pPr'))
        snap = ppr.find(qn('w:snapToGrid')) if ppr is not None else None
        if snap is None or snap.get(qn('w:val')) != '0':
            body_grid += 1
        ind = ppr.find(qn('w:ind')) if ppr is not None else None
        if ind is None or ind.get(qn('w:firstLineChars')) != '200':
            body_bad_indent += 1
        for r in p.runs:
            rpr = r._element.find(qn('w:rPr'))
            rf = rpr.find(qn('w:rFonts')) if rpr is not None else None
            if rf is None or rf.get(qn('w:eastAsia')) != SONG:
                body_bad_font += 1
                break
    info['body_checked'] = body_checked
    if body_grid:
        issues.append(f'{body_grid} 个正文段未取消「对齐到网格」(w:snapToGrid val="0")')
    if body_bad_indent:
        issues.append(f'{body_bad_indent} 个正文段首行缩进不是 2 字符 (firstLineChars="200")')
    if body_bad_font:
        issues.append(f'{body_bad_font} 个正文段的中文字体不是宋体 (w:eastAsia="宋体")')

    # 文字颜色：正文/标题/图表题都应为黑色。
    # python-docx 内置模板的 Heading 样式自带主题蓝（365F91 / 4F81BD），
    # 不清掉的话渲染出来标题就是蓝的——所以这里必须查。
    bad_colors: dict[str, str] = {}
    for name in ('Normal', 'Heading 1', 'Heading 2', 'Heading 3',
                 '摘要题目', 'ABSTRACT', '图名中文', '图名英文', '参考文献正文',
                 '附录', '致谢', '作者简介'):
        try:
            st = doc.styles[name]
        except KeyError:
            continue
        c = st.element.find('.//' + qn('w:color'))
        if c is not None:
            v = (c.get(qn('w:val')) or '').upper()
            if v not in ('', 'AUTO', '000000'):
                bad_colors[f'样式 {name}'] = '#' + v
    for p in doc.paragraphs:
        for r in p.runs:
            rpr = r._element.find(qn('w:rPr'))
            c = rpr.find(qn('w:color')) if rpr is not None else None
            if c is None:
                continue
            v = (c.get(qn('w:val')) or '').upper()
            if v not in ('', 'AUTO', '000000'):
                kind = classify_paragraph(p)
                bad_colors[f'{kind} 段落 {p.text[:12]!r}'] = '#' + v
    info['nonzero_colors'] = bad_colors
    if bad_colors:
        sample = '; '.join(f'{k}={v}' for k, v in list(bad_colors.items())[:4])
        issues.append(f'{len(bad_colors)} 处文字不是黑色（应为 #000000）：{sample}')

    # 三线表
    for i, t in enumerate(doc.tables):
        tblPr = t._tbl.tblPr
        b = tblPr.find(qn('w:tblBorders'))
        if b is not None:
            for edge in ('left', 'right', 'insideV'):
                el = b.find(qn('w:' + edge))
                if el is not None and el.get(qn('w:val')) != 'none':
                    issues.append(f'表 {i + 1} 有竖线/内框，不是三线表')
                    break
    info['tables'] = len(doc.tables)

    info['issues'] = issues
    info['warnings'] = warnings
    info['ok'] = not issues
    return info


def print_report(result: dict, title: str = 'DUT 格式报告') -> str:
    """把 :func:`reformat` / :func:`audit` 的返回值打印成可读报告，并返回文本。"""
    lines = [f'===== {title} =====']
    if 'issues' in result:
        if result['issues']:
            lines.append(f'发现 {len(result["issues"])} 个问题：')
            lines += [f'  ✗ {x}' for x in result['issues']]
        else:
            lines.append('  ✓ 全部检查通过')
        for w in result.get('warnings') or []:
            lines.append(f'  ! 提示：{w}')
        for k in ('sections', 'page', 'even_odd_headers', 'update_fields',
                  'toc_field', 'page_field', 'headings', 'tables',
                  'nonzero_colors', 'heading_format_problems'):
            if k in result:
                lines.append(f'  · {k}: {result[k]}')
    else:
        lines.append(f"  段落总数 {result.get('total')}，已套样式 {result.get('styled')}")
        lines.append(f"  标题 {result.get('headings')}，图表题 {result.get('captions')}，"
                     f"参考文献 {result.get('refs')}，分页 {result.get('page_breaks')}")
        lines.append(f"  分类明细: {result.get('by_kind')}")
        refs = result.get('references')
        if refs:
            lines.append(f"  参考文献重排: kept={refs.get('kept')} "
                         f"removed={len(refs.get('removed') or [])} "
                         f"reordered={refs.get('reordered')}")
    text = '\n'.join(lines)
    print(text)
    return text


def norm_ref_spacing(entry: str) -> str:
    """规范化参考文献条目的空格与标点（不改序号）。"""
    return format_reference_entry(entry, 0).replace('[0] ', '', 1)
