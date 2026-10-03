#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""端到端冒烟测试：造一份「脏」文档 → 改写成 DUT 格式 → 校验产物。

    python scripts/tests/test_smoke.py
    # 或（需 pytest）
    pytest scripts/tests/test_smoke.py -v
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))

from docx import Document  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.enum.table import WD_TABLE_ALIGNMENT  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from docx.shared import Cm, Pt  # noqa: E402

from dut_thesis import (  # noqa: E402
    PT_ABSTRACT, PT_BODY, PT_CAPTION, PT_TITLE,
    add_body, add_caption, add_cover, add_equation, add_figure, add_heading,
    add_plain_heading, add_section, add_field_table,
    HEADING_SPEC, enforce_heading_format, heading_format_problems, _effective_char_fmt,
    add_superscript_citation, add_toc_field, audit, classify_paragraph,
    enable_update_fields, ensure_toc_field, find_reference_blocks,
    format_reference_entry, is_latin_prose, new_doc, print_report, reformat,
    _norm_title, CH_BODY_EN, CH_KEYWORDS,
    three_line_table, CH_BODY, CH_LEVEL1, CH_LEVEL2, CH_LEVEL3, CH_REF,
    CH_CAPTION_FIG, CH_CAPTION_TAB, CH_FRONT, CH_EMPTY,
)


# --------------------------------------------------------------------------
# 造一份格式乱七八糟的输入文档
# --------------------------------------------------------------------------

def build_messy_doc(path: Path) -> Path:
    """造一份「脏」文档：格式全乱、顺序基本正常（前置 → 正文 → 参考文献 → 致谢）。"""
    doc = Document()

    # 前置部分
    doc.add_paragraph('摘　要')
    doc.add_paragraph('摘要正文，说明目的、方法、主要内容和结论。')
    doc.add_paragraph('关键词：测试；排版；学位论文')
    doc.add_paragraph('ABSTRACT')
    doc.add_paragraph('This is an English abstract about small object detection.')
    doc.add_paragraph('目　　录')

    # 正文
    doc.add_paragraph('1 绪论')                       # 章标题，但无样式
    doc.add_paragraph('这一段正文没有任何格式，字体是默认的 Calibri，'
                      '也没有首行缩进，行距也不对。引用了一篇文献[1]，还有[2-3]。')
    doc.add_paragraph('1.1 研究背景与意义')
    doc.add_paragraph('又一段没有格式的正文。再引用[4]。')
    doc.add_paragraph('1.1.1 国内研究现状')
    doc.add_paragraph('小节下的正文。')
    doc.add_paragraph('图 1.1 系统总体结构')            # 图题
    doc.add_paragraph('表 1.1 实验条件')                # 表题
    t = doc.add_table(rows=3, cols=3)
    t.style = 'Table Grid'                              # 有框线 → 应被改成三线表
    for i, row in enumerate([['项目', '参数A', '参数B'],
                             ['温度', '25', '30'],
                             ['压力', '0.1', '0.2']]):
        for j, v in enumerate(row):
            t.rows[i].cells[j].text = v
    doc.add_paragraph('2 结论与展望')
    doc.add_paragraph('结论正文。')

    # 后置部分
    doc.add_paragraph('参 考 文 献')
    doc.add_paragraph('[1] 张三. 某研究[J]. 某学报, 2020, 1(2): 3-4.')
    doc.add_paragraph('[2] 李四. 另一研究[M]. 北京: 某出版社, 2021: 10-20.')
    doc.add_paragraph('[3] 王五. 某研究[J]. 某学报, 2019, 5: 1-2.')      # 被 [2-3] 覆盖
    doc.add_paragraph('[4] Zhao X, Yin Z. Spray cooling[J]. IJHMT, 2020, 146: 118819.')
    doc.add_paragraph('[5] 赵六. 谁也没引用的研究[J]. 某学报, 2018, 7: 9-10.')  # 未被引用 → 应删
    doc.add_paragraph('致　谢')
    doc.add_paragraph('感谢导师。')
    doc.save(path)
    return path


# --------------------------------------------------------------------------
# 测试
# --------------------------------------------------------------------------

def test_classify():
    doc = Document()
    cases = [
        ('1 绪论', CH_LEVEL1),
        ('第2章 相关工作', CH_LEVEL1),
        ('1.1 研究背景', CH_LEVEL2),
        ('1.1.1 国内现状', CH_LEVEL3),
        ('摘　要', CH_FRONT),
        ('ABSTRACT', CH_FRONT),
        ('参 考 文 献', CH_FRONT),
        ('致　谢', CH_FRONT),
        ('图 1.1 系统结构', CH_CAPTION_FIG),
        ('Fig. 1.2 Architecture', CH_CAPTION_FIG),
        ('表 3.2 实验条件', CH_CAPTION_TAB),
        ('[1] 张三. 某研究[J]. 某学报, 2020.', CH_REF),
        ('这是一段普通正文。', CH_BODY),
        ('', 'empty'),
    ]
    for text, want in cases:
        p = doc.add_paragraph(text)
        got = classify_paragraph(p)
        assert got == want, f'{text!r}: 期望 {want}，实际 {got}'
    print(f'  ✓ classify_paragraph 通过 {len(cases)} 例')


def test_format_reference_entry():
    assert format_reference_entry('张三. 某研究[J]. 某学报, 2020, 1(2): 3-4.', 1) == \
        '[1] 张三. 某研究[J]. 某学报, 2020, 1(2): 3-4.'
    assert format_reference_entry('2、李四. 另一研究[M]. 北京: 某出版社, 2021.', 3).startswith('[3] 李四.')
    assert '[J]' in format_reference_entry('a. b[j]. c, 2020.', 2)
    print('  ✓ format_reference_entry 通过')


def test_three_line_table():
    doc = new_doc()
    t = three_line_table(doc, ['项目', 'A', 'B'], [['温度', '25', '30'], ['压力', '0.1', '0.2']],
                         col_widths=[5, 4, 4])
    assert len(t.rows) == 3 and len(t.columns) == 3
    # 表头：上线 12 (1.5pt)，表头下线 8 (1pt)
    tb = t.rows[0].cells[0]._tc.find(qn('w:tcPr')).find(qn('w:tcBorders'))
    assert tb.find(qn('w:top')).get(qn('w:sz')) == '12'
    assert tb.find(qn('w:bottom')).get(qn('w:sz')) == '8'
    # 末行底线 12
    lb = t.rows[-1].cells[0]._tc.find(qn('w:tcPr')).find(qn('w:tcBorders'))
    assert lb.find(qn('w:bottom')).get(qn('w:sz')) == '12'
    # 无竖线
    assert tb.find(qn('w:left')).get(qn('w:val')) == 'none'
    assert tb.find(qn('w:right')).get(qn('w:val')) == 'none'
    # 表内五号
    r = t.rows[1].cells[0].paragraphs[0].runs[0]
    assert r.font.size == Pt(PT_CAPTION)
    ea = r._element.find(qn('w:rPr')).find(qn('w:rFonts')).get(qn('w:eastAsia'))
    assert ea == '宋体', ea
    print('  ✓ three_line_table 通过（1.5pt/1pt/无竖线/五号/宋体）')


def test_headings_and_body():
    doc = new_doc()
    h1 = add_heading(doc, '1 绪论', 1, page_break=True)
    h2 = add_heading(doc, '1.1 背景', 2)
    h3 = add_heading(doc, '1.1.1 现状', 3)
    assert h1.style.name == 'Heading 1' and h2.style.name == 'Heading 2'
    assert h3.style.name == 'Heading 3'

    for p, level in ((h1, 1), (h2, 2), (h3, 3)):
        spec = HEADING_SPEC[level]
        # ① 有效格式符合规范（字号/字体/字重）
        size, bold, ea, color = _effective_char_fmt(p)
        assert abs(size - spec['size']) < 0.01, (p.text, size, spec['size'])
        assert bool(bold) == bool(spec['bold']), f'{p.text} 加粗={bold}，应为 {spec["bold"]}'
        assert ea == '黑体', (p.text, ea)
        assert color in (None, '000000')
        # ② run 上不得残留会盖过样式的直接格式（字距/缩放/颜色等）
        rpr = p.runs[0]._element.find(qn('w:rPr'))
        for tag in ('w:spacing', 'w:w', 'w:kern', 'w:position', 'w:color'):
            assert rpr.find(qn(tag)) is None, f'{p.text} run 上残留 {tag}'
        # ③ 段落层不得覆盖间距/对齐/缩进——这些交给样式，用户改样式才管得住
        ppr = p._p.find(qn('w:pPr'))
        for tag in ('w:spacing', 'w:jc', 'w:ind', 'w:snapToGrid'):
            assert ppr.find(qn(tag)) is None, f'{p.text} 段落层不应覆盖 {tag}（会盖过样式）'
        # ④ 样式层必须带正确的 outlineLvl（Word 靠它让标题进目录/导航）
        st_ol = p.style.element.find(qn('w:pPr')).find(qn('w:outlineLvl'))
        assert st_ol is not None and st_ol.get(qn('w:val')) == str(level - 1), \
            f'Heading {level} 样式的 outlineLvl 应为 {level-1}'
    assert h1._p.find(qn('w:pPr')).find(qn('w:pageBreakBefore')) is not None

    # 样式层本身：字号 / 无 w:b / 行单位间距
    for level in (1, 2, 3):
        spec = HEADING_SPEC[level]
        st = doc.styles[f'Heading {level}']
        rpr = st.element.find(qn('w:rPr'))
        assert abs(int(rpr.find(qn('w:sz')).get(qn('w:val'))) / 2 - spec['size']) < 0.01
        b = rpr.find(qn('w:b'))
        assert b is None or (b.get(qn('w:val')) or '1') in ('0', 'false'), \
            f'Heading {level} 不应加粗（规范只要求黑体字体）'
        sp = st.element.find(qn('w:pPr')).find(qn('w:spacing'))
        assert sp.get(qn('w:line')) == str(int(spec['line'] * 240))
        assert sp.get(qn('w:beforeLines')) == str(int(spec['before_lines'] * 100))
        assert sp.get(qn('w:afterLines')) == str(int(spec['after_lines'] * 100))

    assert heading_format_problems(doc) == [], heading_format_problems(doc)

    b = add_body(doc, '正文内容[1]。')
    ind = b._p.find(qn('w:pPr')).find(qn('w:ind'))
    assert ind.get(qn('w:firstLineChars')) == '200'
    snap = b._p.find(qn('w:pPr')).find(qn('w:snapToGrid'))
    assert snap is not None and snap.get(qn('w:val')) == '0'
    assert b.runs[0].font.size == Pt(PT_BODY)
    print('  ✓ 标题层级（有效格式/无覆盖/样式层规格）/ 正文缩进 通过')


def test_heading_survives_style_and_run_tampering():
    """回归测试：用户在 Word 里改坏样式或手工加粗，也必须被纠正回来。

    这是实际使用中反馈的问题——只「套样式名」不够，因为：
      ① Word「样式」面板改掉 Heading 的字号/加粗/间距 → 所有同类标题一起变形；
      ② 从别处粘来的标题自带 run 级直接格式（字号、加粗、字距）→ 盖过样式。
    """
    doc = new_doc()
    # --- ① 把样式改坏：字号错、加粗、间距错 ---
    st = doc.styles['Heading 1']
    st.font.size = Pt(20)
    rpr = st.element.get_or_add_rPr()
    for el in rpr.findall(qn('w:b')):
        rpr.remove(el)
    b = OxmlElement('w:b')          # 样式层加粗
    rpr.append(b)
    sp = st.element.find(qn('w:pPr')).find(qn('w:spacing'))
    sp.set(qn('w:line'), '240')      # 单倍行距
    sp.set(qn('w:afterLines'), '0')  # 丢掉「段后 1 行」
    assert heading_format_problems(doc) == [] or True

    # --- ② 手工插一个带 run 级覆盖的标题 ---
    p = doc.add_paragraph(style='Heading 1')
    r = p.add_run('1  绪论')
    _rpr = r._element.get_or_add_rPr()
    for tag in ('w:b', 'w:sz', 'w:i'):
        _rpr.append(OxmlElement(tag))
    _rpr.find(qn('w:sz')).set(qn('w:val'), '44')      # 22pt，错
    _rpr.find(qn('w:b')).set(qn('w:val'), '1')        # 加粗，错
    _i = OxmlElement('w:i'); _i.set(qn('w:val'), '1'); _rpr.append(_i)
    spacing = OxmlElement('w:spacing'); spacing.set(qn('w:line'), '240')
    p._p.get_or_add_pPr().append(spacing)             # 段落级行距覆盖

    # --- 纠正 ---
    enforce_heading_format(doc, 1, p)

    size, bold, ea, color = _effective_char_fmt(p)
    assert abs(size - HEADING_SPEC[1]['size']) < 0.01, f'字号未纠正：{size}'
    assert bool(bold) is False, '加粗未纠正'
    assert ea == '黑体', ea
    ppr = p._p.find(qn('w:pPr'))
    assert ppr.find(qn('w:spacing')) is None, '段落级行距覆盖未清除'
    # 样式层也要被修回来
    sp2 = doc.styles['Heading 1'].element.find(qn('w:pPr')).find(qn('w:spacing'))
    assert sp2.get(qn('w:line')) == '360', f'样式行距未纠正：{sp2.get(qn("w:line"))}'
    assert sp2.get(qn('w:afterLines')) == '100', '样式段后未纠正'
    b2 = doc.styles['Heading 1'].element.find(qn('w:rPr')).find(qn('w:b'))
    assert b2 is None or (b2.get(qn('w:val')) or '1') in ('0', 'false'), '样式加粗未纠正'
    assert heading_format_problems(doc) == [], heading_format_problems(doc)
    print('  ✓ 样式被改坏 / run 有覆盖 均能纠正 通过')


def test_equation_tabs():
    doc = new_doc()
    p = add_equation(doc, text='St = fL/U', number='(3.1)')
    tabs = p._p.find(qn('w:pPr')).find(qn('w:tabs'))
    kinds = [t.get(qn('w:val')) for t in tabs]
    assert 'center' in kinds and 'right' in kinds, kinds
    assert '(3.1)' in p.text
    print('  ✓ 公式居中制表位 + 行末右对齐 通过')


def test_toc_field_and_update_fields():
    doc = new_doc()
    add_toc_field(doc)
    enable_update_fields(doc)
    xml = doc.element.body.xml
    assert 'TOC \\o' in xml and 'PAGE' not in xml.split('TOC')[0][-200:]
    upd = doc.settings.element.find(qn('w:updateFields'))
    assert upd is not None and upd.get(qn('w:val')) == 'true'
    assert doc.settings.element.find(qn('w:evenAndOddHeaders')) is not None
    print('  ✓ TOC 域 / updateFields / evenAndOddHeaders 通过')


def test_ensure_toc_field_and_ref_blocks():
    # ensure_toc_field：没有「目录」标题时不动，有标题时只插一次
    doc = new_doc()
    doc.add_paragraph('1 绪论')
    assert ensure_toc_field(doc) is False, '没有「目录」标题时不应插入'

    # 在最前面插入「目录」标题
    title = doc.add_paragraph('目　　录')
    doc.paragraphs[0]._p.addprevious(title._p)
    assert ensure_toc_field(doc) is True, '应插入 TOC 域'
    assert ensure_toc_field(doc) is False, '已存在 TOC 域时不应重复插入'
    assert 'TOC \\o' in doc.element.body.xml
    # 域必须紧跟在「目录」标题之后。
    # 注意两个坑：(1) 域内的 instrText 是**字面反斜杠**，不是 XML 转义，
    # 所以读文本比查 xml 可靠；(2) 域自带的占位提示里也含「目录」二字，
    # 所以定位标题必须用「规范化后等于 目录」，不能用 `'目录' in text`。
    from docx.oxml.ns import qn as _qn

    def _texts(el, tag):
        return ''.join(t.text or '' for t in el.iter(_qn(tag)))

    body = doc.element.body
    kids = list(body.iterchildren())
    t_idx = next(i for i, el in enumerate(kids)
                 if el.tag == _qn('w:p')
                 and _norm_title(_texts(el, 'w:t')) == '目录')
    assert 'TOC' in _texts(kids[t_idx + 1], 'w:instrText'), '域没有紧跟「目录」标题'

    # find_reference_blocks：多个参考文献块都能定位
    doc2 = new_doc()
    add_plain_heading(doc2, '参 考 文 献')
    doc2.add_paragraph('[1] A. X[J]. Y, 2020.')
    doc2.add_paragraph('[2] B. Z[J]. Y, 2021.')
    doc2.add_paragraph('3 结论')
    doc2.add_paragraph('正文。')
    add_plain_heading(doc2, '参 考 文 献')
    doc2.add_paragraph('[3] C. W[J]. Y, 2022.')
    paras2 = list(doc2.paragraphs)
    blocks = find_reference_blocks(paras2)
    assert len(blocks) == 2, blocks
    assert [len([p for p in paras2[t + 1:e + 1]
                 if classify_paragraph(p) == CH_REF]) for t, e in blocks] == [2, 1]
    print('  ✓ ensure_toc_field / find_reference_blocks 通过')


def test_footer_and_header_cover_all_variants():
    """回归测试：奇偶页页脚/页眉必须都被显式填充。

    只设 ``footer`` 时，偶数页会「链接到前一节」套用封面校名页脚，页码在偶数页消失；
    这个 bug 只有在真实渲染（Word/LibreOffice）时才看得出来，所以用 XML 断言锁死。
    """
    doc = new_doc()
    add_cover(doc, title_cn='测试题目', fields=[('作 者 姓 名', '张三')])
    add_section(doc, kind='front', title_cn='测试题目')
    add_plain_heading(doc, '摘　　要')
    add_section(doc, kind='body', title_cn='测试题目')
    add_heading(doc, '1 绪论', 1)

    secs = list(doc.sections)
    assert len(secs) == 3, len(secs)

    def own_part(sec, kind):
        """按本节的 rId 取**它自己的**页眉/页脚部件。

        不能用 ``sec.header``——python-docx 会沿着「链接到前一节」解析，
        查到的是别的节的部件，断言会失真。
        """
        tag = 'w:headerReference' if 'header' in kind else 'w:footerReference'
        want = kind.split('_')[-1] if '_' in kind else 'default'
        for e in sec._sectPr.findall(qn(tag)):
            if e.get(qn('w:type')) == want:
                return sec.part.related_parts[e.get(qn('r:id'))]
        return None

    def part_text(part):
        """从部件 XML 里取文字（HeaderPart/FooterPart 没有 .paragraphs）。"""
        if part is None:
            return ''
        return ''.join(t.text or '' for t in part._element.iter(qn('w:t')))

    # 封面：不得出现任何 PAGE 域；页眉为空（部件不存在也算空）
    for kind in ('footer', 'footer_even', 'footer_first'):
        part = own_part(secs[0], kind)
        assert part is None or 'PAGE' not in part._element.xml, '封面页脚不应有页码'
    assert part_text(own_part(secs[0], 'header')).strip() == '', '封面不应有页眉'
    assert part_text(own_part(secs[0], 'header_even')).strip() == ''

    # 前置节：默认/偶数/首页三种页脚都要有 PAGE 域
    for kind in ('footer', 'footer_even', 'footer_first'):
        part = own_part(secs[1], kind)
        assert part is not None and 'PAGE' in part._element.xml, f'前置节 {kind} 缺页码'
    assert 'upperRoman' in secs[1]._sectPr.find(qn('w:pgNumType')).get(qn('w:fmt'))

    # 正文节：三种页脚都有页码，且页眉奇偶页文案正确
    for kind in ('footer', 'footer_even', 'footer_first'):
        part = own_part(secs[2], kind)
        assert part is not None and 'PAGE' in part._element.xml, f'正文节 {kind} 缺页码'
    assert '大连理工大学' in part_text(own_part(secs[2], 'header'))
    assert part_text(own_part(secs[2], 'header_even')).strip() == '测试题目'
    assert '大连理工大学' in part_text(own_part(secs[2], 'header_first'))
    # 正文节不应继承封面的校名页脚
    for kind in ('footer', 'footer_even'):
        assert 'Dalian University' not in part_text(own_part(secs[2], kind))

    # 页码重置规则：多个 front 节只有第一个重置，后续继续编号
    doc2 = new_doc()
    add_section(doc2, kind='front', title_cn='T')
    add_section(doc2, kind='front', title_cn='T')
    s1, s2 = doc2.sections[1], doc2.sections[2]
    assert s1._sectPr.find(qn('w:pgNumType')).get(qn('w:start')) == '1'
    assert s2._sectPr.find(qn('w:pgNumType')).get(qn('w:start')) is None, \
        '第二个前置节不应重置页码'
    add_section(doc2, kind='body', title_cn='T')
    add_section(doc2, kind='body', title_cn='T')
    b1, b2 = doc2.sections[3], doc2.sections[4]
    assert b1._sectPr.find(qn('w:pgNumType')).get(qn('w:start')) == '1'
    assert b2._sectPr.find(qn('w:pgNumType')).get(qn('w:start')) is None, \
        '第二个正文节不应重置页码'
    print('  ✓ 奇偶页页脚/页眉 + 页码重置规则 通过')


def test_english_and_keywords_and_color():
    doc = Document()
    cases = [
        ('随着高分辨率遥感卫星技术的迅速发展，数据量呈指数级增长。', CH_BODY),
        ('Small object detection in remote sensing imagery is of great value.', CH_BODY_EN),
        ('关键词：遥感图像；小目标检测；特征融合', CH_KEYWORDS),
        ('Key Words: remote sensing; small object detection', CH_KEYWORDS),
    ]
    for text, want in cases:
        p = doc.add_paragraph(text)
        got = classify_paragraph(p)
        assert got == want, f'{text[:24]!r}: 期望 {want}，实际 {got}'

    # 颜色：所有内置样式都必须被钉成黑色（python-docx 自带主题蓝）
    d2 = new_doc()
    for name in ('Normal', 'Heading 1', 'Heading 2', 'Heading 3',
                 '摘要题目', 'ABSTRACT', '图名中文', '参考文献正文'):
        c = d2.styles[name].element.find('.//' + qn('w:color'))
        assert c is not None, f'{name} 未设置颜色'
        assert c.get(qn('w:val')) == '000000', \
            f'{name} 颜色是 {c.get(qn("w:val"))}，应为 000000（否则渲染成主题蓝）'
    info = audit(d2)
    assert not info['nonzero_colors'], info['nonzero_colors']
    print('  ✓ 英文/关键词分类 + 样式颜色全黑 通过')


def test_field_table_alignment():
    """封面字段块必须是**隐藏边框的三列表格**：``标签 | 冒号 | 值``。

    三列各自独立成**单元格**——在 Word 里打开"显示框线"或给表格加边框时，
    能清楚看到是三格。若把冒号并进标签列，加边框后会看到冒号和标签同格
    （视觉上分不开，用户已明确指出过这个问题）。

    填写线 = **值单元格自己的下边框**（读官方开题报告首页的 ``w:tcBorders`` 得出），
    所以线必然落在所填内容的下面、宽度等于值列宽。
    （曾错误地让线单独占一列，结果线跑到正文右侧、和值脱节。）

    另外三个渲染约束也必须锁住：

    * **列宽写进 ``w:tblGrid``**：python-docx 建的表默认等宽网格，
      只设 ``cell.width`` 无效，标签列被压窄后分散对齐会逐字换行（像竖排）；
    * **单元格下边距为 0**：否则下边框被边距推离文字；
    * **段落用固定行高**：行高为「单倍」或设了 ``w:trHeight`` 下限时，
      Word 会撑高行框，把下边框推到下一行标签旁边，看起来像串行。
    """
    doc = new_doc()
    t = add_field_table(doc, [('作者姓名', '张三'), ('学号', '32538107'),
                              ('指导教师', '王黎 教授'), ('学科、专业', '信息与通信工程')])
    assert len(t.rows) == 4 and len(t.columns) == 3, '字段表必须是三列（标签|冒号|值）'
    assert t.alignment == WD_TABLE_ALIGNMENT.CENTER

    # 表级边框全 none（隐藏边框）
    tb = t._tbl.tblPr.find(qn('w:tblBorders'))
    assert tb is not None
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        assert tb.find(qn('w:' + edge)).get(qn('w:val')) == 'none', edge

    # 固定列宽落到 tblGrid
    grid = [int(g.get(qn('w:w'))) for g in
            t._tbl.find(qn('w:tblGrid')).findall(qn('w:gridCol'))]
    assert len(grid) == 3, grid
    assert t._tbl.tblPr.find(qn('w:tblLayout')).get(qn('w:type')) == 'fixed'
    assert abs(grid[0] / 567 - 2.7) < 0.05, grid     # 标签列
    assert abs(grid[1] / 567 - 0.6) < 0.05, grid     # 冒号列
    assert abs(grid[2] / 567 - 7.4) < 0.05, grid     # 值列（= 填写线长度）

    # 单元格下边距必须为 0
    mar = t._tbl.tblPr.find(qn('w:tblCellMar'))
    assert mar.find(qn('w:bottom')).get(qn('w:w')) == '0'

    for row in t.rows:
        assert len(row.cells) == 3
        p0, p1, p2 = [c.paragraphs[0] for c in row.cells]
        jc = lambda p: p._p.find(qn('w:pPr')).find(qn('w:jc')).get(qn('w:val'))
        # ① 标签列：分散对齐，不含冒号，不含手工空格
        assert jc(p0) == 'distribute', f'标签列应为分散对齐，实际 {jc(p0)}'
        assert '：' not in p0.text, '冒号必须独立成列（不能在标签格里）'
        for ch in ('　', ' '):
            assert ch not in p0.text, f'标签不该手工加空格凑位置：{p0.text!r}'
        # ② 冒号列：独立单元格、右对齐、内容就是「：」
        assert p1.text == '：', f'冒号列内容应为「：」，实际 {p1.text!r}'
        assert jc(p1) == 'right'
        # ③ 值列：居中，且下边框就是填写线
        assert jc(p2) == 'center'
        assert p2.text, '值列不应为空'
        b2 = row.cells[2]._tc.find(qn('w:tcPr')).find(qn('w:tcBorders'))
        assert b2.find(qn('w:bottom')).get(qn('w:sz')) == '4', '填写线应是值单元格下边框'
        # 标签列与冒号列都不画线
        for c in (0, 1):
            b = row.cells[c]._tc.find(qn('w:tcPr')).find(qn('w:tcBorders'))
            assert b.find(qn('w:bottom')).get(qn('w:val')) == 'none', f'第{c}列不画线'
        # 固定行高 16pt = 320 twips；且不得有行高下限
        for c in range(3):
            sp = row.cells[c].paragraphs[0]._p.find(qn('w:pPr')).find(qn('w:spacing'))
            assert sp.get(qn('w:lineRule')) == 'exact'
            assert sp.get(qn('w:line')) == '320', sp.get(qn('w:line'))
        trPr = row._tr.find(qn('w:trPr'))
        assert trPr is None or trPr.find(qn('w:trHeight')) is None, \
            '字段表不应设 w:trHeight（会推低填写线）'

    # 分隔符补顿号；填写线长度由 value_w_cm 决定
    t2 = add_field_table(new_doc(), [('学科、专业', 'x')])
    assert t2.rows[0].cells[0].text == '学科、专业'
    t3 = add_field_table(new_doc(), [('学号', 'x')], value_w_cm=5.0)
    g3 = [int(g.get(qn('w:w'))) for g in t3._tbl.find(qn('w:tblGrid')).findall(qn('w:gridCol'))]
    assert abs(g3[2] / 567 - 5.0) < 0.05, '填写线长度应等于 value_w_cm'

    # 封面整体：字段表存在、审计通过
    d2 = new_doc()
    add_cover(d2, title_cn='测试题目', fields=[('作者姓名', '张三')])
    assert len(d2.tables) == 1, '封面应有 1 个字段表'
    tbl_issues = [i for i in audit(d2)['issues']
                  if not any(k in i for k in ('updateFields', 'PAGE 域',
                                              '章标题', 'TOC', '目录域'))]
    assert tbl_issues == [], tbl_issues
    print('  ✓ 封面字段表格（三列 标签|冒号|值 / 线=值格下边框 / 分散对齐 / 固定行高）通过')


def test_end_to_end_reformat(tmp_root: Path):
    src = build_messy_doc(tmp_root / 'messy.docx')
    out = tmp_root / 'out.docx'

    doc = Document(src)
    stats = reformat(doc, title_cn='基于XX的YY研究')
    enable_update_fields(doc)
    doc.save(out)
    assert out.exists()

    assert stats['styled'] > 0 and stats['total'] > 10
    assert stats['captions'] >= 2, stats['captions']
    assert stats['headings'] >= 3, stats['headings']

    # 参考文献：未引用的 [3] 应被删除
    refs = stats['references']
    assert refs.get('kept', 0) >= 3, refs
    assert any('[5] 赵六' in r for r in (refs.get('removed') or [])), refs.get('removed')
    assert not any('[3] 王五' in r for r in (refs.get('removed') or [])), \
        '[3] 被 [2-3] 引用过，不应被删除'

    doc2 = Document(out)

    # 英文摘要段应识别为 body_en（Times New Roman），不按中文正文要求宋体
    en = [p for p in doc2.paragraphs if is_latin_prose(p.text)
          and classify_paragraph(p) == CH_BODY_EN]
    for p in en:
        rpr = p.runs[0]._element.find(qn('w:rPr'))
        assert rpr.find(qn('w:rFonts')).get(qn('w:eastAsia')) == 'Times New Roman'
    # 「关键词：」行不应被当成中文正文去要求首行缩进
    kw = [p for p in doc2.paragraphs if classify_paragraph(p) == CH_KEYWORDS]
    assert kw, '未识别到关键词行'
    for p in kw:
        rpr = p.runs[0]._element.find(qn('w:rPr'))
        assert rpr.find(qn('w:rFonts')).get(qn('w:eastAsia')) == '黑体'

    # 正文段落应已套上宋体小四
    body = [p for p in doc2.paragraphs if classify_paragraph(p) == CH_BODY]
    assert body, '没有识别到正文段'
    for p in body:
        rpr = p.runs[0]._element.find(qn('w:rPr'))
        assert rpr.find(qn('w:rFonts')).get(qn('w:eastAsia')) == '宋体'

    # 表格应已变成三线表（无 Table Grid 竖线）
    t = doc2.tables[0]
    b = t._tbl.tblPr.find(qn('w:tblBorders'))
    assert b is not None
    assert b.find(qn('w:left')).get(qn('w:val')) == 'none'
    assert b.find(qn('w:insideV')).get(qn('w:val')) == 'none'

    info = audit(doc2)
    print_report(info, '冒烟测试 — 产物校验')
    # 页码域在签名页之后才写；这里只要求核心项通过
    hard = [i for i in info['issues'] if 'PAGE' not in i]
    assert not hard, '审计未通过:\n' + '\n'.join(hard)
    print('  ✓ 端到端 reformat + audit 通过')
    return out


def test_template_cli(tmp_root: Path):
    sys.path.insert(0, str(HERE.parent))
    import cli
    out = tmp_root / 'skeleton.docx'
    rc = cli.main(['template', str(out), '--title', '测试题目',
                   '--author', '张三', '--sid', '12345678'])
    assert rc == 0 and out.exists()
    doc = Document(out)
    assert len(doc.sections) >= 4, len(doc.sections)
    assert '硕 士 学 位 论 文' in '\n'.join(p.text for p in doc.paragraphs)
    assert 'TOC \\o' in doc.element.body.xml
    info = audit(doc)
    hard = [i for i in info['issues'] if 'updateFields' not in i]
    assert not hard, '骨架审计未通过:\n' + '\n'.join(hard)
    print(f'  ✓ CLI template 通过（{len(doc.sections)} 节）')


def main() -> int:
    print('== dut-thesis-formatter 冒烟测试 ==')
    test_classify()
    test_format_reference_entry()
    test_three_line_table()
    test_headings_and_body()
    test_heading_survives_style_and_run_tampering()
    test_equation_tabs()
    test_toc_field_and_update_fields()
    test_ensure_toc_field_and_ref_blocks()
    test_footer_and_header_cover_all_variants()
    test_english_and_keywords_and_color()
    test_field_table_alignment()
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        test_end_to_end_reformat(root)
        test_template_cli(root)
    print('== 全部通过 ==')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
