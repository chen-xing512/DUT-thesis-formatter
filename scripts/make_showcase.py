# -*- coding: utf-8 -*-
"""用本技能库造一份内容完整的示例论文，供渲染真实效果图。"""
import sys
from docx.enum.text import WD_ALIGN_PARAGRAPH
sys.path.insert(0, 'scripts')
from docx import Document
from docx.shared import Cm
from dut_thesis import *

OUT = '/Users/apple/Documents/deepseek harness work/DUT-thesis-formatter/docs/showcase.docx'
TITLE = '基于深度学习的遥感小目标检测方法研究'

# 造一张示例插图（不依赖 matplotlib 的复杂逻辑）
import struct, zlib
def png(w, h, rgb):
    raw = b''.join(b'\x00' + bytes(rgb) * w for _ in range(h))
    def ch(t, d):
        c = t + d
        return struct.pack('>I', len(d)) + c + struct.pack('>I', zlib.crc32(c) & 0xffffffff)
    return (b'\x89PNG\r\n\x1a\n' + ch(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0))
            + ch(b'IDAT', zlib.compress(raw)) + ch(b'IEND', b''))
open('/tmp/f1.png', 'wb').write(png(360, 150, (238, 243, 248)))

doc = new_doc()

# ---------- 封面 ----------
add_cover(doc, title_cn=TITLE, title_en='Research on Small Object Detection in Remote Sensing Imagery Based on Deep Learning',
          fields=[('作者姓名', '张三'), ('学号', '12345678'),
                  ('指导教师', '李四 教授'), ('学科、专业', '信息与通信工程'),
                  ('答辩日期', '2026 年 6 月')])

# ---------- 前置：摘要（罗马页码）----------
add_section(doc, kind='front', title_cn=TITLE)
add_plain_heading(doc, '摘　　要')
for t in [
    '遥感图像中的小目标检测在军事侦察、灾害评估与城市规划等领域具有重要的应用价值。'
    '然而，受限于目标尺度小、背景复杂以及类间差异微弱等因素，现有检测方法在小目标上的'
    '精度仍明显低于中大型目标。本文围绕该问题展开研究。',
    '首先，本文分析了遥感小目标检测中特征金字塔的语义—空间失衡机理，指出浅层特征在'
    '多次下采样后信息损失严重是精度下降的主要原因。其次，提出了一种双向特征融合模块'
    '（Bi-FFM），通过自顶向下与自底向上两条路径的交叉聚合，增强了小目标的特征表达。'
    '最后，在公开数据集上进行了系统的消融实验与对比实验。',
    '实验结果表明，本文方法在 mAP 指标上较基线模型提升了 3.55 个百分点，'
    '且在参数量基本持平的前提下，对小目标的召回率提升了 5.1 个百分点，'
    '验证了所提模块的有效性。',
]:
    add_body(doc, t)
add_blank(doc)
p = doc.add_paragraph()
set_run(p.add_run('关键词：'), cn=HEI, size=PT_BODY)
set_run(p.add_run('遥感图像；小目标检测；特征融合；深度学习；注意力机制'), cn=SONG, size=PT_BODY)
fmt_para(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, line=LINE_BODY, snap_to_grid=False)

# ---------- 英文摘要 ----------
add_section(doc, kind='body', title_cn=TITLE)
add_plain_heading(doc, 'ABSTRACT', cn=EN_FONT)
add_body(doc, 'Small object detection in remote sensing imagery is of great value in '
              'military reconnaissance, disaster assessment and urban planning. '
              'However, due to small object scale, cluttered background and subtle '
              'inter-class differences, existing detectors perform notably worse on '
              'small objects than on medium and large ones.', indent=False)
add_body(doc, 'This thesis first analyzes the semantic-spatial imbalance of the feature '
              'pyramid, and then proposes a bidirectional feature fusion module (Bi-FFM) '
              'that cross-aggregates top-down and bottom-up paths. Systematic ablation '
              'and comparison experiments are conducted on public datasets.', indent=False)
add_blank(doc)
p = doc.add_paragraph()
set_run(p.add_run('Key Words：'), cn=EN_FONT, size=PT_BODY, bold=True)
set_run(p.add_run('remote sensing imagery; small object detection; feature fusion; '
                  'deep learning; attention mechanism'), cn=EN_FONT, size=PT_BODY)
fmt_para(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, line=LINE_BODY, snap_to_grid=False)

# ---------- 目录 ----------
add_section(doc, kind='body', title_cn=TITLE)
add_plain_heading(doc, '目　　录')
add_toc_field(doc)

# ---------- 正文 ----------
add_section(doc, kind='body', title_cn=TITLE)
add_heading(doc, '1  绪论', 1, page_break=True)
add_heading(doc, '1.1  研究背景与意义', 2)
add_body(doc, '随着高分辨率遥感卫星技术的迅速发展，遥感图像的获取成本不断降低，'
              '数据量呈指数级增长。如何从海量遥感数据中快速、准确地定位感兴趣目标，'
              '已成为遥感信息处理领域的核心问题之一[1]。')
add_body(doc, '与自然场景图像不同，遥感图像通常具有以下特点：一是成像视角为俯视，'
              '目标缺乏明显的姿态先验；二是目标尺度跨度极大，同一类目标在不同分辨率下'
              '的像素面积可能相差两个数量级[2-3]；三是背景纹理复杂，地物类别之间'
              '存在显著的类间相似性。上述特点使得通用目标检测器直接迁移到遥感场景时'
              '性能明显下降。')
add_heading(doc, '1.1.1  国内外研究现状', 3)
add_body(doc, '早期方法主要基于手工特征，如方向梯度直方图与可变形部件模型。'
              '这类方法依赖先验的模板设计，对目标形变与背景变化的鲁棒性较差[4]。'
              '随着卷积神经网络的兴起，基于锚框的两阶段与单阶段检测器逐步成为主流，'
              '并在多个遥感数据集上取得了显著进展。')
add_body(doc, '近年来，研究者开始关注小目标检测这一子问题。常见思路包括：'
              '提高输入分辨率、设计更密集的锚框、引入特征金字塔的多尺度融合，'
              '以及在损失函数中增加针对小目标的权重[5]。然而，这些方法大多以'
              '增加计算量为代价，难以兼顾精度与效率。')

add_heading(doc, '2  相关理论与方法', 1, page_break=True)
add_heading(doc, '2.1  特征金字塔网络', 2)
add_body(doc, '特征金字塔网络（FPN）通过自顶向下的路径与横向连接，将高层语义'
              '信息传递给浅层特征，从而在多个尺度上获得兼具语义与空间细节的特征表示。'
              '不同方法在公开数据集上的性能对比如表 2.1 所示。')
add_blank(doc)
add_caption(doc, '表 2.1  不同方法在公开数据集上的性能对比',
            'Tab. 2.1  Performance comparison of different methods on public datasets')
three_line_table(
    doc,
    ['方法', '主干网络', 'mAP/%', '小目标召回/%', '参数量/M'],
    [['Faster R-CNN', 'ResNet-50', '74.21', '48.3', '41.5'],
     ['RetinaNet', 'ResNet-50', '75.86', '50.1', '37.8'],
     ['YOLOv5s', 'CSPDarknet', '76.93', '51.7', '7.2'],
     ['本文方法', 'ResNet-50', '78.42', '55.4', '26.1']],
    col_widths=[3.4, 3.2, 2.6, 3.6, 3.2])
add_blank(doc)
add_body(doc, '由表 2.1 可见，本文方法在参数量与基线模型基本持平的前提下，'
              '在小目标召回率上取得了最明显的提升。所提双向特征融合模块的'
              '结构如图 2.1 所示。')
add_figure(doc, '/tmp/f1.png', '图 2.1  双向特征融合模块结构示意图',
           'Fig. 2.1  Architecture of the bidirectional feature fusion module',
           width_cm=11.0)

add_heading(doc, '3  实验结果与分析', 1, page_break=True)
add_heading(doc, '3.1  实验设置', 2)
add_body(doc, '实验在 NVIDIA A100 上完成，采用 PyTorch 框架。输入图像统一缩放至'
              '1024×1024，批次大小为 8，初始学习率设为 0.01，共训练 36 个 epoch。'
              '评价指标采用 COCO 标准下的 mAP@0.5:0.95。')

# ---------- 参考文献 ----------
add_section(doc, kind='body', title_cn=TITLE)
add_plain_heading(doc, '参 考 文 献', toc=True)
for i, r in enumerate([
    'Zhang X, Wang Y, Li J. A survey on small object detection in remote sensing imagery[J]. IEEE Transactions on Geoscience and Remote Sensing, 2022, 60: 1-21.',
    'Li H, Chen S. Multi-scale feature fusion for aerial object detection[J]. Remote Sensing, 2021, 13(9): 1721.',
    '刘洋, 陈明, 王强. 基于注意力机制的遥感图像小目标检测[J]. 电子学报, 2023, 51(4): 892-901.',
    'Zhao K, Liu Q. Deformable part models revisited[J]. Pattern Recognition, 2020, 103: 107-118.',
    'Sun M, Zhou B. Loss reweighting for tiny object detection[C]. Proceedings of the IEEE/CVF Conference on Computer Vision and Pattern Recognition, 2021: 12345-12354.',
], start=1):
    set_run(doc.add_paragraph().add_run(format_reference_entry(r, i)),
            cn=SONG, size=PT_CAPTION)
    last = doc.paragraphs[-1]
    fmt_para(last, align=WD_ALIGN_PARAGRAPH.JUSTIFY, line=LINE_BODY,
             before=0, after=0, snap_to_grid=False, left_indent_pt=8.5, hanging_pt=8.5)

add_section(doc, kind='body', title_cn=TITLE)
add_plain_heading(doc, '致　　谢', toc=True)
add_body(doc, '衷心感谢导师李四教授在选题、研究方案与论文撰写过程中给予的悉心指导。'
              '感谢实验室各位同学在实验与讨论中提供的帮助。感谢家人一直以来的理解与支持。')

enable_update_fields(doc)
doc.save(OUT)
print('已生成', OUT)
