#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""make_demo.py — 生成 `docs/demo.png`（格式效果示意图）。

本机若无 Word/LibreOffice，无法把 .docx 渲染成位图，因此这里用 matplotlib
**按 DUT 规范的实际数值**绘制页面示意图：页边距、页眉、三级标题字号、
正文 1.25 倍行距与首行缩进、三线表线宽、图题位置，全部与
`references/dut-format-rules.md` 一一对应。

    python scripts/make_demo.py
"""
from __future__ import annotations

import sys
from pathlib import Path

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import Rectangle

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / 'docs' / 'demo.png'


def pick_font(candidates=('Songti SC', 'STSong', 'SimSong', 'SimSun',
                          'Noto Serif CJK SC', 'PingFang SC', 'Heiti SC',
                          'Arial Unicode MS')):
    have = {f.name for f in font_manager.fontManager.ttflist}
    for c in candidates:
        if c in have:
            return c
    return 'DejaVu Sans'


CN = pick_font()
matplotlib.rcParams['axes.unicode_minus'] = False

# A4 比例（21 × 29.7 cm），页边距按 mm 换算
A4_W, A4_H = 210.0, 297.0
M_TOP, M_BOT, M_LEFT, M_RIGHT = 35.0, 25.0, 25.0, 25.0
TEXT_W = A4_W - M_LEFT - M_RIGHT          # 160 mm
RIGHT = A4_W - M_RIGHT


def pt(mm: float) -> float:
    """把 pt 值换算成页面毫米单位（1 pt ≈ 0.3528 mm），用于字号示意。"""
    return mm * 0.3528


def draw_page(ax, kind: str):
    ax.set_xlim(0, A4_W)
    ax.set_ylim(A4_H, 0)                  # y 向下，和纸张坐标一致
    ax.axis('off')
    # 纸张
    ax.add_patch(Rectangle((0, 0), A4_W, A4_H, facecolor='white',
                           edgecolor='#d8d8d8', lw=1.2, zorder=0))
    # 版心虚线
    ax.add_patch(Rectangle((M_LEFT, M_TOP), TEXT_W, A4_H - M_TOP - M_BOT,
                           facecolor='none', edgecolor='#cfd8e3',
                           lw=0.8, ls=(0, (3, 3)), zorder=1))
    # 页眉：奇数页「大连理工大学硕士学位论文」，宋体五号居中
    # 封面既无页眉也无页码（规范：封一、封二不编入页码）
    if kind == 'body':
        ax.plot([M_LEFT, RIGHT], [M_TOP - 6, M_TOP - 6], color='#9aa7b4',
                lw=0.7, zorder=2)
        ax.text(A4_W / 2, M_TOP - 12, '大连理工大学硕士学位论文',
                ha='center', va='center', fontsize=6.2, family=CN,
                color='#333', zorder=3)
        ax.text(A4_W / 2, A4_H - M_BOT + 9, '- 3 -', ha='center', va='center',
                fontsize=6.0, family='serif', color='#333', zorder=3)

    y = M_TOP + 6
    if kind == 'cover':
        # 封面：硕士学位论文（华文细黑 24pt）、题目（22pt）、字段块（15pt）
        y = 70
        ax.text(A4_W / 2, y, '硕 士 学 位 论 文', ha='center', va='center',
                fontsize=15, family=CN, fontweight='bold', color='#111', zorder=3)
        y += 24
        ax.text(A4_W / 2, y, '基于深度学习的XX研究', ha='center', va='center',
                fontsize=13.5, family=CN, fontweight='bold', color='#111', zorder=3)
        y += 18
        ax.text(A4_W / 2, y, 'Research on XX Based on Deep Learning',
                ha='center', va='center', fontsize=9.5, family='serif',
                fontweight='bold', color='#111', zorder=3)
        y += 26
        for lab, val in (('作 者 姓 名', '张三'), ('学　　　号', '12345678'),
                         ('指 导 教 师', '李四 教授'),
                         ('学科、 专业', 'XX 工程'),
                         ('答 辩 日 期', '2026 年 6 月')):
            ax.text(A4_W / 2 - 34, y, lab + '：', ha='right', va='center',
                    fontsize=8.6, family=CN, color='#222', zorder=3)
            ax.text(A4_W / 2 - 32, y, '  ' + val + '  ', ha='left', va='center',
                    fontsize=8.6, family=CN, color='#222', zorder=3)
            ax.plot([A4_W / 2 - 32, A4_W / 2 + 34], [y + 4.2, y + 4.2],
                    color='#222', lw=0.7, zorder=3)
            y += 13
        return

    # ---- 正文页 ----
    # 章标题：黑体三号(16pt)居左，段后 1 行
    ax.text(M_LEFT, y, '1  绪论', ha='left', va='top', fontsize=13.5,
            family=CN, fontweight='bold', color='#111', zorder=3)
    y += 17                        # 章标题：三号 + 段后 1 行

    # 节标题：黑体四号(14pt)，段前 0.5 行
    ax.text(M_LEFT, y, '1.1  研究背景与意义', ha='left', va='top', fontsize=11.8,
            family=CN, fontweight='bold', color='#111', zorder=3)
    y += 14                        # 节标题：四号 + 段前 0.5 行
    # 小节标题：黑体小四(12pt)，段前 0.5 行
    ax.text(M_LEFT, y, '1.1.1  国内研究现状', ha='left', va='top', fontsize=10.2,
            family=CN, fontweight='bold', color='#111', zorder=3)
    y += 12                        # 小节标题：小四 + 段前 0.5 行

    # 正文：宋体小四(12pt)，1.25 倍行距，首行缩进 2 字符
    para1 = ('随着深度学习技术的快速发展，目标检测在工业质检、遥感解译等领域'
             '得到了广泛应用[1]。然而，现有方法在复杂背景下的小目标检测上仍存在'
             '精度不足的问题[2-4]。')
    y = draw_body(ax, para1, y, first_line_indent=True)
    y += 3.5                     # 段间距（规范：段前段后 0 行）
    para2 = ('为解决上述问题，本文提出了一种改进的特征融合结构，并在公开数据集'
             '上进行了验证。')
    y = draw_body(ax, para2, y, first_line_indent=True)

    # 三线表：上线/下线 1.5pt，表头线 1pt，无竖线
    y += 9
    y = draw_table(ax, y)

    # 图：居中 + 图题在图下方（宋体五号）
    y += 11
    ax.add_patch(Rectangle((A4_W / 2 - 32, y), 64, 26, facecolor='#eef3f8',
                           edgecolor='#a8bccf', lw=0.8, zorder=3))
    ax.plot([A4_W / 2 - 22, A4_W / 2 + 2, A4_W / 2 + 22], [y + 19, y + 9, y + 17],
            color='#3f6f9f', lw=1.4, marker='o', ms=2.6, zorder=4)
    ax.text(A4_W / 2, y + 13, '图', ha='center', va='center', fontsize=8,
            family=CN, color='#5b7794', zorder=4, alpha=0.0)
    y += 30
    ax.text(A4_W / 2, y, '图 1.1  特征融合结构示意图', ha='center', va='center',
            fontsize=7.4, family=CN, color='#111', zorder=3)
    y += 5.5
    ax.text(A4_W / 2, y, 'Fig. 1.1  Architecture of the feature fusion module',
            ha='center', va='center', fontsize=7.0, family='serif', color='#111', zorder=3)


def draw_body(ax, text, y, first_line_indent=True, size=8.4, line_pitch=6.6,
              chars_per_line=31, indent_chars=2):
    """按 1.25 倍行距绘制正文，首行缩进 2 个字符。

    逐字定位会把字距拉散、看起来像乱码，所以整行一次性绘制，
    只在太长时按固定字数折行。
    """
    lines, cur, first = [], '', True
    limit = chars_per_line - (indent_chars if first_line_indent else 0)
    for ch in text:
        cur += ch
        if len(cur) >= limit:
            lines.append((cur, first))
            cur, first = '', False
            limit = chars_per_line
    if cur:
        lines.append((cur, first))
    for line, is_first in lines:
        x = M_LEFT + (2.4 * indent_chars if (is_first and first_line_indent) else 0)
        ax.text(x, y, line, ha='left', va='baseline', fontsize=size, family=CN,
                color='#222', zorder=3)
        y += line_pitch
    return y


def draw_table(ax, y):
    """三线表：上线 1.5pt、表头线 1pt、下线 1.5pt，无竖线；表内五号居中。"""
    cols = 4
    widths = [0.30, 0.24, 0.23, 0.23]
    xs = [M_LEFT]
    for w in widths:
        xs.append(xs[-1] + TEXT_W * w)
    rows = [['方法', 'mAP/%', '参数量/M', '推理时延/ms'],
            ['基线模型', '78.42', '25.6', '18.3'],
            ['本文方法', '81.97', '26.1', '19.0']]
    # 表题（表上方，宋体五号）
    ax.text(A4_W / 2, y, '表 1.1  不同方法的性能对比', ha='center', va='center',
            fontsize=7.4, family=CN, color='#111', zorder=3)
    y += 5.5
    ax.text(A4_W / 2, y, 'Tab. 1.1  Performance comparison of different methods',
            ha='center', va='center', fontsize=7.0, family='serif', color='#111', zorder=3)
    y += 4.5

    row_h = 7.0
    lw15 = 1.5 * 0.3528 * 2.0      # pt → 图上视觉粗细
    lw10 = 1.0 * 0.3528 * 2.0

    # 上线 1.5pt
    ax.plot([M_LEFT, RIGHT], [y, y], color='#111', lw=lw15, zorder=4,
            solid_capstyle='butt')
    for j, h in enumerate(rows[0]):
        cx = (xs[j] + xs[j + 1]) / 2
        ax.text(cx, y + row_h * 0.62, h, ha='center', va='center', fontsize=7.4,
                family=CN, color='#111', zorder=3)
    y += row_h
    # 表头线 1pt
    ax.plot([M_LEFT, RIGHT], [y, y], color='#111', lw=lw10, zorder=4,
            solid_capstyle='butt')
    for r in rows[1:]:
        for j, v in enumerate(r):
            cx = (xs[j] + xs[j + 1]) / 2
            ax.text(cx, y + row_h * 0.62, v, ha='center', va='center', fontsize=7.4,
                    family=CN, color='#222', zorder=3)
        y += row_h
    # 下线 1.5pt
    ax.plot([M_LEFT, RIGHT], [y, y], color='#111', lw=lw15, zorder=4,
            solid_capstyle='butt')
    return y


def spec_table(ax):
    """右侧规格表：把关键规范值集中展示。"""
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)
    ax.axis('off')
    ax.text(0.05, 9.75, 'DUT 关键格式值', fontsize=11, family=CN,
            fontweight='bold', color='#111', va='top')
    rows = [
        ('页面 / 边距', 'A4；上 3.5 / 下 2.5 / 左 2.5 / 右 2.5 cm'),
        ('页眉 / 页脚距', '2.5 cm / 2.0 cm'),
        ('章标题', '黑体 三号(16pt) 居左，段后 1 行，另起一页'),
        ('节标题', '黑体 四号(14pt) 居左，段前 0.5 行'),
        ('小节标题', '黑体 小四(12pt) 居左，段前 0.5 行'),
        ('正文', '宋体 小四(12pt)，两端对齐，1.25 倍行距'),
        ('正文缩进', '首行缩进 2 字符，取消「对齐到网格」'),
        ('图题 / 表题', '宋体 五号(10.5pt) 居中，中英文两行'),
        ('表', '三线表：上线/下线 1.5pt，表头线 1pt，无竖线'),
        ('参考文献', '五号宋体，悬挂缩进，GB/T 7714，按引用顺序'),
        ('页码', '封面无；前置罗马数字；正文阿拉伯从 1 起'),
        ('页眉', '奇数页校名，偶数页论文题目，宋体五号居中'),
    ]
    y = 9.15
    for k, v in rows:
        ax.text(0.05, y, k, fontsize=8.2, family=CN, fontweight='bold',
                color='#1f3a5f', va='top')
        ax.text(2.85, y, v, fontsize=8.0, family=CN, color='#333', va='top')
        ax.plot([0.05, 9.95], [y - 0.30, y - 0.30], color='#e6e6e6', lw=0.6)
        y -= 0.70
    ax.text(0.05, y - 0.05, '完整规则见 references/dut-format-rules.md',
            fontsize=7.6, family=CN, color='#7a8794', style='italic', va='top')


def main() -> int:
    fig = plt.figure(figsize=(13.0, 8.6), dpi=170, facecolor='white')
    gs = fig.add_gridspec(1, 3, width_ratios=[1.02, 1.02, 1.30],
                          wspace=0.10, left=0.015, right=0.985,
                          top=0.858, bottom=0.050)

    ax0 = fig.add_subplot(gs[0, 0])
    draw_page(ax0, 'cover')
    ax0.set_title('封面（不编页码）', fontsize=10.0, family=CN, pad=14, color='#1f3a5f')

    ax1 = fig.add_subplot(gs[0, 1])
    draw_page(ax1, 'body')
    ax1.set_title('正文页（三线表 + 中英文图题）', fontsize=10.0, family=CN, pad=14,
                  color='#1f3a5f')

    ax2 = fig.add_subplot(gs[0, 2])
    spec_table(ax2)

    fig.text(0.5, 0.992, 'dut-thesis-formatter', ha='center', va='top',
             fontsize=16, family='serif', fontweight='bold', color='#12263f')
    fig.text(0.5, 0.953, '大连理工大学硕士学位论文格式 — 排版效果与关键规范值',
             ha='center', va='top', fontsize=10.5, family=CN, color='#41556e')
    fig.lines.append(plt.Line2D([0.30, 0.70], [0.925, 0.925], color='#d6dee8',
                                lw=1.0, transform=fig.transFigure))
    fig.text(0.5, 0.003, '示意图按《大连理工大学硕士学位论文格式规范》数值绘制；'
                         '实际 .docx 由 scripts/dut_thesis.py 生成',
             ha='center', va='bottom', fontsize=7.6, family=CN, color='#8894a2')

    OUT.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(OUT, dpi=170, facecolor='white')
    plt.close(fig)
    print(f'已生成 {OUT}  ({OUT.stat().st_size / 1024:.0f} KB, 字体 {CN})')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
