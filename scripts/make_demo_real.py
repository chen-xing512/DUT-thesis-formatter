# -*- coding: utf-8 -*-
"""用 Word 真实渲染的 PDF 页面 + 规范值面板，合成 docs/demo.png。"""
import sys
sys.path.insert(0, 'scripts')
import pymupdf
from PIL import Image, ImageDraw, ImageFont

PDF = 'docs/showcase.pdf'
OUT = 'docs/demo.png'
D = pymupdf.open(PDF)

def page_png(i, dpi=150):
    pix = D[i].get_pixmap(dpi=dpi)
    return Image.frombytes('RGB', (pix.width, pix.height), pix.samples)

# 选页：封面 / 摘要(罗马页码) / 正文首页(阿拉伯页码) / 图表页
pages = [page_png(0, 130), page_png(2, 130), page_png(6, 130), page_png(7, 130)]

GAP = 18
PANEL_W = 560
TOP = 96
BOTTOM = 44
ph = pages[0].height
pw = pages[0].width
W = GAP + len(pages) * (pw + GAP) + PANEL_W + GAP
H = TOP + ph + BOTTOM
canvas = Image.new('RGB', (W, H), '#f4f6f9')
draw = ImageDraw.Draw(canvas)

def font(path, size):
    try:
        return ImageFont.truetype(path, size)
    except Exception:
        return ImageFont.load_default()

HEI = '/System/Library/Fonts/STHeiti Medium.ttc'
SONG = '/System/Library/Fonts/Supplemental/Songti.ttc'
f_title = font(HEI, 34)
f_sub = font(SONG, 19)
f_cap = font(SONG, 13)
f_key = font(HEI, 14)
f_val = font(SONG, 13.5)
f_hdr = font(HEI, 20)

# 标题带
draw.text((GAP + 6, 16), 'dut-thesis-formatter', font=f_title, fill='#12263f')
draw.text((GAP + 6, 58), '大连理工大学硕士学位论文格式 — 以下均为 Word 真实渲染结果',
          font=f_sub, fill='#41556e')

CAPS = ['封面：无页码 / 无页眉', '中文摘要：罗马页码 - I -',
        '正文：阿拉伯页码 - 3 -', '三线表 + 中英文图题']
x = GAP
for im, cap in zip(pages, CAPS):
    canvas.paste(im, (x, TOP))
    draw.rectangle([x, TOP, x + im.width - 1, TOP + im.height - 1], outline='#c8d2dd')
    tw = draw.textlength(cap, font=f_cap)
    draw.text((x + (im.width - tw) / 2, TOP + im.height + 12), cap,
              font=f_cap, fill='#41556e')
    x += im.width + GAP

# 右侧规范值面板
px = x
draw.rectangle([px, TOP, px + PANEL_W, TOP + ph], fill='white', outline='#c8d2dd')
draw.text((px + 22, TOP + 20), 'DUT 关键格式值', font=f_hdr, fill='#12263f')
draw.line([px + 22, TOP + 52, px + PANEL_W - 22, TOP + 52], fill='#dbe3ec', width=2)

rows = [
    ('页面 / 边距', 'A4；上 3.5 / 下 2.5 / 左 2.5 / 右 2.5 cm'),
    ('页眉 / 页脚距', '2.5 cm / 2.0 cm'),
    ('章标题', '黑体 三号 16pt 居左，段后 1 行，另起一页'),
    ('节标题', '黑体 四号 14pt 居左，段前 0.5 行'),
    ('小节标题', '黑体 小四 12pt 居左，段前 0.5 行'),
    ('摘要 / 目录标题', '黑体 小三 15pt 加粗居中，段后 1 行'),
    ('正文', '宋体 小四 12pt，两端对齐，1.25 倍行距'),
    ('正文缩进', '首行缩进 2 字符，取消「对齐到网格」'),
    ('关键词', '「关键词：」黑体小四；分号分隔，末尾无标点'),
    ('图题 / 表题', '宋体 五号 10.5pt 居中，中英文两行'),
    ('表', '三线表：上线/下线 1.5pt，表头线 1pt，无竖线'),
    ('参考文献', '五号宋体，悬挂缩进，GB/T 7714，按引用顺序'),
    ('正文引用', '方括号阿拉伯数字右上角标 [1]、[2-4]'),
    ('页码', '封面无；前置罗马数字；正文阿拉伯从 1 起'),
    ('页眉', '奇数页校名，偶数页论文题目，宋体五号居中'),
]
y = TOP + 70
for k, v in rows:
    draw.text((px + 22, y), k, font=f_key, fill='#1f3a5f')
    draw.text((px + 178, y), v, font=f_val, fill='#333333')
    draw.line([px + 22, y + 24, px + PANEL_W - 22, y + 24], fill='#eef1f5')
    y += 34

draw.text((GAP + 6, H - 28),
          '实测字号 — 论文题目 21.9pt · 章标题 15.9pt（三号16）· 节标题 13.8pt（四号14）· '
          '小节 11.9pt（小四12）· 正文 12.0pt · 摘要 15.0pt（小三）— 均与规范一致',
          font=f_cap, fill='#7a8794')

canvas.save(OUT, optimize=True)
print('已生成', OUT, canvas.size)
