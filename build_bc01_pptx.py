#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成第1章“原书目录版”自制课件。内容按朱圣庚《生物化学》第4版第1章节/子节组织。"""
import os, re, json
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(ROOT, "downloads")
os.makedirs(OUT, exist_ok=True)
PPTX = os.path.join(OUT, "bc01-original-style.pptx")

# ---------- palette ----------
NAVY = RGBColor(0x16, 0x2E, 0x47)
NAVY2 = RGBColor(0x1E, 0x3F, 0x5C)
BLUE = RGBColor(0x1F, 0x6F, 0xA8)
TEAL = RGBColor(0x16, 0x8B, 0x8B)
GOLD = RGBColor(0xC9, 0x96, 0x2E)
INK = RGBColor(0x22, 0x30, 0x3C)
MUTED = RGBColor(0x5D, 0x6D, 0x7E)
LIGHT = RGBColor(0xF4, 0xF7, 0xFA)
PALE = RGBColor(0xE8, 0xF0, 0xF7)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0xD5, 0xE1, 0xEC)
RED = RGBColor(0x9B, 0x3A, 0x32)

FONT = "Noto Sans CJK SC"
SERIF = "Noto Serif CJK SC"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

def rgb_hex(c):
    return "%02X%02X%02X" % (c[0], c[1], c[2])

def set_bg(slide, color):
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = color

def box(slide, x, y, w, h, fill=None, line=None, radius=True, dash=False):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
                                 Inches(x), Inches(y), Inches(w), Inches(h))
    shp.shadow.inherit = False
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid(); shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(1.2)
        if dash:
            shp.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    shp.text_frame.clear()
    shp.text_frame.word_wrap = True
    return shp

def text(slide, x, y, w, h, content, size=18, color=INK, bold=False,
         font=FONT, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP, line_spacing=1.12,
         margins=(0.02, 0.02, 0.02, 0.02), italic=False):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = Inches(margins[0]); tf.margin_right = Inches(margins[1])
    tf.margin_top = Inches(margins[2]); tf.margin_bottom = Inches(margins[3])
    if isinstance(content, str):
        content = [content]
    for i, item in enumerate(content):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        if isinstance(item, dict):
            t = item.get("text", "")
            r = p.add_run(); r.text = t
            r.font.size = Pt(item.get("size", size))
            r.font.bold = item.get("bold", bold)
            r.font.color.rgb = item.get("color", color)
            r.font.name = item.get("font", font)
            r.font.italic = item.get("italic", italic)
            if item.get("space_before") is not None:
                p.space_before = Pt(item["space_before"])
        else:
            r = p.add_run(); r.text = str(item)
            r.font.size = Pt(size); r.font.bold = bold
            r.font.color.rgb = color; r.font.name = font; r.font.italic = italic
    return tb

def bullet_text(slide, x, y, w, h, items, size=18, color=INK, gap=8):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = Inches(0.02); tf.margin_right = Inches(0.02)
    tf.margin_top = Inches(0.02); tf.margin_bottom = Inches(0.02)
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap)
        if isinstance(it, dict) and it.get("sub"):
            p.level = 1
            lead = "– "
            txt = it.get("text", "")
            sz = it.get("size", size - 2)
            col = it.get("color", MUTED)
        else:
            txt = it.get("text", it) if isinstance(it, dict) else str(it)
            sz = it.get("size", size) if isinstance(it, dict) else size
            col = it.get("color", color) if isinstance(it, dict) else color
            lead = "▸ " if not (isinstance(it, dict) and it.get("plain")) else ""
        # bold before Chinese colon if present
        r = p.add_run(); r.text = lead + txt
        r.font.size = Pt(sz); r.font.color.rgb = col; r.font.name = FONT
        p.line_spacing = 1.1
    return tb

def kicker(slide, txt, page):
    text(slide, 0.62, 0.36, 9.6, 0.3, txt, size=10.5, color=BLUE, bold=True)
    text(slide, 11.4, 0.36, 1.3, 0.3, "原书 p%d" % page if page else "", size=10,
         color=MUTED, align=PP_ALIGN.RIGHT)

def footer(slide, num, page):
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(0.62), Inches(7.05), Inches(12.72), Inches(7.05))
    line.line.color.rgb = LINE; line.line.width = Pt(0.8)
    text(slide, 0.62, 7.1, 8.0, 0.24, "朱圣庚《生物化学》第4版 · 第1章 生命的分子基础 · 本站自制原书版课件",
         size=8.5, color=MUTED)
    text(slide, 11.8, 7.1, 1.0, 0.24, "%02d" % num, size=9, color=MUTED, align=PP_ALIGN.RIGHT)

def section_slide(num, title, page, questions, sec_no):
    s = prs.slides.add_slide(BLANK); set_bg(s, NAVY)
    # left gold rule
    box(s, 0, 0, 0.12, 7.5, fill=GOLD, line=None, radius=False)
    text(s, 0.85, 0.75, 1.2, 0.5, "SECTION %02d" % sec_no, size=11, color=GOLD, bold=True)
    text(s, 0.85, 1.25, 2.4, 1.2, "%02d" % sec_no, size=88, color=WHITE, bold=True, font=SERIF)
    text(s, 3.05, 1.45, 8.9, 1.25, title, size=36, color=WHITE, bold=True, font=SERIF)
    text(s, 3.08, 2.82, 7.8, 0.4, "原书节级目录 · 第 %d 节" % sec_no, size=13, color=RGBColor(0xB8, 0xCE, 0xDD))
    text(s, 3.08, 3.35, 8.6, 0.38, "本节先回答", size=12, color=GOLD, bold=True)
    y = 3.78
    for q in questions:
        box(s, 3.08, y, 0.14, 0.14, fill=GOLD, line=None, radius=False)
        text(s, 3.34, y - 0.06, 8.4, 0.42, q, size=16, color=WHITE)
        y += 0.55
    text(s, 0.88, 6.85, 9, 0.3, "朱圣庚《生物化学》第4版 · 原书 p%d" % page, size=9.5, color=RGBColor(0x9E, 0xB6, 0xC8))
    text(s, 11.8, 6.85, 0.8, 0.3, "%02d" % num, size=9.5, color=RGBColor(0x9E, 0xB6, 0xC8), align=PP_ALIGN.RIGHT)
    return s

def content_slide(num, sec_label, title, page, bullets, notes, visual=None):
    s = prs.slides.add_slide(BLANK); set_bg(s, LIGHT)
    box(s, 0, 0, 13.333, 0.10, fill=BLUE, line=None, radius=False)
    kicker(s, sec_label, page)
    text(s, 0.62, 0.78, 12.0, 0.72, title, size=27, color=NAVY, bold=True, font=SERIF)
    # content area
    body_w = 7.2 if visual else 7.45
    bullet_text(s, 0.7, 1.72, body_w, 4.95, bullets, size=(18 if visual else 17.2), gap=(9 if visual else 7))
    if visual:
        visual(s, 8.25, 1.62, 4.35, 5.05)
    else:
        cnmap = {"一":1,"二":2,"三":3,"四":4,"五":5,"六":6,"七":7,"八":8}
        cur = cnmap.get(sec_label[:1], 0)
        v_secmap(s, 8.45, 1.62, 4.15, 5.05, cur)
    footer(s, num, page)
    s.notes_slide.notes_text_frame.text = notes
    return s

# ---------- visuals ----------
def v_hierarchy(s, x, y, w, h):
    box(s, x, y, w, h, fill=WHITE, line=LINE)
    text(s, x+0.2, y+0.18, w-0.4, 0.3, "生命结构层次", size=12, color=BLUE, bold=True)
    labels = ["单体亚基", "生物大分子", "超分子复合体", "细胞器", "细胞"]
    colors = [RGBColor(0xE3, 0xF1, 0xF6), RGBColor(0xD7, 0xEA, 0xF3), RGBColor(0xCB, 0xE2, 0xEE), RGBColor(0xBF, 0xD9, 0xE8), NAVY2]
    yy = y + 0.62
    for i, lab in enumerate(labels):
        b = box(s, x+0.55, yy, w-1.1, 0.48, fill=colors[i], line=None)
        text(s, x+0.55, yy+0.09, w-1.1, 0.3, lab, size=12.5 if i < 4 else 13,
             color=INK if i < 4 else WHITE, bold=True, align=PP_ALIGN.CENTER)
        if i < len(labels)-1:
            ar = s.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(x+w/2-0.07), Inches(yy+0.50), Inches(0.14), Inches(0.18))
            ar.fill.solid(); ar.fill.fore_color.rgb = GOLD; ar.line.fill.background()
        yy += 0.7
    text(s, x+0.2, y+4.45, w-0.4, 0.42, "共价键聚合 → 非共价作用组装", size=10.5, color=MUTED, align=PP_ALIGN.CENTER)
    cap(s, x+0.2, y+4.78, w-0.4, "按教材结构层次绘制（矢量）")

def v_dogma(s, x, y, w, h):
    box(s, x, y, w, h, fill=WHITE, line=LINE)
    text(s, x+0.2, y+0.18, w-0.4, 0.3, "遗传信息的传递", size=12, color=BLUE, bold=True)
    items = [("DNA", "贮存遗传信息", NAVY2), ("RNA", "转录/传递", BLUE), ("蛋白质", "执行功能", TEAL)]
    yy = y + 0.72
    for i, (a, b, col) in enumerate(items):
        sh = box(s, x+0.5, yy, w-1.0, 0.62, fill=col, line=None)
        text(s, x+0.5, yy+0.08, w-1.0, 0.26, a, size=16, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        text(s, x+0.5, yy+0.35, w-1.0, 0.2, b, size=9.5, color=RGBColor(0xDD,0xE9,0xF2), align=PP_ALIGN.CENTER)
        if i < 2:
            text(s, x+0.5, yy+0.64, w-1.0, 0.3, "↓  " + ("复制" if i == 0 else "翻译"), size=10, color=GOLD, bold=True, align=PP_ALIGN.CENTER)
        yy += 1.02
    text(s, x+0.2, y+3.55, w-0.4, 1.1, "中心法则：DNA → RNA → 蛋白质；\nDNA复制保证遗传信息忠实传递。", size=10.5, color=MUTED)

def v_energy(s, x, y, w, h):
    box(s, x, y, w, h, fill=WHITE, line=LINE)
    text(s, x+0.2, y+0.18, w-0.4, 0.3, "开放系统与能量流", size=12, color=BLUE, bold=True)
    sun = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x+0.35), Inches(y+0.82), Inches(0.55), Inches(0.55))
    sun.fill.solid(); sun.fill.fore_color.rgb = GOLD; sun.line.fill.background()
    text(s, x+0.30, y+1.42, 0.75, 0.22, "日光", size=9, color=MUTED, align=PP_ALIGN.CENTER)
    box(s, x+1.85, y+0.72, 1.55, 2.1, fill=PALE, line=BLUE)
    text(s, x+1.95, y+1.18, 1.35, 0.7, "活细胞\n开放系统", size=13, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    text(s, x+1.15, y+0.98, 0.55, 0.26, "→", size=18, color=GOLD, bold=True)
    text(s, x+3.52, y+1.02, 0.55, 0.26, "→", size=18, color=TEAL, bold=True)
    text(s, x+0.30, y+2.95, 1.5, 0.6, "低熵\n高能", size=10, color=MUTED, align=PP_ALIGN.CENTER)
    text(s, x+3.55, y+2.95, 0.7, 0.6, "热\n代谢物", size=10, color=MUTED, align=PP_ALIGN.CENTER)
    text(s, x+0.2, y+3.75, w-0.4, 0.95, "摄取低熵高能物质，排出低能高熵产物；\n生命靠持续能量流维持有序。", size=10.5, color=MUTED)

def v_secmap(s, x, y, w, h, cur):
    box(s, x, y, w, h, fill=WHITE, line=LINE)
    text(s, x+0.18, y+0.14, w-0.36, 0.3, "本章位置", size=10.5, color=BLUE, bold=True)
    items = [
        ("1", "生命属性", 15), ("2", "生物化学与分子生物学", 18),
        ("3", "生命物质的化学组成", 21), ("4", "三维结构：构型与构象", 26),
        ("5", "非共价相互作用", 31), ("6", "水和生命", 34),
        ("7", "生命的基本单位——细胞", 42), ("8", "起源与进化", 46),
    ]
    yy = y + 0.55
    for no, name, pg in items:
        on = (int(no) == cur)
        if on:
            box(s, x+0.12, yy-0.02, w-0.24, 0.48, fill=RGBColor(0xE3,0xEF,0xF8), line=None)
        text(s, x+0.20, yy+0.08, 0.28, 0.25, no, size=9.5, color=(BLUE if on else MUTED), bold=True)
        text(s, x+0.50, yy+0.04, w-1.15, 0.36, name, size=8.8 if len(name) > 12 else 9.6,
             color=(NAVY if on else INK), bold=on)
        text(s, x+w-0.62, yy+0.08, 0.45, 0.22, "p%d" % pg, size=7.5, color=MUTED, align=PP_ALIGN.RIGHT)
        yy += 0.53
    text(s, x+0.18, y+h-0.30, w-0.36, 0.22, "当前：第 %d 节 / 共 8 节" % cur if cur else "按左侧目录定位",
         size=8.2, color=GOLD if cur else MUTED, bold=True)

def v_ionpair(s, x, y, w, h):
    box(s, x, y, w, h, fill=WHITE, line=LINE)
    text(s, x+0.2, y+0.18, w-0.4, 0.3, "离子相互作用", size=12, color=BLUE, bold=True)
    b1 = box(s, x+0.35, y+1.30, 1.25, 0.78, fill=RGBColor(0xE3,0xEF,0xF8), line=BLUE)
    text(s, x+0.35, y+1.47, 1.25, 0.3, "–COO⁻", size=15, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    b2 = box(s, x+w-1.60, y+1.30, 1.25, 0.78, fill=RGBColor(0xE4,0xF4,0xF1), line=TEAL)
    text(s, x+w-1.60, y+1.47, 1.25, 0.3, "H₃N⁺", size=15, color=TEAL, bold=True, align=PP_ALIGN.CENTER)
    ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x+1.65), Inches(y+1.69), Inches(x+w-1.65), Inches(y+1.69))
    ln.line.color.rgb = GOLD; ln.line.width = Pt(2.2); ln.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    text(s, x+0.2, y+2.45, w-0.4, 1.9,
         "• 带电基团之间的静电吸引\n• 蛋白质：Asp/Glu 与 Lys/Arg\n• DNA：磷酸骨架与组蛋白、Mg²⁺\n• 高盐可屏蔽电荷，削弱作用",
         size=10.2, color=INK)

def v_ecoli(s, x, y, w, h):
    box(s, x, y, w, h, fill=WHITE, line=LINE)
    text(s, x+0.2, y+0.16, w-0.4, 0.3, "大肠杆菌（原核细胞）", size=11.5, color=BLUE, bold=True)
    # capsule
    cellcap = box(s, x+0.25, y+0.78, w-0.50, 2.35, fill=RGBColor(0xE6,0xF2,0xF7), line=BLUE)
    inner = box(s, x+0.48, y+1.02, w-0.96, 1.86, fill=RGBColor(0xF6,0xFB,0xFD), line=RGBColor(0x9C,0xC7,0xDD))
    text(s, x+0.30, y+0.55, 1.8, 0.24, "细胞壁/细胞膜", size=8.5, color=MUTED)
    # nucleoid
    nuc = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x+1.25), Inches(y+1.35), Inches(1.55), Inches(0.85))
    nuc.fill.solid(); nuc.fill.fore_color.rgb = RGBColor(0xF5,0xD9,0x9A); nuc.line.color.rgb = GOLD
    text(s, x+1.25, y+1.60, 1.55, 0.24, "拟核（环状DNA）", size=7.8, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    # ribosomes
    for dx, dy in [(0.62,1.35),(0.75,1.95),(3.05,1.30),(3.18,1.95),(2.75,2.48),(1.05,2.45)]:
        d = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x+dx), Inches(y+dy), Inches(0.12), Inches(0.12))
        d.fill.solid(); d.fill.fore_color.rgb = TEAL; d.line.fill.background()
    text(s, x+2.82, y+2.58, 1.1, 0.22, "核糖体 70S", size=8.4, color=MUTED)
    # flagellum
    fl = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x+w-0.12), Inches(y+1.75), Inches(x+w+0.02), Inches(y+2.35))
    fl.line.color.rgb = MUTED; fl.line.width = Pt(1.3)
    text(s, x+w-0.98, y+2.32, 0.85, 0.22, "鞭毛", size=8.2, color=MUTED)
    text(s, x+0.22, y+3.35, w-0.44, 1.2,
         "• 无核膜，遗传物质在拟核\n• 环状DNA；可有质粒\n• 70S 核糖体\n• 细胞壁、细胞膜、鞭毛",
         size=9.6, color=INK)
    cap(s, x+0.2, y+4.72, w-0.4, "仿原书 图1-16（矢量重绘）")

def v_euk(s, x, y, w, h):
    box(s, x, y, w, h, fill=WHITE, line=LINE)
    text(s, x+0.2, y+0.16, w-0.4, 0.3, "真核细胞的主要细胞器", size=11.5, color=BLUE, bold=True)
    box(s, x+0.18, y+0.58, w-0.36, 3.05, fill=RGBColor(0xF4,0xFA,0xF8), line=TEAL)
    nu = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x+0.38), Inches(y+0.78), Inches(1.10), Inches(0.98))
    nu.fill.solid(); nu.fill.fore_color.rgb = RGBColor(0xD9,0xE7,0xF5); nu.line.color.rgb = BLUE
    text(s, x+0.38, y+1.08, 1.10, 0.24, "细胞核", size=9.0, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    for dx, dy in [(2.05,0.82),(2.18,1.55)]:
        mi = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x+dx), Inches(y+dy), Inches(1.05), Inches(0.50))
        mi.fill.solid(); mi.fill.fore_color.rgb = RGBColor(0xF7,0xDF,0xCF); mi.line.color.rgb = RED
        text(s, x+dx, y+dy+0.13, 1.05, 0.22, "线粒体", size=8.5, color=RED, bold=True, align=PP_ALIGN.CENTER)
    box(s, x+1.55, y+2.18, 1.15, 0.55, fill=RGBColor(0xE3,0xEF,0xF8), line=BLUE)
    for k in range(3):
        ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x+1.64), Inches(y+2.28+k*0.10), Inches(x+2.58), Inches(y+2.28+k*0.10))
        ln.line.color.rgb = BLUE; ln.line.width = Pt(0.9)
    text(s, x+1.55, y+2.78, 1.15, 0.22, "内质网", size=8.2, color=BLUE, align=PP_ALIGN.CENTER)
    for k in range(3):
        ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x+0.48), Inches(y+2.18+k*0.12), Inches(x+1.18), Inches(y+2.18+k*0.12))
        ln.line.color.rgb = GOLD; ln.line.width = Pt(2)
    text(s, x+0.42, y+2.58, 0.85, 0.22, "高尔基体", size=8.0, color=MUTED)
    ly = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x+2.95), Inches(y+2.38), Inches(0.42), Inches(0.42))
    ly.fill.solid(); ly.fill.fore_color.rgb = RGBColor(0xE8,0xD7,0xF0); ly.line.color.rgb = RGBColor(0x8A,0x5A,0x9A)
    text(s, x+2.92, y+2.84, 0.62, 0.22, "溶酶体", size=7.6, color=MUTED, align=PP_ALIGN.CENTER)
    text(s, x+0.22, y+3.78, w-0.44, 0.65, "核膜、细胞核、线粒体、叶绿体、\n内质网、高尔基体、溶酶体。", size=8.6, color=INK)
    cap(s, x+0.2, y+4.72, w-0.4, "仿原书 图1-17（矢量重绘）")

def v_chirality(s, x, y, w, h):
    box(s, x, y, w, h, fill=WHITE, line=LINE)
    text(s, x+0.2, y+0.18, w-0.4, 0.3, "手性碳与对映体", size=12, color=BLUE, bold=True)
    cx, cy = x+w/2, y+1.95
    c = box(s, x+w/2-0.34, y+1.63, 0.68, 0.56, fill=NAVY2, line=None)
    text(s, x+w/2-0.34, y+1.77, 0.68, 0.28, "C", size=17, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    groups = [("—COOH", cx, y+0.72), ("—NH₂", cx, y+2.76), ("—H", x+0.30, cy), ("—CH₃", x+w-1.28, cy)]
    ends = [((cx, y+1.63), (cx, y+1.40)), ((cx, y+2.19), (cx, y+2.42)),
            ((x+w/2-0.34, cy), (x+0.55, cy)), ((x+w/2+0.34, cy), (x+w-0.55, cy))]
    for (lab, lx, ly), (a, b) in zip(groups, ends):
        ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(a[0]), Inches(a[1]), Inches(b[0]), Inches(b[1]))
        ln.line.color.rgb = MUTED; ln.line.width = Pt(1.4)
        tw = 1.0; tx = lx - tw/2
        if lx < cx: tx = x+0.12
        if lx > cx: tx = x+w-1.12
        text(s, tx, ly-0.16, tw, 0.3, lab, size=10.8, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    text(s, x+0.2, y+3.25, w-0.4, 1.35,
         "四个不同基团连在同一碳上 → 手性碳。\n对映体：互为镜像、不能重合；\n旋光方向相反，生物活性可不同。\nL-氨基酸、D-单糖是生物体常用构型。",
         size=10.2, color=MUTED)
def v_water(s, x, y, w, h):
    box(s, x, y, w, h, fill=WHITE, line=LINE)
    text(s, x+0.2, y+0.18, w-0.4, 0.3, "水分子间的氢键", size=12, color=BLUE, bold=True)
    pts = [(x+0.72, y+1.25), (x+2.5, y+1.05), (x+1.75, y+2.55)]
    for i, (px, py) in enumerate(pts):
        o = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(px), Inches(py), Inches(0.54), Inches(0.54))
        o.fill.solid(); o.fill.fore_color.rgb = BLUE; o.line.fill.background()
        text(s, px, py+0.12, 0.54, 0.25, "O", size=10, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        for dx, dy in [(-0.28, -0.2), (0.55, -0.2)]:
            hh = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(px+dx), Inches(py+dy), Inches(0.22), Inches(0.22))
            hh.fill.solid(); hh.fill.fore_color.rgb = RGBColor(0xE8,0xF0,0xF7); hh.line.color.rgb = BLUE
            text(s, px+dx, py+dy+0.025, 0.22, 0.16, "H", size=6.5, color=NAVY, align=PP_ALIGN.CENTER)
    for (px1,py1),(px2,py2) in [(pts[0],pts[1]),(pts[1],pts[2]),(pts[2],pts[0])]:
        ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(px1+0.3), Inches(py1+0.3), Inches(px2+0.3), Inches(py2+0.3))
        ln.line.color.rgb = GOLD; ln.line.width = Pt(1.5); ln.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    text(s, x+0.2, y+3.15, w-0.4, 1.2, "每个水分子最多形成4个氢键；\n氢键网络 → 高热容、高沸点、高表面张力。\n水是极性溶剂，也是许多反应的参与者。", size=10.2, color=MUTED)
    cap(s, x+0.2, y+4.72, w-0.4, "仿原书 图1-12（矢量重绘）")

def v_table(s, x, y, w, h):
    rows = [
        ("作用", "要点与例子"),
        ("离子相互作用", "电荷吸引：–COO⁻ … ⁺H₃N–"),
        ("氢键", "N/O—H … 孤对电子；DNA碱基配对"),
        ("范德华力", "瞬时偶极；紧密堆积（疏水核心）"),
        ("疏水作用", "非极性基团避水；熵驱动折叠"),
    ]
    tbl = s.shapes.add_table(len(rows), 2, Inches(x), Inches(y), Inches(w), Inches(h)).table
    tbl.columns[0].width = Inches(1.28)
    tbl.columns[1].width = Inches(w - 1.28)
    for i, row in enumerate(rows):
        tbl.rows[i].height = Inches(0.55 if i == 0 else 0.90)
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            cell.text = val
            cell.margin_left = Inches(0.06); cell.margin_right = Inches(0.04)
            cell.margin_top = Inches(0.04); cell.margin_bottom = Inches(0.04)
            cell.fill.solid(); cell.fill.fore_color.rgb = (NAVY2 if i == 0 else (WHITE if i % 2 else PALE))
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT if j else PP_ALIGN.CENTER
                for r in p.runs:
                    r.font.name = FONT
                    r.font.size = Pt(9.2 if i == 0 else 9.0)
                    r.font.bold = (i == 0)
                    r.font.color.rgb = WHITE if i == 0 else INK
def v_timeline(s, x, y, w, h):
    box(s, x, y, w, h, fill=WHITE, line=LINE)
    text(s, x+0.2, y+0.18, w-0.4, 0.3, "生物化学的三个阶段", size=12, color=BLUE, bold=True)
    yline = y + 1.25
    ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x+0.45), Inches(yline), Inches(x+w-0.45), Inches(yline))
    ln.line.color.rgb = LINE; ln.line.width = Pt(2)
    stages = [("1770–1900", "静态/叙述", "分离组成\n结构测定"), ("1900–1950", "动态", "代谢途径\n酶促反应"), ("1950—", "分子", "大分子结构\n分子生物学")]
    for i, (yr, name, desc) in enumerate(stages):
        cx = x + 0.85 + i * (w-1.4)/2.0
        dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx-0.11), Inches(yline-0.11), Inches(0.22), Inches(0.22))
        dot.fill.solid(); dot.fill.fore_color.rgb = GOLD; dot.line.fill.background()
        text(s, cx-0.7, yline-0.62, 1.4, 0.25, yr, size=10, color=MUTED, align=PP_ALIGN.CENTER)
        text(s, cx-0.7, yline+0.22, 1.4, 0.26, name, size=12, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
        text(s, cx-0.7, yline+0.56, 1.4, 0.7, desc, size=10.2, color=MUTED, align=PP_ALIGN.CENTER)
    text(s, x+0.25, y+3.55, w-0.5, 0.55, "研究重点：组成与结构 → 代谢与能量 → 大分子与分子机制", size=9.8, color=GOLD, bold=True, align=PP_ALIGN.CENTER)

def v_ph(s, x, y, w, h):
    box(s, x, y, w, h, fill=WHITE, line=LINE)
    text(s, x+0.2, y+0.18, w-0.4, 0.3, "pH与缓冲", size=12, color=BLUE, bold=True)
    bar_x, bar_y, bar_w = x+0.35, y+1.1, w-0.7
    for i in range(14):
        b = box(s, bar_x+i*bar_w/14, bar_y, bar_w/14+0.01, 0.22, fill=(RGBColor(0xDD,0x6B,0x5B) if i<3 else RGBColor(0xF0,0xC4,0x5E) if i<5 else RGBColor(0x7D,0xC8,0xB4) if i<9 else RGBColor(0x6E,0x9F,0xC8) if i<12 else RGBColor(0x7D,0x86,0xC8)), line=None, radius=False)
    for i in range(0,15,2):
        text(s, bar_x+i*bar_w/14-0.12, bar_y+0.28, 0.3, 0.2, str(i), size=9.0, color=MUTED, align=PP_ALIGN.CENTER)
    text(s, bar_x, bar_y+0.62, bar_w, 0.28, "pH = pKa + lg([A⁻]/[HA])", size=14, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    text(s, x+0.2, y+2.2, w-0.4, 1.9, "• 当[A⁻]=[HA]时，pH=pKa\n• 缓冲液抵抗少量酸/碱引起的pH变化\n• 蛋白质、核酸的电荷状态随pH改变\n• 等电点pI：净电荷为零", size=11.8, color=INK)

def cap(s, x, y, w, txt):
    text(s, x, y, w, 0.24, txt, size=8.0, color=MUTED, italic=True)

def v_carbon(s, x, y, w, h):
    box(s, x, y, w, h, fill=WHITE, line=LINE)
    text(s, x+0.2, y+0.16, w-0.4, 0.3, "碳原子的四面体构型", size=11.5, color=BLUE, bold=True)
    cx, cy = x+w/2, y+2.08
    # tetrahedron
    pts = [(cx, y+0.93), (cx, y+3.18), (x+0.72, y+2.45), (x+w-0.72, y+2.45)]
    for px, py in pts:
        ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(cx), Inches(cy), Inches(px), Inches(py))
        ln.line.color.rgb = BLUE; ln.line.width = Pt(2)
    c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx-0.22), Inches(cy-0.22), Inches(0.44), Inches(0.44))
    c.fill.solid(); c.fill.fore_color.rgb = NAVY2; c.line.fill.background()
    text(s, cx-0.22, cy-0.13, 0.44, 0.24, "C", size=12, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    labels = [("H", pts[0][0], pts[0][1]-0.26), ("H", pts[1][0], pts[1][1]+0.02),
              ("—OH", pts[2][0]-0.42, pts[2][1]-0.05), ("—NH₂", pts[3][0]-0.30, pts[3][1]-0.05)]
    for lab, lx, ly in labels:
        text(s, lx, ly, 0.85, 0.26, lab, size=9.5, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    text(s, x+0.18, y+3.55, w-0.36, 0.9, "碳原子四价，四个单键指向四面体四个顶点；\n可连接 H、O、N、S 等，形成多样骨架。", size=9.0, color=INK)
    cap(s, x+0.2, y+4.72, w-0.4, "仿原书 图1-2（矢量重绘）")

def v_functional(s, x, y, w, h):
    box(s, x, y, w, h, fill=WHITE, line=LINE)
    text(s, x+0.2, y+0.16, w-0.4, 0.3, "生物分子中常见的官能团", size=11.5, color=BLUE, bold=True)
    rows = [("羟基", "—OH"), ("羧基", "—COOH"), ("氨基", "—NH₂"), ("磷酸基", "—OPO₃²⁻"), ("巯基", "—SH")]
    yy = y+0.62
    for i, (name, formula) in enumerate(rows):
        box(s, x+0.18, yy, w-0.36, 0.56, fill=(PALE if i % 2 == 0 else WHITE), line=None)
        text(s, x+0.30, yy+0.13, 1.2, 0.26, name, size=9.6, color=NAVY, bold=True)
        text(s, x+1.55, yy+0.12, w-1.7, 0.3, formula, size=11.0, color=TEAL, bold=True)
        yy += 0.64
    text(s, x+0.2, y+4.05, w-0.4, 0.6, "官能团决定极性和化学反应性，\n碳骨架决定分子形状。", size=9.0, color=INK)
    cap(s, x+0.2, y+4.72, w-0.4, "仿原书 图1-3（矢量重绘）")

def v_stereo(s, x, y, w, h):
    box(s, x, y, w, h, fill=WHITE, line=LINE)
    text(s, x+0.2, y+0.16, w-0.4, 0.3, "立体异构：顺反与对映", size=11.5, color=BLUE, bold=True)
    for j, (lab, right_top) in enumerate([("顺式", "CH₃"), ("反式", "H")]):
        bx = x + 0.28 + j*(w/2-0.10)
        yy = y + 1.55
        for dy in (-0.06, 0.06):
            ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(bx+0.38), Inches(yy+dy), Inches(bx+0.95), Inches(yy+dy))
            ln.line.color.rgb = NAVY2; ln.line.width = Pt(1.8)
        for px in (bx+0.38, bx+0.95):
            ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(px), Inches(yy-0.08), Inches(px), Inches(yy-0.42))
            ln.line.color.rgb = NAVY2; ln.line.width = Pt(1.2)
            ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(px), Inches(yy+0.08), Inches(px), Inches(yy+0.42))
            ln.line.color.rgb = NAVY2; ln.line.width = Pt(1.2)
        text(s, bx+0.18, yy-1.02, 0.55, 0.22, "CH₃", size=7.6, color=INK, align=PP_ALIGN.CENTER)
        text(s, bx+0.78, yy-1.02, 0.55, 0.22, right_top, size=7.6, color=INK, align=PP_ALIGN.CENTER)
        text(s, bx+0.18, yy+0.58, 0.55, 0.22, "H", size=7.6, color=INK, align=PP_ALIGN.CENTER)
        text(s, bx+0.78, yy+0.58, 0.55, 0.22, ("H" if right_top == "CH₃" else "CH₃"), size=7.6, color=INK, align=PP_ALIGN.CENTER)
        text(s, bx+0.30, yy+0.95, 0.8, 0.24, lab, size=8.6, color=(TEAL if j == 0 else GOLD), bold=True, align=PP_ALIGN.CENTER)
    box(s, x+0.28, y+2.85, w-0.56, 1.15, fill=PALE, line=None)
    text(s, x+0.42, y+2.98, w-0.84, 0.26, "对映体", size=9.0, color=NAVY, bold=True)
    text(s, x+0.42, y+3.28, w-0.84, 0.62, "互为镜像、不能重合；\n旋光方向相反，生物活性可不同。\n具体 D/L 构型见下页。", size=8.2, color=INK)
    cap(s, x+0.2, y+4.72, w-0.4, "仿原书 图1-4（矢量重绘）")

def v_glyceraldehyde(s, x, y, w, h):
    box(s, x, y, w, h, fill=WHITE, line=LINE)
    text(s, x+0.2, y+0.16, w-0.4, 0.3, "甘油醛的 D/L 构型", size=11.5, color=BLUE, bold=True)
    for j, lab in enumerate(["D-甘油醛", "L-甘油醛"]):
        bx = x+0.45 + j*1.85
        top = y+1.10; bot = y+3.25; mid = y+2.18
        ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(bx), Inches(top), Inches(bx), Inches(bot))
        ln.line.color.rgb = NAVY2; ln.line.width = Pt(1.4)
        ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(bx-0.50), Inches(mid), Inches(bx+0.50), Inches(mid))
        ln.line.color.rgb = NAVY2; ln.line.width = Pt(1.4)
        text(s, bx-0.40, top-0.25, 0.8, 0.22, "CHO", size=8.6, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
        text(s, bx-0.40, bot+0.04, 0.8, 0.22, "CH₂OH", size=8.6, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
        if j == 0:
            text(s, bx-0.85, mid-0.12, 0.32, 0.22, "H", size=9, color=RED, bold=True)
            text(s, bx+0.50, mid-0.12, 0.36, 0.22, "OH", size=9, color=TEAL, bold=True)
        else:
            text(s, bx-0.85, mid-0.12, 0.36, 0.22, "HO", size=9, color=TEAL, bold=True)
            text(s, bx+0.50, mid-0.12, 0.32, 0.22, "H", size=9, color=RED, bold=True)
        text(s, bx-0.48, bot+0.34, 0.96, 0.24, lab, size=8.4, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x+w/2-0.65), Inches(y+0.85), Inches(x+w/2-0.65), Inches(y+3.85))
    ln.line.color.rgb = GOLD; ln.line.width = Pt(1.1); ln.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    text(s, x+w/2-1.0, y+3.92, 1.1, 0.22, "镜面", size=7.5, color=GOLD, align=PP_ALIGN.CENTER)
    text(s, x+0.2, y+4.25, w-0.4, 0.4, "L-氨基酸、D-单糖是生物体常用构型", size=8.6, color=INK, bold=True)
    cap(s, x+0.2, y+4.72, w-0.4, "仿原书 图1-5（矢量重绘）")

def v_models(s, x, y, w, h):
    box(s, x, y, w, h, fill=WHITE, line=LINE)
    text(s, x+0.2, y+0.16, w-0.4, 0.3, "三维结构的分子模型", size=11.5, color=BLUE, bold=True)
    rows = [("骨架式", "显示连接次序"), ("球棍式", "显示键角与空间"), ("空间填充", "显示分子外形")]
    for i, (name, note) in enumerate(rows):
        yy = y + 0.62 + i*1.25
        box(s, x+0.22, yy, w-0.44, 1.05, fill=(PALE if i % 2 == 0 else WHITE), line=LINE)
        cx = x+0.95
        if i == 0:
            pts = [(cx-0.35,yy+0.72),(cx-0.12,yy+0.30),(cx+0.14,yy+0.70),(cx+0.40,yy+0.30),(cx+0.62,yy+0.66)]
            for a,b in zip(pts,pts[1:]):
                ln=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(a[0]), Inches(a[1]), Inches(b[0]), Inches(b[1]))
                ln.line.color.rgb=NAVY2; ln.line.width=Pt(1.6)
        elif i == 1:
            pts = [(cx-0.32,yy+0.62),(cx-0.08,yy+0.32),(cx+0.18,yy+0.66),(cx+0.44,yy+0.34)]
            for a,b in zip(pts,pts[1:]):
                ln=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(a[0]), Inches(a[1]), Inches(b[0]), Inches(b[1]))
                ln.line.color.rgb=MUTED; ln.line.width=Pt(2)
            for k,(px,py) in enumerate(pts):
                c=s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(px-0.10), Inches(py-0.10), Inches(0.20), Inches(0.20))
                c.fill.solid(); c.fill.fore_color.rgb=[BLUE,TEAL,GOLD,RED][k]; c.line.fill.background()
        else:
            pts = [(cx-0.28,yy+0.52,0.48),(cx+0.02,yy+0.38,0.52),(cx+0.26,yy+0.62,0.40)]
            for px,py,rr in pts:
                c=s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(px), Inches(py), Inches(rr), Inches(rr))
                c.fill.solid(); c.fill.fore_color.rgb=RGBColor(0xCF,0xE3,0xF0); c.line.color.rgb=BLUE
        text(s, x+2.10, yy+0.18, 1.85, 0.26, name, size=10.2, color=NAVY, bold=True)
        text(s, x+2.10, yy+0.50, 1.85, 0.24, note, size=8.4, color=MUTED)
    text(s, x+0.2, y+4.35, w-0.4, 0.28, "不同模型强调不同结构信息", size=8.8, color=INK)
    cap(s, x+0.2, y+4.72, w-0.4, "仿原书 图1-9（矢量重绘）")

def v_tree(s, x, y, w, h):
    box(s, x, y, w, h, fill=WHITE, line=LINE)
    text(s, x+0.2, y+0.16, w-0.4, 0.3, "生物三域系统树", size=11.5, color=BLUE, bold=True)
    root = (x+w/2, y+3.85)
    tips = [(x+0.62, y+1.05, "细菌"), (x+w/2, y+0.88, "古菌"), (x+w-0.62, y+1.05, "真核生物")]
    for tx, ty, lab in tips:
        ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(root[0]), Inches(root[1]), Inches(tx), Inches(ty+0.35))
        ln.line.color.rgb = BLUE; ln.line.width = Pt(1.6)
        c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(tx-0.11), Inches(ty), Inches(0.22), Inches(0.22))
        c.fill.solid(); c.fill.fore_color.rgb = [TEAL, GOLD, BLUE][tips.index((tx,ty,lab))]; c.line.fill.background()
        text(s, tx-0.6, ty-0.30, 1.2, 0.24, lab, size=8.6, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    text(s, root[0]-0.65, root[1]+0.08, 1.3, 0.24, "共同祖先", size=8.2, color=MUTED, align=PP_ALIGN.CENTER)
    text(s, x+0.2, y+3.90, w-0.4, 0.55, "细菌、古菌为原核；真核生物有核膜和细胞器。", size=8.4, color=INK)
    cap(s, x+0.2, y+4.72, w-0.4, "仿原书 图1-15（矢量重绘）")

def v_rnaworld(s, x, y, w, h):
    box(s, x, y, w, h, fill=WHITE, line=LINE)
    text(s, x+0.2, y+0.16, w-0.4, 0.3, "“RNA 世界”假说", size=11.5, color=BLUE, bold=True)
    items = [("RNA", "信息 + 催化", TEAL), ("复制", "自然选择", GOLD), ("DNA/蛋白质", "分工", BLUE)]
    yy = y+0.78
    for i, (a, b, col) in enumerate(items):
        box(s, x+0.55, yy, w-1.1, 0.62, fill=col, line=None)
        text(s, x+0.55, yy+0.07, w-1.1, 0.26, a, size=11.5, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        text(s, x+0.55, yy+0.35, w-1.1, 0.2, b, size=8.2, color=WHITE, align=PP_ALIGN.CENTER)
        if i < 2:
            text(s, x+0.55, yy+0.64, w-1.1, 0.28, "↓", size=13, color=GOLD, bold=True, align=PP_ALIGN.CENTER)
        yy += 1.02
    text(s, x+0.2, y+3.82, w-0.4, 0.7, "RNA 既能贮存信息又能催化；\n后来 DNA 贮存信息、蛋白质承担催化。", size=8.8, color=INK)
    cap(s, x+0.2, y+4.72, w-0.4, "仿原书 图1-19（矢量重绘）")

def v_endosym(s, x, y, w, h):
    box(s, x, y, w, h, fill=WHITE, line=LINE)
    text(s, x+0.2, y+0.16, w-0.4, 0.3, "内共生学说", size=11.5, color=BLUE, bold=True)
    host = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x+0.35), Inches(y+1.05), Inches(1.55), Inches(1.55))
    host.fill.solid(); host.fill.fore_color.rgb = RGBColor(0xE7,0xF1,0xF8); host.line.color.rgb = BLUE
    text(s, x+0.55, y+1.70, 1.15, 0.24, "宿主细胞", size=8.6, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    bact = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x+1.10), Inches(y+1.55), Inches(0.60), Inches(0.36))
    bact.fill.solid(); bact.fill.fore_color.rgb = RGBColor(0xF4,0xDA,0xCF); bact.line.color.rgb = RED
    text(s, x+1.08, y+1.03, 0.75, 0.22, "细菌", size=7.5, color=RED, bold=True)
    text(s, x+2.15, y+1.65, 0.5, 0.32, "→", size=18, color=GOLD, bold=True)
    mito = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x+2.65), Inches(y+1.35), Inches(1.15), Inches(0.58))
    mito.fill.solid(); mito.fill.fore_color.rgb = RGBColor(0xF9,0xE1,0xD2); mito.line.color.rgb = RED
    text(s, x+2.65, y+1.48, 1.15, 0.24, "线粒体", size=8.6, color=RED, bold=True, align=PP_ALIGN.CENTER)
    # inner membrane cristae
    for i in range(3):
        ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x+2.78), Inches(y+1.62+i*0.10), Inches(x+2.95), Inches(y+1.55+i*0.10))
        ln.line.color.rgb = RED; ln.line.width = Pt(0.8)
    text(s, x+0.2, y+3.05, w-0.4, 1.0, "被吞入的需氧细菌未被消化，\n与宿主形成稳定内共生 → 线粒体。\n叶绿体同理来自蓝细菌。", size=8.8, color=INK)
    cap(s, x+0.2, y+4.72, w-0.4, "仿原书 图1-20（矢量重绘）")

# ---------- deck ----------
slides = []
def add(kind, **kw):
    slides.append((kind, kw))

# Cover
s = prs.slides.add_slide(BLANK); set_bg(s, NAVY)
box(s, 0, 0, 0.14, 7.5, fill=GOLD, line=None, radius=False)
text(s, 0.95, 0.85, 8.2, 0.4, "BIOCHEMISTRY · CHAPTER 01", size=11, color=GOLD, bold=True)
text(s, 0.95, 1.55, 11.2, 0.7, "第1章", size=22, color=RGBColor(0xB8,0xCE,0xDD), bold=True)
text(s, 0.95, 2.05, 11.6, 1.3, "生命的分子基础", size=52, color=WHITE, bold=True, font=SERIF)
text(s, 0.98, 3.48, 10.6, 0.5, "朱圣庚《生物化学》第4版 · 原书目录版课件", size=17, color=RGBColor(0xD5,0xE3,0xEE))
box(s, 0.98, 4.22, 4.4, 0.42, fill=RGBColor(0x24,0x4C,0x6C), line=None)
text(s, 1.08, 4.30, 4.2, 0.28, "自制教学课件 · 非老师原PPT", size=10.5, color=GOLD, bold=True)
text(s, 0.98, 6.55, 10.5, 0.3, "按原书 8 节 / 33 个子节组织 · 适合预习、课堂讲解与复习", size=9.5, color=RGBColor(0x9E,0xB6,0xC8))
text(s, 11.8, 6.55, 0.8, 0.3, "01", size=9.5, color=RGBColor(0x9E,0xB6,0xC8), align=PP_ALIGN.RIGHT)

# Map
s = prs.slides.add_slide(BLANK); set_bg(s, LIGHT)
box(s, 0, 0, 13.333, 0.10, fill=BLUE, line=None, radius=False)
kicker(s, "本章学习地图", 0)
text(s, 0.62, 0.78, 11.0, 0.55, "第1章 生命的分子基础 · 8 节 33 个子节", size=26, color=NAVY, bold=True, font=SERIF)
text(s, 0.65, 1.35, 11.7, 0.35, "先建立“生命属性 → 生物分子 → 三维结构 → 非共价作用 → 水 → 细胞 → 起源进化”的主线。", size=11.5, color=MUTED)
secs = [
 ("01", "生命属性与生物化学", "p15", "生命是什么？生物化学研究什么？"),
 ("02", "生物化学与分子生物学", "p18", "学科如何发展？"),
 ("03", "生命物质的化学组成", "p21", "生命由哪些元素和分子组成？"),
 ("04", "生物分子的三维结构", "p26", "构型与构象有何区别？"),
 ("05", "生物系统中的非共价相互作用", "p31", "弱作用如何稳定结构？"),
 ("06", "水和生命", "p34", "水为什么是生命介质？"),
 ("07", "生命的基本单位——细胞", "p42", "原核与真核如何区分？"),
 ("08", "生物分子的起源与进化", "p46", "生命如何起源？"),
]
for i, (no, t, p, q) in enumerate(secs):
    col = i % 2; row = i // 2
    x = 0.62 + col * 6.25; y = 1.95 + row * 1.19
    box(s, x, y, 5.95, 1.02, fill=WHITE, line=LINE)
    box(s, x, y, 0.08, 1.02, fill=GOLD if i < 4 else TEAL, line=None, radius=False)
    text(s, x+0.22, y+0.12, 0.55, 0.3, no, size=15, color=GOLD if i < 4 else TEAL, bold=True)
    text(s, x+0.75, y+0.12, 4.3, 0.3, t, size=12.8, color=NAVY, bold=True)
    text(s, x+0.75, y+0.47, 4.9, 0.3, q, size=10.2, color=MUTED)
    text(s, x+5.22, y+0.12, 0.55, 0.28, p, size=8.5, color=MUTED, align=PP_ALIGN.RIGHT)
footer(s, 2, 0)
s.notes_slide.notes_text_frame.text = "先用学习地图建立全章框架。第1章不是零散概念，而是从生命属性出发，依次回答生命物质是什么、如何形成三维结构、靠什么作用稳定、水与细胞的意义，最后落到生命起源。"

# Section 1
section_slide(3, "生命属性与生物化学", 15, ["生命属性可以归纳为哪三个方面？", "为什么说生物分子既复杂又同一？", "自复制与自装配为什么是生命的精华？"], 1)
content_slide(4, "一、生命属性与生物化学 · 1/3", "（一）化学成分复杂而同一、结构错综而有序", 15,
 ["细胞含上千种分子：水、无机离子、蛋白质、核酸、糖类、脂质。",
  "蛋白质、核酸、多糖是由单体共价聚合而成的生物大分子；脂质是“准生物大分子”。",
  "生物界化学成分高度复杂，但结构模式和功能高度同一。",
  "结构层次：单体亚基 → 大分子 → 超分子复合体 → 细胞器 → 细胞。",
  "共价键负责聚合；非共价作用负责更高层次组装。"],
 "这一页是本章的总纲。强调“复杂”指种类多、结构精细；“同一”指所有生物都用相似的生物分子和同一套构建逻辑。结构层次要能顺背，并区分共价键与非共价作用分别负责哪一层。",
 v_hierarchy)
content_slide(5, "一、生命属性与生物化学 · 2/3", "（二）利用环境的能量和物质进行自我更新", 16,
 ["新陈代谢=合成代谢（同化）+分解代谢（异化）。",
  "代谢途径：上一步产物是下一步底物，每一步由专一酶催化。",
  "分解代谢释放能量，贮存于ATP、NADPH；合成代谢利用能量建造大分子。",
  "生物是开放系统：摄取低熵高能物质，排出低能高熵产物。",
  "能量转化效率高，且能在恒温恒压下完成。"],
 "用“开放系统”解释生命为什么能维持有序：不是违背热力学第二定律，而是把熵增转移到环境中。强调ATP是通用能量货币，NADPH主要提供还原力。",
 v_energy)
content_slide(6, "一、生命属性与生物化学 · 3/3", "（三）自我复制和自我装配是生命状态的精华", 17,
 ["DNA是主要遗传物质：信息量大、稳定、可忠实复制。",
  "DNA半保留复制；突变与重组是变异和进化的分子基础。",
  "遗传信息表达：DNA→RNA→蛋白质；三联体密码决定氨基酸。",
  "自装配：亚基自发缔合成复合体，驱动力主要是疏水作用。",
  "装配信息主要包含在亚基中；逐级装配节省遗传信息、减少缺陷。"],
 "这一页把“遗传”与“自装配”放在一起：前者保证世代之间信息连续，后者保证分子自动形成有序结构。中心法则只讲信息流，不要把它误解为所有信息都只从DNA单向流动。",
 v_dogma)
# Section 2
section_slide(7, "生物化学与分子生物学", 18, ["生物化学经历哪三个阶段？", "分子生物学与生物化学是什么关系？", "中国生物化学有哪些重要贡献？"], 2)
content_slide(8, "二、生物化学与分子生物学 · 1/3", "（一）现代生物化学的发展和分子生物学的兴起", 18,
 ["第一阶段（1770—1900）：静态/叙述生物化学，分离、纯化、测定组成与结构。",
  "第二阶段（1900—1950）：动态生物化学，研究代谢途径、酶促反应与能量转化。",
  "第三阶段（1950—）：分子生物化学，研究生物大分子结构与功能。",
  "1953年DNA双螺旋模型开启分子生物学；分子生物学研究DNA、RNA、蛋白质及其关系。",
  "分子生物学是生物化学发展的新阶段，二者高度交叉。"],
 "按“静态—动态—分子”三阶段讲，学生最容易记住。提醒：阶段不是截然断开，而是研究重点转移；1953年DNA双螺旋是标志性事件。",
 v_timeline)
content_slide(9, "二、生物化学与分子生物学 · 2/3", "（二）现代生物化学发展中的一些重要学术中心和学者", 19,
 ["Scheele：从生物材料分离柠檬酸、乳酸、尿酸、甘油等，静态生物化学开始。",
  "Lavoisier：呼吸是缓慢氧化；Liebig：提出新陈代谢概念。",
  "Hoppe-Seyler：建立生理化学学科，提出“biochemie”。",
  "Pasteur、Fischer、Hopkins、Pavlov、Sumner 等推动微生物学、结构化学、维生素、消化生理与酶化学。",
  "Watson、Crick、Wilkins：DNA双螺旋；Sanger：蛋白质与DNA测序。"],
 "人物不宜逐个背履历，抓住“贡献关键词”。Scheele—分离有机酸；Lavoisier—呼吸氧化；Liebig—代谢概念；Hoppe-Seyler—学科建立；Sumner—酶是蛋白质；Watson/Crick—双螺旋；Sanger—测序。")
content_slide(10, "二、生物化学与分子生物学 · 3/3", "（三）发展中的我国生物化学", 20,
 ["古代发酵实践：酿酒、制醋、造酱、制饴；《齐民要术》记载制曲。",
  "“曲”是促进淀粉糖化和酒化的粗酶制剂，接近现代酶的概念。",
  "《黄帝内经》提出“五谷为养，五畜为益，五果为助，五菜为充”。",
  "现代：吴宪等在蛋白质变性等方面奠基；1965年人工合成牛胰岛素。",
  "1981年酵母丙氨酸tRNA人工合成；中国生物化学逐步走向世界前沿。"],
 "这一页兼顾古代实践与现代成就。重点讲“曲”与酶的关系、牛胰岛素人工合成，既能体现学科历史，也能联系蛋白质结构与功能。")
# Section 3
section_slide(11, "生命物质的化学组成", 21, ["生命元素如何分类？", "为什么说生物分子是碳的化合物？", "小分子、大分子与单体是什么关系？"], 3)
content_slide(12, "三、生命物质的化学组成 · 1/4", "（一）生命元素", 21,
 ["大量元素：H、O、C、N、Ca、P、K、S、Na、Cl、Mg。",
  "微量元素：Fe、Zn、Cu、Mn、Co、Mo、I、Se 等。",
  "C、H、O、N 占生物体质量绝大部分；C 是生物分子的骨架。",
  "大量元素构成水和有机物；微量元素多作为酶、辅酶或激素的组成成分。",
  "元素种类与无机界相同，但比例和组合方式不同。"],
 "区分“大量/微量”按含量，不按重要性；微量元素含量低但常不可替代。考试常考：Fe、Zn、Cu、Mn、Co、Mo、I、Se 等。",
 v_carbon)
content_slide(13, "三、生命物质的化学组成 · 2/4", "（二）生物分子是碳的化合物", 23,
 ["碳原子四价，可形成四个稳定共价键。",
  "碳链可长可短、可分支、可成环，并可连接多种官能团。",
  "常见官能团：羟基、羧基、氨基、磷酸基、巯基、醛/酮基。",
  "碳骨架的多样性 → 生物分子种类繁多、功能多样。",
  "生物分子是碳化合物，但并非所有含碳化合物都是生物分子。"],
 "强调碳的“四价”和“成链/成环”能力。官能团决定化学性质，碳骨架决定分子形状；二者共同决定生物分子功能。",
 v_functional)
content_slide(14, "三、生命物质的化学组成 · 3/4", "（三）生物小分子与代谢物", 24,
 ["生物小分子：氨基酸、核苷酸、单糖、脂肪酸、甘油等。",
  "代谢物是代谢途径中的中间产物和终产物。",
  "小分子既是生物大分子的构件，也是能量代谢和信号分子。",
  "单体通过共价聚合形成大分子；代谢物通过酶促反应相互转化。",
  "“小分子”不是“不重要”，ATP、NAD⁺、辅酶A 都是小分子。"],
 "把“构件”和“代谢物”两个身份讲清：同一个分子既可能是单体，也可能是代谢中间物。举例：乙酰CoA 是小分子，却是代谢枢纽。",
 None)
content_slide(15, "三、生命物质的化学组成 · 4/4", "（四）生物大分子及其单体亚基", 25,
 ["蛋白质—氨基酸；核酸—核苷酸；多糖—单糖。",
  "生物大分子通常指相对分子质量大于5000的多聚体。",
  "单体排列顺序决定大分子的一级结构，进而决定高级结构和功能。",
  "脂质不是典型多聚体，但其聚集体具有大分子性质，称“准生物大分子”。",
  "生物大分子是生命结构与功能的分子基础。"],
 "这一页是全章高频考点：三大类生物大分子与单体对应关系、5000这个界限、脂质的特殊地位。强调“序列决定结构，结构决定功能”。",
 None)
# Section 4
section_slide(16, "生物分子的三维结构：构型与构象", 26, ["构型与构象的本质区别是什么？", "手性碳和对映体如何判断？", "生物分子间的识别为什么是立体专一的？"], 4)
content_slide(17, "四、生物分子的三维结构 · 1/6", "（一）生物分子的大小", 26,
 ["小分子：相对分子质量几十到几百；生物大分子：通常 >5000。",
  "蛋白质：约 10⁴—10⁶；核酸：可达 10⁶—10⁹。",
  "核糖体、病毒、细胞器属于更大的超分子复合体。",
  "生物分子的大小跨越多个数量级，需用 nm—µm 尺度理解。",
  "大小不是唯一标准，结构层次和功能更重要。"],
 "用数量级建立尺度感：小分子Å—nm，蛋白质nm，细胞µm。提醒学生：生物大分子不是“越大越好”，功能取决于结构和相互作用。",
 None)
content_slide(18, "四、生物分子的三维结构 · 2/6", "（二）立体异构与构型", 26,
 ["结构异构：原子连接次序不同。",
  "立体异构：连接次序相同，空间排列不同。",
  "构型包括对映体、非对映体和顺反异构。",
  "构型改变必须断裂共价键，不能靠单键旋转互变。",
  "手性碳：连接四个不同基团的碳原子。"],
 "先区分“连接次序”与“空间排列”，再讲构型。判断手性碳只看四个基团是否不同，不看碳上连了几根键。构型是“换了就回不去”的稳定空间排列。",
 v_stereo)
content_slide(19, "四、生物分子的三维结构 · 3/6", "（三）构型的命名系统", 29,
 ["D/L 命名：以甘油醛为参照，主要用于氨基酸和糖。",
  "L-氨基酸、D-单糖是生物体常用的构型。",
  "R/S 命名：按 CIP 规则给四个基团排优先级，再看空间方向。",
  "D/L 与 R/S 是两套独立系统，不能简单等同。",
  "构型命名必须结合具体分子，不能只看旋光方向。"],
 "易错点：D/L 不是“左/右旋”，而是相对甘油醛的构型；R/S 按原子序数排优先级。L-氨基酸与D-单糖是常见记忆点。",
 v_glyceraldehyde)
content_slide(20, "四、生物分子的三维结构 · 4/6", "（四）构象", 29,
 ["构象：同一构型下，通过单键旋转产生的不同空间形态。",
  "构象改变不需要断裂共价键，通常可快速互变。",
  "例子：乙烷的交叉式/重叠式；蛋白质的α螺旋、β折叠。",
  "构象决定生物大分子的形状、活性和识别能力。",
  "构型≠构象：前者需断键，后者只需旋转。"],
 "用“构型换了回不去，构象转一下还能回来”帮助记忆。蛋白质的α螺旋、β折叠都是构象/二级结构，不是构型。这个区分是考试高频易错点。",
 None)
content_slide(21, "四、生物分子的三维结构 · 5/6", "（五）三维结构的分子模型", 30,
 ["球棍模型：显示原子位置和化学键，适合看连接关系。",
  "空间填充模型：显示原子实际占据空间，适合看分子外形。",
  "飘带/卡通模型：显示蛋白质二级结构和折叠，适合看整体构象。",
  "X射线晶体学、NMR、冷冻电镜是主要结构测定方法。",
  "模型是理解结构的工具，不同模型强调不同信息。"],
 "让学生建立“模型目的不同”的意识：球棍看键，空间填充看体积，飘带看折叠。结构测定方法只需知道X射线、NMR、冷冻电镜三者的基本特点。",
 v_models)
content_slide(22, "四、生物分子的三维结构 · 6/6", "（六）生物分子间的相互作用是立体专一的", 31,
 ["酶与底物、受体与配体、抗体与抗原的结合都具有立体专一性。",
  "结合位点的形状、电荷和氢键供受体必须互补。",
  "手性识别是生物专一性的基础：对映体生物活性常不同。",
  "药物对映体可能一个有效、另一个无效甚至有害。",
  "立体专一性来自非共价相互作用的精确空间匹配。"],
 "从“锁钥/诱导契合”过渡到立体专一性。强调不是“形状像”就行，还要求电荷、氢键和手性匹配。药物手性例子能说明其现实意义。",
 None)
# Section 5
section_slide(23, "生物系统中的非共价相互作用", 31, ["生物系统中主要有哪四类非共价作用？", "它们为什么能稳定生物大分子？", "疏水作用为什么常是折叠的主要驱动力？"], 5)
content_slide(24, "五、非共价相互作用 · 1/5", "（一）离子相互作用", 32,
 ["带电基团之间的静电吸引或排斥。",
  "蛋白质中：Asp/Glu 的 –COO⁻ 与 Lys/Arg 的 –NH₃⁺。",
  "DNA中：磷酸骨架的负电荷与组蛋白、Mg²⁺ 等相互作用。",
  "强度受介电常数、离子强度和环境 pH 影响。",
  "在water中较强，但可被高盐竞争削弱。"],
 "离子键本质是静电作用。注意环境：水的高介电常数会削弱静电吸引；高盐会屏蔽电荷。蛋白质盐析、离子交换层析都与离子作用有关。",
 v_ionpair)
content_slide(25, "五、非共价相互作用 · 2/5", "（二）氢键", 32,
 ["氢键：与电负性原子（N、O、F）共价结合的 H，与另一个电负性原子的孤对电子作用。",
  "具有方向性和饱和性，强度介于共价键和范德华力之间。",
  "DNA碱基配对：A=T 两个氢键，G≡C 三个氢键。",
  "蛋白质二级结构：α螺旋、β折叠靠主链氢键稳定。",
  "水的高热容、高沸点也与氢键网络有关。"],
 "氢键不是“氢原子之间的键”，而是H与孤对电子之间的作用。方向性使氢键具有立体专一性；G≡C三个氢键解释DNA稳定性差异。",
 None)
content_slide(26, "五、非共价相互作用 · 3/5", "（三）范德华力", 32,
 ["瞬时偶极诱导邻近原子产生偶极，形成微弱吸引。",
  "包括取向力、诱导力、色散力；色散力最普遍。",
  "作用距离很短，要求原子间紧密堆积。",
  "在疏水核心、分子内部堆积中贡献显著。",
  "单个很弱，但数量多时总贡献可观。"],
 "范德华力常被忽略，但大量紧密堆积的原子间作用对结构稳定很重要。强调“距离短、数量多”的特点。",
 None)
content_slide(27, "五、非共价相互作用 · 4/5", "（四）疏水相互作用", 33,
 ["非极性基团在水中有避水聚集的倾向。",
  "本质不是“非极性基团相互吸引”，而是水分子氢键网络被破坏后的熵驱动。",
  "蛋白质折叠时，疏水侧链埋入内部，极性侧链暴露表面。",
  "疏水作用是膜脂双层形成、蛋白质折叠的重要驱动力。",
  "去垢剂、有机溶剂可破坏疏水作用，导致蛋白质变性。"],
 "这是最容易被误解的一条：疏水作用不是“油水相斥”的简单吸引，而是水分子为了维持氢键网络产生的熵驱动。用“熵驱动”四个字点明本质。",
 None)
content_slide(28, "五、非共价相互作用 · 5/5", "（五）非共价相互作用对大分子的结构和功能是关键的", 34,
 ["单个非共价作用弱，但数量多、可协同，形成稳定结构。",
  "它们允许生物大分子在温和条件下组装和解离。",
  "决定蛋白质折叠、核酸双螺旋、膜结构和分子识别。",
  "也是酶催化、信号转导、免疫识别的基础。",
  "非共价作用“弱而多”，是生命动态性的基础。"],
 "这一页总结四类作用的分工：离子键—电荷；氢键—方向性识别；范德华—紧密堆积；疏水—折叠驱动力。强调“弱”不是缺点，而是可逆、可调控的前提。",
 v_table)
# Section 6
section_slide(29, "水和生命", 34, ["水有哪些特殊性质？", "水为什么是生命介质？", "弱酸弱碱和缓冲如何影响生物分子？"], 6)
content_slide(30, "六、水和生命 · 1/4", "（一）水的结构和性质", 34,
 ["水是极性分子，O—H 键极性，分子间形成氢键。",
  "每个水分子最多形成4个氢键，形成动态网络。",
  "结果：高沸点、高熔点、高蒸发热、高热容、高表面张力。",
  "水的高介电常数使离子化合物易溶解。",
  "水既是溶剂，也是许多生化反应的底物或产物。"],
 "把“结构—性质—功能”连起来：极性→氢键→高热容/高沸点→缓冲温度、支持生命。水的高热容是重点。",
 v_water)
content_slide(31, "六、水和生命 · 2/4", "（二）水是生命的介质", 35,
 ["水占细胞质量的大部分，是物质运输和反应的介质。",
  "极性分子、离子易溶于水；非极性分子难溶。",
  "水参与水解、水合、氧化还原等反应。",
  "水的流动性支持扩散、渗透和胞内运输。",
  "水的高热容帮助生物体缓冲温度变化。"],
 "水不是惰性背景，而是积极参与生命过程。强调“介质+反应物+温度缓冲”三重角色。",
 None)
content_slide(32, "六、水和生命 · 3/4", "（三）溶质影响水的依数性", 36,
 ["依数性：只与溶质粒子数有关，与粒子性质关系较小。",
  "包括蒸汽压下降、沸点升高、凝固点降低、渗透压。",
  "渗透压对细胞体积、水分分布至关重要。",
  "生物体通过渗透调节维持细胞内外水平衡。",
  "大分子溶液也产生渗透压，但单位质量贡献较小。"],
 "依数性只数“粒子数”，不数“粒子种类”。联系医学：输液必须考虑渗透压；细胞在低渗溶液中吸水胀破，在高渗溶液中皱缩。",
 None)
content_slide(33, "六、水和生命 · 4/4", "（四）水、弱酸和弱碱的电离", 37,
 ["水可微弱电离：H₂O ⇌ H⁺ + OH⁻；pH = −lg[H⁺]。",
  "弱酸解离：HA ⇌ H⁺ + A⁻；Ka = [H⁺][A⁻]/[HA]。",
  "Henderson-Hasselbalch：pH = pKa + lg([A⁻]/[HA])。",
  "当[A⁻]=[HA]时，pH = pKa；此时缓冲能力较强。",
  "蛋白质、核酸的电荷状态随pH改变；等电点时净电荷为零。"],
 "这一页是计算题基础。强调pH、pKa、缓冲范围三者关系；缓冲液在pKa±1范围内最有效。等电点pI是蛋白质/氨基酸净电荷为零的pH。",
 v_ph)
# Section 7
section_slide(34, "生命的基本单位——细胞", 42, ["生物三域如何划分？", "为什么细胞是生命的基本单位？", "原核细胞与真核细胞有何区别？"], 7)
content_slide(35, "七、生命的基本单位——细胞 · 1/4", "（一）生物三域——细菌、古菌和真核生物", 42,
 ["三域：细菌（Bacteria）、古菌（Archaea）、真核生物（Eukarya）。",
  "细菌和古菌是原核生物，无核膜包被的细胞核。",
  "古菌在分子特征上更接近真核生物，常生活在极端环境。",
  "真核生物有核膜、细胞核和多种膜性细胞器。",
  "三域分类基于分子系统发生，而不仅是形态。"],
 "三域不是“细菌、植物、动物”，而是细菌、古菌、真核。古菌常被误认为细菌，要强调其分子特征和极端环境适应性。",
 v_tree)
content_slide(36, "七、生命的基本单位——细胞 · 2/4", "（二）细胞是所有生物的结构和功能单位", 43,
 ["细胞是生命结构和功能的基本单位。",
  "单细胞生物一个细胞即个体；多细胞生物由许多细胞协作。",
  "细胞膜界定边界，维持内环境相对稳定。",
  "细胞体积小，有利于物质交换；表面积/体积比是关键。",
  "细胞学说：所有生物由细胞构成，细胞来自已有细胞。"],
 "从“边界、代谢、遗传、繁殖”四个角度说明细胞为什么是基本单位。表面积/体积比解释为什么细胞不能无限大。",
 None)
content_slide(37, "七、生命的基本单位——细胞 · 3/4", "（三）原核细胞的代表——大肠杆菌", 44,
 ["大肠杆菌：革兰氏阴性、杆状、无核膜。",
  "遗传物质：环状DNA，位于拟核；可有质粒。",
  "核糖体：70S（50S+30S），是蛋白质合成场所。",
  "细胞膜外有细胞壁；部分菌株有鞭毛、菌毛。",
  "代谢多样、繁殖快，是分子生物学模式生物。"],
 "以大肠杆菌为原核代表，记住“无核膜、拟核、70S核糖体、环状DNA、质粒”。这些也是抗生素作用靶点的背景。",
 v_ecoli)
content_slide(38, "七、生命的基本单位——细胞 · 4/4", "（四）真核细胞含有多种膜性细胞器", 45,
 ["细胞核：遗传信息储存、转录场所。",
  "线粒体：氧化磷酸化、ATP合成；有自身DNA。",
  "内质网：蛋白质/脂质合成与加工；粗面ER有核糖体。",
  "高尔基体：加工、分选和运输；溶酶体：降解。",
  "叶绿体：光合作用；植物细胞还有细胞壁、液泡。"],
 "按“遗传信息—能量—合成加工—运输降解”讲细胞器，避免孤立记忆。线粒体和叶绿体有自身DNA，为内共生起源提供证据。",
 v_euk)
# Section 8
section_slide(39, "生物分子的起源与进化", 46, ["生物分子如何通过化学进化产生？", "为什么说RNA可能是最早的基因和催化剂？", "真核细胞如何从原核细胞进化而来？"], 8)
content_slide(40, "八、生物分子的起源与进化 · 1/4", "（一）最初的生物分子是通过化学进化产生的", 47,
 ["早期地球：还原性大气、强紫外线、雷电、火山活动。",
  "无机小分子 → 有机小分子（氨基酸、核苷酸、单糖等）。",
  "米勒-尤里实验：模拟早期地球条件，可生成多种氨基酸。",
  "有机小分子在合适条件下聚合为生物大分子。",
  "化学进化是生命起源的前奏，不等于生命已经出现。"],
 "区分“化学进化”和“生物进化”：前者是无机到有机、小分子到大分子；后者是细胞出现后的繁殖、变异和选择。米勒实验是经典证据，但不是生命起源的完整答案。",
 None)
content_slide(41, "八、生物分子的起源与进化 · 2/4", "（二）RNA或相关前体可能是最早的基因和催化剂", 48,
 ["RNA既能贮存遗传信息，又能催化反应（核酶）。",
  "“RNA世界”假说：早期生命可能以RNA同时承担信息与催化功能。",
  "核酶的存在说明“核酸只是信息分子”不成立。",
  "后来DNA更适合长期贮存信息，蛋白质催化效率更高。",
  "RNA世界是假说，不是已经证实的唯一路径。"],
 "RNA世界的核心逻辑：信息+催化双重功能，解决“先有DNA还是先有蛋白质”的循环。核酶是支持证据；要说明这是假说，不是定论。",
 v_rnaworld)
content_slide(42, "八、生物分子的起源与进化 · 3/4", "（三）原始细胞的出现——生物学进化的开始", 49,
 ["生物大分子聚集成膜包被的原始细胞（protocell）。",
  "膜将内部环境与外界分隔，形成相对独立的化学系统。",
  "原始细胞能摄取物质、进行反应、复制遗传信息。",
  "自然选择开始作用于可复制的原始细胞。",
  "原始细胞出现是化学进化走向生物进化的关键一步。"],
 "强调“边界”的意义：没有膜，分子扩散稀释，无法形成稳定的选择单位。原始细胞不必很复杂，关键是能自我维持和复制。",
 None)
content_slide(43, "八、生物分子的起源与进化 · 4/4", "（四）真核细胞是从原核细胞进化而来的", 50,
 ["内共生学说：线粒体源于被吞入的需氧细菌。",
  "叶绿体源于被吞入的蓝细菌；两者都有自身DNA和双层膜。",
  "真核细胞获得高效产能和光合作用能力。",
  "细胞核、内膜系统、细胞骨架逐渐形成。",
  "内共生学说的证据：DNA、核糖体、双层膜、分裂方式。"],
 "内共生学说是重点：线粒体/叶绿体有自身环状DNA、70S核糖体、双层膜，支持其细菌起源。不要把“内共生”误解为“细胞吞了一个细胞器”。",
 v_endosym)
# Summary
s = prs.slides.add_slide(BLANK); set_bg(s, LIGHT)
box(s, 0, 0, 13.333, 0.10, fill=GOLD, line=None, radius=False)
kicker(s, "本章总结", 50)
text(s, 0.62, 0.78, 12.0, 0.55, "第1章 · 三条主线与高频考点", size=27, color=NAVY, bold=True, font=SERIF)
cols = [
 ("主线一：生命是什么", ["生命属性三方面", "化学成分复杂而同一", "结构层次与自装配", "开放系统与能量流"]),
 ("主线二：生物分子", ["生命元素与碳骨架", "小分子/大分子/单体", "构型 vs 构象", "非共价作用四类"]),
 ("主线三：水、细胞、起源", ["水的性质与pH/缓冲", "原核/真核/三域", "化学进化与RNA世界", "内共生起源"]),
]
for i, (h, items) in enumerate(cols):
    x = 0.62 + i*4.18
    box(s, x, 1.65, 3.95, 3.50, fill=WHITE, line=LINE)
    box(s, x, 1.65, 3.95, 0.08, fill=[BLUE, TEAL, GOLD][i], line=None, radius=False)
    text(s, x+0.22, 1.85, 3.5, 0.35, h, size=13.5, color=NAVY, bold=True)
    yy = 2.35
    for it in items:
        text(s, x+0.25, yy, 3.45, 0.4, "▸ " + it, size=12.2, color=INK)
        yy += 0.60
box(s, 0.62, 5.55, 12.0, 0.85, fill=RGBColor(0xE8,0xF0,0xF7), line=None)
text(s, 0.85, 5.70, 11.6, 0.5, "复习顺序：先框架（8节） → 再概念（生物大分子/构型/构象） → 后应用（水、pH、细胞、起源）", size=11.5, color=NAVY, bold=True)
footer(s, 44, 50)
s.notes_slide.notes_text_frame.text = "用三条主线收束全章：生命属性、生物分子（结构/构象/非共价作用）、水与细胞及起源。让学生先能复述框架，再补细节。"

# Quiz
s = prs.slides.add_slide(BLANK); set_bg(s, LIGHT)
box(s, 0, 0, 13.333, 0.10, fill=TEAL, line=None, radius=False)
kicker(s, "课堂自测", 50)
text(s, 0.62, 0.78, 12.0, 0.55, "5 道题检查本章掌握度", size=27, color=NAVY, bold=True, font=SERIF)
qs = [
 ("1", "下列属于微量元素的是（  ）", "A. C    B. H    C. Fe    D. O"),
 ("2", "蛋白质的α螺旋变为β折叠属于（  ）", "A. 结构异构    B. 构型改变    C. 构象改变    D. 共价键断裂"),
 ("3", "生物大分子通常指相对分子质量大于（  ）", "A. 500    B. 1000    C. 5000    D. 50000"),
 ("4", "水的高热容主要与（  ）有关", "A. 分子小    B. 氢键    C. 不含碳    D. 中性"),
 ("5", "原核细胞与真核细胞的主要区别是（  ）", "A. 有无DNA    B. 有无核膜    C. 有无核糖体    D. 有无细胞膜"),
]
for i, (no, q, opts) in enumerate(qs):
    y = 1.55 + i*1.02
    box(s, 0.62, y, 12.0, 0.86, fill=WHITE, line=LINE)
    text(s, 0.76, y+0.12, 0.35, 0.3, no, size=14, color=GOLD, bold=True)
    text(s, 1.12, y+0.10, 10.8, 0.28, q, size=11.5, color=INK, bold=True)
    text(s, 1.12, y+0.45, 10.8, 0.26, opts, size=11, color=MUTED)
footer(s, 45, 50)
s.notes_slide.notes_text_frame.text = "先让学生独立作答，再翻答案。题目只覆盖本章最基础的概念，不涉及后续章节细节。"

# Answers
s = prs.slides.add_slide(BLANK); set_bg(s, NAVY)
text(s, 0.85, 0.55, 8.0, 0.4, "课堂自测 · 答案与复习清单", size=12, color=GOLD, bold=True)
text(s, 0.85, 1.05, 10.0, 0.65, "答案", size=34, color=WHITE, bold=True, font=SERIF)
ans = [("1", "C", "Fe 属于微量元素；C、H、O 是大量元素。"),
       ("2", "C", "α螺旋、β折叠都是二级结构，通过单键旋转互变，属于构象改变。"),
       ("3", "C", "教材以 5000 作为生物大分子的常用界限。"),
       ("4", "B", "水分子间大量氢键吸收热量，使水具有高热容。"),
       ("5", "B", "原核细胞无核膜包被的细胞核；两者都有DNA、核糖体和细胞膜。")]
y = 1.9
for no, a, e in ans:
    text(s, 0.95, y, 0.35, 0.3, no, size=14, color=GOLD, bold=True)
    text(s, 1.35, y, 0.65, 0.35, a, size=16, color=WHITE, bold=True)
    text(s, 2.1, y, 9.9, 0.38, e, size=11.2, color=RGBColor(0xEE,0xF4,0xF8))
    y += 0.72
text(s, 0.95, 5.55, 11.2, 0.32, "复习清单", size=12, color=GOLD, bold=True)
text(s, 0.95, 5.9, 11.4, 1.3,
     "① 生命属性三方面  ② 生物大分子与单体  ③ 构型/构象  ④ 四类非共价作用  ⑤ 水与pH/缓冲  ⑥ 原核/真核  ⑦ RNA世界与内共生",
     size=10.5, color=WHITE)
text(s, 11.8, 6.9, 0.8, 0.3, "46", size=9.5, color=RGBColor(0x9E,0xB6,0xC8), align=PP_ALIGN.RIGHT)

# ---------- appendix: calculation ----------
s = prs.slides.add_slide(BLANK); set_bg(s, NAVY)
box(s, 0, 0, 0.14, 7.5, fill=GOLD, line=None, radius=False)
text(s, 0.95, 0.95, 9.0, 0.35, "APPENDIX · CALCULATION", size=11, color=GOLD, bold=True)
text(s, 0.95, 1.55, 10.5, 0.9, "附录：计算题专项", size=40, color=WHITE, bold=True, font=SERIF)
text(s, 0.98, 2.75, 10.8, 0.4, "原书第1章课后计算题涉及：溶液配制、稀释、pH、Ka、缓冲液。", size=15, color=RGBColor(0xD5,0xE3,0xEE))
for i, q in enumerate(["① 0.4 mol/L NaOH 500 mL 怎么配？", "② 浓盐酸稀释：37%、1.20 g/mL 怎么换算？", "③ 弱酸 pH=4.6，Ka、pKa 怎么求？", "④ 醋酸缓冲液 pH 与配比怎么算？"]):
    text(s, 1.05, 3.55 + i*0.58, 10.8, 0.35, q, size=13, color=WHITE)
text(s, 0.98, 6.7, 9.0, 0.3, "计算题不靠背答案，抓住“单位—物质的量—守恒”三条线。", size=9.5, color=RGBColor(0x9E,0xB6,0xC8))
text(s, 11.8, 6.85, 0.8, 0.3, "47", size=9.5, color=RGBColor(0x9E,0xB6,0xC8), align=PP_ALIGN.RIGHT)

s = prs.slides.add_slide(BLANK); set_bg(s, LIGHT)
box(s, 0, 0, 13.333, 0.10, fill=BLUE, line=None, radius=False)
kicker(s, "附录 · 计算题专项", 0)
text(s, 0.62, 0.78, 12.0, 0.55, "计算工具箱：四个公式解决本章计算", size=25, color=NAVY, bold=True, font=SERIF)
cards = [
 ("物质的量", "n = cV", "m = nM", "V 用 L；m 用 g"),
 ("质量分数→浓度", "c = 1000ρw / M", "ρ: g/mL；w: 质量分数", "浓盐酸、浓硫酸换算"),
 ("渗透压浓度", "i × c", "i=粒子数", "NaOH: i=2；葡萄糖: i=1"),
 ("pH 与缓冲", "pH = −lg[H⁺]", "pH = pKa + lg([A⁻]/[HA])", "弱酸 Ka、缓冲配比"),
]
for i,(h,f1,f2,note) in enumerate(cards):
    col=i%2; row=i//2; x=0.62+col*6.25; y=1.75+row*2.25
    box(s,x,y,5.95,1.95,fill=WHITE,line=LINE)
    box(s,x,y,0.08,1.95,fill=[BLUE,TEAL,GOLD,RED][i],line=None,radius=False)
    text(s,x+0.28,y+0.18,4.9,0.3,h,size=12.5,color=NAVY,bold=True)
    text(s,x+0.28,y+0.58,5.3,0.35,f1,size=17,color=[BLUE,TEAL,GOLD,RED][i],bold=True)
    text(s,x+0.28,y+1.02,5.3,0.3,f2,size=13,color=INK)
    text(s,x+0.28,y+1.45,5.3,0.28,note,size=10.2,color=INK)
footer(s,48,0)
s.notes_slide.notes_text_frame.text = "本章计算题核心是单位换算和守恒：n=cV、m=nM、c=1000ρw/M、pH与Henderson-Hasselbalch。先把已知量换成统一单位，再代公式。"

s = prs.slides.add_slide(BLANK); set_bg(s, LIGHT)
box(s, 0, 0, 13.333, 0.10, fill=GOLD, line=None, radius=False)
kicker(s, "附录 · 计算题专项", 0)
text(s, 0.62, 0.78, 12.0, 0.55, "例题：溶液配制、pH 与缓冲液", size=25, color=NAVY, bold=True, font=SERIF)
# left column
box(s, 0.62, 1.55, 6.05, 4.95, fill=WHITE, line=LINE)
text(s, 0.85, 1.72, 5.6, 0.32, "例1  NaOH 与浓盐酸稀释", size=13, color=BLUE, bold=True)
text(s, 0.85, 2.12, 5.55, 4.1,
     "① 0.4 mol/L NaOH 500 mL\nn=0.4×0.500=0.20 mol\nm=0.20×40=8.0 g\n质量浓度=8.0÷0.500=16 g/L\n渗透压浓度=2×0.4=0.8 mol/L\n\n② 浓盐酸 37%、1.20 g/mL\nc=1000×1.20×0.37÷36.5\n  ≈12.16 mol/L\nn(HCl)=0.5×2.00=1.00 mol\nV=1.00÷12.16≈0.0822 L=82.2 mL",
     size=11.3, color=INK, line_spacing=1.12)
# right column
box(s, 6.90, 1.55, 5.80, 4.95, fill=WHITE, line=LINE)
text(s, 7.12, 1.72, 5.3, 0.32, "例2  弱酸 pH 与醋酸缓冲液", size=13, color=TEAL, bold=True)
text(s, 7.12, 2.12, 5.3, 4.1,
     "① 0.02 mol/L 弱酸，pH=4.6\n[H⁺]=10⁻⁴·⁶≈2.51×10⁻⁵ mol/L\nKa≈[H⁺]²/(c−[H⁺])\n  ≈(2.51×10⁻⁵)²/0.02\n  ≈3.15×10⁻⁸\npKa≈7.50\n\n② 0.20 mol/L NaAc + 0.60 mol/L HAc\npH=4.76+lg(0.20/0.60)\n  =4.76−0.477≈4.28\n\n③ pH=5.0 的 0.2 mol/L 缓冲液\nx/y=10⁰·²⁴≈1.74，x+y=0.2\nx≈0.127，y≈0.073 mol/L",
     size=11.1, color=INK, line_spacing=1.10)
footer(s,49,0)
s.notes_slide.notes_text_frame.text = "例1强调配制与稀释：先求物质的量，再换算质量或体积。例2强调pH、Ka和缓冲液：先写平衡式/方程，再代入数据；缓冲液还可用总浓度守恒解方程组。"

# Save
prs.save(PPTX)
print("SAVED", PPTX, "slides", len(prs.slides))
