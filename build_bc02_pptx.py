#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""bc02 AI PPT (bc01 style): amino acids, peptides, proteins."""
import os
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE

ROOT=os.path.dirname(os.path.abspath(__file__)); OUT=os.path.join(ROOT,'downloads'); os.makedirs(OUT,exist_ok=True)
PPTX=os.path.join(OUT,'bc02-ai-style.pptx')
NAVY=RGBColor(0x16,0x2E,0x47); NAVY2=RGBColor(0x1E,0x3F,0x5C); BLUE=RGBColor(0x1F,0x6F,0xA8)
TEAL=RGBColor(0x16,0x8B,0x8B); GOLD=RGBColor(0xC9,0x96,0x2E); INK=RGBColor(0x22,0x30,0x3C)
MUTED=RGBColor(0x5D,0x6D,0x7E); LIGHT=RGBColor(0xF4,0xF7,0xFA); PALE=RGBColor(0xE8,0xF0,0xF7)
WHITE=RGBColor(0xFF,0xFF,0xFF); LINE=RGBColor(0xD5,0xE1,0xEC); RED=RGBColor(0x9B,0x3A,0x32)
FONT='Noto Sans CJK SC'; SERIF='Noto Serif CJK SC'
prs=Presentation(); prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5); BLANK=prs.slide_layouts[6]

def bg(s,c):
 s.background.fill.solid(); s.background.fill.fore_color.rgb=c
def box(s,x,y,w,h,fill=None,line=None,radius=True):
 sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,Inches(x),Inches(y),Inches(w),Inches(h))
 sh.shadow.inherit=False
 if fill is None: sh.fill.background()
 else: sh.fill.solid(); sh.fill.fore_color.rgb=fill
 if line is None: sh.line.fill.background()
 else: sh.line.color.rgb=line; sh.line.width=Pt(1)
 sh.text_frame.clear(); sh.text_frame.word_wrap=True
 return sh
def text(s,x,y,w,h,content,size=18,color=INK,bold=False,font=FONT,align=PP_ALIGN.LEFT,valign=MSO_ANCHOR.TOP,line_spacing=1.12):
 tb=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=tb.text_frame; tf.word_wrap=True; tf.vertical_anchor=valign
 tf.margin_left=Inches(.02);tf.margin_right=Inches(.02);tf.margin_top=Inches(.02);tf.margin_bottom=Inches(.02)
 if isinstance(content,str): content=[content]
 for i,item in enumerate(content):
  p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); p.alignment=align; p.line_spacing=line_spacing
  r=p.add_run(); r.text=str(item); r.font.size=Pt(size); r.font.bold=bold; r.font.color.rgb=color; r.font.name=font
 return tb
def bullets(s,x,y,w,h,items,size=18,gap=8):
 tb=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=tb.text_frame; tf.word_wrap=True
 tf.margin_left=Inches(.02);tf.margin_right=Inches(.02);tf.margin_top=Inches(.02);tf.margin_bottom=Inches(.02)
 for i,it in enumerate(items):
  p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); p.space_after=Pt(gap); p.line_spacing=1.1
  r=p.add_run(); r.text='▸ '+str(it); r.font.size=Pt(size); r.font.color.rgb=INK; r.font.name=FONT
 return tb
def kicker(s,txt,page):
 text(s,.62,.36,10.2,.3,txt,size=10.5,color=BLUE,bold=True)
 text(s,11.4,.36,1.3,.3,'原书 p%d'%page if page else '',size=10,color=MUTED,align=PP_ALIGN.RIGHT)
def footer(s,num,page):
 ln=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Inches(.62),Inches(7.05),Inches(12.72),Inches(7.05));ln.line.color.rgb=LINE;ln.line.width=Pt(.8)
 text(s,.62,7.1,8.5,.24,'朱圣庚《生物化学》第4版 · 第2章 氨基酸、肽和蛋白质 · AI讲义',size=8.5,color=MUTED)
 text(s,11.8,7.1,.8,.24,'%02d'%num,size=9,color=MUTED,align=PP_ALIGN.RIGHT)

def section(num,no,title,page,questions):
 s=prs.slides.add_slide(BLANK);bg(s,NAVY);box(s,0,0,.12,7.5,fill=GOLD,radius=False)
 text(s,.85,.75,2.0,.4,'SECTION %02d'%no,size=11,color=GOLD,bold=True)
 text(s,.85,1.2,2.4,1.2,'%02d'%no,size=88,color=WHITE,bold=True,font=SERIF)
 text(s,3.05,1.45,9.2,1.2,title,size=34,color=WHITE,bold=True,font=SERIF)
 text(s,3.08,2.78,8.4,.4,'原书节级目录 · 第 %d 节'%no,size=12.5,color=RGBColor(0xB8,0xCE,0xDD))
 text(s,3.08,3.30,8.4,.36,'本节先回答',size=11.5,color=GOLD,bold=True)
 y=3.75
 for q in questions:
  box(s,3.08,y,.13,.13,fill=GOLD,radius=False);text(s,3.34,y-.06,8.2,.4,q,size=17.5,color=WHITE);y+=.52
 text(s,.9,6.85,8,.3,'朱圣庚《生物化学》第4版 · 原书 p%d'%page,size=9.5,color=RGBColor(0x9E,0xB6,0xC8))
 text(s,11.8,6.85,.8,.3,'%02d'%num,size=9.5,color=RGBColor(0x9E,0xB6,0xC8),align=PP_ALIGN.RIGHT)

def fig_reaction(num):
 s=prs.slides.add_slide(BLANK);bg(s,LIGHT);box(s,0,0,13.333,.1,fill=GOLD,radius=False)
 kicker(s,'第4节 · 反应式导读',65);text(s,.62,.72,12,.55,'图2-7 甘氨酸缩合与水解：怎么读反应式',size=25,color=NAVY,bold=True,font=SERIF)
 # panels
 for i,(h,b) in enumerate([('① 两分子甘氨酸','H₂N—CH₂—COOH\n+ H₂N—CH₂—COOH'),('② 二酮吡嗪','六元环，含两个肽键\n（环状二肽）'),('③ 双甘二肽','H₂N—CH₂—CO—NH—CH₂—COOH\n（线性二肽）')]):
  x=.7+i*4.15;box(s,x,1.55,3.75,2.2,fill=WHITE,line=LINE);text(s,x+.18,1.7,3.4,.3,h,size=13,color=NAVY,bold=True);text(s,x+.18,2.15,3.4,1.4,b,size=10.5,color=INK,align=PP_ALIGN.CENTER)
  if i<2: text(s,x+3.78,2.35,.35,.3,'→',size=20,color=GOLD,bold=True)
 text(s,4.0,1.35,1.5,.3,'加热/乙二醇',size=9,color=GOLD,bold=True,align=PP_ALIGN.CENTER)
 text(s,8.2,1.35,1.5,.3,'煮沸/浓HCl',size=9,color=GOLD,bold=True,align=PP_ALIGN.CENTER)
 # steps
 box(s,.7,4.15,11.9,1.95,fill=WHITE,line=LINE);text(s,.9,4.28,11.5,.28,'读反应式四步',size=11,color=TEAL,bold=True)
 text(s,.9,4.62,11.5,.35,'① 找反应物和条件 → ② 圈出被脱去的 —OH 和 —H（2 H₂O）→ ③ 数环状/线性产物中的肽键 → ④ 对比体内酶促合成',size=10.2,color=INK)
 text(s,.9,5.15,11.5,.7,'重点：缩合是“脱2分子水成环”；水解是“加2分子水开环”。环状二酮吡嗪有2个肽键，线性双甘二肽有1个肽键。',size=10.2,color=MUTED)
 text(s,.9,5.95,11.5,.5,'体内不是简单加热缩合：需要酶、活化的羧基和能量，并防止肽键被水解。',size=10.5,color=GOLD,bold=True)
 footer(s,num,65);s.notes_slide.notes_text_frame.text='图2-7反应式读法：两分子甘氨酸缩合脱2水成环状二酮吡嗪；浓盐酸水解开环成双甘二肽。反应式的核心是原子记账和肽键数。'

def content(num,sec_label,title,page,items,keypoint,pitfall,notes):
 s=prs.slides.add_slide(BLANK);bg(s,LIGHT);box(s,0,0,13.333,.1,fill=BLUE,radius=False)
 kicker(s,sec_label,page);text(s,.62,.78,7.8,.72,title,size=30,color=NAVY,bold=True,font=SERIF)
 bullets(s,.7,1.70,7.3,4.85,items,size=22.0,gap=8)
 box(s,8.35,1.55,4.35,2.25,fill=WHITE,line=LINE);text(s,8.55,1.70,3.95,.28,'本页考点',size=12,color=TEAL,bold=True);text(s,8.55,2.02,3.95,1.55,keypoint,size=14.0,color=INK)
 box(s,8.35,3.98,4.35,1.60,fill=RGBColor(0xFB,0xF3,0xDE),line=None);text(s,8.55,4.10,3.95,.26,'易错点',size=12,color=RGBColor(0x7A,0x4A,0x10),bold=True);text(s,8.55,4.38,3.95,.95,pitfall,size=13.0,color=RGBColor(0x6B,0x3E,0x0E))
 footer(s,num,page);s.notes_slide.notes_text_frame.text=notes

def wrap(num,no,title,sentences,questions,pitfall,notes):
 s=prs.slides.add_slide(BLANK);bg(s,LIGHT);box(s,0,0,13.333,.1,fill=GOLD,radius=False)
 kicker(s,'第 %d 节 · 本节收束'%no,0);text(s,.62,.72,12,.55,'第%d节 · 3句话 + 2题自测'%no,size=26,color=NAVY,bold=True,font=SERIF)
 box(s,.62,1.42,7.55,4.55,fill=WHITE,line=LINE);text(s,.84,1.58,7.1,.28,'三句话记住',size=11,color=BLUE,bold=True)
 y=1.98
 for i,x in enumerate(sentences):
  text(s,.86,y,.35,.28,str(i+1),size=11,color=GOLD,bold=True);text(s,1.20,y-.02,6.75,1.05,x,size=16.0,color=INK);y+=1.32
 box(s,8.35,1.42,4.35,4.55,fill=WHITE,line=LINE);text(s,8.55,1.58,3.95,.28,'2 题自测',size=10.5,color=RED,bold=True)
 qy=1.95
 for j,(q,a) in enumerate(questions):
  text(s,8.55,qy,3.95,.55,'Q%d %s'%(j+1,q),size=12.0,color=INK,bold=True);text(s,8.55,qy+.55,3.95,.48,'答：'+a,size=11.0,color=MUTED);qy+=1.15
 box(s,.62,6.1,12.08,.62,fill=RGBColor(0xFB,0xF3,0xDE),line=None);text(s,.82,6.22,11.7,.36,'⚠ 易错点：'+pitfall,size=13.0,color=RGBColor(0x7A,0x4A,0x10),bold=True)
 footer(s,num,0);s.notes_slide.notes_text_frame.text=notes

def breakdown(num,no,title,steps,example,pitfall,notes):
 s=prs.slides.add_slide(BLANK);bg(s,LIGHT);box(s,0,0,13.333,.1,fill=RED,radius=False)
 kicker(s,'第 %d 节 · 难点拆解'%no,0);text(s,.62,.72,12,.55,title,size=24,color=NAVY,bold=True,font=SERIF)
 for i,(h,b) in enumerate(steps):
  x=.62+i*4.12;box(s,x,1.48,3.92,2.45,fill=WHITE,line=LINE);box(s,x,1.48,3.92,.08,fill=[BLUE,TEAL,GOLD][i],radius=False)
  text(s,x+.20,1.64,.4,.3,str(i+1),size=15,color=[BLUE,TEAL,GOLD][i],bold=True);text(s,x+.66,1.64,3.1,.3,h,size=11.2,color=NAVY,bold=True);text(s,x+.20,2.04,3.55,1.7,b,size=10.3,color=INK)
 box(s,.62,4.15,7.55,1.55,fill=WHITE,line=LINE);text(s,.82,4.28,7.1,.26,'例子 / 怎么用',size=10.4,color=TEAL,bold=True);text(s,.82,4.58,7.1,.95,example,size=9.8,color=INK)
 box(s,8.35,4.15,4.35,1.55,fill=RGBColor(0xFB,0xF3,0xDE),line=None);text(s,8.55,4.28,3.95,.26,'易错点',size=10.4,color=RGBColor(0x7A,0x4A,0x10),bold=True);text(s,8.55,4.58,3.95,.95,pitfall,size=9.4,color=RGBColor(0x6B,0x3E,0x0E))
 footer(s,num,0);s.notes_slide.notes_text_frame.text=notes

def fig_amino(num):
 s=prs.slides.add_slide(BLANK);bg(s,LIGHT);kicker(s,'第2章 · 图2-1 示意图',54);text(s,.62,.75,12,.55,'α-氨基酸的一般结构',size=27,color=NAVY,bold=True,font=SERIF)
 cx,cy=6.7,3.45
 pts=[(-2.4,-1.45,'H₃N⁺'),(2.4,-1.45,'COO⁻'),(-2.4,1.45,'H'),(2.4,1.45,'R')]
 for dx,dy,lab in pts:
  ln=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Inches(cx),Inches(cy),Inches(cx+dx),Inches(cy+dy));ln.line.color.rgb=BLUE;ln.line.width=Pt(2.4)
 box(s,cx-.34,cy-.34,.68,.68,fill=NAVY2,radius=True);text(s,cx-.34,cy-.17,.68,.3,'α-C',size=12,color=WHITE,bold=True,align=PP_ALIGN.CENTER)
 for dx,dy,lab in pts:
  text(s,cx+dx-1.0,cy+dy-.24,2.0,.4,lab,size=16,color=NAVY,bold=True,align=PP_ALIGN.CENTER)
 text(s,.8,5.7,11.8,.8,'α-碳连接 —NH₃⁺、—COO⁻、—H 和侧链 R；除甘氨酸外，α-碳是手性碳。',size=16.0,color=INK)
 text(s,.8,6.5,11.8,.3,'AI 示意图（非原书扫描）',size=8,color=MUTED)
 footer(s,num,54)

def fig_charge(num):
 s=prs.slides.add_slide(BLANK);bg(s,LIGHT);kicker(s,'第2章 · 等电点示意',62);text(s,.62,.75,12,.55,'pH 对氨基酸电荷状态的影响',size=27,color=NAVY,bold=True,font=SERIF)
 x0,y0,w=1.0,2.55,11.3
 for i in range(14):
  box(s,x0+i*w/14,y0,w/14+.01,.25,fill=(RED if i<3 else RGBColor(0xF0,0xC4,0x5E) if i<6 else TEAL if i<9 else BLUE if i<12 else RGBColor(0x7D,0x86,0xC8)),radius=False)
 for i in range(0,15,2): text(s,x0+i*w/14-.15,y0+.3,.35,.2,str(i),size=9,color=MUTED,align=PP_ALIGN.CENTER)
 cards=[('低 pH','—NH₃⁺ 占优；净电荷 +1',2.0,RED),('pI','净电荷=0；溶解度最低',6.0,TEAL),('高 pH','—COO⁻ 占优；净电荷 −1',10.2,BLUE)]
 for h,b,pos,col in cards:
  x=x0+pos*w/14-1.1
  box(s,x,4.0,2.4,1.55,fill=WHITE,line=LINE);text(s,x+.18,4.14,2.05,.3,h,size=11,color=col,bold=True);text(s,x+.18,4.48,2.05,.85,b,size=10,color=INK)
 text(s,.9,5.95,11.5,.6,'中性氨基酸 pI≈(pK₁+pK₂)/2；酸性/碱性氨基酸还要考虑侧链 pKR。',size=12,color=INK)
 footer(s,num,62)

def fig_peptide(num):
 s=prs.slides.add_slide(BLANK);bg(s,LIGHT);kicker(s,'第2章 · 肽键示意',74);text(s,.62,.75,12,.55,'肽键的形成与平面性',size=27,color=NAVY,bold=True,font=SERIF)
 text(s,1.0,2.2,3.2,.5,'氨基酸1',size=13,color=NAVY,bold=True);text(s,9.1,2.2,3.2,.5,'氨基酸2',size=13,color=NAVY,bold=True)
 box(s,1.0,2.8,3.3,1.1,fill=WHITE,line=LINE);text(s,1.2,3.05,2.9,.5,'H₂N—CH(R₁)—COOH',size=14,color=INK,align=PP_ALIGN.CENTER)
 box(s,8.9,2.8,3.3,1.1,fill=WHITE,line=LINE);text(s,9.1,3.05,2.9,.5,'H₂N—CH(R₂)—COOH',size=14,color=INK,align=PP_ALIGN.CENTER)
 ln=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Inches(4.3),Inches(3.35),Inches(8.9),Inches(3.35));ln.line.color.rgb=GOLD;ln.line.width=Pt(2.2)
 text(s,5.0,2.65,3.2,.35,'—CO—NH— 肽键',size=15,color=GOLD,bold=True,align=PP_ALIGN.CENTER)
 text(s,5.4,3.45,2.4,.3,'+ H₂O',size=11,color=MUTED,align=PP_ALIGN.CENTER)
 text(s,.9,4.8,11.5,.9,'肽键具有部分双键性质：C—N 键不能自由旋转，六个原子共平面；多为反式构型。',size=16.0,color=INK)
 text(s,.9,5.8,11.5,.7,'肽链方向：N端 → C端；n 个氨基酸形成 n−1 个肽键，释放 n−1 分子水。',size=15.0,color=MUTED)
 footer(s,num,74)

def fig_levels(num):
 s=prs.slides.add_slide(BLANK);bg(s,LIGHT);kicker(s,'第2章 · 结构层次',81);text(s,.62,.75,12,.55,'蛋白质的四级结构层次',size=27,color=NAVY,bold=True,font=SERIF)
 labels=[('一级结构','氨基酸序列\n共价键：肽键/二硫键'),('二级结构','α螺旋、β折叠\n主链氢键'),('三级结构','一条肽链的\n三维折叠'),('四级结构','多个亚基\n非共价缔合')]
 for i,(h,b) in enumerate(labels):
  x=.7+i*3.1;box(s,x,1.75,2.75,2.15,fill=WHITE,line=LINE);box(s,x,1.7,2.75,.08,fill=[BLUE,TEAL,GOLD,RED][i],radius=False);text(s,x+.18,1.88,2.4,.3,h,size=13,color=NAVY,bold=True);text(s,x+.18,2.35,2.4,1.2,b,size=13.5,color=INK,align=PP_ALIGN.CENTER)
  if i<3: text(s,x+2.78,2.65,.35,.3,'→',size=18,color=GOLD,bold=True)
 text(s,.9,5.35,11.6,.7,'结构层次：一级→二级→三级→四级；一级结构决定高级结构，高级结构决定功能。',size=12.5,color=INK)
 footer(s,num,81)

def fig_seq(num):
 s=prs.slides.add_slide(BLANK);bg(s,LIGHT);kicker(s,'第2章 · 测序策略',87);text(s,.62,.75,12,.55,'蛋白质测序：先拆后拼',size=27,color=NAVY,bold=True,font=SERIF)
 steps=[('1','末端与组成','Edman测N端；羧肽酶测C端；氨基酸组成分析'),('2','拆成肽段','Trypsin切Lys/Arg；CNBr切Met；得到可读短肽'),('3','测序与重叠','Edman/质谱测短肽；用重叠序列拼回全长；定位二硫键')]
 y=1.6
 for no,h,b in steps:
  box(s,.8,y,11.7,1.25,fill=WHITE,line=LINE);text(s,1.0,y+.2,.5,.4,no,size=18,color=GOLD,bold=True);text(s,1.55,y+.18,3.2,.3,h,size=12.5,color=NAVY,bold=True);text(s,4.4,y+.18,7.7,.75,b,size=13.5,color=INK)
  y+=1.55
 text(s,.9,6.1,11.6,.5,'关键：两套裂解方式产生重叠肽段，才能确定片段先后次序。',size=11.5,color=MUTED)
 footer(s,num,87)

def origfig(num,fig,title,page,items):
 s=prs.slides.add_slide(BLANK);bg(s,LIGHT);box(s,0,0,13.333,.1,fill=TEAL,radius=False)
 kicker(s,'原书配图 · 图%s（扫描原图）'%fig,page);text(s,.62,.70,12,.55,title,size=23,color=NAVY,bold=True,font=SERIF)
 text(s,.64,1.18,8,.25,'原书 图%s · 原图未重绘'%fig,size=9,color=GOLD,bold=True)
 path=os.path.join(ROOT,'assets/img/bc02orig/fig%s.png'%fig)
 im=Image.open(path);iw,ih=im.size;ar=iw/ih;maxw,maxh=8.0,5.15
 if ar>1.5:maxw,maxh=8.2,4.8
 scale=min(maxw/iw,maxh/ih);w,h=iw*scale,ih*scale;x=.62+(maxw-w)/2;y=1.5+(maxh-h)/2
 s.shapes.add_picture(path,Inches(x),Inches(y),Inches(w),Inches(h))
 box(s,9.0,1.4,3.9,5.15,fill=WHITE,line=LINE);text(s,9.2,1.55,3.5,.28,'看图顺序 + 知识点',size=10.8,color=TEAL,bold=True)
 bullets(s,9.2,1.95,3.5,4.35,['① '+x for x in items],size=9.6,gap=6)
 footer(s,num,page);s.notes_slide.notes_text_frame.text='原书图%s：%s'%(fig,title)


def table20(num):
 s=prs.slides.add_slide(BLANK);bg(s,LIGHT);box(s,0,0,13.333,.1,fill=TEAL,radius=False)
 kicker(s,'第2章 · 20种氨基酸分类总表',56);text(s,.62,.72,12,.55,'20种蛋白质氨基酸：五类一次记住',size=25,color=NAVY,bold=True,font=SERIF)
 rows=[('非极性/疏水','Gly 甘、Ala 丙、Val 缬、Leu 亮、Ile 异亮、Met 甲硫、Pro 脯','侧链疏水；Pro为环状亚氨基酸'),
 ('芳香族','Phe 苯丙、Tyr 酪、Trp 色','280 nm吸收；Trp/Tyr最强'),
 ('极性不带电','Ser 丝、Thr 苏、Cys 半胱、Asn 天冬酰胺、Gln 谷氨酰胺','—OH/—SH/—CONH₂，可形成氢键'),
 ('酸性','Asp 天冬、Glu 谷','侧链含—COOH，pI偏低'),
 ('碱性','Lys 赖、Arg 精、His 组','侧链含氨基/胍基/咪唑基，pI偏高')]
 y=1.5
 for i,(a,b,c) in enumerate(rows):
  box(s,.65,y,12.0,.86,fill=WHITE if i%2==0 else PALE,line=LINE)
  text(s,.85,y+.12,2.1,.3,a,size=11.5,color=[BLUE,TEAL,GOLD,RED,NAVY2][i],bold=True)
  text(s,3.0,y+.08,6.7,.68,b,size=9.8,color=INK)
  text(s,9.9,y+.12,2.6,.55,c,size=8.8,color=MUTED)
  y+=1.0
 text(s,.9,6.5,11.6,.3,'口诀：非极性—芳香—极性不带电—酸性—碱性；先分组，再记侧链特殊基团。',size=10.5,color=GOLD,bold=True)
 footer(s,num,56);s.notes_slide.notes_text_frame.text='20种氨基酸先按R基五类分组，再记每组的代表和特殊侧链。'

def mechfig(num,fig,title,page,rows):
 s=prs.slides.add_slide(BLANK);bg(s,LIGHT);box(s,0,0,13.333,.1,fill=RED,radius=False)
 kicker(s,'反应式四步卡 · 图%s'%fig,page);text(s,.62,.70,12,.55,title,size=22,color=NAVY,bold=True,font=SERIF)
 path=os.path.join(ROOT,'assets/img/bc02orig/fig%s.png'%fig);im=Image.open(path);iw,ih=im.size;ar=iw/ih
 maxw,maxh=7.4,5.1
 if ar>1.5:maxw,maxh=7.5,4.8
 sc=min(maxw/iw,maxh/ih);w,h=iw*sc,ih*sc;x=.55;y=1.45+(maxh-h)/2
 s.shapes.add_picture(path,Inches(x),Inches(y),Inches(w),Inches(h))
 box(s,8.25,1.38,4.45,5.15,fill=WHITE,line=LINE)
 labels=['① 反应物/条件','② 原子记账','③ 键数/键变化','④ 体内外对比']
 yy=1.55
 for i,(lab,txt) in enumerate(zip(labels,rows)):
  text(s,8.45,yy,3.95,.26,lab,size=9.8,color=[BLUE,TEAL,GOLD,RED][i],bold=True)
  text(s,8.45,yy+.27,3.95,.78,txt,size=8.9,color=INK)
  yy+=1.15
 footer(s,num,page);s.notes_slide.notes_text_frame.text='反应式四步卡：'+title+'。先找反应物条件，再原子记账、键数、体内外对比。'


def peptide_map(num):
 s=prs.slides.add_slide(BLANK);bg(s,LIGHT);box(s,0,0,13.333,.1,fill=GOLD,radius=False)
 kicker(s,'第7节 · 肽化学反应式总览',74);text(s,.62,.70,12,.55,'肽化学：五类反应，一张图串起来',size=24,color=NAVY,bold=True,font=SERIF)
 cards=[('肽键形成','氨基+羧基→肽键\n脱1分子水'),('肽键水解','肽键+H₂O→氨基+羧基\n酸/碱/酶'),('茚三酮','α-氨基氧化脱氨\n生成紫色色素'),('二硫键','2 Cys—SH→Cys—S—S—Cys\n氧化/还原'),('Edman/固相','N端标记测序\nC→N固相合成')]
 for i,(h,b) in enumerate(cards):
  x=.65+(i%5)*2.5;y=1.55;box(s,x,y,2.3,1.8,fill=WHITE,line=LINE);text(s,x+.12,y+.12,2.0,.3,h,size=10.5,color=NAVY,bold=True);text(s,x+.12,y+.55,2.0,1.0,b,size=8.8,color=INK,align=PP_ALIGN.CENTER)
 text(s,.9,3.75,11.6,.35,'每类反应都按四步读：①条件 ②原子记账 ③键数 ④体内外对比',size=11,color=TEAL,bold=True)
 text(s,.9,4.25,11.6,1.2,'例：肽键形成——羧基—OH与氨基—H脱去1分子水；形成1个C—N肽键；体内由核糖体催化并耗能，体外可加热/缩合剂。',size=10.5,color=INK)
 footer(s,num,74);s.notes_slide.notes_text_frame.text='肽化学总览：肽键形成/水解、茚三酮、二硫键、Edman、固相合成；统一按四步卡读反应式。'

def reaction_card(num,title,equation,rows,notes):
 s=prs.slides.add_slide(BLANK);bg(s,LIGHT);box(s,0,0,13.333,.1,fill=RED,radius=False)
 kicker(s,'反应式四步卡',74);text(s,.62,.70,12,.55,title,size=22,color=NAVY,bold=True,font=SERIF)
 box(s,.62,1.40,12.0,1.15,fill=WHITE,line=LINE);text(s,.85,1.55,11.5,.28,'反应式（抓底物/产物/条件）',size=9.8,color=BLUE,bold=True);text(s,.85,1.88,11.5,.55,equation,size=13.5,color=INK,align=PP_ALIGN.CENTER)
 for i,(h,b) in enumerate(zip(['① 条件','② 原子记账','③ 键数/键变化','④ 体内外对比'],rows)):
  x=.62+i*3.03;box(s,x,2.85,2.85,2.55,fill=WHITE,line=LINE);text(s,x+.16,3.0,2.5,.28,h,size=9.8,color=[BLUE,TEAL,GOLD,RED][i],bold=True);text(s,x+.16,3.35,2.5,1.85,b,size=9.3,color=INK)
 text(s,.9,5.75,11.8,.55,'学生任务：圈出被加入/脱去的原子，数键数，再解释体内为什么需要酶或能量。',size=10.5,color=GOLD,bold=True)
 footer(s,num,74);s.notes_slide.notes_text_frame.text=notes

# content data
# content data
# content data
# content data
SECTIONS=[
('一、氨基酸—蛋白质的单体亚基',54,['蛋白质的水解','α-氨基酸的一般结构'],[
('蛋白质的水解',54,['酸水解：6 mol/L HCl，水解完全，但色氨酸破坏、酰胺键水解。','碱水解：水解完全，但会引起消旋，色氨酸不破坏。','酶水解：条件温和、不消旋，但往往不完全，适合部分水解。'],[ '不同水解方法的“完全性/消旋/色氨酸”取舍。'],'把“完全水解”当成“都适合测组成”；色氨酸在酸水解中被破坏。','水解方法决定后续分析：测氨基酸组成常用酸水解，但色氨酸需另法。'),
('α-氨基酸的一般结构',54,['α-碳连接 —NH₃⁺、—COO⁻、—H 和侧链 R。','除甘氨酸外，α-碳连接四个不同基团，是手性碳。','蛋白质中常见氨基酸多为 L-构型。'],[ 'α-氨基酸的通用骨架；R 基决定性质。'],'把“L-构型”说成“左旋”；L/D 是构型，不是旋光方向。','α-碳是氨基酸分类、手性、酸碱性质的中心。')],
[('生命中的氨基酸为什么几乎都是L型？','酶促反应的立体专一性；L型与酶活性中心匹配。'),('酸水解最容易被破坏的氨基酸是？','色氨酸。')],'L/D不是左右旋；酸水解破坏色氨酸。'),
('二、氨基酸的分类',55,['常见的蛋白质氨基酸','不常见的蛋白质氨基酸','非蛋白质氨基酸'],[
('常见的蛋白质氨基酸',56,['常见蛋白质氨基酸约20种。','按R基极性分：非极性、极性不带电、酸性、碱性；芳香族单列。','R基决定溶解度、电荷、氢键和疏水作用。'],[ '20种氨基酸按R基分类；名字—结构—性质对应。'],'把“必需氨基酸”与“常见蛋白质氨基酸”混为一谈。','分类服务于性质：疏水/极性/带电决定蛋白质折叠与分离。'),
('不常见的蛋白质氨基酸',58,['由常见氨基酸修饰而来，如羟脯氨酸、羟赖氨酸。','多出现在胶原蛋白等特定蛋白质中。','修饰改变结构稳定性或功能。'],[ '不常见≠非蛋白质；它们仍在蛋白质中。'],'把修饰氨基酸当成“非蛋白质氨基酸”。','胶原蛋白的羟脯氨酸与三股螺旋稳定有关。'),
('非蛋白质氨基酸',60,['不参与蛋白质合成的氨基酸，如鸟氨酸、瓜氨酸。','可作为代谢中间物或天然产物。','D-氨基酸也可见于细菌细胞壁、抗生素。'],[ '非蛋白质氨基酸：存在但不作为蛋白质构件。'],'把鸟氨酸/瓜氨酸当成蛋白质20种氨基酸。','尿素循环中间物鸟氨酸、瓜氨酸属于非蛋白质氨基酸。')],
[('20种氨基酸按R基分为哪几类？','非极性、极性不带电、酸性、碱性；芳香族单列。'),('鸟氨酸属于哪类？','非蛋白质氨基酸，尿素循环中间物。')],'“不常见蛋白质氨基酸”和“非蛋白质氨基酸”不是一回事。'),
('三、氨基酸的酸碱性质',60,['氨基酸的解离','氨基酸的等电点'],[
('氨基酸的解离',60,['氨基酸是两性电解质，含—COOH和—NH₃⁺。','低pH质子化带正电；高pH去质子化带负电。','pK₁≈羧基，pK₂≈氨基，pKR≈侧链。'],[ 'pH决定质子化状态；pK是半解离点。'],'把pK₁、pK₂固定对应所有氨基酸；侧链pKR也要考虑。','滴定曲线平台对应pK，拐点对应pI附近。'),
('氨基酸的等电点',62,['pI是净电荷为零时的pH。','中性氨基酸pI≈(pK₁+pK₂)/2。','酸性/碱性氨基酸要结合侧链pKR计算；pI时溶解度最低。'],[ 'pI定义、计算、溶解度和电泳行为。'],'把pI当成“中性pH=7”；pI由分子本身决定。','离子交换/电泳中，pH>pI带负电，pH<pI带正电。')],
[('中性氨基酸的pI如何估算？','pI≈(pK₁+pK₂)/2；酸性/碱性需结合侧链pKR。'),('pH高于pI时氨基酸带什么电？','带负电。')],'pH与pI比较决定电荷方向：pH>pI为负，pH<pI为正。'),
('四、氨基酸的化学反应',64,['α-羧基参加的反应','α-氨基参加的反应','α-羧基和α-氨基共同参加的反应','侧链官能团参加的反应'],[
('α-羧基参加的反应',64,['成酯、成酰胺、脱羧。','可用于保护羧基或制备衍生物。','脱羧生成相应胺类。'],[ '羧基反应：酯化/酰胺化/脱羧。'],'把脱羧和脱氨混淆。','脱羧丢CO₂，脱氨丢NH₃。'),
('α-氨基参加的反应',64,['酰化、烷基化、与醛形成Schiff碱。','与茚三酮反应可显色，用于检测。','氨基反应常用于标记和测序。'],[ '氨基亲核；酰化/ Schiff碱/茚三酮。'],'茚三酮不是只与氨基反应，也受其他基团影响。','Edman试剂PITC与N端氨基反应。'),
('α-羧基和α-氨基共同参加的反应',64,['与茚三酮反应生成蓝紫色产物（脯氨酸黄色）。','加热可成二酮哌嗪。','肽键形成是羧基与氨基的缩合。'],[ '共同反应：茚三酮、成肽。'],'茚三酮对所有氨基酸显色相同；脯氨酸例外。','肽键形成需活化羧基，生物体内由核糖体催化。'),
('侧链官能团参加的反应',66,['—SH可氧化成二硫键，参与稳定结构。','—OH、咪唑基、胍基可参与氢键/离子键。','侧链反应用于化学修饰和活性位点研究。'],[ '侧链决定特殊反应与修饰。'],'把二硫键当成肽键；二硫键是Cys侧链间共价键。','二硫键可在还原剂下断裂，用于测序和复性。')],
[('茚三酮与脯氨酸显什么颜色？','黄色（其他多数氨基酸蓝紫色）。'),('二硫键由哪两个氨基酸形成？','两个半胱氨酸（Cys）的—SH氧化。')],'茚三酮脯氨酸黄色；二硫键不是肽键。'),
('五、氨基酸的旋光性和光谱性质',67,['氨基酸的旋光性和立体化学','氨基酸的光谱性质'],[
('氨基酸的旋光性和立体化学',67,['手性碳使氨基酸具有旋光性；L/D是构型命名。','多数天然氨基酸为L型；D型在细菌细胞壁等出现。','旋光方向（+/-）与L/D不是同一概念。'],[ '手性、L/D、旋光方向三者区分。'],'把L等同于左旋（−）；L/D看构型，旋光看实验。','对映体旋光方向相反，生物活性常不同。'),
('氨基酸的光谱性质',68,['芳香族氨基酸（Phe/Tyr/Trp）在280 nm有吸收。','Trp、Tyr贡献最大；可用于蛋白质定量。','紫外吸收受pH和微环境影响。'],[ '280 nm吸收来自芳香族侧链。'],'核酸也在260 nm吸收；蛋白质定量常用280 nm或Bradford。','Phe吸收较弱，Trp/Tyr更强。')],
[('L/D与旋光方向是什么关系？','无固定对应；L是构型，+/-是实验旋光方向。'),('蛋白质280 nm吸收主要来自？','Trp、Tyr（Phe较弱）。')],'L≠左旋；280 nm不是核酸的260 nm。'),
('六、氨基酸混合物的分析和分离',69,['分配层析','离子交换层析'],[
('分配层析',69,['纸层析/薄层层析：按在两相中的分配系数不同分离。','Rf值与分子极性、展开剂有关。','可用于氨基酸定性鉴定。'],[ '分配层析=Rf/极性/两相分配。'],'Rf值不是固定常数，受展开剂、温度、滤纸影响。','Rf越大，移动越远。'),
('离子交换层析',72,['阳离子交换树脂在低pH结合带正电氨基酸。','不同氨基酸pI不同，结合强弱不同。','用pH/盐梯度洗脱，按电荷差异分离。'],[ '离子交换=pI/电荷/梯度洗脱。'],'pH>pI时氨基酸带负电，阳离子交换结合减弱。','离子交换常用于氨基酸分析仪。')],
[('离子交换层析主要按什么分离？','电荷/pI差异；用pH或盐梯度洗脱。'),('pH>pI时氨基酸带什么电？','负电。')],'阳离子交换结合正电；pH升高使结合减弱。'),
('七、肽',74,['肽和肽键的结构','肽的物理和化学性质','天然存在的活性肽'],[
('肽和肽键的结构',74,['肽键是—CO—NH—，由α-羧基与α-氨基缩合。','肽键有部分双键性质，C—N不能自由旋转，六个原子共平面。','肽链有N端和C端；n个氨基酸形成n−1个肽键。'],[ '肽键平面性、方向性、n−1关系。'],'把肽键当成普通单键；其部分双键性限制旋转。','肽键多为反式；脯氨酸可顺式。'),
('肽的物理和化学性质',76,['肽的酸碱性质由N端、C端和侧链决定。','肽可水解，肽键断裂释放小肽/氨基酸。','肽的旋光性、溶解度与序列有关。'],[ '肽=多个氨基酸，仍有两性。'],'肽没有游离α-氨基/羧基？端基和侧链仍有。','肽键水解需要酸/碱/酶。'),
('天然存在的活性肽',77,['谷胱甘肽（GSH）：抗氧化，含γ-谷氨酰半胱氨酰甘氨酸。','胰岛素：多肽激素，A/B链二硫键连接。','脑啡肽、催产素等短肽信号分子。'],[ '活性肽例子：GSH/胰岛素/脑啡肽。'],'谷胱甘肽的肽键不是普通α-肽键（γ-谷氨酰）。','GSH的—SH参与氧化还原。')],
[('肽键为什么不能自由旋转？','C—N具有部分双键性质，六原子共平面。'),('n个氨基酸形成多少肽键？','n−1个。')],'肽键有部分双键性；n个氨基酸→n−1肽键。'),
('八、蛋白质的组成、分类、分子大小和结构层次',78,['蛋白质的化学组成和分类','蛋白质分子的形状和大小','蛋白质分子结构的组织层次'],[
('蛋白质的化学组成和分类',78,['元素组成：C、H、O、N，常含S；N≈16%。','单纯蛋白质/结合蛋白质；球状/纤维状。','结合蛋白质含辅基：血红素、糖、脂、金属等。'],[ '元素、分类、辅基。'],'蛋白质含氮量16%是凯氏定氮的基础。','分类维度不同，不要混用。'),
('蛋白质分子的形状和大小',79,['球状蛋白近似球形，纤维状蛋白呈长链。','相对分子质量从几千到数百万。','可用凝胶过滤、SDS-PAGE、超速离心、质谱测分子量。'],[ '形状/大小/测定方法。'],'SDS-PAGE测的是亚基分子量，不是天然完整蛋白。','凝胶过滤按流体力学半径分离。'),
('蛋白质分子结构的组织层次',81,['一级：氨基酸序列；二级：α螺旋/β折叠；三级：三维折叠；四级：亚基组装。','一级结构由肽键/二硫键维持；高级结构主要靠非共价作用。','结构层次决定功能。'],[ '四级结构层次与作用力。'],'四级结构不是所有蛋白都有；单链蛋白只有三级。','二硫键属于一级结构交联，但也稳定三级结构。')],
[('蛋白质中氮含量平均约多少？','16%。'),('SDS-PAGE测定的是什么分子量？','亚基分子量。')],'SDS-PAGE测亚基；四级结构不是所有蛋白都有。'),
('九、蛋白质的一级结构',82,['序列决定功能','一级结构举例','蛋白质序列数据库','蛋白质化学测序的策略'],[
('蛋白质的氨基酸序列决定蛋白质的功能',82,['一级结构=氨基酸排列顺序。','序列决定折叠，折叠决定功能。','突变可改变功能：镰状细胞贫血HbS。'],[ '序列→结构→功能；突变例子。'],'一级结构不仅决定结构，也决定定位、修饰。','HbS是单个氨基酸突变。'),
('蛋白质一级结构的举例',82,['胰岛素：A链21、B链30，二硫键连接。','核糖核酸酶：124个氨基酸，4对二硫键。','血红蛋白：α₂β₂四聚体。'],[ '经典一级结构例子。'],'胰岛素A/B链不是一条链；二硫键连接。','核糖核酸酶测序是经典工作。'),
('蛋白质序列数据库',83,['UniProt、GenBank、PDB等。','序列比对可找同源蛋白和保守位点。','数据库是序列—结构—功能关联的入口。'],[ '数据库/比对/同源。'],'数据库序列不等于实验验证的功能。','PDB是结构数据库，UniProt是序列/功能数据库。'),
('蛋白质化学测序的策略',84,['先测末端和组成，再裂解成肽段。','测短肽序列，用重叠肽段重建顺序。','定位二硫键；现代多用质谱。'],[ '测序策略：末端—裂解—重叠—二硫键。'],'不能只靠一种酶裂解就拼出全序列。','两套裂解产生重叠是重建关键。')],
[('一级结构决定什么？','决定折叠和功能；突变可致病。'),('测序中“重叠肽段”的作用？','确定片段先后次序，重建全长。')],'序列→结构→功能；测序需重叠片段。'),
('十、蛋白质测序的一些常用方法',84,['末端残基分析','二硫键的断裂','氨基酸组成的分析','多肽链的部分裂解','肽段序列测定','全序列重建','二硫键定位'],[
('末端残基分析',84,['N端：Edman降解/PITC；DNS-Cl等。','C端：羧肽酶A/B/Y。','末端分析只能确定两端，不能给全序列。'],[ 'N端Edman、C端羧肽酶。'],'Edman从N端逐个切，不能测C端。','C端测定常因羧肽酶特异性受限。'),
('二硫键的断裂',85,['氧化法：过甲酸氧化Cys为磺基丙氨酸。','还原法：DTT/β-巯基乙醇，随后烷基化保护。','断裂二硫键使肽链展开，便于测序。'],[ '二硫键断裂：氧化/还原+烷基化。'],'还原后不烷基化，—SH会重新氧化。','过甲酸氧化不可逆。'),
('氨基酸组成的分析',86,['完全水解后用氨基酸分析仪/色谱定量。','酸水解破坏Trp，需另法。','组成分析给“有哪些、多少”，不给顺序。'],[ '组成分析≠序列分析。'],'组成相同不等于序列相同。','Trp需碱水解或特殊方法。'),
('多肽链的部分裂解',86,['Trypsin：切Lys/Arg的C端。','Chymotrypsin：切芳香族/Leu等疏水残基。','CNBr：切Met的C端。'],[ '酶切特异性决定片段。'],'Trypsin切Lys/Arg的C端，不是N端。','CNBr切Met，不是所有蛋白都有Met。'),
('肽段氨基酸序列的测定',87,['Edman降解逐轮测N端。','质谱（MS/MS）通过碎片离子测序。','短肽更易得到清晰序列。'],[ 'Edman/质谱测短肽。'],'Edman对N端封闭肽无效。','质谱需正确解释碎片。'),
('肽段在原多肽链中的次序确定',91,['用不同裂解方式得到重叠肽段。','重叠序列确定相邻片段顺序。','重建全长序列并核对。'],[ '重叠法重建顺序。'],'只有一套裂解通常无法确定顺序。','重叠区越长越可靠。'),
('二硫键位置的确定',92,['先在不还原条件下酶切。','比较还原/非还原电泳或质谱。','对角线电泳可定位二硫键肽段。'],[ '二硫键定位：非还原酶切+比较。'],'还原后二硫键信息消失，需先保留。','对角线电泳是经典方法。')],
[('Trypsin切割位点？','Lys/Arg的C端。'),('Edman降解测哪一端？','N端。')],'Trypsin切C端；Edman测N端；组成≠序列。'),
('十一、氨基酸序列与生物进化',92,['序列同源性、同源蛋白质和蛋白质家族','同源蛋白质氨基酸序列的物种差异','同源蛋白质具有共同的进化起源'],[
('序列的同源性、同源蛋白质和蛋白质家族',92,['同源蛋白：不同物种中序列/结构相似的蛋白。','序列同一性、保守替换可反映亲缘关系。','蛋白质家族由共同祖先基因演化。'],[ '同源/家族/保守位点。'],'同源不是“功能相同”的同义词；看共同起源。','相似性高不一定同源，需统计检验。'),
('同源蛋白质氨基酸序列的物种差异',93,['物种越近，序列差异越小。','差异积累可作为分子钟。','保守残基多与结构/功能关键位点有关。'],[ '物种差异/分子钟/保守残基。'],'分子钟速率不恒定；受选择影响。','功能关键位点更保守。'),
('同源蛋白质具有共同的进化起源',95,['序列比对和系统树支持共同祖先。',' invariant residues 提示关键功能。','基因复制/突变产生新功能。'],[ '共同起源：序列/结构/系统树证据。'],'序列相似也可能趋同演化，不一定是同源。','系统树分支长度表示差异，不表示“高级”。')],
[('同源蛋白的“同源”指什么？','有共同进化起源，不是单纯功能相同。'),('保守残基通常说明什么？','对结构/功能重要，受选择压力。')],'同源=共同起源；相似≠同源；保守位点重要。'),
('十二、肽和蛋白质的化学合成：固相肽合成',96,['固相肽合成'],[
('固相肽合成',96,['将C端氨基酸固定在树脂上。','循环：脱保护→活化→偶联→洗涤。','最后裂解树脂并脱侧链保护基，得到肽。'],[ 'Merrifield固相合成循环。'],'合成方向是C端→N端，与生物合成N端→C端相反。','每步需保护侧链，避免副反应。')],
[('固相肽合成的方向？','C端→N端（在树脂上从C端延长）。'),('每轮循环的关键步骤？','脱保护、活化、偶联、洗涤。')],'固相合成C→N；生物合成N→C。')
]

num=1
s=prs.slides.add_slide(BLANK);bg(s,NAVY);box(s,0,0,.14,7.5,fill=GOLD,radius=False)
text(s,.95,.9,9,.4,'BIOCHEMISTRY · CHAPTER 02',size=11,color=GOLD,bold=True)
text(s,.95,1.6,11,.7,'第2章',size=22,color=RGBColor(0xB8,0xCE,0xDD),bold=True)
text(s,.95,2.1,11.5,1.2,'氨基酸、肽和蛋白质',size=48,color=WHITE,bold=True,font=SERIF)
text(s,.98,3.5,10.5,.45,'朱圣庚《生物化学》第4版 · AI 讲义版',size=16,color=RGBColor(0xD5,0xE3,0xEE))
text(s,.98,4.25,7,.4,'AI 讲义 · 12节 · 36个子节',size=10.5,color=GOLD,bold=True)
text(s,.98,6.55,10.5,.3,'先框架，再结构/性质，最后测序与合成',size=9.5,color=RGBColor(0x9E,0xB6,0xC8));num+=1

# map
s=prs.slides.add_slide(BLANK);bg(s,LIGHT);box(s,0,0,13.333,.1,fill=BLUE,radius=False)
kicker(s,'本章学习地图',0);text(s,.62,.78,12,.55,'第2章 · 12节 36个子节',size=27,color=NAVY,bold=True,font=SERIF)
for i,sec in enumerate(SECTIONS):
 x=.62+(i%2)*6.25;y=1.55+(i//2)*.86
 box(s,x,y,5.95,.74,fill=WHITE,line=LINE);box(s,x,y,.07,.74,fill=GOLD if i<4 else TEAL,radius=False)
 text(s,x+.2,y+.10,.45,.25,'%02d'%(i+1),size=11,color=GOLD if i<4 else TEAL,bold=True)
 text(s,x+.72,y+.08,4.9,.3,sec[0].split('、',1)[1],size=12.5,color=NAVY,bold=True)
 text(s,x+.72,y+.38,4.9,.24,'原书 p%d · %d个子节'%(sec[1],len(sec[2])),size=9.8,color=MUTED)
footer(s,num,0);num+=1

# Build sections
for si,(title,page,subs,items,wqs,wrap_pit) in enumerate(SECTIONS,1):
 section(num,si,title.split('、',1)[1],page,[it[0] for it in items][:3]);num+=1
 if si==7: peptide_map(num);num+=1
 for j,(st,sp,bul,kp,pf,note) in enumerate(items):
  content(num,'%s · %d/%d'%(title,j+1,len(items)),st,sp,bul,kp,pf,note);num+=1

  if si==2 and j==0: origfig(num,'2-2','图2-2 20种常见的蛋白质氨基酸',56,['20种氨基酸按R基五类分组。','同一类侧链性质相似。','记名字、缩写和特殊侧链。']);num+=1
  if si==2 and j==1: origfig(num,'2-3','图2-3 一些不常见的蛋白质氨基酸',59,['由常见氨基酸修饰而来。','多出现在胶原等特定蛋白。','修饰改变结构/功能。']);num+=1
  if si==2 and j==2: origfig(num,'2-4','图2-4 某些非蛋白质氨基酸',60,['不参与蛋白质合成。','如鸟氨酸、瓜氨酸。','可作代谢中间物/天然产物。']);num+=1
  if si==3 and j==1: origfig(num,'2-6','图2-6 谷氨酸和赖氨酸的滴定曲线',63,['酸性/碱性氨基酸有第三个pK。','pI要结合侧链pKR计算。','曲线平台对应缓冲区。']);num+=1
  if si==4 and j==1: origfig(num,'2-8','图2-8 氨基酸与茚三酮的反应',66,['茚三酮用于氨基酸检测。','多数氨基酸显蓝紫色。','脯氨酸显黄色。']);num+=1
  if si==4 and j==1: mechfig(num,'2-8','图2-8 茚三酮反应：原子去哪了？',66,['α-氨基 + 水合茚三酮，加热。','脱羧、氧化脱氨，生成醛、CO₂、NH₃和还原茚三酮。','C—N/C—C断裂；生成紫色色素。','体外检测反应，不是体内正常代谢。']);num+=1
  if si==4 and j==2: fig_reaction(num);num+=1
  if si==5 and j==1: origfig(num,'2-11','图2-11 芳香族氨基酸的紫外吸收',69,['Phe/Tyr/Trp在280 nm吸收。','Trp/Tyr贡献最大。','用于蛋白质定量。']);num+=1
  if si==9 and j==1: origfig(num,'2-28','图2-28 牛胰岛素的一级结构',83,['A链21、B链30个残基。','链间/链内二硫键。','一级结构决定高级结构。']);num+=1
  if si==10 and j==0: mechfig(num,'2-33','图2-33 Edman降解：N端怎么测？',87,['PITC + N端氨基，弱碱，无水。','PITC标记N端，酸处理切下N端残基。','每轮断1个肽键，只测1个N端。','体外化学测序；体内蛋白质降解不是Edman。']);num+=1
  if si==7 and j==1: reaction_card(num,'肽键水解：水加到哪？','—CO—NH— + H₂O  →  —COOH + H₂N—',['酸/碱/蛋白酶催化；加热或酶。','H₂O的—H加到N，—OH加到羰基C。','断1个C—N肽键；生成1个—COOH和1个—NH₂。','体内由蛋白酶/肽酶催化；强酸强碱是体外水解。'],'肽键水解：加水断1个肽键，体内靠蛋白酶。');num+=1
  if si==10 and j==1: reaction_card(num,'二硫键：氧化成键/还原断键','2 Cys—SH  ⇌  Cys—S—S—Cys + 2H⁺ + 2e⁻',['氧化剂/还原剂；体内由巯基交换/还原系统。','两个—SH脱2H成S—S；还原加2H断键。','形成/断裂1个S—S键（二硫键）。','体外用DTT/过甲酸；体内由谷氧还蛋白等调控。'],'二硫键：氧化脱2H成S—S，还原加2H断键。');num+=1
  # original textbook figures inline
  if si==2 and j==0: table20(num);num+=1
  if si==1 and j==1: origfig(num,'2-1','图2-1 α-氨基酸的一般结构',54,['α-碳连接—NH₃⁺、—COO⁻、—H和R。','除Gly外，α-碳是手性碳。','20种氨基酸的R基不同。']);num+=1
  if si==3 and j==0: origfig(num,'2-5','图2-5 甘氨酸的滴定曲线',62,['羧基pK₁≈2.34，氨基pK₂≈9.60。','pI≈(pK₁+pK₂)/2≈5.97。','曲线平坦区是缓冲区。']);num+=1
  if si==5 and j==0: origfig(num,'2-9','图2-9 丙氨酸与甘油醛的构型关系',68,['L-丙氨酸的构型参照L-甘油醛。','D/L是构型，不是旋光方向。','蛋白质中常见L-氨基酸。']);num+=1
  if si==6 and j==1: origfig(num,'2-17','图2-17 氨基酸分析仪的洗脱顺序',73,['离子交换按电荷/pI分离。','不同氨基酸洗脱时间不同。','用pH/盐梯度洗脱。']);num+=1
  if si==7 and j==0:
    origfig(num,'2-19','图2-19 五肽的结构',74,['肽键—CO—NH—连接氨基酸。','N端→C端有方向性。','侧链决定肽的性质。']);num+=1
    origfig(num,'2-21','图2-21 肽键的结构特征',75,['肽键C—N有约40%双键性。','六个原子共平面，多为反式。','C=O与N—H形成氢键。']);num+=1
  if si==8 and j==2: origfig(num,'2-27','图2-27 蛋白质结构的组织层次',81,['一级：氨基酸序列。','二级：α螺旋/β折叠。','三级：三维折叠；四级：亚基。']);num+=1
  if si==12 and j==0: origfig(num,'2-42','图2-42 固相肽合成原理',96,['C端固定在树脂上。','循环：脱保护→偶联→洗涤。','最后裂解并脱侧链保护基。']);num+=1
  if si==12 and j==0: mechfig(num,'2-42','图2-42 固相肽合成：每轮成几个键？',96,['树脂固定C端；脱保护→活化→偶联→洗涤。','羧基与氨基缩合，脱1分子水。','每轮形成1个肽键。','固相C→N；生物合成N→C，需酶/能量。']);num+=1
  # figures after selected
  if si==1 and j==1: fig_amino(num);num+=1
  if si==3 and j==1: fig_charge(num);num+=1
  if si==7 and j==0: fig_peptide(num);num+=1
  if si==8 and j==2: fig_levels(num);num+=1
  if si==10 and j==5: fig_seq(num);num+=1
 # wrap
 wrap(num,si,title.split('、',1)[1],
  [ '第%d节：%s 的主线是“%s”。'%(si,title.split('、',1)[1],items[0][0]),
    '关键概念：%s。'%(items[0][3]),
    '考试落点：%s'%items[0][4] ],
  wqs,
  wrap_pit,
  '第%d节收束：先抓主线，再记核心概念与易错点。'%si);num+=1

# summary
s=prs.slides.add_slide(BLANK);bg(s,LIGHT);box(s,0,0,13.333,.1,fill=GOLD,radius=False)
kicker(s,'本章总结',96);text(s,.62,.78,12,.55,'第2章 · 三条主线',size=27,color=NAVY,bold=True,font=SERIF)
cols=[('结构',['α-氨基酸：α-C、R基、L型','20种氨基酸按R基分类','肽键平面性/方向性']),('性质',['两性电解质，pI决定电荷','280 nm芳香族吸收','离子交换/层析分离']),('序列与合成',['一级结构决定功能','测序：末端—裂解—重叠','固相合成C→N'])]
for i,(h,items) in enumerate(cols):
 x=.62+i*4.18;box(s,x,1.65,3.95,3.9,fill=WHITE,line=LINE);box(s,x,1.65,3.95,.08,fill=[BLUE,TEAL,GOLD][i],radius=False);text(s,x+.22,1.85,3.5,.3,h,size=13.5,color=NAVY,bold=True)
 for k,t in enumerate(items): text(s,x+.25,2.35+k*.72,3.45,.4,'▸ '+t,size=15.0,color=INK)
footer(s,num,96);num+=1
# quiz
s=prs.slides.add_slide(BLANK);bg(s,LIGHT);box(s,0,0,13.333,.1,fill=TEAL,radius=False)
kicker(s,'课堂自测',96);text(s,.62,.78,12,.55,'5 道题检查本章掌握度',size=27,color=NAVY,bold=True,font=SERIF)
qs=[('1','下列属于碱性氨基酸的是？','A. Asp  B. Glu  C. Lys  D. Ser'),('2','肽键的C—N键为什么不能自由旋转？','A. 离子键  B. 部分双键性质  C. 氢键  D. 二硫键'),('3','中性氨基酸pI近似等于？','A. pK₁  B. pK₂  C. (pK₁+pK₂)/2  D. 7.0'),('4','Edman降解测哪一端？','A. N端  B. C端  C. 侧链  D. 二硫键'),('5','固相肽合成方向是？','A. N→C  B. C→N  C. 任意  D. 只加Cys')]
for i,(n,q,o) in enumerate(qs):
 y=1.55+i*1.02;box(s,.62,y,12,.86,fill=WHITE,line=LINE);text(s,.78,y+.12,.35,.3,n,size=14,color=GOLD,bold=True);text(s,1.12,y+.1,10.8,.28,q,size=15.0,color=INK,bold=True);text(s,1.12,y+.45,10.8,.26,o,size=13.0,color=MUTED)
footer(s,num,96)

# renumber footers
for i,sl in enumerate(prs.slides,1):
 for sh in sl.shapes:
  if not sh.has_text_frame: continue
  try:
   if sh.left<Inches(11.3) or sh.top<Inches(6.65): continue
  except Exception: continue
  tf=sh.text_frame; t=tf.text.strip()
  if t.isdigit() and len(t)<=2:
   if tf.paragraphs and tf.paragraphs[0].runs: tf.paragraphs[0].runs[0].text='%02d'%i
prs.save(PPTX);print('SAVED',PPTX,len(prs.slides))
