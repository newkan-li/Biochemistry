import fitz,os,json,re
BASE='/home/newkan/Documents/biology/biochemistry/Biochemistry';PDF='/tmp/opencode/pptpdf/8.pdf';d=fitz.open(PDF)
fp=os.path.join(BASE,'data','bc06.js');s=open(fp,encoding='utf-8').read();ch=json.loads(re.search(r'=(\{.*\});?\s*$',s,re.S).group(1))
# preserve checks from current selected pages
old=ch['modules'][0]['slides']; checks={x.get('page'):x.get('check') for x in old if x.get('check')}
slides=[]
for page in range(1,d.page_count+1):
 txt=d[page-1].get_text('text').strip();norm=re.sub(r'\s+',' ',txt).strip()
 lines=[x.strip() for x in txt.splitlines() if x.strip()]
 title=lines[0][:60] if lines else f'第6章 · PPT第{page}页'
 if len(title)<2:title=f'第6章 · PPT第{page}页'
 points=[x.strip() for x in re.split(r'[。；;]',norm) if len(x.strip())>8][:3]
 fn=f'assets/img/bc06/ppt{page}.jpg'
 pix=d[page-1].get_pixmap(dpi=100,colorspace=fitz.csRGB);pix.save(os.path.join(BASE,fn),jpg_quality=80)
 slides.append({'i':page-1,'sec':'第6章 酶通论·老师PPT逐页','title':title,'notes':norm[:1600] or '（此页为过渡页/无文字）','points':points,'fig':f'中文PPT：王镜岩《生物化学》课件 · 第{page}页','anim':None,'check':checks.get(page),'cards':None,'img':fn,'page':page})
ch['modules'][0]['slides']=slides
ch['sub2']='第6章 · 87页老师PPT全量 + 重点页考点题 + 课后题/真题'
open(fp,'w',encoding='utf-8').write('window.CHAPTERS=window.CHAPTERS||{};window.CHAPTERS["bc06"]='+json.dumps(ch,ensure_ascii=False)+';\n')
print('slides',len(slides),'checks',sum(1 for x in slides if x.get('check')))
