import fitz,os,json,re,math,glob
BASE='/home/newkan/Documents/biology/biochemistry/Biochemistry';PDF='/tmp/opencode/pptpdf'
NOTES='/home/newkan/Documents/biology/生物化学 (第4版) 上下册 by 朱圣庚  徐长法+笔记和课后习题（含考研真题）详解'
ocr='/home/newkan/Documents/biology/biochemistry/_incoming/notes_ocr'
# chapters using Chinese Wang PPT
src={2:'3.pdf',3:'6.pdf',4:'6.pdf',5:'7.pdf',6:'8.pdf',7:'9.pdf',8:'10.pdf',9:'1.pdf',10:'2.pdf',11:'13.pdf',12:'14.pdf',13:'11.pdf',14:'17.pdf',15:'19.pdf',16:'20.pdf',17:'22.pdf',25:'30.pdf',26:'32.pdf',32:'36.pdf',33:'38.pdf',34:'39.pdf',35:'40.pdf'}
up=fitz.open([x for x in glob.glob(NOTES+'/*上册*.pdf') if '笔记' not in os.path.basename(x)][0]);down=fitz.open([x for x in glob.glob(NOTES+'/*下册*.pdf') if '笔记' not in os.path.basename(x)][0])
# notes PDFs for OCR? Use book pages for images and OCR files for text; mapping arrays
up_starts=[3,8,24,37,50,62,73,93,107,122,135,148,162,172,185];down_starts=[3,13,24,38,52,67,77,83,94,101,126,139,148,166,181,194,215,227,247,269,287,313,318]
def grams(s):
 s=re.sub(r'[^\u4e00-\u9fffA-Za-z0-9]+',' ',s).lower();out=set()
 for tok in s.split():
  for i in range(len(tok)-1):out.add(tok[i:i+2])
 return out
for num in range(1,37):
 cid=f'bc{num:02d}'
 if cid=='bc06':continue
 p=f'data/{cid}.js';s=open(p,encoding='utf-8').read();ch=json.loads(re.search(r'=(\{.*\});?\s*$',s,re.S).group(1));qs=ch.get('short') or []
 slides=[]
 if num in src:
  d=fitz.open(os.path.join(PDF,src[num])); total=d.page_count
  for pg in range(1,total+1):
   txt=d[pg-1].get_text('text').strip();norm=re.sub(r'\s+',' ',txt).strip();lines=[x.strip() for x in txt.splitlines() if x.strip()]
   title=lines[0][:70] if lines else f'PPT第{pg}页'
   fn=f'assets/img/{cid}/full{pg}.jpg';pix=d[pg-1].get_pixmap(dpi=100,colorspace=fitz.csRGB);pix.save(os.path.join(BASE,fn),jpg_quality=78)
   slides.append({'i':pg-1,'sec':'老师PPT逐页','title':title,'notes':norm[:1800] or '（此页为过渡页/无文字）','points':[x.strip() for x in re.split(r'[。；;]',norm) if len(x.strip())>8][:3],'fig':'','page':None,'anim':None,'check':None,'cards':None,'img':fn})
 else:
  if num<=14:doc=up;start=up_starts[num-1];end=up_starts[num] if num<14 else 185;pref='up'
  else:idx=num-15;doc=down;start=down_starts[idx];end=down_starts[idx+1];pref='down'
  for pg in range(start,end):
   fn=f'assets/img/{cid}/full{pg}.jpg';pix=doc[pg-1].get_pixmap(dpi=100,colorspace=fitz.csRGB);pix.save(os.path.join(BASE,fn),jpg_quality=78)
   f=os.path.join(ocr,f'{pref}_{pg:03d}.txt');txt=open(f,encoding='utf-8',errors='ignore').read().strip() if os.path.exists(f) else ''
   norm=re.sub(r'\s+',' ',txt).strip();lines=[x.strip() for x in txt.splitlines() if x.strip()]
   title=lines[0][:70] if lines else f'笔记原页 p{pg}'
   slides.append({'i':pg-start,'sec':'中文笔记原页','title':title,'notes':norm[:1800] or '（此页无文字/过渡页）','points':[x.strip() for x in re.split(r'[。；;]',norm) if len(x.strip())>8][:3],'fig':'','page':None,'anim':None,'check':None,'cards':None,'img':fn})
 # map existing after-class questions to best slides (max one each)
 unused=list(range(len(qs)));used=set()
 for si,sl in enumerate(slides):
  if si in used:continue
  sg=grams((sl.get('title') or '')+' '+(sl.get('notes') or '')[:1200]);best=None;bs=0
  for qi in unused:
   q=qs[qi];qg=grams((q.get('q') or '')+' '+(q.get('a') or '')[:250]);
   if not qg:continue
   sc=len(qg&sg)/(math.sqrt(len(qg))*math.sqrt(max(1,len(sg))))
   if sc>bs:bs=sc;best=qi
  if best is not None and bs>=0.025:
   q=qs[best];sl['check']={'kind':'short','id':q.get('id'),'q':q.get('q'),'a':q.get('a'),'e':q.get('e',''),'src':'配套笔记课后习题'};unused.remove(best);used.add(si)
 ch['modules'][0]['slides']=slides;ch['sub2']=f'第{num}章 · 全量PPT/笔记页 + 本页考点 + 课后题'
 open(p,'w',encoding='utf-8').write('window.CHAPTERS=window.CHAPTERS||{};window.CHAPTERS["'+cid+'"]='+json.dumps(ch,ensure_ascii=False)+';\n')
 print(cid,'slides',len(slides),'checks',sum(1 for x in slides if x.get('check')))
