import fitz,glob,json,re,os
BASE='/home/newkan/Documents/biology/biochemistry/Biochemistry'
BOOKS='/home/newkan/Documents/biology/生物化学 (第4版) 上下册 by 朱圣庚  徐长法+笔记和课后习题（含考研真题）详解'
up=fitz.open(glob.glob(BOOKS+'/*上册*笔记*.pdf')[0]); down=fitz.open(glob.glob(BOOKS+'/*下册*笔记*.pdf')[0])
up_starts=[3,8,24,37,50,62,73,93,107,122,135,148,162,172,185]
down_starts=[3,13,24,38,52,67,77,83,94,101,126,139,148,166,181,194,215,227,247,269,287,313,318]
for num in range(1,37):
 cid=f'bc{num:02d}';p=os.path.join(BASE,cid+'.html')
 # parse embedded JSON
 s=open(os.path.join(BASE,'data',cid+'.js'),encoding='utf-8').read()
 m=re.search(r'=(\{.*\});?\s*$',s,re.S);ch=json.loads(m.group(1))
 slides=ch['modules'][0]['slides']
 if num<=14: doc=up;start=up_starts[num-1];end=up_starts[num] if num<14 else 185
 else: doc=down;idx=num-15;start=down_starts[idx];end=down_starts[idx+1]
 outdir=os.path.join(BASE,'assets','img',cid);os.makedirs(outdir,exist_ok=True)
 n=min(len(slides),6)
 pages=list(range(start,min(end,start+n)))
 for i,pg in enumerate(pages):
  # 1-based page index; fitz zero-based
  pix=doc[pg-1].get_pixmap(dpi=100,colorspace=fitz.csRGB)
  fn=os.path.join(outdir,f'p{pg:03d}.jpg')
  pix.save(fn,jpg_quality=78)
  slides[i]['img']=f'assets/img/{cid}/p{pg:03d}.jpg'
  slides[i]['page']=pg
  slides[i]['fig']=(slides[i].get('fig') or '')+' · 笔记原页'
 # remaining slides reuse last image? no
 s2='window.CHAPTERS=window.CHAPTERS||{};window.CHAPTERS["'+cid+'"]='+json.dumps(ch,ensure_ascii=False)+';\n'
 open(os.path.join(BASE,'data',cid+'.js'),'w',encoding='utf-8').write(s2)
 print(cid,len(pages),'pages')
