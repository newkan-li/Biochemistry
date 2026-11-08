import fitz,os,json,re,glob
BASE='/home/newkan/Documents/biology/biochemistry/Biochemistry';PDF='/tmp/opencode/pptpdf'
# Chinese Wang PPT mapping
src={2:'3.pdf',3:'6.pdf',4:'6.pdf',5:'7.pdf',6:'8.pdf',7:'9.pdf',8:'10.pdf',9:'1.pdf',10:'2.pdf',11:'13.pdf',12:'14.pdf',13:'11.pdf',14:'17.pdf',15:'19.pdf',16:'20.pdf',17:'22.pdf',25:'30.pdf',26:'32.pdf',32:'36.pdf',33:'38.pdf',34:'39.pdf',35:'40.pdf'}
for num,pdfname in src.items():
 cid=f'bc{num:02d}';d=fitz.open(os.path.join(PDF,pdfname));outdir=os.path.join(BASE,'assets','img',cid);os.makedirs(outdir,exist_ok=True)
 fp=os.path.join(BASE,'data',cid+'.js');s=open(fp,encoding='utf-8').read();ch=json.loads(re.search(r'=(\{.*\});?\s*$',s,re.S).group(1));slides=ch['modules'][0]['slides'];n=min(6,len(slides),d.page_count-1)
 for i in range(n):
  pg=i+1;pix=d[pg].get_pixmap(dpi=100,colorspace=fitz.csRGB);fn=f'assets/img/{cid}/ppt{i+1}.jpg';pix.save(os.path.join(BASE,fn),jpg_quality=80);slides[i]['img']=fn;slides[i]['page']=pg+1;slides[i]['fig']=f'中文PPT：王镜岩《生物化学》课件 · 第{pg+1}页'
 ch['modules'][0]['slides']=slides;open(fp,'w',encoding='utf-8').write('window.CHAPTERS=window.CHAPTERS||{};window.CHAPTERS["'+cid+'"]='+json.dumps(ch,ensure_ascii=False)+';\n');print('cn',cid,pdfname)
# Rebuild notes-page visuals for chapters without usable Chinese PPT
notes=[1,18,19,20,21,22,23,24,27]
BOOKS='/home/newkan/Documents/biology/生物化学 (第4版) 上下册 by 朱圣庚  徐长法+笔记和课后习题（含考研真题）详解'
up=fitz.open(glob.glob(BOOKS+'/*上册*笔记*.pdf')[0]);down=fitz.open(glob.glob(BOOKS+'/*下册*笔记*.pdf')[0])
up_starts=[3,8,24,37,50,62,73,93,107,122,135,148,162,172,185];down_starts=[3,13,24,38,52,67,77,83,94,101,126,139,148,166,181,194,215,227,247,269,287,313,318]
for num in notes:
 cid=f'bc{num:02d}'
 if num<=14:doc=up;start=up_starts[num-1];end=up_starts[num] if num<14 else 185
 else:idx=num-15;doc=down;start=down_starts[idx];end=down_starts[idx+1]
 fp=os.path.join(BASE,'data',cid+'.js');s=open(fp,encoding='utf-8').read();ch=json.loads(re.search(r'=(\{.*\});?\s*$',s,re.S).group(1));slides=ch['modules'][0]['slides'];n=min(6,len(slides),end-start)
 outdir=os.path.join(BASE,'assets','img',cid);os.makedirs(outdir,exist_ok=True)
 for i in range(n):
  pg=start+i;pix=doc[pg-1].get_pixmap(dpi=100,colorspace=fitz.csRGB);fn=f'assets/img/{cid}/note{i+1}.jpg';pix.save(os.path.join(BASE,fn),jpg_quality=78);slides[i]['img']=fn;slides[i]['page']=pg;slides[i]['fig']='中文笔记原页（该章暂无对应中文PPT）'
 ch['modules'][0]['slides']=slides;open(fp,'w',encoding='utf-8').write('window.CHAPTERS=window.CHAPTERS||{};window.CHAPTERS["'+cid+'"]='+json.dumps(ch,ensure_ascii=False)+';\n');print('notes',cid,n)
