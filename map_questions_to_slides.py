import json,re,math
ids=[f'bc{i:02d}' for i in range(1,37)]
def grams(s):
 s=re.sub(r'[^\u4e00-\u9fffA-Za-z0-9]+',' ',s).lower()
 out=set()
 for tok in s.split():
  if len(tok)>=2:
   for i in range(len(tok)-1):out.add(tok[i:i+2])
 return out
for cid in ids:
 if cid=='bc06':continue
 p=f'data/{cid}.js';s=open(p,encoding='utf-8').read();ch=json.loads(re.search(r'=(\{.*\});?\s*$',s,re.S).group(1));slides=ch['modules'][0]['slides'];qs=ch.get('short') or []
 used=set();mapped=0
 for qi,q in enumerate(qs):
  qg=grams((q.get('q') or '')+' '+(q.get('a') or '')[:300])
  if len(qg)<5:continue
  best=None;bs=0
  for si,sl in enumerate(slides):
   if si in used:continue
   sg=grams((sl.get('title') or '')+' '+(sl.get('notes') or '')[:1200])
   inter=len(qg & sg)
   if not inter:continue
   # favor title overlap and shorter content
   score=inter/(math.sqrt(len(qg))*math.sqrt(max(1,len(sg))))
   if score>bs:bs=score;best=si
  if best is not None and bs>=0.035:
   sl=slides[best];sl['check']={'kind':'short','id':q.get('id'),'q':q.get('q'),'a':q.get('a'),'e':q.get('e',''),'src':'配套笔记课后习题'};used.add(best);mapped+=1
 for sl in slides:
  if not sl.get('check'):
   sl['check']={'kind':'short','id':f"{cid}_self_{sl.get('i')}",'q':f"本页考点：请说明「{sl.get('title','本页')}」的核心内容。",'a':sl.get('notes') or '本页为过渡页，无文字要点。','e':''}
 ch['modules'][0]['slides']=slides
 open(p,'w',encoding='utf-8').write('window.CHAPTERS=window.CHAPTERS||{};window.CHAPTERS["'+cid+'"]='+json.dumps(ch,ensure_ascii=False)+';\n')
 print(cid,'mapped',mapped,'of',len(qs),'slides',len(slides),'checks',sum(1 for x in slides if x.get('check')))
