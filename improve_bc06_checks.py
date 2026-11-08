import json,re
p='data/bc06.js';s=open(p,encoding='utf-8').read();ch=json.loads(re.search(r'=(\{.*\});?\s*$',s,re.S).group(1));slides=ch['modules'][0]['slides'];qs=ch['short']
# page -> after-class short question index (0-based)
m={7:0,15:2,18:3,21:3,30:4,41:4,42:1,49:5,55:5,60:5,64:6,71:6,75:7,57:8,59:9,62:10,63:11}
for pg,idx in m.items():
 if 0<=idx<len(qs):
  q=qs[idx];sl=slides[pg-1];sl['check']={'kind':'short','id':q['id'],'q':q['q'],'a':q['a'],'e':q.get('e',''),'src':'课后习题'}
for sl in slides:
 if not sl.get('check'):
  sl['check']={'kind':'short','id':sl.get('check',{}).get('id') or f"bc06_self_{sl.get('i')}",'q':f"本页考点：请说明「{sl.get('title','本页')}」的核心内容。",'a':sl.get('notes') or '本页为过渡页，无文字要点。','e':''}
 # ensure page-specific answer for existing generic
 if sl.get('check',{}).get('kind')=='short' and not sl.get('check',{}).get('src'):
  sl['check']['q']=f"本页考点：请说明「{sl.get('title','本页')}」的核心内容。"
  sl['check']['a']=sl.get('notes') or '本页为过渡页，无文字要点。'
ch['modules'][0]['slides']=slides
open(p,'w',encoding='utf-8').write('window.CHAPTERS=window.CHAPTERS||{};window.CHAPTERS["bc06"]='+json.dumps(ch,ensure_ascii=False)+';\n')
print('mapped',len(m),'checks',sum(1 for x in slides if x.get('check')))
