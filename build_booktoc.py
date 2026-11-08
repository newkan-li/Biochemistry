import fitz,glob,json,re,os
base='/home/newkan/Documents/biology/生物化学 (第4版) 上下册 by 朱圣庚  徐长法+笔记和课后习题（含考研真题）详解'
res={}
def clean(t):return re.sub(r'\s+',' ',t.replace('\u3000',' ')).strip()
# upper: chapter level1, sections level2
p=[x for x in glob.glob(base+'/*上册*.pdf') if '笔记' not in os.path.basename(x)][0];d=fitz.open(p);toc=d.get_toc(simple=True);cur=None
for lvl,title,page in toc:
 t=clean(title)
 m=re.match(r'第(\d+)章',t)
 if lvl==1 and m:
  cur='bc%02d'%int(m.group(1));res[cur]={'page':page,'title':t,'sections':[]}
 elif cur and lvl==2 and t and not re.match(r'(提要|习题|主要参考书目|索引)',t):
  res[cur]['sections'].append({'title':t,'page':page})
 elif cur and lvl==1 and t.startswith('第') and not m:
  cur=None
# lower: chapter level2, sections level3
p=[x for x in glob.glob(base+'/*下册*.pdf') if '笔记' not in os.path.basename(x)][0];d=fitz.open(p);toc=d.get_toc(simple=True);cur=None
for lvl,title,page in toc:
 t=clean(title)
 m=re.match(r'第(\d+)章',t)
 if lvl==2 and m:
  cur='bc%02d'%int(m.group(1));res[cur]={'page':page,'title':t,'sections':[]}
 elif cur and lvl==3 and t and not re.match(r'(提要|习题|主要参考书目|索引)',t):
  res[cur]['sections'].append({'title':t,'page':page})
 elif cur and lvl<=2 and not m and re.match(r'第[一二三四五六七八九十]+篇',t):
  cur=None
for k in sorted(res):print(k,res[k]['page'],len(res[k]['sections']),res[k]['sections'][:2])
open('data/booktoc.js','w',encoding='utf-8').write('window.BOOKTOC='+json.dumps(res,ensure_ascii=False)+';\n')
print('entries',len(res))
