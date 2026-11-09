import json,re
M=json.loads(re.search(r'window\.MANIFEST=(\[.*\]);?',open('data/manifest.js').read(),re.S).group(1));BC=[x for x in M if x['id'].startswith('bc')]
BT=json.loads(re.search(r'window\.BOOKTOC=(\{.*\});?',open('data/booktoc.js').read(),re.S).group(1))
def chdata(cid):
 s=open(f'data/{cid}.js',encoding='utf-8').read();return json.loads(re.search(r'=(\{.*\});?\s*$',s,re.S).group(1))
groups=[[BC[:14],'第一篇 生物分子：结构和功能（第1—14章）'],[BC[14:28],'第二篇 新陈代谢：途径和能学（第15—28章）'],[BC[28:],'第三篇 遗传信息：传递和表达（第29—36章）']]
head='''<!doctype html><html lang="zh-CN"><head><meta name="robots" content="noindex,nofollow,noarchive,nosnippet"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="color-scheme" content="light dark"><title>生物化学 · 朱圣庚第4版 · 36章详细目录</title><link rel="stylesheet" href="assets/style.css?v=20261108cr"><script src="assets/theme.js?v=20261108cr"></script><style>.bw{max-width:1180px;margin:0 auto;padding:20px 16px 50px}.bw h1{font-size:26px;margin:4px 0}.bw .sub{color:var(--muted);font-size:14px;margin:0 0 14px;line-height:1.7}.bw .banner{background:linear-gradient(135deg,#1f5c8b,#12324d);color:#fff;border-radius:14px;padding:14px 18px;margin:10px 0 16px}.bw .banner p{margin:5px 0 0;font-size:13px;line-height:1.7;opacity:.94}.bw h2{font-size:18px;margin:18px 0 8px}.bw .ch{background:var(--card);border:1px solid var(--line);border-radius:11px;padding:9px 12px;margin:7px 0}.bw .chh{display:flex;justify-content:space-between;gap:8px;align-items:center;flex-wrap:wrap}.bw .chh a{font-weight:700;color:#1f5c8b;text-decoration:none;font-size:14px}.bw .meta{font-size:11.5px;color:var(--muted)}.bw details{margin-top:3px}.bw summary{cursor:pointer;color:#7a2a24;font-size:12px;font-weight:600}.bw .subs{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:2px 12px;padding:6px 0 2px 10px}.bw .subs a{font-size:12px;color:var(--muted);text-decoration:none;line-height:1.55}.bw .subs a:hover{color:#1f5c8b}.bw .pg{color:#9a6b00;font-size:11px;margin-left:4px}.bw .kbh{grid-column:1/-1;margin:8px 0 2px;padding:4px 8px;background:#eef4fb;border-left:3px solid #1f5c8b;border-radius:4px;font-size:12.5px;font-weight:700;color:#1f5c8b}.bw .kbh span{font-weight:400;color:var(--muted);font-size:11px;margin-left:6px}html[data-theme="dark"] .bw .kbh{background:#22303d;border-left-color:#5b9bd5;color:#a8c6e0}.bw .kb{grid-column:1/-1;margin:4px 0 3px 4px;border-left:2px solid #cbd9ea;padding-left:8px}.bw .kb>summary{cursor:pointer;font-size:12.5px;font-weight:700;color:#1f5c8b;list-style:none;line-height:1.55}.bw .kb>summary::-webkit-details-marker{display:none}.bw .kb>summary::before{content:"▸ ";color:#7a5a00}.bw .kb[open]>summary::before{content:"▾ "}.bw .kbsub{grid-column:1/-1;margin:3px 0 3px 12px;border-left:1px dashed #d5e3ef;padding-left:8px}.bw .kbsub>summary{cursor:pointer;font-size:12px;font-weight:600;color:#31516f;list-style:none;line-height:1.5}.bw .kbsub>summary::-webkit-details-marker{display:none}.bw .kbsub>summary::before{content:"· ";color:#7a5a00}.bw .kb span,.bw .kbsub span{font-weight:400;color:var(--muted);font-size:11px;margin-left:6px}html[data-theme="dark"] .bw .kb{border-left-color:#33475a}html[data-theme="dark"] .bw .kb>summary,html[data-theme="dark"] .bw .kbsub>summary{color:#a8c6e0}.bw .links{display:flex;gap:8px;flex-wrap:wrap;margin-top:14px}.bw .links a{text-decoration:none;color:#1f5c8b;border:1px solid #cbd9ea;background:#eef4fb;border-radius:16px;padding:5px 11px;font-size:13px}.bw .note{font-size:12.5px;color:var(--muted);line-height:1.8;background:var(--card);border:1px solid var(--line);border-radius:10px;padding:9px 12px;margin:10px 0}html[data-theme="dark"] .bw .ch,html[data-theme="dark"] .bw .note{background:#1b2531;border-color:#2c3a4a}html[data-theme="dark"] .bw .chh a{color:#a8c6e0}html[data-theme="dark"] .bw .links a{background:#22303d;border-color:#33475a;color:#a8c6e0}.bw .bsec{grid-column:1/-1;margin:2px 0 3px;border-left:2px solid #cbd9ea;padding-left:8px}.bw .bsec>summary{cursor:pointer;font-size:12.5px;font-weight:700;color:#31516f;list-style:none;line-height:1.55}.bw .bsec>summary::-webkit-details-marker{display:none}.bw .bsec>summary::before{content:"▸ ";color:#7a5a00}.bw .bsec[open]>summary::before{content:"▾ "}.bw .bsubs{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:1px 10px;padding:2px 0 2px 12px}html[data-theme="dark"] .bw .bsec{border-left-color:#33475a}html[data-theme="dark"] .bw .bsec>summary{color:#a8c6e0}</style></head><body><div class="bw"><div style="display:flex;justify-content:space-between;gap:10px;flex-wrap:wrap;align-items:center"><div><h1>生物化学 · 朱圣庚第4版</h1><div class="sub">36章详细目录 · 原书章节子节 + 本站PPT逐页入口</div></div><a class="navbtn" style="width:auto;margin:0;text-decoration:none" href="index.html">← 首页</a></div><div class="banner"><b>目录按朱圣庚《生物化学》第4版原书TOC</b><p>原书子节带页码；点章节进入学习，点子节进入该章；每章另附“本站PPT逐页目录”。</p></div>'''
body=''
for g,title in groups:
 body+=f'<h2>{title}</h2>'
 for x in g:
  cid=x['id'];d=chdata(cid);bt=BT.get(cid,{});secs=bt.get('sections',[]);slides=d['modules'][0]['slides']
  subs=''; nsub=0
  for sec in secs:
   pg=sec.get('page',0);pg='p'+str(pg) if pg and pg>0 else '—'
   ss=sec.get('subs',[]); nsub+=len(ss)
   head=f'{sec.get("title","")}<span class="pg">{pg}</span>'
   if ss:
    inner=''.join(f'<a href="{cid}.html">{z.get("title","")}<span class="pg">{("p"+str(z.get("page"))) if z.get("page",0)>0 else "—"}</span></a>' for z in ss)
    subs+=f'<details class="bsec"><summary>{head}</summary><div class="bsubs">{inner}</div></details>'
   else:
    subs+=f'<a href="{cid}.html">{head}</a>'
  psubs=''
  if cid=='bc06' and d.get('toc'):
   for bi,b in enumerate(d['toc']):
    psubs+=f'<details class="kb"{" open" if bi==0 else ""}><summary>{b.get("title","")}<span>{b.get("pages","")} · {b.get("end",0)-b.get("start",0)}页</span></summary>'
    for it in b.get('items',[]):
     psubs+=f'<details class="kbsub"><summary>{it.get("title","")}<span>{it.get("pages","")}</span></summary>'
     for sl in slides[int(it.get('start',0)):int(it.get('end',0))]:
      t=(sl.get('title') or f'PPT第{sl.get("i",0)+1}页').strip()
      psubs+=f'<a href="{cid}.html#s_{cid}_s0_{sl.get("i",0)}">{t[:55]}</a>'
     psubs+='</details>'
    psubs+='</details>'
   psum=f'本章知识块目录（{len(d["toc"])}块）'
  else:
   for sl in slides:
    t=(sl.get('title') or f'PPT第{sl.get("i",0)+1}页').strip()
    psubs+=f'<a href="{cid}.html#s_{cid}_s0_{sl.get("i",0)}">{t[:55]}</a>'
   psum=f'本站PPT逐页目录（{len(slides)}）'
  body+=f'<div class="ch" id="ch-{cid}"><div class="chh"><a href="{cid}.html">{x["title"]}</a><span class="meta">{x["n_slides"]}页PPT · {x["n_q"]}题</span></div><details open><summary>📖 原书章节子节（{len(secs)} 节 · {nsub} 子节）</summary><div class="subs">{subs or "（原书目录未提取到）"}</div></details><details><summary>📑 {psum}</summary><div class="subs">{psubs}</div></details></div>'
body+='<div class="note"><b>说明：</b>“原书章节子节”来自朱圣庚第4版PDF目录，页码为PDF页；本站PPT页与笔记不保证逐页对应原书页码，只按知识点组织。</div><div class="links"><a href="index.html">← 首页</a><a href="870-resources.html">资料分析</a><a href="870-ppt-map.html">PPT对照表</a></div></div></body></html>'
open('learn.html','w',encoding='utf-8').write(head+body)
print('learn original toc')
