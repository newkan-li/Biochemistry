import json,re,os
M=json.loads(re.search(r'window\.MANIFEST=(\[.*\]);?',open('data/manifest.js').read(),re.S).group(1))
BC=[x for x in M if x['id'].startswith('bc')]
def chdata(cid):
 s=open(f'data/{cid}.js',encoding='utf-8').read();return json.loads(re.search(r'=(\{.*\});?\s*$',s,re.S).group(1))
groups=[[BC[:14],'第一篇 生物分子：结构和功能（第1—14章）'],[BC[14:28],'第二篇 新陈代谢：途径和能学（第15—28章）'],[BC[28:],'第三篇 遗传信息：传递和表达（第29—36章）']]
# key chapter 6 section groups by PPT page ranges; anchor first slide index = page-1
bc06_groups=[(1,15,'1. 酶的定义与催化特点'),(16,24,'2. 酶的化学本质与组成'),(25,41,'3. 酶的命名与分类'),(42,47,'4. 酶的专一性'),(48,60,'5. 酶活力与测定'),(61,63,'6. 酶的分离纯化'),(64,70,'7. 核酶'),(71,74,'8. 抗体酶'),(75,87,'9. 酶工程与固定化酶')]
head='''<!doctype html><html lang="zh-CN"><head><meta name="robots" content="noindex,nofollow,noarchive,nosnippet"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="color-scheme" content="light dark"><title>生物化学 · 朱圣庚第4版 · 36章详细目录</title><link rel="stylesheet" href="assets/style.css?v=20261106a"><script src="assets/theme.js?v=20261106a"></script><style>
.bw{max-width:1180px;margin:0 auto;padding:20px 16px 50px}.bw h1{font-size:26px;margin:4px 0}.bw .sub{color:var(--muted);font-size:14px;margin:0 0 14px;line-height:1.7}.bw .banner{background:linear-gradient(135deg,#1f5c8b,#12324d);color:#fff;border-radius:14px;padding:14px 18px;margin:10px 0 16px}.bw .banner p{margin:5px 0 0;font-size:13px;line-height:1.7;opacity:.94}.bw h2{font-size:18px;margin:18px 0 8px}.bw .ch{background:var(--card);border:1px solid var(--line);border-radius:11px;padding:9px 12px;margin:7px 0}.bw .chh{display:flex;justify-content:space-between;gap:8px;align-items:center;flex-wrap:wrap}.bw .chh a{font-weight:700;color:#1f5c8b;text-decoration:none;font-size:14px}.bw .meta{font-size:11.5px;color:var(--muted)}.bw details{margin-top:3px}.bw summary{cursor:pointer;color:#7a2a24;font-size:12px;font-weight:600}.bw .subs{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:2px 12px;padding:6px 0 2px 10px}.bw .subs a{font-size:12px;color:var(--muted);text-decoration:none;line-height:1.5}.bw .subs a:hover{color:#1f5c8b}.bw .links{display:flex;gap:8px;flex-wrap:wrap;margin-top:14px}.bw .links a{text-decoration:none;color:#1f5c8b;border:1px solid #cbd9ea;background:#eef4fb;border-radius:16px;padding:5px 11px;font-size:13px}.bw .note{font-size:12.5px;color:var(--muted);line-height:1.8;background:var(--card);border:1px solid var(--line);border-radius:10px;padding:9px 12px;margin:10px 0}html[data-theme="dark"] .bw .ch,html[data-theme="dark"] .bw .note{background:#1b2531;border-color:#2c3a4a}html[data-theme="dark"] .bw .chh a{color:#a8c6e0}html[data-theme="dark"] .bw .links a{background:#22303d;border-color:#33475a;color:#a8c6e0}
</style></head><body><div class="bw"><div style="display:flex;justify-content:space-between;gap:10px;flex-wrap:wrap;align-items:center"><div><h1>生物化学 · 朱圣庚第4版</h1><div class="sub">36章详细目录 · 章节子节 · 讲义/PPT/习题/自测</div></div><a class="navbtn" style="width:auto;margin:0;text-decoration:none" href="index.html">← 首页</a></div><div class="banner"><b>按朱圣庚《生物化学》第4版36章组织</b><p>每章下面展开子节目录；点章节标题进入学习，点子节跳到对应 PPT/讲义页；每页都有掌握知识点、详细讲解和本页自测。</p></div>'''
body=''
for g,title in groups:
 body+=f'<h2>{title}</h2>'
 for x in g:
  cid=x['id'];d=chdata(cid);slides=d['modules'][0]['slides'];subs=''
  if cid=='bc06':
   for a,b,label in bc06_groups:
    idx=max(0,min(len(slides)-1,a-1));subs+=f'<a href="{cid}.html#s_{cid}_s0_{idx}">{label}</a>'
  else:
   for sl in slides:
    t=(sl.get('title') or f'PPT第{sl.get("i",0)+1}页').strip()
    subs+=f'<a href="{cid}.html#s_{cid}_s0_{sl.get("i",0)}">{t[:60]}</a>'
  body+=f'<div class="ch"><div class="chh"><a href="{cid}.html">{x["title"]}</a><span class="meta">{x["n_slides"]}页 · {x["n_q"]}题</span></div><details><summary>展开章节子节目录</summary><div class="subs">{subs}</div></details></div>'
body+='<div class="note"><b>说明：</b>子节目录来自各章讲义/PPT页；第6章按知识板块分组，其他章按PPT页列出。原始扫描笔记和完整PPT不上传公网。</div><div class="links"><a href="index.html">← 首页</a><a href="870-resources.html">资料分析</a><a href="870-ppt-map.html">PPT对照表</a></div></div></body></html>'
open('learn.html','w',encoding='utf-8').write(head+body)
print('learn toc written')
