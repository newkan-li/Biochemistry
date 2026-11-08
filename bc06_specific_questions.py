import json,re
p='data/bc06.js';s=open(p,encoding='utf-8').read();ch=json.loads(re.search(r'=(\{.*\});?\s*$',s,re.S).group(1));slides=ch['modules'][0]['slides']
# page -> specific question; answers use that page's PPT text
M={
2:'酶的认识最早可追溯到哪些中国古代生产和医药实践？',
5:'enzyme 一词来自哪两个希腊词根？字面含义是什么？',
13:'过氧化氢酶催化双氧水裂解，说明酶的什么特性？',
16:'哪些实验/性质证据说明酶的化学本质是蛋白质？',
17:'T.Cech 发现的第一个有催化活性的天然 RNA 是什么？',
19:'羧肽酶属于哪类酶？它从肽链哪一端水解肽键？',
20:'NADPH-脱氢酶和过氧化氢酶分别需要什么辅酶/辅基？',
22:'酶蛋白和辅助因子在催化中各自起什么作用？',
25:'酶的惯用命名法主要依据哪些原则？',
26:'习惯命名法有什么优点和缺点？',
27:'国际系统命名法包括哪些内容？',
28:'以乳酸脱氢酶为例，说明系统名的组成。',
29:'系统名和惯用名有什么区别？',
33:'氧化还原酶催化什么反应？主要包括哪两类？',
34:'脱氢酶和氧化酶催化的反应有何不同？',
35:'加氧酶分为哪两类？分别催化什么反应？',
36:'转移酶催化什么反应？举例说明。',
37:'水解酶催化什么反应？举出常见例子。',
38:'裂合酶催化什么反应？',
39:'异构酶催化什么反应？',
40:'连接酶（合成酶）催化什么反应？为什么必须与 ATP 偶联？',
43:'解释酶专一性的学说主要有哪些？',
44:'锁钥学说的核心内容是什么？有什么局限？',
46:'羧肽酶的诱导契合模式说明了什么？',
47:'三点结合催化理论如何解释酶的立体专一性？',
50:'测定酶活力时为什么通常选择反应初速度？',
52:'酶活力单位（U）是如何定义的？',
53:'国际单位 IU 的定义是什么？',
54:'Kat 单位如何定义？1 Kat 等于多少 IU？',
56:'转换数（turnover number）的定义是什么？',
58:'酶活力测定常用哪些方法？各依据什么信号？',
61:'酶的分离纯化通常包括哪些步骤？',
65:'L19RNA 具备哪些酶促催化特征？',
66:'L19RNA 除核酶活性外，还表现出哪些催化作用？',
67:'1997 年 Zhang 和 Cech 的实验证明 RNA 具有什么活性？',
68:'核酶的自我剪切和自我剪接有何区别？',
69:'核酶的典型结构特点包括哪些？',
70:'核酶的研究意义和应用前景是什么？',
72:'第一代抗体酶是如何获得的？',
73:'第二代和第三代抗体酶分别采用什么策略？',
77:'化学修饰酶常修饰哪些氨基酸残基？目的是什么？',
78:'酶分子交联反应的作用是什么？',
79:'大分子修饰酶蛋白侧链有什么作用？',
81:'固定化酶有哪些主要优点？',
82:'固定化酶常用哪些方法？',
83:'酶电极由哪些部分组成？其工作原理是什么？',
84:'固定化葡萄糖异构酶在工业上有什么应用？',
86:'生物酶工程的定义和主要优点是什么？',
87:'生物酶工程改造和生产酶有哪几条路线？'
}
for page,q in M.items():
 if 1<=page<=len(slides):
  sl=slides[page-1]
  sl['check']={'kind':'short','id':f'bc06_p{page}','q':q,'a':sl.get('notes') or '本页无文字说明。','e':'','src':'PPT内容自测'}
# remove old generic checks where no meaningful page question; keep mapped, existing mcq, after-class
for sl in slides:
 c=sl.get('check') or {}
 if c.get('kind')=='short' and not c.get('src'):
  sl['check']=None
 if sl['check'] is None and sl.get('page') in M:
  pass
ch['modules'][0]['slides']=slides
open(p,'w',encoding='utf-8').write('window.CHAPTERS=window.CHAPTERS||{};window.CHAPTERS["bc06"]='+json.dumps(ch,ensure_ascii=False)+';\n')
print('specific',len(M),'checks',sum(1 for x in slides if x.get('check')))
