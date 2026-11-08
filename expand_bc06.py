import fitz,os,json,re
BASE='/home/newkan/Documents/biology/biochemistry/Biochemistry';PDF='/tmp/opencode/pptpdf/8.pdf';d=fitz.open(PDF)
# page: (first line/title hint, question, options, answer, explanation)
items=[
(6,'什么是酶?','酶的化学本质（除少数RNA外）主要是','A. 蛋白质|B. 糖类|C. 脂质|D. 核酸','A','酶是活细胞产生的生物催化剂，绝大多数酶的化学本质是蛋白质。'),
(7,'酶与一般催化剂的共同点','酶与一般催化剂的共同点是','A. 改变平衡常数|B. 降低反应活化能|C. 催化本来不能发生的反应|D. 必须在高温下工作','B','二者都降低活化能、加快反应，但不改变平衡常数。'),
(10,'活化能','酶加速反应的主要机制是','A. 提高活化能|B. 降低活化能|C. 改变平衡常数|D. 增加底物浓度','B','酶稳定过渡态，降低反应所需活化自由能。'),
(11,'酶易失活','下列最易使酶失活的条件是','A. 常温|B. 最适pH|C. 高温、强酸、强碱|D. 低浓度底物','C','高温、强酸、强碱、重金属等使蛋白质变性，酶失活。'),
(12,'酶的高效性','酶催化效率通常比非催化反应高','A. 2倍|B. 10—100倍|C. 10^2—10^3倍|D. 10^6—10^16倍','D','酶具有极高催化效率，通常可提高10^6—10^16倍。'),
(14,'酶的专一性','酶作用的物质称为','A. 辅酶|B. 底物|C. 产物|D. 抑制剂','B','底物是酶作用的物质；专一性指酶对底物的严格选择。'),
(15,'酶活性可调节','下列不属于酶活性调节方式的是','A. 反馈抑制|B. 共价修饰|C. 别构调节|D. 改变DNA序列','D','酶活性可通过反馈抑制、别构、共价修饰、激素等调节；改变DNA序列属于遗传变异。'),
(18,'酶的化学组成','全酶由什么组成','A. 酶蛋白+辅因子|B. 酶蛋白+底物|C. 辅酶+金属离子|D. 核酸+蛋白质','A','全酶=酶蛋白+辅因子；酶蛋白决定专一性，辅因子参与催化。'),
(21,'酶蛋白与辅因子','决定酶促反应专一性的是','A. 辅酶|B. 金属离子|C. 酶蛋白|D. 底物浓度','C','酶蛋白的结构决定活性中心，从而决定专一性。'),
(24,'寡聚酶与多酶复合体','由几种酶靠非共价键嵌合、催化连续反应的复合物是','A. 单体酶|B. 寡聚酶|C. 多酶复合体|D. 同工酶','C','多酶复合体由多种酶非共价嵌合，有利于连续反应高效进行。'),
(30,'酶的分类','国际酶学委员会按反应类型将酶分为','A. 4大类|B. 5大类|C. 6大类|D. 7大类','C','氧化还原酶、转移酶、水解酶、裂合酶、异构酶、连接酶六大类。'),
(41,'EC编号','EC编号的第一位数字表示','A. 亚类|B. 大类|C. 亚亚类|D. 排号','B','EC编号四位依次为大类、亚类、亚亚类和排号。'),
(42,'酶的专一性类型','酶只作用于一种底物，属于','A. 绝对专一性|B. 相对专一性|C. 立体专一性|D. 无专一性','A','绝对专一性指酶只作用于一种底物。'),
(45,'诱导契合','诱导契合学说认为','A. 酶活性中心完全刚性|B. 底物诱导酶构象改变|C. 酶改变平衡常数|D. 酶不需要活性中心','B','底物与酶接触后诱导酶活性中心构象改变，形成互补契合。'),
(49,'酶活力','酶活力通常用什么表示','A. 酶浓度|B. 酶促反应速率|C. 蛋白质含量|D. 分子量','B','酶活力用一定条件下酶催化反应的速率表示。'),
(55,'比活力','比活力的常用单位是','A. U/mg蛋白|B. U/mL|C. mol/L|D. kDa','A','比活力=每毫克蛋白质所含的酶活力单位数，反映纯度。'),
(60,'分离纯化指标','纯化倍数等于','A. 每次总活力/第一次总活力|B. 每次比活力/第一次比活力|C. 每次蛋白量/第一次蛋白量|D. 每次体积/第一次体积','B','纯化倍数=某步比活力/第一次比活力；回收率=某步总活力/第一次总活力。'),
(64,'核酶','核酶是','A. 有催化活性的RNA|B. 有催化活性的蛋白质|C. 辅酶|D. 底物','A','核酶是具有催化功能的RNA，支持RNA世界假说。'),
(71,'抗体酶','抗体酶的本质是','A. 免疫球蛋白|B. RNA|C. 多糖|D. 脂质','A','抗体酶是赋予免疫球蛋白催化活性的催化性抗体。'),
(80,'固定化酶','固定化酶的主要优点是','A. 不可重复使用|B. 易与产物分离、可重复使用|C. 稳定性下降|D. 增加污染','B','固定化酶不溶于水，便于分离和重复利用，稳定性提高。')
]
# load existing question arrays
fp=os.path.join(BASE,'data','bc06.js');s=open(fp,encoding='utf-8').read();ch=json.loads(re.search(r'=(\{.*\});?\s*$',s,re.S).group(1))
slides=[]
for n,(pg,title,q,opts,ans,exp) in enumerate(items,1):
 t=d[pg-1].get_text('text').strip();t=re.sub(r'\s+',' ',t).strip()
 # title from first nonempty line / given hint
 lines=[x.strip() for x in d[pg-1].get_text('text').splitlines() if x.strip()]
 stitle=lines[0][:50] if lines else title
 points=[x.strip() for x in re.split(r'[。；;]',t) if len(x.strip())>8][:3]
 pix=d[pg-1].get_pixmap(dpi=100,colorspace=fitz.csRGB);fn=f'assets/img/bc06/ppt{n}.jpg';pix.save(os.path.join(BASE,fn),jpg_quality=80)
 check={'id':f'bc06_check_{n}','q':q,'o':opts.split('|'),'a':ans,'e':exp,'kp':f'bc06_s0_{n-1}'}
 slides.append({'i':n-1,'sec':'第6章 酶通论·逐页PPT','title':stitle or title,'notes':t[:1400],'points':points,'fig':f'中文PPT：王镜岩《生物化学》课件 · 第{pg}页','anim':None,'check':check,'cards':None,'img':fn,'page':pg})
ch['modules'][0]['slides']=slides
ch['modules'][0]['short']=ch.get('short',[])
ch['preview']={'main':'第6章按老师PPT逐页学习：从酶的定义、催化特点、化学组成、分类命名、专一性、活力测定到核酶、抗体酶和酶工程。每页PPT都配一道本页考点题。','questions':[x[2] for x in items[:4]]}
ch['sub2']='第6章 · 老师PPT逐页 + 本页考点题 + 课后题/真题'
open(fp,'w',encoding='utf-8').write('window.CHAPTERS=window.CHAPTERS||{};window.CHAPTERS["bc06"]='+json.dumps(ch,ensure_ascii=False)+';\n')
print('slides',len(slides),'checks',sum(1 for x in slides if x.get('check')))
