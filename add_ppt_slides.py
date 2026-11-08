import fitz,os,json,re
BASE='/home/newkan/Documents/biology/biochemistry/Biochemistry';PDF='/tmp/opencode/pptpdf'
# source PDF per chapter, only where a matching PPT exists
src={
1:'0 绪论.pdf',2:'01 Amino acids.pdf',3:'02 Structure of proteins.pdf',4:'03 Structure and function relationships of proteins.pdf',5:'04 Protein properties,classification and research methods.pdf',
6:'08 Introduction of enzymology.pdf',7:'09 Enzyme kinetics.pdf',8:'10 Mechanisms of enzyme catalysis.pdf',9:'15 Carbohydrates.pdf',10:'16 Lipids and biological membranes.pdf',
11:'06 Structure and function of Nucleic acids.pdf',12:'07 Properties of Nucleic acids and Research methods.pdf',13:'14 Vitamins and coenzymes.pdf',14:'17 激素及其受体介导的信号转导.pdf',
15:'19.pdf',16:'20.pdf',17:'22 糖酵解.pdf',18:'23 三羧酸循环.pdf',19:'20 生物氧化.pdf',20:'24 磷酸戊糖途径.pdf',
21:'25 糖异生.pdf',22:'27 糖原代谢.pdf',23:'26 光合作用.pdf',24:'29 脂肪酸代谢.pdf',25:'30.pdf',26:'32.pdf',27:'32 核苷酸代谢.pdf'
}
for num,pdfname in src.items():
 cid=f'bc{num:02d}';path=os.path.join(PDF,pdfname)
 if not os.path.exists(path):print('missing',num,pdfname);continue
 d=fitz.open(path);outdir=os.path.join(BASE,'assets','img',cid);os.makedirs(outdir,exist_ok=True)
 # parse current data
 fp=os.path.join(BASE,'data',cid+'.js');s=open(fp,encoding='utf-8').read();m=re.search(r'=(\{.*\});?\s*$',s,re.S);ch=json.loads(m.group(1));slides=ch['modules'][0]['slides']
 n=min(6,len(slides),max(1,d.page_count-1))
 for i in range(n):
  pg=i+1 # skip cover: PDF page 2 -> index1; use i+1 zero-based = page 2
  if pg>=d.page_count:pg=d.page_count-1
  pix=d[pg].get_pixmap(dpi=100,colorspace=fitz.csRGB)
  fn=f'assets/img/{cid}/ppt{i+1}.jpg';pix.save(os.path.join(BASE,fn),jpg_quality=80)
  slides[i]['img']=fn;slides[i]['page']=pg+1;slides[i]['fig']=f'对应PPT：{pdfname.replace(".pdf","")} · 第{pg+1}页'
 ch['modules'][0]['slides']=slides
 open(fp,'w',encoding='utf-8').write('window.CHAPTERS=window.CHAPTERS||{};window.CHAPTERS["'+cid+'"]='+json.dumps(ch,ensure_ascii=False)+';\n')
 print(cid,pdfname,d.page_count,'imgs',n)
