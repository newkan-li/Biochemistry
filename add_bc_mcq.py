import os,re,json
OCR='/home/newkan/Documents/biology/biochemistry/_incoming/notes_ocr'
up_starts=[3,8,24,37,50,62,73,93,107,122,135,148,162,172,185]
down_starts=[3,13,24,38,52,67,77,83,94,101,126,139,148,166,181,194,215,227,247,269,287,313,318]
def pages(pref,a,b):
 out=[]
 for i in range(a,b):
  f=os.path.join(OCR,f'{pref}_{i:03d}.txt')
  if os.path.exists(f):out.append(open(f,encoding='utf-8',errors='ignore').read())
 return '\n'.join(out)
def parse_mcq(raw,prefix):
 # isolate multiple choice section
 m=re.search(r'(?:一、|一，|一\s*、)?\s*选择题',raw)
 if not m:return []
 seg=raw[m.end():]
 m2=re.search(r'\n\s*(?:二、|二，|二\s*、)\s*(?:填空题|判断|名词)',seg)
 if m2:seg=seg[:m2.start()]
 blocks=re.split(r'(?m)^\s*(\d{1,3})\s*[.、．]\s*',seg)
 out=[]
 for j in range(1,len(blocks),2):
  b=blocks[j+1]
  # stop if next section text
  ans=re.search(r'【答案】\s*([A-D]|[对错])',b)
  if not ans:continue
  body=b[:ans.start()]
  # options: A. xxx B. xxx
  opts=[]
  for om in re.finditer(r'(?<![A-Za-z])([A-E])[.、．]\s*([^A-E\n]{1,260})',body):
   txt=om.group(1)+'. '+om.group(2).strip()
   opts.append(txt)
  # dedupe by letter
  seen=set();opts2=[]
  for o in opts:
   letter=o[0]
   if letter not in seen:seen.add(letter);opts2.append(o)
  if len(opts2)<2 or ans.group(1) not in 'ABCDE':continue
  q=re.split(r'(?<![A-Za-z])[A-E][.、．]',body)[0].strip()
  q=re.sub(r'\s+',' ',q)
  if len(q)<5:continue
  out.append({'id':f'{prefix}_mcq_{len(out)+1}','modI':0,'q':q,'o':opts2,'a':ans.group(1),'e':'','kp':''})
  if len(out)>=8:break
 return out
for num in range(1,37):
 cid=f'bc{num:02d}'
 if num<=14:a=up_starts[num-1];b=up_starts[num] if num<14 else 185;raw=pages('up',a,b)
 else:idx=num-15;a=down_starts[idx];b=down_starts[idx+1];raw=pages('down',a,b)
 mcq=parse_mcq(raw,cid)
 if not mcq:continue
 p=f'data/{cid}.js';s=open(p,encoding='utf-8').read();m=re.search(r'=(\{.*\});?\s*$',s,re.S);ch=json.loads(m.group(1));ch['mcq']=mcq;ch['modules'][0]['mcq']=mcq;open(p,'w',encoding='utf-8').write('window.CHAPTERS=window.CHAPTERS||{};window.CHAPTERS["'+cid+'"]='+json.dumps(ch,ensure_ascii=False)+';\n');print(cid,len(mcq))
