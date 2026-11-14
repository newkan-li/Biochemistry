import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
const root = path.resolve('data');
const rows = [];
for (const file of fs.readdirSync(root).filter(f => /^bc\d+\.js$/.test(f)).sort()) {
  const code = fs.readFileSync(path.join(root, file), 'utf8');
  const sandbox = { window: {} };
  try { vm.runInNewContext(code, sandbox); } catch { continue; }
  const id = file.replace(/\.js$/, '');
  const c = sandbox.window.CHAPTERS?.[id];
  if (!c) continue;
  const conceptual = (c.exercises || []).filter(q => !['选择', '计算'].includes(q.type));
  const same = conceptual.filter(q => String(q.answer || '').trim() && String(q.answer).trim() === String(q.analysis || '').trim());
  const short = conceptual.filter(q => String(q.answer || '').trim().length < 15);
  const emptyE = [];
  for (const k of ['mcq', 'judge', 'fill']) for (const q of c[k] || []) if (!String(q.e || '').trim()) emptyE.push(q.id);
  for (const m of c.modules) for (const s of m.slides) if (s.check && !String(s.check.e || '').trim()) emptyE.push(s.check.id);
  if (same.length || short.length || emptyE.length) rows.push({ chapter: id, webConceptual: conceptual.length, answerEqualsAnalysis: same.length, answerUnder15: short.length, emptyExplanation: emptyE.length });
}
console.table(rows);
console.log('TOTAL', rows.reduce((a, r) => ({ answerEqualsAnalysis: a.answerEqualsAnalysis + r.answerEqualsAnalysis, answerUnder15: a.answerUnder15 + r.answerUnder15, emptyExplanation: a.emptyExplanation + r.emptyExplanation }), { answerEqualsAnalysis: 0, answerUnder15: 0, emptyExplanation: 0 }));
