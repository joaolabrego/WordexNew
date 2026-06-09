const fs = require('fs');
function readJson(p) {
  let t = fs.readFileSync(p, 'utf8');
  if (t.charCodeAt(0) === 0xfeff) t = t.slice(1);
  return JSON.parse(t);
}
function findErros(obj, path = '') {
  const r = [];
  if (!obj || typeof obj !== 'object') return r;
  if (obj.Erro) r.push({ path, erro: obj.Erro });
  if (Array.isArray(obj)) obj.forEach((x, i) => r.push(...findErros(x, `${path}[${i}]`)));
  else Object.keys(obj).forEach((k) => r.push(...findErros(obj[k], path ? `${path}.${k}` : k)));
  return r;
}
const a = readJson('D:/WordexNew/wordex.json');
const b = readJson('D:/WordexNew/wordex-correto.json');
console.log('wordex.json:', fs.statSync('D:/WordexNew/wordex.json').size, 'bytes');
console.log('correto:', fs.statSync('D:/WordexNew/wordex-correto.json').size, 'bytes');
console.log('Identicos:', JSON.stringify(a) === JSON.stringify(b));
console.log('\nErros no wordex.json:');
findErros(a).forEach((e) => console.log(' -', e.path, '=>', e.erro));
