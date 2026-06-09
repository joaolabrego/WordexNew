const fs = require('fs');

function readJson(p) {
  let t = fs.readFileSync(p, 'utf8');
  if (t.charCodeAt(0) === 0xfeff) t = t.slice(1);
  return JSON.parse(t);
}

function listIssues(a, b, path = '') {
  const out = [];
  if (Array.isArray(a) && Array.isArray(b)) {
    if (a.length !== b.length) out.push({ path, issue: 'array length', a: a.length, b: b.length });
    for (let i = 0; i < Math.max(a.length, b.length); i++) {
      if (!a[i]) out.push({ path: `${path}[${i}]`, issue: 'missing in wordex' });
      else if (!b[i]) out.push({ path: `${path}[${i}]`, issue: 'extra in wordex' });
      else out.push(...listIssues(a[i], b[i], `${path}[${i}]`));
    }
    return out;
  }
  if (a && b && typeof a === 'object' && typeof b === 'object' && !Array.isArray(a)) {
    for (const k of Object.keys(b)) {
      if (!(k in a)) out.push({ path: path ? `${path}.${k}` : k, issue: 'missing in wordex' });
      else out.push(...listIssues(a[k], b[k], path ? `${path}.${k}` : k));
    }
    for (const k of Object.keys(a)) {
      if (!(k in b)) out.push({ path: path ? `${path}.${k}` : k, issue: 'extra in wordex' });
    }
    return out;
  }
  if (JSON.stringify(a) !== JSON.stringify(b)) {
    const short = (v) => {
      const s = typeof v === 'string' ? v : JSON.stringify(v);
      return s.length > 90 ? s.slice(0, 90) + `…[${s.length}]` : s;
    };
    out.push({ path: path || '(root)', issue: 'value differs', a: short(a), b: short(b) });
  }
  return out;
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
const issues = listIssues(a, b);
const erros = findErros(a);

console.log('=== COMPARACAO ===');
console.log('wordex.json:        ', fs.statSync('D:/WordexNew/wordex.json').size, 'bytes');
console.log('wordex-correto.json:', fs.statSync('D:/WordexNew/wordex-correto.json').size, 'bytes');
console.log('Identicos:          ', issues.length === 0 ? 'SIM' : 'NAO');
console.log('Diferencas totais:  ', issues.length);
console.log('Erros no wordex:    ', erros.length);

const groups = {};
issues.forEach((i) => {
  groups[i.issue] = (groups[i.issue] || 0) + 1;
});
console.log('Por tipo:', groups);

console.log('\n--- Extras em wordex (nao existem no correto) ---');
issues.filter((i) => i.issue === 'extra in wordex').forEach((i) => console.log(' +', i.path));

console.log('\n--- Ausentes em wordex ---');
issues.filter((i) => i.issue === 'missing in wordex').forEach((i) => console.log(' -', i.path));

console.log('\n--- Valores diferentes ---');
issues.filter((i) => i.issue === 'value differs').forEach((i) => {
  console.log(` * ${i.path}`);
  console.log(`   wordex:  ${i.a}`);
  console.log(`   correto: ${i.b}`);
});

if (erros.length) {
  console.log('\n--- Objetos Erro ---');
  erros.forEach((e) => console.log(' !', e.path, '=>', e.erro));
}
