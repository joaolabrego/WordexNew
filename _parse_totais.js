const fs = require('fs');
const path = require('path');

const base = 'D:/WordexNew/_xlsm_live/xl';
const shared = [];
const ss = fs.readFileSync(path.join(base, 'sharedStrings.xml'), 'utf8');
for (const m of ss.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)) {
  shared.push(m[1].replace(/&quot;/g, '"').replace(/&amp;/g, '&'));
}

function parseSheet(file) {
  const xml = fs.readFileSync(path.join(base, 'worksheets', file), 'utf8');
  const si = xml.match(/<sheetData>([\s\S]*)<\/sheetData>/)[1];
  const rows = [...si.matchAll(/<row r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)];
  const grid = {};
  for (const row of rows) {
    const r = row[1];
    const cells = [
      ...row[2].matchAll(
        /<c r="([A-Z]+)(\d+)"([^>]*)>(?:<f[^>]*>([\s\S]*?)<\/f>)?(?:<v>([\s\S]*?)<\/v>)?/g
      ),
    ];
    for (const c of cells) {
      const ref = c[1] + c[2];
      const type = (c[3].match(/t="([^"]+)"/) || [])[1] || '';
      const v = c[5] || '';
      grid[ref] = type === 's' ? shared[Number(v)] : v;
    }
  }
  return grid;
}

const map = {
  TotaisProdutos: 'sheet5.xml',
  TotaisProdutosGrafico: 'sheet6.xml',
  TotaisProdutosGerais: 'sheet7.xml',
};

for (const [name, file] of Object.entries(map)) {
  const g = parseSheet(file);
  console.log('\n=== ' + name + ' ===');
  const refs = Object.keys(g).sort((a, b) => {
    const ra = parseInt(a.slice(1), 10);
    const rb = parseInt(b.slice(1), 10);
    if (ra !== rb) return ra - rb;
    return a.localeCompare(b);
  });
  for (const ref of refs) {
    console.log(ref + ': ' + (g[ref] || '').toString().slice(0, 80));
  }
}
