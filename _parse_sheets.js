const fs = require('fs');
const path = require('path');

const base = 'D:/WordexNew/_xlsm_live/xl';
const dir = path.join(base, 'worksheets');

const shared = [];
const ss = fs.readFileSync(path.join(base, 'sharedStrings.xml'), 'utf8');
for (const m of ss.matchAll(/<si>(?:<t[^>]*>([\s\S]*?)<\/t>|<r>[\s\S]*?<\/r>)<\/si>/g)) {
  shared.push((m[1] || '').replace(/&quot;/g, '"').replace(/&amp;/g, '&'));
}

const wb = fs.readFileSync(path.join(base, 'workbook.xml'), 'utf8');
const rels = fs.readFileSync(path.join(base, '_rels/workbook.xml.rels'), 'utf8');
const ridToFile = {};
for (const m of rels.matchAll(/Id="(rId\d+)"[^>]+Target="worksheets\/(sheet\d+\.xml)"/g)) {
  ridToFile[m[1]] = m[2];
}
const sheets = [...wb.matchAll(/<sheet name="([^"]+)"[^>]+r:id="(rId\d+)"/g)];

function cellText(type, v, t) {
  if (type === 's') return shared[Number(v)] ?? ('s:' + v);
  return t || v || '';
}

function parseSheet(file, sheetName) {
  const xml = fs.readFileSync(path.join(dir, file), 'utf8');
  const si = xml.match(/<sheetData>([\s\S]*)<\/sheetData>/);
  if (!si) return;
  console.log('\n=== ' + sheetName + ' (' + file + ') ===');
  const rows = [...si[1].matchAll(/<row r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)];
  for (const row of rows.slice(0, 5)) {
    const r = row[1];
    const cells = [
      ...row[2].matchAll(
        /<c r="([A-Z]+)(\d+)"([^>]*)>(?:<f[^>]*>([\s\S]*?)<\/f>)?(?:<is>[\s\S]*?<t>([\s\S]*?)<\/t>[\s\S]*?<\/is>)?(?:<v>([\s\S]*?)<\/v>)?/g
      ),
    ];
    for (const c of cells) {
      const ref = c[1] + c[2];
      const attrs = c[3] || '';
      const type = (attrs.match(/t="([^"]+)"/) || [])[1] || '';
      const f = c[4] ? c[4].replace(/&quot;/g, '"').replace(/&amp;/g, '&') : '';
      const t = c[5] || '';
      const v = c[6] || '';
      const val = cellText(type, v, t);
      console.log('  ' + ref + ': ' + (f ? 'F=' + f : 'V=' + val));
    }
  }
}

for (const m of sheets) {
  const name = m[1];
  const file = ridToFile[m[2]];
  if (file) parseSheet(file, name);
}
