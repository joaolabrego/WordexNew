const fs = require('fs');
const ss = fs.readFileSync('D:/WordexNew/_xlsm_live/xl/sharedStrings.xml', 'utf8');
const shared = [...ss.matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => {
  const parts = m[1].match(/<t[^>]*>([\s\S]*?)<\/t>/g);
  if (!parts) return '';
  return parts
    .map((x) => x.replace(/<\/?t[^>]*>/g, ''))
    .join('')
    .replace(/&amp;/g, '&');
});

const names = {
  'sheet5.xml': 'TotaisProdutos',
  'sheet6.xml': 'TotaisProdutosGrafico',
  'sheet7.xml': 'TotaisProdutosGerais',
};

for (const [file, label] of Object.entries(names)) {
  const xml = fs.readFileSync(`D:/WordexNew/_xlsm_live/xl/worksheets/${file}`, 'utf8');
  const rows = [
    ...xml.match(/<sheetData>([\s\S]*)<\/sheetData>/)[1].matchAll(/<row r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g),
  ];
  console.log('\n=== ' + label + ' ===');
  for (const row of rows) {
    const cells = [
      ...row[2].matchAll(
        /<c r="([A-Z]+)(\d+)"([^>]*)>(?:<f[^>]*>([\s\S]*?)<\/f>)?(?:<v>([\s\S]*?)<\/v>)?/g
      ),
    ];
    for (const c of cells) {
      const ref = c[1] + c[2];
      const type = (c[3].match(/t="([^"]+)"/) || [])[1] || '';
      const v = c[5] || '';
      const val = type === 's' ? shared[Number(v)] : v;
      const f = c[4] ? ' [F]' : '';
      console.log(ref + ': ' + val + f);
    }
  }
}
