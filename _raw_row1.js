const fs = require('fs');
for (const file of ['sheet5.xml', 'sheet6.xml', 'sheet7.xml']) {
  const xml = fs.readFileSync(`D:/WordexNew/_xlsm_live/xl/worksheets/${file}`, 'utf8');
  const row1 = xml.match(/<row r="1"[^>]*>([\s\S]*?)<\/row>/);
  console.log('\n' + file);
  if (!row1) continue;
  for (const c of row1[1].matchAll(/<c r="([^"]+)"([^>]*)>([\s\S]*?)<\/c>/g)) {
    console.log(c[1], c[2], c[3].replace(/\s+/g, ' ').slice(0, 120));
  }
}
