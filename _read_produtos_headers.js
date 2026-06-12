const AdmZip = require("adm-zip");
const z = new AdmZip("d:/WordexNew/wordex.xlsm");
const xml = z.readAsText("xl/worksheets/sheet3.xml", "utf8");
const s = z.readAsText("xl/sharedStrings.xml", "utf8");
const strings = [...s.matchAll(/<t>([^<]*)<\/t>/g)].map(m => m[1]);

for (const col of ["A", "B", "C", "D", "E", "F", "G"]) {
  const re = new RegExp(`<c r="${col}1"[^>]*t="s"><v>(\\d+)</v>`);
  const m = xml.match(re);
  console.log(col, m ? strings[Number(m[1])] : "(missing)");
}
