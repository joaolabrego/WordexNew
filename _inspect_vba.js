const CFB = require("cfb");
const AdmZip = require("adm-zip");

function decompressVba(compressed) {
  const out = [];
  let i = 0;
  while (i < compressed.length) {
    const flags = compressed[i++];
    for (let bit = 0; bit < 8; bit++) {
      if (i >= compressed.length) break;
      if (flags & (1 << bit)) out.push(compressed[i++]);
      else {
        if (i + 1 >= compressed.length) break;
        const ref = compressed[i] | (compressed[i + 1] << 8);
        i += 2;
        const length = (ref & 0xf) + 3;
        const offset = (ref >> 4) + 1;
        const start = out.length - offset;
        for (let j = 0; j < length; j++) out.push(out[start + j]);
      }
    }
  }
  return Buffer.from(out);
}

function decompressModule(raw) {
  const buf = Buffer.from(raw);
  if (buf[0] !== 0x01) return null;
  let payload = buf.subarray(1);
  const parts = [];
  while (payload.length >= 2) {
    const length = (payload.readUInt16LE(0) & 0x0fff) + 3;
    if (length > payload.length) break;
    parts.push(decompressVba(payload.subarray(2, length)));
    payload = payload.subarray(length);
  }
  return Buffer.concat(parts);
}

const zip = new AdmZip("d:/WordexNew/wordex.xlsm");
const cfb = CFB.read(zip.getEntry("xl/vbaProject.bin").getData());
for (const name of ["Wordex", "WordexConsulta"]) {
  const idx = cfb.FullPaths.indexOf(`Root Entry/VBA/${name}`);
  if (idx < 0) {
    console.log(name, "missing");
    continue;
  }
  const plain = decompressModule(Buffer.from(cfb.FileIndex[idx].content));
  const text = plain.toString("latin1");
  console.log("\n===", name, "===");
  for (const key of ["K_HEADERS_ROW", "K_TYPES_ROW", "K_DETAILS_ROW", "ObterNumeroColuna", "Wordex_ObterKindColuna"]) {
    const i = text.indexOf(key);
    console.log(key, i >= 0 ? text.slice(i, i + 120).replace(/\r/g, "") : "NOT FOUND");
  }
}
