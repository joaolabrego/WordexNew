const fs = require("fs");
const CFB = require("cfb");
const AdmZip = require("adm-zip");

// copy compress/decompress from _inject_vba.js
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

const zip = new AdmZip("wordex.backup.xlsm");
const cfb = CFB.read(zip.getEntry("xl/vbaProject.bin").getData());
const idx = cfb.FullPaths.indexOf("Root Entry/VBA/Wordex");
const raw = Buffer.from(cfb.FileIndex[idx].content);
const plain = decompressModule(raw);
console.log("plain len", plain.length);
console.log(plain.toString("latin1").slice(0, 200));
