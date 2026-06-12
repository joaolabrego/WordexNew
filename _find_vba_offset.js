const AdmZip = require("adm-zip");
const CFB = require("cfb");

function unpackHeader(bytes) {
  const intHeader = bytes.readUInt16LE(0);
  return { compressed: (intHeader & 0x8000) >> 15, length: (intHeader & 0x0fff) + 3 };
}

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

function tryDecompressContainer(buf) {
  if (buf.length < 3 || buf[0] !== 0x01) return null;
  let payload = buf.subarray(1);
  const parts = [];
  while (payload.length >= 2) {
    const { compressed, length } = unpackHeader(payload);
    if (length < 2 || length > payload.length) break;
    const body = payload.subarray(2, length);
    if (compressed) parts.push(decompressVba(body));
    else parts.push(body.subarray(0, Math.min(body.length, 4096)));
    payload = payload.subarray(length);
  }
  return Buffer.concat(parts).toString("latin1");
}

const zip = new AdmZip("wordex.xlsm");
const cfb = CFB.read(zip.getEntry("xl/vbaProject.bin").getData());
const raw = Buffer.from(cfb.FileIndex[cfb.FullPaths.indexOf("Root Entry/VBA/Wordex")].content);

for (let off = 0; off < Math.min(raw.length, 5000); off++) {
  const slice = raw.subarray(off);
  if (slice[0] !== 0x01) continue;
  const text = tryDecompressContainer(slice);
  if (text && text.includes("Attribute VB_Name")) {
    console.log("Found at offset", off);
    console.log(text.slice(0, 120));
    break;
  }
}
