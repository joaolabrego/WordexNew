const fs = require("fs");
const path = require("path");
const CFB = require("cfb");
const AdmZip = require("adm-zip");

const ROOT = __dirname;
const XLSM = path.join(ROOT, "wordex.xlsm");

function unpackHeader(bytes) {
  const intHeader = bytes.readUInt16LE(0);
  const compressed = (intHeader & 0x8000) >> 15;
  const length = (intHeader & 0x0fff) + 3;
  return { compressed, length };
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

function decompressModule(raw) {
  const buf = Buffer.from(raw);
  if (buf.length < 1 || buf[0] !== 0x01) return null;
  let payload = buf.subarray(1);
  const parts = [];
  while (payload.length >= 2) {
    const { compressed, length } = unpackHeader(payload);
    if (length > payload.length) break;
    const body = payload.subarray(2, length);
    parts.push(compressed ? decompressVba(body) : body);
    payload = payload.subarray(length);
  }
  return Buffer.concat(parts);
}

function ceilLog2(n) {
  let i = 4;
  while (2 ** i < n) i++;
  return i;
}

function copyTokenHelp(difference) {
  const bitCount = ceilLog2(difference);
  return { bitCount, maxLength: (0xffff << bitCount) + 3 };
}

function packCopyToken(length, offset, help) {
  return ((offset - 1) << (16 - help.bitCount)) | (length - 3);
}

function compressChunk(data) {
  const active = Buffer.from(data);
  let uncompressed = Buffer.from(data);
  let compressed = Buffer.alloc(0);

  function matching() {
    let bestLength = 0;
    let bestCandidate = 0;
    for (let candidate = active.length - uncompressed.length - 1; candidate >= 0; candidate--) {
      let c = candidate;
      let d = active.length - uncompressed.length;
      let length = 0;
      while (d < active.length && active[d] === active[c]) {
        c++;
        d++;
        length++;
      }
      if (length > bestLength) {
        bestLength = length;
        bestCandidate = candidate;
      }
    }
    if (bestLength >= 3) {
      const help = copyTokenHelp(active.length - uncompressed.length);
      return {
        offset: active.length - uncompressed.length - bestCandidate,
        length: Math.min(help.maxLength, bestLength),
      };
    }
    return { offset: 0, length: 0 };
  }

  function compressToken() {
    const { offset, length } = matching();
    if (offset > 0) {
      const help = copyTokenHelp(active.length - uncompressed.length);
      const token = packCopyToken(length, offset, help);
      uncompressed = uncompressed.subarray(length);
      return { packed: Buffer.from([token & 0xff, (token >> 8) & 0xff]), flag: 1 };
    }
    const b = uncompressed[0];
    uncompressed = uncompressed.subarray(1);
    return { packed: Buffer.from([b]), flag: 0 };
  }

  while (uncompressed.length > 0) {
    let tokenFlag = 0;
    const tokens = [];
    for (let i = 0; i < 8 && uncompressed.length > 0; i++) {
      const { packed, flag } = compressToken();
      tokenFlag |= flag << i;
      tokens.push(packed);
    }
    compressed = Buffer.concat([compressed, Buffer.from([tokenFlag]), ...tokens]);
  }

  let chunkSizeMinusThree = compressed.length - 1;
  let headerValue = 0xb000;
  let body = compressed;
  if (chunkSizeMinusThree > 4095) {
    body = Buffer.concat([data, Buffer.alloc(Math.max(0, 4096 - data.length))]);
    headerValue = 0x3000;
    chunkSizeMinusThree = 4095;
  }
  const header = headerValue | chunkSizeMinusThree;
  return Buffer.concat([Buffer.from([header & 0xff, (header >> 8) & 0xff]), body]);
}

function compressVba(data) {
  const out = [Buffer.from([0x01])];
  for (let i = 0; i < data.length; i += 4096) {
    out.push(compressChunk(data.subarray(i, i + 4096)));
  }
  return Buffer.concat(out);
}

function readModuleSource(basPath) {
  let text = fs.readFileSync(basPath, "utf8");
  if (!text.startsWith("Attribute VB_Name")) {
    const name = path.basename(basPath, ".bas");
    text = `Attribute VB_Name = "${name}"\r\n` + text;
  }
  if (!text.endsWith("\r\n")) text += "\r\n";
  return Buffer.from(text, "latin1");
}

function patchVbaProject(binBuffer, modules) {
  const cfb = CFB.read(binBuffer);
  for (const { streamName, basPath } of modules) {
    const fullPath = `Root Entry/VBA/${streamName}`;
    const idx = cfb.FullPaths.indexOf(fullPath);
    if (idx < 0) throw new Error(`Stream nao encontrado: ${streamName}`);
    const source = readModuleSource(basPath);
    cfb.FileIndex[idx].content = compressVba(source);
    console.log(`OK ${streamName}: ${source.length} bytes`);
  }
  return CFB.write(cfb);
}

// round-trip test
const zip0 = new AdmZip(XLSM);
const cfb0 = CFB.read(zip0.getEntry("xl/vbaProject.bin").getData());
const idx0 = cfb0.FullPaths.indexOf("Root Entry/VBA/Wordex");
const original = decompressModule(cfb0.FileIndex[idx0].content);
if (!original || !original.toString("latin1").includes("ObterRegistros")) {
  throw new Error("Falha ao descomprimir modulo original.");
}
console.log("Round-trip test OK, original", original.length, "bytes");

fs.copyFileSync(XLSM, path.join(ROOT, "wordex.backup.xlsm"));
const zip = new AdmZip(XLSM);
const patched = patchVbaProject(zip.getEntry("xl/vbaProject.bin").getData(), [
  { streamName: "Wordex", basPath: path.join(ROOT, "Wordex.bas") },
  { streamName: "WordexConsulta", basPath: path.join(ROOT, "WordexConsulta.bas") },
]);
zip.updateFile("xl/vbaProject.bin", patched);
zip.writeZip(XLSM);
console.log("wordex.xlsm atualizado.");
