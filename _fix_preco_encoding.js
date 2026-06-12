const fs = require("fs");
const AdmZip = require("adm-zip");

const xlsmPath = "d:/WordexNew/wordex.xlsm";
const zip = new AdmZip(xlsmPath);

let sheet2 = zip.readAsText("xl/worksheets/sheet2.xml", "utf8");
sheet2 = sheet2.replace(/PreÃ§o/g, "Preço");
zip.updateFile("xl/worksheets/sheet2.xml", Buffer.from(sheet2, "utf8"));

let sheet1 = zip.readAsText("xl/worksheets/sheet1.xml", "utf8");
sheet1 = sheet1.replace(/PreÃ§o/g, "Preço");
zip.updateFile("xl/worksheets/sheet1.xml", Buffer.from(sheet1, "utf8"));

zip.writeZip(xlsmPath);
console.log("Formulas corrigidas no xlsm.");
