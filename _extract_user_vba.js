const fs = require("fs");
const line = fs
  .readFileSync(
    "C:/Users/joaol/.cursor/projects/d-WordexNew/agent-transcripts/9e193440-c9a8-4459-bf6a-1122db6494cf/9e193440-c9a8-4459-bf6a-1122db6494cf.jsonl",
    "utf8"
  )
  .split("\n")[0];
const j = JSON.parse(line);
const text = j.message.content[0].text;
const start = text.indexOf("Option Explicit");
const end = text.lastIndexOf("End Function") + "End Function".length;
if (start < 0) {
  console.log("not found", start);
  process.exit(1);
}
const code = text.slice(start, end + 2000);
fs.writeFileSync("D:/WordexNew/_user_pasted_totais.bas", code);
console.log("written", code.length);
