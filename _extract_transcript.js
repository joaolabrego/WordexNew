const fs = require("fs");
const path =
  "C:/Users/joaol/.cursor/projects/d-WordexNew/agent-transcripts/8a38f22b-e626-4de7-931f-bb1b5eaaaab2/8a38f22b-e626-4de7-931f-bb1b5eaaaab2.jsonl";
const lines = fs.readFileSync(path, "utf8").split(/\n/);

for (const line of lines) {
  if (!line.includes("Wordex.bas")) continue;
  if (!line.includes('"Write"')) continue;
  try {
    const j = JSON.parse(line);
    for (const c of j.message.content) {
      if (
        c.name === "Write" &&
        c.input &&
        c.input.path &&
        c.input.path.toLowerCase().includes("wordex.bas")
      ) {
        fs.writeFileSync("D:/WordexNew/_from_transcript_Wordex.bas", c.input.contents);
        console.log("written", c.input.contents.length);
        process.exit(0);
      }
    }
  } catch (e) {
    // ignore
  }
}

console.log("not found");
