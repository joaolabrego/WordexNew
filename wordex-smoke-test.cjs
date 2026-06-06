/**
 * Smoke test Wordex — fluxos críticos para apresentação (Node + jsdom).
 * Executar: node wordex-smoke-test.cjs
 */
const fs = require("fs");
const path = require("path");
const { JSDOM } = require("jsdom");

const ROOT = __dirname;
const results = [];

function pass(name) {
  results.push({ name, ok: true });
  console.log("  OK  " + name);
}

function fail(name, error) {
  const message = error instanceof Error ? error.message : String(error);
  results.push({ name, ok: false, message });
  console.log("  FAIL " + name + ": " + message);
}

function assert(name, condition, detail) {
  if (condition) {
    pass(name);
  } else {
    fail(name, detail || "assertion failed");
  }
}

function loadWordex() {
  const vm = require("vm");
  const htmlFile = fs.readFileSync(path.join(ROOT, "WORDEX.html"), "utf8");
  const scriptStart = htmlFile.indexOf("<script>\n    class WordexEditor");
  const scriptEnd = htmlFile.indexOf('const wordex = new WordexEditor("document");');

  if (scriptStart < 0 || scriptEnd < 0) {
    throw new Error("Não foi possível extrair o script do WordexEditor.");
  }

  const editorScript = htmlFile.slice(scriptStart + 8, scriptEnd) +
    'window.wordex = new WordexEditor("document");';

  const dom = new JSDOM("<!DOCTYPE html><html><body><div id=\"document\"></div></body></html>", {
    url: "file://" + path.join(ROOT, "WORDEX.html").replace(/\\/g, "/"),
    runScripts: "outside-only"
  });

  const { window } = dom;
  const context = dom.getInternalVMContext();

  window.WORDEX_PAGED_TEMPLATE = fs.readFileSync(path.join(ROOT, "wordex-paged.html"), "utf8");
  window.WORDEX_PDF_TEMPLATE = fs.readFileSync(path.join(ROOT, "wordex-pdf.html"), "utf8");
  window.alert = () => {};
  window.confirm = () => true;
  window.prompt = (msg, def) => {
    if (String(msg || "").includes("Macro")) return "TesteMacro";
    if (String(msg || "").includes("Texto inicial")) return "Titulo";
    if (String(msg || "").includes("Largura")) return "40mm";
    if (String(msg || "").includes("Altura")) return "10mm";
    if (String(msg || "").includes("Tipo do gráfico")) return "bar";
    if (String(msg || "").includes("linhas")) return "2";
    if (String(msg || "").includes("colunas")) return "2";
    if (String(msg || "").includes("Datasource")) return "ROOT";
    return def ?? null;
  };
  window.requestAnimationFrame = callback => {
    callback();
    return 1;
  };

  vm.runInContext(editorScript, context);

  return { dom, wordex: window.wordex, window };
}

function stubLayout(window) {
  window.Element.prototype.getBoundingClientRect = function () {
    const width = Number.parseFloat(this.style?.width) || 200;
    const height = Number.parseFloat(this.style?.height) || 100;
    return {
      x: 0,
      y: 0,
      top: 0,
      left: 0,
      right: width,
      bottom: height,
      width,
      height
    };
  };

  Object.defineProperty(window.HTMLElement.prototype, "offsetHeight", {
    configurable: true,
    get() {
      return Number.parseFloat(this.style?.height) || 100;
    }
  });

  Object.defineProperty(window.HTMLElement.prototype, "offsetWidth", {
    configurable: true,
    get() {
      return Number.parseFloat(this.style?.width) || 200;
    }
  });

  Object.defineProperty(window.HTMLElement.prototype, "clientWidth", {
    configurable: true,
    get() {
      return 600;
    }
  });

  Object.defineProperty(window.HTMLElement.prototype, "clientHeight", {
    configurable: true,
    get() {
      const min = Number.parseFloat(this.style?.minHeight);
      return Number.isFinite(min) && min > 0 ? min : 120;
    }
  });
}

console.log("\nWordex smoke test\n");

try {
  const script = fs.readFileSync(path.join(ROOT, "WORDEX.html"), "utf8");
  const start = script.indexOf("<script>");
  const end = script.lastIndexOf("</script>");
  new Function(script.slice(start + 8, end));
  pass("Sintaxe JS do WORDEX.html");
} catch (error) {
  fail("Sintaxe JS do WORDEX.html", error);
}

let wordex;
let win;

try {
  const loaded = loadWordex();
  wordex = loaded.wordex;
  win = loaded.window;
  stubLayout(win);
  assert("WordexEditor instancia", !!wordex && typeof wordex.render === "function");
} catch (error) {
  fail("Carregar WordexEditor no jsdom", error);
  process.exit(1);
}

if (wordex) {
  try {
    wordex.render();
    pass("render() inicial");
  } catch (error) {
    fail("render() inicial", error);
  }

  try {
    const blocksBefore = wordex.wordexDocument.bodyFlow.length;
    wordex.selected.region = "bodyFlow";
    wordex.selectedBlockId = null;
    wordex.selectedParagraphId = null;
    wordex.selectedObject = null;

    const chart = wordex.createChartElement("Vendas", "bar", "left", "100mm", "60mm");
    wordex.insertMainObjectInNewParagraph(chart);

    const blocksAfter = wordex.wordexDocument.bodyFlow.length;
    assert("Objeto principal cria parágrafo novo", blocksAfter === blocksBefore + 1);

    const chartBlockId = wordex.selectedBlockId;
    const paragraph = win.document.querySelector(`.paragraph[data-block-id="${chartBlockId}"]`);
    const chartNode = paragraph?.querySelector(".wordex-chart-box");
    assert("Parágrafo novo contém o gráfico", !!chartNode);
    assert(
      "Gráfico usa posição absoluta livre",
      chartNode?.style?.position === "absolute"
    );

    wordex.selectObject(chartNode);
    const blocksBeforeImage = wordex.wordexDocument.bodyFlow.length;
    const image = wordex.createDynamicImageElement("Logo", "left", "80mm", "45mm");
    wordex.insertMainObjectInNewParagraph(image);

    assert(
      "Segundo objeto cria outro parágrafo (não empilha no mesmo)",
      wordex.wordexDocument.bodyFlow.length === blocksBeforeImage + 1
    );

    assert(
      "Parágrafo do gráfico continua com um só objeto principal",
      paragraph?.querySelectorAll(".dynamic-image-box, .wordex-table, .wordex-chart-box").length === 1
    );
  } catch (error) {
    fail("Inserção de objetos principais", error);
  }

  try {
    wordex.selected.region = "bodyFlow";
    wordex.selectedBlockId = null;
    wordex.selectedParagraphId = null;
    wordex.selectedObject = null;

    let alertMessage = "";
    win.alert = msg => {
      alertMessage = String(msg);
    };

    const textbox = wordex.createTextboxElement("Titulo", "free", "40mm", "10mm");
    wordex.insertTextboxIntoObjectParagraph(textbox);
    assert(
      "Textbox recusado sem parágrafo de objeto seleccionado",
      alertMessage.includes("parágrafo de objeto")
    );

    const chart = win.document.querySelector(".body-flow .wordex-chart-box");
    wordex.selectObject(chart);
    alertMessage = "";
    const textbox2 = wordex.createTextboxElement("Titulo grafico", "free", "40mm", "10mm");
    wordex.insertTextboxIntoObjectParagraph(textbox2);

    const objectParagraph = chart.closest(".paragraph");
    const textboxes = objectParagraph?.querySelectorAll(".wordex-textbox") || [];
    assert("Textbox inserido no parágrafo do gráfico", textboxes.length >= 1);
    assert(
      "Textbox com posição absoluta",
      textboxes[textboxes.length - 1]?.style?.position === "absolute"
    );
    assert(
      "Textbox z-index acima do gráfico",
      Number.parseInt(textboxes[textboxes.length - 1]?.style?.zIndex, 10) >= 2
    );
  } catch (error) {
    fail("Inserção de textbox", error);
  }

  try {
    const paragraph = win.document.querySelector(".body-flow .wordex-chart-box")?.closest(".paragraph");
    const textbox = paragraph?.querySelector(".wordex-textbox");
    const chart = paragraph?.querySelector(".wordex-chart-box");

    if (paragraph && textbox && chart) {
      wordex.applyObjectFreePosition(textbox, paragraph, 80, 12);
      assert("Textbox move para left/top independentes", textbox.style.left === "80px" && textbox.style.top === "12px");
      assert("Parágrafo cresce com minHeight", Number.parseFloat(paragraph.style.minHeight) >= 12 + 10);
    } else {
      fail("Movimentação livre do textbox", "parágrafo/gráfico/textbox não encontrados");
    }
  } catch (error) {
    fail("Movimentação livre do textbox", error);
  }

  try {
    const sampleData = JSON.parse(
      fs.readFileSync(path.join(ROOT, "WORDEX.json"), "utf8").replace(/^\uFEFF/, "")
    );
    const html = wordex.buildGeneratedReportHtml(Array.isArray(sampleData) ? sampleData[0] : sampleData);
    assert("buildGeneratedReportHtml produz HTML", typeof html === "string" && html.length > 500);
    assert("HTML gerado inclui body-flow ou pagex", /body-flow|pagex-page|wordex-generated-report/i.test(html));
  } catch (error) {
    fail("Montagem HTML", error);
  }

  try {
    const paged = fs.readFileSync(path.join(ROOT, "wordex-paged.html"), "utf8");
    assert("wordex-paged.html define template", paged.includes("WORDEX_PAGED_TEMPLATE"));
    assert("wordex-paged.html tem pagex-page", paged.includes("pagex-page"));
    assert("wordex-paged.html função overflows ou paginação", /function overflows|function paginate|splitParagraph/i.test(paged));
  } catch (error) {
    fail("wordex-paged.html", error);
  }

  try {
    const pdf = fs.readFileSync(path.join(ROOT, "wordex-pdf.html"), "utf8");
    assert("wordex-pdf.html define template", pdf.includes("WORDEX_PDF") || pdf.includes("pdf"));
    assert("wordex-pdf.html usa html2canvas ou jsPDF", /html2canvas|jspdf/i.test(pdf));
  } catch (error) {
    fail("wordex-pdf.html", error);
  }
}

const failed = results.filter(item => !item.ok);
console.log("\n--- Resumo ---");
console.log("Total: " + results.length + " | OK: " + (results.length - failed.length) + " | Falhas: " + failed.length);

if (failed.length) {
  console.log("\nFalhas:");
  failed.forEach(item => console.log(" - " + item.name + ": " + item.message));
  process.exit(1);
}

console.log("\nSmoke test concluído com sucesso.\n");
