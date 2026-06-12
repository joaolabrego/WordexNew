# Wordex — referência

## Carregamento de templates

`WORDEX.html` inclui:

```html
<script src="wordex-paged.html"></script>
<script src="wordex-pdf.html"></script>
```

Cada arquivo define `window.WORDEX_PAGED_TEMPLATE` / `window.WORDEX_PDF_TEMPLATE` (string HTML com `@@PLACEHOLDERS@@`).

`WordexEditor.loadPagedTemplate()` / `loadPdfTemplate()` injetam helpers compartilhados no marcador `/*WORDEX_TEMPLATE_HELPERS*/` via `getTemplateHelpersScript()` (placeholders, nomes de arquivo, utilitários de imagem).

Helpers compartilhados: `applyTemplatePlaceholders`, `sanitizeHtmlFileName`, `sanitizeFileName`, `fetchImageAsDataUrl`, `loadImageAsDataUrl`, `omitUnembeddableImage`, etc.

## Placeholders principais (paginado)

| Placeholder | Conteúdo |
|-------------|----------|
| `@@WORDEX_SOURCE_HTML@@` | Clone do documento gerado |
| `@@WORDEX_PAGEX_CONFIG@@` | JSON: header/footer templates, watermark, pageConfig, imageDataUrlMap |
| `@@WORDEX_PAGE_WIDTH@@` etc. | Dimensões da página |

## Placeholders principais (PDF)

| Placeholder | Conteúdo |
|-------------|----------|
| `@@WORDEX_PAGEX_HTML@@` | HTML das `.pagex-page` já inline |
| `@@WORDEX_PDF_CONFIG@@` | JSON: fileName, title, imageDataUrlMap, returnBase64 |
| `@@WORDEX_PAGEX_STYLE@@` | CSS da paginação |
| `@@WORDEX_REPORT_STYLES@@` | CSS do relatório |

## Classes CSS críticas

- `.pagex-page` — uma página A4 no relatório
- `.pagex-watermark` — marca d'água (overlay, z-index 2)
- `.wordex-chart-box` — gráfico; borda essencial
- `.wordex-table` — tabela; modos `data-wordex-selection-mode`
- `.paragraph-block.move-selected` — parágrafo em modo reordenar
- `.wordex-paragraph-drop-indicator` — linha de drop no arrasto
- `.selected-object` — borda laranja de objeto selecionado
- `.wordex-resize-overlay` — cantoneiras de redimensionamento (fixo no body)

### Editor vs relatório gerado

No editor (`body:not(.wordex-generated-report)`):

- Regiões **sem** `outline` de seleção (`.region.selected` / `:hover` desativados); etiquetas Header/Body/Footer mantidas.
- `body-flow` / header / footer **sem** bordas tracejadas de guia.
- `.paragraph` usa `display: flow-root`; `::after` clearfix desativado no editor.
- Contorno tracejado do `.paragraph` **não** some quando há `.selected-object` dentro.

No `@media print` e relatório paginado/PDF: outlines de edição removidos.

## JSON do relatório

Carregado em `reportData`. Imagens externas: prefetch para data URL (`prefetchReportImageDataUrls`). Base64 no JSON evita omissão no PDF.

## ObterPDF — fluxo interno

1. `normalizeReportJsonInput(json)`
2. `buildGeneratedReportHtml(data)` → iframe oculto
3. Espera `wordex-pagex-ready` no body do iframe
4. `wordexBuildPdfHtml({ returnBase64: true })`
5. Segundo iframe com template PDF
6. Espera `wordexPdfExportResult` com `dataUri` / `base64`

`postMessage`: tipo `wordex-obter-pdf` → resposta `wordex-obter-pdf-result`.

## Backend PDF

Preferir Puppeteer (`page.pdf`, `preferCSSPageSize`, `printBackground`) sobre Chrome CLI em Node. Ver `helpCHROME.txt`.

## Undo

`saveUndoStateIfChanged()` após mutações estruturais. Não amend commit hooks falhos.

## Armadilhas conhecidas

| Problema | Causa usual | Direção de fix |
|----------|-------------|----------------|
| Marca d'água oculta | Fundo opaco em tabela | `background: transparent` no paginado/PDF |
| Espaço entre blocos | `min-height` inflado na paginação | Medir sem mutar; `tightenMainObjectBlockHeights` |
| Gráfico cortado no PDF | overflow hidden na captura | `overflow: visible` no clone |
| Folga interna do gráfico | Altura da caixa > SVG | Borda resolve visualmente; não encolher SVG auto |
| Imagens faltando no PDF | CORS / URL externa | Embutir base64 ou servir HTTP acessível ao gerador |
| Tabela não arrasta | Seleção limpa no mousedown da célula | `prepareTableCellForEditing` não deve correr em modo `table`; `defaultPrevented` não bloqueia `startInlineObjectMouseDragCandidate` |
| Linha-guia azul no editor | Contorno/borda de região Body | Remover outline `.region.selected` e bordas `body-flow` no editor (ver SKILL) |

## Arquivos auxiliares na raiz

- `WORDEX.json`, `wordex-correto.json` — exemplos de dados
- Scripts `_*.js` — utilitários pontuais; não são o app principal

## Versão / título

Título da página: `Wordex - MVP Contínuo v299 header/footer compacto`. Classe principal: `WordexEditor`.
