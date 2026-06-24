---
name: wordex
description: >-
  Wordex report editor and PDF pipeline (WORDEX.html, paginated HTML, PDF,
  ObterPDF iframe API). Use when editing Wordex, fixing layout/watermark/
  pagination/charts/tables/paragraphs, generating reports from JSON, or
  integrating via iframe or backend PDF.
---

# Wordex

Cliente-side report designer: template no editor → JSON → HTML paginado → PDF.

## Arquivos principais

| Arquivo | Papel |
|---------|--------|
| `WORDEX.html` | Editor (`WordexEditor`), toolbar, geração de relatório, `ObterPDF` |
| `wordex-paged.html` | Template embutido (`window.WORDEX_PAGED_TEMPLATE`): paginação, toolbar paginado |
| `wordex-pdf.html` | Template embutido (`window.WORDEX_PDF_TEMPLATE`): captura jsPDF/html2canvas |
| `helpIFRAME.txt` | API `ObterPDF(json)` para iframe |
| `helpCHROME.txt` | PDF no backend (Chrome CLI / Puppeteer) |
| `helpBACKEND.txt` | Windows Service: template + JSON → headless → PDF |

Não editar `backup/`, `_zip_extract/`, `joao.html`, `teste.html` salvo pedido explícito.

## Modelo de documento

```text
wordexDocument = {
  page, watermark,
  headerTemplate: [ blocos ],
  bodyFlow:       [ blocos ],
  footerTemplate: [ blocos ]
}
```

Cada bloco é `{ id, type: "paragraph", html, style }`. Objetos (gráfico, tabela, imagem, textbox) vivem **dentro** do `html` do parágrafo.

Regiões DOM: `.header-template`, `.body-flow`, `.footer-template` → wrappers `.paragraph-block` → `.paragraph`.

Mutações: alterar `wordexDocument[region]` → `render()`. Antes de salvar/gerar: `syncCurrentEditingState()`.

## Pipeline de relatório

1. `buildGeneratedReportHtml(data)` monta HTML com placeholders `@@...@@`
2. `wordex-paged.html` pagina (`startPagex`, `.pagex-page`, classe `wordex-pagex-ready`)
3. `wordex-pdf.html` captura páginas e gera PDF

Config JSON em `#wordex-pagex-config` / `#wordex-pdf-config`.

## API iframe

`window.ObterPDF(json)` → `Promise<dataUri>` (`data:application/pdf;base64,...`).

Orquestra iframes ocultos: paginado → `wordexBuildPdfHtml` → PDF com `returnBase64: true`.

Detalhes: [helpIFRAME.txt](../../helpIFRAME.txt) na raiz do projeto.

## Gestos do editor

Princípio: **toolbar = ações frequentes e contextuais** (grupos aparecem conforme seleção). Gestos avançados ficam no mouse/teclado — não sobrecarregar a barra.

### Seleção e edição

| Gestão | Comportamento |
|--------|----------------|
| Tabela — clique normal | Edita célula (modo célula) |
| Tabela — **Ctrl+clique** | Cicla **tabela → linha → coluna → célula → tabela** |
| Tabela — **Shift+clique** | Seleção em bloco de células |
| Objeto (gráfico, imagem, textbox) — clique normal | Edita conteúdo interno (textbox) ou foco normal |
| Objeto — **Ctrl+clique** | Seleciona objeto (borda laranja, cantoneiras, arrasto livre) |
| Textbox — **clique normal** | Edita texto interno; grupo **Texto** da toolbar visível |
| Textbox — **Ctrl+clique** | Seleciona a caixa como objeto; grupo **Texto** oculto; **Alinhar** / **Borda** / **Margem** |
| Textbox — **duplo-clique** | Só edita/seleciona texto (sem diálogo legado de alinhamento) |
| Parágrafo — **Ctrl+clique** no bloco | Modo movimentação (`.move-selected`); **arrastar** reordena na região; clique simples cancela |

### Arrastar objeto no parágrafo

Requisito: objeto selecionado com **Ctrl+clique** (tabela exige modo **tabela** inteira).

| Gestão | Efeito |
|--------|--------|
| **Arrastar** (sem Ctrl) | Move objeto; vertical altera margem de **cima** (`blockY` / `top`) |
| **Ctrl + arrastar** na vertical | Objeto mantém altura; só cresce/diminui margem de **baixo** (`blockBottomGap`) |
| **Horizontal** (com ou sem Ctrl) | Move esquerda/direita (`blockX`) |

Conversão automática para `data-align="free"` + `position:absolute` ao iniciar arrasto (`prepareObjectForFreeDrag`, `applyObjectFreePosition`).

### Teclado (objeto selecionado no parágrafo)

| Gestão | Efeito |
|--------|--------|
| **Ctrl + ←/→** | Move horizontalmente (passo 24px; **Ctrl+Shift** = 80px) |
| **Ctrl + ↑/↓** | Move verticalmente (margem de cima) |
| **Ctrl + Shift + ↑/↓** | Ajusta só margem de **baixo** (equivalente ao Ctrl+arrastar vertical) |

Setas **sem Ctrl** = navegação normal do texto; não mover objeto.

### Tabela — arrasto e redimensionamento

- **Mover tabela inteira**: Ctrl+clique até modo **tabela** → arrastar (não entra em edição de célula no mousedown).
- **Redimensionar coluna**: Ctrl+clique até modo **coluna** → alças horizontais.
- **Redimensionar linha**: Ctrl+clique até modo **linha** → alças verticais.
- **Modo célula**: clique normal ou ciclo Ctrl+clique; arrasto livre da tabela **não** se aplica.

### Feedback visual no editor

- **Parágrafo**: contorno tracejado azul quando focado/selecionado — **permanece visível** mesmo com objeto selecionado dentro dele.
- **Regiões Header/Body/Footer**: no editor, **sem** contorno azul sólido nem bordas tracejadas de guia; só etiquetas “Header” / “Body” / “Footer”.
- **Objeto selecionado**: borda laranja (`.selected-object`) + overlay de resize (`.wordex-resize-overlay`).

Reordenar parágrafo: `moveBlockToIndex` — alterar `wordexDocument[region]`, **não** só DOM.

Código: `handleTableOrObjectClick`, `cycleTableComponentSelection`, `startInlineObjectMouseDragCandidate`, `handleInlineObjectMouseDragMove`, `handleFreeObjectKeyboard`.

### Textbox

Inserir só em **parágrafo de objeto** (imagem, gráfico ou tabela) — botão **▤ Textbox** (grupo Inserir).

| Contexto | Toolbar **Texto** | Combo **Vertical** (grupo Alinhar) |
|----------|-------------------|-------------------------------------|
| Caret **dentro** do textbox (edição) | Visível | Ativa — alinha o **texto dentro** da caixa |
| Textbox **Ctrl+clique** (objeto) | Oculta | Ativa enquanto o textbox é o alvo |
| Caret no **parágrafo** (fora do textbox) | Conforme parágrafo | **Desativada** — não afeta textbox “fantasma” |

**Alinhamento vertical do texto** (não move a caixa no parágrafo): combo `#textboxVerticalAlign` — **Nenhum**, **Topo**, **Centro**, **Base**. Persiste em `data-text-vertical-align`; aplica `display:flex` + `justify-content` via `syncTextboxContentVerticalAlignStyles`. Caixa precisa de **altura** suficiente para o efeito ser visível.

**Alinhamento horizontal do texto** (dentro da caixa): botões esquerda/centro/direita/justificar do grupo **Alinhar** com caret no textbox (`applyTextboxTextAlign`).

**Macros em parágrafo de objeto:** texto e macros (`@PageNumber`, campos JSON) **só dentro do textbox**, nunca diretamente no parágrafo. Caret no textbox ou combo **Campo** com alvo guardado (`getTextboxForMacroInsertion`, `prepareMacroInsertionTarget`). Sem JSON: macros de sistema (`@PageNumber`, `@Today`, …) na combo **Campo**.

**Caso de uso típico (header):** logotipo à direita + títulos em textbox(s) no mesmo parágrafo; alinhar verticalmente o texto das caixas com **Centro** ou **Base** para alinhar com o logo.

Código: `getTargetTextboxForVerticalAlign`, `applyTextboxContentVerticalAlign`, `getTextboxForMacroInsertion`, `insertMacroTextInEditableContainer`, `isTextToolbarContext`, `isParagraphSelectedForMacroContext`.

## Regras de layout (não reverter sem pedido)

- **Gráfico**: borda obrigatória (1px solid preta padrão); usuário pode alterar estilo, não remover (`applyTableBorderPreset('none')` bloqueado)
- **Marca d'água**: `.pagex-watermark` z-index 2, por cima do conteúdo; tabela/células transparentes no paginado/PDF
- **Paginação**: altura real dos blocos (`getBlockLayoutBounds`, `tightenMainObjectBlockHeights`); não inflar `min-height` durante paginação
- **PDF**: `printBackground: true`; `dynamic-image-box` / `object-block` com overflow visível na captura

Não reintroduzir `fitChartBoxHeightToContent` (quebrou resize pelas cantoneiras).

## Conceito de domínio (tabelas de dados)

Linha detalhe = profundidade fixa no datasource; grupos = ancestrais na árvore. Especificar só a profundidade do detalhe; o resto se deduz.

## Onde buscar no código

| Tarefa | `WORDEX.html` |
|--------|----------------|
| Bordas | `applyTableBorderPreset`, `getCurrentBoxBorderTarget`, `ensureChartEssentialBorder` |
| Parágrafo mover | `selectParagraphBlockForMove`, `paragraphBlockReorderDrag` |
| Tabela seleção | `cycleTableComponentSelection`, `handleTableOrObjectClick` |
| Arrasto de objeto | `startInlineObjectMouseDragCandidate`, `handleInlineObjectMouseDragMove`, `applyObjectFreePosition`, `prepareObjectForFreeDrag` |
| Teclado objeto livre | `handleFreeObjectKeyboard` |
| Textbox — alinhamento vertical do texto | `applyTextboxContentVerticalAlign`, `syncTextboxContentVerticalAlignStyles`, `#textboxVerticalAlign` |
| Textbox — macros | `getTextboxForMacroInsertion`, `insertMacroTextInEditableContainer`, `isParagraphSelectedForMacroContext` |
| Toolbar contextual | `isTextToolbarContext`, `updateToolbarContext`, `getToolbarContextState` |
| Geração HTML | `buildGeneratedReportHtml`, `buildGeneratedDocumentClone` |
| PDF iframe | `obterPdfFromJson`, `ObterPDF` |
| Paginação | `wordex-paged.html`: `paginateBlocks`, `startPagex`, `buildPdfViewerHtml` |
| PDF captura | `wordex-pdf.html`: `generatePdfFromPagexPages`, `capturePagexPage` |

## Princípios ao editar

- Diff mínimo; seguir estilo existente (vanilla JS, sem framework)
- Responder em português
- Não commitar sem pedido
- Servir por HTTP para imagens externas e testes de PDF (`file://` limita CORS)
- Testar fluxo: editor → Montar HTML → PDF e, se iframe, `ObterPDF`

## Referência estendida

Mapa de funções, placeholders e armadilhas: [reference.md](reference.md)
