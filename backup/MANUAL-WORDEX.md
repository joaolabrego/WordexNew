# Manual operacional — Wordex

> Documento vivo. Completar e corrigir conforme o sistema evoluir.  
> Última revisão: junho/2026.

---

## 1. O que é o Wordex (e o que não é)

| | **Wordex** | **Crudex** (outro sistema) |
|---|------------|----------------------------|
| Papel | Editor e gerador de **relatórios** (layout, gráficos, tabelas, PDF) | Gerador e executor de **sistemas de informação** |
| Entrada | Template visual + JSON de dados | Definição do sistema / dados de negócio |
| Saída | HTML paginado e PDF | Aplicação em execução |

O Wordex **não substitui** o Crudex. Em geral: o Crudex produz o sistema e os dados; o Wordex desenha o **documento final** (relatório) a partir de um template e de um JSON.

---

## 2. Arquivos principais

| Arquivo | Função |
|---------|--------|
| `WORDEX.html` | Editor principal (`WordexEditor`), barra de ferramentas, montagem de relatório, API `ObterPDF` |
| `wordex-paged.html` | Template de paginação (páginas A4, header/footer, marca d'água) |
| `wordex-pdf.html` | Captura das páginas e geração do PDF |
| `helpIFRAME.txt` | Integração via iframe (`ObterPDF`) |
| `helpCHROME.txt` | PDF no backend (Chrome / Puppeteer / `WordexChromePdf`) |
| `helpBACKEND.txt` | Windows Service: template + JSON → PDF (headless, sem paginar em C#) |
| `WORDEX.json`, `wordex-correto.json` | Exemplos de JSON de relatório |

**Como abrir:** servir a pasta por **HTTP** (não confiar só em `file://` — imagens externas e PDF podem falhar por CORS).

---

## 3. Fluxo operacional resumido

```text
1. Desenhar o template no editor (WORDEX.html)
      ↓
2. Salvar o template (HTML interno vira wordexDocument em memória / export)
      ↓
3. Carregar JSON de dados (📂) ou receber via integração
      ↓
4. Montar template HTML (botão HTML primário) — preenche macros e dados
      ↓
5. Paginar (automático no fluxo) → visualizar páginas
      ↓
6. Gerar PDF (na janela paginada ou via ObterPDF)
```

---

## 4. Estrutura do documento

O documento tem **três regiões**:

- **Cabeçalho** (`headerTemplate`) — repetido no topo de cada página
- **Corpo** (`bodyFlow`) — fluxo principal do relatório
- **Rodapé** (`footerTemplate`) — repetido no fim de cada página

Cada região é uma lista de **blocos**. Hoje o bloco usual é um **parágrafo** (`type: "paragraph"`), com HTML editável.

Objetos (gráfico, tabela, imagem, textbox) ficam **dentro do HTML do parágrafo**, não como blocos separados no modelo (salvo legado de `object-block`, em migração).

---

## 5. Barra de ferramentas — visão geral

### Inserir
- **¶ Parágrafo** — bloco de texto
- **📊 Histograma** — gráfico de barras (vertical ou horizontal)
- **▤ Textbox** — caixa de texto livre
- **▦ Tabela** — tabela Wordex
- **▧ Imagem (JSON)** — imagem ligada a campo do JSON (aparece com JSON carregado)
- **⚙ Configurar objeto** — dimensões, alinhamento etc. (com objeto selecionado)

### Página
- Formato (A4, etc.) e orientação (retrato / paisagem)

### Histograma (com gráfico selecionado — **Ctrl+clique** no gráfico)
- **Horizontal** — barras horizontais
- **Modo:** `normal` | `acumulado` | `misto`
- **Cor da série** — clique numa barra para escolher qual série colorir
- **Coleção / Histograma** (seletores de JSON) — ligar o gráfico a um histograma do Crudex (ver §8)

### Texto / Alinhar / Borda / Margem
- Formatação de texto, alinhamento de parágrafo ou objeto
- Bordas (tabela, célula, gráfico, imagem, textbox)
- Margem do **bloco** parágrafo (não confundir com gaps internos do gráfico — ver §8)

### Dados
- **📂 Carregar JSON**
- **🖼 Imagem local** — no objeto selecionado ou novo (Ctrl+clique limpa imagem)
- **💧 Marca d'água** — imagem de fundo (Ctrl+clique remove)
- **Coleção / Campo macro** — macros de texto, imagem ou histograma (com JSON carregado; contexto depende do objeto selecionado)

### Template
- **Montar template HTML** — gera relatório a partir do JSON
- **💾 Salvar template HTML** — exporta o template do editor

Na janela **paginada**: **Salvar HTML** e **Gerar PDF**.

---

## 6. Gestos do mouse e atalhos

### Parágrafo (bloco)
| Ação | Como |
|------|------|
| Editar texto | Clique no parágrafo |
| Modo mover / reordenar | **Ctrl+clique** no bloco (`.move-selected`) |
| Reordenar na região | Arrastar o bloco no modo mover |
| Cancelar modo mover | Clique simples fora |

### Objeto (gráfico, imagem, tabela, textbox)
| Ação | Como |
|------|------|
| Editar conteúdo (texto em textbox / célula) | Clique normal |
| Selecionar objeto (borda, resize, arrasto) | **Ctrl+clique** |
| Arrastar objeto no parágrafo | Arrastar (após seleção implícita ao arrastar, ou com objeto selecionado) |
| Mover com teclado | **Ctrl + setas** (Shift = passo maior) |

**Textbox:** só arrasta livremente depois de **Ctrl+clique** (clique normal edita o texto).

### Tabela
| Ação | Como |
|------|------|
| Editar célula | Clique na célula |
| Ciclar seleção tabela → linha → coluna → célula | **Ctrl+clique** |
| Seleção em bloco de células | **Shift+clique** |
| Arrastar tabela inteira | Modo seleção **tabela** + arrasto |

### Gráfico
| Ação | Como |
|------|------|
| Selecionar (borda, resize, toolbar) | **Ctrl+clique** |
| **Dados do histograma** (categorias, séries, valores) | **Duplo-clique** no gráfico |
| Configurar tamanho/alinhamento | **Ctrl+duplo-clique** |
| Cor da série | Com gráfico selecionado, clique numa barra |
| Legenda arrastável | Com gráfico selecionado, arrastar a legenda dentro da área |

---

## 7. Parágrafo com um objeto principal (gráfico, imagem ou tabela)

Quando o parágrafo contém **um único** objeto principal (gráfico, imagem ou tabela), o Wordex trata o parágrafo como **caixa de layout** para esse objeto.

### Por que existem gaps superior e inferior separados

Casos de uso típico: colocar **textboxes** (títulos, notas, percentuais explicativos) **acima** ou **abaixo** do gráfico, sem esmagar o desenho.

### Arrasto vertical — gaps independentes

| Gesto | O que altera |
|-------|----------------|
| **Arrastar** (sem Ctrl) | Distância **superior** do objeto à borda do parágrafo (`marginTop` / `blockY`) |
| **Ctrl + arrastar** (eixo vertical) | Distância **inferior** do objeto à borda do parágrafo (`marginBottom` / `blockBottomGap`) |
| Arrasto horizontal (com ou sem Ctrl) | Posição horizontal (`marginLeft` / `blockX`) |

A altura mínima do parágrafo acompanha: **gap superior + altura do objeto + gap inferior**. Outros objetos no mesmo parágrafo (ex.: textboxes) também entram no cálculo da altura.

### Teclado (parágrafo com objeto)
| Atalho | Efeito |
|--------|--------|
| Ctrl + ↑ / ↓ | Ajusta gap **superior** |
| Ctrl + Shift + ↑ / ↓ | Ajusta gap **inferior** |
| Ctrl + ← / → | Move horizontalmente |

Gap inferior padrão: **4 px** (se nunca foi ajustado). Valores ficam em `data-block-y` e `data-block-bottom-gap` no elemento do objeto e são persistidos no HTML do parágrafo.

### Fluxo sugerido: gráfico + textos auxiliares

1. Inserir parágrafo com histograma.
2. **Arrastar** o gráfico para abrir espaço **em cima**.
3. **Ctrl+arrastar** para abrir espaço **embaixo**.
4. Inserir **textbox** no mesmo parágrafo e posicionar nos vãos (Ctrl+clique no textbox para mover/redimensionar).
5. Montar HTML / PDF e conferir se a paginação respeitou a altura.

---

## 8. Histogramas (gráficos de barras)

### Tipos e modos
- **Vertical** (padrão) ou **Horizontal** (checkbox na toolbar)
- **normal** — barras lado a lado por categoria
- **acumulado** — barras empilhadas; rótulos em % por segmento
- **misto** — empilhamento com tratamento de valores negativos separado

### Regras visuais
- Gráfico tem **borda obrigatória** (1 px preta por padrão). O usuário pode mudar estilo/espessura, mas **não remover** completamente a borda.
- Rótulos de valor nos modos acumulado/misto são posicionados automaticamente (legenda percentual por categoria).
- Gráficos **normais** usam rótulos rotacionados no topo das barras (com detecção de colisão dentro da mesma categoria).

### Tipografia do gráfico (isolada do parágrafo)

A fonte, tamanho e cor do histograma afetam **somente** o gráfico (rótulos, legenda, categorias, eixos) — **não** o parágrafo que o contém nem os parágrafos seguintes.

| Situação | Comportamento esperado |
|----------|------------------------|
| Gráfico selecionado + mudar fonte na toolbar | Atualiza só o gráfico |
| Parágrafo de texto selecionado + mudar fonte | Atualiza só esse parágrafo |
| Inserir novo parágrafo (¶) ou Enter com gráfico ainda selecionado | Novo bloco usa a tipografia do **parágrafo hospedeiro** (ou Arial padrão), não a do gráfico |
| Selecionar parágrafo depois de editar o gráfico | Toolbar volta a mostrar a fonte **do parágrafo** |

**Por que:** o caso de uso típico é gráfico com fonte própria (ex.: Times nos rótulos) dentro de um relatório em Arial, com textboxes nos gaps (§7). Antes, a fonte do gráfico “vazava” para parágrafos novos porque a toolbar compartilhada refletia o gráfico e novos blocos liam esse valor.

**Persistência técnica:** tipografia do gráfico em `data-chart-font-family`, `data-chart-font-size`, `data-chart-font-color`, etc. — não em `style.fontFamily` no elemento do gráfico. Gráficos antigos com fonte no `style` inline são migrados automaticamente ao carregar.

**Se parágrafos já ficaram com fonte errada:** selecione cada bloco de texto, corrija a fonte na toolbar (com o **parágrafo** selecionado, não o gráfico), ou recrie os blocos.

### Dados — passo a passo (não esquecer amanhã)

1. **Duplo-clique** no histograma (ou **Ctrl+clique** para selecionar e depois duplo-clique).
2. No diálogo:
   - **Linhas** = categorias (Item 1, Item 2, …)
   - **Colunas** = séries (Preço, Quantidade, …)
   - **Células** = valores em texto formatado (`350,00`, `100`, …)
   - **+ / − Linha** e **+ / − Coluna** para ajustar a grade
3. **Confirmar** — grava e redesenha o gráfico.

**Inserir histograma novo:** cursor no parágrafo → **📊 Histograma** (grupo Inserir). **Ctrl+duplo-clique** para tamanho e alinhamento.

### Dados a partir do JSON (Crudex)

Com JSON carregado (📂) e histograma selecionado (**Ctrl+clique**):

| Onde está o gráfico | Seletor **Coleção** | Seletor **Histograma** |
|---------------------|---------------------|-------------------------|
| Parágrafo (fora de tabela) | `ROOT` ou histograma do ROOT (ex.: `TotaisProdutosGrafico`) | Com `ROOT`: escolher o campo histograma |
| Célula de tabela (header, detail ou footer) | **Fixo** na coleção da linha (ex.: `Clientes`) | Histogramas da linha (ex.: `TotaisProdutosGrafico`) |

Ao escolher o histograma, o Wordex **importa** os valores do JSON para o **mesmo grid** do ▥ (não fica um caminho separado). Depois pode abrir **▥** e editar — útil para **simulações**.

- Trocar a macro no seletor **reimporta** do JSON (sobrescreve edições manuais anteriores).
- No relatório/PDF, usa os dados do grid (inclui simulações salvas no template).
- Histograma numa tabela importa a amostra da primeira linha; se editar o grid, o mesmo gráfico editado repete em todas as linhas.

### Formato JSON do histograma (Crudex)

Wrapper com `"Kind": "histogram"` e `Items`. Cada item traz:

- Campo formatado: `"PreçoSum": { "Kind": "string", "Value": "350,00" }`
- Legenda da série: `"PreçoSum_Label": { "Kind": "string", "Value": "Preço" }`

Não é necessário campo numérico `*_Value` — o Wordex interpreta o texto formatado (`150,00`, `3`, etc.).

### Legenda
- Com **2+ séries**, aparece painel de legenda **arrastável** dentro do gráfico (§6).

---

## 9. Textboxes

- Caixa de texto editável, borda tracejada no editor.
- **Clique** — editar texto.
- **Ctrl+clique** — selecionar como objeto (mover, redimensionar, borda, margem).
- Úteis nos **gaps** acima/abaixo do gráfico (ver §7).

---

## 10. Tabelas

- Modos de seleção: célula, linha, coluna, tabela inteira (Ctrl+clique cicla).
- Linhas estruturais (cabeçalho, grupo, rodapé) e linhas **Free** — toolbar **Tabela**: **Ctrl+clique** numa célula cicla até **linha** ou **coluna**; **+** / **−** inserem ou exclui a linha ou coluna selecionada; as setas movem a **linha** (↑↓) ou a **coluna** (←→), conforme o modo de seleção.
- Mesclagem de linhas/colunas/células; botão para restaurar footers colunizados.
- No relatório gerado: células e fundos transparentes onde necessário para a **marca d'água** aparecer por cima.

### Conceito de domínio (dados hierárquicos)
- **Linha detalhe** = profundidade fixa no datasource.
- **Grupos** = ancestrais na árvore.
- Na configuração, especifica-se sobretudo a **profundidade do detalhe**; o restante deduz-se.

---

## 11. Marca d'água

- Sempre **imagem** (Data URL), não texto.
- Configurar: botão **💧**; remover: **Ctrl+clique** no mesmo botão.
- No PDF paginado: overlay por cima do conteúdo (`z-index` elevado). Tabelas/células não devem tapar a marca.

---

## 12. JSON de relatório

1. **Carregar** JSON de exemplo ou produção (📂). Exemplo: `crudex.json`, `wordex-correto.json`.
2. Macros `{{NomeCampo}}` no template são resolvidas na montagem.
3. **Kind** nos wrappers Crudex: `"collection"`, `"total"`, `"histogram"` (gráficos `*Grafico`).
4. Campos escalares: `"string"` (varchar), `"number"`, `"datetime"`, `"boolean"`, `"image"`. No Crudex oficial com API, `text` = varchar(max) e `string` = varchar(n); no Wordex/VBA exportamos **`string`**. JSON antigo com `"Kind": "text"` ainda é aceito no editor.
5. Histogramas ligados no editor viram dados no grid do gráfico (§8); no PDF usam o que está salvo no template.
6. Imagens: preferir **base64** no JSON para PDF confiável; URLs externas exigem HTTP acessível ao gerador.
7. **Montar template HTML** aplica os dados e abre o fluxo paginado.

---

## 13. Geração de PDF

**Regra:** o PDF sai sempre do **HTML paginado** (`.pagex-page`), não do template/editor.

| Origem | Vai para PDF? |
|--------|----------------|
| Janela paginada após "Montar template HTML" (`wordex-paged.html`, Salvar HTML) | Sim |
| `WORDEX.html` (editor, `#wordexDocument`, `.body-flow`) | Não |
| HTML montado antes da paginação (sem `.pagex-page`) | Não |

No backend C# (`WordexChromePdf`), `RequirePaginatedHtml` vem `true` por padrão e rejeita arquivos que não sejam paginados.

### Pelo editor (fluxo manual)
1. **Montar template HTML** a partir do JSON (abre a janela paginada).
2. Na toolbar paginada: **Gerar PDF** (ou **Salvar HTML** e converter no servidor).

### Pela API (iframe / integração)
```javascript
const dataUri = await iframe.contentWindow.ObterPDF(json);
// retorna: data:application/pdf;base64,...
```
O `ObterPDF` monta e pagina internamente antes de capturar — o usuário não precisa salvar o HTML.

### Backend (Windows Service / servidor)

**Não pagine em C#.** A paginação usa DOM e JavaScript (`wordex-paged.html`). Com template + JSON, o caminho é headless + `ObterPDF(json)` — o mesmo pipeline do iframe, no servidor.

```text
DataSet → JSON (C#) ─┐
                     ├→ HTTP + WORDEX.html → ObterPDF(json) → PDF
Template (WORDEX.html salvo) ─┘
```

Detalhes completos: **`helpBACKEND.txt`**.

| Variante | Como | Quando |
|----------|------|--------|
| **A — ObterPDF** | Playwright/Puppeteer chama `ObterPDF(json)` | Mais simples; um passo |
| **B — Chrome print** | Headless pagina → HTML `.pagex-page` → `WordexChromePdf` | Qualidade de impressão Chromium |

Publicar na pasta do serviço: `WORDEX.html` (💾 Salvar template HTML), `wordex-paged.html`, `wordex-pdf.html`. Servir por HTTP.

```csharp
// Variante B — só se já tiver HTML paginado salvo (janela paginada)
crudex.GerarPdfDeArquivoHtml(@"D:\saida\relatorio-paginado.html", @"D:\saida\relatorio.pdf");
```

Ver também `helpCHROME.txt` (Chrome CLI / Puppeteer).

---

## 14. Pipeline técnico (referência rápida)

```text
wordexDocument (template)
    → buildGeneratedReportHtml(data)
    → wordex-paged.html (paginateBlocks, .pagex-page)
    → wordex-pdf.html (captura html2canvas / jsPDF)
    → PDF
```

Placeholders `@@...@@` nos templates embutidos. Config em `#wordex-pagex-config` e `#wordex-pdf-config`.

---

## 15. Decisões e armadilhas (não esquecer)

| Tema | Decisão / cuidado |
|------|-------------------|
| Gaps do gráfico | Superior e inferior **independentes** para permitir textboxes nos vãos |
| Tipografia do gráfico | Isolada em `data-chart-font-*`; não usar `style` inline no `.wordex-chart-box` para fonte/cor |
| Novos parágrafos | Não herdam fonte da toolbar quando ela reflete objeto (gráfico); usam estilo do bloco hospedeiro |
| Borda do gráfico | Obrigatória; não remover via preset `none` |
| Altura do gráfico | **Não** reintroduzir `fitChartBoxHeightToContent` (quebrou resize pelas cantoneiras) |
| Paginação | Usar altura **real** dos blocos; evitar inflar `min-height` só para paginar |
| PDF | `printBackground: true`; overflow visível em objetos na captura |
| Marca d'água | Fundos de tabela transparentes no paginado/PDF |
| Histograma + JSON | Importa para o grid ▥; edição manual = simulação persistida no template |
| Picker de histograma | Fora de tabela: ROOT + histogramas; em célula: coleção da linha fixa |
| Testes | Servir por HTTP; `file://` limita imagens e PDF |

---

## 16. Desfazer

O editor mantém pilha de **undo** após mudanças estruturais (`saveUndoStateIfChanged`). Usar quando disponível na UI ou atalho configurado.

---

## 17. O que ainda pode entrar neste manual

- [x] Tipografia do gráfico isolada do parágrafo (§8)
- [x] Passo a passo: selecionar histograma, ▥ Dados, JSON → grid, simulações (§8)
- [ ] Passo a passo com capturas (parágrafo + gráfico + textboxes)
- [ ] Estrutura completa do JSON (`reportData`) com exemplos campo a campo
- [ ] Tabelas: linha a linha dos tipos de row (`Header`, `Group`, `Detail`, `Free`, …)
- [ ] Integração Crudex → Wordex (quando o fluxo estiver fechado)
- [x] Backend: template + JSON → PDF sem paginar em C# (`helpBACKEND.txt`)
- [ ] Troubleshooting PDF (imagem faltando, corte, página em branco)

---

## 18. Manutenção deste documento

Ao implementar comportamento novo no Wordex:

1. Registrar **o que** o usuário faz (gesto / botão).
2. Registrar **por que** existe (caso de uso).
3. Atualizar §15 se for decisão que não deve ser revertida sem motivo.

Arquivos de apoio para quem edita código: `.cursor/skills/wordex/SKILL.md` e `reference.md`.
