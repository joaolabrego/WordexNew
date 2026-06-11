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
| `helpCHROME.txt` | PDF no backend (Chrome / Puppeteer) |
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

### Histograma (com gráfico selecionado)
- **Horizontal** — barras horizontais
- **Modo:** `normal` | `acumulado` | `misto`
- **Dados do histograma** — editor de categorias, séries e valores
- **Cor da série** — clique numa barra para escolher qual série colorir

### Texto / Alinhar / Borda / Margem
- Formatação de texto, alinhamento de parágrafo ou objeto
- Bordas (tabela, célula, gráfico, imagem, textbox)
- Margem do **bloco** parágrafo (não confundir com gaps internos do gráfico — ver §8)

### Dados
- **📂 Carregar JSON**
- **🖼 Imagem local** — no objeto selecionado ou novo (Ctrl+clique limpa imagem)
- **💧 Marca d'água** — imagem de fundo (Ctrl+clique remove)
- Seletores de coleção/campo macro (com JSON carregado)

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

### Gráfico — legenda
Com o gráfico selecionado, a **legenda** (várias séries) pode ser **arrastada** dentro da área do gráfico.

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

### Dados
- Diálogo **Dados do histograma**: categorias (linhas), séries (colunas), valores.
- Pode vir de JSON (`chartData` no objeto) ou edição manual no diálogo.
- Com 2+ séries, aparece **painel de legenda** (arrastável).

---

## 9. Textboxes

- Caixa de texto editável, borda tracejada no editor.
- **Clique** — editar texto.
- **Ctrl+clique** — selecionar como objeto (mover, redimensionar, borda, margem).
- Úteis nos **gaps** acima/abaixo do gráfico (ver §7).

---

## 10. Tabelas

- Modos de seleção: célula, linha, coluna, tabela inteira (Ctrl+clique cicla).
- Linhas estruturais (cabeçalho, grupo, rodapé) e linhas **Free** — ver toolbar **Linhas** / **Colunas**.
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

1. **Carregar** JSON de exemplo ou produção (📂).
2. Macros `{{NomeCampo}}` no template são resolvidas na montagem.
3. Imagens: preferir **base64** no JSON para PDF confiável; URLs externas exigem HTTP acessível ao gerador.
4. **Montar template HTML** aplica os dados e abre o fluxo paginado.

---

## 13. Geração de PDF

### Pelo editor (fluxo manual)
1. Montar HTML a partir do JSON.
2. Na toolbar paginada: **Gerar PDF**.

### Pela API (iframe / integração)
```javascript
const dataUri = await iframe.contentWindow.ObterPDF(json);
// retorna: data:application/pdf;base64,...
```

Detalhes: `helpIFRAME.txt` (mesma origem) e `postMessage` para origens diferentes.

### Backend
Preferir Puppeteer com `printBackground: true`. Ver `helpCHROME.txt`.

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
| Testes | Servir por HTTP; `file://` limita imagens e PDF |

---

## 16. Desfazer

O editor mantém pilha de **undo** após mudanças estruturais (`saveUndoStateIfChanged`). Usar quando disponível na UI ou atalho configurado.

---

## 17. O que ainda pode entrar neste manual

- [x] Tipografia do gráfico isolada do parágrafo (§8)
- [ ] Passo a passo com capturas (parágrafo + gráfico + textboxes)
- [ ] Estrutura completa do JSON (`reportData`) com exemplos campo a campo
- [ ] Tabelas: linha a linha dos tipos de row (`Header`, `Group`, `Detail`, `Free`, …)
- [ ] Integração Crudex → Wordex (quando o fluxo estiver fechado)
- [ ] Troubleshooting PDF (imagem faltando, corte, página em branco)

---

## 18. Manutenção deste documento

Ao implementar comportamento novo no Wordex:

1. Registrar **o que** o usuário faz (gesto / botão).
2. Registrar **por que** existe (caso de uso).
3. Atualizar §15 se for decisão que não deve ser revertida sem motivo.

Arquivos de apoio para quem edita código: `.cursor/skills/wordex/SKILL.md` e `reference.md`.
