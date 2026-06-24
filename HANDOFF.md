# Doodle Studio — Handoff Document

> **Para a próxima IA:** este documento cobre 100% do projeto. Leia antes de tocar em qualquer arquivo.

---

## O que é este projeto

**Doodle Studio** é uma ferramenta de desenho vetorial animado 100% offline, em arquivo único HTML/CSS/JS — sem build, sem dependências externas, sem CDN. Tudo roda no browser diretamente abrindo o arquivo.

Existem **três versões**:
- `index.html` — **V1 (NUNCA MODIFICAR).** Versão estável de referência.
- `v2.html` — **V2 (checkpoint).** Redesign de UX mantendo 100% da lógica da V1.
- `v3.html` — **V3 (versão ativa).** Cópia da V2 + grupos, export cropado, sidebar enxuta, timeline melhor.

### Novidades da V3 (sobre a V2)
- **Tema claro:** fundo `#EEECE6`, acento magenta (`#C8185C`). Tokens em `:root`.
- **Sidebar enxuta:** painel fixo `.panel-fixed` "Pincel" com cor + paleta (compacta) + opacidade
  + tamanho + **Estabilizador + Suavização Vetorial** sempre visíveis; só "Textura do Traço"
  (`#texAcc`) colapsa. "Imagem de Referência" também virou `.panel-fixed` (sem colapsar).
  Só "Falhas na Tinta" permanece `.acc` colapsável.
- **Grupos de traços:** `stroke.groupId`; `state.multiSel` (Set de índices), `state.nextGroupId`.
  Ctrl/Shift+clique nos blocos da timeline multi-seleciona → botão **Agrupar** (`#tlGroupBtn`).
  Helpers `groupsMap()`, `groupColor(gid)`, `renderGroupBar()`, `syncGroupsUI()`. Chips de grupo
  com "×" para excluir; clique no chip re-seleciona os traços. Tags `G{n}` nos blocos.
- **Export PNG cropado:** `contentBBox(strokeList)` + `renderCropPNG(strokeList, colorOverride,
  margin=10)` + `downloadCanvas()`. `confirmExport` agora: imagem principal (1920×1080, ou
  cropada no conteúdo se `#cropContent` marcado) **+** 1 PNG por grupo cropado com 10px de margem.
- **Defaults de export animado:** GIF/`state.anim.gifScale` = **1920**; FPS/`state.anim.fps` = **30**.
- **GIF/WebM por escopo:** `exportGif(job)` / `exportWebM(job)` aceitam `{strokes,label,crop}`.
  `runAnimExport(format)` lê `#animScope` (Canvas/Grupos/Ambos) + `#animCrop` e dispara os jobs
  em sequência (await; o GIF tem `_gifBusy`). Helpers `revealsFor(list,p,e)` (reveals re-baseados
  ao subconjunto), `renderAnimatedList()`, `jobViewport(list,crop)` (full-frame ou bbox+10px).
  GIF mantém densidade `gifScale/1920`; WebM é 1:1. Grupos cropados saem com nome `-g<id>`.
- **GIF poster frame:** o **primeiro frame do GIF é o estado FINAL** (desenho completo), com
  `disposal=2` (restore to background = transparente) para limpar antes do reveal animar. Assim
  thumbnails/documentos estáticos (que mostram o frame 0) exibem a arte pronta; no play mostra o
  full rapidinho → limpa → anima → full → loop. Total de frames = `N+1`. Helpers internos no
  `exportGif`: `quantize(data)` e `emitFrame(delay, disposal)`.
- **Timeline:** Sequencial/Juntos viraram segmented (`.seg`/`.seg-btn`, classe `.active` marca o
  modo) com `ⓘ` explicativo (`#tlInfoPop`). Toggle **esconder/mostrar** a timeline na action bar
  (`#tlToggleBtn` → classe `.tl-off` no `.dock`; `state.tlVisible`).
- **Excluir traço qualquer:** `deleteSelection()` (botão 🗑 na barra de seleção + tecla Del/
  Backspace) remove o(s) traço(s) selecionado(s) por identidade — não só o último. Não entra no
  undo (limpa o redo).
- **Trocar de camada (z-order):** botões ⬆/⬇ na barra de seleção (`moveSelectedLayer(dir)`) **ou**
  arrastar o bloco na **vertical** na timeline (lane alvo destacada com `.drop-target`). Horizontal
  continua = `animStart`. `remapSelection()` mantém seleção/multiSel por identidade após reordenar.
- **Duração padrão:** `state.anim.duration` = **2s** (slider `#animDur` value 20).
- **Colar imagem (Ctrl+V):** listener `paste` no document → se houver imagem no clipboard,
  `setRefImage(dataUrl)` a usa como imagem de referência com a opacidade atual (`#refOp`).

---

## Arquitetura geral

### Canvas e desenho
- Canvas HTML5 2D, resolução interna **1920×1080**
- O usuário desenha com mouse/touch; os pontos brutos ficam em `state.strokes[i].raw`
- Após soltar o mouse (`finishStroke`), os pontos passam por pipeline de vetorização:
  1. **Gaussian smoothing** (estabilizador de movimento, slider `mouseSmooth`)
  2. **RDP simplification** (remove pontos redundantes)
  3. **Chaikin subdivision** (suavização vetorial, slider `vectorSmooth`)
- O resultado vetorizado é cacheado em `stroke._cache` para não reprocessar a cada render

### Texturas de traço
Implementadas em `drawStroke(ctx, stroke, ptsOverride)`. Cada textura usa `rand(x, y)` baseado em **posição** (não em tempo), então o traço é estável entre frames da animação.

| Textura | data-tex | Técnica |
|---------|----------|---------|
| Giz | `giz` | Múltiplas linhas com jitter de posição e opacidade |
| Sólido | `solido` | Linha simples com globalAlpha |
| Marcador | `marcador` | Linha grossa semi-transparente |
| Lápis | `lapis` | Ruído de grão fino |
| Aquarela | `aquarela` | Blobs de cor com blur |
| Spray | `spray` | Pontos aleatórios em raio variável |
| Nanquim | `nanquim` | Linha fina com variação de espessura |
| Pontilhado | `pontilhado` | Série de pontos ao longo do traço |
| Patinhas | `patinhas` | Elipses simulando patas de cachorro (palma + 4 dedos), orientadas pela direção do traço |

Podem ser combinadas **até 2 texturas simultâneas** (primeira = primária, segunda = `.secondary`).

### Falhas na tinta
Controladas por `state.fault` (`none` | `quebra` | `falha` | `respingo`) + sliders `faultLvl` (intensidade) e `gapSize` (tamanho da quebra). Aplicadas sobre os pontos vetorizados antes de renderizar o traço.

> **Nota:** Os sliders `faultLvl` e `gapSize` também controlam os parâmetros da textura **Patinhas** (distância e tamanho das marcas).

---

## Estado global (`state`)

```js
state = {
  strokes: [],          // array de traços na ordem temporal
  selectedStrokeIdx: null,
  currentColor: '#FD5E6D',
  currentOpacity: 100,
  brushSize: 8,
  textures: ['giz'],    // até 2
  mouseSmooth: 50,
  vectorSmooth: 0,
  fault: 'none',
  faultLvl: 16,
  gapSize: 54,
  refImage: null,
  refOpacity: 0.4,
  zoom: 1,              // V2 only
  redo: [],             // V2 only
  anim: {
    playing: false,
    duration: 4,
    easing: 'linear',
    loop: true,
    holdMs: 600,
    fps: 15,
    animScale: '640x360',
    webmBg: '#FFFFFF',
    _raf: null,
    _t0: 0,
  }
}
```

### Estrutura de um traço (`stroke`)

```js
{
  raw: [{x, y}, ...],     // pontos brutos na ordem do mouse
  color: '#FD5E6D',
  opacity: 100,
  size: 8,
  textures: ['giz'],
  fault: 'none',
  faultLvl: 16,
  gapSize: 54,
  animStart: 0,           // instante de início na timeline (em unidades de arco)
  _cache: null,           // pontos vetorizados cacheados
  _arc: null,             // comprimento de arco cacheado
}
```

---

## Sistema de animação

### Motor de revelação progressiva

```js
strokeArcLen(stroke)        // comprimento de arco do traço vetorizado
strokeDur(stroke)           // Math.max(arcLen, 4) — duração proporcional
strokeReveals(progress, easing) // → array com quanto revelar de cada traço em [0, dur]
```

A timeline mapeia `progress ∈ [0,1]` → tempo real `tc` usando `T = max(animStart + dur)` para todos os traços. Cada traço revela `clamp(tc - animStart, 0, dur)` de seu comprimento de arco.

### Timeline arrastável

- Cada traço tem `animStart` (em unidades de comprimento de arco)
- **Sequencial** (default): `animStart` cumulativo — um começa quando o anterior termina
- **Juntos**: todos `animStart = 0`
- Arrastar blocos na timeline muda `animStart` com snap automático
- O playhead reflete o progresso durante o preview

### Curvas de easing disponíveis

```js
EASING = {
  linear:    t => t,
  easeIn:    t => t * t,
  easeOut:   t => 1 - (1-t)*(1-t),
  easeInOut: t => t < 0.5 ? 2*t*t : 1 - Math.pow(-2*t+2,2)/2
}
```

---

## Export

### PNG (`exportBtn`)
- Modal com paleta de cores selecionável
- Checkbox "Exportar com cores originais" (default marcado)
- Se cores selecionadas: exporta um PNG por cor substituindo todas as cores pelo tom escolhido
- Se original: exporta o canvas como está

### SVG (`svgBtn`)
- Converte os traços vetorizados para `<path>` SVG
- Preserva todas as cores e opacidades

### GIF (`gifBtn`)
- Encoder LZW GIF89a **embutido** (100% offline, sem bibliotecas)
- Fundo transparente (GIF com canal alpha via índice transparente)
- Transparência binária: `alpha < 128` → transparente
- Quadros revelam progressivamente usando `strokeReveals()`
- **Timing canônico Weiner/giflib**: o incremento de code-size acontece DENTRO de `output()`, DEPOIS de emitir, quando `freeEnt > maxcode` — isso é crítico, off-by-one causa corrupção total
- Disposal method = 1 ("do not dispose") — necessário porque frames são cumulativos
- `_gifBusy` flag para evitar exports concorrentes corrompendo o `LZW_DICT` (Int32Array global)
- Loop infinito via `NETSCAPE2.0`

### WebM (`webmBtn`)
- Usa `MediaRecorder` nativo do browser
- Fundo sólido (sem transparência — WebM com alpha não é suportado amplamente)
- `canvas.captureStream(0)` + `track.requestFrame()` para controle determinístico de frames

---

## Layout V2 (v2.html)

### CSS Grid structure

```
body { grid-template-rows: auto 1fr auto }
  ├── header.topbar        (auto)
  ├── main                 (1fr, grid-template-columns: 312px 1fr)
  │   ├── aside.tools      (sidebar colapsável)
  │   └── .canvas-area     (área do canvas com barra de ações)
  └── .dock                (286px, grid-template-columns: 312px 1fr 188px)
      ├── col 1: transporte de animação
      ├── col 2: timeline com régua
      └── col 3: exports
```

A variável `--sidebar-w: 312px` é usada tanto no `main` quanto no `.dock` para garantir alinhamento perfeito entre sidebar e coluna de transporte.

### Tokens de design

```css
--bg: #EEECE6           /* fundo geral — beige quente */
--surface: #F5F3ED      /* cards e painéis */
--surface-2: #E6E4DC    /* inputs e hover */
--surface-3: #DAD8CF    /* destaque leve */
--border: #C8C5BC
--border-2: #D6D3CA
--text: #1C1A17         /* quase-preto quente */
--muted: #7A776E
--muted-2: #524E47
--accent: #C8185C       /* magenta — cor de marca principal */
--accent-2: #E0226E     /* magenta brilhante */
--accent-soft: rgba(200,24,92,0.10)
--danger: #C03040
```

### Paleta de cores do usuário (5 fixas)
```
#FD5E6D  (coral)
#436AE1  (azul)
#EEECE6  (creme — igual ao fundo)
#FFFFFF  (branco)
#000000  (preto)
```

### Sidebar — Acordeão

5 seções colapsáveis via `.acc` / `.acc-head` / `.acc-body`:
- **Cor** — paleta, hex, color picker, opacidade
- **Pincel** — tamanho + sub-acordeão de Textura (colapsado por default)
- **Suavização** — estabilizador do movimento, suavização vetorial
- **Falhas na Tinta** — tipo de falha + intensidade + tamanho da quebra
- **Referência** — imagem de referência (carregar/remover/opacidade)

Estado do acordeão: classe `.collapsed` no `.acc`. Sub-acordeão de textura: `.sub-acc.collapsed`.

### Barra de ações do canvas (V2)

Botões: Desfazer | Refazer | Limpar | Zoom − | 100% | Zoom + | Ajustar

**Redo stack**: `state.redo = []`. Ao desfazer: `redo.push(strokes.pop())`. Ao refazer: `strokes.push(redo.pop())`. Limpar o redo ao desenhar um novo traço (`finishStroke`).

**Zoom**: aplicado via `transform: scale(state.zoom)` no `#canvasWrap`. O `getPos()` usa `getBoundingClientRect()` então as coordenadas de desenho continuam corretas com qualquer zoom.

### Indicador de pincel ativo (header)

Chip no header mostrando textura + cor + tamanho atual. Atualiza via `updateBrushChip()` que é injetado com wrapper nas funções `selectColor`, `updateTextureUI`, e no `input` do slider de tamanho.

### Atalhos de teclado (V2)

| Tecla | Ação |
|-------|------|
| `Space` | Play/Pause |
| `Ctrl+Z` | Desfazer |
| `Ctrl+Shift+Z` | Refazer |
| `Esc` | Deselecionar traço |

---

## IDs críticos (o script referencia todos estes)

```
hexInput, hexSw, colorPicker, pickBtn, palette, opacity, opacityVal
size, sizeVal, textureGrid (.tex-btn[data-tex]), texHint
mouseSmooth, mouseSmoothVal, vectorSmooth, vectorSmoothVal
faultGrid (.fault-btn[data-fault]), faultLvl, faultLvlVal, gapSize, gapSizeVal
refOp, loadImgBtn, removeImgBtn, imgInput, refImg
strokeSelectedInfo, deselectBtn
playBtn, resetAnimBtn, animDur, animDurVal, animHold, animHoldVal
animEasing, animLoop, animFps, animScale, webmBg, webmBgPick
gifBtn, webmBtn, svgBtn, exportBtn
undoBtn, redoBtn (V2), clearBtn
drawCanvas, canvasWrap, statChip, status, resLabel, toast
exportModal (+ filhos: exportOriginal, modalColors, exportConfirmBtn, modalCloseBtn)
tlWrap, tlLanes, tlPlayhead, tlEmpty, tlSeqBtn, tlTogetherBtn
zoomLabel, brushChip (V2), bcSw, bcTex, bcSize
```

---

## Bugs conhecidos e correções aplicadas

### GIF — timing LZW canônico (RESOLVIDO)
Sintoma: GIF abre rosa no Windows / branco no browser / `truncated` no PIL.
Causa: o incremento de `codeSize` ocorria 1 código antes do esperado pelos decodificadores.
Correção: mover o `if (nextCode > (1<<codeSize)-1) codeSize++` para DENTRO de `output()`, APÓS emitir os bits. Nunca reverter isso.

### GIF — índice de transparência
O `bgColorIndex` no header GIF deve apontar para o índice transparente, não para `palette[0]`.

### GIF — disposal method
`disposalMethod = 1` (do not dispose). Com `disposalMethod = 2` os frames anteriores são apagados, quebrando a animação cumulativa.

### GIF — exports concorrentes
`LZW_DICT` é um `Int32Array` global compartilhado. Flag `_gifBusy` previne dois exports simultâneos corrompendo o dicionário.

---

## O que está pendente / próximas ideias

- [ ] Persistência em `localStorage` (abas abertas, última cor, etc.)
- [ ] Arrastar para redimensionar duração de um traço na timeline
- [ ] Tema claro/escuro toggle (a V2 já é tema claro)
- [ ] Export de frame único (PNG do momento atual da timeline)
- [ ] Numeração de traços na timeline com hover highlight no canvas
- [ ] Subir para GitHub + deploy no Vercel (próximo passo imediato)

---

## Estrutura de arquivos

```
Doodle Maker/
├── index.html      ← V1 — NUNCA MODIFICAR
├── v2.html         ← V2 — versão ativa
├── HANDOFF.md      ← este arquivo
└── .claude/
    └── launch.json ← config do servidor local (npx http-server porta 5173)
```

Para rodar localmente:
```bash
npx http-server . -p 5173 -c-1
# Abre: http://localhost:5173/v2.html
```

---

## Contexto do usuário

- **Nome da conta GitHub:** `herickcba`
- **Email:** herickmf@gmail.com
- O usuário trabalha no Windows 11 com PowerShell
- Prefere respostas diretas e sem enrolação
- A V1 é sagrada — nunca toque nela
