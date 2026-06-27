# Doodle Studio — Web (completo) × Add-in PowerPoint

> **Add-in vira plataforma multi-módulo.** O task pane tem **abas no topo** (`src/shared/modules.js`): cada módulo é uma `view`. Módulo 1 = **Doodle**. Módulo 2 = **Tipografia** (estilos B+G). Módulo 3 = **Imagem** (Nano Banana). Próximos módulos entram como novas abas.

## Módulo Imagem (Nano Banana / Gemini) — add-in
Gera/edita imagens **1920×1080** (16:9). Modelos selecionáveis: **Nano Banana 2** (`gemini-3.1-flash-image`, default) e **Pro** (`gemini-3-pro-image`). Arquivos: `src/modules/image-gen.js`, `src/shared/img-library.js`, dialog `src/dialog/imgedit.html`, funções `api/*`.

| Recurso | Como |
|---|---|
| Prompt → imagem | `improve-prompt` (Gemini texto → JSON) → `generate-image` (16:9) → reescala p/ 1920×1080 → galeria |
| Toggle de **consistência** | chips `imagem·estilo·universo·textura·personagem·composição` enviados ao `improve-prompt` (varia o prompt mantendo o que estiver ligado) |
| **Imagem enviada: Referência ou Editar** | toggle por imagem ativa. **Referência** = estilo numa geração nova. **Editar** = a imagem vira a BASE → abre o editor (marcar a região + prompt) p/ uma edição fiel que preserva o fundo/cena |
| **Style reference** | upload/colar imagem (`DoodleImgRefs`, ≤1280px, localStorage) → enviada junto na geração |
| Edição fiel (inpainting com máscara) | manda **2 imagens**: base + **máscara branco/preto** (pinte a região → vira branco). O modelo edita só a área branca e mantém o preto idêntico. Sem pintar = edição semântica. Instrução reforçada de preservação. (`toMask` no painel; `edit-image`/`gemini.editImage`) |
| **Batch** | 1–4 imagens em sequência |
| Inserir no slide | `OfficeBridge.insertImage` preenche o slide (16:9) |
| **Editar** | tela grande (`imgedit.html`) com **pincel sólido** (sem giz) → base + marcação + prompt → `edit-image` |
| Galeria/biblioteca | **persistente** (IndexedDB, `DoodleGallery`) — sobrevive reabrir; **excluir** por card (×); **editar/inserir/★ref** qualquer uma. Refs persistem (downscale) |
| Editar imagem (já editada) | editor recebe cópia **leve (JPEG ≤1280px)** p/ display (não estoura o localStorage); composição base+marcação no painel em alta |

**Dois modos (dispatch em `image-gen.js` → `src/shared/gemini.js`):**
- **Serviço (produção):** sem chave local → chama `api/improve-prompt` · `api/generate-image` · `api/edit-image` (mesma origem); `GEMINI_API_KEY` fica **só no servidor** (env var). Setup: `vercel env add GEMINI_API_KEY`.
- **Teste (BYO key):** o usuário cola a chave na caixa "Chave Gemini (modo teste)" → fica no `localStorage` e o add-in chama o **Gemini direto** do navegador (`x-goog-api-key`). Destrava o teste sem servidor; trocar pra serviço é só limpar a chave.

Chave **nunca** no repo (`.env.local`/`.gitignore`; `.env.example` = placeholder).

## Módulo Tipografia (estilos B+G) — add-in
Estilos **Hero / Mega / H1** aplicáveis a um objeto ou texto selecionado, com prévia clicável (sem dropdown). Arquivos: `src/modules/typography.js`, `OfficeBridge.applyTextStyle` / `insertStylesReference`, asset `assets/estilos-bg.pptx` (gerado por `tools/build-styles-reference.py`).

| Atributo | Como é aplicado |
|---|---|
| Fonte (Avenir Next), negrito, tamanho, cor | **Na hora**, à seleção, via Office.js `TextRange.font` (PowerPointApi **1.6** p/ ler a seleção) |
| Ponto final em outra cor (Hero) | `getSubstring(len-1,1).font.color` |
| Cor da fonte conforme o fundo (Hero) | Best-effort: lê o preenchimento do objeto; se `436AE1` → branca (ponto `FC5E6D`) |
| **Entrelinha (0.8x/0.9x)** | **Não é setável via Office.js** → vem da **âncora**: `insertStylesReference()` insere um slide de referência (`estilos-bg.pptx`, entrelinha assada no OOXML) e o usuário **copia a caixa** ou usa o **Pincel de Formatação** (nativos, preservam a entrelinha) |

Estilos (11): **Hero** 120·0.8·`FC5E6D` (ponto `436AE1`; sobre fundo azul → branca) · **Mega** 80·0.9·`436AE1` · **H1** 60·0.9·`FC5E6D` · **Label de Seção** 60·0.9·`FC5E6D` · **Corpo de Texto** 44·1.15·`436AE1`·Reg · **H3** 34·0.95·`FC5E6D` · **H4** 28·1.0·`436AE1` · **H5** 24·1.0·`436AE1`·Reg · **Corpo de Pilar** 20·1.3·`436AE1`·Reg · **Eyebrow** 18·1.0·`FC5E6D` · **Caption** 16·1.3·`436AE1`·Reg. (Bold salvo onde não marcado Reg.) Entrelinha sempre **múltiplo** (spcPct) na referência.

---

Duas versões mantidas em paralelo, com o **mesmo motor de desenho**:

- **Web (v3)** — o app **completo**. Arquivo: [`v2.html`](v2.html). No ar: https://doodle-studio-sigma.vercel.app/v2.html
- **Add-in PowerPoint** — a **versão app** (subconjunto focado em inserir no slide). Pasta: [`v3-powerpoint-addin/`](v3-powerpoint-addin/). No ar: https://doodle-studio-app.vercel.app/

> **Regra:** a Web é o *superset* (tem tudo). O Add-in é um *subconjunto* enxuto. Toda feature nova deve ser classificada aqui: **web-only**, **add-in-only** ou **ambos**. O motor de desenho (`doodle-engine.js` no add-in) deve ficar sincronizado com a lógica de `v2.html`.

## Tabela de features

| Feature | Web | Add-in | Notas |
|---|:---:|:---:|---|
| Desenho vetorial + suavização (gaussian/RDP/Chaikin) | ✅ | ✅ | mesmo motor |
| `Shift` = linha reta | ✅ | ✅ | |
| Desfazer / Refazer / Limpar | ✅ | ✅ | |
| Cor + paleta + hex + color picker | ✅ | ✅ | |
| Opacidade do traço | ✅ | ❌ | removida no add-in (simplicidade) |
| Tamanho do pincel | ✅ slider | ❌ | fixo em 8px no add-in |
| Texturas | ✅ 9 | ⚠️ só **giz** | sólido, marcador, lápis, aquarela, spray, nanquim, pontilhado, patinhas só na web |
| Combinar 2 texturas | ✅ | ❌ | |
| Falhas na tinta (quebra/falha/respingo) | ✅ | ❌ | removidas no add-in |
| Animação (revelar traço por traço) | ✅ timeline arrastável | ⚠️ sequencial + duração | add-in: versão simples, sem timeline |
| GIF animado | ✅ (exporta arquivo) | ✅ (insere no slide) | encoder GIF89a embutido, compartilhado. No add-in **só na tela grande** |
| GIF: 1º quadro = estado final | ✅ | ✅ | quadro-pôster (disposal=2) garante que PDF/render estático mostre o desenho completo; reveal continua começando do zero |
| Controles de animação do GIF (curva/FPS/segurar no final) | ✅ | ✅ tela grande | curva linear/easeIn/easeOut/easeInOut, FPS, hold ms |
| GIF com/sem loop · GIFs por traço separado | ✅ | ✅ tela grande | |
| Export WebM | ✅ | ❌ | só web |
| Export PNG / SVG (arquivo) | ✅ | ❌ | só web (no add-in, vira "inserir no slide") |
| Zoom do canvas | ✅ | ✅ tela grande | −/100%/+/Ajustar na barra da tela grande |
| Selecionar traço (clique) | ✅ | ✅ | clique sem arrasto destaca o traço |
| **Biblioteca de traços** (salvar/reusar) | ✅ | ✅ | salvar desenho inteiro ou traço selecionado; **clicar na miniatura insere no slide no tamanho salvo**; global entre decks (`localStorage`); seção colapsável no painel; salvar também na tela grande |
| Biblioteca: salvar **GIF** (3 versões) | n/a | ✅ | "Salvar como GIF" grava **sempre 3 itens**: GIF em loop + GIF sem loop + PNG do estado final. Disponível na **barra lateral** (configs padrão) e na **tela grande** (configs de animação). Não há mais opção de loop |
| Inserir GIF = inserir + salvar | n/a | ✅ | na tela grande, "Inserir GIF no slide" insere a versão **sem loop** e salva as 3 versões na biblioteca |
| Biblioteca: tag de **loop / sem loop** | n/a | ✅ | badge `GIF ↻` (em loop) vs `GIF 1×` (sem loop) |
| **Inserir no slide** (PNG) | ❌ | ✅ | **só add-in** (`setSelectedDataAsync`) |
| Fundo = imagem do slide atual | ❌ | ✅ | via copiar+colar (Cmd+V / "Fundo do slide") |
| Tela grande (Office Dialog) | ❌ | ✅ | concentra GIF + animação + zoom (painel fica enxuto) |
| Detectar tamanho real do slide | ❌ | ✅ | lê `sldSz` do `.pptx` |
| Inserir traços **separados ou juntos** | n/a | ✅ | espelha o modelo de "vários traços" da web |
| Botão na faixa (ribbon) do PowerPoint | n/a | ✅ | atalho como ChatGPT/Claude |

## Compatibilidade / paralelismo

- **Motor compartilhado:** `v3-powerpoint-addin/src/engine/doodle-engine.js` é o motor de `v2.html` extraído (já com animação/GIF, seleção de traço, `loadStrokes`/`thumbnailOf`). Bug/ajuste no motor deve ser replicado nos dois.
- **Biblioteca:** `v3-powerpoint-addin/src/shared/library.js` guarda em `localStorage['doodle.library']`; painel e tela grande compartilham (mesma origem). Salvar funciona dos dois lados; navegar/carregar fica no painel.
- **Tokens visuais:** ambos usam a mesma paleta (cinza-neutro + azul `#436AE1`) e tipografia.
- **O que NÃO migra pro add-in:** animação e exports de arquivo (dependências/peso e fora do caso de uso "desenhar no slide").
- **O que é exclusivo do add-in:** tudo que toca a API do Office (inserir, ler slide, dialog, ribbon).

_Atualize esta tabela sempre que adicionar/remover uma feature de qualquer versão._
