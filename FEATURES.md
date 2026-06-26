# Doodle Studio — Web (completo) × Add-in PowerPoint

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
| Biblioteca: salvar **GIF** (com tag) | n/a | ✅ | na tela grande, "Salvar GIF na biblioteca" guarda traços + opções de animação (regenera no insert) com tag `GIF`; salva **também** um PNG do estado final estático |
| Biblioteca: tag de **loop / sem loop** | n/a | ✅ | badge `GIF ↻` (em loop) vs `GIF 1×` (sem loop), conforme a escolha ao salvar |
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
