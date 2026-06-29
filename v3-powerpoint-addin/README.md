# CBA Studio v3 — Add-in de PowerPoint

Desenhe doodles vetoriais (textura giz, suavização, falhas na tinta) **sobre o slide atual** e insira o resultado como **imagem PNG transparente**, posicionada onde você desenhou.

Reaproveita o motor de desenho do CBA Studio web (`../v2.html`). Sem animação (MVP).

## O que está incluído (MVP)
Cor + paleta, opacidade, tamanho, textura **giz**, suavização (movimento + vetorial), falhas na tinta (quebra/falha/respingo + intensidade + tamanho da quebra), **linha reta com `Shift`**, desfazer/refazer, e **inserir no slide**.

Duas superfícies de desenho:
- **Painel lateral** — rascunho rápido, sem trocar de janela.
- **⤢ Tela grande** — janela ampla (Office Dialog) com o **slide ao fundo em opacidade ajustável** para desenhar com referência. Ao "Inserir", o desenho volta pro painel, que o coloca no slide. (Painel e janela trocam dados por `localStorage` de mesma origem — sem limite de mensagem do Office. Se o seu build particionar o `localStorage` do dialog, o fundo/entrega pode falhar; nesse caso usamos mensagens fragmentadas.)

## Estrutura
```
v3-powerpoint-addin/
├── manifest.xml                 ← manifesto do add-in (sideload)
├── assets/                      ← ícones
└── src/
    ├── engine/doodle-engine.js  ← motor de desenho extraído do v2.html
    ├── office/office-bridge.js  ← Office.js: ler slide + inserir imagem
    └── taskpane/                ← UI do painel (html/css/js)
```

## Como o desenho entra no slide
- **Fundo**: `getActiveSlide().getImageAsBase64()` (API **preview**) renderiza o slide atual; você desenha por cima. Se a API não existir no seu build do Office, cai num quadro vazio 16:9 (a inserção continua funcionando).
- **Inserção**: o traço vira um PNG transparente recortado na sua bounding box → `slide.shapes.addImage(base64)` → posicionado em pontos (slide 16:9 = 960×540 pt) onde foi desenhado. É **imagem**, não forma editável (limitação do Office.js).

## Rodar localmente (sideload)
Pré-requisito: Node.js.

1. **Certificado HTTPS confiável** (Office só carrega add-in por HTTPS):
   ```bash
   npx office-addin-dev-certs install
   ```
   Isso instala um cert/chave em `~/.office-addin-dev-certs/`.

2. **Servir a pasta do add-in na porta 3000 (HTTPS)** — sirva ESTA pasta como raiz:
   ```bash
   cd v3-powerpoint-addin
   npx http-server . -S \
     -C ~/.office-addin-dev-certs/localhost.crt \
     -K ~/.office-addin-dev-certs/localhost.key \
     -p 3000 -c-1
   ```
   Confirme que abre: `https://localhost:3000/src/taskpane/taskpane.html`

3. **Validar o manifesto** (opcional):
   ```bash
   npx office-addin-manifest validate manifest.xml
   ```

4. **Sideload no PowerPoint**:
   - **Windows/Mac (desktop)**: PowerPoint → **Inserir → Suplementos → Meus Suplementos → Carregar Meu Suplemento** → escolha `manifest.xml`.
   - **Web (powerpoint.office.com)**: **Inserir → Suplementos → Carregar Meu Suplemento** → `manifest.xml`.
   - O painel "CBA Studio" abre à direita.

5. **Usar**: abra um slide → **⟳ Atualizar slide** (carrega o fundo) → desenhe → **Inserir no slide**.

## Teste no navegador (sem PowerPoint)
Para validar só o desenho/engine, abra `src/taskpane/taskpane.html` num servidor estático qualquer. Em "Modo navegador":
- desenho, sliders, falhas, undo/redo funcionam normalmente;
- **Inserir no slide** baixa o PNG (em vez de inserir).

## Ressalvas conhecidas
- `getActiveSlide` / `getImageAsBase64` são **APIs preview** — exigem build recente do PowerPoint e podem variar entre Windows/Mac/Web. Há fallback para quadro vazio.
- Slide assumido como **16:9**. Suporte a 4:3 fica para depois.
- O painel é estreito; arraste a borda para alargar e desenhar com mais espaço.

## Roadmap (fora do MVP)
Outras texturas, imagem de referência, animação → GIF/WebM no slide, janela ampla (Office Dialog), 4:3, deploy centralizado/AppSource.
