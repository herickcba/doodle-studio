# CBA Studio — Handoff

> **Para a próxima sessão:** leia este documento inteiro antes de tocar em
> qualquer arquivo. Ele reflete o estado real do projeto em **21/07/2026**,
> versão **v1.5.0**, último commit `d568aa9`.
> Complementos: `tools/BUILD.md` (build/deploy, fonte da verdade do processo),
> `CLAUDE.md` (regras críticas), `FEATURES.md` (mapa de features).
> O handoff antigo (ferramenta web V1/V2, hoje legado) está em
> `HANDOFF-v1-v2-legacy.md`.

---

## 1. O que é o produto

Toolkit de PowerPoint para a marca **CBA B+G**, distribuído em **3 partes**
que compartilham uma versão única:

| Parte | O que é | Fonte | Servido em |
|---|---|---|---|
| **Faixa** | Add-in VBA `.ppam` — aba "CBA Studio" no ribbon | `v3-powerpoint-addin/assets/BG-DoodleStudio.bas` + `ribbon/` | `doodle-studio-sigma.vercel.app/download/` |
| **Extensão** | Painel lateral Office.js (abas Doodle / Imagem / Otimização) | `v3-powerpoint-addin/src/` | `doodle-studio-app.vercel.app` |
| **Landing** | Site de instalação + página de padrões da marca | `index.html`, `config.html`, `download/` | `doodle-studio-sigma.vercel.app` |

**Usuários:** o próprio usuário (Herick) e o time/sócios da CBA B+G. Rodam em
**Mac** (principal) e Windows (beta, não validado a fundo).

---

## 2. ⚠️ Regras críticas (nunca violar)

1. **Segredos:** a chave Gemini (`AIzaSy…`) NUNCA pode ser commitada.
   Rode `git diff | grep -iE "AIza|api[_-]?key|secret|token"` antes de TODO commit.
   A chave vive só no ambiente do Vercel.

2. **Decks reais do usuário:** ele deixa decks de trabalho abertos no PowerPoint
   (ex.: `ITAU_IGA_BRANDING`, `1.CBABG_XP`, `RGM2_Conceito`, `2 [Autosaved]`).
   - Em diálogos de fechar/quit desses decks: **SEMPRE Salvar**, nunca descartar.
   - **NUNCA rode macros que escrevem** (FixPageSize, Padronizar*, etc.) pelo
     Immediate window sem antes conferir `ActivePresentation.Name` — o VBA age
     na apresentação ATIVA, que pode ser um deck real.
   - Para testar: crie um deck descartável (`Cmd+N`), ative-o explicitamente
     (`Presentations("Presentation4").Windows(1).Activate`), teste, feche com
     "Don't Save". Decks de teste = descartáveis; decks reais = intocáveis.

3. **Git e Vercel são separados.** `git push` não publica nada. Publicar exige
   `vercel --prod --yes` (ver §5). Os dois projetos Vercel NÃO são git-connected.

4. **Push para `master`** pode ser bloqueado pelo classificador de segurança
   (push direto na branch padrão). Se bloquear, peça autorização ao usuário.

---

## 3. Arquitetura

### 3.1 Faixa (VBA)
- **Arquivo único:** `BG-DoodleStudio.bas` (~2.250 linhas, ~47 subs/functions públicas).
- **Ribbon:** `ribbon/customUI14.xml` (RibbonX) + `ribbon/_rels/customUI14.xml.rels`
  (todo ícone precisa de entry no rels!) + `ribbon/images/*.png` (83 ícones,
  gerados por `tools/gen-ribbon-icons.py`).
- **Grupos da aba:** Inserir · Tipografia (11 estilos) · Entrelinha · Texto ·
  Fonte · Preench. · Contorno · Formas · Alinhar · Auditoria · Padrões.

**Constantes no topo do módulo** (nunca hard-code de novo):
`CBA_VERSION`, `PT_PER_CM` (28.3465), `ANCHOR_DEFAULT_CM` (1.27),
`GUIDE_MARGIN_CM` (3.15), `MAX_DEPTH` (12), `STYLE_INSERT` ("dsH5"),
`SAMPLE_MAX` (24).
> VBA exige que TODA declaração de nível de módulo (`Const`/`Dim`/`Private`)
> fique no topo, antes de qualquer procedimento. Const no meio do arquivo =
> erro de compilação "Only comments may appear after End Sub".

**Escopo das operações** (decisão do usuário, leva 6):
- `TargetSlides()` é a regra única para Padronizar tipografia/entrelinha/cores:
  slides selecionados no painel → slide dos objetos selecionados → slide atual.
  **Nunca a apresentação inteira** (era destrutivo demais).
- Rounded/Unround também operam por seleção.

### 3.2 Extensão (Office.js)
- `src/office/office-bridge.js` (~1.200 linhas) — "god module": leitura do .pptx
  por slices de ZIP, parsing XML, inserção de shapes, otimização de imagem,
  diálogos. **Refatorar em 3 módulos está fora de escopo** (risco de regressão no Mac).
- `src/engine/doodle-engine.js` — motor de desenho vetorial (vetorização
  gaussiana + RDP + Chaikin, texturas, encoder GIF89a próprio).
- `src/taskpane/` — UI; `src/dialog/` — tela grande e editor de imagem;
  `src/modules/` — image-gen (Gemini) e img-audit (otimização);
  `src/shared/presets.js` — biblioteca de doodles embarcada (`window.DoodlePresets`).

### 3.3 Landing
- `index.html` e `config.html` usam o **design system CBA/RGM**
  (mesmos tokens dos protótipos em `~/Desktop/CBA/RGM/RGM 2.0/prototipos/assets/cba.css`).
- Tokens: `--blue:#436AE1`, `--navy:#264BB9`, `--coral:#FC5E6D`,
  `--coral-btn:#FB3E50`, `--ice:#E4EBFF`. Fonte Avenir Next, **só pesos 400 e 700**.
- Logo oficial em `assets/logo-cba.png`.
- `v1.html` / `v2.html` são a ferramenta web antiga (mantidas; v1 é referência histórica).

---

## 4. Ciclo de build da faixa (o passo manual que dói)

O `.bas` do repo é a FONTE, mas **o VBA precisa ser compilado dentro do
PowerPoint** — não há compilador headless no Mac. Ciclo completo em
`tools/BUILD.md`; resumo:

1. Se mexeu em ícone: `python3 tools/gen-ribbon-icons.py` **+ adicionar
   `<Relationship>` no `.rels`** (o build valida e aborta se faltar).
2. `open ~/Downloads/BG-DoodleStudio.pptm` → Ativar Macros.
3. `Tools > Macro > Visual Basic Editor`.
4. No projeto **BG-DoodleStudio** (cuidado: há outros projetos na lista!):
   clique-direito no módulo `BG_DoodleStudio` → Remove → **No** (não exportar).
5. Clique-direito no projeto → **Import File…** → o `.bas` do repo.
   (No diálogo de arquivo, `Cmd+Shift+G` e colar o caminho é mais confiável
   que clicar na lista.)
6. Focar a janela do VBE (clicar na barra de título) → `Debug > Compile VBAProject`.
   Sem dialog de erro = compilou.
7. Salvar direcionado pelo Immediate window (evita salvar o deck errado):
   `Presentations("BG-DoodleStudio.pptm").Save`
8. `bash tools/build-ribbon-ppam.sh` → gera `~/Downloads/BG-DoodleStudio.ppam`.
9. Instalar + empacotar:
   ```bash
   cp ~/Downloads/BG-DoodleStudio.ppam \
     "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized/DoodleStudio/"
   cp ~/Downloads/BG-DoodleStudio.ppam download/BG-DoodleStudio.ppam
   (cd download && zip -X CBA-Studio-instalador.zip BG-DoodleStudio.ppam)
   ```
10. Recarregar no PowerPoint: `Tools > PowerPoint Add-ins…` → desmarcar
    DoodleStudio → OK → reabrir → marcar → OK. **Sem isso a sessão continua
    com o .ppam antigo em memória.**
11. Conferir que o código novo entrou:
    ```bash
    cd /tmp && rm -rf pk && mkdir pk && cd pk && unzip -q <ppam> \
      && strings ppt/vbaProject.bin | grep -i "MinhaFuncaoNova"
    ```

**Extensão e landing:** só editar e deployar — nada de compilar.

---

## 5. Deploy (dois projetos Vercel)

```bash
# secret-scan primeiro!
git add -A && git commit -m "..." && git push origin master

vercel --prod --yes                              # RAIZ  → sigma (landing, downloads, config, APIs)
(cd v3-powerpoint-addin && vercel --prod --yes)   # → app (painel/extensão + install-mac.sh)
```
Regra prática: mexeu em `index.html`/`config.html`/`download/` → deploy da raiz.
Mexeu em `v3-powerpoint-addin/src/` ou `manifest.xml` → deploy da pasta do addin.
Mexeu no `.ppam` → deploy da raiz (é servido de `download/`).

**Versão** (`v1.5.0`) vive em 5 lugares — atualizar todos ao lançar:
`CBA_VERSION` no `.bas`, `CBA_VERSION` em `taskpane.js`, `<Version>` nos dois
`manifest.xml`, `download/version.json`, e o badge/textos em `index.html`.

---

## 6. Quirks do Mac (não redescobrir — custaram caro)

**PowerPoint / VBA**
- `.ppam` **não** expõe macros na caixa de Macros. Callbacks só via ribbon.
- **PageSetup rejeita silenciosamente** mudar uma dimensão se a proporção
  intermediária ficar extrema (1583×540 = 2,9:1 é bloqueado). Solução:
  **altura primeiro**, em 4 passadas.
- O PowerPoint **escala a geometria** ao mudar o tamanho do slide, mas o corpo
  da fonte **não aparece alterado em `Font.Size`** (ele usa autoajuste interno).
  **Nunca multiplique `Font.Size`** — foi o que quebrou layouts inteiros.
- Placeholders/objetos especiais retornam dimensões-sentinela absurdas que
  estouram na multiplicação (**erro 6 Overflow**). Sempre validar faixa antes
  de gravar e proteger com `On Error`.
- `On Error Resume Next` é local ao procedimento — o erro pode escapar no
  código do CHAMADOR. Envolva blocos de risco em função própria com handler.
- `View.PasteSpecial` não existe; usar `Application.CommandBars.ExecuteMso`
  late-bound (`Dim cb As Object`) — padrão validado em `PasteTextOnly` e `CropPicture`.
- Strings do VBA: **evite acentos e travessões (—)** em MsgBox/InputBox. O VBE
  do Mac interpreta o UTF-8 do `.bas` como MacRoman e vira mojibake (`‚Äî`).
  Use hífen simples.

**Office.js**
- `getActiveSlide()` e `shapes.addImage()` **não são confiáveis**; usar
  `slides.getItemAt(index)` + `setSelectedDataAsync`.
- `shapes.load('items')` → `sync` → `item.load('name')` → `sync` (dois passos;
  o atalho `'items/name'` não popula).
- `window.confirm()` **não funciona** no webview do taskpane → confirmação em 2 cliques.
- Ler o .pptx custa caro: há cache de fatias compartilhado (TTL ~2 min) e LRU
  de 5 imagens decodificadas em `office-bridge.js`.

**GIF (motor de doodle)**
- GIF só tem transparência **binária**; a textura de giz é semi-transparente.
  Tentativas de "melhorar" (dithering Bayer, matte branco) foram **rejeitadas
  pelo usuário** e revertidas. O render é o original (CAP 720 + corte alpha 128),
  com o traço gerado a **70% do peso** (`GIF_STROKE_WEIGHT`) para ficar arejado
  como o pincel. **Não mexa nisso sem comparativo visual aprovado.**

**Ferramentas / computer-use**
- O VBE às vezes rouba/perde o foco do menu: se o menu bar mostrar os menus do
  PowerPoint (Arrange, Slide Show) em vez dos do VBE (Debug, Run), clique na
  janela do VBE antes.
- Diálogos do sistema podem bloquear cliques ("SecurityAgent"). Teclado
  (`Return`, `Tab`) costuma funcionar quando o clique não.

---

## 7. Estado atual e decisões recentes

**Últimas levas entregues (todas commitadas, publicadas e testadas ao vivo):**

- **Leva 4** — versionamento v1.5.0 (`version.json`, botão "Sobre" na faixa,
  aviso de update no painel), constantes VBA, âncoras esq/topo unificadas,
  cache de .pptx, validação de ícones no build, `tools/BUILD.md` + `CLAUDE.md`,
  cache HTTP dos estáticos (abertura do painel deixou de fazer 13 round-trips).
- **Leva 5** — GIF vai pra biblioteca e o slide recebe **sempre o PNG estático**;
  badges PNG/GIF nos cards; preview com fundo azul para traço branco/cinza;
  presets embarcados + botão ⧉ que copia o JSON do doodle.
- **Leva 6** — Padronizar tipografia/entrelinha/cores passam a agir **só nos
  slides selecionados**.
- **Leva 7** — Page Size refeito (confia no resize nativo, detecta por
  amostragem, oferece ajuste só se o PPT não mexeu, **nunca gera erro**);
  caixa de texto e rounded box nascem no estilo **Texto 24**; botão **Crop**;
  landing e `config.html` redesenhadas no design system CBA/RGM.
- **Último commit** — botões **Esquerda/Centro** (alinhamento de parágrafo) no
  grupo Texto, à direita do Negrito.

**Fora de escopo (decidido, não refazer sem pedir):**
- Quebrar `office-bridge.js` em 3 módulos.
- Git LFS / GitHub Releases para os binários (repo tem ~11 MB).
- Consolidar os dois domínios Vercel (funciona; documentar em vez de mexer).
- Testes automatizados de Office.js.
- Refatorar os ~10 loops de seleção do VBA num helper genérico (sem delegates
  em VBA vira gambiarra; custo > ganho).

**Pendências conhecidas:**
- `presets.js` está **vazio** — o usuário vai desenhar os doodles e mandar os
  JSONs (copiados pelo botão ⧉) para colar lá.
- Instalador Windows continua **beta** (checklist em `download/WINDOWS-TEST.md`).
- O caminho "PowerPoint não escala nada" do Page Size não pôde ser reproduzido
  em deck sintético — se o usuário relatar de novo, pedir para ele contar o que
  a mensagem final disse (ela informa qual caminho foi usado).

---

## 8. Como o usuário trabalha (preferências observadas)

- Fala **português**; responda em português.
- Quer **direto ao ponto** — pediu explicitamente menos preâmbulo.
- Prefere **UI direta e integrada**, poucos cards, cinza neutro + azul.
- Em mudança de render/visual: **mostre um comparativo e peça aprovação antes
  de publicar** (aprendido no episódio da textura do GIF).
- Fluxo padrão ao pedir uma feature: implementar → build → **testar ao vivo no
  PowerPoint** → commit → push → deploy. Ele valoriza a verificação real, não
  só "o código está certo".
- Ao terminar, ele costuma pedir o resumo do que mudou **em termos de benefício
  para o usuário**, não em termos técnicos.
