# BUILD.md — ciclo completo de build e release do CBA Studio

Três artefatos saem deste repo:

| Artefato | O que é | Onde é servido |
|---|---|---|
| **Faixa** `BG-DoodleStudio.ppam` | Add-in VBA (aba "CBA Studio") | `doodle-studio-sigma.vercel.app/download/` |
| **Extensão** (painel) | Office.js taskpane (`v3-powerpoint-addin/src/`) | `doodle-studio-app.vercel.app` (manifest aponta pra cá) |
| **Landing** `index.html` + `download/` + `config.html` | Site + instaladores | `doodle-studio-sigma.vercel.app` |

**Dois projetos Vercel, de propósito** (deploys independentes):
- `doodle-studio` (raiz do repo, exclui `v3-powerpoint-addin/` via `.vercelignore`) → domínio **sigma** — landing, downloads, `version.json`, APIs Gemini.
- `doodle-studio-app` (pasta `v3-powerpoint-addin/`) → domínio **app** — o painel e o `install-mac.sh` referenciado pela landing.
- Nenhum é git-connected: publica-se com `vercel --prod` (ver §4).

---

## 1. Faixa (.ppam) — quando mexer em `assets/BG-DoodleStudio.bas` ou `ribbon/`

O `.bas` do repo é a FONTE, mas o VBA precisa ser compilado DENTRO do
PowerPoint (não existe compilador headless no Mac). O ciclo manual está
abaixo. *Pista de automação futura (testada em parte): o AppleScript do
PowerPoint Mac aceita `run VB macro macro name "..."` — um macro "builder"
dentro do .pptm poderia reimportar o .bas via VBE object model; falta validar
`Application.VBE` no Mac com NENHUM deck real aberto (macros mutantes rodam
na apresentação ativa!).* O ciclo:

1. **Ícones** (só se adicionou botão/ícone novo):
   ```bash
   python3 tools/gen-ribbon-icons.py         # gera ribbon/images/*.png
   ```
   e adicione a `<Relationship>` correspondente em
   `ribbon/_rels/customUI14.xml.rels`. O build (passo 3) valida XML×PNG×rels
   e **aborta com lista do que faltou** se algo não bater.

2. **Compilar o VBA no deck de build** `~/Downloads/BG-DoodleStudio.pptm`
   (se não existir: crie um .pptm qualquer com esse nome e siga o mesmo passo):
   - Abra o .pptm no PowerPoint → **Ativar Macros**.
   - `Tools > Macro > Visual Basic Editor`.
   - No Project pane: remova o módulo `BG_DoodleStudio` antigo
     (clique-direito → Remove… → **No** para não exportar).
   - Clique-direito no projeto → **Import File…** → escolha
     `v3-powerpoint-addin/assets/BG-DoodleStudio.bas`.
   - `Debug > Compile VBAProject` (não pode dar erro).
   - Volte ao PowerPoint (ícone no topo-esquerdo do VBE) e **Cmd+S**.

3. **Empacotar**:
   ```bash
   bash tools/build-ribbon-ppam.sh    # ~/Downloads/*.pptm -> ~/Downloads/*.ppam
   ```

4. **Instalar local + pacote**:
   ```bash
   cp ~/Downloads/BG-DoodleStudio.ppam \
     "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized/DoodleStudio/"
   cp ~/Downloads/BG-DoodleStudio.ppam download/BG-DoodleStudio.ppam
   (cd download && zip -X CBA-Studio-instalador.zip BG-DoodleStudio.ppam)
   ```

5. **Recarregar no PowerPoint aberto**: `Tools > PowerPoint Add-ins…` →
   desmarque **DoodleStudio** → OK → reabra o diálogo → marque de novo → OK.
   (Sem isso a sessão continua com o .ppam antigo em memória.)

6. **Smoke test ao vivo** nos botões que mudaram + `Padrões > Sobre`
   (a versão exibida deve ser a nova).

## 2. Extensão (painel)

Só editar `v3-powerpoint-addin/src/**` e deployar (§4) — o painel carrega
do site a cada abertura, ninguém precisa reinstalar. `node --check` em cada
.js editado antes.

## 3. Versão (release)

Uma versão única pro produto. Ao lançar, atualize **nos 5 lugares**:
1. `v3-powerpoint-addin/assets/BG-DoodleStudio.bas` → `CBA_VERSION` (+ ciclo §1!)
2. `v3-powerpoint-addin/src/taskpane/taskpane.js` → `const CBA_VERSION`
3. `v3-powerpoint-addin/manifest.xml` **e** `download/manifest.xml` → `<Version>`
4. `download/version.json` → `version`, `date`, `notes` (alimenta o aviso de
   update no painel e o instalador)
5. `index.html` (badge no eyebrow + notas "🔄 Já tem instalado?") e
   `download/LEIA-ME.txt` (título)

Atualizar = usuário roda o instalador de novo (o `install-mac.sh` sempre baixa
o `.ppam` mais recente e imprime a versão do `version.json`).

## 4. Deploy

```bash
# secret-scan antes de todo commit (a chave Gemini NUNCA pode entrar):
git diff | grep -iE "AIza|api[_-]?key|secret|token" || echo ok

git add -A && git commit && git push          # push: pedir autorização

vercel --prod --yes                           # na RAIZ  -> sigma (landing/downloads)
(cd v3-powerpoint-addin && vercel --prod --yes)  # -> app (painel)
```

## 5. Schema do `cba-config.txt` (Padrões da marca)

Gerado em `config.html` → colado em `Padrões > Aplicar config`. Formato
`chave=valor;chave=valor` (uma linha). Chaves:

| Chave | Exemplo | Significado |
|---|---|---|
| `fonte` | `Avenir Next` | fonte única da marca |
| `pal0…pal4` | `FD5E6D` | paleta: 0 rosa/magenta, 1 azul, 2 bege, 3 branco, 4 preto (hex sem #) |
| `radius` | `25` | raio padrão em px @ altura 1080 |
| `s_<id>` | `s_dsHero=120,1,0` | estilo: tamanho, negrito (0/1), papel de cor (0=rosa 1=azul) |
| `e_<size>` | `e_44=1.0` | entrelinha (múltiplo) por tamanho |

Persistido em `~/Library/Containers/com.microsoft.Powerpoint/Data/cba-config.txt`
(Mac) / `%APPDATA%` (Windows). `SetDefaults` no .bas é a verdade de fábrica.

## 6. Gotchas Mac (aprendidos a caro)

- `.ppam` não expõe macros na caixa de Macros — callbacks só via ribbon.
- `getActiveSlide()` / `shapes.addImage()` do Office.js falham no Mac —
  usar `getItemAt` + `setSelectedDataAsync` (já embutido no office-bridge).
- `window.confirm()` não funciona no webview do taskpane — confirmação em
  2 cliques.
- PowerPoint **rejeita silenciosamente** SlideWidth/Height se a proporção
  intermediária for extrema — altura primeiro, em passadas (FixPageSize).
- Decks REAIS abertos (ITAU, XP…): nunca descartar ao fechar — sempre Save.
