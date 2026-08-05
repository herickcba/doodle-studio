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

2. **Compilar o VBA no deck de build** `~/Downloads/BG-DoodleStudio.pptm`.

   > **A cada release, comece de um arquivo-base LIMPO** (deck novo, `Cmd+N`,
   > salvo como `BG-DoodleStudio.pptm`, substituindo o anterior). Reaproveitar o
   > mesmo .pptm entre builds faz a tabela de nomes do projeto VBA acumular
   > identificadores de procedimentos já removidos — o `verify-ppam.sh` acusa
   > isso como "identificadores no P-code sem correspondência no fonte", que é
   > exatamente a assinatura de VBA stomping numa auditoria externa. No .ppam
   > v1.5.0 havia 39 resíduos assim (`ScaleAllSlides`, `C_AZUL`, `temBackup`…),
   > todos de código deletado de verdade. Ver `docs/SECURITY-REVIEW-v1.5.0B.md` §5.

   - Abra o .pptm no PowerPoint → **Ativar Macros**.
   - `Tools > Macro > Visual Basic Editor`.
   - No Project pane: remova o módulo `BG_DoodleStudio` antigo
     (clique-direito → Remove… → **No** para não exportar).
   - Clique-direito no projeto → **Import File…** → escolha
     `v3-powerpoint-addin/assets/BG-DoodleStudio.bas`.
   - `Debug > Compile VBAProject` (não pode dar erro).
   - Volte ao PowerPoint (ícone no topo-esquerdo do VBE) e **Cmd+S**.

3. **Empacotar e verificar**:
   ```bash
   bash tools/build-ribbon-ppam.sh    # ~/Downloads/*.pptm -> ~/Downloads/*.ppam
   bash tools/verify-ppam.sh          # gate de integridade + build-manifest.json
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

## 4.1 Checklist de publicação (gates obrigatórios)

Ordem obrigatória. **Qualquer gate que falhar interrompe a publicação** — é o
que a análise de risco do sócio exige (resposta completa em
`docs/SECURITY-REVIEW-v1.5.0B.md`).

| # | Gate | Comando | Falha significa |
|---|---|---|---|
| 1 | Segredos | `git diff \| grep -iE "AIza\|api[_-]?key\|secret\|token"` | chave prestes a vazar — **pare** |
| 2 | Análise estática do VBA | `bash tools/vba-static-scan.sh` | termo proibido sem justificativa na allowlist |
| 3 | Compilação | `Debug > Compile VBAProject` no VBE (§1, arquivo-base **limpo**) | erro de compilação |
| 4 | Empacotamento | `bash tools/build-ribbon-ppam.sh` | ícone sem PNG ou sem `<Relationship>` |
| 5 | Integridade + manifesto | `bash tools/verify-ppam.sh` | fonte do binário ≠ `.bas` do repo, callback sem procedimento, ou repo sujo |
| 5b | Documentação em dia | `python3 tools/design-system/check-tokens.py` | mexeu num estilo, cor ou constante e esqueceu o `design.md` / `tokens.py` |
| 6 | Smoke test ao vivo | percorrer os grupos da faixa **num deck descartável** | regressão funcional |
| 7 | Commit + push | `git add -A && git commit && git push` | — |
| 8 | Deploy | `vercel --prod --yes` (§4) | — |

Notas:

- O gate 5 **exige árvore limpa**: o `build-manifest.json` declara um commit, e
  com o repo sujo esse vínculo seria falso. Para teste local, `--allow-dirty`
  (o manifesto sai marcado como `sourceDirty: true`).
- Rode o gate 5 **uma vez** e commite o manifesto junto no commit do release. Ele
  sempre aponta para o commit anterior ao que o contém (é a natureza da coisa:
  o hash do commit não existe antes do commit). Rodar o verify de novo depois só
  reescreve `buildDate`/`sourceCommit` — não persiga isso.
- Os "fantasmas de P-code" do gate 5 **avisam, não bloqueiam** (existe `--strict`
  para bloquear, mas não use na rotina: todo build produz ~5 tokens de lixo
  binário aleatório, que mudariam a cada release). O que importa é a **forma**
  do token: nome que parece identificador (`ScaleAllSlides`, `C_AZUL`) = resíduo
  real, investigue; sequência curta sem sentido (`sMZfMz`) = ruído, ignore.
- O gate 5b compara o `SetDefaults` do `.bas` com `tools/design-system/tokens.py`
  (estilos, paleta, entrelinhas, raio, âncora) e com os rótulos do
  `customUI14.xml`. Se falhar, atualize `tokens.py` + `design.md` e regenere o
  deck: `python3 tools/design-system/build_deck.py`.
- **O projeto VBA não é assinado digitalmente.** Não é um passo esquecido: é uma
  decisão registrada (o VBE do macOS não assina e não há certificado corporativo).
  Justificativa em `docs/SECURITY-REVIEW-v1.5.0B.md` §6. Se um dia houver
  certificado, assine **antes** do gate 5 — a assinatura altera o arquivo e o
  hash tem de ser calculado depois dela.

## 5. Schema do `cba-config.txt` (Padrões da marca)

Gerado em `config.html` → colado em `Padrões > Aplicar config`. Formato
`chave=valor;chave=valor` (uma linha). Chaves:

| Chave | Exemplo | Significado |
|---|---|---|
| `fonte` | `fonte=Avenir Next` | fonte única da marca |
| `pal0…pal4` | `pal0=FD5E6D` | paleta: 0 rosa/magenta, 1 azul, 2 bege, 3 branco, 4 preto (hex sem #) |
| `radiusPx` | `radiusPx=25` | raio padrão em px @ altura 1080 |
| `s_<id>` | `s_dsHero=120\|1\|0` | estilo: tamanho **\|** negrito (0/1) **\|** papel de cor (0=rosa 1=azul) |
| `ent_<size>` | `ent_44=1.0` | entrelinha (múltiplo) por tamanho |

> Os nomes acima são os que `ApplyKV` (no `.bas`) realmente reconhece, e são os
> que `config.html` gera — conferido nos dois lados. **Chave desconhecida é
> ignorada em silêncio**, sem erro: se você montar uma config à mão e ela não
> surtir efeito, quase sempre é nome de chave errado. Note o separador do
> `s_<id>`: barra vertical, não vírgula.

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
