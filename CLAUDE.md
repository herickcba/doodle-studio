# CBA Studio (Doodle Maker)

Toolkit de PowerPoint pra marca CBA B+G, em 3 partes:
1. **Faixa** — add-in VBA `.ppam` (aba "CBA Studio": tipografia, cores,
   alinhamento/âncoras, rounded, linhas-guia, auditoria, padrões).
   Fonte: `v3-powerpoint-addin/assets/BG-DoodleStudio.bas` + `ribbon/`.
2. **Extensão** — painel Office.js (abas Doodle / Imagem / Otimização).
   Fonte: `v3-powerpoint-addin/src/`.
3. **Landing + instalador** — `index.html`, `download/`, `install-mac.sh`
   (curl|bash), `config.html` (Brand Standards).

## Regras críticas

- **Segredos**: a chave Gemini (`AIzaSy…`) NUNCA pode ser commitada. Rode um
  secret-scan no diff antes de TODO commit (`git diff | grep -iE "AIza|api[_-]?key|secret|token"`).
- **Decks reais do usuário** (ITAU_IGA_BRANDING, 1.CBABG_XP, etc.): em
  diálogos de fechar/quit, SEMPRE Salvar — nunca descartar. Decks de teste
  descartáveis: Don't Save.
- **Versão**: uma só pro produto (hoje definida em `download/version.json`).
  Ao lançar, atualizar os 5 lugares listados em `tools/BUILD.md` §3.

## Build & deploy

Ciclo completo (incl. o passo manual do VBE pro .ppam, os DOIS projetos
Vercel e o schema do `cba-config.txt`): **`tools/BUILD.md`**.
Resumo: `.bas`/`ribbon/` mudou → compilar no VBE → `tools/build-ribbon-ppam.sh`
→ copiar pro caminho instalado + `download/` → re-zipar instalador → reload
add-in → smoke test. Extensão/landing: só editar e `vercel --prod` (raiz =
sigma/landing; `v3-powerpoint-addin/` = app/painel — NÃO são git-connected).

## Quirks Mac (não redescobrir)

Ver `tools/BUILD.md` §6 e a memória do projeto: Office.js sem
`getActiveSlide`/`addImage` confiáveis (usar `getItemAt` + `setSelectedDataAsync`),
sem `window.confirm` no taskpane, PageSetup rejeita proporção intermediária
extrema (altura primeiro), `.ppam` não mostra macros na caixa de Macros.

## Docs

- `tools/BUILD.md` — build/release/deploy (fonte da verdade)
- `HANDOFF.md` — arquitetura do motor de desenho (V1/V2)
- `FEATURES.md` — mapa de features
- `download/LEIA-ME.txt` — instruções pro usuário final
