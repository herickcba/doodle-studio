# CBA Studio — Checklist de validação no Windows (~10 min)

O instalador Windows é **BETA** (o desenvolvimento é feito em Mac). Antes de
liberar para o time Windows, uma pessoa com PowerPoint 2016/365 no Windows
precisa rodar este checklist uma vez e reportar o resultado.

## Preparação
1. Baixe `BG-DoodleStudio.ppam` e `instalar-faixa-windows.bat` de
   https://doodle-studio-sigma.vercel.app (seção Windows) para a pasta Downloads.

## Instalação
2. [ ] Dê duplo-clique em `instalar-faixa-windows.bat`.
   - Se o PowerPoint estiver aberto, o instalador pede para fechar (não força).
   - Esperado no fim: `[OK] Instalado em: %APPDATA%\Microsoft\AddIns\BG-DoodleStudio.ppam`
3. [ ] Abra o PowerPoint → clique **Habilitar Macros** se perguntar.
4. [ ] A aba **CBA Studio** aparece na faixa.

## Funções-chave (crie uma apresentação em branco)
5. [ ] **Tipografia**: digite um texto no título, selecione, clique `Hero 120`
   → texto vira Avenir Next Bold 120pt rosa (se a fonte Avenir não existir no
   Windows, anote qual fonte apareceu — isso é esperado e será tratado à parte).
6. [ ] **Inserir → Rounded box** → aparece um retângulo arredondado bege/branco.
7. [ ] **Texto → Caps**: selecione um texto e clique → vira MAIÚSCULAS com
   espaçamento maior.
8. [ ] **Alinhar → Esquerda/Topo**: desenhe 2 retângulos, selecione os dois,
   clique cada botão → bordas esquerda/topo alinham.
9. [ ] **Formas → Rounded**: selecione um retângulo → cantos arredondam.
10. [ ] **Padrões → Aplicar config**: cole `pal0=00AA00` e OK → mensagem
    "Padroes aplicados e salvos." Selecione um texto e clique `H1 60` → fica
    **verde**. Feche e reabra o PowerPoint → `H1 60` continua verde
    (config persistiu em `%APPDATA%\CBAStudio\cba-config.txt`).
11. [ ] Para desfazer o teste: apague `%APPDATA%\CBAStudio\cba-config.txt`
    e reabra o PowerPoint.

## Extensão (painel lateral Doodle/ImaGen)
12. [ ] O sideload do manifest.xml no Windows usa uma pasta compartilhada de
    catálogo (diferente do Mac). Se o time Windows for usar o painel, validar
    à parte com as instruções da landing (seção Windows).

## Reportar
- Tudo passou → rollout Windows liberado.
- Algo falhou → mande print + o passo que falhou.
