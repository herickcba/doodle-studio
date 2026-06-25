# Product

## Register

product

## Users

Designers, ilustradores e criadores de conteúdo que querem desenhar traços vetoriais à mão e animá-los progressivamente (revelação tipo "lousa") para exportar como PNG, SVG, GIF ou WebM. Usam em desktop, geralmente com mouse ou caneta/tablet, num fluxo focado: desenhar → ajustar textura/suavização → organizar a timeline → exportar. Contexto de trabalho concentrado, sessão única, sem login.

## Product Purpose

Ferramenta de desenho vetorial animado 100% offline, em arquivo único HTML/CSS/JS (sem build, sem dependências, sem CDN). Existe para tornar simples criar animações de "desenho à mão" com controle preciso (curvas Bezier, texturas de traço, falhas de tinta, timeline arrastável) e exportar em múltiplos formatos sem ferramentas pesadas. Sucesso = o usuário desenha, anima e exporta um resultado limpo rapidamente, com o canvas no centro e os controles fora do caminho. Há duas versões: `index.html` (V1, estável, intocável) e `v2.html` (versão ativa, redesenhada).

## Brand Personality

Minimalista, técnico, preciso. Voz objetiva e sem firulas; a interface se apaga para o trabalho aparecer. Estética de ferramenta profissional (mais perto de um editor/CAD leve do que de um app lúdico), com identidade carregada por tipografia, neutros e um único accent azul — não por decoração.

## Anti-references

- **SaaS genérico / "cara de IA"**: gradientes, glassmorphism, grids de cards idênticos, eyebrows minúsculos em maiúsculas, fundo creme/bege padrão.
- **Interface lotada de cards/boxes e menus aninhados** — o oposto da UI integrada e plana adotada (seções separadas por divisores, não containers empilhados).
- **Ferramenta infantil** — nada de emojis, cores berrantes ou ar de brinquedo. Ícones outline, paleta contida.

## Design Principles

- **O canvas é o protagonista.** Controles densos e fora do caminho; é possível ocultar o painel de animação/timeline para desenhar em tela cheia.
- **Integrar, não empilhar.** Seções planas com divisores em vez de cards aninhados; cada controle aparece uma vez, no lugar onde faz sentido.
- **Precisão visível.** Mostrar estado (pincel ativo, status, resolução, contagem de traços) de forma compacta e legível; números tabulares.
- **Contenção cromática.** Neutros + um accent azul; cor sinaliza ação/estado, não decora.
- **Offline e sem peso.** Arquivo único, zero dependências; nada que dependa de rede.

## Accessibility & Inclusion

Alvo WCAG 2.1 AA: contraste de texto ≥4.5:1 (≥3:1 para texto grande), foco visível e operável por teclado nos controles principais (botões, sliders, selects, modais). O desenho em si é por ponteiro (mouse/caneta/touch) — inerente à ferramenta —, mas toda a cadeia de UI ao redor deve ser navegável por teclado. Respeitar `prefers-reduced-motion` quando houver transições.
