# Design System CBA B+G — regras

> **Fonte da verdade:** o `SetDefaults` de
> [`BG-DoodleStudio.bas`](v3-powerpoint-addin/assets/BG-DoodleStudio.bas). Este
> documento e o `tools/design-system/tokens.py` são espelhos dele. Quando
> divergirem, o código vence — e o gate
> [`check-tokens.py`](tools/design-system/check-tokens.py) quebra o build até
> alguém acertar os três.
>
> Escrito para atravessar mídia: todo token vem em **pt** (PowerPoint), **px**
> (web, canvas de 1080 de altura) e **cm** (impresso e a régua do usuário).

---

## 1. A ideia em uma página

O sistema tem **dois eixos** e o resto é consequência.

**Eixo 1 — a escala tipográfica é fechada.** Treze estilos, nada entre eles. Não
existe "um pouquinho maior": se o corpo não está na tabela, não é do sistema.
Cada estilo já traz peso, cor e entrelinha juntos — escolher "Título 60" decide
as quatro coisas de uma vez, e é por isso que a faixa aplica tudo num clique.

**Eixo 2 — o tipo só tem duas cores.** Rosa e azul. Bege e branco são fundo;
preto é reserva e não entra em texto da marca. Toda decisão de cor de texto se
reduz a "qual dos dois papéis" — e a exceção (branco) não é escolha, é a regra
de contraste agindo sozinha.

Tudo o mais — raio, espaçamento, margem — deriva de um módulo de **20pt**.

---

## 2. Página

| | pt | px @1080 | cm |
|---|---|---|---|
| Largura | 1583,13 | 1920 | 55,8 |
| Altura | 890,63 | 1080 | 31,4 |

16:9 customizado — **não** é o 16:9 padrão do PowerPoint (1280×720pt). Decks em
outro formato entram distorcidos; o botão **Page Size** da faixa converte.

A ponte entre mídias é a altura: **1080px ≡ 890,63pt**, logo `1pt = 1,2126px`.
É a razão que o `radiusPx` da configuração usa e o que permite o mesmo token
descrever um slide e uma tela.

---

## 3. Grid e espaçamento

| Token | pt | px | cm | Papel |
|---|---|---|---|---|
| Margem de segurança | 76 | 92 | 2,68 | Onde o conteúdo **não** entra |
| Módulo de espaçamento | 20 | 24 | 0,71 | Toda distância é múltiplo dele |
| Âncora de encaixe | 36 | 44 | 1,27 | Alinhamento entre objetos |
| Raio | 20,62 | 25 | 0,73 | Cantos arredondados |

**Área útil: 1431 × 765pt.**

**Espaçamento** anda em múltiplos do módulo: 20 · 40 · 60 · 80pt. Distância que
não é múltiplo de 20 é acidente, não decisão.

**Raio** é um valor **visual constante**, não uma proporção: um card grande e um
botão pequeno têm o mesmo arredondamento de 20,62pt. Em OOXML isso vira
`adj = (raio / (menor_lado / 2)) × 50000`. Pills usam `adj = 50000`.

**A âncora não é a margem.** São papéis diferentes: a margem delimita a área
útil da página; a âncora é o passo de encaixe ao alinhar objetos entre si. A
faixa oferece 0,5 · 1 · 1,27 · 1,5 · 2 · 2,5 · 3 · 4 · 5 cm.

---

## 4. Paleta

| # | Nome | Hex | RGB | Papel |
|---|---|---|---|---|
| `pal0` | Rosa | `#FD5E6D` | 253, 94, 109 | **Cor de tipo 1.** Statements, títulos, chapéus, números |
| `pal1` | Azul | `#436AE1` | 67, 106, 225 | **Cor de tipo 2.** Corpo, tópicos, legendas, manchetes |
| `pal2` | Bege | `#EEECE6` | 238, 236, 230 | Fundo de conteúdo padrão. Nunca é cor de tipo |
| `pal3` | Branco | `#FFFFFF` | 255, 255, 255 | Fundo claro **e** texto sobre fundo saturado |
| `pal4` | Preto | `#000000` | 0, 0, 0 | Reserva. Não usar em tipo da marca |

Os índices importam: a configuração de marca (`cba-config.txt`) referencia por
número, não por nome.

### Fundos permitidos

Bege (padrão), branco, azul e rosa. Sobre azul e rosa o texto é branco — não por
escolha, mas pela regra de contraste abaixo.

---

## 5. Escala tipográfica

Uma família: **Avenir Next**, Regular e Bold. Nada mais.

| Estilo | pt | px | Peso | Cor | Entrelinha | Quando usar |
|---|---|---|---|---|---|---|
| Big Number 250 | 250 | 303 | Bold | rosa | 1,00× | O número que carrega a página. Um por slide |
| Statement 120 | 120 | 146 | Bold | rosa | 0,80× | A frase-manifesto |
| Manchete 80 | 80 | 97 | Bold | azul | 0,90× | Abertura de capítulo |
| Título 60 | 60 | 73 | Bold | rosa | 0,90× | Título de página |
| Destaque 44 | 44 | 53 | Regular | azul | 1,15× | Corpo de destaque, manifesto |
| Subtítulo 34 | 34 | 41 | Bold | rosa | 0,95× | Subtítulo, descritivo |
| Tópico 28 | 28 | 34 | Bold | azul | 1,00× | Título de tópico ou card |
| Texto 24 | 24 | 29 | Regular | azul | 1,00× | Texto corrido — **o padrão** |
| Apoio 20 | 20 | 24 | Regular | azul | 1,30× | Apoio em pilares e colunas |
| Chapéu 18 | 18 | 22 | Bold | rosa | 1,00× | Chapéu acima do título |
| Texto 15 | 15 | 18 | Regular | azul | 1,30× | Texto corrido em slide denso |
| Legenda 12 | 12 | 15 | Bold | azul | 1,30× | Legendas, notas, fontes |
| CAPS 12 | 12 | 15 | Bold | azul | 1,00× | Rótulo em caixa alta |

### A lógica da entrelinha

Não é arbitrária: **quanto maior o corpo, mais apertada a linha.** Em 250 e 120pt
o espaço entre linhas já é enorme em valor absoluto, então o multiplicador
comprime (0,80×) para o bloco ler como uma massa só. Em 12 e 15pt acontece o
inverso — a linha precisa de ar para o olho achar a próxima (1,30×).

A entrelinha é **exata** (múltiplo do corpo), nunca "simples" ou "duplo", que
variam com a fonte e quebram a consistência entre máquinas.

**Legenda 12 e CAPS 12 dividem o corpo e divergem na entrelinha** (1,30× contra
1,00×). Por isso a entrelinha pertence ao **estilo**, não ao tamanho — uma tabela
indexada por corpo não conseguiria descrever os dois.

> Existe uma segunda tabela, indexada por corpo, que alimenta os botões de
> Entrelinha e o Padronizar. Ela age em texto **sem estilo associado** e é só um
> padrão por tamanho. Onde os dois discordam, o estilo manda.

---

## 6. Regras de composição

### Contraste automático

**Se a cor do texto for exatamente a cor do fundo, o texto vira branco.**

O fundo relevante é, nesta ordem: o preenchimento da própria forma; senão o
preenchimento do grupo que a contém; senão o fundo do slide — e "fundo do slide"
inclui uma forma sólida que cubra a página inteira, que é como a maioria dos
decks reais é montada.

Vale na criação e em toda aplicação de estilo. Não é sugestão: texto rosa sobre
fundo rosa é invisível, e o sistema não deixa acontecer.

### Ponto final colorido

O **Statement 120** pinta o ponto final de azul — só o ponto, só se a frase
terminar em ".". Sobre fundo azul, o texto vira branco e o ponto vira rosa.

É o único estilo com comportamento próprio. Os outros doze são declarativos.

### Caixa alta

O **CAPS 12** aplica caixa alta e **3pt de espaçamento entre letras**. Caixa alta
sem espaçamento fecha demais em Avenir Next Bold.

É reversível: aplicar outro estilo por cima desfaz a caixa alta e zera o
espaçamento. Nenhum estilo deixa resíduo.

---

## 7. Anti-vícios

O sistema **nunca** emite:

- Linha decorativa, barra vertical de destaque, divisória improvisada
- Gradiente, sombra, brilho, contorno em texto
- Cor fora dos cinco slots da paleta
- Corpo fora dos treze da escala
- Entrelinha fora da tabela do estilo
- Fonte que não seja Avenir Next
- Ícone inventado

Quando bater a vontade de "colocar uma linha azul fina para separar", a resposta
é hierarquia: espaço em branco (múltiplo de 20pt), mudança de corpo, ou um bloco
de cor. Nunca um traço.

---

## 8. Migração

Corpos que existiram em versões anteriores e saíram da escala fechada:

| Saiu | Vai para | Observação |
|---|---|---|
| 55pt | **60pt** | Título |
| 36pt / 32pt | **34pt** | Subtítulo |
| 22pt | **24pt** | Texto |
| 16pt / 14pt | **15pt** | Texto pequeno |
| Rótulo 60 | **Título 60** | Eram formatação idêntica; o botão foi removido |

O botão **Padronizar tipografia** faz esse arredondamento sozinho, encaixando
cada corpo no vizinho mais próximo da escala.

> Atenção: com 250pt na escala, texto acima de ~185pt agora arredonda para 250,
> não para 120.

---

## 9. Fora do PowerPoint

Os tokens viajam; o que muda é a unidade.

**Web.** Use `px` da tabela (base 1080). A entrelinha vira `line-height` sem
unidade — o multiplicador é o mesmo número. O raio de 25px é literal. O módulo
de espaçamento vira 24px. Se o layout não for 1920×1080, escale por proporção,
nunca reescolha corpos: a hierarquia mora nas razões entre os treze, não nos
valores absolutos.

**Impresso.** Use `cm`. A página é 55,8 × 31,4cm — em A4 paisagem (29,7 × 21cm)
o fator é 0,53.

**Social.** Recorte a área útil, nunca a página inteira: a margem de 76pt existe
para o conteúdo sobreviver a qualquer recorte.

**Regra que não muda em nenhuma mídia:** duas cores de tipo, escala fechada,
entrelinha do estilo, espaçamento em múltiplos do módulo, e o contraste
automático.

---

## 10. Dívidas conhecidas

Divergências reais entre o que este documento diz e o que existe hoje. Estão
aqui para não virarem surpresa:

1. **Margem × linhas-guia.** A margem documentada é 76pt (2,68cm), mas a faixa
   desenha as linhas-guia em 3,15cm (89,3pt). Corrigir exige mudar
   `GUIDE_MARGIN_CM` no `.bas` e um ciclo de build. O `check-tokens.py` avisa a
   cada execução.
2. **O rosa em outras mídias.** A landing (`--coral: #FC5E6D`) e a skill
   `cba-bg-design-system-v3` usam `FC5E6D`; o padrão é **`FD5E6D`**. Um dígito de
   diferença, invisível a olho nu e detectável em auditoria.
3. **Corpos herdados do v3.** 55, 36, 32, 22, 16 e 14pt aparecem em decks
   antigos. A tabela de migração acima resolve.

---

## 11. Onde cada coisa vive

| Arquivo | Papel |
|---|---|
| `v3-powerpoint-addin/assets/BG-DoodleStudio.bas` | **A fonte da verdade** (`SetDefaults`) |
| `tools/design-system/tokens.py` | Espelho em Python, para gerar e validar |
| `tools/design-system/check-tokens.py` | Gate que compara os dois |
| `config.html` | Página de Brand Standards — muda os tokens sem recompilar |
| `docs/CBA-Studio-Design-System.pptx` | Este documento em slides |
| `design.md` | Este arquivo |
