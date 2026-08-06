# Design System CBA B+G: regras · 2.0

> **Fonte da verdade:** o `SetDefaults` de
> [`BG-DoodleStudio.bas`](v3-powerpoint-addin/assets/BG-DoodleStudio.bas). Este
> documento e o `tools/design-system/tokens.py` são espelhos dele. Quando
> divergirem, o código vence, e o gate
> [`check-tokens.py`](tools/design-system/check-tokens.py) quebra o build até
> alguém acertar os três.
>
> Escrito para atravessar mídia: todo token vem em **pt** (PowerPoint), **px**
> (web, canvas de 1080 de altura) e **cm** (impresso e a régua do usuário).

---

## 1. A ideia em uma página

O sistema tem **dois eixos** e o resto é consequência.

**Eixo 1: a escala tipográfica é fechada.** Treze estilos, nada entre eles. Não
existe "um pouquinho maior": se o corpo não está na tabela, não é do sistema.
Cada estilo já traz peso, cor e entrelinha juntos. Escolher "Título 60" decide
as quatro coisas de uma vez, e é por isso que a faixa aplica tudo num clique.

**Eixo 2: o tipo só tem duas cores.** Rosa e azul. Bege e branco são fundo;
preto é reserva e não entra em texto da marca. Toda decisão de cor de texto se
reduz a "qual dos dois papéis", e a exceção (branco) não é escolha, é a regra
de contraste agindo sozinha.

Todo o resto (raio, espaçamento, margem) deriva de um módulo de **20pt**.

---

## 2. Página

| | pt | px @1080 | cm |
|---|---|---|---|
| Largura | 1583,13 | 1920 | 55,8 |
| Altura | 890,63 | 1080 | 31,4 |

16:9 customizado. **Não** é o 16:9 padrão do PowerPoint (1280×720pt). Decks em
outro formato entram distorcidos; o botão **Page Size** da faixa converte.

A ponte entre mídias é a altura: **1080px ≡ 890,63pt**, logo `1pt = 1,2126px`.
É a razão que o `radiusPx` da configuração usa e o que permite o mesmo token
descrever um slide e uma tela.

---

## 3. Grid e espaçamento

Tudo é múltiplo de **20**. Margem, módulo e raio fecham entre si. No 1.0 não
fechavam (76 de margem com raio 20,62 não conversa com nada), e era isso que
fazia o alinhamento parecer arbitrário.

| Token | pt | px | cm | Papel |
|---|---|---|---|---|
| Módulo | 20 | 24 | 0,71 | A unidade. Tudo deriva dele |
| Raio | 20 | 24 | 0,71 | 1 módulo |
| Gutter | 40 | 49 | 1,41 | 2 módulos, o vão entre colunas |
| Margem | 60 | 73 | 2,12 | 3 módulos, nos **quatro** lados |
| Âncora de encaixe | 36 | 44 | 1,27 | Alinhamento entre objetos |

**Área útil: 1463 × 771pt.**

### As 4 colunas

| | |
|---|---|
| Coluna | 335,78pt |
| Bordas em x | 60 · 435,78 · 811,57 · 1187,35 |
| Larguras válidas | 1col 335,78 · 2col 711,57 · 3col 1087,35 · 4col 1463,13 |

Todo bloco de conteúdo começa numa borda de coluna e ocupa **1, 2, 3 ou 4
colunas**. Largura escolhida no olho não existe.

### A escala vertical

Gap entre blocos escolhe um degrau, e o degrau tem significado:

| Degrau | pt | Quando |
|---|---|---|
| 1× | 20 | Amarra rótulo ao valor, título ao subtítulo |
| 2× | 40 | Separa itens de uma lista |
| 3× | 60 | Separa blocos irmãos |
| 4× | 80 | Separa blocos |
| 6× | 120 | Separa **zonas** do slide |

**Raio** é um valor **visual constante**, não uma proporção: um card grande e um
botão pequeno têm o mesmo arredondamento de 20pt. Em OOXML isso vira
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

Bege (padrão), branco, azul e rosa. Sobre azul e rosa o texto é branco, não por
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
| Texto 24 | 24 | 29 | Regular | azul | 1,00× | Texto corrido: **o padrão** |
| Apoio 20 | 20 | 24 | Regular | azul | 1,30× | Apoio em pilares e colunas |
| Chapéu 18 | 18 | 22 | Bold | rosa | 1,00× | Chapéu acima do título |
| Texto 15 | 15 | 18 | Regular | azul | 1,30× | Texto corrido em slide denso |
| Legenda 12 | 12 | 15 | Bold | azul | 1,30× | Legendas, notas, fontes |
| CAPS 12 | 12 | 15 | Bold | azul | 1,00× | Rótulo em caixa alta |

### A lógica da entrelinha

Não é arbitrária: **quanto maior o corpo, mais apertada a linha.** Em 250 e 120pt
o espaço entre linhas já é enorme em valor absoluto, então o multiplicador
comprime (0,80×) para o bloco ler como uma massa só. Em 12 e 15pt acontece o
inverso: a linha precisa de ar para o olho achar a próxima (1,30×).

A entrelinha é **exata** (múltiplo do corpo), nunca "simples" ou "duplo", que
variam com a fonte e quebram a consistência entre máquinas.

**Legenda 12 e CAPS 12 dividem o corpo e divergem na entrelinha** (1,30× contra
1,00×). Por isso a entrelinha pertence ao **estilo**, não ao tamanho. Uma tabela
indexada por corpo não conseguiria descrever os dois.

> Existe uma segunda tabela, indexada por corpo, que alimenta os botões de
> Entrelinha e o Padronizar. Ela age em texto **sem estilo associado** e é só um
> padrão por tamanho. Onde os dois discordam, o estilo manda.

---

## 6. Comportamento do tipo

### Contraste automático

**Se a cor do texto for exatamente a cor do fundo, o texto vira branco.**

O fundo relevante é, nesta ordem: o preenchimento da própria forma; senão o
preenchimento do grupo que a contém; senão o fundo do slide, e "fundo do slide"
inclui uma forma sólida que cubra a página inteira, que é como a maioria dos
decks reais é montada.

Vale na criação e em toda aplicação de estilo. Não é sugestão: texto rosa sobre
fundo rosa é invisível, e o sistema não deixa acontecer.

### Ponto final colorido

O **Statement 120** pinta o ponto final de azul, só o ponto, e só se a frase
terminar em ".". Sobre fundo azul, o texto vira branco e o ponto vira rosa.

É o único estilo com comportamento próprio. Os outros doze são declarativos.

### Caixa alta

O **CAPS 12** aplica caixa alta e **3pt de espaçamento entre letras**. Caixa alta
sem espaçamento fecha demais em Avenir Next Bold.

É reversível: aplicar outro estilo por cima desfaz a caixa alta e zera o
espaçamento. Nenhum estilo deixa resíduo.

---

## 7. Composição

O 1.0 tinha os tokens certos e compunha errado. Estas são as regras que faltavam.

### Altura de caixa vem do conteúdo

Card ou bloco de cor = **padding (40) + o maior conteúdo do grupo + padding**.
Todos os cards de uma linha têm a mesma altura, definida pelo mais alto.

Nunca esticar a caixa até o rodapé porque "sobrou espaço". Espaço vazio embaixo
de um card curto é resultado correto; caixa alta com texto no topo é vício.

### O conteúdo ocupa o canvas

Se o conteúdo termina antes de **60% da altura útil**, o slide está mal composto.
A saída não é esticar caixa nem aumentar corpo: é **distribuir em zonas**. Um
bloco de argumento em cima, uma faixa de desdobramento ancorada na base, com 6
módulos de ar entre as duas.

### Lista é uma caixa só

Itens de mesma função e mesmo estilo vão num **único bloco de texto**, separados
por quebra de parágrafo. Uma caixa por item multiplica o trabalho de quem edita
e desalinha na primeira mudança de conteúdo.

### Texto que não cabe

Nesta ordem:

1. **Primeira frase no estilo pedido, o resto um degrau abaixo.** Dá ritmo
   editorial e mantém a hierarquia.
2. Se ainda não couber, o bloco inteiro desce um degrau.
3. Se ainda não couber, o texto é longo demais para o slide. Corte conteúdo.

Nunca encolher para um corpo fora da escala, nunca apertar a entrelinha.

### Texto corrido é Apoio 20

**Apoio 20** é o padrão de texto corrido. **Texto 15** existe para conteúdo
micro e denso: grades com muitos itens, fichas técnicas, tabelas.

### Capa

Sempre fundo saturado, **azul de preferência**, rosa como alternativa. Nunca
bege nem branco.

### Numeração

`1.` `2.` `3.`: com ponto. Nunca `01`, `02`.

---

## 8. Escrita

A tipografia não salva um texto mal escrito, e o sistema recusa três vícios:

- **Sem travessão () e sem hífen duplo (`--`).** Reescreva com vírgula, ponto ou
  dois-pontos. Ninguém escreve com hífen duplo.
- **Sem comentário solto no rodapé do slide.** Aquela frase pequena embaixo,
  explicando o que já está dito acima, só polui.
- **Sem "nota:", "atenção:", "obs.:"** pendurados no fim.

> A regra de ASCII puro do `.bas` (mojibake do VBE no Mac) vale **só** para
> strings do VBA. Markdown e PPTX levam acento e pontuação normais.

---

## 9. Anti-vícios

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

## 10. Migração

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

## 11. Fora do PowerPoint

Os tokens viajam; o que muda é a unidade.

**Web.** Use `px` da tabela (base 1080). A entrelinha vira `line-height` sem
unidade: o multiplicador é o mesmo número. O raio de 25px é literal. O módulo
de espaçamento vira 24px. Se o layout não for 1920×1080, escale por proporção,
nunca reescolha corpos: a hierarquia mora nas razões entre os treze, não nos
valores absolutos.

**Impresso.** Use `cm`. A página é 55,8 × 31,4cm. Em A4 paisagem (29,7 × 21cm)
o fator é 0,53.

**Social.** Recorte a área útil, nunca a página inteira: a margem de 76pt existe
para o conteúdo sobreviver a qualquer recorte.

**Regra que não muda em nenhuma mídia:** duas cores de tipo, escala fechada,
entrelinha do estilo, espaçamento em múltiplos do módulo, e o contraste
automático.

---

## 12. Dívidas conhecidas

Divergências reais entre o que este documento diz e o que existe hoje. Estão
aqui para não virarem surpresa:

As três primeiras são a **próxima leva do `.bas`**: exigem ciclo do VBE e novo
release, e o `check-tokens.py` avisa sobre elas a cada execução.

1. **Margem × linhas-guia.** O 2.0 usa 60pt (2,12cm); a faixa desenha as guias em
   3,15cm (89,3pt). Ajustar `GUIDE_MARGIN_CM` para `2.12`.
2. **Raio.** O 2.0 usa 20pt; a faixa aplica 20,62pt, porque guarda `radiusPx = 25`
   no canvas de 1080. Ajustar para `24` (19,79pt, o inteiro mais próximo). O
   ideal seria a config guardar pt em vez de px, mas isso muda o schema do
   `cba-config.txt` e quebra configurações salvas.
3. **A grade de colunas não existe na faixa.** As linhas-guia já desenham 4
   colunas com gutter de 2 raios, o que bate com o 2.0. Depois de ajustar margem
   e raio, a grade da faixa e a do sistema passam a ser a mesma.
4. **O rosa em outras mídias.** A landing (`--coral: #FC5E6D`) e a skill
   `cba-bg-design-system-v3` usam `FC5E6D`; o padrão é **`FD5E6D`**. Um dígito de
   diferença, invisível a olho nu e detectável em auditoria.
5. **Corpos herdados do v3.** 55, 36, 32, 22, 16 e 14pt aparecem em decks
   antigos. A tabela de migração acima resolve.

---

## 13. Onde cada coisa vive

| Arquivo | Papel |
|---|---|
| `v3-powerpoint-addin/assets/BG-DoodleStudio.bas` | **A fonte da verdade** (`SetDefaults`) |
| `tools/design-system/tokens.py` | Espelho em Python, para gerar e validar |
| `tools/design-system/check-tokens.py` | Gate que compara os dois |
| `config.html` | Página de Brand Standards, muda os tokens sem recompilar |
| `docs/CBA-Studio-Design-System.pptx` | Este documento em slides |
| `design.md` | Este arquivo |
