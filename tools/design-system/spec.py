"""O conteúdo do deck, como dados.

Os slides são dicts; quem sabe desenhar é o layouts.py. Onde o conteúdo É o
token (os 13 estilos, as 5 cores), a lista é GERADA a partir do tokens.py: assim o deck não pode divergir do sistema que documenta.
"""

from __future__ import annotations

import tokens as T

# Amostras por estilo. O corpo grande só cabe com pouca palavra; o pequeno
# precisa de frase inteira para a entrelinha aparecer.
SAMPLES = {
    "dsBigNumber": "47",
    "dsHero": "A marca não\né o que você diz.",
    "dsMega": "O começo de\num capítulo",
    "dsH1": "Título de página",
    "dsCorpo": "O corpo de destaque carrega a ideia principal\nsem virar título.",
    "dsH3": "Subtítulo que explica o título",
    "dsH4": "Título de tópico ou de card",
    "dsH5": "Texto corrido. É o estilo que toda caixa nova recebe ao nascer,\n"
            "porque é o que mais aparece num deck de verdade.",
    "dsCorpoPilar": "Texto de apoio em pilares e colunas. Respira mais que o\n"
                    "texto corrido porque costuma vir em bloco estreito.",
    "dsEyebrow": "CHAPÉU ACIMA DO TÍTULO",
    "dsTexto15": "Texto corrido pequeno, para slides densos onde 24pt não caberia.\n"
                 "A entrelinha de 1,3x mantém a leitura confortável mesmo apertado.",
    "dsLegenda12": "Legenda de imagem, nota de rodapé ou fonte de dado.\n"
                   "Fonte: CBA B+G, 2026.",
    "dsCaps12": "Rótulo em caixa alta",
}


def _footer_section(name):
    return name.upper()


def build_spec():
    S = []
    add = S.append

    # ---------------------------------------------------------- abertura
    add({"kind": "hero_cover", "bg": T.AZUL,
         "eyebrow": "CBA B+G  ·  DESIGN SYSTEM",
         "title": "Um sistema\ncabe em treze\nestilos.",
         "sub": "Documentação da tipografia, cor e grid do CBA Studio."})

    add({"kind": "spec_page", "section": "ABERTURA", "eyebrow": "O QUE É ISTO",
         "title": "A ferramenta é a definição",
         "intro": "Este documento não inventa um padrão: ele descreve o que a "
                  "faixa CBA Studio já aplica com um clique. A fonte da verdade "
                  "é o código do add-in; aqui e nos slides seguintes ele está "
                  "escrito em português.",
         "rows": [("FONTE DA VERDADE", "SetDefaults, no .bas"),
                  ("ESPELHO EM CÓDIGO", "tools/design-system/tokens.py"),
                  ("REGRAS EM TEXTO", "design.md"),
                  ("GATE ANTI-DERIVA", "check-tokens.py")]})

    add({"kind": "card_grid_5", "section": "ABERTURA", "cols": 4,
         "title": "O caminho",
         "intro": "Sete blocos, do que é invariável para o que é aplicado.",
         "items": [
             {"kicker": "1.", "title": "Fundamentos", "body": "Página, margem, módulo, raio."},
             {"kicker": "2.", "title": "Cor", "body": "Cinco slots, dois papéis de tipo."},
             {"kicker": "3.", "title": "Tipografia", "body": "Os treze estilos, um a um."},
             {"kicker": "4.", "title": "Componentes", "body": "Caixas, cards, imagens."},
             {"kicker": "5.", "title": "Arquétipos", "body": "Layouts que o sistema aceita."},
             {"kicker": "6.", "title": "Regras", "body": "Anti-vícios e migração."},
             {"kicker": "7.", "title": "Outras mídias", "body": "Web, impresso, social."},
             {"kicker": "8.", "title": "Aplicação", "body": "A faixa e o Brand Standards."},
         ]})

    # ---------------------------------------------------------- 01 fundamentos
    add({"kind": "chapter_divider", "bg": T.AZUL, "number": "1.",
         "title": "Fundamentos", "section": "FUNDAMENTOS",
         "sub": "A página, a margem, o módulo de espaçamento e o raio. "
                "Todo o resto deriva daqui."})

    add({"kind": "diagram_page", "section": "FUNDAMENTOS",
         "title": "A página",
         "intro": "16:9 customizado, não o 16:9 padrão do PowerPoint. Deck em "
                  "outro formato entra distorcido: o botão Page Size da faixa "
                  "converte.\n\nA altura é a ponte entre mídias: 1080px equivale "
                  "a 890,63pt, logo 1pt = 1,2126px.",
         "show_margin": False,
         "legend": [("LARGURA", "1583,13 pt  ·  1920 px  ·  55,8 cm"),
                    ("ALTURA", "890,63 pt  ·  1080 px  ·  31,4 cm"),
                    ("PROPORÇÃO", "16:9 customizado")]})

    add({"kind": "diagram_page", "section": "FUNDAMENTOS",
         "title": "Margem e área útil",
         "intro": "A margem tem 60pt, três módulos exatos, nos quatro lados. "
                  "Ela existe para o slide sobreviver a recorte, projeção e "
                  "reaproveitamento em outra mídia.",
         "show_margin": True,
         "legend": [("MARGEM", "60 pt  ·  73 px  ·  2,12 cm"),
                    ("ÁREA ÚTIL", "1463 x 771 pt"),
                    ("COLUNAS", "4 de 335,78 pt")]})

    add({"kind": "diagram_page", "section": "FUNDAMENTOS",
         "title": "O módulo de espaçamento",
         "intro": "Toda distância entre colunas, cards e blocos é múltiplo de "
                  "20pt, e o raio e a margem também: 20 e 60. Distância que não é "
                  "múltiplo de 20 é acidente, não decisão.",
         "show_margin": True, "show_module": True,
         "legend": [("MÓDULO", "20 pt  ·  24 px  ·  0,71 cm"),
                    ("MÚLTIPLOS", "20 · 40 · 60 · 80 pt"),
                    ("RELAÇÃO", "raio = 1 módulo, gutter = 2")]})

    add({"kind": "spec_page", "section": "FUNDAMENTOS",
         "title": "Âncora de encaixe",
         "intro": "A âncora não é a margem. A margem delimita a página; a âncora "
                  "é o passo de alinhamento entre objetos. São papéis diferentes "
                  "e convivem.",
         "rows": [("PADRÃO", "1,27 cm  ·  36 pt"),
                  ("EM PIXELS", "44 px @1080"),
                  ("PASSOS", "0,5 · 1 · 1,27 · 1,5 cm"),
                  ("PASSOS (cont.)", "2 · 2,5 · 3 · 4 · 5 cm")]})

    add({"kind": "spec_page", "section": "FUNDAMENTOS",
         "title": "Raio",
         "intro": "O raio é um valor visual constante, não uma proporção: um card "
                  "grande e um botão pequeno tem o mesmo arredondamento. É o que "
                  "faz formas de tamanhos diferentes parecerem da mesma família.",
         "rows": [("RAIO", "20 pt  ·  24 px  ·  0,71 cm"),
                  ("EM OOXML", "adj = raio / (menor lado / 2) x 50000"),
                  ("PILL", "adj = 50000"),
                  ("RETO", "adj = 0")]})

    add({"kind": "quote_side_image", "section": "FUNDAMENTOS",
         "eyebrow": "COMO OS TRÊS SE COMBINAM",
         "title": "Margem, módulo e raio",
         "body": "A margem diz onde começa. O módulo diz de quanto em quanto se "
                 "anda. O raio diz como a forma termina.\n\n"
                 "Com os três fixos, dois slides feitos por pessoas diferentes "
                 "saem parecidos sem combinação prévia: que é o único teste "
                 "que importa num design system.",
         "panel_bg": T.AZUL, "panel_style": "dsH3",
         "panel_text": "60 de margem.\n20 de módulo.\n20 de raio.\n\n"
                       "Tudo múltiplo de 20."})

    # ---------------------------------------------------------- 02 cor
    add({"kind": "chapter_divider", "bg": T.ROSA, "number": "2.",
         "title": "Cor", "section": "COR",
         "sub": "Cinco slots. Dois carregam tipo. Um resolve contraste sozinho."})

    add({"kind": "card_grid_5", "section": "COR", "cols": 5,
         "title": "A paleta",
         "intro": "Os índices importam: a configuração de marca referência por "
                  "número, não por nome.",
         "items": [{"kicker": "PAL%d" % i,
                    "title": T.PALETTE[i]["name"].capitalize(),
                    "body": "#" + T.PALETTE[i]["hex"]} for i in range(5)]})

    for i in range(5):
        add({"kind": "swatch_page", "section": "COR",
             "palette": T.PALETTE[i], "index": i})

    add({"kind": "multi_card_grid", "section": "COR",
         "eyebrow": "PAPÉIS DE COR", "title": "O tipo só tem duas cores",
         "cards": [
             {"title": "Rosa\npal0", "body": "Statements, títulos, chapéus e "
              "números-destaque. É a cor que marca hierarquia alta."},
             {"title": "Azul\npal1", "body": "Corpo de texto, tópicos, legendas "
              "e manchetes. É a cor que se le em volume."},
             {"title": "Branco\npal3", "body": "Não é escolha: aparece quando a "
              "regra de contraste age sobre fundo saturado."},
         ]})

    add({"kind": "quote_side_image", "section": "COR",
         "eyebrow": "REGRA DE CONTRASTE",
         "title": "Texto nunca some no fundo",
         "body": "Se a cor do texto for exatamente a cor do fundo, o texto vira "
                 "branco.\n\nO fundo relevante é, nesta ordem: o preenchimento da "
                 "própria forma; senão o do grupo que a contém; senão o do slide "
                 "e isso inclui uma forma sólida cobrindo a página, que é como "
                 "a maioria dos decks reais é montada.",
         "panel_bg": T.ROSA, "panel_style": "dsH1",
         "panel_text": "Título 60 é rosa.\nSobre fundo rosa,\nele nasce branco."})

    add({"kind": "multi_card_grid", "section": "COR", "card_bg": T.BEGE,
         "eyebrow": "FUNDOS", "title": "Fundos permitidos",
         "cards": [
             {"title": "Bege", "body": "O fundo de conteúdo padrão. Tipo em rosa e azul."},
             {"title": "Branco", "body": "Cards e areas de respiro. Tipo em rosa e azul."},
             {"title": "Azul", "body": "Aberturas e destaques. Tipo em branco."},
             {"title": "Rosa", "body": "Aberturas e destaques. Tipo em branco."},
         ]})

    add({"kind": "do_dont", "section": "COR", "title": "Cor: o que vale",
         "intro": "A paleta é fechada. Não existe 'um azul um pouco mais claro'.",
         "do": ["Usar os cinco slots como estao",
                "Deixar a regra de contraste decidir o branco",
                "Bege como fundo padrão de conteúdo",
                "Rosa e azul como os dois papéis de tipo"],
         "dont": ["Criar tom intermediário 'para variar'",
                  "Texto preto em conteúdo de marca",
                  "Cor de fundo fora dos quatro permitidos",
                  "Transparência para 'suavizar' uma cor"]})

    # ---------------------------------------------------------- 03 tipografia
    add({"kind": "chapter_divider", "bg": T.AZUL, "number": "3.",
         "title": "Tipografia", "section": "TIPOGRAFIA",
         "sub": "Uma família, treze estilos. Cada um já traz peso, cor e "
                "entrelinha: escolher o estilo decide as quatro coisas."})

    add({"kind": "card_grid_5", "section": "TIPOGRAFIA", "cols": 5,
         "title": "A escala",
         "intro": "Treze corpos, nada entre eles. Se não está aqui, não é do sistema.",
         "items": [{"kicker": "%gPT" % s["size"],
                    "title": s["label"].rsplit(" ", 1)[0],
                    "body": "%s · %s · %.2fx" % (
                        "Bold" if s["bold"] else "Regular",
                        T.PALETTE[0 if s["role"] == 0 else 1]["name"],
                        s["ent"])}
                   for s in T.STYLES]})

    for s in T.STYLES:
        add({"kind": "type_specimen", "section": "TIPOGRAFIA",
             "style_id": s["id"], "sample": SAMPLES[s["id"]]})

    add({"kind": "zoned_content", "section": "TIPOGRAFIA",
         "eyebrow": "ENTRELINHA",
         "title": "Quanto maior o corpo,\nmais apertada a linha",
         "body": "Não é arbitrario. Em 250 e 120pt o espaço entre linhas já é "
                 "enorme em valor absoluto: o multiplicador comprime para o bloco "
                 "ler como uma massa só.\n\nEm 12 e 15pt acontece o inverso: a "
                 "linha precisa de ar para o olho achar a proxima.",
         "items": [
             {"title": "Sempre exata",
              "body": "Multiplo do corpo, nunca 'simples' ou 'duplo', que variam "
                      "com a fonte e quebram entre máquinas."},
             {"title": "Pertence ao estilo",
              "body": "Legenda 12 e CAPS 12 dividem o corpo e divergem na "
                      "entrelinha. Uma tabela por tamanho não descreveria os dois."},
             {"title": "Há uma tabela por tamanho",
              "body": "Ela alimenta os botões de Entrelinha e age em texto sem "
                      "estilo. Onde discordam, o estilo manda."},
         ]})

    ent_rows = [("%gPT" % s["size"], "%.2fx   =   %.1f pt" % (s["ent"], s["size"] * s["ent"]))
                for s in T.STYLES]
    add({"kind": "spec_page", "section": "TIPOGRAFIA",
         "title": "Entrelinha por estilo",
         "intro": "O valor efetivo em pontos é o corpo vezes o multiplicador.",
         "rows": ent_rows})

    add({"kind": "quote_side_image", "section": "TIPOGRAFIA",
         "eyebrow": "COMPORTAMENTO PROPRIO",
         "title": "O ponto final",
         "body": "O Statement 120 pinta o ponto final de azul. Só o ponto, e só "
                 "se a frase terminar em ponto.\n\nSobre fundo azul, o texto vira "
                 "branco e o ponto vira rosa: as duas regras agindo juntas sem "
                 "se anular.\n\nÉ o único estilo com comportamento próprio. Os "
                 "outros doze são declarativos.",
         "panel_bg": T.AZUL, "panel_style": "dsHero",
         "panel_text": "Branco.\nPonto rosa."})

    add({"kind": "quote_side_image", "section": "TIPOGRAFIA",
         "eyebrow": "COMPORTAMENTO PROPRIO",
         "title": "Caixa alta",
         "body": "O CAPS 12 aplica caixa alta e 3pt de espaçamento entre letras. "
                 "Caixa alta sem espaçamento fecha demais em Avenir Next Bold.\n\n"
                 "É reversível: aplicar outro estilo por cima desfaz a caixa alta "
                 "e zera o espaçamento. Nenhum estilo deixa resíduo.",
         "panel_bg": T.BRANCO, "panel_style": "dsCaps12",
         "panel_text": "Rótulo em caixa alta\ncom três pontos\nde espaçamento"})

    # ---------------------------------------------------------- 04 componentes
    add({"kind": "chapter_divider", "bg": T.ROSA, "number": "4.",
         "title": "Componentes", "section": "COMPONENTES",
         "sub": "Poucas pecas, todas derivadas do raio e do módulo."})

    add({"kind": "multi_card_grid", "section": "COMPONENTES", "card_bg": T.BRANCO,
         "eyebrow": "AS DUAS PECAS BASE", "title": "Caixa e bloco",
         "cards": [
             {"title": "Caixa de texto", "body": "Nasce em Texto 24, âncorada na "
              "medida escolhida. Sem preenchimento e sem contorno."},
             {"title": "Rounded box", "body": "Bloco de cor com o raio do sistema. "
              "Tambem nasce em Texto 24."},
             {"title": "Pill", "body": "Bloco totalmente arredondado. Para rotulos "
              "e banners curtos, nunca para texto longo."},
         ]})

    add({"kind": "card_grid_5", "section": "COMPONENTES", "cols": 4,
         "title": "Cards",
         "intro": "Card é rounded box com padding de um módulo. A grade sai da "
                  "divisão da área útil com gaps múltiplos de 20pt.",
         "items": [
             {"kicker": "2 CARDS", "title": "706 pt", "body": "gap de 20pt"},
             {"kicker": "3 CARDS", "title": "464 pt", "body": "gap de 20pt"},
             {"kicker": "4 CARDS", "title": "343 pt", "body": "gap de 20pt"},
             {"kicker": "PADDING", "title": "40 pt", "body": "dois módulos"},
         ]})

    add({"kind": "spec_page", "section": "COMPONENTES",
         "title": "Imagens",
         "intro": "Imagem entra com o mesmo raio das formas. Recorte é pelo botão "
                  "Crop da faixa, que usa o comando nativo do PowerPoint: nada "
                  "de esticar para caber.",
         "rows": [("RAIO", "20 pt, igual às formas"),
                  ("RECORTE", "Formas > Crop"),
                  ("PROPORÇÃO", "nunca distorcer"),
                  ("LEGENDA", "Legenda 12, abaixo")]})

    add({"kind": "do_dont", "section": "COMPONENTES", "title": "Componentes: o que vale",
         "intro": "Quando bater a vontade de separar com um traco, a resposta é hierarquia.",
         "do": ["Espaço em branco em múltiplo de 20pt",
                "Mudanca de corpo para separar assunto",
                "Bloco de cor para agrupar",
                "O mesmo raio em tudo"],
         "dont": ["Linha fina decorativa",
                  "Barra vertical de destaque",
                  "Sombra, gradiente ou brilho",
                  "Raio diferente 'porque a forma é menor'"]})

    # ---------------------------------------------------------- 05 arquétipos
    add({"kind": "chapter_divider", "bg": T.AZUL, "number": "5.",
         "title": "Arquétipos", "section": "ARQUÉTIPOS",
         "sub": "Os layouts que o sistema aceita. Fora deles, é composição nova "
                "e composição nova pede revisão."})

    add({"kind": "card_grid_5", "section": "ARQUÉTIPOS", "cols": 4,
         "title": "O catálogo, 1 de 2",
         "intro": "Os arquétipos de abertura, especificação e grade.",
         "items": [
             {"kicker": "1.", "title": "Hero cover", "body": "Statement que abre. Fundo sempre saturado."},
             {"kicker": "2.", "title": "Chapter divider", "body": "Número e título de bloco."},
             {"kicker": "3.", "title": "Spec page", "body": "Título, intro e pares rótulo/valor."},
             {"kicker": "4.", "title": "Type specimen", "body": "Amostra no corpo real e ficha técnica."},
             {"kicker": "5.", "title": "Swatch", "body": "Uma cor com hex, RGB e papel."},
             {"kicker": "6.", "title": "Multi card", "body": "2 a 4 cards de mesma altura."},
             {"kicker": "7.", "title": "Card grid", "body": "Grade densa, com caixa."},
         ]})

    add({"kind": "card_grid_5", "section": "ARQUÉTIPOS", "cols": 4,
         "title": "O catálogo, 2 de 2",
         "intro": "Os arquétipos de composição, regra e fecho.",
         "items": [
             {"kicker": "8.", "title": "Grid plain", "body": "A mesma grade, sem caixa."},
             {"kicker": "9.", "title": "Zoned content", "body": "Zona superior e faixa inferior."},
             {"kicker": "10.", "title": "Stat band", "body": "Faixa inferior de números."},
             {"kicker": "11.", "title": "Quote side", "body": "Texto e painel de cor."},
             {"kicker": "12.", "title": "Do e nunca", "body": "Duas colunas de regra."},
             {"kicker": "13.", "title": "Diagram", "body": "Esquema desenhado em escala."},
             {"kicker": "14.", "title": "Closing", "body": "Fecho do documento."},
         ]})

    add({"kind": "hero_cover", "bg": T.ROSA, "section": "ARQUÉTIPOS",
         "eyebrow": "ARQUÉTIPO 01  ·  HERO COVER",
         "title": "Uma frase que\nmarca a página.",
         "sub": "Statement 120 sobre fundo saturado: o texto vira branco e o ponto, rosa."})

    add({"kind": "chapter_divider", "bg": T.BEGE, "number": "2.",
         "title": "Chapter divider\nsobre bege", "section": "ARQUÉTIPOS",
         "sub": "O mesmo arquétipo do bloco 01, com fundo claro: Big Number e "
                "Manchete mantém a cor de tipo."})

    add({"kind": "zoned_content", "section": "ARQUÉTIPOS",
         "eyebrow": "ARQUÉTIPO 08", "title": "Pillar dense",
         "body": "Coluna de argumento à esquerda, desdobramento à direita. É o "
                 "layout de maior densidade que o sistema aceita sem virar "
                 "poluição.\n\nUse quando o conteúdo tem uma tese e três a cinco "
                 "consequências.",
         "items": [
             {"title": "Uma tese por slide", "body": "Se há duas, são dois slides."},
             {"title": "Três a cinco itens", "body": "Menos que três não pede o layout; mais que cinco não cabe."},
             {"title": "Título em Tópico 28", "body": "Corpo em Texto 15."},
         ]})

    add({"kind": "quote_side_image", "section": "ARQUÉTIPOS",
         "eyebrow": "ARQUÉTIPO 09", "title": "Quote side",
         "body": "Texto à esquerda, painel de cor à direita. O painel não é "
                 "decoracao: ele carrega a citação, o número ou a amostra que "
                 "sustenta o argumento do lado esquerdo.",
         "panel_bg": T.ROSA, "panel_style": "dsCorpo",
         "panel_text": "O painel carrega\no que o texto\nafirma."})

    add({"kind": "multi_card_grid", "section": "ARQUÉTIPOS", "card_bg": T.BEGE,
         "eyebrow": "ARQUÉTIPO 06", "title": "Multi card",
         "cards": [
             {"title": "Dois a quatro", "body": "Mais que quatro vira grade densa, "
              "que é outro arquétipo."},
             {"title": "Mesma altura", "body": "Cards de alturas diferentes "
              "quebram a leitura horizontal."},
             {"title": "Gap de 20pt", "body": "Sempre o módulo, nunca uma folga "
              "escolhida no olho."},
         ]})

    add({"kind": "spec_page", "section": "ARQUÉTIPOS", "bg": T.BRANCO,
         "eyebrow": "ARQUÉTIPO 03", "title": "Spec page sobre branco",
         "intro": "O mesmo arquétipo de especificacao, em fundo branco. Todos os "
                  "arquétipos aceitam os quatro fundos permitidos sem ajuste.",
         "rows": [("FUNDO", "Branco"), ("TIPO", "Rosa e azul"),
                  ("CONTRASTE", "Não dispara"), ("USO", "Areas de respiro")]})

    add({"kind": "card_grid_5", "section": "ARQUÉTIPOS", "cols": 2,
         "title": "Card grid em duas colunas",
         "intro": "A mesma grade, com menos colunas e mais corpo por item.",
         "items": [
             {"kicker": "VARIAÇÃO", "title": "Duas colunas", "body": "Dois módulos de largura cada, para item com texto."},
             {"kicker": "VARIAÇÃO", "title": "Quatro colunas", "body": "Uma coluna cada, para rótulo e número."},
         ]})

    add({"kind": "grid_plain", "section": "ARQUÉTIPOS", "cols": 4,
         "title": "Grid plain",
         "intro": "A mesma grade sem caixa. Quando o item é curto, a caixa não "
                  "acrescenta informação, só peso.",
         "items": [
             {"kicker": "1.", "title": "Sem caixa", "body": "O alinhamento sustenta o agrupamento."},
             {"kicker": "2.", "title": "Mesma grade", "body": "As colunas são as mesmas do card grid."},
             {"kicker": "3.", "title": "Item curto", "body": "Rótulo, título e uma linha de apoio."},
             {"kicker": "4.", "title": "Mais leve", "body": "Preferir quando houver seis itens ou mais."},
         ]})

    add({"kind": "stat_band", "section": "ARQUÉTIPOS",
         "eyebrow": "ARQUÉTIPO 10", "title": "Stat band",
         "body": "Variação da faixa inferior para números. A zona de cima "
                 "sustenta o argumento e a faixa entrega a evidência, sem "
                 "competir por atenção.",
         "stats": [
             {"value": "13", "label": "estilos na escala fechada"},
             {"value": "5", "label": "cores na paleta"},
             {"value": "20", "label": "pontos de módulo"},
             {"value": "4", "label": "colunas na grade"},
         ]})

    add({"kind": "multi_card_grid", "section": "ARQUÉTIPOS", "bg": T.AZUL,
         "card_bg": T.BRANCO, "eyebrow": "VARIAÇÃO DE FUNDO",
         "title": "Multi card sobre azul",
         "cards": [
             {"title": "Fundo saturado", "body": "O título vira branco pela regra "
              "de contraste, sem ninguém escolher."},
             {"title": "Card claro", "body": "O card mantém o fundo branco e o "
              "tipo volta às cores da marca."},
         ]})

    # ---------------------------------------------------------- 06 regras
    add({"kind": "chapter_divider", "bg": T.ROSA, "number": "6.",
         "title": "Regras", "section": "REGRAS",
         "sub": "O que o sistema nunca emite, e para onde vai o que ficou de fora."})

    add({"kind": "do_dont", "section": "REGRAS", "title": "Anti-vícios",
         "intro": "A lista do que o sistema recusa. Não é preferência estética: "
                  "é o que mantém dois decks parecidos.",
         "do": ["Corpo da escala dos treze",
                "Cor dos cinco slots",
                "Entrelinha do estilo",
                "Avenir Next, Regular ou Bold"],
         "dont": ["Linha decorativa ou divisória",
                  "Gradiente, sombra, brilho, contorno",
                  "Cor ou corpo fora da tabela",
                  "Icone inventado"]})

    add({"kind": "card_grid_5", "section": "REGRAS", "cols": 3,
         "title": "Checklist de auditoria",
         "intro": "O que olhar antes de mandar um deck para fora.",
         "items": [
             {"kicker": "1.", "title": "Fonte", "body": "Toda run em Avenir Next."},
             {"kicker": "2.", "title": "Corpo", "body": "Todo corpo na escala dos treze."},
             {"kicker": "3.", "title": "Cor", "body": "Toda cor nos cinco slots."},
             {"kicker": "4.", "title": "Entrelinha", "body": "Bate com a do estilo."},
             {"kicker": "5.", "title": "Margem", "body": "Nada invade os 76pt."},
             {"kicker": "6.", "title": "Contraste", "body": "Nenhum texto na cor do fundo."},
         ]})

    add({"kind": "spec_page", "section": "REGRAS",
         "title": "Migração",
         "intro": "Corpos de versões anteriores que sairam da escala fechada. O "
                  "botão Padronizar tipografia faz o arredondamento sozinho.",
         "rows": [("55 PT", "vira 60 pt"), ("36 E 32 PT", "vira 34 pt"),
                  ("22 PT", "vira 24 pt"), ("16 E 14 PT", "vira 15 pt"),
                  ("ROTULO 60", "vira Título 60"), ("ACIMA DE 185 PT", "vira 250 pt")]})

    # ---------------------------------------------------------- 07 outras mídias
    add({"kind": "chapter_divider", "bg": T.AZUL, "number": "7.",
         "title": "Outras mídias", "section": "OUTRAS MÍDIAS",
         "sub": "Os tokens viajam. O que muda é a unidade."})

    add({"kind": "multi_card_grid", "section": "OUTRAS MÍDIAS", "card_bg": T.BRANCO,
         "eyebrow": "A MESMA REGRA EM TRÊS UNIDADES", "title": "Web, impresso, social",
         "cards": [
             {"title": "Web", "body": "Use px da tabela (base 1080). A entrelinha "
              "vira line-height sem unidade: o multiplicador é o mesmo número. "
              "Raio de 25px, módulo de 24px."},
             {"title": "Impresso", "body": "Use cm. A página é 55,8 x 31,4cm; em "
              "A4 paisagem o fator é 0,53. A escala não muda, só o tamanho fisico."},
             {"title": "Social", "body": "Recorte a área útil, nunca a página "
              "inteira. A margem de 76pt existe para o conteúdo sobreviver ao corte."},
         ]})

    add({"kind": "quote_side_image", "section": "OUTRAS MÍDIAS",
         "eyebrow": "O QUE NAO MUDA",
         "title": "A hierarquia mora nas razões",
         "body": "Se o layout não for 1920x1080, escale por proporção: nunca "
                 "reescolha corpos.\n\nA diferenca entre Título 60 e Texto 24 não "
                 "está em 60 e 24: está na razão entre eles. Mantida a razão, o "
                 "sistema sobrevive a qualquer tamanho de tela ou papel.",
         "panel_bg": T.BEGE, "panel_style": "dsH3",
         "panel_text": "Duas cores de tipo.\nEscala fechada.\nEntrelinha do estilo.\n"
                       "Espacamento em módulo.\nContraste automatico."})

    # ---------------------------------------------------------- 08 aplicação
    add({"kind": "chapter_divider", "bg": T.ROSA, "number": "8.",
         "title": "Aplicação", "section": "APLICAÇÃO",
         "sub": "Onde este sistema vira um clique."})

    add({"kind": "card_grid_5", "section": "APLICAÇÃO", "cols": 4,
         "title": "A faixa CBA Studio",
         "intro": "Os grupos da aba, e o que cada um resolve.",
         "items": [
             {"kicker": "INSERIR", "title": "Caixa e bloco", "body": "Já nascem no padrão."},
             {"kicker": "TIPOGRAFIA", "title": "13 estilos", "body": "Um clique aplica quatro decisões."},
             {"kicker": "ENTRELINHA", "title": "Por tamanho", "body": "Para texto sem estilo."},
             {"kicker": "TEXTO", "title": "Peso e alinhamento", "body": "Negrito, esquerda, centro."},
             {"kicker": "FONTE / FUNDO", "title": "Paleta", "body": "Só os cinco slots."},
             {"kicker": "FORMAS", "title": "Raio e crop", "body": "Rounded, tirar, recortar."},
             {"kicker": "ALINHAR", "title": "Âncoras", "body": "Encaixe em medida fixa."},
             {"kicker": "AUDITORIA", "title": "Padronizar", "body": "Encaixa o que fugiu."},
         ]})

    add({"kind": "spec_page", "section": "APLICAÇÃO",
         "title": "Mudar os padroes",
         "intro": "A página de Brand Standards gera uma linha de configuração que "
                  "a faixa aplica sem recompilar nada. É como o sistema evolui "
                  "sem virar código novo.",
         "rows": [("PÁGINA", "Padroes > Abrir página"),
                  ("APLICAR", "Padroes > Aplicar config"),
                  ("FORMATO", "chave=valor separado por ;"),
                  ("ESTILO", "s_<id> = corpo|bold|papel|entrelinha")]})

    add({"kind": "closing", "bg": T.AZUL,
         "title": "Treze estilos,\ncinco cores,\num módulo."})

    return S
