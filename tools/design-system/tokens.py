"""Tokens do design system CBA Studio.

FONTE DA VERDADE: o `SetDefaults` de v3-powerpoint-addin/assets/BG-DoodleStudio.bas.
Este arquivo e' um ESPELHO em Python, para o gerador do deck e para qualquer
ferramenta que precise dos tokens fora do PowerPoint.

O espelho pode mentir se alguem mexer no .bas e esquecer daqui -- por isso
existe o check-tokens.py, que compara os dois e falha. Ele roda no checklist
de publicacao (tools/BUILD.md 4.1).

Unidades: o PowerPoint pensa em PONTOS, a web em PIXELS e o usuario em
CENTIMETROS. Todo token de medida traz as tres, para o sistema atravessar
midia sem ninguem ter que reinterpretar.
"""

from __future__ import annotations

# ---------------------------------------------------------------- pagina
PAGE_W_PT = 1583.13
PAGE_H_PT = 890.63
EMU_PER_PT = 12700
PT_PER_CM = 28.3465

# Canvas de referencia da extensao (o motor de doodle desenha em 1920x1080).
# E' a ponte entre o mundo "px" e o mundo "pt": tudo que a config guarda em px
# assume esta altura.
CANVAS_H_PX = 1080


def pt_to_px(pt: float) -> float:
    """Pontos -> pixels no canvas de 1080 de altura."""
    return pt * CANVAS_H_PX / PAGE_H_PT


def px_to_pt(px: float) -> float:
    """Pixels (canvas 1080) -> pontos."""
    return px * PAGE_H_PT / CANVAS_H_PX


def cm_to_pt(cm: float) -> float:
    return cm * PT_PER_CM


# ---------------------------------------------------------------- cor
# Os cinco slots da paleta, na ordem em que o .bas os declara (pal0..pal4).
# Os indices IMPORTAM: a config (cba-config.txt) referencia por numero.
PALETTE = {
    0: {"name": "rosa",   "hex": "FD5E6D", "rgb": (253, 94, 109),
        "role": "Cor de tipo 1. Statements, títulos, chapéus, números-destaque."},
    1: {"name": "azul",   "hex": "436AE1", "rgb": (67, 106, 225),
        "role": "Cor de tipo 2. Corpo de texto, tópicos, legendas, manchetes."},
    2: {"name": "bege",   "hex": "EEECE6", "rgb": (238, 236, 230),
        "role": "Fundo de conteúdo padrão. Nunca é cor de tipo."},
    3: {"name": "branco", "hex": "FFFFFF", "rgb": (255, 255, 255),
        "role": "Fundo claro e texto sobre fundo saturado (regra de contraste)."},
    4: {"name": "preto",  "hex": "000000", "rgb": (0, 0, 0),
        "role": "Reserva. Não usar em tipo da marca."},
}

# Atalhos por nome, para o gerador nao ficar cheio de indice magico.
ROSA = PALETTE[0]["hex"]
AZUL = PALETTE[1]["hex"]
BEGE = PALETTE[2]["hex"]
BRANCO = PALETTE[3]["hex"]
PRETO = PALETTE[4]["hex"]

# So' estes dois carregam tipo. "role" nos estilos abaixo aponta pra ca.
TYPE_ROLES = {0: ROSA, 1: AZUL}

# ---------------------------------------------------------------- tipografia
FONT = "Avenir Next"

# Os 13 estilos da faixa, na ordem em que aparecem no ribbon.
#   id      -> id do botao no customUI14.xml (e chave em cba-config.txt)
#   label   -> o rotulo que o usuario le
#   size    -> corpo em pt
#   bold    -> Avenir Next Bold
#   role    -> 0 rosa / 1 azul
#   ent     -> entrelinha (multiplo exato, nao "simples/duplo")
#   caps    -> caixa alta
#   spacing -> espacamento entre letras em pt (0 = normal)
#   use     -> quando usar
STYLES = [
    {"id": "dsBigNumber",  "label": "Big Number 250", "size": 250, "bold": True,  "role": 0, "ent": 1.0,
     "use": "O número que carrega a página inteira. Um por slide, sem concorrência."},
    {"id": "dsHero",       "label": "Statement 120",  "size": 120, "bold": True,  "role": 0, "ent": 0.8,
     "period": 1, "bg_aware": True,
     "use": "A frase-manifesto. Ponto final em azul; sobre fundo azul, o texto vira branco e o ponto rosa."},
    {"id": "dsMega",       "label": "Manchete 80",    "size": 80,  "bold": True,  "role": 1, "ent": 0.9,
     "use": "Título de abertura de capítulo."},
    {"id": "dsH1",         "label": "Título 60",      "size": 60,  "bold": True,  "role": 0, "ent": 0.9,
     "use": "Título de página."},
    {"id": "dsCorpo",      "label": "Destaque 44",    "size": 44,  "bold": False, "role": 1, "ent": 1.15,
     "use": "Corpo de destaque e manifesto. O maior corpo que ainda é texto, não título."},
    {"id": "dsH3",         "label": "Subtítulo 34",   "size": 34,  "bold": True,  "role": 0, "ent": 0.95,
     "use": "Subtítulo e descritivo abaixo do título."},
    {"id": "dsH4",         "label": "Tópico 28",      "size": 28,  "bold": True,  "role": 1, "ent": 1.0,
     "use": "Título de tópico ou de card."},
    {"id": "dsH5",         "label": "Texto 24",       "size": 24,  "bold": False, "role": 1, "ent": 1.0,
     "use": "Texto corrido. O padrão de toda caixa nova."},
    {"id": "dsCorpoPilar", "label": "Apoio 20",       "size": 20,  "bold": False, "role": 1, "ent": 1.3,
     "use": "Texto de apoio em pilares e colunas."},
    {"id": "dsEyebrow",    "label": "Chapéu 18",      "size": 18,  "bold": True,  "role": 0, "ent": 1.0,
     "use": "Chapéu acima do título."},
    {"id": "dsTexto15",    "label": "Texto 15",       "size": 15,  "bold": False, "role": 1, "ent": 1.3,
     "use": "Texto corrido pequeno, em slides densos."},
    {"id": "dsLegenda12",  "label": "Legenda 12",     "size": 12,  "bold": True,  "role": 1, "ent": 1.3,
     "use": "Legendas, notas e fontes."},
    {"id": "dsCaps12",     "label": "CAPS 12",        "size": 12,  "bold": True,  "role": 1, "ent": 1.0,
     "caps": True, "spacing": 3,
     "use": "Rótulo em caixa alta. Mesmo corpo da Legenda 12, entrelinha diferente."},
]

STYLE_BY_ID = {s["id"]: s for s in STYLES}

# Tabela por TAMANHO: alimenta os botoes de Entrelinha e o Padronizar, que agem
# em texto SEM estilo associado. Onde dois estilos dividem o corpo (12pt), aqui
# vale o padrao -- o estilo carrega o proprio valor e nao depende desta tabela.
LINE_HEIGHT_BY_SIZE = {
    250: 1.0, 120: 0.8, 80: 0.9, 60: 0.9, 44: 1.15, 34: 0.95,
    28: 1.0, 24: 1.0, 20: 1.3, 18: 1.0, 15: 1.3, 12: 1.3,
}

ALLOWED_SIZES = frozenset(s["size"] for s in STYLES)
ALLOWED_HEX = frozenset(c["hex"] for c in PALETTE.values())

# ---------------------------------------------------------------- forma e grid
#
#  DESIGN SYSTEM 2.0 -- 60 / 20 / 20
#  Margem, modulo e raio sao todos multiplos de 20. Era o problema do 1.0:
#  76 de margem com modulo 20 e raio 20,62 nao fechava com nada.
#
SPACING_UNIT_PT = 20                # o modulo. Tudo deriva dele.
RADIUS_PT = 20.0                    # raio canonico (visual constante)
GUTTER_PT = 2 * RADIUS_PT           # 40 pt -- o vao entre colunas e' 2 raios
MARGIN_PT = 3 * SPACING_UNIT_PT     # 60 pt, nos QUATRO lados

# Escala vertical. Cada degrau tem significado; gap fora dela e' acidente.
#   1x amarra rotulo ao valor      3x separa blocos irmaos
#   2x separa itens de uma lista   4x separa blocos
#   6x separa ZONAS do slide (topo x faixa inferior)
SPACING = {n: SPACING_UNIT_PT * n for n in (1, 2, 3, 4, 6)}
V_TIGHT, V_ITEM, V_BLOCK, V_SECTION, V_ZONE = (SPACING[n] for n in (1, 2, 3, 4, 6))

MARGIN_LEFT_PT = MARGIN_PT
MARGIN_RIGHT_PT = PAGE_W_PT - MARGIN_PT
MARGIN_TOP_PT = MARGIN_PT
MARGIN_BOTTOM_PT = MARGIN_PT
CONTENT_W_PT = MARGIN_RIGHT_PT - MARGIN_LEFT_PT          # 1463.13
CONTENT_H_PT = PAGE_H_PT - MARGIN_TOP_PT - MARGIN_BOTTOM_PT  # 770.63
CONTENT_BOTTOM_PT = PAGE_H_PT - MARGIN_BOTTOM_PT         # 830.63

# Grade de 4 colunas. Bloco de conteudo ocupa 1, 2, 3 ou 4 colunas -- nunca
# uma largura escolhida no olho.
COLUMNS = 4
COL_W_PT = (CONTENT_W_PT - (COLUMNS - 1) * GUTTER_PT) / COLUMNS   # 335.78


def col_x(i: int) -> float:
    """x da borda esquerda da coluna i (0-based)."""
    return MARGIN_LEFT_PT + i * (COL_W_PT + GUTTER_PT)


def col_span(n: int) -> float:
    """Largura de um bloco de n colunas, gutters incluidos."""
    return n * COL_W_PT + (n - 1) * GUTTER_PT


# Padding interno de card/box: 2 modulos.
BOX_PAD_PT = SPACING[2]

# Uma caixa nao pode ser muito mais alta que o conteudo que carrega. Acima
# disto o build avisa -- era o vicio de esticar o card ate' o rodape.
BOX_STRETCH_MAX = 1.6

# O conteudo tem de ocupar o canvas. Se terminar antes disto, o layout deve
# distribuir em zonas em vez de deixar tudo empilhado no topo.
CANVAS_FILL_MIN = 0.60

# ---- divida com a extensao (ver design.md) -------------------------------
# O .bas ainda guarda o raio em px e as guias em 3,15 cm. Os dois numeros
# abaixo existem so' para o check-tokens.py medir a distancia e avisar.
RADIUS_PX_BAS = 25                  # = 20.62 pt
RADIUS_PT_BAS = round(px_to_pt(RADIUS_PX_BAS), 2)

ANCHOR_DEFAULT_CM = 1.27            # ancora de encaixe padrao (esq. e topo)
ANCHOR_DEFAULT_PT = round(cm_to_pt(ANCHOR_DEFAULT_CM), 2)   # 36.0 pt
ANCHOR_STEPS_CM = [0.5, 1.0, 1.27, 1.5, 2.0, 2.5, 3.0, 4.0, 5.0]

GUIDE_MARGIN_CM = 3.15              # onde a faixa desenha as linhas-guia hoje
GUIDE_MARGIN_PT = round(cm_to_pt(GUIDE_MARGIN_CM), 2)       # 89.29 pt

MAX_DEPTH = 12                      # recursao maxima em grupos aninhados
