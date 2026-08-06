"""Medicao de texto com a fonte real.

Sem isto, o gerador chuta a altura de cada bloco e o texto transborda por cima
do bloco seguinte -- foi exatamente o que aconteceu na primeira versao do deck.
Medindo com a Avenir Next de verdade, o layout empilha blocos sabendo quanto
cada um ocupa.

Usa Pillow so' para medir; nada e' rasterizado.
"""

from __future__ import annotations

import functools

from PIL import ImageFont

import tokens as T

FONT_FILE = "/System/Library/Fonts/Avenir Next.ttc"
FACE_REGULAR = 7   # ('Avenir Next', 'Regular')
FACE_BOLD = 0      # ('Avenir Next', 'Bold')

# Sem a fonte instalada caimos numa estimativa por largura media de caractere.
# Pior, mas nao quebra o build em outra maquina.
_FALLBACK_CHAR_W = 0.52


@functools.lru_cache(maxsize=256)
def _font(size_pt: float, bold: bool):
    try:
        return ImageFont.truetype(FONT_FILE, int(round(size_pt)),
                                  index=FACE_BOLD if bold else FACE_REGULAR)
    except Exception:
        return None


def text_width(text: str, size_pt: float, bold: bool) -> float:
    f = _font(size_pt, bold)
    if f is None:
        return len(text) * size_pt * _FALLBACK_CHAR_W
    return f.getlength(text)


def wrap_lines(text: str, size_pt: float, bold: bool, width_pt: float,
               caps: bool = False, spacing_pt: float = 0) -> int:
    """Quantas linhas o texto ocupa numa caixa desta largura.

    Respeita as quebras explicitas (\\n) e quebra o resto por palavra, como o
    PowerPoint faz.
    """
    if width_pt <= 0:
        return 1
    total = 0
    for para in text.split("\n"):
        if not para.strip():
            total += 1
            continue
        src = para.upper() if caps else para
        words = src.split()
        line, n = "", 1
        for w in words:
            probe = (line + " " + w).strip()
            wpx = text_width(probe, size_pt, bold)
            if spacing_pt:
                wpx += spacing_pt * max(0, len(probe) - 1)
            if wpx <= width_pt or not line:
                line = probe
            else:
                n += 1
                line = w
        total += n
    return total


def block_height(text: str, style: dict, width_pt: float) -> float:
    """Altura que o bloco precisa, na entrelinha do estilo."""
    n = wrap_lines(text, style["size"], bool(style["bold"]), width_pt,
                   caps=bool(style.get("caps")),
                   spacing_pt=float(style.get("spacing") or 0))
    # A entrelinha multiplica o corpo; a primeira linha ainda ocupa o corpo
    # inteiro mesmo quando o multiplicador e' menor que 1 (0,8x do Statement).
    lead = style["size"] * style["ent"]
    return max(style["size"], lead) + lead * (n - 1)


def fits(text: str, style_id: str, width_pt: float, height_pt: float) -> bool:
    return block_height(text, T.STYLE_BY_ID[style_id], width_pt) <= height_pt + 0.5


# Degraus de descida quando o texto nao cabe. So' estilos de TEXTO CORRIDO --
# nao faz sentido "descer" um Statement para uma Legenda.
STEP_DOWN = {
    "dsHero": "dsMega",
    "dsMega": "dsH1",
    "dsH1": "dsH3",
    "dsCorpo": "dsH3",
    "dsH3": "dsH5",
    "dsH4": "dsH5",
    "dsH5": "dsCorpoPilar",
    "dsCorpoPilar": "dsTexto15",
}


def split_first_sentence(text: str):
    """Separa a primeira frase do resto. Devolve (frase, resto) ou (text, '')."""
    for i, ch in enumerate(text):
        if ch in ".!?" and i + 1 < len(text) and text[i + 1] in " \n":
            return text[:i + 1], text[i + 2:].lstrip()
    return text, ""


def fit_text(text: str, style_id: str, width_pt: float, height_pt: float):
    """Encaixa o texto na altura disponivel, do jeito editorial.

    1) Cabe no estilo pedido? usa ele.
    2) Senao: primeira frase no estilo pedido, o resto um degrau abaixo.
       E' o que da' ritmo em vez de encolher o bloco inteiro.
    3) Ainda nao cabe? desce os dois degraus, e repete.

    Devolve lista de (style_id, texto) para o chamador empilhar.
    """
    if fits(text, style_id, width_pt, height_pt):
        return [(style_id, text)]

    head, tail = split_first_sentence(text)
    cur = style_id
    while cur:
        sub = STEP_DOWN.get(cur)
        if tail and sub:
            h = (block_height(head, T.STYLE_BY_ID[cur], width_pt)
                 + T.SPACING[1]
                 + block_height(tail, T.STYLE_BY_ID[sub], width_pt))
            if h <= height_pt + 0.5:
                return [(cur, head), (sub, tail)]
        nxt = STEP_DOWN.get(cur)
        if not nxt:
            break
        if fits(text, nxt, width_pt, height_pt):
            return [(nxt, text)]
        cur = nxt
    # Chegou no menor degrau e ainda nao cabe: devolve no menor e deixa o
    # gate do build reclamar. Encolher mais viraria ilegivel.
    return [(cur or style_id, text)]
