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
