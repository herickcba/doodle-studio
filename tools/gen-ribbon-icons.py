#!/usr/bin/env python3
"""
Gera os icones PNG do ribbon "Doodle Studio" (estilos de tipografia + entrelinha).
Glifos em Avenir Next nas cores da marca, fundo transparente, 32x32.
Saida: v3-powerpoint-addin/ribbon/images/<id>.png  (ids batem com customUI14.xml image=)

Uso: python3 tools/gen-ribbon-icons.py
"""
import os
from PIL import Image, ImageDraw, ImageFont

ROSA = (252, 94, 109, 255)   # FC5E6D
AZUL = (67, 106, 225, 255)    # 436AE1
SIZE = 32

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "v3-powerpoint-addin", "ribbon", "images")
os.makedirs(OUT, exist_ok=True)

AVENIR = "/System/Library/Fonts/Avenir Next.ttc"
HELV = "/System/Library/Fonts/Helvetica.ttc"


def font(px, bold=True):
    # Avenir Next.ttc: index 0=Regular ... usa bold via index quando possivel.
    for path, idx in ((AVENIR, 2 if bold else 0), (HELV, 1 if bold else 0)):
        try:
            return ImageFont.truetype(path, px, index=idx)
        except Exception:
            continue
    return ImageFont.load_default()


def draw_glyph(text, color, dot=None, bold=True):
    """Glifo centralizado; opcional ponto final colorido (dot=RGBA)."""
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    px = 26 if len(text) == 1 else 18
    f = font(px, bold)
    # mede e centraliza
    bb = d.textbbox((0, 0), text, font=f)
    w, h = bb[2] - bb[0], bb[3] - bb[1]
    x = (SIZE - w) / 2 - bb[0]
    y = (SIZE - h) / 2 - bb[1]
    if dot:
        x -= 3  # abre espaco pro ponto
    d.text((x, y), text, font=f, fill=color)
    if dot:
        # ponto final colorido proximo da baseline, a direita
        r = 3
        cx, cy = x + w + 3, y + h
        d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=dot)
    return img


def draw_lines(color, n=3, arrow=True):
    """Icone de entrelinha: n linhas horizontais + seta vertical dupla."""
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    left, right = 11, 29
    ys = [7 + i * (18 / (n - 1)) for i in range(n)]
    for y in ys:
        d.line((left, y, right, y), fill=color, width=2)
    if arrow:
        ax = 5
        d.line((ax, ys[0], ax, ys[-1]), fill=color, width=2)
        for yy, dy in ((ys[0], 3), (ys[-1], -3)):
            d.line((ax, yy, ax - 2, yy + dy), fill=color, width=2)
            d.line((ax, yy, ax + 2, yy + dy), fill=color, width=2)
    return img


# id -> (glifo, cor, dot)
GLYPHS = {
    "dsHero": ("A", ROSA, AZUL),
    "dsMega": ("M", AZUL, None),
    "dsH1": ("H1", ROSA, None),
    "dsLabelSec": ("L", ROSA, None),
    "dsCorpo": ("B", AZUL, None),
    "dsH3": ("H3", ROSA, None),
    "dsH4": ("H4", AZUL, None),
    "dsH5": ("H5", AZUL, None),
    "dsCorpoPilar": ("P", AZUL, None),
    "dsEyebrow": ("E", ROSA, None),
    "dsCaption": ("c", AZUL, None),
}

for name, (txt, col, dot) in GLYPHS.items():
    draw_glyph(txt, col, dot).save(os.path.join(OUT, name + ".png"))

draw_lines(AZUL, n=3).save(os.path.join(OUT, "entreSel.png"))
draw_lines(AZUL, n=4).save(os.path.join(OUT, "entreTudo.png"))

print("Icones gerados em", os.path.normpath(OUT))
for fn in sorted(os.listdir(OUT)):
    print("  ", fn)
