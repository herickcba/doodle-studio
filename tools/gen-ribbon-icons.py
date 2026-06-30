#!/usr/bin/env python3
"""
Gera os icones PNG do ribbon "CBA Studio" (estilos de tipografia + entrelinha).
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


def draw_step(label, color=AZUL):
    """Icone de step de entrelinha: o multiplicador como numero."""
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    px = 17 if len(label) <= 3 else 13
    f = font(px, bold=True)
    bb = d.textbbox((0, 0), label, font=f)
    w, h = bb[2] - bb[0], bb[3] - bb[1]
    d.text(((SIZE - w) / 2 - bb[0], (SIZE - h) / 2 - bb[1]), label, font=f, fill=color)
    return img


STEPS = {"ent08": "0,8", "ent09": "0,9", "ent095": "0,95",
         "ent10": "1,0", "ent115": "1,15", "ent13": "1,3"}
for sid, lbl in STEPS.items():
    draw_step(lbl).save(os.path.join(OUT, sid + ".png"))


def draw_align_anchor():
    """Linha de ancora (rosa) + barras alinhadas a' esquerda (azul)."""
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    ax = 8
    d.line((ax, 4, ax, 28), fill=ROSA, width=2)
    for y, ln in ((10, 15), (16, 21), (22, 11)):
        d.line((ax + 3, y, ax + 3 + ln, y), fill=AZUL, width=3)
    return img


def draw_rounded(multi=False):
    """Quadrado(s) de cantos arredondados."""
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    try:
        if not multi:
            d.rounded_rectangle((7, 7, 25, 25), radius=6, outline=AZUL, width=3)
        else:
            d.rounded_rectangle((3, 10, 16, 24), radius=4, outline=AZUL, width=2)
            d.rounded_rectangle((16, 5, 29, 19), radius=4, outline=ROSA, width=2)
    except AttributeError:
        d.rectangle((7, 7, 25, 25), outline=AZUL, width=3)
    return img


draw_align_anchor().save(os.path.join(OUT, "alignAnchor.png"))
draw_rounded(False).save(os.path.join(OUT, "rounded.png"))
draw_rounded(True).save(os.path.join(OUT, "roundedAll.png"))


# --- Formas: Tudo rounded / Tirar rounded ---
GRAY = (120, 120, 120, 255)


def draw_round_every():
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    try:
        d.rounded_rectangle((3, 9, 17, 25), radius=4, fill=AZUL)
        d.rounded_rectangle((15, 5, 29, 21), radius=4, fill=ROSA)
    except AttributeError:
        d.rectangle((3, 9, 17, 25), fill=AZUL)
        d.rectangle((15, 5, 29, 21), fill=ROSA)
    return img


def draw_round_off():
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    try:
        d.rounded_rectangle((7, 7, 25, 25), radius=6, outline=GRAY, width=2)
    except AttributeError:
        d.rectangle((7, 7, 25, 25), outline=GRAY, width=2)
    d.line((6, 26, 26, 6), fill=ROSA, width=3)   # slash = remover
    return img


draw_round_every().save(os.path.join(OUT, "roundEvery.png"))
draw_round_off().save(os.path.join(OUT, "roundOff.png"))


# --- Barra de cores (paleta da marca): preenchimento e fonte ---
# 0 rosa, 1 azul, 2 bege, 3 branco, 4 preto  (5 = transparente)
# Botoes agora size="large": chips GRANDES (canvas 64), cheios.
PALETTE = [(253, 94, 109), (67, 106, 225), (238, 236, 230),
           (255, 255, 255), (0, 0, 0)]
CSIZE = 64


def _checker(d, x0, y0, x1, y1, cell=9,
             c1=(255, 255, 255, 255), c2=(198, 198, 198, 255)):
    row = 0
    yy = y0
    while yy < y1:
        col = 0
        xx = x0
        while xx < x1:
            c = c1 if (row + col) % 2 == 0 else c2
            d.rectangle((xx, yy, min(xx + cell, x1), min(yy + cell, y1)), fill=c)
            xx += cell
            col += 1
        yy += cell
        row += 1


def draw_fill_swatch(rgb):
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    box = (4, 4, CSIZE - 5, CSIZE - 5)
    try:
        d.rounded_rectangle(box, radius=11, fill=rgb + (255,), outline=GRAY, width=2)
    except AttributeError:
        d.rectangle(box, fill=rgb + (255,), outline=GRAY, width=2)
    return img


def draw_fill_transp():
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    _checker(d, 5, 5, CSIZE - 5, CSIZE - 5)
    try:
        d.rounded_rectangle((4, 4, CSIZE - 5, CSIZE - 5), radius=11, outline=GRAY, width=2)
    except AttributeError:
        d.rectangle((4, 4, CSIZE - 5, CSIZE - 5), outline=GRAY, width=2)
    return img


def draw_font_swatch(rgb):
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    f = font(46, bold=True)
    txt = "A"
    bb = d.textbbox((0, 0), txt, font=f)
    w, h = bb[2] - bb[0], bb[3] - bb[1]
    x = (CSIZE - w) / 2 - bb[0]
    y = (CSIZE - h) / 2 - bb[1] - 6
    d.text((x, y), txt, font=f, fill=(30, 30, 30, 255))
    d.rectangle((10, CSIZE - 14, CSIZE - 11, CSIZE - 6), fill=rgb + (255,), outline=GRAY, width=1)
    return img


def draw_font_transp():
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    f = font(46, bold=True)
    txt = "A"
    bb = d.textbbox((0, 0), txt, font=f)
    w, h = bb[2] - bb[0], bb[3] - bb[1]
    x = (CSIZE - w) / 2 - bb[0]
    y = (CSIZE - h) / 2 - bb[1] - 6
    d.text((x, y), txt, font=f, fill=(30, 30, 30, 255))
    _checker(d, 10, CSIZE - 14, CSIZE - 11, CSIZE - 6, cell=6)
    d.rectangle((10, CSIZE - 14, CSIZE - 11, CSIZE - 6), outline=GRAY, width=1)
    return img


for i, rgb in enumerate(PALETTE):
    draw_fill_swatch(rgb).save(os.path.join(OUT, "fill%d.png" % i))
    draw_font_swatch(rgb).save(os.path.join(OUT, "font%d.png" % i))
draw_fill_transp().save(os.path.join(OUT, "fillT.png"))
draw_font_transp().save(os.path.join(OUT, "fontT.png"))


print("Icones gerados em", os.path.normpath(OUT))
for fn in sorted(os.listdir(OUT)):
    print("  ", fn)
