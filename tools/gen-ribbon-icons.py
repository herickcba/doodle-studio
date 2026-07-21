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


# --- Contorno (outline): swatch VAZADO (anel colorido sobre anel cinza) ---
def draw_outline_swatch(rgb):
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    box = (4, 4, CSIZE - 5, CSIZE - 5)
    try:
        d.rounded_rectangle(box, radius=12, outline=GRAY, width=7)
        d.rounded_rectangle(box, radius=12, outline=rgb + (255,), width=4)
    except AttributeError:
        d.rectangle(box, outline=GRAY, width=7)
        d.rectangle(box, outline=rgb + (255,), width=4)
    return img


def draw_outline_transp():
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    box = (4, 4, CSIZE - 5, CSIZE - 5)
    try:
        d.rounded_rectangle(box, radius=12, outline=GRAY, width=2)
    except AttributeError:
        d.rectangle(box, outline=GRAY, width=2)
    d.line((10, CSIZE - 10, CSIZE - 10, 10), fill=ROSA, width=4)   # slash = sem linha
    return img


for i, rgb in enumerate(PALETTE):
    draw_outline_swatch(rgb).save(os.path.join(OUT, "outline%d.png" % i))
draw_outline_transp().save(os.path.join(OUT, "outlineT.png"))


# --- Inserir: caixa de texto / rounded box ---
def draw_ins_textbox():
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rectangle((8, 13, CSIZE - 9, CSIZE - 14), outline=AZUL, width=3)
    f = font(30, bold=True)
    bb = d.textbbox((0, 0), "T", font=f)
    w, h = bb[2] - bb[0], bb[3] - bb[1]
    d.text(((CSIZE - w) / 2 - bb[0], (CSIZE - h) / 2 - bb[1] - 4), "T", font=f, fill=AZUL)
    return img


def draw_ins_roundbox():
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    box = (8, 15, CSIZE - 9, CSIZE - 12)
    try:
        d.rounded_rectangle(box, radius=11, outline=AZUL, width=3)
    except AttributeError:
        d.rectangle(box, outline=AZUL, width=3)
    cx, cy = CSIZE - 16, 15
    d.line((cx - 7, cy, cx + 7, cy), fill=ROSA, width=3)
    d.line((cx, cy - 7, cx, cy + 7), fill=ROSA, width=3)
    return img


draw_ins_textbox().save(os.path.join(OUT, "insTextBox.png"))
draw_ins_roundbox().save(os.path.join(OUT, "insRoundBox.png"))


# --- Auditoria ---
def draw_audit_scan():
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.ellipse((10, 10, 40, 40), outline=AZUL, width=4)
    d.line((37, 37, CSIZE - 11, CSIZE - 11), fill=AZUL, width=6)
    return img


def draw_audit_fonts():
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    f = font(30, bold=True)
    bb = d.textbbox((0, 0), "Aa", font=f)
    w, h = bb[2] - bb[0], bb[3] - bb[1]
    d.text(((CSIZE - w) / 2 - bb[0] - 4, (CSIZE - h) / 2 - bb[1] - 8), "Aa", font=f, fill=(30, 30, 30, 255))
    d.line((CSIZE - 28, CSIZE - 18, CSIZE - 19, CSIZE - 9), fill=ROSA, width=4)
    d.line((CSIZE - 19, CSIZE - 9, CSIZE - 6, CSIZE - 27), fill=ROSA, width=4)
    return img


def draw_audit_colors():
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.ellipse((9, 12, 27, 30), fill=ROSA)
    d.ellipse((23, 12, 41, 30), fill=AZUL)
    d.ellipse((37, 12, 55, 30), fill=(0, 0, 0, 255))
    d.line((16, 46, 26, 56), fill=ROSA, width=4)
    d.line((26, 56, 47, 37), fill=ROSA, width=4)
    return img


draw_audit_scan().save(os.path.join(OUT, "auditScan.png"))
draw_audit_fonts().save(os.path.join(OUT, "auditFonts.png"))
draw_audit_colors().save(os.path.join(OUT, "auditColors.png"))


# --- Texto: Negrito / Colar texto / Caps / Expandir (hi-res 64px) ---
INK = (32, 32, 38, 255)


def draw_bold():
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    f = font(52, bold=True)
    bb = d.textbbox((0, 0), "B", font=f)
    w, h = bb[2] - bb[0], bb[3] - bb[1]
    d.text(((CSIZE - w) / 2 - bb[0], (CSIZE - h) / 2 - bb[1] - 4), "B", font=f, fill=INK)
    return img


def draw_paste():
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    try:
        d.rounded_rectangle((14, 13, 50, 55), radius=6, outline=AZUL, width=4)
        d.rounded_rectangle((25, 7, 39, 18), radius=3, fill=AZUL)
    except AttributeError:
        d.rectangle((14, 13, 50, 55), outline=AZUL, width=4)
        d.rectangle((25, 7, 39, 18), fill=AZUL)
    d.line((21, 30, 43, 30), fill=AZUL, width=3)
    d.line((21, 38, 43, 38), fill=AZUL, width=3)
    d.line((21, 46, 35, 46), fill=AZUL, width=3)
    return img


def draw_caps():
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    f = font(32, bold=True)
    bb = d.textbbox((0, 0), "AA", font=f)
    w, h = bb[2] - bb[0], bb[3] - bb[1]
    d.text(((CSIZE - w) / 2 - bb[0], (CSIZE - h) / 2 - bb[1] - 9), "AA", font=f, fill=INK)
    for x0 in (12, 28, 44):
        d.line((x0, CSIZE - 11, x0 + 9, CSIZE - 11), fill=ROSA, width=3)
    return img


def draw_expand():
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    try:
        d.rounded_rectangle((23, 23, 41, 41), radius=4, outline=AZUL, width=3)
    except AttributeError:
        d.rectangle((23, 23, 41, 41), outline=AZUL, width=3)
    corners = [
        (8, 8, 18, 18, (8, 16), (16, 8)),
        (56, 8, 46, 18, (56, 16), (48, 8)),
        (8, 56, 18, 46, (8, 48), (16, 56)),
        (56, 56, 46, 46, (56, 48), (48, 56)),
    ]
    for ax, ay, bx, by, h1, h2 in corners:
        d.line((ax, ay, bx, by), fill=ROSA, width=3)
        d.line((ax, ay, h1[0], h1[1]), fill=ROSA, width=3)
        d.line((ax, ay, h2[0], h2[1]), fill=ROSA, width=3)
    return img


draw_bold().save(os.path.join(OUT, "dsBold.png"))
draw_paste().save(os.path.join(OUT, "dsPaste.png"))
draw_caps().save(os.path.join(OUT, "dsCaps.png"))
draw_expand().save(os.path.join(OUT, "dsExpand.png"))


# --- Alinhar / Distribuir (hi-res, mesma familia do grupo Texto) ---
def draw_align_left():
    """Barra-guia vertical (rosa) a' esquerda + barras alinhadas (azul)."""
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.line((10, 8, 10, CSIZE - 8), fill=ROSA, width=3)
    for y, ln in ((16, 34), (29, 46), (42, 26)):
        d.rectangle((14, y - 4, 14 + ln, y + 4), fill=AZUL)
    return img


def draw_align_top():
    """Barra-guia horizontal (rosa) no topo + barras alinhadas (azul)."""
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.line((8, 10, CSIZE - 8, 10), fill=ROSA, width=3)
    for x, ln in ((16, 34), (29, 46), (42, 26)):
        d.rectangle((x - 4, 14, x + 4, 14 + ln), fill=AZUL)
    return img


def draw_dist_h():
    """Tres barras verticais com espacamento igual (azul) + setas (rosa)."""
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    for x in (12, 30, 48):
        d.rectangle((x - 4, 14, x + 4, CSIZE - 14), fill=AZUL)
    d.line((20, CSIZE - 8, 40, CSIZE - 8), fill=ROSA, width=2)
    return img


draw_align_left().save(os.path.join(OUT, "alignLeft.png"))
draw_align_top().save(os.path.join(OUT, "alignTop.png"))
draw_dist_h().save(os.path.join(OUT, "distH.png"))


# --- Padrões (Brand Standards): engrenagem / colar-aplicar ---
def draw_config_open():
    """Engrenagem azul (abrir a página de padrões)."""
    import math
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    cx, cy, rO, rI = 32, 32, 22, 14
    teeth = 8
    pts = []
    for i in range(teeth * 2):
        ang = math.pi * i / teeth
        r = rO if (i % 2 == 0) else rI
        pts.append((cx + r * math.cos(ang), cy + r * math.sin(ang)))
    d.polygon(pts, fill=AZUL)
    d.ellipse((cx - 8, cy - 8, cx + 8, cy + 8), fill=(0, 0, 0, 0))
    d.ellipse((cx - 8, cy - 8, cx + 8, cy + 8), outline=(255, 255, 255, 255), width=3)
    return img


def draw_config_apply():
    """Prancheta com check rosa (colar/aplicar a config)."""
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    try:
        d.rounded_rectangle((14, 13, 50, 55), radius=6, outline=AZUL, width=4)
        d.rounded_rectangle((25, 7, 39, 18), radius=3, fill=AZUL)
    except AttributeError:
        d.rectangle((14, 13, 50, 55), outline=AZUL, width=4)
        d.rectangle((25, 7, 39, 18), fill=AZUL)
    d.line((22, 36, 30, 45), fill=ROSA, width=5)
    d.line((30, 45, 44, 26), fill=ROSA, width=5)
    return img


draw_config_open().save(os.path.join(OUT, "configOpen.png"))
draw_config_apply().save(os.path.join(OUT, "configApply.png"))


# --- Leva 3: raio (dropdown), page size, auditoria de imagens ---
def draw_radius_pick():
    """Canto arredondado em destaque + cotinha do raio."""
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    try:
        d.rounded_rectangle((10, 14, 54, 52), radius=14, outline=AZUL, width=4)
    except AttributeError:
        d.rectangle((10, 14, 54, 52), outline=AZUL, width=4)
    # seta apontando pro canto arredondado (o "raio")
    d.line((36, 38, 17, 20), fill=ROSA, width=3)
    d.line((17, 20, 24, 21), fill=ROSA, width=3)
    d.line((17, 20, 18, 27), fill=ROSA, width=3)
    return img


def draw_page_size():
    """Retangulo 16:9 com setas de expansao nos cantos."""
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rectangle((14, 20, 50, 44), outline=AZUL, width=4)
    for ax, ay, bx, by in ((8, 14, 16, 22), (56, 50, 48, 42)):
        d.line((ax, ay, bx, by), fill=ROSA, width=3)
        d.line((ax, ay, ax + (4 if ax < 32 else -4), ay), fill=ROSA, width=3)
        d.line((ax, ay, ax, ay + (4 if ay < 32 else -4)), fill=ROSA, width=3)
    return img


def draw_img_audit():
    """Moldura de foto (montanha + sol) + lupa."""
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rectangle((8, 12, 44, 42), outline=AZUL, width=3)
    d.ellipse((14, 17, 21, 24), fill=ROSA)                     # sol
    d.polygon([(11, 39), (24, 25), (33, 34), (38, 29), (41, 39)], fill=AZUL)  # montanhas
    d.ellipse((36, 34, 52, 50), outline=ROSA, width=4)          # lupa
    d.line((50, 48, 58, 56), fill=ROSA, width=5)
    return img


draw_radius_pick().save(os.path.join(OUT, "radiusPick.png"))
draw_page_size().save(os.path.join(OUT, "pageSize.png"))
draw_img_audit().save(os.path.join(OUT, "imgAudit.png"))


def draw_rad_item(label, rpx):
    """Item do dropdown de raio: quadradinho com o canto no raio + rotulo baked-in
    (no Mac, itens de gallery so renderizam com imagem)."""
    W, H = 88, 36
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    r = 8 if rpx == 0 else max(2, min(14, int(rpx * 28 / 50)))
    try:
        d.rounded_rectangle((6, 4, 34, 32), radius=r, outline=AZUL, width=3)
    except AttributeError:
        d.rectangle((6, 4, 34, 32), outline=AZUL, width=3)
    f = font(15, bold=True)
    bb = d.textbbox((0, 0), label, font=f)
    d.text((42, (H - (bb[3] - bb[1])) / 2 - bb[1]), label, font=f,
           fill=ROSA if rpx == 0 else (30, 30, 30, 255))
    return img


for _rid, _lab, _rv in (("radItemP", "Padrão", 0), ("radItem10", "10 px", 10),
                        ("radItem15", "15 px", 15), ("radItem20", "20 px", 20),
                        ("radItem25", "25 px", 25), ("radItem30", "30 px", 30),
                        ("radItem40", "40 px", 40), ("radItem50", "50 px", 50)):
    draw_rad_item(_lab, _rv).save(os.path.join(OUT, _rid + ".png"))


def draw_anchor_pick():
    """Icone do dropdown de medida da ancora: barra esquerda + caixa encostada."""
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.line((14, 12, 14, 52), fill=AZUL, width=5)          # barra de ancoragem
    d.rectangle((20, 24, 48, 44), outline=ROSA, width=4)  # caixa encostada
    d.line((40, 34, 25, 34), fill=ROSA, width=3)          # seta empurrando <-
    d.polygon([(23, 34), (29, 30), (29, 38)], fill=ROSA)
    return img


def draw_anc_item(label, is_default=False):
    """Item do dropdown de medida: barra + caixinha + rotulo baked-in
    (no Mac, itens de gallery so renderizam com imagem)."""
    W, H = 88, 36
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.line((9, 5, 9, 31), fill=AZUL, width=3)
    d.rectangle((13, 12, 30, 24), outline=ROSA if is_default else AZUL, width=2)
    f = font(15, bold=True)
    bb = d.textbbox((0, 0), label, font=f)
    d.text((38, (H - (bb[3] - bb[1])) / 2 - bb[1]), label, font=f,
           fill=ROSA if is_default else (30, 30, 30, 255))
    return img


draw_anchor_pick().save(os.path.join(OUT, "anchorPick.png"))
for _aid, _alab, _adef in (("ancItemD", "Padrão", True), ("ancItem50", "0,5 cm", False),
                           ("ancItem100", "1 cm", False), ("ancItem150", "1,5 cm", False),
                           ("ancItem200", "2 cm", False), ("ancItem250", "2,5 cm", False),
                           ("ancItem300", "3 cm", False), ("ancItem400", "4 cm", False),
                           ("ancItem500", "5 cm", False)):
    draw_anc_item(_alab, _adef).save(os.path.join(OUT, _aid + ".png"))


# ---- Âncora topo (botão + dropdown) — espelho vertical dos de esquerda ----
def draw_align_anchor_top():
    """Linha de âncora no topo (rosa) + barras alinhadas ao topo (azul)."""
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    ay = 8
    d.line((4, ay, 28, ay), fill=ROSA, width=2)
    for x, ln in ((10, 15), (16, 21), (22, 11)):
        d.line((x, ay + 3, x, ay + 3 + ln), fill=AZUL, width=3)
    return img


def draw_anchor_top_pick():
    """Dropdown de medida da âncora topo: barra no topo + caixa encostada embaixo."""
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.line((12, 14, 52, 14), fill=AZUL, width=5)          # barra de ancoragem (topo)
    d.rectangle((22, 20, 42, 48), outline=ROSA, width=4)  # caixa encostada
    d.line((32, 40, 32, 25), fill=ROSA, width=3)          # seta empurrando p/ cima
    d.polygon([(32, 23), (28, 29), (36, 29)], fill=ROSA)
    return img


draw_align_anchor_top().save(os.path.join(OUT, "alignAnchorTop.png"))
draw_anchor_top_pick().save(os.path.join(OUT, "anchorTopPick.png"))


# ---- Linhas-guia (toggle): moldura vermelha + colunas ----
def draw_guides():
    VERM = (237, 28, 36, 255)
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rectangle((5, 5, 26, 26), outline=VERM, width=2)   # margem
    d.line((13, 6, 13, 25), fill=VERM, width=1)          # colunas
    d.line((18, 6, 18, 25), fill=VERM, width=1)
    return img


draw_guides().save(os.path.join(OUT, "guides.png"))


# ---- Sobre (versao instalada): "i" num circulo azul ----
def draw_about():
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.ellipse((10, 10, 54, 54), outline=AZUL, width=4)
    d.ellipse((29, 18, 35, 24), fill=AZUL)               # pingo do i
    d.rectangle((29, 28, 35, 46), fill=AZUL)             # haste
    return img


draw_about().save(os.path.join(OUT, "about.png"))


# ---- Crop de imagem: as duas alças em L cruzadas (símbolo universal) ----
def draw_crop():
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    # retângulo da foto, ao fundo, em cinza claro
    d.rectangle((18, 18, 46, 46), fill=(232, 235, 245, 255))
    # alça superior-esquerda (azul)
    d.rectangle((14, 6, 20, 50), fill=AZUL)
    d.rectangle((14, 44, 58, 50), fill=AZUL)
    # alça inferior-direita (rosa)
    d.rectangle((44, 14, 50, 58), fill=ROSA)
    d.rectangle((6, 14, 50, 20), fill=ROSA)
    return img


draw_crop().save(os.path.join(OUT, "cropPic.png"))


# ---- Alinhar texto: 4 linhas (2a em rosa) alinhadas à esquerda / ao centro ----
def _text_lines(centered):
    img = Image.new("RGBA", (CSIZE, CSIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    ys = (16, 26, 36, 46)
    widths = (34, 22, 30, 20)   # linhas de comprimentos variados
    for i, (y, w) in enumerate(zip(ys, widths)):
        x0 = (CSIZE - w) / 2 if centered else 14
        col = ROSA if i == 1 else AZUL
        d.rounded_rectangle((x0, y, x0 + w, y + 5), radius=2, fill=col)
    return img


_text_lines(False).save(os.path.join(OUT, "alignTextLeft.png"))
_text_lines(True).save(os.path.join(OUT, "alignTextCenter.png"))


print("Icones gerados em", os.path.normpath(OUT))
for fn in sorted(os.listdir(OUT)):
    print("  ", fn)
