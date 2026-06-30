#!/usr/bin/env python3
"""
Miniaturas de PREVIEW dos estilos de tipografia para a GALERIA do ribbon.
Cada thumb mostra a amostra na FONTE/COR/PESO reais, com o glifo escalado
pelo tamanho do estilo (Hero grande -> Caption pequeno) e ponto azul no Hero.
Saida: v3-powerpoint-addin/ribbon/images/prev<Key>.png  (ids batem com customUI14.xml image=)

Uso: python3 tools/gen-style-previews.py
"""
import os
from PIL import Image, ImageDraw, ImageFont

ROSA = (252, 94, 109, 255)   # FC5E6D
AZUL = (67, 106, 225, 255)    # 436AE1
W, H = 88, 52                 # tamanho da miniatura (landscape)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "v3-powerpoint-addin", "ribbon", "images")
os.makedirs(OUT, exist_ok=True)

AVENIR = "/System/Library/Fonts/Avenir Next.ttc"
HELV = "/System/Library/Fonts/Helvetica.ttc"


def font(px, bold):
    for path, idx in ((AVENIR, 2 if bold else 0), (HELV, 1 if bold else 0)):
        try:
            return ImageFont.truetype(path, px, index=idx)
        except Exception:
            continue
    return ImageFont.load_default()


# (id, sizePt, bold, cor, ponto_azul)
STYLES = [
    ("prevHero",       120, True,  ROSA, True),
    ("prevMega",        80, True,  AZUL, False),
    ("prevH1",          60, True,  ROSA, False),
    ("prevLabelSec",    60, True,  ROSA, False),
    ("prevCorpo",       44, False, AZUL, False),
    ("prevH3",          34, True,  ROSA, False),
    ("prevH4",          28, True,  AZUL, False),
    ("prevH5",          24, False, AZUL, False),
    ("prevCorpoPilar",  20, False, AZUL, False),
    ("prevEyebrow",     18, True,  ROSA, False),
    ("prevCaption",     16, False, AZUL, False),
]

# pt (16..120) -> px legivel dentro da thumb (18..40), preservando a hierarquia
def px_for(pt):
    lo_pt, hi_pt, lo_px, hi_px = 16, 120, 18, 40
    t = (pt - lo_pt) / (hi_pt - lo_pt)
    return int(round(lo_px + t * (hi_px - lo_px)))


for name, pt, bold, col, blue_dot in STYLES:
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    f = font(px_for(pt), bold)
    txt = "Aa"
    bb = d.textbbox((0, 0), txt, font=f)
    tw, th = bb[2] - bb[0], bb[3] - bb[1]
    x = (W - tw) / 2 - bb[0]
    if blue_dot:
        x -= 5
    y = (H - th) / 2 - bb[1]
    d.text((x, y), txt, font=f, fill=col)
    if blue_dot:
        r = max(2, px_for(pt) // 9)
        cx, cy = x + tw + 4, y + th
        d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=AZUL)
    img.save(os.path.join(OUT, name + ".png"))

# --- versao QUADRADA p/ os botoes inline da faixa ---
# Glifo BEM maior (pedido do usuario): enche o icone, mantendo uma
# leve hierarquia de tamanho (Hero > Caption). Encolhe pra caber.
SQ = 64
PAD = 4


def target_px_sq(pt):
    # alvo de tamanho por hierarquia (piso alto p/ ler grande)
    lo_pt, hi_pt, lo_px, hi_px = 16, 120, 44, 62
    t = (pt - lo_pt) / (hi_pt - lo_pt)
    return int(round(lo_px + t * (hi_px - lo_px)))


def fit_font(d, txt, pt, bold, extra):
    px = target_px_sq(pt)
    while px > 12:
        f = font(px, bold)
        bb = d.textbbox((0, 0), txt, font=f)
        w, h = bb[2] - bb[0], bb[3] - bb[1]
        if w + extra <= SQ - PAD and h <= SQ - PAD:
            return f, px, bb
        px -= 1
    f = font(px, bold)
    return f, px, d.textbbox((0, 0), txt, font=f)


for name, pt, bold, col, blue_dot in STYLES:
    img = Image.new("RGBA", (SQ, SQ), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    txt = "Aa"
    extra = 0
    f, px, bb = fit_font(d, txt, pt, bold, extra)
    tw, th = bb[2] - bb[0], bb[3] - bb[1]
    if blue_dot:
        # reservar espaco do ponto: refit com 'extra'
        f, px, bb = fit_font(d, txt, pt, bold, max(4, px // 6))
        tw, th = bb[2] - bb[0], bb[3] - bb[1]
    x = (SQ - tw) / 2 - bb[0]
    if blue_dot:
        x -= px // 8
    y = (SQ - th) / 2 - bb[1]
    d.text((x, y), txt, font=f, fill=col)
    if blue_dot:
        r = max(2, px // 8)
        cx, cy = x + tw + r + 1, y + th
        d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=AZUL)
    img.save(os.path.join(OUT, name.replace("prev", "sq") + ".png"))

print("Previews gerados em", os.path.normpath(OUT))
for s in STYLES:
    print("  ", s[0] + ".png  /  " + s[0].replace("prev", "sq") + ".png")
