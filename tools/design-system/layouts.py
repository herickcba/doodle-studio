"""Construtores de slide do design system CBA.

Todo shape sai daqui, e tudo que sai daqui passa pelos tokens. Se um valor nao
esta' em tokens.py, este modulo nao consegue emitir -- e' o mesmo principio do
vba-static-scan: a regra mora no codigo, nao na boa vontade de quem usa.

A modelagem dos arquetipos (uma funcao por familia, recebendo um dict de spec)
segue a da skill cba-bg-design-system-v3; o que muda e' a emissao (python-pptx
em vez de OOXML cru) e, principalmente, a tabela de tokens.

Altura de texto e' MEDIDA com a fonte real (measure.py), nunca estimada: blocos
empilhados sabem quanto o anterior ocupou. Chutar altura foi o que fez a
primeira versao do deck sobrepor texto.
"""

from __future__ import annotations

import copy

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Pt

import measure
import tokens as T

ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
ANCHOR = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}

M = T.SPACING          # {1: 20, 2: 40, 3: 60, 4: 80}


def rgb(hex6: str) -> RGBColor:
    if hex6.upper() not in T.ALLOWED_HEX:
        raise ValueError("cor fora da paleta: %s" % hex6)
    return RGBColor.from_string(hex6.upper())


def style(style_id: str) -> dict:
    s = T.STYLE_BY_ID.get(style_id)
    if not s:
        raise ValueError("estilo fora da escala: %s" % style_id)
    return s


def blank_slide(prs, bg=None):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    fill = sl.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb(bg or T.BEGE)
    return sl


def _resolve_color(st, bg_hex, override=None):
    """A regra de contraste, identica a' do VBA: se a cor do texto for
    exatamente a do fundo, o texto vira branco."""
    color = override or T.TYPE_ROLES[st["role"]]
    if bg_hex and color.upper() == bg_hex.upper():
        return T.BRANCO
    return color


def needed(style_id, text, w):
    """Altura que este texto ocupa nesta largura."""
    return measure.block_height(text, style(style_id), w)


def add_text(slide, style_id, text, x, y, w, h=None, *, bg=None, color=None,
             align="left", anchor="top"):
    """Caixa de texto num estilo do sistema. h=None mede a altura necessaria.

    Aplica de uma vez o que a faixa aplica: fonte, corpo, peso, cor, entrelinha
    exata, caixa alta, espacamento entre letras e o ponto final colorido.
    """
    st = style(style_id)
    if h is None:
        h = needed(style_id, text, w)
    box = slide.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = ANCHOR[anchor]
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

    final_color = _resolve_color(st, bg, color)
    period_color = None
    if st.get("period") is not None and text.rstrip().endswith("."):
        period_color = T.TYPE_ROLES[st["period"]]
        if final_color == T.BRANCO:      # sobre fundo saturado o ponto vai pro rosa
            period_color = T.ROSA

    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ALIGN[align]
        p.line_spacing = st["ent"]
        if period_color and i == len(lines) - 1 and line.rstrip().endswith("."):
            _run(p, line.rstrip()[:-1], st, final_color)
            _run(p, ".", st, period_color)
        else:
            _run(p, line, st, final_color)
    return box


def _run(p, text, st, color_hex):
    r = p.add_run()
    r.text = text
    f = r.font
    f.name = T.FONT
    f.size = Pt(st["size"])
    f.bold = bool(st["bold"])
    f.color.rgb = rgb(color_hex)
    # caixa alta e espacamento entre letras nao tem API no python-pptx:
    # vao direto no rPr, do mesmo jeito que o VBA usa o TextFrame2.
    rPr = r._r.get_or_add_rPr()
    if st.get("caps"):
        rPr.set("cap", "all")
    if st.get("spacing"):
        rPr.set("spc", str(int(st["spacing"] * 100)))   # centesimos de ponto
    return r


class Stack:
    """Empilha blocos de texto medindo cada um. O y anda sozinho."""

    def __init__(self, slide, x, y, w, bg=None):
        self.slide, self.x, self.y, self.w, self.bg = slide, x, y, w, bg

    def add(self, style_id, text, gap=0, **kw):
        if not text:
            return self
        self.y += gap
        w = kw.pop("w", self.w)
        h = needed(style_id, text, w)
        add_text(self.slide, style_id, text, kw.pop("x", self.x), self.y, w, h,
                 bg=self.bg, **kw)
        self.y += h
        return self

    def skip(self, pts):
        self.y += pts
        return self


def add_box(slide, x, y, w, h, *, fill=None, radius="std", line=None, line_w=1):
    """Retangulo do sistema. radius: 'std' (20,62pt), 'pill' ou 'hard'."""
    shape_type = MSO_SHAPE.RECTANGLE if radius == "hard" else MSO_SHAPE.ROUNDED_RECTANGLE
    shp = slide.shapes.add_shape(shape_type, Pt(x), Pt(y), Pt(w), Pt(h))
    if radius != "hard":
        # adj do OOXML: 0..50000 (50000 = pill). python-pptx normaliza /100000.
        if radius == "pill":
            shp.adjustments[0] = 0.5
        else:
            shp.adjustments[0] = min(0.5, (T.RADIUS_PT / (min(w, h) / 2)) * 0.5)
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb(fill)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = rgb(line)
        shp.line.width = Pt(line_w)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False          # anti-vicios: nada de sombra
    shp.text_frame.word_wrap = True
    return shp


def footer(slide, section, number, bg=None):
    y = T.PAGE_H_PT - T.MARGIN_BOTTOM_PT - 16
    if section:
        add_text(slide, "dsCaps12", section, T.MARGIN_LEFT_PT, y, 700, 16, bg=bg)
    add_text(slide, "dsLegenda12", str(number), T.MARGIN_RIGHT_PT - 100, y, 100, 16,
             bg=bg, align="right")


# ---------------------------------------------------------------- arquetipos

def hero_cover(prs, spec):
    bg = spec.get("bg", T.BEGE)
    sl = blank_slide(prs, bg)
    s = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT + 40, T.CONTENT_W_PT * 0.86, bg)
    s.add("dsEyebrow", spec.get("eyebrow"))
    s.add(spec.get("style", "dsHero"), spec["title"], gap=M[3])
    s.add("dsCorpo", spec.get("sub"), gap=M[4], w=T.CONTENT_W_PT * 0.62)
    return sl


def chapter_divider(prs, spec):
    bg = spec.get("bg", T.AZUL)
    sl = blank_slide(prs, bg)
    s = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT + 20, T.CONTENT_W_PT * 0.8, bg)
    s.add("dsBigNumber", spec.get("number"))
    s.add("dsMega", spec["title"], gap=M[2])
    s.add("dsCorpoPilar", spec.get("sub"), gap=M[2], w=T.CONTENT_W_PT * 0.5)
    return sl


def spec_page(prs, spec):
    """Titulo + intro + pares rotulo/valor. O cavalo de batalha da documentacao."""
    bg = spec.get("bg", T.BEGE)
    sl = blank_slide(prs, bg)
    s = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT, T.CONTENT_W_PT, bg)
    s.add("dsCaps12", spec.get("eyebrow"))
    s.add("dsH1", spec["title"], gap=M[1] if spec.get("eyebrow") else 0)
    s.add("dsCorpo", spec.get("intro"), gap=M[2], w=T.CONTENT_W_PT * 0.72)

    rows = spec.get("rows", [])
    if rows:
        y = s.y + M[3]
        col_w = (T.CONTENT_W_PT - M[2]) / 2
        per_col = (len(rows) + 1) // 2
        row_h = max(
            needed("dsCaps12", k, col_w) + needed("dsH4", v, col_w) + M[1]
            for k, v in rows)
        for i, (label, value) in enumerate(rows):
            col, row = divmod(i, per_col)
            rx = T.MARGIN_LEFT_PT + col * (col_w + M[2])
            ry = y + row * row_h
            add_text(sl, "dsCaps12", label, rx, ry, col_w, bg=bg)
            add_text(sl, "dsH4", value, rx, ry + needed("dsCaps12", label, col_w) + 6,
                     col_w, bg=bg)
    if spec.get("note"):
        nw = T.CONTENT_W_PT * 0.66
        nh = needed("dsTexto15", spec["note"], nw)
        add_text(sl, "dsTexto15", spec["note"], T.MARGIN_LEFT_PT,
                 T.PAGE_H_PT - T.MARGIN_BOTTOM_PT - 40 - nh, nw, nh, bg=bg)
    return sl


def type_specimen(prs, spec):
    """Um estilo por slide: amostra no proprio corpo + ficha tecnica."""
    st = style(spec["style_id"])
    bg = spec.get("bg", T.BEGE)
    sl = blank_slide(prs, bg)

    s = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT, T.CONTENT_W_PT, bg)
    s.add("dsCaps12", spec.get("eyebrow", "ESCALA TIPOGRÁFICA"))
    s.add("dsH1", st["label"], gap=M[1])

    # ficha tecnica ancorada na base; a amostra fica com o espaco entre as duas
    cells = [("CORPO", "%gpt" % st["size"]),
             ("PESO", "Bold" if st["bold"] else "Regular"),
             ("COR", T.PALETTE[0 if st["role"] == 0 else 1]["name"].capitalize()),
             ("ENTRELINHA", "%.2fx" % st["ent"]),
             ("PX @1080", "%dpx" % round(T.pt_to_px(st["size"])))]
    use_w = T.CONTENT_W_PT * 0.8
    use_h = needed("dsTexto15", st["use"], use_w)
    fy = T.PAGE_H_PT - T.MARGIN_BOTTOM_PT - 40 - use_h - M[1] - 46
    cw = T.CONTENT_W_PT / len(cells)
    for i, (k, v) in enumerate(cells):
        cx = T.MARGIN_LEFT_PT + i * cw
        add_text(sl, "dsCaps12", k, cx, fy, cw - 10, bg=bg)
        add_text(sl, "dsH4", v, cx, fy + 22, cw - 10, bg=bg)
    add_text(sl, "dsTexto15", st["use"], T.MARGIN_LEFT_PT, fy + 46 + M[1],
             use_w, use_h, bg=bg)

    # a amostra, no corpo real do estilo, centrada no espaco que sobrou
    top = s.y + M[3]
    avail = fy - M[3] - top
    sh = needed(spec["style_id"], spec["sample"], T.CONTENT_W_PT)
    add_text(sl, spec["style_id"], spec["sample"], T.MARGIN_LEFT_PT,
             top + max(0, (avail - sh) / 2), T.CONTENT_W_PT, sh, bg=bg)
    return sl


def swatch_page(prs, spec):
    bg = spec.get("bg", T.BEGE)
    sl = blank_slide(prs, bg)
    pal = spec["palette"]
    s = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT, T.CONTENT_W_PT, bg)
    s.add("dsCaps12", "PALETA")
    s.add("dsH1", pal["name"].capitalize(), gap=M[1])

    sw_y = s.y + M[3]
    sw_w = 620
    sw_h = T.PAGE_H_PT - T.MARGIN_BOTTOM_PT - 40 - sw_y
    add_box(sl, T.MARGIN_LEFT_PT, sw_y, sw_w, sw_h, fill=pal["hex"],
            line=T.AZUL if pal["hex"] == T.BRANCO else None)

    tx = T.MARGIN_LEFT_PT + sw_w + M[3]
    tw = T.MARGIN_RIGHT_PT - tx
    ts = Stack(sl, tx, sw_y, tw, bg)
    for k, v in (("HEX", "#" + pal["hex"]),
                 ("RGB", "%d, %d, %d" % pal["rgb"]),
                 ("SLOT", "pal%d" % spec["index"])):
        ts.add("dsCaps12", k, gap=M[2] if ts.y > sw_y else 0)
        ts.add("dsH4", v, gap=6)
    ts.add("dsTexto15", pal["role"], gap=M[3])
    return sl


def multi_card_grid(prs, spec):
    bg = spec.get("bg", T.BEGE)
    sl = blank_slide(prs, bg)
    s = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT, T.CONTENT_W_PT, bg)
    s.add("dsCaps12", spec.get("eyebrow"))
    s.add("dsH1", spec["title"], gap=M[1])

    cards = spec["cards"]
    n = len(cards)
    cw = (T.CONTENT_W_PT - M[1] * (n - 1)) / n
    cy = s.y + M[3]
    ch = T.PAGE_H_PT - T.MARGIN_BOTTOM_PT - 40 - cy
    card_bg = spec.get("card_bg", T.BRANCO)
    pad = M[2]
    for i, c in enumerate(cards):
        cx = T.MARGIN_LEFT_PT + i * (cw + M[1])
        add_box(sl, cx, cy, cw, ch, fill=card_bg)
        cs = Stack(sl, cx + pad, cy + pad, cw - pad * 2, card_bg)
        cs.add("dsH4", c["title"])
        cs.add("dsCorpoPilar", c.get("body"), gap=M[1])
    return sl


def card_grid_5(prs, spec):
    bg = spec.get("bg", T.BEGE)
    sl = blank_slide(prs, bg)
    s = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT, T.CONTENT_W_PT, bg)
    s.add("dsH1", spec["title"])
    s.add("dsCorpoPilar", spec.get("intro"), gap=M[1], w=T.CONTENT_W_PT * 0.7)

    items = spec["items"]
    cols = spec.get("cols", 4)
    cw = (T.CONTENT_W_PT - M[1] * (cols - 1)) / cols
    rows = (len(items) + cols - 1) // cols
    top = s.y + M[3]
    avail = T.PAGE_H_PT - T.MARGIN_BOTTOM_PT - 40 - top
    ch = (avail - M[1] * (rows - 1)) / rows
    card_bg = spec.get("card_bg", T.BRANCO)
    pad = M[1]
    for i, it in enumerate(items):
        r, c = divmod(i, cols)
        cx = T.MARGIN_LEFT_PT + c * (cw + M[1])
        cy = top + r * (ch + M[1])
        add_box(sl, cx, cy, cw, ch, fill=card_bg)
        cs = Stack(sl, cx + pad, cy + pad, cw - pad * 2, card_bg)
        cs.add("dsCaps12", it.get("kicker"))
        cs.add("dsH4", it["title"], gap=8)
        cs.add("dsTexto15", it.get("body"), gap=10)
    return sl


def pillar_card_dense(prs, spec):
    bg = spec.get("bg", T.BEGE)
    sl = blank_slide(prs, bg)
    half = (T.CONTENT_W_PT - M[4]) / 2
    ls = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT, half, bg)
    ls.add("dsEyebrow", spec.get("eyebrow"))
    ls.add("dsH1", spec["title"], gap=M[1])
    ls.add("dsCorpoPilar", spec.get("body"), gap=M[2])

    rs = Stack(sl, T.MARGIN_LEFT_PT + half + M[4], T.MARGIN_TOP_PT + 30, half, bg)
    for item in spec.get("items", []):
        rs.add("dsH4", item["title"], gap=M[2] if rs.y > T.MARGIN_TOP_PT + 30 else 0)
        rs.add("dsTexto15", item.get("body"), gap=8)
    return sl


def quote_side_image(prs, spec):
    bg = spec.get("bg", T.BEGE)
    sl = blank_slide(prs, bg)
    tw = T.CONTENT_W_PT * 0.5
    s = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT, tw, bg)
    s.add("dsEyebrow", spec.get("eyebrow"))
    s.add("dsH1", spec["title"], gap=M[1])
    s.add("dsCorpo", spec["body"], gap=M[2])

    px = T.MARGIN_LEFT_PT + tw + M[3]
    pw = T.MARGIN_RIGHT_PT - px
    ph = T.PAGE_H_PT - T.MARGIN_TOP_PT - T.MARGIN_BOTTOM_PT - 40
    panel_bg = spec.get("panel_bg", T.AZUL)
    add_box(sl, px, T.MARGIN_TOP_PT, pw, ph, fill=panel_bg)
    if spec.get("panel_text"):
        pstyle = spec.get("panel_style", "dsCorpo")
        pad = M[2]
        th = needed(pstyle, spec["panel_text"], pw - pad * 2)
        add_text(sl, pstyle, spec["panel_text"], px + pad,
                 T.MARGIN_TOP_PT + max(pad, (ph - th) / 2), pw - pad * 2, th,
                 bg=panel_bg)
    return sl


def do_dont(prs, spec):
    bg = spec.get("bg", T.BEGE)
    sl = blank_slide(prs, bg)
    s = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT, T.CONTENT_W_PT, bg)
    s.add("dsH1", spec["title"])
    s.add("dsCorpoPilar", spec.get("intro"), gap=M[1], w=T.CONTENT_W_PT * 0.72)

    cw = (T.CONTENT_W_PT - M[1]) / 2
    cy = s.y + M[3]
    ch = T.PAGE_H_PT - T.MARGIN_BOTTOM_PT - 40 - cy
    pad = M[2]
    for i, (head, items, tone) in enumerate([
            (spec.get("do_title", "SIM"), spec["do"], T.AZUL),
            (spec.get("dont_title", "NUNCA"), spec["dont"], T.ROSA)]):
        cx = T.MARGIN_LEFT_PT + i * (cw + M[1])
        add_box(sl, cx, cy, cw, ch, fill=T.BRANCO)
        cs = Stack(sl, cx + pad, cy + pad, cw - pad * 2, T.BRANCO)
        cs.add("dsCaps12", head, color=tone)
        for it in items:
            cs.add("dsCorpoPilar", it, gap=M[1])
    return sl


def diagram_page(prs, spec):
    """Esquema em escala: os retangulos sao proporcionais de verdade."""
    bg = spec.get("bg", T.BEGE)
    sl = blank_slide(prs, bg)
    lw = T.CONTENT_W_PT * 0.34
    s = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT, lw, bg)
    s.add("dsH1", spec["title"])
    s.add("dsTexto15", spec.get("intro"), gap=M[2])

    dw = T.CONTENT_W_PT * 0.56
    dh = dw * T.PAGE_H_PT / T.PAGE_W_PT
    dx = T.MARGIN_RIGHT_PT - dw
    dy = T.MARGIN_TOP_PT + 60
    scale = dw / T.PAGE_W_PT
    add_box(sl, dx, dy, dw, dh, fill=T.BRANCO, radius="hard", line=T.AZUL)
    if spec.get("show_margin", True):
        m = T.MARGIN_PT * scale
        add_box(sl, dx + m, dy + m, dw - 2 * m,
                dh - m - T.MARGIN_BOTTOM_PT * scale, radius="hard", line=T.ROSA)
    if spec.get("show_module"):
        u = T.SPACING_UNIT_PT * scale
        for i in range(8):
            add_box(sl, dx + T.MARGIN_PT * scale + i * u * 2, dy + dh - u * 3,
                    u, u, fill=T.AZUL, radius="hard")

    ls = Stack(sl, dx, dy + dh + M[2], dw, bg)
    for label, value in spec.get("legend", []):
        ls.add("dsCaps12", label, gap=M[1] if ls.y > dy + dh + M[2] else 0)
        ls.add("dsTexto15", value, gap=4)
    return sl


def closing(prs, spec):
    bg = spec.get("bg", T.AZUL)
    sl = blank_slide(prs, bg)
    w = T.CONTENT_W_PT * 0.82
    th = needed("dsHero", spec["title"], w)
    add_text(sl, "dsHero", spec["title"], T.MARGIN_LEFT_PT,
             (T.PAGE_H_PT - th) / 2 - 40, w, th, bg=bg)
    if spec.get("sub"):
        sw = T.CONTENT_W_PT * 0.6
        sh = needed("dsCorpoPilar", spec["sub"], sw)
        add_text(sl, "dsCorpoPilar", spec["sub"], T.MARGIN_LEFT_PT,
                 T.PAGE_H_PT - T.MARGIN_BOTTOM_PT - 40 - sh, sw, sh, bg=bg)
    return sl


BUILDERS = {
    "hero_cover": hero_cover,
    "chapter_divider": chapter_divider,
    "spec_page": spec_page,
    "type_specimen": type_specimen,
    "swatch_page": swatch_page,
    "multi_card_grid": multi_card_grid,
    "card_grid_5": card_grid_5,
    "pillar_card_dense": pillar_card_dense,
    "quote_side_image": quote_side_image,
    "do_dont": do_dont,
    "diagram_page": diagram_page,
    "closing": closing,
}


def build(prs, spec, index):
    kind = spec["kind"]
    fn = BUILDERS.get(kind)
    if not fn:
        raise ValueError("arquetipo desconhecido: %s" % kind)
    sl = fn(prs, copy.deepcopy(spec))
    if spec.get("footer", True) and kind not in ("hero_cover", "closing"):
        footer(sl, spec.get("section", ""), index, bg=spec.get("bg", T.BEGE))
    return sl
