"""Construtores de slide do design system CBA 2.0.

Todo shape sai daqui, e tudo que sai daqui passa pelos tokens.

O que mudou do 1.0, tudo vindo da revisao feita a mao no PPTX:

  ritmo vertical   gaps escolhem um degrau da escala (20/40/60/80/120) e o
                   degrau tem significado. O 1.0 usava 1x em lugar de 4x.
  altura de caixa  card = padding + o MAIOR conteudo do grupo + padding.
                   Nunca esticar ate' o rodape.
  zonas            conteudo curto se distribui em zona superior + faixa
                   inferior, em vez de empilhar tudo no topo.
  lista            itens de mesma funcao vao numa caixa so', com quebra de
                   paragrafo. Uma caixa por item e' um inferno de editar.
  texto que estoura primeira frase no estilo pedido, resto um degrau abaixo.
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

# atalhos da escala vertical
TIGHT, ITEM, BLOCK, SECTION, ZONE = (T.V_TIGHT, T.V_ITEM, T.V_BLOCK,
                                     T.V_SECTION, T.V_ZONE)

SATURATED = (T.AZUL, T.ROSA)   # fundos que exigem tipo branco


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
    """Regra de contraste, identica a' do VBA: cor do texto igual a' do fundo
    vira branco."""
    color = override or T.TYPE_ROLES[st["role"]]
    if bg_hex and color.upper() == bg_hex.upper():
        return T.BRANCO
    return color


def needed(style_id, text, w):
    return measure.block_height(text, style(style_id), w)


def add_text(slide, style_id, text, x, y, w, h=None, *, bg=None, color=None,
             align="left", anchor="top", space_after=0):
    """Caixa de texto num estilo do sistema. h=None mede a altura necessaria."""
    st = style(style_id)
    if h is None:
        h = needed(style_id, text, w)
        if space_after:
            h += space_after * (text.count("\n"))
    box = slide.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = ANCHOR[anchor]
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

    final_color = _resolve_color(st, bg, color)
    period_color = T.TYPE_ROLES[st["period"]] if st.get("period") is not None else None

    # Regra de MARCA do Statement (bg_aware, herdada do VBA): sobre o azul da
    # marca o texto vira branco e o ponto, rosa. Nao e' legibilidade, e' marca:
    # rosa sobre azul seria legivel, mas nao e' o que a marca faz.
    if st.get("bg_aware") and bg and bg.upper() == T.AZUL and not color:
        final_color = T.BRANCO
        if period_color:
            period_color = T.ROSA

    if period_color and text.rstrip().endswith("."):
        # o ponto tambem obedece o contraste. Sem isto o ponto azul do Statement
        # desaparecia sobre fundo azul: o texto era protegido e ele nao.
        if final_color == T.BRANCO and period_color.upper() == T.BRANCO:
            period_color = T.ROSA
        elif bg and period_color.upper() == bg.upper():
            period_color = T.BRANCO
    else:
        period_color = None

    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ALIGN[align]
        p.line_spacing = st["ent"]
        if space_after and i < len(lines) - 1:
            p.space_after = Pt(space_after)
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
    rPr = r._r.get_or_add_rPr()
    if st.get("caps"):
        rPr.set("cap", "all")
    if st.get("spacing"):
        rPr.set("spc", str(int(st["spacing"] * 100)))
    return r


def add_list(slide, style_id, items, x, y, w, *, bg=None, color=None,
             gap=None):
    """Lista numa CAIXA SO'. Itens viram paragrafos, separados por espaco de
    paragrafo. Uma caixa por item era o vicio do 1.0."""
    gap = ITEM if gap is None else gap
    text = "\n".join(items)
    return add_text(slide, style_id, text, x, y, w, bg=bg, color=color,
                    space_after=gap)


def list_height(style_id, items, w, gap=None):
    gap = ITEM if gap is None else gap
    return (sum(needed(style_id, it, w) for it in items)
            + gap * max(0, len(items) - 1))


class Stack:
    """Empilha blocos medindo cada um. O y anda sozinho."""

    def __init__(self, slide, x, y, w, bg=None):
        self.slide, self.x, self.y, self.w, self.bg = slide, x, y, w, bg
        self.top = y

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

    def fit(self, style_id, text, max_bottom, gap=0, **kw):
        """Como add(), mas encaixa na altura disponivel: primeira frase no
        estilo pedido, resto um degrau abaixo."""
        if not text:
            return self
        self.y += gap
        w = kw.pop("w", self.w)
        parts = measure.fit_text(text, style_id, w, max_bottom - self.y)
        for i, (sid, txt) in enumerate(parts):
            h = needed(sid, txt, w)
            add_text(self.slide, sid, txt, kw.get("x", self.x), self.y, w, h,
                     bg=self.bg, **{k: v for k, v in kw.items() if k != "x"})
            self.y += h
            if i < len(parts) - 1:
                self.y += TIGHT
        return self

    def add_list(self, style_id, items, gap=0, item_gap=None, **kw):
        if not items:
            return self
        self.y += gap
        w = kw.pop("w", self.w)
        add_list(self.slide, style_id, items, kw.pop("x", self.x), self.y, w,
                 bg=self.bg, gap=item_gap, **kw)
        self.y += list_height(style_id, items, w, item_gap)
        return self

    def skip(self, pts):
        self.y += pts
        return self


def add_box(slide, x, y, w, h, *, fill=None, radius="std", line=None, line_w=1):
    """Retangulo do sistema. radius: 'std' (20pt), 'pill' ou 'hard'."""
    shape_type = MSO_SHAPE.RECTANGLE if radius == "hard" else MSO_SHAPE.ROUNDED_RECTANGLE
    shp = slide.shapes.add_shape(shape_type, Pt(x), Pt(y), Pt(w), Pt(h))
    if radius != "hard":
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
    """Rodape de localizacao. Fica na ultima linha da area util."""
    y = T.CONTENT_BOTTOM_PT - 16
    if section:
        add_text(slide, "dsCaps12", section, T.MARGIN_LEFT_PT, y, 700, 16, bg=bg)
    add_text(slide, "dsLegenda12", str(number), T.MARGIN_RIGHT_PT - 100, y, 100,
             16, bg=bg, align="right")


def distribute(top, bottom, count, block_h, min_gap=None):
    """Passo entre `count` blocos de altura `block_h` para que o ULTIMO termine
    em `bottom`. Nunca aperta abaixo de min_gap.

    Sem descontar a altura do ultimo bloco, o passo empurra o final para fora da
    pagina: foi o que gerou 41 transbordos na primeira versao do 2.0.
    """
    min_gap = ITEM if min_gap is None else min_gap
    if count <= 1:
        return block_h + min_gap
    span = bottom - top - block_h
    ideal = span / (count - 1)
    return max(block_h + min_gap, min(ideal, block_h + ZONE))


def anchor_bottom(top, bottom, count, pitch, block_h):
    """Se o grupo nao preenche o espaco, desce inteiro em vez de esticar.
    Espalhar alem de 120pt desconecta os blocos em vez de dar ritmo."""
    used = (count - 1) * pitch + block_h
    return max(top, bottom - used)


def _card_heights(cards, cw, pad, blocks):
    """Altura unica dos cards do grupo: o MAIOR conteudo manda."""
    inner = cw - pad * 2
    tallest = 0
    for c in cards:
        h, first = 0, True
        for sid, key, gap in blocks:
            txt = c.get(key)
            if not txt:
                continue
            if not first:
                h += gap
            h += needed(sid, txt, inner)
            first = False
        tallest = max(tallest, h)
    return tallest + pad * 2


# ---------------------------------------------------------------- arquetipos

def hero_cover(prs, spec):
    """Capa. Fundo SEMPRE saturado: azul (preferencial) ou rosa."""
    bg = spec.get("bg", T.AZUL)
    if bg not in SATURATED:
        raise ValueError("capa exige fundo azul ou rosa, recebeu %s" % bg)
    sl = blank_slide(prs, bg)
    w = T.col_span(3)
    s = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT, w, bg)
    s.add("dsLegenda12", spec.get("eyebrow"))
    s.fit(spec.get("style", "dsHero"), spec["title"],
          T.CONTENT_BOTTOM_PT - 160, gap=SECTION)
    if spec.get("sub"):
        sw = T.col_span(2)
        sh = needed("dsCorpoPilar", spec["sub"], sw)
        add_text(sl, "dsCorpoPilar", spec["sub"], T.MARGIN_LEFT_PT,
                 T.CONTENT_BOTTOM_PT - sh, sw, sh, bg=bg)
    return sl


def chapter_divider(prs, spec):
    """Abertura de bloco. Numero em '1.', nunca '01'."""
    bg = spec.get("bg", T.AZUL)
    sl = blank_slide(prs, bg)
    s = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT, T.col_span(3), bg)
    s.add("dsBigNumber", spec.get("number"))
    s.add("dsMega", spec["title"], gap=SECTION)
    if spec.get("sub"):
        sw = T.col_span(2)
        sh = needed("dsCorpoPilar", spec["sub"], sw)
        add_text(sl, "dsCorpoPilar", spec["sub"], T.MARGIN_LEFT_PT,
                 T.CONTENT_BOTTOM_PT - sh, sw, sh, bg=bg)
    return sl


def spec_page(prs, spec):
    """Titulo + intro + pares rotulo/valor, distribuidos ate' a base."""
    bg = spec.get("bg", T.BEGE)
    sl = blank_slide(prs, bg)
    s = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT, T.CONTENT_W_PT, bg)
    s.add("dsCaps12", spec.get("eyebrow"))
    s.add("dsH1", spec["title"], gap=ITEM if spec.get("eyebrow") else 0)
    if spec.get("intro"):
        s.fit("dsCorpo", spec["intro"], T.CONTENT_BOTTOM_PT * 0.55,
              gap=BLOCK, w=T.col_span(3))

    rows = spec.get("rows", [])
    if rows:
        # ate' 8 pares cabem em 2 blocos de 2 colunas; acima disso, 4 blocos
        # de 1 coluna. Apertar a linha em vez de abrir coluna era o que fazia
        # a tabela de entrelinha estourar a pagina.
        ncols = 2 if len(rows) <= 8 else 4
        span = T.COLUMNS // ncols
        col_w = T.col_span(span)
        per_col = (len(rows) + ncols - 1) // ncols
        lab_h = max(needed("dsCaps12", k, col_w) for k, _ in rows)
        val_h = max(needed("dsH4", v, col_w) for _, v in rows)
        row_h = lab_h + TIGHT + val_h
        # distribui as linhas ate' a base da area util
        top = s.y + ZONE
        bottom = T.CONTENT_BOTTOM_PT - SECTION
        pitch = distribute(top, bottom, per_col, row_h)
        top = anchor_bottom(top, bottom, per_col, pitch, row_h)
        for i, (label, value) in enumerate(rows):
            col, row = divmod(i, per_col)
            rx = T.col_x(col * span)
            ry = top + row * pitch
            add_text(sl, "dsCaps12", label, rx, ry, col_w, bg=bg)
            add_text(sl, "dsH4", value, rx, ry + lab_h + TIGHT, col_w, bg=bg)
    return sl


def type_specimen(prs, spec):
    """Um estilo por slide: amostra no proprio corpo + ficha tecnica."""
    st = style(spec["style_id"])
    bg = spec.get("bg", T.BEGE)
    sl = blank_slide(prs, bg)

    s = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT, T.CONTENT_W_PT, bg)
    s.add("dsCaps12", spec.get("eyebrow", "ESCALA TIPOGRÁFICA"))
    s.add("dsH1", st["label"], gap=ITEM)

    cells = [("CORPO", "%gpt" % st["size"]),
             ("PESO", "Bold" if st["bold"] else "Regular"),
             ("COR", T.PALETTE[0 if st["role"] == 0 else 1]["name"].capitalize()),
             ("ENTRELINHA", "%.2fx" % st["ent"]),
             ("PX @1080", "%dpx" % round(T.pt_to_px(st["size"])))]
    use_w = T.col_span(3)
    use_h = needed("dsCorpoPilar", st["use"], use_w)
    band_h = 12 + TIGHT + 28
    fy = T.CONTENT_BOTTOM_PT - 40 - use_h - BLOCK - band_h
    cw = T.CONTENT_W_PT / len(cells)
    for i, (k, v) in enumerate(cells):
        cx = T.MARGIN_LEFT_PT + i * cw
        add_text(sl, "dsCaps12", k, cx, fy, cw - ITEM, bg=bg)
        add_text(sl, "dsH4", v, cx, fy + 12 + TIGHT, cw - ITEM, bg=bg)
    add_text(sl, "dsCorpoPilar", st["use"], T.MARGIN_LEFT_PT,
             fy + band_h + BLOCK, use_w, use_h, bg=bg)

    top = s.y + BLOCK
    avail = fy - BLOCK - top
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
    s.add("dsH1", pal["name"].capitalize(), gap=ITEM)

    sw_y = s.y + BLOCK
    sw_w = T.col_span(2)
    sw_h = T.CONTENT_BOTTOM_PT - 40 - sw_y
    add_box(sl, T.MARGIN_LEFT_PT, sw_y, sw_w, sw_h, fill=pal["hex"],
            line=T.AZUL if pal["hex"] == T.BRANCO else None)

    tx = T.col_x(2)
    tw = T.col_span(2)
    ts = Stack(sl, tx, sw_y, tw, bg)
    for i, (k, v) in enumerate([("HEX", "#" + pal["hex"]),
                                ("RGB", "%d, %d, %d" % pal["rgb"]),
                                ("SLOT", "pal%d" % spec["index"])]):
        ts.add("dsCaps12", k, gap=BLOCK if i else 0)
        ts.add("dsH4", v, gap=TIGHT)
    ts.add("dsCorpoPilar", pal["role"], gap=SECTION)
    return sl


def multi_card_grid(prs, spec):
    """2 a 4 cards. Altura pelo MAIOR conteudo, nunca ate' o rodape."""
    bg = spec.get("bg", T.BEGE)
    sl = blank_slide(prs, bg)
    s = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT, T.CONTENT_W_PT, bg)
    s.add("dsCaps12", spec.get("eyebrow"))
    s.add("dsH1", spec["title"], gap=ITEM)

    cards = spec["cards"]
    n = len(cards)
    span = T.COLUMNS // n if T.COLUMNS % n == 0 else 1
    cw = T.col_span(span) if T.COLUMNS % n == 0 else \
        (T.CONTENT_W_PT - T.GUTTER_PT * (n - 1)) / n
    pad = T.BOX_PAD_PT
    blocks = [("dsH4", "title", 0), ("dsCorpoPilar", "body", ITEM)]
    ch = _card_heights(cards, cw, pad, blocks)
    cy = s.y + BLOCK
    card_bg = spec.get("card_bg", T.BRANCO)
    for i, c in enumerate(cards):
        cx = (T.col_x(i * span) if T.COLUMNS % n == 0
              else T.MARGIN_LEFT_PT + i * (cw + T.GUTTER_PT))
        add_box(sl, cx, cy, cw, ch, fill=card_bg)
        cs = Stack(sl, cx + pad, cy + pad, cw - pad * 2, card_bg)
        cs.add("dsH4", c["title"])
        cs.add("dsCorpoPilar", c.get("body"), gap=ITEM)
    return sl


def card_grid_5(prs, spec):
    """Grade densa de itens curtos, em cards."""
    return _grid(prs, spec, cards=True)


def grid_plain(prs, spec):
    """A mesma grade, SEM caixa. Alternativa mais limpa quando o item e' curto."""
    return _grid(prs, spec, cards=False)


def _grid(prs, spec, cards):
    bg = spec.get("bg", T.BEGE)
    sl = blank_slide(prs, bg)
    s = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT, T.CONTENT_W_PT, bg)
    s.add("dsH1", spec["title"])
    s.add("dsCorpoPilar", spec.get("intro"), gap=ITEM, w=T.col_span(3))

    items = spec["items"]
    cols = spec.get("cols", 4)
    span = max(1, T.COLUMNS // cols)
    cw = T.col_span(span) if cols <= T.COLUMNS else \
        (T.CONTENT_W_PT - T.GUTTER_PT * (cols - 1)) / cols
    pad = T.BOX_PAD_PT if cards else 0
    rows = (len(items) + cols - 1) // cols
    top = s.y + BLOCK
    bottom = T.CONTENT_BOTTOM_PT - 40

    # Grade densa e' conteudo MICRO: se o corpo em Apoio 20 nao couber,
    # desce para Texto 15, que existe pra isso (item 4 da revisao).
    body_style = "dsCorpoPilar"
    for candidate in ("dsCorpoPilar", "dsTexto15"):
        blocks = [("dsCaps12", "kicker", 0), ("dsH4", "title", TIGHT),
                  (candidate, "body", TIGHT)]
        ch = _card_heights(items, cw, pad, blocks)
        body_style = candidate
        if top + rows * ch + (rows - 1) * T.GUTTER_PT <= bottom:
            break
    pitch = distribute(top, bottom, rows, ch, T.GUTTER_PT)
    top = anchor_bottom(top, bottom, rows, pitch, ch)
    card_bg = spec.get("card_bg", T.BRANCO)
    for i, it in enumerate(items):
        r, c = divmod(i, cols)
        cx = (T.col_x(c * span) if cols <= T.COLUMNS
              else T.MARGIN_LEFT_PT + c * (cw + T.GUTTER_PT))
        cy = top + r * pitch
        if cards:
            add_box(sl, cx, cy, cw, ch, fill=card_bg)
        cs = Stack(sl, cx + pad, cy + pad, cw - pad * 2,
                   card_bg if cards else bg)
        cs.add("dsCaps12", it.get("kicker"))
        cs.add("dsH4", it["title"], gap=TIGHT)
        cs.add(body_style, it.get("body"), gap=TIGHT)
    return sl


def zoned_content(prs, spec):
    """Zona superior (argumento) + faixa inferior (desdobramento).

    Substitui o pillar_card_dense do 1.0, que empilhava tudo no topo e deixava
    metade do slide vazia.
    """
    bg = spec.get("bg", T.BEGE)
    sl = blank_slide(prs, bg)
    items = spec.get("items", [])
    n = max(1, len(items))
    span = max(1, T.COLUMNS // n)
    iw = T.col_span(span)

    # faixa inferior, ancorada na base
    band_h = 0
    for it in items:
        h = needed("dsH4", it["title"], iw)
        if it.get("body"):
            h += TIGHT + needed("dsCorpoPilar", it["body"], iw)
        band_h = max(band_h, h)
    band_y = T.CONTENT_BOTTOM_PT - 40 - band_h
    for i, it in enumerate(items):
        ix = T.col_x(i * span)
        cs = Stack(sl, ix, band_y, iw, bg)
        cs.add("dsH4", it["title"])
        cs.add("dsCorpoPilar", it.get("body"), gap=TIGHT)

    # zona superior, com ZONE de ar antes da faixa
    s = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT, T.col_span(3), bg)
    s.add("dsEyebrow", spec.get("eyebrow"))
    s.add("dsH1", spec["title"], gap=ITEM)
    s.fit("dsH5", spec.get("body", ""), band_y - ZONE, gap=BLOCK)
    return sl


def stat_band(prs, spec):
    """Zona superior + faixa inferior de NUMEROS. Variacao do zoned_content."""
    bg = spec.get("bg", T.BEGE)
    sl = blank_slide(prs, bg)
    stats = spec["stats"]
    n = len(stats)
    span = max(1, T.COLUMNS // n)
    iw = T.col_span(span)

    band_h = max(needed("dsMega", st["value"], iw)
                 + TIGHT + needed("dsCorpoPilar", st["label"], iw)
                 for st in stats)
    band_y = T.CONTENT_BOTTOM_PT - 40 - band_h
    for i, st in enumerate(stats):
        ix = T.col_x(i * span)
        cs = Stack(sl, ix, band_y, iw, bg)
        cs.add("dsMega", st["value"])
        cs.add("dsCorpoPilar", st["label"], gap=TIGHT)

    s = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT, T.col_span(3), bg)
    s.add("dsCaps12", spec.get("eyebrow"))
    s.add("dsH1", spec["title"], gap=ITEM)
    s.fit("dsH5", spec.get("body", ""), band_y - ZONE, gap=BLOCK)
    return sl


def quote_side_image(prs, spec):
    """Texto a' esquerda, painel de cor a' direita ocupando a altura util."""
    bg = spec.get("bg", T.BEGE)
    sl = blank_slide(prs, bg)
    tw = T.col_span(2)
    px = T.col_x(2)
    pw = T.col_span(2)
    ph = T.CONTENT_H_PT
    panel_bg = spec.get("panel_bg", T.AZUL)
    add_box(sl, px, T.MARGIN_TOP_PT, pw, ph, fill=panel_bg)
    if spec.get("panel_text"):
        pstyle = spec.get("panel_style", "dsH3")
        pad = T.BOX_PAD_PT
        th = needed(pstyle, spec["panel_text"], pw - pad * 2)
        add_text(sl, pstyle, spec["panel_text"], px + pad,
                 T.MARGIN_TOP_PT + max(pad, (ph - th) / 2), pw - pad * 2, th,
                 bg=panel_bg)

    s = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT, tw, bg)
    s.add("dsEyebrow", spec.get("eyebrow"))
    s.add("dsH1", spec["title"], gap=ITEM)
    s.fit("dsH3", spec["body"], T.CONTENT_BOTTOM_PT - 40, gap=BLOCK)
    return sl


def do_dont(prs, spec):
    """Duas colunas. Cabecalho em Subtitulo 34 e itens numa CAIXA SO'."""
    bg = spec.get("bg", T.BEGE)
    sl = blank_slide(prs, bg)
    s = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT, T.CONTENT_W_PT, bg)
    s.add("dsH1", spec["title"])
    s.add("dsCorpoPilar", spec.get("intro"), gap=ITEM, w=T.col_span(3))

    cw = T.col_span(2)
    pad = T.BOX_PAD_PT
    inner = cw - pad * 2
    cols = [(spec.get("do_title", "Sim"), spec["do"], T.AZUL),
            (spec.get("dont_title", "Nunca"), spec["dont"], T.ROSA)]
    ch = pad * 2 + max(needed("dsH3", h, inner) + BLOCK
                       + list_height("dsCorpoPilar", it, inner)
                       for h, it, _ in cols)
    cy = s.y + BLOCK
    for i, (head, items, tone) in enumerate(cols):
        cx = T.col_x(i * 2)
        add_box(sl, cx, cy, cw, ch, fill=T.BRANCO)
        cs = Stack(sl, cx + pad, cy + pad, inner, T.BRANCO)
        cs.add("dsH3", head, color=tone)
        cs.add_list("dsCorpoPilar", items, gap=BLOCK)
    return sl


def diagram_page(prs, spec):
    """Esquema em escala: os retangulos sao proporcionais de verdade."""
    bg = spec.get("bg", T.BEGE)
    sl = blank_slide(prs, bg)
    lw = T.col_span(2)     # 1 coluna quebrava titulo de 60pt no meio da palavra
    s = Stack(sl, T.MARGIN_LEFT_PT, T.MARGIN_TOP_PT, lw, bg)
    s.add("dsH1", spec["title"])
    s.fit("dsCorpoPilar", spec.get("intro", ""), T.CONTENT_BOTTOM_PT - 200,
          gap=BLOCK)

    legend = spec.get("legend", [])
    leg_h = (len(legend) * (12 + TIGHT + 24) + ITEM * max(0, len(legend) - 1)
             if legend else 0)
    dy = T.MARGIN_TOP_PT + BLOCK
    max_dh = T.CONTENT_BOTTOM_PT - 40 - dy - (leg_h + BLOCK if legend else 0)
    dw = min(T.col_span(2), max_dh * T.PAGE_W_PT / T.PAGE_H_PT)
    dh = dw * T.PAGE_H_PT / T.PAGE_W_PT
    dx = T.MARGIN_RIGHT_PT - dw
    scale = dw / T.PAGE_W_PT
    add_box(sl, dx, dy, dw, dh, fill=T.BRANCO, radius="hard", line=T.AZUL)
    if spec.get("show_margin", True):
        m = T.MARGIN_PT * scale
        add_box(sl, dx + m, dy + m, dw - 2 * m, dh - 2 * m,
                radius="hard", line=T.ROSA)
    if spec.get("show_module"):
        u = T.SPACING_UNIT_PT * scale
        for i in range(8):
            add_box(sl, dx + T.MARGIN_PT * scale + i * u * 2, dy + dh - u * 3,
                    u, u, fill=T.AZUL, radius="hard")

    if legend:
        ls = Stack(sl, dx, dy + dh + BLOCK, dw, bg)
        for i, (label, value) in enumerate(legend):
            ls.add("dsCaps12", label, gap=ITEM if i else 0)
            ls.add("dsCorpoPilar", value, gap=TIGHT)
    return sl


def closing(prs, spec):
    """Fecho. Sem linha de versao: o 'ultimo slide' de verdade vira arquetipo
    proprio mais adiante."""
    bg = spec.get("bg", T.AZUL)
    sl = blank_slide(prs, bg)
    w = T.col_span(3)
    th = needed("dsHero", spec["title"], w)
    add_text(sl, "dsHero", spec["title"], T.MARGIN_LEFT_PT,
             (T.PAGE_H_PT - th) / 2, w, th, bg=bg)
    return sl


BUILDERS = {
    "hero_cover": hero_cover,
    "chapter_divider": chapter_divider,
    "spec_page": spec_page,
    "type_specimen": type_specimen,
    "swatch_page": swatch_page,
    "multi_card_grid": multi_card_grid,
    "card_grid_5": card_grid_5,
    "grid_plain": grid_plain,
    "zoned_content": zoned_content,
    "stat_band": stat_band,
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
