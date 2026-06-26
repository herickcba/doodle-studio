#!/usr/bin/env python3
"""
Build assets/estilos-bg.pptx — the "style anchor" deck for the Typography module.

One slide holds the B+G typographic styles as REAL, exactly-formatted text boxes
(font, weight, size, color, the colored final period AND the line spacing baked in).
The add-in inserts this slide via insertSlidesFromBase64(KeepSourceFormatting); the
user then copy/pastes a box or uses the Format Painter onto their object — both
native paths carry the line spacing that the Office.js API cannot set.

Run:  python3 tools/build-styles-reference.py
Out:  assets/estilos-bg.pptx
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN

# --- palette ---
RED = RGBColor(0xFC, 0x5E, 0x6D)      # cor 1
BLUE = RGBColor(0x43, 0x6A, 0xE1)     # cor 2
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY = RGBColor(0x76, 0x76, 0x7C)
FONT = "Avenir Next"

# Each style: (key, label, sample, size_pt, line_spacing, text_color, period_color, fill_color)
STYLES = [
    ("hero",      "HERO STATEMENT · 120pt · 0.8x",        "Ideia grande", 120, 0.8, RED,   BLUE, None),
    ("hero_blue", "HERO sobre fundo 436AE1 · branco",     "Ideia grande", 120, 0.8, WHITE, RED,  BLUE),
    ("mega",      "MEGA STATEMENT · 80pt · 0.9x",         "Mega statement", 80, 0.9, BLUE,  BLUE, None),
    ("h1",        "H1 TÍTULO DE PÁGINA · 60pt · 0.9x",    "Título de página", 60, 0.9, RED,  RED,  None),
]


def add_caption(slide, top_in, text):
    box = slide.shapes.add_textbox(Inches(0.6), Inches(top_in), Inches(12), Inches(0.35))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run(); run.text = text
    f = run.font
    f.name = FONT; f.size = Pt(14); f.bold = True; f.color.rgb = GRAY
    return box


def add_style_box(slide, top_in, height_in, sample, size_pt, spacing, text_color, period_color, fill_color):
    box = slide.shapes.add_textbox(Inches(0.6), Inches(top_in), Inches(12.1), Inches(height_in))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = Inches(0.12); tf.margin_right = Inches(0.12)
    tf.margin_top = Inches(0.06); tf.margin_bottom = Inches(0.06)

    if fill_color is not None:
        box.fill.solid(); box.fill.fore_color.rgb = fill_color
        box.line.fill.background()
    else:
        box.fill.background(); box.line.fill.background()

    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.line_spacing = spacing            # float => multiple (0.8 == 80%)

    # main text run + final-period run in the accent color
    run = p.add_run(); run.text = sample
    f = run.font; f.name = FONT; f.size = Pt(size_pt); f.bold = True; f.color.rgb = text_color

    dot = p.add_run(); dot.text = "."
    df = dot.font; df.name = FONT; df.size = Pt(size_pt); df.bold = True; df.color.rgb = period_color
    return box


def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(14.0)     # tall canvas so 120pt boxes don't collide

    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    # heading
    head = slide.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(12), Inches(0.6))
    hp = head.text_frame.paragraphs[0]
    hr = hp.add_run(); hr.text = "Estilos B+G — copie um bloco ou use o Pincel de Formatação. Apague este slide depois."
    hf = hr.font; hf.name = FONT; hf.size = Pt(16); hf.bold = True; hf.color.rgb = GRAY

    # layout: caption + box per style, generous vertical spacing
    tops = [1.1, 4.5, 7.9, 10.6]   # caption tops (inches)
    box_h = [2.9, 2.9, 2.2, 1.7]
    for (key, label, sample, size, spacing, tcol, pcol, fill), top, bh in zip(STYLES, tops, box_h):
        add_caption(slide, top, label)
        add_style_box(slide, top + 0.45, bh, sample, size, spacing, tcol, pcol, fill)

    out_dir = os.path.join(os.path.dirname(__file__), "..", "assets")
    out = os.path.normpath(os.path.join(out_dir, "estilos-bg.pptx"))
    prs.save(out)
    print("wrote", out)


if __name__ == "__main__":
    main()
