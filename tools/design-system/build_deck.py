#!/usr/bin/env python3
"""build_deck.py -- gera o deck de documentacao do design system.

    python3 tools/design-system/build_deck.py [saida.pptx]

Depois de gerar, VALIDA o proprio arquivo: le cada run de volta e confere fonte,
corpo, cor e entrelinha contra os tokens. Mesmo principio do vba-static-scan --
o gerador nao pode emitir nada fora do sistema, e a checagem e' no artefato
final, nao na intencao do codigo.
"""

from __future__ import annotations

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
sys.path.insert(0, HERE)

from pptx import Presentation  # noqa: E402
from pptx.util import Pt  # noqa: E402

import layouts  # noqa: E402
import spec as spec_mod  # noqa: E402
import tokens as T  # noqa: E402

OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    REPO, "docs", "CBA-Studio-Design-System.pptx")


def generate(path):
    prs = Presentation()
    prs.slide_width = Pt(T.PAGE_W_PT)
    prs.slide_height = Pt(T.PAGE_H_PT)

    slides = spec_mod.build_spec()
    for i, s in enumerate(slides, 1):
        try:
            layouts.build(prs, s, i)
        except Exception as e:
            raise SystemExit("ERRO no slide %d (%s): %s" % (i, s.get("kind"), e))

    os.makedirs(os.path.dirname(path), exist_ok=True)
    prs.save(path)
    return len(slides)


def validate(path):
    """Le o PPTX de volta e confere cada run contra os tokens.

    Alem dos tokens, confere TRANSBORDO: mede o texto com a fonte real e
    compara com a caixa. Texto que nao cabe sobrepoe o bloco de baixo -- foi o
    unico defeito que a validacao de tokens deixou passar na primeira versao.
    """
    import measure  # noqa: E402

    prs = Presentation(path)
    sizes_ok = {float(s) for s in T.ALLOWED_SIZES}
    hexes_ok = {h.upper() for h in T.ALLOWED_HEX}
    ents_ok = {round(s["ent"], 3) for s in T.STYLES}
    by_size = {}
    for st in T.STYLES:
        by_size.setdefault((st["size"], bool(st["bold"]), round(st["ent"], 3)), st)

    errs = []
    runs = 0
    for n, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            # transbordo: reconstroi o estilo a partir do que esta' no arquivo
            txt = shape.text_frame.text
            if txt.strip() and shape.width and shape.height:
                p0 = shape.text_frame.paragraphs[0]
                r0 = p0.runs[0] if p0.runs else None
                if r0 is not None and r0.font.size and p0.line_spacing:
                    key = (r0.font.size.pt, bool(r0.font.bold),
                           round(float(p0.line_spacing), 3))
                    st = by_size.get(key)
                    if st:
                        need = measure.block_height(txt, st, shape.width / 12700)
                        have = shape.height / 12700
                        if need > have + 2:
                            errs.append(
                                "slide %d: texto transborda %.0fpt (precisa %.0f, "
                                "caixa %.0f) em %r"
                                % (n, need - have, need, have, txt[:32].replace("\n", " ")))
            for p in shape.text_frame.paragraphs:
                if p.line_spacing is not None:
                    ls = round(float(p.line_spacing), 3)
                    if ls not in ents_ok:
                        errs.append("slide %d: entrelinha %.3f fora da tabela" % (n, ls))
                for r in p.runs:
                    if not r.text.strip():
                        continue
                    runs += 1
                    f = r.font
                    if f.name != T.FONT:
                        errs.append("slide %d: fonte %r em %r" % (n, f.name, r.text[:24]))
                    if f.size is None or f.size.pt not in sizes_ok:
                        errs.append("slide %d: corpo %s fora da escala em %r"
                                    % (n, f.size.pt if f.size else None, r.text[:24]))
                    try:
                        h = str(f.color.rgb).upper()
                    except Exception:
                        h = None
                    if h and h not in hexes_ok:
                        errs.append("slide %d: cor #%s fora da paleta em %r"
                                    % (n, h, r.text[:24]))
    return len(prs.slides), runs, errs


def main():
    n = generate(OUT)
    rel = os.path.relpath(OUT, REPO)
    print("== deck do design system ==")
    print("   %s" % rel)
    print("   %d slides" % n)
    print()

    slides, runs, errs = validate(OUT)
    print("Validacao estrutural (fonte, corpo, cor, entrelinha):")
    print("   %d slides, %d runs de texto conferidas" % (slides, runs))
    if errs:
        print()
        print("FORA DO SISTEMA (%d):" % len(errs))
        for e in errs[:40]:
            print("  - " + e)
        if len(errs) > 40:
            print("  ... e mais %d" % (len(errs) - 40))
        return 1
    print("   OK: nada fora dos tokens.")
    if slides < 50:
        print()
        print("AVISO: %d slides -- o combinado eram 50+." % slides)
    return 0


if __name__ == "__main__":
    sys.exit(main())
