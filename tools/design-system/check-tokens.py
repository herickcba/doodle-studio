#!/usr/bin/env python3
"""check-tokens.py -- o .bas e o tokens.py contam a mesma historia?

O design system tem duas encarnacoes: o `SetDefaults` do BG-DoodleStudio.bas
(o que o usuario aplica com um clique) e o tokens.py (o que a documentacao e o
gerador do deck dizem). Se elas divergirem, a documentacao mente -- e mentira
em documento de marca custa caro, porque ninguem descobre na hora.

Este script le os dois e compara estilo a estilo. Divergencia = exit 1.
Roda no checklist de publicacao (tools/BUILD.md 4.1).

Uso: python3 tools/design-system/check-tokens.py [caminho/do/.bas]
"""

from __future__ import annotations

import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))

# Os tokens moram na skill (fonte unica). Ver comentario no build_deck.py.
sys.path.insert(0, os.path.join(REPO, ".claude", "skills", "cba-visual-v2",
                                "scripts"))

import tokens as T  # noqa: E402

BAS = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    REPO, "v3-powerpoint-addin", "assets", "BG-DoodleStudio.bas")


def parse_bas(path):
    """Extrai paleta, estilos e entrelinhas do SetDefaults."""
    src = open(path, encoding="utf-8", errors="replace").read()
    m = re.search(r"Private Sub SetDefaults\(\)(.*?)^End Sub", src,
                  re.S | re.M)
    if not m:
        raise SystemExit("ERRO: nao achei SetDefaults no .bas")
    body = m.group(1)

    pal = {}
    for i, r, g, b in re.findall(
            r"gPal\((\d)\)\s*=\s*RGB\((\d+),\s*(\d+),\s*(\d+)\)", body):
        pal[int(i)] = (int(r), int(g), int(b))

    styles = []
    for sid, size, bold, role, ent in re.findall(
            r'AddStyle\s+"([A-Za-z0-9_]+)",\s*([\d.]+),\s*(True|False),\s*(\d+),\s*([\d.#]+)',
            body):
        styles.append({
            "id": sid,
            "size": float(size),
            "bold": bold == "True",
            "role": int(role),
            # VBA escreve 1# para "1 como Single"; Python nao entende o #
            "ent": float(ent.rstrip("#")),
        })

    ents = {}
    for sz, mult in re.findall(r"AddEnt\s+(\d+),\s*([\d.#]+)", body):
        ents[int(sz)] = float(mult.rstrip("#"))

    fonte = re.search(r'gFonte\s*=\s*"([^"]+)"', body)
    radius = re.search(r"gRadiusPt\s*=\s*([\d.]+)", body)
    return {
        "palette": pal,
        "styles": styles,
        "ents": ents,
        "font": fonte.group(1) if fonte else None,
        "radius_pt": float(radius.group(1)) if radius else None,
    }


def parse_consts(path):
    """Constantes de layout declaradas no topo do modulo."""
    src = open(path, encoding="utf-8", errors="replace").read()
    out = {}
    for name in ("ANCHOR_DEFAULT_CM", "GUIDE_MARGIN_PT", "PT_PER_CM", "MAX_DEPTH"):
        m = re.search(r"Const\s+%s\s+As\s+\w+\s*=\s*([\d.]+)" % name, src)
        if m:
            out[name] = float(m.group(1))
    return out


def main():
    if not os.path.exists(BAS):
        raise SystemExit("ERRO: nao achei o .bas: %s" % BAS)

    bas = parse_bas(BAS)
    consts = parse_consts(BAS)
    errs, warns = [], []

    print("== tokens.py x SetDefaults do .bas ==")
    print("   .bas: %s" % os.path.relpath(BAS, REPO))
    print()

    # ---- fonte
    if bas["font"] != T.FONT:
        errs.append("fonte: .bas=%r tokens=%r" % (bas["font"], T.FONT))

    # ---- paleta
    for i, rgb in sorted(bas["palette"].items()):
        want = T.PALETTE.get(i)
        if not want:
            errs.append("paleta: pal%d existe no .bas e nao no tokens.py" % i)
            continue
        if tuple(want["rgb"]) != rgb:
            errs.append("paleta pal%d (%s): .bas=%s tokens=%s"
                        % (i, want["name"], rgb, tuple(want["rgb"])))
    if len(T.PALETTE) != len(bas["palette"]):
        errs.append("paleta: %d cores no .bas, %d no tokens.py"
                    % (len(bas["palette"]), len(T.PALETTE)))

    # ---- estilos
    pending = {s["id"] for s in T.STYLES if s.get("pending_bas")}
    bas_ids = [s["id"] for s in bas["styles"]]
    tok_ids = [s["id"] for s in T.STYLES if s["id"] not in pending]
    for pid in sorted(pending):
        st = T.STYLE_BY_ID[pid]
        warns.append("estilo %r (%s) existe na documentacao e ainda nao na faixa.\n"
                     "       Proxima leva do .bas: AddStyle \"%s\", %g, %s, %d, %s"
                     % (st["label"], pid, pid, st["size"],
                        "True" if st["bold"] else "False", st["role"], st["ent"]))
    if bas_ids != tok_ids:
        so_bas = [i for i in bas_ids if i not in tok_ids]
        so_tok = [i for i in tok_ids if i not in bas_ids]
        if so_bas:
            errs.append("estilos so' no .bas: %s" % ", ".join(so_bas))
        if so_tok:
            errs.append("estilos so' no tokens.py: %s" % ", ".join(so_tok))
        if not so_bas and not so_tok:
            errs.append("estilos na ORDEM diferente (a ordem e' a da faixa):\n"
                        "       .bas   = %s\n       tokens = %s"
                        % (", ".join(bas_ids), ", ".join(tok_ids)))

    for b in bas["styles"]:
        t = T.STYLE_BY_ID.get(b["id"])
        if not t:
            continue
        for field in ("size", "bold", "role", "ent"):
            if b[field] != t[field]:
                errs.append("%s.%s: .bas=%r tokens=%r"
                            % (b["id"], field, b[field], t[field]))

    # ---- rotulos: o que o tokens.py chama de cada estilo tem de ser o que o
    #      usuario le' no botao. Sem isto, a documentacao usa um nome e a faixa
    #      outro -- foi assim que "Titulo 60" apareceu sem acento no deck.
    ribbon = os.path.join(REPO, "v3-powerpoint-addin", "ribbon", "customUI14.xml")
    if os.path.exists(ribbon):
        xml = open(ribbon, encoding="utf-8").read()
        labels = dict(re.findall(r'<button id="(ds[A-Za-z0-9]+)"\s+label="([^"]+)"', xml))
        for t in T.STYLES:
            want = labels.get(t["id"])
            if want and want != t["label"]:
                errs.append("%s.label: ribbon=%r tokens=%r"
                            % (t["id"], want, t["label"]))

    # ---- entrelinha por tamanho
    if bas["ents"] != T.LINE_HEIGHT_BY_SIZE:
        for sz in sorted(set(bas["ents"]) | set(T.LINE_HEIGHT_BY_SIZE)):
            a, c = bas["ents"].get(sz), T.LINE_HEIGHT_BY_SIZE.get(sz)
            if a != c:
                errs.append("entrelinha do corpo %s: .bas=%r tokens=%r" % (sz, a, c))

    # ---- raio e constantes de layout
    # Desde a v1.6.0 o raio e' um valor visual constante em PONTOS, e tem de
    # ser o mesmo dos dois lados. Nao ha' mais tolerancia aqui.
    if bas["radius_pt"] != T.RADIUS_PT_BAS:
        errs.append("raio: .bas=%rpt tokens=%rpt"
                    % (bas["radius_pt"], T.RADIUS_PT_BAS))
    if consts.get("ANCHOR_DEFAULT_CM") != T.ANCHOR_DEFAULT_CM:
        errs.append("ancora padrao: .bas=%rcm tokens=%rcm"
                    % (consts.get("ANCHOR_DEFAULT_CM"), T.ANCHOR_DEFAULT_CM))
    if consts.get("GUIDE_MARGIN_PT") != T.GUIDE_MARGIN_PT:
        errs.append("margem das guias: .bas=%rpt tokens=%rpt"
                    % (consts.get("GUIDE_MARGIN_PT"), T.GUIDE_MARGIN_PT))
    if consts.get("MAX_DEPTH") != T.MAX_DEPTH:
        errs.append("MAX_DEPTH: .bas=%r tokens=%r"
                    % (consts.get("MAX_DEPTH"), T.MAX_DEPTH))

    if warns:
        print("AVISOS (nao bloqueiam):")
        for w in warns:
            print("  - " + w)
        print()

    if errs:
        print("DIVERGENCIA entre o codigo e a documentacao (%d):" % len(errs))
        for e in errs:
            print("  - " + e)
        print()
        print("  O .bas e' a fonte da verdade. Atualize tools/design-system/tokens.py")
        print("  e o design.md, e regenere o deck (build_deck.py).")
        return 1

    print("OK: %d estilos, %d cores e as constantes de layout batem." %
          (len(T.STYLES), len(T.PALETTE)))
    print("    Font %s | design 2.0: margem %.0fpt, modulo %dpt, raio %.0fpt, "
          "gutter %.0fpt, %d colunas de %.2fpt"
          % (T.FONT, T.MARGIN_PT, T.SPACING_UNIT_PT, T.RADIUS_PT, T.GUTTER_PT,
             T.COLUMNS, T.COL_W_PT))
    return 0


if __name__ == "__main__":
    sys.exit(main())
