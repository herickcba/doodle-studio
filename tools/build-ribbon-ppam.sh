#!/bin/bash
# ============================================================
#  build-ribbon-ppam.sh
#  Pega um .pptm (com a VBA BG_DoodleStudio ja' compilada no
#  PowerPoint) e gera o BG-DoodleStudio.ppam: troca o content-type
#  para add-in e injeta a faixa (customUI14.xml + rels + imagens).
#
#  Uso:  tools/build-ribbon-ppam.sh [entrada.pptm] [saida.ppam]
#  Default: ~/Downloads/BG-DoodleStudio.pptm -> ~/Downloads/BG-DoodleStudio.ppam
# ============================================================
set -euo pipefail

HERE="$(cd "$(dirname "$0")" && pwd)"
REPO="$(cd "$HERE/.." && pwd)"
RIBBON="$REPO/v3-powerpoint-addin/ribbon"

PPTM="${1:-$HOME/Downloads/BG-DoodleStudio.pptm}"
OUT="${2:-$HOME/Downloads/BG-DoodleStudio.ppam}"

command -v python3 >/dev/null || { echo "Erro: python3 nao encontrado (necessario para o build)."; exit 1; }
command -v zip >/dev/null || { echo "Erro: zip nao encontrado."; exit 1; }
[ -f "$PPTM" ] || { echo "Erro: nao achei o .pptm: $PPTM"; echo "(Cole a VBA num deck, compile e salve como .pptm primeiro.)"; exit 1; }
unzip -t "$PPTM" >/dev/null 2>&1 || { echo "Erro: $PPTM nao e' um .pptm valido (zip corrompido?)."; exit 1; }
[ -f "$RIBBON/customUI14.xml" ] || { echo "Erro: $RIBBON/customUI14.xml ausente"; exit 1; }

# saida absoluta
mkdir -p "$(dirname "$OUT")"
OUT="$(cd "$(dirname "$OUT")" && pwd)/$(basename "$OUT")"

TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

unzip -q "$PPTM" -d "$TMP"

# 1) [Content_Types].xml: main -> addin ; garante Default png
python3 - "$TMP" <<'PY'
import sys, os
tmp = sys.argv[1]
ct = os.path.join(tmp, '[Content_Types].xml')
s = open(ct, encoding='utf-8').read()
s = s.replace('application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml',
              'application/vnd.ms-powerpoint.addin.macroEnabled.main+xml')
if 'Extension="png"' not in s:
    s = s.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>')
open(ct, 'w', encoding='utf-8').write(s)
PY

# 2) injeta a faixa (ribbon/ -> customUI/)
mkdir -p "$TMP/customUI/_rels" "$TMP/customUI/images"
cp "$RIBBON/customUI14.xml" "$TMP/customUI/customUI14.xml"
cp "$RIBBON/_rels/customUI14.xml.rels" "$TMP/customUI/_rels/customUI14.xml.rels"
cp "$RIBBON/images/"*.png "$TMP/customUI/images/"

# 3) relacao de extensibilidade no _rels/.rels
python3 - "$TMP" <<'PY'
import sys, os
tmp = sys.argv[1]
rels = os.path.join(tmp, '_rels', '.rels')
s = open(rels, encoding='utf-8').read()
rel = ('<Relationship Id="DoodleStudioRibbon" '
       'Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" '
       'Target="customUI/customUI14.xml"/>')
if 'ui/extensibility' not in s:
    s = s.replace('</Relationships>', rel + '</Relationships>')
open(rels, 'w', encoding='utf-8').write(s)
PY

# 4) rezip como .ppam ([Content_Types].xml primeiro)
rm -f "$OUT"
( cd "$TMP" && zip -q -X "$OUT" '[Content_Types].xml' >/dev/null \
  && zip -q -rX "$OUT" . -x '[Content_Types].xml' >/dev/null )

echo "OK -> $OUT"
echo "--- conteudo relevante ---"
unzip -l "$OUT" | grep -E 'customUI|vbaProject|Content_Types' || true
