#!/usr/bin/env bash
# pack-skill.sh -- reempacota a skill cba-visual-v2 para download/.
#
# A skill e' a FONTE UNICA dos tokens (o tools/design-system/ importa dela), e
# vai publicada junto com a ferramenta em todo release: quem baixa o add-in
# leva o mesmo sistema visual para o Claude.
#
# Antes de substituir o zip antigo, PROVA que o novo funciona: extrai num
# diretorio temporario, gera o deck de exemplo e roda a auditoria. Zip que nao
# gera nao vai para a landing.
set -euo pipefail

REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SRC="$REPO/.claude/skills"
NAME="cba-visual-v2"
OUT="$REPO/download/$NAME.zip"
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

[ -d "$SRC/$NAME" ] || { echo "ERRO: nao achei $SRC/$NAME"; exit 1; }

find "$SRC/$NAME" -name "__pycache__" -type d -exec rm -rf {} + 2>/dev/null || true
( cd "$SRC" && zip -qr "$TMP/$NAME.zip" "$NAME" -x "*.DS_Store" "*__pycache__*" )

# prova de vida: o zip extraido gera e audita sozinho
( cd "$TMP" && unzip -q "$NAME.zip" \
  && python3 "$NAME/scripts/build.py" "$NAME/examples/minimal.json" prova.pptx >/dev/null )

mv "$TMP/$NAME.zip" "$OUT"
echo "OK -> $(basename "$OUT") ($(du -h "$OUT" | cut -f1)), testado a partir do zip"
