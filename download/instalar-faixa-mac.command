#!/bin/bash
# ============================================================
#  Instalador da faixa "CBA Studio" (.ppam VBA + RibbonX).
#  Copia o add-in para a pasta do Office e registra no
#  MicrosoftRegistrationDB (AutoLoad + Path) p/ a aba aparecer
#  sozinha ao abrir o PowerPoint. (Espelha o BrightSlide.)
#
#  Duplo-clique, ou: tools/install-doodle-ribbon.command [caminho.ppam]
# ============================================================
set -euo pipefail

ADDIN="DoodleStudio"
HERE="$(cd "$(dirname "$0")" && pwd)"

# 1) localizar o .ppam
SRC="${1:-}"
if [ -z "$SRC" ]; then
  for c in "$HERE/BG-DoodleStudio.ppam" "$HOME/Downloads/BG-DoodleStudio.ppam"; do
    [ -f "$c" ] && SRC="$c" && break
  done
fi
[ -n "$SRC" ] && [ -f "$SRC" ] || { echo "❌ Nao achei BG-DoodleStudio.ppam. Passe o caminho como argumento."; exit 1; }

GC="$HOME/Library/Group Containers/UBF8T346G9.Office"
DEST_DIR="$GC/User Content.localized/Add-Ins.localized/$ADDIN"
DB="$GC/MicrosoftRegistrationDB.reg"

echo "→ Instalando CBA Studio…"

# 2) fechar o PowerPoint p/ recarregar o add-in
osascript -e 'tell application "Microsoft PowerPoint" to quit' >/dev/null 2>&1 || true
sleep 1

# 3) copiar o .ppam
mkdir -p "$DEST_DIR"
cp "$SRC" "$DEST_DIR/BG-DoodleStudio.ppam"
PPAM="$DEST_DIR/BG-DoodleStudio.ppam"

[ -f "$DB" ] || { echo "❌ RegistrationDB nao existe ($DB). Abra o PowerPoint uma vez e rode de novo."; exit 1; }

# timestamp no formato Windows FILETIME (hex big-endian), como o BrightSlide
timestamp() {
  local T=10000000 E=11644473600
  local ft=$((($(date +"%s") + E) * T))
  local hex='0'$(printf '%x\n' "$ft") out=''
  local i
  for ((i=0; i<16; i=i+2)); do out=${hex:$i:2}$out; done
  echo "$out"
}

# 4) registrar: caminhar/criar Software/Microsoft/Office/16.0/PowerPoint/AddIns
ADDINS='Software/Microsoft/Office/16.0/PowerPoint/AddIns'
IFS='/' read -r -a arr <<< "$ADDINS"
pid="-1"
for key in "${arr[@]}"; do
  cid=$(echo "SELECT node_id FROM HKEY_CURRENT_USER WHERE name='$key' AND parent_id=$pid" | sqlite3 "$DB")
  if [[ -z "$cid" ]]; then
    echo "INSERT INTO HKEY_CURRENT_USER (parent_id, name, write_time) VALUES ($pid, '$key', X'$(timestamp)')" | sqlite3 "$DB"
    cid=$(echo "SELECT MAX(node_id) FROM HKEY_CURRENT_USER" | sqlite3 "$DB")
  fi
  pid=$cid
done
addins_id=$pid

# nó do add-in
addin_id=$(echo "SELECT node_id FROM HKEY_CURRENT_USER WHERE name='$ADDIN' AND parent_id=$addins_id" | sqlite3 "$DB")
if [[ -z "$addin_id" ]]; then
  addin_id=$(echo "INSERT INTO HKEY_CURRENT_USER (parent_id, name, write_time) VALUES ($addins_id, '$ADDIN', X'$(timestamp)'); SELECT last_insert_rowid()" | sqlite3 "$DB")
fi

# valor AutoLoad = 1 (type 4 = DWORD)
nid=$(echo "SELECT node_id FROM HKEY_CURRENT_USER_values WHERE node_id=$addin_id AND name='AutoLoad'" | sqlite3 "$DB")
[[ -z "$nid" ]] && echo "INSERT INTO HKEY_CURRENT_USER_values (node_id, name, type, value) VALUES ($addin_id, 'AutoLoad', 4, '1')" | sqlite3 "$DB"

# valor Path (type 1 = string) — atualiza sempre
pnid=$(echo "SELECT node_id FROM HKEY_CURRENT_USER_values WHERE node_id=$addin_id AND name='Path'" | sqlite3 "$DB")
if [[ -z "$pnid" ]]; then
  echo "INSERT INTO HKEY_CURRENT_USER_values (node_id, name, type, value) VALUES ($addin_id, 'Path', 1, '$PPAM')" | sqlite3 "$DB"
else
  echo "UPDATE HKEY_CURRENT_USER_values SET value='$PPAM' WHERE node_id=$addin_id AND name='Path'" | sqlite3 "$DB"
fi

echo "✅ Instalado: $PPAM"
echo "   Abra o PowerPoint → clique 'Enable Macros' na 1ª vez → a aba 'CBA Studio' aparece."
