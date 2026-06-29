#!/bin/bash
# ============================================================
#  Desinstalador da faixa "Doodle Studio".
#  Remove o registro no MicrosoftRegistrationDB e apaga o .ppam.
#  Duplo-clique, ou: tools/uninstall-doodle-ribbon.command
# ============================================================
set -euo pipefail

ADDIN="DoodleStudio"
GC="$HOME/Library/Group Containers/UBF8T346G9.Office"
DEST_DIR="$GC/User Content.localized/Add-Ins.localized/$ADDIN"
DB="$GC/MicrosoftRegistrationDB.reg"

echo "→ Desinstalando Doodle Studio…"
osascript -e 'tell application "Microsoft PowerPoint" to quit' >/dev/null 2>&1 || true
sleep 1

if [ -f "$DB" ]; then
  ADDINS='Software/Microsoft/Office/16.0/PowerPoint/AddIns'
  IFS='/' read -r -a arr <<< "$ADDINS"
  pid="-1"
  ok=1
  for key in "${arr[@]}"; do
    pid=$(echo "SELECT node_id FROM HKEY_CURRENT_USER WHERE name='$key' AND parent_id=$pid" | sqlite3 "$DB")
    [[ -z "$pid" ]] && { ok=0; break; }
  done
  if [[ "$ok" == "1" ]]; then
    node_id=$(echo "SELECT node_id FROM HKEY_CURRENT_USER WHERE name='$ADDIN' AND parent_id=$pid" | sqlite3 "$DB")
    if [[ -n "$node_id" ]]; then
      echo "DELETE FROM HKEY_CURRENT_USER WHERE node_id='$node_id'" | sqlite3 "$DB"
      echo "DELETE FROM HKEY_CURRENT_USER_values WHERE node_id='$node_id'" | sqlite3 "$DB"
      echo "  registro removido (node $node_id)"
    fi
  fi
fi

rm -rf "$DEST_DIR"
echo "✅ Removido. Reabra o PowerPoint — a aba 'Doodle Studio' some."
