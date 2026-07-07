#!/bin/bash
# ============================================================
#  CBA Studio — instalador Mac (faixa .ppam + extensão), sem
#  arquivo baixado (não passa pelo Gatekeeper). Uso:
#    curl -fsSL https://doodle-studio-app.vercel.app/install-mac.sh | bash
#  Instala a extensão (manifesto -> pasta wef) e a faixa (.ppam +
#  registro AutoLoad no MicrosoftRegistrationDB, estilo BrightSlide).
# ============================================================
set -uo pipefail

PPAM_URL="https://doodle-studio-sigma.vercel.app/download/BG-DoodleStudio.ppam"
MANIFEST_URL="https://doodle-studio-app.vercel.app/manifest.xml"
ADDIN="DoodleStudio"

echo "→ Instalando CBA Studio…"

# fecha o PowerPoint p/ recarregar tudo
osascript -e 'tell application "Microsoft PowerPoint" to quit' >/dev/null 2>&1 || true
sleep 1

# ---------- 1) EXTENSÃO (manifesto -> pasta wef) ----------
WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
mkdir -p "$WEF"
if curl -fsSL "$MANIFEST_URL" -o "$WEF/cba-studio.xml"; then
  echo "  ✅ Extensão instalada (painel lateral)."
else
  echo "  ⚠️  Não consegui baixar a extensão — verifique a internet."
fi

# ---------- 2) FAIXA (.ppam + registro AutoLoad) ----------
GC="$HOME/Library/Group Containers/UBF8T346G9.Office"
DEST_DIR="$GC/User Content.localized/Add-Ins.localized/$ADDIN"
DB="$GC/MicrosoftRegistrationDB.reg"
mkdir -p "$DEST_DIR"
PPAM="$DEST_DIR/BG-DoodleStudio.ppam"

if curl -fsSL "$PPAM_URL" -o "$PPAM"; then
  if [ -f "$DB" ]; then
    # timestamp no formato Windows FILETIME (hex big-endian)
    timestamp() {
      local T=10000000 E=11644473600
      local ft=$((($(date +"%s") + E) * T))
      local hex='0'$(printf '%x\n' "$ft") out='' i
      for ((i=0; i<16; i=i+2)); do out=${hex:$i:2}$out; done
      echo "$out"
    }
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
    addin_id=$(echo "SELECT node_id FROM HKEY_CURRENT_USER WHERE name='$ADDIN' AND parent_id=$addins_id" | sqlite3 "$DB")
    if [[ -z "$addin_id" ]]; then
      addin_id=$(echo "INSERT INTO HKEY_CURRENT_USER (parent_id, name, write_time) VALUES ($addins_id, '$ADDIN', X'$(timestamp)'); SELECT last_insert_rowid()" | sqlite3 "$DB")
    fi
    nid=$(echo "SELECT node_id FROM HKEY_CURRENT_USER_values WHERE node_id=$addin_id AND name='AutoLoad'" | sqlite3 "$DB")
    [[ -z "$nid" ]] && echo "INSERT INTO HKEY_CURRENT_USER_values (node_id, name, type, value) VALUES ($addin_id, 'AutoLoad', 4, '1')" | sqlite3 "$DB"
    pnid=$(echo "SELECT node_id FROM HKEY_CURRENT_USER_values WHERE node_id=$addin_id AND name='Path'" | sqlite3 "$DB")
    if [[ -z "$pnid" ]]; then
      echo "INSERT INTO HKEY_CURRENT_USER_values (node_id, name, type, value) VALUES ($addin_id, 'Path', 1, '$PPAM')" | sqlite3 "$DB"
    else
      echo "UPDATE HKEY_CURRENT_USER_values SET value='$PPAM' WHERE node_id=$addin_id AND name='Path'" | sqlite3 "$DB"
    fi
    echo "  ✅ Faixa instalada (aba CBA Studio)."
  else
    echo "  ⚠️  Abra o PowerPoint uma vez e rode o comando de novo (o registro do Office ainda não existe)."
  fi
else
  echo "  ⚠️  Não consegui baixar a faixa — verifique a internet."
fi

echo ""
echo "✅ Pronto! Abra o PowerPoint:"
echo "   • Na 1ª vez, clique em \"Ativar Macros\" → a aba \"CBA Studio\" aparece."
echo "   • Painel: Inserir → Suplementos → Meus Suplementos → CBA Studio."
