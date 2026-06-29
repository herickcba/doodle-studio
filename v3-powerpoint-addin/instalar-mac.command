#!/bin/bash
# Instalador do CBA Studio para PowerPoint (Mac).
# Baixa o manifesto e o coloca na pasta de suplementos do PowerPoint.

WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
URL="https://doodle-studio-app.vercel.app/manifest.xml"

mkdir -p "$WEF"
if curl -fsSL "$URL" -o "$WEF/doodle-studio.xml"; then
  osascript -e 'tell application "Microsoft PowerPoint" to quit' >/dev/null 2>&1 || true
  osascript -e 'display dialog "CBA Studio instalado! ✅

Abra o PowerPoint e vá em:
Inserir → Add-ins → Developer Add-ins → CBA Studio." buttons {"OK"} default button "OK" with title "CBA Studio"' >/dev/null 2>&1
  echo "Instalado em: $WEF/doodle-studio.xml"
else
  osascript -e 'display dialog "Não foi possível baixar o suplemento. Verifique sua conexão e tente de novo." buttons {"OK"} with title "CBA Studio" with icon stop' >/dev/null 2>&1
  echo "Falha ao baixar de $URL"
  exit 1
fi
