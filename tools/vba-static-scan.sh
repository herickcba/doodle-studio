#!/bin/bash
# ============================================================
#  vba-static-scan.sh
#  Analise estatica do fonte VBA da Faixa, conforme o item 4 da
#  analise de risco ("Validacao tecnica obrigatoria").
#
#  Roda a lista de termos sensiveis contra o .bas. Toda ocorrencia
#  precisa estar REMOVIDA do codigo ou LISTADA em
#  tools/vba-scan-allowlist.txt com justificativa tecnica.
#  Qualquer ocorrencia nao justificada -> exit 1 (bloqueia a publicacao).
#
#  Uso:  tools/vba-static-scan.sh [arquivo.bas]
#  Default: v3-powerpoint-addin/assets/BG-DoodleStudio.bas
# ============================================================
set -euo pipefail

HERE="$(cd "$(dirname "$0")" && pwd)"
REPO="$(cd "$HERE/.." && pwd)"

BAS="${1:-$REPO/v3-powerpoint-addin/assets/BG-DoodleStudio.bas}"
ALLOW="$HERE/vba-scan-allowlist.txt"

command -v python3 >/dev/null || { echo "Erro: python3 nao encontrado."; exit 1; }
[ -f "$BAS" ]   || { echo "Erro: nao achei o .bas: $BAS"; exit 1; }
[ -f "$ALLOW" ] || { echo "Erro: allowlist ausente: $ALLOW"; exit 1; }

python3 - "$BAS" "$ALLOW" <<'PY'
import re, sys, os

bas_path, allow_path = sys.argv[1], sys.argv[2]
src = open(bas_path, encoding='utf-8', errors='replace').read().splitlines()

# ---------------------------------------------------------------
# Termos. TIER 1 = lista literal do item 4 do documento de analise.
# TIER 2 = superficies proibidas no item 3 que a lista do item 4
#          nao nomeia (sistema de arquivos, automacao de UI, rede).
# Regex com \b so' no INICIO de proposito: "Execute" precisa casar
# com "ExecuteMso", "Shell" com "ShellExecute".
# ---------------------------------------------------------------
TIER1 = [
    ('Shell',            r'\bShell'),
    ('WScript',          r'\bWScript'),
    ('PowerShell',       r'PowerShell'),
    ('cmd.exe',          r'cmd\.exe'),
    ('osascript',        r'osascript'),
    ('AppleScript',      r'AppleScript'),
    ('CreateObject',     r'\bCreateObject'),
    ('GetObject',        r'\bGetObject'),
    ('CallByName',       r'\bCallByName'),
    ('Application.Run',  r'\bApplication\s*\.\s*Run\b'),
    ('Declare',          r'\bDeclare\b'),
    ('PtrSafe',          r'\bPtrSafe\b'),
    ('LoadLibrary',      r'\bLoadLibrary'),
    ('RegWrite',         r'\bRegWrite'),
    ('RegRead',          r'\bRegRead'),
    ('RegDelete',        r'\bRegDelete'),
    ('schtasks',         r'schtasks'),
    ('XMLHTTP',          r'\bXMLHTTP'),
    ('ServerXMLHTTP',    r'ServerXMLHTTP'),
    ('WinHTTP',          r'WinHTTP'),
    ('ADODB',            r'\bADODB'),
    ('URLDownloadToFile',r'URLDownloadToFile'),
    ('VBProject',        r'\bVBProject'),
    ('VBComponents',     r'\bVBComponents'),
    ('CodeModule',       r'\bCodeModule'),
    ('Eval',             r'\bEval'),
    ('Execute',          r'\bExecute'),
]
TIER2 = [
    ('FileSystemObject', r'FileSystemObject'),
    ('ScriptControl',    r'ScriptControl'),
    ('MSXML2',           r'MSXML2'),
    ('InternetExplorer', r'InternetExplorer'),
    ('WinInet',          r'WinInet'),
    ('SendKeys',         r'\bSendKeys'),
    ('Kill',             r'\bKill\b'),
    ('RmDir',            r'\bRmDir'),
    ('MkDir',            r'\bMkDir'),
    ('FileCopy',         r'\bFileCopy'),
    ('SetAttr',          r'\bSetAttr'),
    ('ChDir',            r'\bChDir'),
    ('ChDrive',          r'\bChDrive'),
    ('SaveAs',           r'\bSave(Copy)?As'),
    ('Open',             r'\bOpen\b[^\n]*\bFor\s+(Input|Output|Append|Binary|Random)\b'),
    ('Dir',              r'\bDir\$?\s*\('),
    ('Environ',          r'\bEnviron'),
    ('FollowHyperlink',  r'\bFollowHyperlink'),
    ('URL',              r'https?://'),
]
# Tokens que, dentro de uma string literal, sugerem um comando de
# sistema sendo montado por concatenacao (item 4, ultimo paragrafo).
CONCAT_TOKENS = re.compile(
    r'\.exe|\.ps1|\.bat|\.cmd|\.vbs|\.sh\b|/bin/|powershell|osascript|cmd\s*/c|-command',
    re.I)

# ---------------------------------------------------------------
# Remove comentarios respeitando strings ("" e' aspas escapada no VBA).
# ---------------------------------------------------------------
def strip_comment(line):
    out, in_str, i = [], False, 0
    while i < len(line):
        c = line[i]
        if c == '"':
            if in_str and i + 1 < len(line) and line[i + 1] == '"':
                out.append('""'); i += 2; continue
            in_str = not in_str
        elif c == "'" and not in_str:
            break
        out.append(c); i += 1
    return ''.join(out)

def literals(line):
    return re.findall(r'"((?:[^"]|"")*)"', line)

# ---------------------------------------------------------------
# Mapa linha -> procedimento que a contem.
# ---------------------------------------------------------------
PROC_RE = re.compile(r'^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?'
                     r'(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z0-9_]+)', re.I)
END_RE  = re.compile(r'^\s*End\s+(Sub|Function|Property)\b', re.I)

proc_of, cur = {}, '(nivel de modulo)'
for n, raw in enumerate(src, 1):
    code = strip_comment(raw)
    m = PROC_RE.match(code)
    if m:
        cur = m.group(1)
    proc_of[n] = cur
    if END_RE.match(code):
        cur = '(nivel de modulo)'

# ---------------------------------------------------------------
# Allowlist: TERMO | PROCEDIMENTO | justificativa
# ---------------------------------------------------------------
allow, allow_used = {}, set()
for raw in open(allow_path, encoding='utf-8'):
    ln = raw.strip()
    if not ln or ln.startswith('#'):
        continue
    parts = [p.strip() for p in ln.split('|')]
    if len(parts) < 3:
        print('ERRO: linha malformada na allowlist: ' + ln); sys.exit(1)
    allow[(parts[0], parts[1])] = parts[2]

# ---------------------------------------------------------------
# Varredura
# ---------------------------------------------------------------
justified, blocking, comments_only, concat_hits = [], [], [], []

URL_RE = re.compile(r'https?://[^\s"\')]+', re.I)

for n, raw in enumerate(src, 1):
    code = strip_comment(raw)
    for tier, terms in (('1', TIER1), ('2', TIER2)):
        for label, pat in terms:
            if re.search(pat, code, re.I):
                # URL e' liberada pela URL LITERAL, nunca pelo procedimento:
                # senao bastaria injetar um segundo endereco dentro de um
                # procedimento que ja' tem uma URL justificada.
                if label == 'URL':
                    for u in URL_RE.findall(code):
                        key = ('URL', u)
                        if key in allow:
                            allow_used.add(key)
                            justified.append((n, 'URL', u, allow[key], code.strip()))
                        else:
                            blocking.append((n, tier, 'URL', u, code.strip()))
                    continue
                key = (label, proc_of[n])
                if key in allow:
                    allow_used.add(key)
                    justified.append((n, label, proc_of[n], allow[key], code.strip()))
                else:
                    blocking.append((n, tier, label, proc_of[n], code.strip()))
            elif re.search(pat, raw, re.I):
                comments_only.append((n, label, raw.strip()))
    if '&' in code:
        for lit in literals(code):
            if CONCAT_TOKENS.search(lit):
                concat_hits.append((n, lit, code.strip()))

# ---------------------------------------------------------------
# Relatorio
# ---------------------------------------------------------------
print('== Analise estatica do VBA ==')
print('   fonte: %s (%d linhas)' % (os.path.basename(bas_path), len(src)))
print('   termos verificados: %d (item 4 do documento) + %d (superficies do item 3)'
      % (len(TIER1), len(TIER2)))
print()

if justified:
    print('OCORRENCIAS JUSTIFICADAS (%d) -- ver docs/SECURITY-REVIEW-v1.5.0B.md:' % len(justified))
    for n, label, proc, why, code in justified:
        # o trecho de codigo vai junto de proposito: quem le o relatorio
        # confere a linha real, nao so' a justificativa cadastrada.
        print('  l.%-5d %-16s em %-18s %s' % (n, label, proc, code[:60]))
        print('  %-7s %s' % ('', '-> ' + why))
    print()

if comments_only:
    print('MENCOES SO\' EM COMENTARIO (%d, nao bloqueiam):' % len(comments_only))
    for n, label, txt in comments_only:
        print('  l.%-5d %-16s %s' % (n, label, txt[:70]))
    print()

stale = sorted(set(allow) - allow_used)
if stale:
    print('AVISO: entradas da allowlist que nao casam com nenhuma ocorrencia')
    print('       (codigo mudou? remova a entrada para manter o documento honesto):')
    for label, proc in stale:
        print('  - %s em %s' % (label, proc))
    print()

fail = False
if concat_hits:
    fail = True
    print('BLOQUEIO: string concatenada que pode formar um comando de sistema (%d):' % len(concat_hits))
    for n, lit, code in concat_hits:
        print('  l.%-5d "%s"  ->  %s' % (n, lit[:40], code[:70]))
    print()

if blocking:
    fail = True
    print('BLOQUEIO: ocorrencia sem justificativa (%d):' % len(blocking))
    for n, tier, label, proc, code in blocking:
        print('  l.%-5d [tier %s] %-16s em %-18s %s' % (n, tier, label, proc, code[:60]))
    print()
    print('  Para liberar: remova o codigo, OU adicione em tools/vba-scan-allowlist.txt')
    print('  a linha "TERMO | PROCEDIMENTO | justificativa" e documente em')
    print('  docs/SECURITY-REVIEW-v1.5.0B.md. Nao adicione sem revisar de verdade.')
    sys.exit(1)

if fail:
    sys.exit(1)

print('OK: nenhuma ocorrencia nao justificada. %d justificadas, %d mencoes em comentario.'
      % (len(justified), len(comments_only)))
PY

echo "OK -> analise estatica passou ($(basename "$BAS"))"
