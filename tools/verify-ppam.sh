#!/bin/bash
# ============================================================
#  verify-ppam.sh
#  Gate de integridade do .ppam, conforme os itens 1 e 2 da analise
#  de risco. Roda DEPOIS do build-ribbon-ppam.sh.
#
#  Prova que o binario publicado veio do fonte auditado:
#    1) extrai o fonte VBA de dentro do .ppam e compara com o .bas do repo
#    2) confere que todo callback do ribbon existe no fonte
#    3) procura identificadores no P-code sem correspondencia no fonte
#       (residuo de compilacoes antigas / indicio de VBA stomping)
#    4) calcula o SHA-256 e gera o download/build-manifest.json
#
#  NAO substitui assinatura digital -- o projeto nao e' assinado
#  (justificativa em docs/SECURITY-REVIEW-v1.5.0B.md). Este e' o
#  controle compensatorio de integridade.
#
#  Uso: tools/verify-ppam.sh [arquivo.ppam] [--strict] [--allow-dirty]
#       --strict       fantasmas de P-code passam a bloquear
#       --allow-dirty  permite repo sujo (marca o manifesto como dirty)
# ============================================================
set -euo pipefail

HERE="$(cd "$(dirname "$0")" && pwd)"
REPO="$(cd "$HERE/.." && pwd)"

PPAM="$HOME/Downloads/BG-DoodleStudio.ppam"
STRICT=0
ALLOW_DIRTY=0
for a in "$@"; do
  case "$a" in
    --strict)      STRICT=1 ;;
    --allow-dirty) ALLOW_DIRTY=1 ;;
    -*)            echo "Erro: opcao desconhecida: $a"; exit 1 ;;
    *)             PPAM="$a" ;;
  esac
done

BAS="$REPO/v3-powerpoint-addin/assets/BG-DoodleStudio.bas"
RIBBON="$REPO/v3-powerpoint-addin/ribbon/customUI14.xml"
MANIFEST="$REPO/download/build-manifest.json"

command -v python3 >/dev/null || { echo "Erro: python3 nao encontrado."; exit 1; }
[ -f "$PPAM" ]   || { echo "Erro: nao achei o .ppam: $PPAM"; exit 1; }
[ -f "$BAS" ]    || { echo "Erro: nao achei o .bas: $BAS"; exit 1; }
[ -f "$RIBBON" ] || { echo "Erro: nao achei o customUI14.xml: $RIBBON"; exit 1; }

# --- estado do git: o manifesto declara um commit, entao ele tem de ser verdade
DIRTY=0
if [ -n "$(git -C "$REPO" status --porcelain 2>/dev/null)" ]; then
  DIRTY=1
  if [ "$ALLOW_DIRTY" -eq 0 ]; then
    echo "BLOQUEIO: a arvore de trabalho tem alteracoes nao commitadas."
    echo "  O build-manifest.json declara um commit -- com o repo sujo esse"
    echo "  vinculo seria falso. Commite antes de publicar, ou use --allow-dirty"
    echo "  para um teste local (o manifesto sai marcado como dirty)."
    exit 1
  fi
fi
COMMIT="$(git -C "$REPO" rev-parse HEAD 2>/dev/null || echo 'sem-git')"
VERSION="$(sed -n 's/.*CBA_VERSION As String = "\([^"]*\)".*/\1/p' "$BAS" | head -1)"

python3 - "$PPAM" "$BAS" "$RIBBON" "$MANIFEST" "$COMMIT" "$VERSION" "$STRICT" "$DIRTY" "$HERE" <<'PY'
import hashlib, json, os, re, subprocess, sys, zipfile

ppam, bas, ribbon, manifest, commit, version, strict, dirty, tools = sys.argv[1:10]
strict, dirty = int(strict), int(dirty)
sys.path.insert(0, tools)
import vba_extract as VX

raw = open(ppam, 'rb').read()
sha = hashlib.sha256(raw).hexdigest()

print('== Verificacao de integridade do .ppam ==')
print('   arquivo: %s (%d bytes)' % (os.path.basename(ppam), len(raw)))
print('   sha256 : %s' % sha)
print('   commit : %s%s' % (commit[:12], '  [ARVORE SUJA]' if dirty else ''))
print()

fail = []

# ---------------------------------------------------------------
# 1. Fonte embutido no binario == fonte versionado
# ---------------------------------------------------------------
with zipfile.ZipFile(ppam) as z:
    names = [n for n in z.namelist() if n.endswith('vbaProject.bin')]
    if not names:
        print('BLOQUEIO: o .ppam nao contem ppt/vbaProject.bin.')
        sys.exit(1)
    vba_bin = z.read(names[0])

mods, streams = VX.extract_modules(vba_bin)
if not mods:
    print('BLOQUEIO: nao consegui extrair nenhum modulo do vbaProject.bin.')
    sys.exit(1)

repo_src = open(bas, 'rb').read().replace(b'\r\n', b'\n')
expected_mod = 'BG_DoodleStudio'

print('1) Fonte embutido x fonte versionado')
if expected_mod not in mods:
    fail.append('modulo %s ausente no .ppam' % expected_mod)
    print('   BLOQUEIO: modulo %s nao esta no .ppam. Modulos: %s'
          % (expected_mod, ', '.join(sorted(mods))))
else:
    pcode, src_text = mods[expected_mod]
    emb = src_text.encode('latin-1').replace(b'\r\n', b'\n')
    if emb == repo_src:
        print('   OK: identico byte a byte (%d bytes).' % len(emb))
    elif emb.lower() == repo_src.lower():
        n = sum(1 for a, b in zip(emb, repo_src) if a != b)
        print('   OK: identico ignorando caixa (%d bytes diferentes).' % n)
        print('       Esperado: o VBE normaliza a caixa dos identificadores ao')
        print('       compilar (VBA e case-insensitive). Nao ha' + "'" + ' diferenca de conteudo.')
    else:
        fail.append('fonte do .ppam diverge do .bas do repo')
        print('   BLOQUEIO: o fonte dentro do .ppam NAO corresponde ao .bas do repo.')
        print('   Este binario nao foi gerado a partir deste commit.')
        import difflib
        a = repo_src.decode('latin-1').splitlines()
        b = emb.decode('latin-1').splitlines()
        d = [x for x in difflib.unified_diff(a, b, 'repo/.bas', 'ppam', lineterm='', n=0)]
        print('   linhas: repo=%d ppam=%d' % (len(a), len(b)))
        for line in d[:24]:
            print('     ' + line)
        if len(d) > 24:
            print('     ... (%d linhas de diff no total)' % len(d))

# extras: modulos nao previstos
extra = sorted(set(mods) - {expected_mod})
if extra:
    fail.append('modulos inesperados: %s' % ', '.join(extra))
    print('   BLOQUEIO: modulos alem do esperado no projeto: %s' % ', '.join(extra))
print()

# ---------------------------------------------------------------
# 2. Callbacks do ribbon existem no fonte
# ---------------------------------------------------------------
print('2) Callbacks do ribbon x procedimentos do fonte')
xml = open(ribbon, encoding='utf-8').read()
cbs = sorted(set(re.findall(r'(?:onAction|onLoad|getPressed|getLabel|getText|getItemCount|getItemLabel|getSelectedItemIndex|getEnabled|getVisible|getImage)="([^"]+)"', xml)))
src_low = repo_src.decode('latin-1').lower()
procs = set(re.findall(r'^\s*(?:public\s+|private\s+|friend\s+)?(?:static\s+)?'
                       r'(?:sub|function|property\s+(?:get|let|set))\s+([a-z0-9_]+)',
                       src_low, re.M))
missing = [c for c in cbs if c.lower() not in procs]
if missing:
    fail.append('callbacks sem procedimento: %s' % ', '.join(missing))
    print('   BLOQUEIO: o ribbon chama callbacks que nao existem no fonte:')
    for c in missing:
        print('     - ' + c)
else:
    print('   OK: os %d callbacks do customUI14.xml tem procedimento correspondente.' % len(cbs))

orphan_procs = sorted(p for p in procs
                      if p not in {c.lower() for c in cbs}
                      and not re.search(r'\b%s\b' % re.escape(p),
                                        re.sub(r'^\s*(?:public|private|friend|static|\s)*'
                                               r'(?:sub|function|property\s+\w+)\s+%s\b' % re.escape(p),
                                               '', src_low, flags=re.M)))
if orphan_procs:
    print('   AVISO: procedimentos publicos sem callback e sem chamador interno: %s'
          % ', '.join(orphan_procs))
print()

# ---------------------------------------------------------------
# 3. Fantasmas de P-code (indicio de stomping / residuo de build)
# ---------------------------------------------------------------
print('3) Identificadores no P-code sem correspondencia no fonte')
allow_path = os.path.join(tools, 'pcode-ghosts-allowlist.txt')
extra_allow = []
if os.path.exists(allow_path):
    for ln in open(allow_path, encoding='utf-8'):
        ln = ln.split('#')[0].strip()
        if ln:
            extra_allow.append(ln)

if expected_mod in mods:
    pcode, src_text = mods[expected_mod]
    blob = pcode + streams.get('_VBA_PROJECT', b'')
    ghosts = VX.pcode_ghosts(blob, src_text, extra_allow)
    if not ghosts:
        print('   OK: nenhum identificador orfao no P-code.')
    else:
        print('   %d identificador(es) presentes no P-code e ausentes do fonte:'
              % len(ghosts))
        for i in range(0, len(ghosts), 6):
            print('     ' + '  '.join(ghosts[i:i + 6]))
        print()
        print('   Como ler esta lista -- sao duas coisas diferentes:')
        print('   - Token curto e sem sentido (ex.: "sMZfMz", "JVB7BZ") = lixo binario')
        print('     do proprio P-code que por acaso caiu na faixa ASCII. Muda a cada')
        print('     build e nao significa nada. Esperar ~5 por build.')
        print('   - Token que PARECE identificador ("ScaleAllSlides", "C_AZUL",')
        print('     "temBackup") = residuo real de codigo que ja\' foi removido.')
        print('     Quase sempre porque o .pptm hospedeiro foi reaproveitado entre')
        print('     builds. Remedio: compilar num arquivo-base LIMPO (tools/BUILD.md).')
        print('     Se aparecer mesmo depois de um build limpo, PARE e investigue:')
        print('     e\' o sinal que uma auditoria externa chamaria de VBA stomping.')
        if strict:
            fail.append('%d fantasmas de P-code' % len(ghosts))
            print('   BLOQUEIO (--strict).')
print()

# ---------------------------------------------------------------
# 4. Manifesto do build
# ---------------------------------------------------------------
owner = subprocess.run(['git', 'config', 'user.email'], capture_output=True,
                       text=True).stdout.strip() or 'desconhecido'
build_date = subprocess.run(['date', '-u', '+%Y-%m-%dT%H:%M:%SZ'],
                            capture_output=True, text=True).stdout.strip()
data = {
    'productName': 'BG-DoodleStudio',
    'version': version,
    'buildDate': build_date,
    'sourceCommit': commit,
    'sourceDirty': bool(dirty),
    'sha256': sha,
    'buildOwner': owner,
    'signatureStatus': 'not-signed',
    'certificateSubject': None,
    'certificateThumbprint': None,
    'signatureNote': ('Projeto VBA nao assinado: o VBE do macOS nao oferece '
                      'assinatura de projeto e a CBA nao possui certificado '
                      'corporativo de code signing. Decisao e justificativa em '
                      'docs/SECURITY-REVIEW-v1.5.0B.md.'),
}
print('4) Manifesto do build')
if fail:
    print('   nao gerado: ha bloqueios pendentes.')
else:
    os.makedirs(os.path.dirname(manifest), exist_ok=True)
    with open(manifest, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
        f.write('\n')
    # o manifesto so' vale se o hash conferir com o arquivo final
    assert hashlib.sha256(open(ppam, 'rb').read()).hexdigest() == data['sha256']
    print('   OK -> %s' % os.path.relpath(manifest, os.path.dirname(tools)))
    print('   hash conferido contra o arquivo final.')
print()

if fail:
    print('FALHOU: ' + '; '.join(fail))
    print('Publicacao interrompida.')
    sys.exit(1)
print('OK: .ppam integro e vinculado ao commit %s.' % commit[:12])
PY
