#!/usr/bin/env python3
# ============================================================
#  vba_extract.py
#  Le um vbaProject.bin (OLE Compound File) e devolve, por modulo,
#  o P-code compilado e o CODIGO-FONTE embutido.
#
#  Usa SO' a biblioteca padrao de proposito: o gate de publicacao
#  nao pode depender de `pip install` (oletools/pcodedmp) para
#  rodar em outra maquina ou daqui a dois anos.
#
#  Referencias de formato:
#    - [MS-CFB]  Compound File Binary
#    - [MS-OVBA] 2.4.1 Compression/Decompression
# ============================================================
import math
import re
import struct
import sys


def cfb_streams(data):
    """Devolve {nome_do_stream: bytes} de um OLE Compound File."""
    if data[:8] != b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        raise ValueError('nao e um OLE Compound File (assinatura ausente)')
    ssz = 1 << struct.unpack_from('<H', data, 0x1e)[0]
    msz = 1 << struct.unpack_from('<H', data, 0x20)[0]
    n_fat = struct.unpack_from('<I', data, 0x2c)[0]
    dir_start = struct.unpack_from('<I', data, 0x30)[0]
    cutoff = struct.unpack_from('<I', data, 0x38)[0]
    mini_start = struct.unpack_from('<I', data, 0x3c)[0]
    difat_start = struct.unpack_from('<I', data, 0x44)[0]
    n_difat = struct.unpack_from('<I', data, 0x48)[0]

    def sect(n):
        off = (n + 1) * ssz
        return data[off:off + ssz]

    difat = list(struct.unpack_from('<109I', data, 0x4c))
    nxt = difat_start
    for _ in range(n_difat):
        if nxt >= 0xFFFFFFFA:
            break
        s = sect(nxt)
        difat += list(struct.unpack_from('<%dI' % (ssz // 4 - 1), s, 0))
        nxt = struct.unpack_from('<I', s, ssz - 4)[0]
    difat = [d for d in difat[:n_fat] if d < 0xFFFFFFFA]

    fat = []
    for d in difat:
        fat += list(struct.unpack_from('<%dI' % (ssz // 4), sect(d), 0))

    minifat = []
    cur, seen = mini_start, set()
    while cur < 0xFFFFFFFA and cur not in seen:
        seen.add(cur)
        minifat += list(struct.unpack_from('<%dI' % (ssz // 4), sect(cur), 0))
        cur = fat[cur]

    def chain(start, mini=False):
        table = minifat if mini else fat
        out, cur, seen = [], start, set()
        while cur < 0xFFFFFFFA and cur not in seen and cur < len(table):
            seen.add(cur)
            out.append(cur)
            cur = table[cur]
        return out

    dirdata = b''.join(sect(s) for s in chain(dir_start))
    entries = []
    for i in range(len(dirdata) // 128):
        e = dirdata[i * 128:(i + 1) * 128]
        nlen = struct.unpack_from('<H', e, 0x40)[0]
        entries.append((
            e[:max(0, nlen - 2)].decode('utf-16-le', 'replace'),   # nome
            e[0x42],                                                # tipo
            struct.unpack_from('<I', e, 0x74)[0],                   # setor inicial
            struct.unpack_from('<Q', e, 0x78)[0],                   # tamanho
        ))

    root = next((e for e in entries if e[1] == 5), None)
    ministream = b''.join(sect(s) for s in chain(root[2])) if root else b''

    streams = {}
    for name, etype, start, size in entries:
        if etype != 2 or size == 0:
            continue
        if size < cutoff:
            buf = b''.join(ministream[s * msz:(s + 1) * msz] for s in chain(start, mini=True))
        else:
            buf = b''.join(sect(s) for s in chain(start))
        streams[name] = buf[:size]
    return streams


def _copytoken_masks(difference):
    bits = max(int(math.ceil(math.log(difference, 2))) if difference > 1 else 4, 4)
    length_mask = 0xFFFF >> bits
    return length_mask, (~length_mask) & 0xFFFF, bits


def vba_decompress(data):
    """[MS-OVBA] 2.4.1 - descompacta um CompressedContainer (byte 0x01)."""
    if not data or data[0] != 0x01:
        return None
    out = bytearray()
    i = 1
    while i + 1 < len(data):
        header = struct.unpack_from('<H', data, i)[0]
        i += 2
        size = (header & 0x0FFF) + 3
        if (header >> 12) & 0x07 != 0b011:
            break
        compressed = (header >> 15) & 0x01
        if not compressed:
            out += data[i:i + 4096]
            i += 4096
            continue
        end = i + size - 2
        chunk_start = len(out)
        while i < end and i < len(data):
            flags = data[i]
            i += 1
            for bit in range(8):
                if i >= end or i >= len(data):
                    break
                if not ((flags >> bit) & 1):
                    out.append(data[i])
                    i += 1
                else:
                    if i + 1 >= len(data):
                        return bytes(out)
                    token = struct.unpack_from('<H', data, i)[0]
                    i += 2
                    lmask, omask, bits = _copytoken_masks(len(out) - chunk_start)
                    length = (token & lmask) + 3
                    src = len(out) - (((token & omask) >> (16 - bits)) + 1)
                    if src < 0:
                        return bytes(out)
                    for k in range(length):
                        out.append(out[src + k])
    return bytes(out)


def extract_modules(vba_bin):
    """-> ({modulo: (pcode_bytes, fonte_str)}, {stream: bytes})

    O stream de um modulo e' PerformanceCache (P-code) seguido do
    CompressedContainer com o fonte. Localizamos o container pelo
    primeiro offset 0x01 que descompacta num texto com 'Attribute VB_Name'.
    """
    streams = cfb_streams(vba_bin)
    mods = {}
    for name, buf in streams.items():
        for off in range(len(buf)):
            if buf[off] != 0x01:
                continue
            txt = vba_decompress(buf[off:])
            if txt and b'Attribute VB_Name' in txt[:400]:
                mods[name] = (buf[:off], txt.decode('latin-1'))
                break
    return mods, streams


# ---- Fantasmas de P-code -------------------------------------------
# Identificadores que sobraram no cache compilado sem correspondencia
# no fonte. Num add-in construido a partir de um arquivo-base LIMPO
# essa lista tende a zero; entradas aqui costumam ser residuo de
# compilacoes anteriores no mesmo arquivo hospedeiro.
_VBA_INTERNAL = {
    # constantes de compilacao condicional e nomes do proprio runtime VBA;
    # aparecem no _VBA_PROJECT sem nunca terem estado no nosso fonte.
    'arm64', 'win16', 'win32', 'win64', 'vba6', 'vba7', 'mac',
    'mac_office_version', 'vbaproject', 'project1', 'module1', 'msforms',
    'datatype', 'thisdocument', 'stdole', 'office',
}


def pcode_ghosts(pcode, source, extra_allow=()):
    """Tokens presentes no P-code e ausentes do fonte (HEURISTICA).

    O P-code nao guarda os identificadores delimitados: quase sempre vem
    um byte binario (as vezes printavel) colado no inicio ou no fim, e o
    token aparece cortado. Duas regras de ruido, aplicadas em conjunto:

      a) algum prefixo de >= 6 chars existe no texto do fonte;
      b) o token, tirando ate' 2 chars do fim e/ou 1 do inicio, e' uma
         PALAVRA do fonte (ou um nome interno do VBA).

    O que sobra sao tokens sem nenhuma raiz no fonte -- o sinal que
    interessa. Nao e' prova: e' triagem. Ver docs/SECURITY-REVIEW-v1.5.0B.md §5.
    """
    src = source.lower()
    words = {w.lower() for w in re.findall(r'[A-Za-z_][A-Za-z0-9_]*', source)}
    words |= _VBA_INTERNAL
    allow = {a.lower() for a in extra_allow}
    ghosts = set()
    for m in re.finditer(rb'[A-Za-z_][A-Za-z0-9_]{5,}', pcode):
        tok = m.group().decode('latin-1')
        t = tok.lower()
        if t in allow or t.startswith('_') or re.fullmatch(r'[0-9a-f]{6,}', t):
            continue
        if any(t[:L] in src for L in range(len(t), 5, -1)):        # (a)
            continue
        trimmed = {t, t[:-1], t[:-2], t[1:], t[1:-1], t[1:-2]}     # (b)
        if trimmed & words:
            continue
        ghosts.add(tok)
    return sorted(ghosts)


if __name__ == '__main__':
    mods, streams = extract_modules(open(sys.argv[1], 'rb').read())
    print('streams: ' + ', '.join(sorted(streams)))
    for name, (pc, src) in sorted(mods.items()):
        print('modulo %s: p-code %d bytes, fonte %d linhas'
              % (name, len(pc), src.count('\n')))
