# -*- coding: utf-8 -*-
"""
Atualizador de OPEX — RP
Atualiza o staff (grupo RAMPA) a partir da folha e/ou a quantidade de voos
a partir da malha. Cada parte é independente: envie só o que quiser atualizar.
"""
import csv, io, os, re, tempfile, zipfile
from collections import defaultdict

import pandas as pd
import streamlit as st
from lxml import etree

st.set_page_config(page_title="Atualizador de OPEX · RP", page_icon="📊", layout="wide")

M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
Q = lambda t: '{%s}%s' % (M, t)
RELNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

BASES_PADRAO = ['BYO','CGR','CMG','CWB','DOU','FEN','FLN','IGU','JJG','JOI',
                'MGF','NVT','POA','UDI','XAP']

# ── staff ──
SIM  = {'Trabalhando', 'Férias', 'Atestado'}
ZERO = {'Auxílio Doença', 'Acidente Trabalho', 'Licença Maternidade',
        'Lic. Recl', 'Lic. s/ Remuneração', 'Suspensão Ajuizamento Rec Trab'}
FORA = {'Aposentadoria Invalidez'}
# MGF é a sede: só esse pedaço "operacional" entra na conferência; o resto da
# folha de MGF (gerência/administrativo corporativo) fica de fora, sem tocar
# e sem reportar.
MGF_ESCOPO = {'AGENTE LIDER','AUX SERV AEROPORTUARIO','AUX.SERV GERAIS','AUXILIAR DE RAMPA',
              'AUXILIAR LIDER DE LIMPEZA','AUXILIAR LIDER DE RAMPA','BALANCEIRO(A)',
              'MOTORISTA CAT D/E2','OPERADOR DE EQUIP CAT"D"','OPERADOR DE EQUIP CAT"E"',
              'OPERADOR DE RAMPA','SUPERV.OPERACIONAL','JOVEM APRENDIZ'}
C_STAFF = {'grupo': 19, 'funcao': 20, 'ch_dia': 22, 'qtde': 24}      # S, T, V, X

# ── voos ──
C_VOO = {'base': 3, 'cliente': 4, 'equip': 6, 'tipo': 7, 'qtde': 8}  # C, D, F, G, H


def colnum(ref):
    c = re.match(r'([A-Z]+)', ref).group(1); v = 0
    for ch in c: v = v * 26 + (ord(ch) - 64)
    return v


# ═══════════════════ nomes de função: folha x OPEX ═══════════════════
# A folha e o OPEX nem sempre escrevem a função igual. generico() é a
# normalização "de sempre" (tira o código da base no final, mantém a letra
# JR/PL/SR mas tira o dígito de sub-nível tipo PL3, e junta todo
# SUPERV.OPERACIONAL num bucket só porque a folha não separa nível pra esse
# cargo). sem_nivel() e chave_fuzzy() são usadas só como fallback, quando a
# chave "de sempre" não acha nada em nenhuma CH -- pra não reportar como
# "função nova" um cargo que já existe, só escrito diferente (ex.:
# "COORDENADOR GERAL SPAX" no OPEX == "COORDENADOR DE SPAX JR" na folha).
_BASES_SUFIXO = r'(BYO|CGR|CMG|CWB|DOU|FEN|FLN|IGU|JJG|JOI|MGF|NVT|POA|UDI|XAP)'


def generico(f):
    f = re.sub(rf'\s+{_BASES_SUFIXO}$', '', str(f).strip().upper())
    f = re.sub(r'(JR|PL|SR)\d+$', r'\1', f)          # PL3 -> PL (mantém a letra)
    return 'SUPERV.OPERACIONAL' if f.startswith('SUPERV.OPERACIONAL') else f


def sem_nivel(f):
    """Tira também a letra de nível sozinha no final -- para cargos em que a
    folha não diferencia nível nenhum (nem letra, nem dígito)."""
    return re.sub(r'\s+(JR|PL|SR)$', '', f).strip()


def chave_fuzzy(f):
    """Último recurso: tira parênteses, palavras de ligação, a letra de
    nível sozinha no final e um dígito solto no final -- resolve variação
    de redação tipo '(A)' vs 'DE', 'GERAL'/'FILIAL' presente só de um lado,
    'JR' sozinho vs ausente, ou um '1' sozinho no fim."""
    f = re.sub(r'\([^)]*\)', ' ', f)
    f = re.sub(r'\b(DE|DA|DO|GERAL|FILIAL)\b', ' ', f)
    f = re.sub(r'\s+', ' ', f).strip()
    f = re.sub(r'\s+(JR|PL|SR)$', '', f)
    f = re.sub(r'\d+$', '', f)
    return re.sub(r'\s+', ' ', f).strip()


# ═══════════════════════ FOLHA (FPRE109) ═══════════════════════
def ler_folha_xls(dados):
    """Formato .xls, inclusive o BIFF2 antigo que o sistema exporta."""
    erros = []
    for motor in ('xlrd', None):
        try:
            df = pd.read_excel(io.BytesIO(dados), header=None, dtype=object,
                               **({'engine': motor} if motor else {}))
            return [{i + 1: v for i, v in enumerate(row)
                     if v is not None and str(v) not in ('nan', 'NaT', '')}
                    for row in df.values.tolist()]
        except Exception as e:
            erros.append(f'{motor or "auto"}: {e}')
    raise ValueError('não consegui ler este .xls (' + ' | '.join(erros)[:160] + ')')


def ler_folha(dados):
    """Lê o FPRE109 mesmo com o XML fora do padrão (caminhos com barra invertida)."""
    if dados[:2] != b'PK':                      # não é xlsx
        return ler_folha_xls(dados)
    z = zipfile.ZipFile(io.BytesIO(dados))
    nomes = {i.filename.replace('\\', '/'): i.filename for i in z.infolist()}
    sst = []
    if 'xl/sharedStrings.xml' in nomes:
        r = etree.fromstring(z.read(nomes['xl/sharedStrings.xml']))
        sst = [''.join(t.text or '' for t in si.iter(Q('t'))) for si in r]
    sheet = next((n for n in nomes if re.search(r'sheet\d*\.xml$', n)), None)
    if not sheet:
        raise ValueError('não encontrei a planilha dentro do arquivo')
    root = etree.fromstring(z.read(nomes[sheet]))
    linhas = []
    for row in root.iter(Q('row')):
        d = {}
        for c in row:
            t = c.get('t'); v = c.find(Q('v'))
            if t == 'inlineStr' or v is None:
                i = c.find(Q('is'))
                val = ''.join(x.text or '' for x in i.iter(Q('t'))) if i is not None else None
            elif t == 's':
                val = sst[int(v.text)]
            else:
                val = v.text
            if val not in (None, ''): d[colnum(c.get('r'))] = val
        if d: linhas.append(d)
    return linhas


def ch_mensal(v):
    """A folha traz a CH semanal em decimal (7,5 = 36h/sem = 180h/mês)."""
    s = str(v)
    if ':' in s:                                    # formato antigo '210:00'
        return int(s.split(':')[0])
    try: x = float(s.replace(',', '.'))
    except ValueError: return None
    return round(x * 24) if x < 24 else round(x)


def agregar_folha(linhas):
    base, col = None, {}
    out = defaultdict(lambda: defaultdict(lambda: {'sim': 0, 'zero': 0, 'fora': 0}))
    for d in linhas:
        vals = {k: str(v).strip() for k, v in d.items() if isinstance(v, str)}
        if 'Cadastro' in vals.values() and 'Cargo' in vals.values():
            for k, v in vals.items():
                if v == 'Cargo': col['cargo'] = k
                elif v.replace(' ', '').startswith('C.Hor'): col['ch'] = k
                elif v == 'Situação': col['sit'] = k
            continue
        a = d.get(1)
        if isinstance(a, str) and re.match(r'^\d{3}\s*,\s*\w+', a.strip()):
            base = a.split(',')[1].strip(); continue
        if not (base and col): continue
        cargo, ch, sit = d.get(col.get('cargo')), d.get(col.get('ch')), d.get(col.get('sit'))
        if ch is None:                              # a CH às vezes escorrega de coluna
            for k in range(col.get('ch', 11) - 1, col.get('ch', 11) + 2):
                if k in d and re.match(r'^[\d.,:]+$', str(d[k])): ch = d[k]; break
        if not (cargo and ch and sit): continue
        cargo = generico(cargo)                      # mesma normalização dos dois lados
        chm = ch_mensal(ch)
        if chm is None: continue
        sit = str(sit).strip(); k = (cargo, chm)
        if sit in SIM:    out[base][k]['sim'] += 1
        elif sit in ZERO: out[base][k]['zero'] += 1
        elif sit in FORA: out[base][k]['fora'] += 1
    return {b: dict(v) for b, v in out.items()}


# ═══════════════════════ MALHA ═══════════════════════
def ler_malha(dados, nome):
    if nome.lower().endswith(('.xlsx', '.xlsm', '.xls')):
        df = pd.read_excel(io.BytesIO(dados), dtype=str)
    else:
        txt = dados.decode('utf-8-sig', errors='replace')
        sep = ';' if txt.count(';') > txt.count(',') else ','
        df = pd.read_csv(io.StringIO(txt), sep=sep, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    need = ['Base', 'Nome', 'Mod', 'Tipo']
    faltando = [n for n in need if n not in df.columns]
    if faltando:
        raise ValueError('faltam as colunas: ' + ', '.join(faltando))
    return df.to_dict('records')


def equip_de(tipo, rotulos):
    """Escolhe o rótulo do equipamento conforme o que já existe naquela base/cliente."""
    t = str(tipo).strip().upper()
    if t in ('B737', 'B738'):   cand = ['737/738']
    elif t == 'ATR':            cand = ['AT9']
    elif t == '295':            cand = ['295']
    elif t in ('E190', 'E195'): cand = ['E1']
    elif t == 'C208':           cand = ['CARAVAN']
    elif t == 'A321':           cand = ['321', '319/320']
    elif t in ('A319', 'A320'): cand = ['319/320', '320']
    else:                       cand = [t]
    for c in cand:
        if c in rotulos: return c
    return cand[0]


def tipo_de(cia, mod):
    """LATAM não usa PNT: o pernoite dela vem como TST.N."""
    if str(cia).upper() == 'LATAM':
        return 'PNT' if str(mod).strip() == 'TST.N' else 'TST'
    return 'PNT' if str(mod).strip() == 'PNT' else 'TST'


# ═══════════════════════ OPEX ═══════════════════════
def valor(c, sst):
    t = c.get('t'); v = c.find(Q('v'))
    if t == 'inlineStr':
        i = c.find(Q('is'))
        return ''.join(x.text or '' for x in i.iter(Q('t'))) if i is not None else None
    if v is None: return None
    return sst[int(v.text)] if t == 's' else v.text


def escreve(cel, coluna, num):
    if coluna not in cel: return False
    c = cel[coluna]
    for tag in ('v', 'f', 'is'):
        e = c.find(Q(tag))
        if e is not None: c.remove(e)
    if 't' in c.attrib: del c.attrib['t']
    etree.SubElement(c, Q('v')).text = str(num)
    return True


def processar(opex_bytes, folha, malha, bases, manter):
    pasta = tempfile.mkdtemp()
    with zipfile.ZipFile(io.BytesIO(opex_bytes)) as z: z.extractall(pasta)
    wb = etree.parse(f'{pasta}/xl/workbook.xml').getroot()
    rels = {r.get('Id'): r.get('Target')
            for r in etree.parse(f'{pasta}/xl/_rels/workbook.xml.rels').getroot()}
    mapa = {s.get('name'): rels[s.get(RELNS + 'id')] for s in wb.find(Q('sheets'))}
    sst = []
    p = f'{pasta}/xl/sharedStrings.xml'
    if os.path.exists(p):
        sst = [''.join(t.text or '' for t in si.iter(Q('t'))) for si in etree.parse(p).getroot()]

    rel = {'staff': [], 'realinhado': [], 'faltando': [], 'zerar': [],
           'sem_correspondencia': [], 'voos': [], 'sem_linha': [], 'ausentes': [],
           'comparativo_staff': [], 'comparativo_voos': []}

    # ── STAFF ──
    if folha:
        for b in bases:
            if b not in mapa or b in manter: continue
            if b not in folha:                  # base sem ninguém na folha: não zera, avisa
                rel['ausentes'].append(b); continue

            # por_funcao[fkey] = {chm: qtde} -- mutado (consumido) à medida
            # que vai casando; zeros = combos só com afastados (reporta 0)
            por_funcao = defaultdict(dict)
            zeros = {}
            for (cargo, chm), v in folha.get(b, {}).items():
                if v['sim']: por_funcao[cargo][chm] = v['sim']
                elif v['zero']: zeros[(cargo, chm)] = 0
            fkeys_conhecidos = set(por_funcao.keys())   # snapshot ANTES de consumir

            arq = f"{pasta}/xl/{mapa[b]}"
            tree = etree.parse(arq)
            rows = {int(r.get('r')): r for r in tree.getroot().find(Q('sheetData'))}

            linhas_base, orfaos = [], []
            for rn in sorted(rows):
                cel = {colnum(c.get('r')): c for c in rows[rn]}
                g = valor(cel[C_STAFF['grupo']], sst) if C_STAFF['grupo'] in cel else None
                f = valor(cel[C_STAFF['funcao']], sst) if C_STAFF['funcao'] in cel else None
                if not f: continue
                if b == 'MGF' and generico(f) not in MGF_ESCOPO: continue   # sede: fora
                chd = valor(cel[C_STAFF['ch_dia']], sst) if C_STAFF['ch_dia'] in cel else None
                try: chm = round(float(chd) * 30)
                except (TypeError, ValueError): chm = 0
                atual = valor(cel[C_STAFF['qtde']], sst) if C_STAFF['qtde'] in cel else None
                try: atual = int(float(atual)) if atual not in (None, '') else 0
                except ValueError: continue
                fkey = generico(f)
                linhas_base.append((rn, cel, f, fkey, chm, atual))

            # soma corrente do que fica no OPEX depois de atualizar (começa
            # igual ao que já estava; cada escreve() abaixo ajusta a diferença)
            # -- usada só pro comparativo base a base, não afeta a planilha
            soma_opex_b = sum(atual for *_, atual in linhas_base)

            # passo 1 — casamento exato (mesma função, mesma CH)
            for rn, cel, f, fkey, chm, atual in linhas_base:
                bucket = por_funcao.get(fkey)
                if bucket and bucket.get(chm, 0) > 0:
                    novo = bucket[chm]; bucket[chm] = 0
                    if atual != novo and escreve(cel, C_STAFF['qtde'], novo):
                        soma_opex_b += novo - atual
                        rel['staff'].append({'Base': b, 'Função': f, 'CH': chm,
                                             'De': atual, 'Para': novo, 'Dif': novo - atual})
                else:
                    orfaos.append((rn, cel, f, fkey, chm, atual))

            # passo 2 — pareamento global das órfãs: cada linha órfã pode
            # casar com (a) outra CH da MESMA função já conhecida (ex.:
            # gerente que passou de 220h pra 180h) ou, só quando o nome
            # exato não existe NA FOLHA EM CH NENHUMA, (b) o nome "parecido"
            # (variação de redação, tipo "(A)" vs "DE"). Resolve tudo de uma
            # vez, sempre escolhendo globalmente a CH mais próxima primeiro
            # -- evita que uma linha "errada" (de outra CH, mera coincidência
            # de ordem) roube a sobra que seria da linha certa.
            restantes = list(orfaos)
            while restantes:
                melhor = None   # (distancia, índice em restantes, fk candidata, ch2)
                for i, (rn, cel, f, fkey, chm, atual) in enumerate(restantes):
                    candidatas_fk = [fkey] if fkey in por_funcao else \
                        [fk for fk in por_funcao if chave_fuzzy(fk) == chave_fuzzy(fkey)]
                    for fk in candidatas_fk:
                        for ch2, q in por_funcao[fk].items():
                            if q <= 0: continue
                            d = abs(ch2 - chm)
                            if melhor is None or d < melhor[0]:
                                melhor = (d, i, fk, ch2)
                if melhor is None: break
                _, i, fk, ch2 = melhor
                rn, cel, f, fkey, chm, atual = restantes.pop(i)
                novo = por_funcao[fk][ch2]; por_funcao[fk][ch2] = 0
                if escreve(cel, C_STAFF['ch_dia'], round(ch2 / 30, 2)):
                    escreve(cel, C_STAFF['qtde'], novo)
                    soma_opex_b += novo - atual
                    nome = f if fk == fkey else f'{f} (= "{fk}" na folha)'
                    rel['realinhado'].append({'Base': b, 'Função': nome, 'CH_antiga': chm,
                                              'CH_nova': ch2, 'Qtde': novo})

            # passo 3 — o que não casou de jeito nenhum
            for rn, cel, f, fkey, chm, atual in restantes:
                if fkey in fkeys_conhecidos:
                    # a folha CONHECE essa função -- só não sobrou nenhuma
                    # pessoa pra essa linha (já foi toda pra outra CH/linha
                    # irmã). Isso é normal (zera, igual sempre fez).
                    if atual != 0 and escreve(cel, C_STAFF['qtde'], 0):
                        soma_opex_b += -atual
                        rel['staff'].append({'Base': b, 'Função': f, 'CH': chm,
                                             'De': atual, 'Para': 0, 'Dif': -atual})
                elif atual > 0:
                    # a folha NUNCA viu essa função nessa base, em CH nenhuma
                    # -- não dá pra saber se é gente que saiu ou só um nome
                    # escrito diferente. NÃO zera sozinho, só avisa.
                    rel['sem_correspondencia'].append({'Base': b, 'Função': f, 'CH': chm, 'Qtde': atual})

            for fk, bucket in por_funcao.items():
                if b == 'MGF' and fk not in MGF_ESCOPO: continue   # sede: fora de escopo
                for chm2, q in bucket.items():
                    if q > 0:
                        rel['faltando'].append({'Base': b, 'Função': fk, 'CH': chm2, 'Qtde': q})
            for k in zeros:
                rel['zerar'].append({'Base': b, 'Função': k[0], 'CH': k[1]})
            tree.write(arq, xml_declaration=True, encoding='UTF-8', standalone=True)

            # comparativo base a base: o que fica no OPEX depois de
            # atualizar x o que a folha (109) diz que deveria ter
            total_109_b = sum(v['sim'] for (cargo, _), v in folha.get(b, {}).items()
                              if not (b == 'MGF' and cargo not in MGF_ESCOPO))
            rel['comparativo_staff'].append({
                'Base': b, 'OPEX (após atualizar)': soma_opex_b, '109': total_109_b,
                'Diferença': soma_opex_b - total_109_b,
            })

    # ── VOOS ──
    if malha and 'Voos&Tarifas' in mapa:
        arq = f"{pasta}/xl/{mapa['Voos&Tarifas']}"
        tree = etree.parse(arq)
        rows = {int(r.get('r')): r for r in tree.getroot().find(Q('sheetData'))}
        linhas, rotulos = [], defaultdict(set)
        for rn in sorted(rows):
            cel = {colnum(c.get('r')): c for c in rows[rn]}
            base = valor(cel[C_VOO['base']], sst) if C_VOO['base'] in cel else None
            cli  = valor(cel[C_VOO['cliente']], sst) if C_VOO['cliente'] in cel else None
            eq   = valor(cel[C_VOO['equip']], sst) if C_VOO['equip'] in cel else None
            tp   = valor(cel[C_VOO['tipo']], sst) if C_VOO['tipo'] in cel else None
            if not (base and cli and eq and tp): continue
            if str(tp).strip() not in ('TST', 'PNT'): continue          # pula cabeçalho
            base, cli = str(base).strip(), str(cli).strip().upper()
            eq, tp = str(eq).strip(), str(tp).strip()
            linhas.append((cel, base, cli, eq, tp))
            rotulos[(cli, base)].add(eq)

        cont = defaultdict(int)
        for x in malha:
            cia, base = str(x.get('Nome', '')).strip().upper(), str(x.get('Base', '')).strip()
            if not cia or not base: continue
            e = equip_de(x.get('Tipo'), rotulos.get((cia, base), set()))
            cont[(cia, base, e, tipo_de(cia, x.get('Mod')))] += 1

        usados = set()
        # comparativo por (Base, Cia) -- não só por base: um erro tipo "uma
        # cia aérea a mais numa base" só aparece nesse nível de detalhe
        soma_opex_voos = defaultdict(int)
        for cel, base, cli, eq, tp in linhas:
            k = (cli, base, eq, tp)
            novo = cont.get(k, 0) if k not in usados else 0
            usados.add(k)
            atual = valor(cel[C_VOO['qtde']], sst) if C_VOO['qtde'] in cel else None
            try: atual = int(float(atual)) if atual not in (None, '') else 0
            except (TypeError, ValueError): continue
            soma_opex_voos[(base, cli)] += atual
            if atual != novo and escreve(cel, C_VOO['qtde'], novo):
                soma_opex_voos[(base, cli)] += novo - atual
                rel['voos'].append({'Base': base, 'Cliente': cli, 'Equip': eq, 'Tipo': tp,
                                    'De': atual, 'Para': novo, 'Dif': novo - atual})
        for k, v in cont.items():
            if k not in usados and v > 0:
                rel['sem_linha'].append({'Base': k[1], 'Cliente': k[0], 'Equip': k[2],
                                         'Tipo': k[3], 'Voos': v})
        tree.write(arq, xml_declaration=True, encoding='UTF-8', standalone=True)

        # comparativo Base × Cia: voos que ficam no OPEX x voos da malha
        soma_malha_voos = defaultdict(int)
        for k, v in cont.items():
            cia, base_m = k[0], k[1]
            soma_malha_voos[(base_m, cia)] += v
        for (base, cli) in sorted(set(soma_opex_voos) | set(soma_malha_voos)):
            o, m = soma_opex_voos.get((base, cli), 0), soma_malha_voos.get((base, cli), 0)
            rel['comparativo_voos'].append({
                'Base': base, 'Cia': cli, 'OPEX (após atualizar)': o, 'Malha': m,
                'Diferença': o - m,
            })

    # ── recalcular ao abrir ──
    wp = f'{pasta}/xl/workbook.xml'; wt = etree.parse(wp)
    calc = wt.getroot().find(Q('calcPr'))
    if calc is None: calc = etree.SubElement(wt.getroot(), Q('calcPr'))
    calc.set('fullCalcOnLoad', '1')
    wt.write(wp, xml_declaration=True, encoding='UTF-8', standalone=True)
    rp = f'{pasta}/xl/_rels/workbook.xml.rels'; rt = etree.parse(rp)
    for r in list(rt.getroot()):
        if r.get('Target', '').endswith('calcChain.xml'): rt.getroot().remove(r)
    rt.write(rp, xml_declaration=True, encoding='UTF-8', standalone=True)
    cp = f'{pasta}/[Content_Types].xml'; ct = etree.parse(cp)
    for ov in list(ct.getroot()):
        if ov.get('PartName') == '/xl/calcChain.xml': ct.getroot().remove(ov)
    ct.write(cp, xml_declaration=True, encoding='UTF-8', standalone=True)
    if os.path.exists(f'{pasta}/xl/calcChain.xml'): os.remove(f'{pasta}/xl/calcChain.xml')

    buf = io.BytesIO()
    arqs = []
    for root, _, fs in os.walk(pasta):
        for f in fs: arqs.append(os.path.relpath(os.path.join(root, f), pasta).replace(os.sep, '/'))
    arqs.sort(key=lambda x: (x != '[Content_Types].xml', x))
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        for a in arqs: z.write(os.path.join(pasta, a), a)
    return buf.getvalue(), rel


# ═══════════════════════ interface ═══════════════════════
st.title("📊 Atualizador de OPEX · RP")
st.caption("Envie o OPEX e o que quiser atualizar. **Staff e voos são independentes** — "
           "dá para atualizar só um dos dois.")

f_opex = st.file_uploader("**OPEX** — obrigatório", type=['xlsx', 'xlsm'])

c1, c2 = st.columns(2)
with c1:
    st.markdown("**1 · Staff** — todos os grupos nas abas de base (MGF só a parte operacional)")
    f_folha = st.file_uploader("Folha (FPRE109) — .xlsx ou .xls", type=['xlsx', 'xls'], key="folha")
with c2:
    st.markdown("**2 · Voos** — aba Voos&Tarifas")
    f_malha = st.file_uploader("Malha (CSV padrão Malha RP)", type=['csv', 'xlsx'], key="malha")

with st.sidebar:
    st.header("Opções")
    bases = st.multiselect("Bases a atualizar (staff)", BASES_PADRAO, default=BASES_PADRAO)
    manter = st.multiselect("Manter como está", BASES_PADRAO, default=['FEN'],
                            help="Bases que não vêm na folha, como FEN.")
    st.divider()
    st.caption("**Staff — regras de situação**")
    st.write("Conta: Trabalhando, Férias, Atestado")
    st.write("Fica com 0: Auxílio Doença, Acidente, Maternidade, Licenças")
    st.write("Fora: Aposentadoria por Invalidez")
    st.caption("**Voos — tipo de atendimento**")
    st.write("LATAM: TST.N vira PNT; o resto é TST")
    st.write("Demais cias: PNT quando o solo passa de 4h")

pronto = f_opex and (f_folha or f_malha)
if not f_opex:
    st.info("Comece enviando o OPEX.")
elif not pronto:
    st.warning("Envie a folha, a malha, ou as duas.")

if st.button("Atualizar OPEX", type="primary", use_container_width=True, disabled=not pronto):
    folha = malha = None
    if f_folha:
        try:
            folha = agregar_folha(ler_folha(f_folha.getvalue()))
            if not folha:
                st.error("Não encontrei ninguém na folha."); st.stop()
            tot = sum(v['sim'] for d in folha.values() for v in d.values())
            st.success(f"Folha: **{tot} pessoas** (Trabalhando/Férias/Atestado) em {len(folha)} bases.")
        except Exception as e:
            st.error(f"Não consegui ler a folha: {e}"); st.stop()
    if f_malha:
        try:
            malha = ler_malha(f_malha.getvalue(), f_malha.name)
            if not malha:
                st.error("A malha está vazia — nenhum voo lido. Confira o arquivo."); st.stop()
            st.success(f"Malha: **{len(malha)} voos**.")
        except Exception as e:
            st.error(f"Não consegui ler a malha: {e}"); st.stop()

    try:
        saida, rel = processar(f_opex.getvalue(), folha, malha, bases, manter)
    except Exception as e:
        st.error(f"Não consegui atualizar o OPEX: {e}"); st.stop()

    a, b_, c = st.columns(3)
    a.metric("Staff — células alteradas", len(rel['staff']) + len(rel['realinhado']))
    b_.metric("Voos — linhas alteradas", len(rel['voos']))
    c.metric("Pendências", len(rel['faltando']) + len(rel['sem_linha']) + len(rel['sem_correspondencia']))

    st.subheader("Comparativo base a base")
    comp1, comp2 = st.columns(2)
    with comp1:
        st.caption("**Staff** — OPEX depois de atualizar × 109")
        if rel['comparativo_staff']:
            dfc = pd.DataFrame(rel['comparativo_staff'])
            dfc.insert(0, '', dfc['Diferença'].apply(lambda d: '✅' if d == 0 else '⚠️'))
            st.dataframe(dfc, use_container_width=True, hide_index=True)
        else:
            st.caption("Envie a folha pra ver esse comparativo.")
    with comp2:
        st.caption("**Voos** — OPEX depois de atualizar × malha")
        if rel['comparativo_voos']:
            dfc2 = pd.DataFrame(rel['comparativo_voos'])
            dfc2.insert(0, '', dfc2['Diferença'].apply(lambda d: '✅' if d == 0 else '⚠️'))
            st.dataframe(dfc2, use_container_width=True, hide_index=True)
        else:
            st.caption("Envie a malha pra ver esse comparativo.")

    if folha:
        st.subheader("Staff")
        if rel['staff']:
            df = pd.DataFrame(rel['staff'])
            st.dataframe(df, use_container_width=True, hide_index=True)
            st.dataframe(df.groupby('Base')['Dif'].sum().reset_index().rename(
                columns={'Dif': 'Saldo'}), use_container_width=True, hide_index=True)
        else:
            st.info("O staff já estava batendo com a folha.")
        if rel['realinhado']:
            st.info("Linha que já existia, só mudou de CH (realinhada em vez de reportada "
                    "como função nova):")
            st.dataframe(pd.DataFrame(rel['realinhado']), use_container_width=True, hide_index=True)
        if rel['faltando']:
            st.warning("Linhas que você precisa criar no OPEX (a função existe na folha, "
                       "mas não há linha na aba):")
            st.dataframe(pd.DataFrame(rel['faltando']), use_container_width=True, hide_index=True)
        if rel['sem_correspondencia']:
            st.warning("Função que está no OPEX com gente, mas a folha não tem esse nome em "
                       "CH nenhuma — **não mexi na quantidade**, confira se é gente que saiu "
                       "ou só um nome escrito diferente:")
            st.dataframe(pd.DataFrame(rel['sem_correspondencia']), use_container_width=True, hide_index=True)
        if rel['ausentes']:
            st.error("Estas bases **não apareceram na folha** e foram mantidas como estavam "
                     "(nada foi zerado): " + ", ".join(rel['ausentes']) +
                     ". Se elas realmente não têm mais ninguém, zere na mão.")
        if rel['zerar']:
            st.caption("Funções só com afastados — devem entrar com 0:")
            st.dataframe(pd.DataFrame(rel['zerar']), use_container_width=True, hide_index=True)

    if malha:
        st.subheader("Voos")
        if rel['voos']:
            df = pd.DataFrame(rel['voos'])
            st.dataframe(df, use_container_width=True, hide_index=True)
            st.dataframe(df.groupby('Base')[['De', 'Para']].sum().reset_index(),
                         use_container_width=True, hide_index=True)
        else:
            st.info("Os voos já estavam batendo com a malha.")
        if rel['sem_linha']:
            st.warning("Voos sem linha de tarifa no OPEX (não foram lançados):")
            st.dataframe(pd.DataFrame(rel['sem_linha']), use_container_width=True, hide_index=True)

    nome = f_opex.name.rsplit('.', 1)[0] + '_ATUALIZADO.xlsx'
    st.download_button("⬇️ Baixar OPEX atualizado", saida, file_name=nome,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary", use_container_width=True)
