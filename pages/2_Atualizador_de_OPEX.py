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
        'Lic. Recl', 'Lic. s/ Remuneração'}
FORA = {'Aposentadoria Invalidez'}
RAMPA = {'AGENTE LIDER','AUX SERV AEROPORTUARIO','AUX.SERV GERAIS','AUXILIAR DE RAMPA',
         'AUXILIAR LIDER DE LIMPEZA','AUXILIAR LIDER DE RAMPA','BALANCEIRO(A)',
         'MOTORISTA CAT D/E2','OPERADOR DE EQUIP CAT"D"','OPERADOR DE EQUIP CAT"E"',
         'OPERADOR DE RAMPA','SUPERV.OPERACIONAL'}
C_STAFF = {'grupo': 19, 'funcao': 20, 'ch_dia': 22, 'qtde': 24}      # S, T, V, X

# ── voos ──
C_VOO = {'base': 3, 'cliente': 4, 'equip': 6, 'tipo': 7, 'qtde': 8}  # C, D, F, G, H


def colnum(ref):
    c = re.match(r'([A-Z]+)', ref).group(1); v = 0
    for ch in c: v = v * 26 + (ord(ch) - 64)
    return v


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
        cargo = str(cargo).strip()
        if cargo not in RAMPA: continue
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
def generico(f):
    f = re.sub(r'\s+(BYO|CGR|CMG|CWB|DOU|FEN|FLN|IGU|JJG|JOI|MGF|NVT|POA|UDI|XAP)$',
               '', str(f).strip())
    return 'SUPERV.OPERACIONAL' if f.startswith('SUPERV.OPERACIONAL') else f


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

    rel = {'staff': [], 'faltando': [], 'zerar': [], 'voos': [], 'sem_linha': [], 'ausentes': []}

    # ── STAFF ──
    if folha:
        for b in bases:
            if b not in mapa or b in manter: continue
            if b not in folha:                  # base sem ninguém na folha: não zera, avisa
                rel['ausentes'].append(b); continue
            alvo, zeros = defaultdict(int), {}
            for (cargo, chm), v in folha.get(b, {}).items():
                if v['sim']: alvo[(cargo, chm)] = v['sim']
                elif v['zero']: zeros[(cargo, chm)] = 0
            arq = f"{pasta}/xl/{mapa[b]}"
            tree = etree.parse(arq)
            rows = {int(r.get('r')): r for r in tree.getroot().find(Q('sheetData'))}
            vistos = set()
            for rn in sorted(rows):
                cel = {colnum(c.get('r')): c for c in rows[rn]}
                g = valor(cel[C_STAFF['grupo']], sst) if C_STAFF['grupo'] in cel else None
                f = valor(cel[C_STAFF['funcao']], sst) if C_STAFF['funcao'] in cel else None
                if g != 'RAMPA' or not f: continue
                chd = valor(cel[C_STAFF['ch_dia']], sst) if C_STAFF['ch_dia'] in cel else None
                try: chm = round(float(chd) * 30)
                except (TypeError, ValueError): chm = 0
                k = (generico(f), chm)
                novo = alvo.get(k, 0) if k not in vistos else 0
                vistos.add(k)
                atual = valor(cel[C_STAFF['qtde']], sst) if C_STAFF['qtde'] in cel else None
                try: atual = int(float(atual)) if atual not in (None, '') else 0
                except ValueError: continue
                if atual != novo and escreve(cel, C_STAFF['qtde'], novo):
                    rel['staff'].append({'Base': b, 'Função': f, 'CH': chm,
                                         'De': atual, 'Para': novo, 'Dif': novo - atual})
            for k, q in alvo.items():
                if k not in vistos and q > 0:
                    rel['faltando'].append({'Base': b, 'Função': k[0], 'CH': k[1], 'Qtde': q})
            for k in zeros:
                if k not in vistos:
                    rel['zerar'].append({'Base': b, 'Função': k[0], 'CH': k[1]})
            tree.write(arq, xml_declaration=True, encoding='UTF-8', standalone=True)

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
        for cel, base, cli, eq, tp in linhas:
            k = (cli, base, eq, tp)
            novo = cont.get(k, 0) if k not in usados else 0
            usados.add(k)
            atual = valor(cel[C_VOO['qtde']], sst) if C_VOO['qtde'] in cel else None
            try: atual = int(float(atual)) if atual not in (None, '') else 0
            except (TypeError, ValueError): continue
            if atual != novo and escreve(cel, C_VOO['qtde'], novo):
                rel['voos'].append({'Base': base, 'Cliente': cli, 'Equip': eq, 'Tipo': tp,
                                    'De': atual, 'Para': novo, 'Dif': novo - atual})
        for k, v in cont.items():
            if k not in usados and v > 0:
                rel['sem_linha'].append({'Base': k[1], 'Cliente': k[0], 'Equip': k[2],
                                         'Tipo': k[3], 'Voos': v})
        tree.write(arq, xml_declaration=True, encoding='UTF-8', standalone=True)

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
    st.markdown("**1 · Staff** — grupo RAMPA nas abas de base")
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
                st.error("Não encontrei ninguém do grupo RAMPA na folha."); st.stop()
            tot = sum(v['sim'] for d in folha.values() for v in d.values())
            st.success(f"Folha: **{tot} pessoas** de RAMPA em {len(folha)} bases.")
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
    a.metric("Staff — células alteradas", len(rel['staff']))
    b_.metric("Voos — linhas alteradas", len(rel['voos']))
    c.metric("Pendências", len(rel['faltando']) + len(rel['sem_linha']))

    if folha:
        st.subheader("Staff")
        if rel['staff']:
            df = pd.DataFrame(rel['staff'])
            st.dataframe(df, use_container_width=True, hide_index=True)
            st.dataframe(df.groupby('Base')['Dif'].sum().reset_index().rename(
                columns={'Dif': 'Saldo'}), use_container_width=True, hide_index=True)
        else:
            st.info("O staff já estava batendo com a folha.")
        if rel['faltando']:
            st.warning("Linhas que você precisa criar no OPEX (a função existe na folha, "
                       "mas não há linha na aba):")
            st.dataframe(pd.DataFrame(rel['faltando']), use_container_width=True, hide_index=True)
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
