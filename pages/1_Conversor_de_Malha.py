# -*- coding: utf-8 -*-
"""
Conversor de Malha — RP
Lê as malhas da GOL, AZUL, LATAM e as manuais e gera o CSV no padrão "Malha RP".
"""
import io, re, csv, json, unicodedata
from datetime import datetime, timedelta, date
from collections import defaultdict

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Conversor de Malha · RP", page_icon="✈️", layout="wide")

# ───────────────────────── configuração ─────────────────────────
CIAS_POR_BASE = {
    "BYO": ["AZUL", "GOL", "LATAM"],
    "CGR": ["AZUL", "GOL", "LATAM"],
    "CMG": ["AZUL"],
    "CWB": ["LATAM"],
    "DOU": ["LATAM"],
    "FEN": ["AZUL", "GOL", "LATAM"],
    "FLN": ["ANDES", "AZUL", "GOL", "LATAM", "SKY"],
    "IGU": ["AZUL", "GOL", "JETSMART", "LATAM"],
    "JJG": ["LATAM"],
    "JOI": ["AZUL", "GOL", "LATAM"],
    "MGF": ["AZUL", "GOL", "LATAM"],
    "NVT": ["AZUL", "GOL", "LATAM"],
    "POA": ["GOL", "LATAM", "SKY"],
    "UDI": ["AZUL", "GOL", "LATAM"],
    "XAP": ["AZUL", "GOL", "LATAM"],
}
COD_CLIENTE = {
    "AZUL": "0000000061", "GOL": "0000000151", "LATAM": "0000010717",
    "ANDES": "0000000221", "SKY": "0000000011", "JETSMART": "0000001012",
}
ICAO = {"AZUL": "AZU", "GOL": "GLO", "LATAM": "TAM", "JETSMART": "JAT"}

# equipamento da fonte -> tipo padronizado
EQUIP_TIPO = {
    "295":"295","32A":"A320","32N":"A320","32Q":"A321","330":"A330","332":"A330","339":"A330",
    "733":"B737","738":"B738","73A":"B737","73G":"B737","73H":"B737","73M":"B737",
    "73P":"B737 cargo","73T":"B737","73X":"B737","7M8":"B738","A321":"A321","A330":"A330",
    "AT42":"ATR","AT9":"ATR","B727":"B727","B737":"B737","B738":"B738","B767":"B767","B777":"B777",
    "E95":"E195","EJF":"Embraer cargo","320":"A320","321":"A321","319":"A319","767":"B767",
    "AT72":"ATR","C208":"C208","B747":"B747","777":"B777","A320":"A320","A319":"A319","ATF":"ATR",
    "CN1":"C208","CNF":"C208","787":"B787","B748":"B748","B787":"B787","FK70":"FK70","7MI":"B738",
    "A380":"A380","7ME":"B738","789":"B789","7ML":"B738","38R":"A320","31R":"A321","B190":"B190",
    "73C":"B738","38A":"A320","359":"A350",
}
# tipo -> código do sistema
TIPO_COD = {
    "A124":"53","A300":"6","A318":"2","A319":"3","319":"3","A320":"4","320":"4","A321":"56",
    "321":"56","A330":"10","339":"10","A343":"16","A346":"17","A350":"57","A380":"90","AT42":"51",
    "AT72":"15","B727":"11","B737":"1","73P":"1","B738":"1","B747":"14","B748":"62","B757":"8",
    "B767":"12","767":"12","B777":"21","777":"21","B787":"55","E190":"9","EJF":"9","C208":"47",
    "E195":"9","E90":"9","E95":"9","295":"144","ATR":"15","FK70":"18","B789":"55","B190":"104",
    "A350":"57","B737 cargo":"1","Embraer cargo":"9",
}
MELI = {"73C"}                       # Gol Meli — fora por padrão
CAB = ["Base","Cód. Cliente","Nome","Data","Mod","Tipo","Voo","Hora Chegada","Hora Saída","ICAO","AERONAVE"]

LIM_PNT   = 240     # minutos de solo: acima disso vira PNT
LAT_NOITE = (22*60, 5*60)   # LATAM: saída >=22:00 ou <=05:00 -> TST.N


# ───────────────────────── utilidades ─────────────────────────
def hhmm(v):
    """Normaliza qualquer representação de hora para 'HH:MM'."""
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.strftime("%H:%M")
    if hasattr(v, "hour") and hasattr(v, "minute") and not isinstance(v, (int, float)):
        return f"{v.hour:02d}:{v.minute:02d}"
    if isinstance(v, (int, float)):                       # fração do dia
        m = int(round(float(v) * 24 * 60)) % 1440
        return f"{m//60:02d}:{m%60:02d}"
    s = str(v).strip()
    m = re.search(r"(\d{1,2}):(\d{2})", s)
    if m:
        return f"{int(m.group(1))%24:02d}:{m.group(2)}"
    return None


def data_br(v):
    if isinstance(v, datetime):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, date):
        return v.strftime("%d/%m/%Y")
    s = str(v).strip()
    for f in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s[:19], f).strftime("%d/%m/%Y")
        except ValueError:
            pass
    return None


def para_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    for f in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s[:19], f).date()
        except ValueError:
            pass
    return None


def solo_min(sta, std):
    """Minutos de solo entre chegada e saída, tratando virada de meia-noite."""
    if not sta or not std:
        return None
    a = int(sta[:2]) * 60 + int(sta[3:5])
    b = int(std[:2]) * 60 + int(std[3:5])
    return (b - a) if b >= a else (b + 1440 - a)


def calc_mod(cia, sta, std):
    if cia == "LATAM":
        if not std:
            return "TST.D"
        m = int(std[:2]) * 60 + int(std[3:5])
        return "TST.N" if (m >= LAT_NOITE[0] or m <= LAT_NOITE[1]) else "TST.D"
    s = solo_min(sta, std)
    if s is None:
        return "TST"
    return "TST" if s < LIM_PNT else "PNT"


def conv_equip(eq):
    """equipamento da fonte -> (tipo, código)."""
    if eq is None:
        return None, None
    e = str(eq).strip().upper()
    e = re.sub(r"\.0$", "", e)
    tipo = EQUIP_TIPO.get(e)
    if tipo is None:
        tipo = EQUIP_TIPO.get(e.lstrip("0")) or e
    cod = TIPO_COD.get(tipo) or TIPO_COD.get(str(tipo).upper()) or ""
    return tipo, (f"{int(cod):02d}" if str(cod).isdigit() else cod)


def norm(s):
    s = unicodedata.normalize("NFKD", str(s))
    return "".join(c for c in s if not unicodedata.combining(c)).strip().upper()


# ───────────────────────── leitores ─────────────────────────
def ler_gol(arq):
    """GOL: CSV/Excel já com chegada e saída na mesma linha."""
    nome = arq.name.lower()
    if nome.endswith((".xlsx", ".xlsm", ".xls")):
        df = pd.read_excel(arq, dtype=str)
    else:
        raw = arq.getvalue().decode("utf-8", errors="replace")
        sep = ";" if raw.count(";") > raw.count(",") else ","
        df = pd.read_csv(io.StringIO(raw), sep=sep, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    col = {norm(c): c for c in df.columns}
    need = ["AIRPORT", "DAY", "INBOUND FLT NUM", "INBOUND STA", "OUTBOUND STD", "EQUIP"]
    faltando = [n for n in need if n not in col]
    if faltando:
        raise ValueError(f"GOL — colunas não encontradas: {', '.join(faltando)}")
    out = []
    for _, r in df.iterrows():
        base = str(r[col["AIRPORT"]]).strip().upper()
        dt = data_br(r[col["DAY"]])
        if not base or not dt:
            continue
        out.append({
            "Base": base, "Nome": "GOL",
            "Data": dt,
            "Voo": str(r[col["INBOUND FLT NUM"]]).strip(),
            "Hora Chegada": hhmm(r[col["INBOUND STA"]]),
            "Hora Saída": hhmm(r[col["OUTBOUND STD"]]),
            "equip": r[col["EQUIP"]],
        })
    return out


def ler_azul(arq):
    """AZUL: aba 'Ground time'; a base é o destino da chegada."""
    xl = pd.ExcelFile(arq)
    aba = next((s for s in xl.sheet_names if "GROUND" in norm(s)), xl.sheet_names[0])
    df = xl.parse(aba, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    col = {norm(c): c for c in df.columns}

    def acha(*alvos):
        for a in alvos:
            if a in col:
                return col[a]
        return None

    c_base = acha("INBOUND ARVL STA", "INBOUND ARVL")
    c_day = acha("DAY", "DATE")
    c_flt = acha("INBOUND FLT NUM", "INBOUND FLT NUMBER")
    c_sta = acha("INBOUND ARVL TIME", "INBOUND ARVL")
    c_std = acha("OUTBOUND DEPT TIME", "OUTBOUND DEPT")
    c_eq = acha("OUTBOUND EQUIP", "EQUIP")
    if not all([c_base, c_day, c_sta, c_std]):
        raise ValueError("AZUL — não encontrei as colunas da aba Ground time")

    out = []
    for _, r in df.iterrows():
        base = str(r[c_base]).strip().upper()
        dt = data_br(r[c_day])
        if not base or base == "NAN" or not dt:
            continue
        out.append({
            "Base": base, "Nome": "AZUL", "Data": dt,
            "Voo": str(r[c_flt]).strip() if c_flt else "",
            "Hora Chegada": hhmm(r[c_sta]),
            "Hora Saída": hhmm(r[c_std]),
            "equip": r[c_eq] if c_eq else None,
        })
    return out


def ler_latam(arq):
    """LATAM: uma linha por perna. Pareia pela operação seguinte da mesma aeronave."""
    xl = pd.ExcelFile(arq)
    aba = next((s for s in xl.sheet_names if "CONSOLIDA" in norm(s)), xl.sheet_names[0])
    df = xl.parse(aba, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    col = {norm(c): c for c in df.columns}
    need = {"ID AVION": None, "N°OPERACION POR AVION": None, "VUELO": None,
            "FSL": None, "ORI": None, "STD": None, "FLL": None, "DES": None, "STA": None}
    for k in list(need):
        alt = k.replace("°", "").replace("Ó", "O")
        need[k] = col.get(k) or col.get(alt) or next(
            (v for c, v in col.items() if c.startswith(k[:8])), None)
    if not need["ID AVION"] or not need["DES"]:
        raise ValueError("LATAM — não encontrei 'Id Avión' / 'DES' na aba consolidada")

    cols = list(df.columns)
    P = {k: cols.index(v) for k, v in need.items() if v in cols}
    i_mat = cols.index(col["MAT"]) if "MAT" in col else None
    obrig = ["ID AVION", "N°OPERACION POR AVION", "VUELO", "FSL", "STD", "FLL", "DES", "STA"]
    faltando = [k for k in obrig if k not in P]
    if faltando:
        raise ValueError("LATAM — colunas ausentes: " + ", ".join(faltando))

    dest, orig = {}, {}
    for r in df.itertuples(index=False, name=None):
        try:
            ida = str(r[P["ID AVION"]]).strip()
            op = int(float(r[P["N°OPERACION POR AVION"]]))
        except (TypeError, ValueError):
            continue
        fll, fsl = para_date(r[P["FLL"]]), para_date(r[P["FSL"]])
        if fll:
            dest[(ida, op + 1, fll)] = r      # chegada aponta para a operação seguinte
        if fsl:
            orig[(ida, op, fsl)] = r

    out = []
    for (ida, opn, dia), rd in dest.items():
        base = str(rd[P["DES"]]).strip().upper()
        if not base or base == "NAN":
            continue
        ro = orig.get((ida, opn, dia)) or orig.get((ida, opn, dia + timedelta(days=1)))
        if ro is None:
            continue
        out.append({
            "Base": base, "Nome": "LATAM", "Data": dia.strftime("%d/%m/%Y"),
            "Voo": str(rd[P["VUELO"]]).strip(),
            "Hora Chegada": hhmm(rd[P["STA"]]),
            "Hora Saída": hhmm(ro[P["STD"]]),
            "equip": rd[i_mat] if i_mat is not None else None,
        })
    return out


def ler_manual(arq, base_hint=""):
    """Manual: aba 'MALHA MES' com grade — uma coluna por dia, marcada com x."""
    xl = pd.ExcelFile(arq)
    aba = next((s for s in xl.sheet_names if "MALHA" in norm(s)), None)
    if aba is None:
        raise ValueError(f"{arq.name} — não achei a aba 'MALHA MES'")
    df = xl.parse(aba, header=None, dtype=object)

    # linha do cabeçalho: onde aparece CLIENTE
    lin_cab = None
    for i in range(min(30, len(df))):
        vals = [norm(v) for v in df.iloc[i].tolist() if v is not None]
        if "CLIENTE" in vals and any("VOO" in v for v in vals):
            lin_cab = i
            break
    if lin_cab is None:
        raise ValueError(f"{arq.name} — cabeçalho não encontrado")

    cab = [norm(v) if (v is not None and str(v) != "nan") else "" for v in df.iloc[lin_cab].tolist()]
    def idx(*alvos):
        for a in alvos:
            for j, c in enumerate(cab):
                if c == a or c.startswith(a):
                    return j
        return None
    c_cli, c_voo = idx("CLIENTE"), idx("# VOO", "VOO")
    c_ac, c_sta, c_std = idx("ACFT"), idx("STA"), idx("STD")

    # linha das datas: a primeira abaixo do cabeçalho com muitas datas
    lin_dt, datas = None, {}
    for i in range(lin_cab + 1, min(lin_cab + 5, len(df))):
        d = {j: para_date(v) for j, v in enumerate(df.iloc[i].tolist())
             if para_date(v) is not None}
        if len(d) >= 20:
            lin_dt, datas = i, d
            break
    if not datas:
        raise ValueError(f"{arq.name} — não achei a linha de datas")

    base = base_hint or re.split(r"[_\s\-.]", arq.name)[0].upper()[:3]
    out = []
    for i in range(lin_dt + 1, len(df)):
        row = df.iloc[i].tolist()
        cli = row[c_cli] if c_cli is not None else None
        if cli is None or str(cli).strip() in ("", "nan", "None"):
            continue
        sta, std = hhmm(row[c_sta]), hhmm(row[c_std])
        if not sta and not std:
            continue
        for j, dt in datas.items():
            v = row[j] if j < len(row) else None
            if v is None or str(v).strip().upper() not in ("X", "1"):
                continue
            out.append({
                "Base": base, "Nome": str(cli).strip().upper(),
                "Data": dt.strftime("%d/%m/%Y"),
                "Voo": str(row[c_voo]).strip() if c_voo is not None else "",
                "Hora Chegada": sta, "Hora Saída": std,
                "equip": row[c_ac] if c_ac is not None else None,
            })
    return out


# ───────────────────────── montagem do CSV ─────────────────────────
def montar(regs, cias_por_base, periodo, incluir_meli):
    linhas, descartes = [], defaultdict(int)
    for r in regs:
        base, cia = r["Base"], r["Nome"]
        if base not in cias_por_base:
            descartes["base fora da lista"] += 1
            continue
        if cia not in cias_por_base[base]:
            descartes[f"{cia} não atendida em {base}"] += 1
            continue
        d = para_date(r["Data"])
        if d is None or (d.year, d.month) not in periodo:
            descartes["fora do período"] += 1
            continue
        if not r["Hora Chegada"] or not r["Hora Saída"]:
            descartes["sem horário de chegada ou saída"] += 1
            continue
        eq = str(r.get("equip") or "").strip().upper()
        if eq in MELI and not incluir_meli.get(base):
            descartes[f"Gol Meli em {base}"] += 1
            continue
        tipo, cod = conv_equip(r.get("equip"))
        if not tipo:
            descartes["equipamento sem de-para"] += 1
            continue
        linhas.append({
            "Base": base,
            "Cód. Cliente": COD_CLIENTE.get(cia, "0000000000"),
            "Nome": cia,
            "Data": d.strftime("%d/%m/%Y"),
            "Mod": calc_mod(cia, r["Hora Chegada"], r["Hora Saída"]),
            "Tipo": tipo,
            "Voo": r["Voo"],
            "Hora Chegada": r["Hora Chegada"],
            "Hora Saída": r["Hora Saída"],
            "ICAO": ICAO.get(cia, "(vazio)"),
            "AERONAVE": cod,
        })
    linhas.sort(key=lambda x: (x["Base"], para_date(x["Data"]), x["Hora Chegada"]))
    return linhas, dict(descartes)


def gerar_csv(linhas):
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=CAB, lineterminator="\r\n")
    w.writeheader()
    w.writerows(linhas)
    return ("\ufeff" + buf.getvalue()).encode("utf-8")


# ───────────────────────── interface ─────────────────────────
st.title("✈️ Conversor de Malha · RP")
st.caption("Converte as malhas da GOL, AZUL, LATAM e as manuais para o padrão *Malha RP*.")

if "cias" not in st.session_state:
    st.session_state.cias = {b: list(v) for b, v in CIAS_POR_BASE.items()}
if "meli" not in st.session_state:
    st.session_state.meli = {b: False for b in CIAS_POR_BASE}

with st.sidebar:
    st.header("Período")
    hoje = date.today()
    ano = st.number_input("Ano", 2024, 2035, 2026)
    meses_nome = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"]
    sel = st.multiselect("Meses", options=list(range(1, 13)),
                         default=[9], format_func=lambda m: meses_nome[m-1])
    periodo = {(int(ano), m) for m in sel}

    st.divider()
    st.header("Bases e companhias")
    st.caption("Marque as companhias que a base atende em **voo**.")
    todas_cias = ["AZUL", "GOL", "LATAM", "SKY", "JETSMART", "ANDES"]
    for b in sorted(st.session_state.cias):
        with st.expander(f"{b} · {len(st.session_state.cias[b])} cia(s)"):
            novo = st.multiselect("Companhias", todas_cias,
                                  default=st.session_state.cias[b], key=f"c_{b}",
                                  label_visibility="collapsed")
            st.session_state.cias[b] = novo
            st.session_state.meli[b] = st.checkbox(
                "Atende Gol Meli (73C)", value=st.session_state.meli.get(b, False), key=f"m_{b}")

    st.divider()
    with st.form("nova_base"):
        st.caption("Cadastrar base nova")
        nb = st.text_input("Sigla IATA", max_chars=3, placeholder="ex: REC")
        nc = st.multiselect("Companhias", todas_cias, key="nc")
        if st.form_submit_button("Adicionar") and nb:
            st.session_state.cias[nb.strip().upper()] = nc
            st.session_state.meli[nb.strip().upper()] = False
            st.rerun()

c1, c2 = st.columns(2)
with c1:
    f_gol = st.file_uploader("**GOL** — um ou mais arquivos", type=["csv","xlsx","xlsm","xls"],
                             accept_multiple_files=True)
    f_azul = st.file_uploader("**AZUL** — aba *Ground time*", type=["xlsx","xlsm","xls"],
                              accept_multiple_files=True)
with c2:
    f_lat = st.file_uploader("**LATAM** — aba *Consolidado*", type=["xlsx","xlsm","xls"],
                             accept_multiple_files=True)
    f_man = st.file_uploader("**Manuais** — aba *MALHA MES*", type=["xlsx","xlsm","xls"],
                             accept_multiple_files=True)

if st.button("Gerar CSV", type="primary", use_container_width=True):
    if not periodo:
        st.error("Escolha ao menos um mês.")
        st.stop()

    regs, erros, lidos = [], [], {}
    for arqs, leitor, nome in [(f_gol, ler_gol, "GOL"), (f_azul, ler_azul, "AZUL"),
                               (f_lat, ler_latam, "LATAM"), (f_man, ler_manual, "Manual")]:
        for a in (arqs or []):
            try:
                r = leitor(a)
                regs += r
                lidos[f"{nome} · {a.name}"] = len(r)
            except Exception as e:
                erros.append(f"**{nome} · {a.name}** — {e}")

    if erros:
        for e in erros:
            st.error(e)
    if not regs:
        st.warning("Nenhum voo lido. Envie ao menos um arquivo.")
        st.stop()

    linhas, desc = montar(regs, st.session_state.cias, periodo, st.session_state.meli)

    with st.expander(f"Arquivos lidos ({sum(lidos.values())} registros)", expanded=False):
        for k, v in lidos.items():
            st.write(f"- {k}: **{v}**")
    if desc:
        with st.expander(f"Descartados ({sum(desc.values())})"):
            for k, v in sorted(desc.items(), key=lambda x: -x[1]):
                st.write(f"- {k}: **{v}**")

    if not linhas:
        st.error("Nada sobrou após os filtros. Confira o período e as companhias por base.")
        st.stop()

    st.success(f"**{len(linhas)} voos** no CSV.")
    df = pd.DataFrame(linhas)
    a, b, c = st.columns(3)
    a.metric("Voos", len(df))
    b.metric("Bases", df["Base"].nunique())
    c.metric("Companhias", df["Nome"].nunique())

    st.dataframe(
        df.pivot_table(index="Base", columns="Nome", values="Voo",
                       aggfunc="count", fill_value=0, margins=True, margins_name="Total"),
        use_container_width=True)
    st.dataframe(df.head(50), use_container_width=True, hide_index=True)

    mm = "-".join(f"{m:02d}" for m in sorted(sel))
    st.download_button("⬇️ Baixar CSV", gerar_csv(linhas),
                       file_name=f"Malha_RP_{ano}_{mm}.csv", mime="text/csv",
                       type="primary", use_container_width=True)
