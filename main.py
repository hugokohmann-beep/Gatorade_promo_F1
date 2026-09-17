"""
Gatorade Promo 2026 | Funil de Participação
Dashboard Streamlit — Data Intelligence Bakery

Rodar:  streamlit run app.py
"""
import re
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ------------------------------------------------------------------ config
st.set_page_config(page_title="Gatorade Promo 2026 | Funil", page_icon="🏎️", layout="wide")

ARQUIVO = "dados.xlsx"   # deve ficar na mesma pasta do app.py
SHEET_ABA = "Gatorade"   # nome da guia dentro do arquivo

LARANJA = "#E8491D"
PALETA = ["#E8491D", "#F28C28", "#1a1a1a", "#8c8c8c", "#FFC15E", "#7A2E12"]

METRICAS = {"usuarios": "Usuários", "cadastros": "Cadastros", "pincodes": "Pincodes", "opt_in": "Opt-in"}

COLUNAS = {
    "data_final": "data",
    "UTM Channel": "channel",
    "UTM Campaign": "campaign",
    "UTM Source": "source",
    "UTM Medium": "medium",
    "UTM Content": "content",
    "Ad ID": "ad_id",
    "Usuarios": "usuarios",
    "cadastros": "cadastros",
    "pincodes": "pincodes",
    "opt in": "opt_in",
}


# ------------------------------------------------------------------ estilo (template)
st.markdown(
    f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Roboto+Condensed:wght@400;700;900&display=swap');
html, body, [class*="css"], .stMarkdown, .stMetric {{ font-family: 'Roboto Condensed', sans-serif; }}
.block-container {{ max-width: 1150px; padding-top: 1.5rem; }}
.hero {{ border-radius: 14px; padding: 26px; color: #fff; min-height: 180px; position: relative;
        background: radial-gradient(circle at 70% 40%, #5a2410 0, #1b0d07 45%, #0a0a0a 100%); }}
.hero h1 {{ color:#fff; font-weight:900; font-size:46px; margin:30px 0 0; }}
.hero h1 span {{ color:#F26A21; }}
.hero .b {{ position:absolute; bottom:14px; font-weight:900; font-size:20px; }}
.sec {{ display:flex; align-items:center; justify-content:center; gap:18px; margin:34px 0 14px; }}
.sec:before, .sec:after {{ content:""; height:2px; width:140px; background:{LARANJA}; }}
.sec h2 {{ margin:0; color:{LARANJA}; font-weight:900; font-size:34px; }}
.box-title {{ text-align:center; color:{LARANJA}; font-weight:700; font-size:19px; margin-bottom:6px; }}
.ov {{ border:1.5px solid #f1c3b3; border-radius:12px; padding:16px; font-size:16px; height:100%; }}
.ov b {{ color:{LARANJA}; }}
.kpi {{ background:{LARANJA}; color:#fff; border-radius:10px; padding:14px; text-align:center; margin-bottom:12px; }}
.kpi small {{ font-size:12px; font-weight:700; text-transform:uppercase; opacity:.9; }}
.kpi div {{ font-size:30px; font-weight:900; }}
.kpi.g {{ background:#bdbdbd; color:#333; }}
.helmet {{ border-radius:16px; min-height:320px; display:flex; align-items:center; justify-content:center;
          background: radial-gradient(circle at 50% 40%, #333, #0b0b0b); color:#F26A21;
          font-weight:900; font-size:40px; text-align:center; letter-spacing:2px; }}
.footer {{ margin-top:30px; background:{LARANJA}; color:#fff; border-radius:10px; padding:16px 22px;
          display:flex; justify-content:space-between; align-items:center; font-weight:900; }}
.footer .c {{ text-align:center; font-weight:400; }}
</style>
""",
    unsafe_allow_html=True,
)


# ------------------------------------------------------------------ helpers
def fmt(n: float) -> str:
    return f"{n:,.0f}".replace(",", ".")


def pct(a: float, b: float) -> str:
    return f"{a / b * 100:.1f}%".replace(".", ",") if b else "–"


def num_br(serie: pd.Series) -> pd.Series:
    """Converte '1.088', '1.088,5', '1088' ou '1088.0' em número."""
    def conv(v):
        v = str(v).strip()
        if v in ("", "nan", "None"):
            return 0.0
        if "," in v:
            v = v.replace(".", "").replace(",", ".")
        elif re.fullmatch(r"\d{1,3}(\.\d{3})+", v):
            v = v.replace(".", "")
        try:
            return float(v)
        except ValueError:
            return 0.0
    return serie.map(conv)


def parse_data(serie: pd.Series) -> pd.Series:
    """
    Converte a coluna de data aceitando os formatos que o Excel costuma devolver:
    dd/mm/aaaa, aaaa-mm-dd (com ou sem hora) e o número de série do Excel (ex.: 45901).
    """
    s = serie.astype(str).str.strip().str.replace(r"\s+00:00:00$", "", regex=True)

    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")

    # 1) dd/mm/aaaa ou dd-mm-aaaa
    br = s.str.match(r"^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$", na=False)
    out[br] = pd.to_datetime(s[br], dayfirst=True, errors="coerce")

    # 2) aaaa-mm-dd (ISO, com ou sem hora)
    iso = ~br & s.str.match(r"^\d{4}-\d{2}-\d{2}", na=False)
    out[iso] = pd.to_datetime(s[iso], errors="coerce")

    # 3) número de série do Excel
    serial = ~br & ~iso & s.str.match(r"^\d{5}(\.\d+)?$", na=False)
    if serial.any():
        out[serial] = pd.to_datetime(
            pd.to_numeric(s[serial], errors="coerce"), unit="D", origin="1899-12-30", errors="coerce"
        )

    # 4) qualquer resto
    resto = out.isna()
    if resto.any():
        out[resto] = pd.to_datetime(s[resto], dayfirst=True, errors="coerce")

    return out.dt.normalize()


def secao(titulo: str):
    st.markdown(f'<div class="sec"><h2>{titulo}</h2></div>', unsafe_allow_html=True)


@st.cache_data
def preparar(df: pd.DataFrame) -> pd.DataFrame:
    df = df.rename(columns=lambda c: str(c).strip()).rename(columns=COLUNAS)
    faltando = [c for c in COLUNAS.values() if c not in df.columns]
    if faltando:
        raise ValueError(f"Colunas não encontradas: {', '.join(faltando)}")
    df["data"] = parse_data(df["data"])
    for c in METRICAS:
        df[c] = num_br(df[c])
    for c in ["channel", "campaign", "source", "medium", "content"]:
        df[c] = df[c].fillna("(vazio)").astype(str)
    df["ad_id"] = df["ad_id"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    invalidas = int(df["data"].isna().sum())
    if invalidas:
        st.warning(f"{invalidas} linha(s) foram ignoradas por não terem data válida na coluna 'data_final'.")
    return df.dropna(subset=["data"])


@st.cache_data(show_spinner="Lendo a planilha...")
def ler_planilha(caminho: str, aba: str, versao: float) -> pd.DataFrame:
    """Lê a aba do arquivo Excel. `versao` é a data de modificação, usada para invalidar o cache."""
    return pd.read_excel(caminho, sheet_name=aba, dtype=str).fillna("")


# ------------------------------------------------------------------ dados
caminho = Path(__file__).parent / ARQUIVO

with st.sidebar:
    st.header("Base de dados")
    st.write(f"Arquivo: **{ARQUIVO}** · aba **{SHEET_ABA}**")
    if caminho.exists():
        st.caption(f"Atualizado em {pd.Timestamp(caminho.stat().st_mtime, unit='s', tz='UTC').tz_convert('America/Sao_Paulo'):%d/%m/%Y %H:%M}")
    if st.button("🔄 Recarregar arquivo"):
        ler_planilha.clear()

if not caminho.exists():
    st.error(f"Arquivo '{ARQUIVO}' não encontrado. Coloque-o na mesma pasta do app.py: {caminho.parent}")
    st.stop()

try:
    bruto = ler_planilha(str(caminho), SHEET_ABA, caminho.stat().st_mtime)
    bruto = bruto.loc[:, ~bruto.columns.astype(str).str.startswith("Unnamed")]
    base = preparar(bruto)
except ValueError as e:
    st.error(f"Não foi possível ler a aba '{SHEET_ABA}': {e}")
    st.stop()
except Exception as e:
    st.error(f"Erro ao abrir o arquivo: {e}")
    st.stop()

# ------------------------------------------------------------------ header
st.markdown(
    """<div class="hero"><div style="opacity:.8">PepsiCo · F1 · Gatorade</div>
    <h1>Promo <span>Gatorade</span> 2026</h1><div style="opacity:.8">Funil de participação dos consumidores</div>
    <div class="b" style="left:26px">Bakery.</div><div class="b" style="right:26px">pepsico</div></div>""",
    unsafe_allow_html=True,
)

# ------------------------------------------------------------------ filtros
st.write("")
dmin, dmax = base["data"].min().date(), base["data"].max().date()
f1, f2, f3, f4 = st.columns([1.4, 1, 1, 1])
periodo = f1.date_input("Período", (dmin, dmax), min_value=dmin, max_value=dmax, format="DD/MM/YYYY")
canais = f2.multiselect("UTM Channel", sorted(base["channel"].unique()))
campanhas = f3.multiselect("UTM Campaign", sorted(base["campaign"].unique()))
sources = f4.multiselect("UTM Source", sorted(base["source"].unique()))

ini, fim = (periodo if isinstance(periodo, tuple) and len(periodo) == 2 else (dmin, dmax))
df = base[(base["data"].dt.date >= ini) & (base["data"].dt.date <= fim)]
if canais:
    df = df[df["channel"].isin(canais)]
if campanhas:
    df = df[df["campaign"].isin(campanhas)]
if sources:
    df = df[df["source"].isin(sources)]

if df.empty:
    st.warning("Nenhum dado para os filtros selecionados.")
    st.stop()

t = df[list(METRICAS)].sum()
dias = df["data"].dt.date.nunique()
por_canal = df.groupby("channel", as_index=False)[list(METRICAS)].sum().sort_values("cadastros", ascending=False)

# ------------------------------------------------------------------ overview
secao("overview")
lider = por_canal.iloc[0]
conv = por_canal[por_canal["usuarios"] >= 50].assign(tx=lambda d: d["cadastros"] / d["usuarios"])
melhor = conv.sort_values("tx", ascending=False).iloc[0] if not conv.empty else None

o1, o2, o3 = st.columns(3)
o1.markdown(
    f'<div class="ov"><b>{fmt(t.usuarios)}</b> usuários chegaram à promoção em <b>{dias}</b> dia(s), '
    f'gerando <b>{fmt(t.cadastros)}</b> cadastros ({pct(t.cadastros, t.usuarios)}).</div>',
    unsafe_allow_html=True,
)
txt_melhor = (
    f" Melhor taxa de cadastro: <b>{melhor.channel}</b> ({pct(melhor.cadastros, melhor.usuarios)})." if melhor is not None else ""
)
o2.markdown(
    f'<div class="ov"><b>{lider.channel}</b> lidera em cadastros ({pct(lider.cadastros, t.cadastros)} do total).{txt_melhor}</div>',
    unsafe_allow_html=True,
)
o3.markdown(
    f'<div class="ov"><b>{fmt(t.pincodes)}</b> pincodes cadastrados ({pct(t.pincodes, t.cadastros)} dos cadastros) '
    f'e <b>{fmt(t.opt_in)}</b> opt-ins ({pct(t.opt_in, t.cadastros)}).</div>',
    unsafe_allow_html=True,
)

# ------------------------------------------------------------------ big numbers
secao("big numbers")
kpis = [
    ("Usuários", fmt(t.usuarios), ""), ("Cadastros", fmt(t.cadastros), ""),
    ("Pincodes", fmt(t.pincodes), ""), ("Opt-in", fmt(t.opt_in), ""),
    ("Tx. Cadastro", pct(t.cadastros, t.usuarios), ""), ("Tx. Pincode", pct(t.pincodes, t.cadastros), ""),
    ("Tx. Opt-in", pct(t.opt_in, t.cadastros), ""),
    ("Período", f"{df['data'].min():%d/%m} a {df['data'].max():%d/%m}", "g"),
]
for linha in (kpis[:4], kpis[4:]):
    for col, (rot, val, cls) in zip(st.columns(4), linha):
        col.markdown(f'<div class="kpi {cls}"><small>{rot}</small><div>{val}</div></div>', unsafe_allow_html=True)

# ------------------------------------------------------------------ funil
secao("funil")
c_img, c_fun = st.columns([1, 1.4])
c_img.markdown('<div class="helmet">GATORADE<br>PIT STOP</div>', unsafe_allow_html=True)

etapas = ["Usuários", "Cadastros", "Opt-in", "Pincodes"]
valores = [t.usuarios, t.cadastros, t.opt_in, t.pincodes]
fig_fun = go.Figure(
    go.Funnel(
        y=etapas, x=valores,
        text=[f"{fmt(v)}<br>{s}" for v, s in zip(valores, [
            "100%", f"{pct(t.cadastros, t.usuarios)} dos usuários",
            f"{pct(t.opt_in, t.cadastros)} dos cadastros", f"{pct(t.pincodes, t.cadastros)} dos cadastros"])],
        textinfo="text", textfont=dict(color="white", size=15),
        marker=dict(color=[LARANJA, "#EE5F2A", "#F27A3A", "#F28C28"]),
        connector=dict(fillcolor="#f6d6c9"),
    )
)
fig_fun.update_layout(height=330, margin=dict(l=10, r=10, t=10, b=10), font=dict(family="Roboto Condensed", size=15))
c_fun.plotly_chart(fig_fun, use_container_width=True)
c_fun.markdown(
    f"<p style='text-align:center'>Conversão ponta a ponta (usuário → pincode): "
    f"<b style='color:{LARANJA}'>{pct(t.pincodes, t.usuarios)}</b></p>",
    unsafe_allow_html=True,
)

# ------------------------------------------------------------------ resultados
secao("resultados")

st.markdown('<div class="box-title">evolução por dia</div>', unsafe_allow_html=True)
met = st.radio("Métrica", list(METRICAS), format_func=METRICAS.get, horizontal=True, label_visibility="collapsed")
evo = df.assign(dia=df["data"].dt.normalize()).groupby(["dia", "channel"], as_index=False)[met].sum()
fig_evo = go.Figure()
for i, canal in enumerate(sorted(evo["channel"].unique())):
    d = evo[evo["channel"] == canal].sort_values("dia")
    fig_evo.add_bar(x=d["dia"], y=d[met], name=canal, marker_color=PALETA[i % len(PALETA)])
fig_evo.update_layout(
    barmode="stack", height=360, margin=dict(l=10, r=10, t=10, b=10),
    xaxis=dict(tickformat="%d/%m", dtick="D1"), legend=dict(orientation="h", y=-0.15),
    font=dict(family="Roboto Condensed"),
)
st.plotly_chart(fig_evo, use_container_width=True)

g1, g2 = st.columns(2)
with g1:
    st.markdown('<div class="box-title">resultados por veículo</div>', unsafe_allow_html=True)
    pc = por_canal.sort_values("cadastros")
    fig_ch = go.Figure()
    for (col, rot), cor in zip([("cadastros", "Cadastros"), ("opt_in", "Opt-in"), ("pincodes", "Pincodes")],
                               [LARANJA, "#F28C28", "#1a1a1a"]):
        fig_ch.add_bar(y=pc["channel"], x=pc[col], name=rot, orientation="h", marker_color=cor)
    fig_ch.update_layout(height=380, margin=dict(l=10, r=10, t=10, b=10), legend=dict(orientation="h", y=-0.12),
                         font=dict(family="Roboto Condensed"))
    st.plotly_chart(fig_ch, use_container_width=True)

with g2:
    st.markdown('<div class="box-title">detalhamento por veículo</div>', unsafe_allow_html=True)
    tab = por_canal.copy()
    total = pd.DataFrame([{"channel": "Total", **t.to_dict()}])
    tab = pd.concat([tab, total], ignore_index=True)
    tab["Tx Cad."] = (tab["cadastros"] / tab["usuarios"]).where(tab["usuarios"] > 0)
    tab["Tx Opt-in"] = (tab["opt_in"] / tab["cadastros"]).where(tab["cadastros"] > 0)
    tab["Tx Pin"] = (tab["pincodes"] / tab["cadastros"]).where(tab["cadastros"] > 0)
    tab = tab.rename(columns={"channel": "Canal", **METRICAS})[
        ["Canal", "Usuários", "Cadastros", "Tx Cad.", "Opt-in", "Tx Opt-in", "Pincodes", "Tx Pin"]]
    st.dataframe(
        tab, hide_index=True, use_container_width=True, height=380,
        column_config={
            **{c: st.column_config.NumberColumn(format="%d") for c in ["Usuários", "Cadastros", "Opt-in", "Pincodes"]},
            **{c: st.column_config.NumberColumn(format="percent") for c in ["Tx Cad.", "Tx Opt-in", "Tx Pin"]},
        },
    )

# ------------------------------------------------------------------ rodapé
st.markdown(
    """<div class="footer"><div>Bakery.</div>
    <div class="c">Desenvolvido por Data Intelligence Bakery<br><small>dúvidas e sugestões: bi@ampfy.com</small></div>
    <div>pepsico</div></div>""",
    unsafe_allow_html=True,
)