from __future__ import annotations

import hashlib
import html
import io
import os
import re
import unicodedata
from pathlib import Path
from typing import Iterable
from difflib import SequenceMatcher

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import load_workbook


# =========================================================
# CONFIGURAÇÃO VISUAL
# =========================================================
st.set_page_config(
    page_title="First Intelligence | Business Performance",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

NAVY = "#071B33"
NAVY_2 = "#0B2F55"
BLUE = "#1261A0"
CYAN = "#4F9AD1"
TEAL = "#2F78B7"
GREEN = "#245F96"
RED = "#B42318"
ORANGE = "#6C9FCB"
YELLOW = "#BFD8EC"
LIGHT = "#F3F6FA"
WHITE = "#FFFFFF"
GRAY = "#667085"
DARK = "#24364B"
BORDER = "#E4EAF1"
ANALYSIS_YEAR = 2026

LINES = ["MICROTECH", "LOCACAO", "VENDAS", "ENDOSCOPIA"]
LINE_LABELS = {
    "MICROTECH": "Microtech",
    "LOCACAO": "Locação",
    "VENDAS": "Vendas",
    "ENDOSCOPIA": "Endoscopia",
    "CONSOLIDADO": "Consolidado",
    "NAO CLASSIFICADA": "Não classificada",
}
LINE_COLORS = {
    "MICROTECH": "#4F9AD1",
    "LOCACAO": "#071B33",
    "VENDAS": "#1261A0",
    "ENDOSCOPIA": "#8AB8DC",
}

MANAGER_LINE_MAP = {
    "CELSO": "MICROTECH",
    "RENATO": "VENDAS",
    "AMAURI": "LOCACAO",
    "RONALDO": "ENDOSCOPIA",
}
LINE_MANAGER_MAP = {line: manager for manager, line in MANAGER_LINE_MAP.items()}

st.markdown(
    f"""
    <style>
    :root {{ --navy:{NAVY}; --blue:{BLUE}; --cyan:{CYAN}; --green:{GREEN}; --light:{LIGHT}; }}
    .stApp {{ background:{LIGHT}; }}
    .block-container {{ max-width:1540px; padding-top:1.15rem !important; padding-bottom:1.25rem; }}
    [data-testid="stSidebar"] {{ background:linear-gradient(180deg,#071B33 0%,#0B2F55 100%); border-right:0; }}
    [data-testid="stSidebar"] .block-container {{ padding-top:1rem; }}
    [data-testid="stSidebar"] label, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span {{ color:#F4F8FC !important; opacity:1 !important; }}
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] h4 {{ color:#FFFFFF !important; }}
    [data-testid="stSidebar"] [data-testid="stCaptionContainer"],
    [data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {{ color:#D8E7F3 !important; opacity:1 !important; line-height:1.45; }}
    [data-testid="stSidebar"] div[data-testid="stButton"] button {{
        background:rgba(255,255,255,.10); border:1px solid rgba(255,255,255,.20); color:#FFFFFF !important;
        min-height:2.55rem; box-shadow:none;
    }}
    [data-testid="stSidebar"] div[data-testid="stButton"] button:hover {{
        background:rgba(255,255,255,.17); border-color:rgba(255,255,255,.32);
    }}
    [data-testid="stSidebar"] div[data-testid="stButton"] button p {{ color:#FFFFFF !important; opacity:1 !important; }}
    [data-testid="stSidebar"] [data-baseweb="select"] > div,
    [data-testid="stSidebar"] input {{ background:#FFFFFF; color:#071B33; border-radius:10px; }}
    [data-testid="stSidebar"] [data-baseweb="select"] span,
    [data-testid="stSidebar"] input {{ color:#071B33 !important; }}
    [data-testid="stSidebar"] [role="radiogroup"] label {{ background:rgba(255,255,255,.045); margin-bottom:4px; border:1px solid rgba(255,255,255,.06); }}
    [data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) {{ background:rgba(79,154,209,.22); border-color:rgba(185,217,240,.42); }}
    [data-testid="stPlotlyChart"] {{
        background:linear-gradient(180deg,#FFFFFF 0%,#FCFEFF 100%); border:1px solid #D6E4EF;
        border-radius:18px; padding:4px 7px; box-shadow:0 10px 24px rgba(7,27,51,.06); overflow:hidden;
    }}
    [data-testid="stPlotlyChart"]:hover {{ box-shadow:0 20px 40px rgba(7,27,51,.10); transform:translateY(-1px); transition:.18s ease; }}
    div[data-testid="stDataFrame"] {{ border:1px solid {BORDER}; border-radius:16px; overflow:hidden; background:white; }}
    div[data-testid="stDownloadButton"] button, div[data-testid="stButton"] button {{ border-radius:11px; font-weight:750; }}

    .first-sidebar {{
        background:rgba(255,255,255,.08); color:white; border:1px solid rgba(255,255,255,.10); border-radius:16px;
        padding:16px 17px; margin:0 0 12px 0; box-shadow:none;
    }}
    .first-sidebar .brand {{ font-size:1.30rem; font-weight:950; letter-spacing:.16em; line-height:1; }}
    .first-sidebar .brand small {{ display:block; color:#B9D9F0; font-size:.58rem; letter-spacing:.30em; margin-top:6px; }}
    .first-sidebar p {{ margin:12px 0 0 0; opacity:.78; font-size:.74rem; line-height:1.4; }}

    .dashboard-hero {{
        position:relative; overflow:hidden; background:linear-gradient(120deg,{NAVY} 0%,#0E3762 58%,#1B5D97 120%);
        color:white; border-radius:22px; padding:23px 25px 20px; margin:4px 0 10px;
        box-shadow:0 18px 38px rgba(7,27,51,.18);
    }}
    .dashboard-hero::after {{ content:""; position:absolute; width:320px; height:320px; border-radius:50%; right:-130px; top:-185px; background:rgba(255,255,255,.055); }}
    .hero-grid {{ position:relative; z-index:2; display:flex; align-items:center; justify-content:space-between; gap:20px; }}
    .hero-brand {{ font-size:.72rem; font-weight:900; letter-spacing:.28em; color:#B9D9F0; text-transform:uppercase; }}
    .dashboard-hero h1 {{ font-size:clamp(1.46rem,2.15vw,1.98rem); line-height:1.16; margin:7px 0 1px 0; letter-spacing:-.025em; padding-top:1px; }}
    .dashboard-hero p {{ margin:0; opacity:.82; font-size:.88rem; }}
    .hero-chips {{ display:flex; flex-wrap:wrap; gap:7px; justify-content:flex-end; }}
    .hero-chip {{ background:rgba(255,255,255,.11); border:1px solid rgba(255,255,255,.20); border-radius:999px; padding:7px 11px; font-size:.71rem; font-weight:800; white-space:nowrap; }}

    .kpi-card {{
        background:linear-gradient(145deg,#FFFFFF 0%,#FBFDFF 100%); border:1px solid #DDE7F0; border-radius:19px; padding:13px 15px 12px 15px;
        min-height:118px; height:100%; position:relative; overflow:hidden;
        box-shadow:0 8px 24px rgba(7,27,51,.055); display:flex; flex-direction:column; justify-content:space-between;
    }}
    .kpi-card::after {{ content:""; position:absolute; width:70px; height:70px; border-radius:50%; right:-28px; top:-30px; background:var(--tone,#00A7B5); opacity:.10; }}
    .kpi-top {{ display:flex; align-items:center; gap:9px; }}
    .kpi-dot {{ width:9px; height:9px; border-radius:50%; background:var(--tone,#00A7B5); box-shadow:0 0 0 5px color-mix(in srgb, var(--tone,#00A7B5) 13%, transparent); }}
    .kpi-label {{ color:{GRAY}; font-size:.70rem; font-weight:850; letter-spacing:.055em; text-transform:uppercase; line-height:1.3; }}
    .kpi-value {{ color:{NAVY}; font-size:clamp(1.08rem,1.45vw,1.55rem); font-weight:950; line-height:1.07; margin:11px 0 0 0; letter-spacing:-.035em; white-space:nowrap; }}
    .kpi-note {{ color:{GRAY}; font-size:.72rem; line-height:1.35; margin-top:10px; overflow-wrap:anywhere; }}
    .kpi-delta {{ display:inline-flex; width:max-content; margin-top:8px; border-radius:999px; padding:3px 7px; font-size:.67rem; font-weight:850; background:#EAF3FA; color:{BLUE}; }}
    .kpi-delta.bad {{ background:#E8F2FA; color:{NAVY_2}; }}

    .section-head {{ display:flex; align-items:center; justify-content:space-between; gap:12px; margin:14px 0 8px 0; padding:0 0 7px 2px; border-bottom:1px solid #DDE7F0; }}
    .section-head h3 {{ color:{NAVY}; font-size:1.13rem; font-weight:900; margin:0; letter-spacing:-.015em; }}
    .section-head p {{ color:{GRAY}; font-size:.78rem; margin:4px 0 0 0; display:none; }}
    .section-badge {{ display:inline-flex; align-items:center; border-radius:999px; padding:5px 9px; background:#EAF3FA; color:{BLUE}; font-size:.67rem; font-weight:900; text-transform:uppercase; letter-spacing:.04em; white-space:nowrap; }}

    .insight-grid {{ display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:8px; margin-top:2px; }}
    .insight-card {{ background:white; border:1px solid {BORDER}; border-left:4px solid {BLUE}; border-radius:12px; padding:10px 12px; box-shadow:0 3px 10px rgba(7,27,51,.03); }}
    .insight-card.good {{ border-left-color:{TEAL}; }}
    .insight-card.critical {{ border-left-color:{CYAN}; }}
    .insight-card strong {{ color:{NAVY}; font-size:.84rem; display:block; margin-bottom:4px; }}
    .insight-card span {{ color:#475467; font-size:.76rem; line-height:1.42; }}

    .scope-note {{ background:#EDF4FA; border:1px solid #C9DDED; border-radius:13px; padding:11px 13px; color:#234F74; font-size:.77rem; line-height:1.45; }}
    .warning-note {{ background:#F2F7FB; border:1px solid #CCDDEC; border-radius:13px; padding:11px 13px; color:#315A7B; font-size:.77rem; line-height:1.45; }}
    .secure-note {{ background:rgba(255,255,255,.10); border:1px solid rgba(255,255,255,.14); border-radius:12px; padding:10px 12px; color:#EAF4FB; font-size:.73rem; line-height:1.4; }}
    .login-brand {{ text-align:center; margin:3.8rem auto 1.2rem auto; }}
    .login-brand .logo {{ color:#071B33; font-size:1.55rem; font-weight:950; letter-spacing:.15em; }}
    .login-brand .logo span {{ color:#2F78B7; }}
    .login-brand p {{ color:#667085; margin:.45rem 0 0; font-size:.82rem; }}
    div[data-testid="stForm"] {{ background:#FFFFFF; border:1px solid #DCE5EE; border-radius:20px; padding:1.15rem 1.2rem 1.25rem; box-shadow:0 18px 50px rgba(7,27,51,.10); }}
    div[data-testid="stForm"] button {{ background:linear-gradient(90deg,#071B33,#1261A0); color:#FFFFFF; border:0; min-height:2.8rem; }}
    .user-pill {{ display:flex; align-items:center; justify-content:space-between; gap:10px; background:linear-gradient(180deg,#F9FBFD 0%,#F2F6FA 100%); border:1px solid {BORDER}; border-radius:13px; padding:8px 10px; margin:6px 0 8px; box-shadow:0 5px 14px rgba(7,27,51,.045); }}
    .user-pill .left {{ display:flex; align-items:center; gap:10px; }}
    .user-pill .avatar {{ width:11px; height:11px; border-radius:50%; background:{BLUE}; box-shadow:0 0 0 5px rgba(18,97,160,.12); }}
    .user-pill b {{ color:{NAVY}; font-size:.82rem; display:block; }}
    .user-pill small {{ color:{GRAY}; font-size:.67rem; display:block; margin-top:2px; }}
    .user-pill span {{ color:{BLUE}; font-size:.69rem; font-weight:800; }}

    [data-testid="stSidebar"] [role="radiogroup"] label {{ border-radius:10px; padding:3px 4px; }}
    [data-testid="stSidebar"] hr {{ margin:.55rem 0; }}
    [data-testid="stTabs"] button {{ font-weight:800; color:#475467; }}
    [data-testid="stTabs"] button[aria-selected="true"] {{ color:{NAVY}; border-bottom-color:{BLUE}; }}

    .clean-table-wrap {{
        width:100%; overflow:auto; background:#FFFFFF; border:1px solid #DCE7F0;
        border-radius:17px; box-shadow:0 8px 22px rgba(7,27,51,.045);
    }}
    table.clean-table {{ width:100%; border-collapse:separate; border-spacing:0; font-size:.76rem; color:#344054; }}
    table.clean-table thead th {{
        position:sticky; top:0; z-index:2; background:linear-gradient(180deg,#0B2F55 0%,#123E68 100%);
        color:#FFFFFF; text-align:left; font-weight:850; padding:9px 11px; white-space:nowrap; border:0;
    }}
    table.clean-table tbody td {{ padding:8px 11px; border:0; white-space:nowrap; background:#FFFFFF; }}
    table.clean-table tbody tr:nth-child(even) td {{ background:#F7FAFD; }}
    table.clean-table tbody tr:hover td {{ background:#EDF5FB; }}
    .clean-table-note {{ color:#667085; font-size:.69rem; margin:4px 2px 0; }}
    div[data-testid="stVerticalBlock"] {{ gap:.72rem; }}
    div[data-testid="stHorizontalBlock"] {{ gap:.85rem; }}
    div[data-testid="stElementContainer"] {{ margin-bottom:0; }}
    [data-testid="stExpander"] {{ margin:2px 0 4px; }}

    @media(max-width:900px) {{
      .block-container {{ padding-left:.75rem; padding-right:.75rem; }}
      .dashboard-hero {{ padding:18px; }} .hero-grid {{ align-items:flex-start; flex-direction:column; }} .hero-chips {{ justify-content:flex-start; }}
      .kpi-card {{ min-height:110px; }} .insight-grid {{ grid-template-columns:1fr; }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# FUNÇÕES GERAIS
# =========================================================
def norm(value: object) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    text = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"\s+", " ", text).strip().upper()
    return text


def client_key(value: object) -> str:
    text = norm(value)
    text = re.sub(r"[^A-Z0-9 ]", " ", text)
    text = re.sub(r"\b(LTDA|EPP|EIRELI|ME|S A|SA|ASSOCIACAO|FUNDACAO)\b", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def seller_key(value: object) -> str:
    """Normaliza nomes de vendedores/representantes para cruzamento com a aba Metas."""
    text = norm(value)
    text = re.sub(r"[^A-Z0-9 ]", " ", text)
    text = re.sub(r"\b(LTDA|LT|EPP|EIRELI|ME|S A|SA)\b", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def document_key(value: object) -> str:
    """Normaliza NF/documento para cruzamento entre BASE BI e Centro de Custos."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    digits = re.sub(r"\D", "", text)
    return digits.lstrip("0") if digits else ""


def product_key(value: object) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    if isinstance(value, (int, np.integer)):
        text = str(int(value))
    elif isinstance(value, (float, np.floating)) and float(value).is_integer():
        text = str(int(value))
    else:
        text = str(value).strip()
    return re.sub(r"[^A-Z0-9]", "", norm(text))


def product_base_key(value: object) -> str:
    key = product_key(value)
    # Sufixos usados para revenda/teste não alteram o item comercial de referência.
    for suffix in ("RV1", "RV2", "RV", "TC", "AT"):
        if key.endswith(suffix) and len(key) > len(suffix) + 2:
            return key[:-len(suffix)]
    return key


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    seen: dict[str, int] = {}
    cols: list[str] = []
    for col in df.columns:
        name = str(col).strip()
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        cols.append(name)
    df.columns = cols
    return df


def to_number(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0.0)
    txt = series.astype(str).str.strip().str.replace(r"R\$\s*", "", regex=True).str.replace(" ", "", regex=False)
    both = txt.str.contains(",", na=False) & txt.str.contains(r"\.", na=False)
    txt.loc[both] = txt.loc[both].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    only_comma = txt.str.contains(",", na=False) & ~txt.str.contains(r"\.", na=False)
    txt.loc[only_comma] = txt.loc[only_comma].str.replace(",", ".", regex=False)
    return pd.to_numeric(txt, errors="coerce").fillna(0.0)


def brl(value: float) -> str:
    value = float(value or 0)
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def compact_money(value: float) -> str:
    value = float(value or 0)
    av = abs(value)
    if av >= 1_000_000_000:
        txt = f"{value / 1_000_000_000:,.1f} bi"
    elif av >= 1_000_000:
        txt = f"{value / 1_000_000:,.1f} mi"
    elif av >= 1_000:
        txt = f"{value / 1_000:,.1f} mil"
    else:
        txt = f"{value:,.0f}"
    return txt.replace(",", "X").replace(".", ",").replace("X", ".")


def pct(value: float) -> str:
    number = float(value or 0)
    if 0 < abs(number) < 0.0005:
        return "<0,1%" if number > 0 else ">-0,1%"
    return f"{number * 100:,.1f}%".replace(",", "X").replace(".", ",").replace("X", ".")


def safe_div(a: float, b: float) -> float:
    return float(a / b) if b not in (0, None) and not pd.isna(b) else 0.0


def filter_analysis_year(df: pd.DataFrame, year: int = ANALYSIS_YEAR) -> pd.DataFrame:
    """Mantém somente registros do ano-base na coluna mensal padronizada."""
    if df is None or df.empty or "_MES" not in df.columns:
        return df.copy() if isinstance(df, pd.DataFrame) else df
    months = df["_MES"]
    try:
        years = months.dt.year
    except Exception:
        years = months.map(lambda value: value.year if isinstance(value, pd.Period) else pd.to_datetime(value, errors="coerce").year)
    return df.loc[years.eq(year)].copy()


def token_similarity(a: str, b: str) -> float:
    a_tokens = " ".join(sorted(set(norm(a).split())))
    b_tokens = " ".join(sorted(set(norm(b).split())))
    return SequenceMatcher(None, a_tokens, b_tokens).ratio() * 100


def best_match(query: str, choices: list[str], cutoff: float = 88.0) -> tuple[str | None, float]:
    best_choice = None
    best_score = 0.0
    for choice in choices:
        score = token_similarity(query, choice)
        if score > best_score:
            best_choice, best_score = choice, score
    return (best_choice, best_score) if best_choice is not None and best_score >= cutoff else (None, 0.0)


def to_datetime_mixed(series: pd.Series) -> pd.Series:
    out = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    numeric = pd.to_numeric(series, errors="coerce")
    mask_num = numeric.notna() & numeric.between(20000, 80000)
    if mask_num.any():
        out.loc[mask_num] = pd.to_datetime(numeric.loc[mask_num], unit="D", origin="1899-12-30", errors="coerce")
    if (~mask_num).any():
        out.loc[~mask_num] = pd.to_datetime(series.loc[~mask_num], errors="coerce", dayfirst=True)
    return out


def month_label(period: pd.Period) -> str:
    names = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
    return f"{names[period.month - 1]}/{str(period.year)[2:]}"


def line_label(line: str) -> str:
    return LINE_LABELS.get(norm(line), str(line).title())


def cost_center_label(value: object) -> str:
    """Padroniza a exibição dos centros de custos oficiais sem alterar a base."""
    key = norm(value)
    if key in LINES:
        return line_label(key)
    if value is None or key == "":
        return "Não informado"
    return str(value).strip()


def rateio_line_key(value: object) -> str:
    """Converte o Centro de Custos Rateado para uma das linhas oficiais."""
    key = norm(value)
    if key in LINES:
        return key
    if key in MANAGER_LINE_MAP:
        return MANAGER_LINE_MAP[key]
    if "MICROTECH" in key or "MICRO TECH" in key:
        return "MICROTECH"
    if "LOCAC" in key:
        return "LOCACAO"
    if "ENDOSC" in key:
        return "ENDOSCOPIA"
    if "VENDA" in key:
        return "VENDAS"
    return "NAO CLASSIFICADA"



def card(label: str, value: str, note: str = "", tone: str = CYAN, delta: str = "", bad: bool = False) -> None:
    delta_html = f"<div class='kpi-delta {'bad' if bad else ''}'>{delta}</div>" if delta else ""
    st.markdown(
        f"""
        <div class="kpi-card" style="--tone:{tone}">
          <div>
            <div class="kpi-top"><span class="kpi-dot"></span><div class="kpi-label">{label}</div></div>
            <div class="kpi-value">{value}</div>
            {delta_html}
          </div>
          <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section_header(title: str, subtitle: str = "", badge: str = "") -> None:
    badge_html = f"<span class='section-badge'>{badge}</span>" if badge else ""
    st.markdown(
        f"""
        <div class="section-head">
          <div><h3>{title}</h3></div>{badge_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def insight_cards(items: list[dict[str, str]]) -> None:
    if not items:
        return
    html = ["<div class='insight-grid'>"]
    for item in items:
        cls = item.get("class", "")
        html.append(
            f"<div class='insight-card {cls}'><strong>{item['title']}</strong><span>{item['text']}</span></div>"
        )
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)


def plot_layout(fig: go.Figure, height: int = 350, legend_bottom: bool = True) -> go.Figure:
    legend = (
        dict(
            orientation="h", yanchor="top", y=-0.09, xanchor="left", x=0,
            bgcolor="rgba(239,246,252,.85)", bordercolor="#D9E6F0", borderwidth=1,
            font=dict(size=10, color="#365A78"), itemclick=False, itemdoubleclick=False,
        )
        if legend_bottom
        else dict(
            orientation="v", yanchor="top", y=1, xanchor="left", x=1.01,
            bgcolor="rgba(239,246,252,.85)", bordercolor="#D9E6F0", borderwidth=1,
            font=dict(size=10, color="#365A78"), itemclick=False, itemdoubleclick=False,
        )
    )
    fig.update_layout(
        height=height,
        margin=dict(l=16, r=58, t=78, b=50 if legend_bottom else 22),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#FBFDFF",
        font=dict(family="Arial", color="#344054", size=12),
        legend=legend,
        title=dict(font=dict(color=NAVY, size=16, family="Arial Black"), x=.018, xanchor="left", y=.975),
        hoverlabel=dict(bgcolor="white", bordercolor="#C9DDED", font=dict(color=NAVY, size=11), namelength=-1),
        hovermode="closest",
        separators=",.",
        uniformtext_minsize=9,
        uniformtext_mode="show",
        bargap=.28,
        bargroupgap=.10,
    )
    fig.update_xaxes(
        title_text="", showgrid=False, automargin=True, showline=False,
        tickfont=dict(color="#61758A", size=11), ticks="",
    )
    fig.update_yaxes(
        title_text="", showgrid=False, zeroline=False, automargin=True,
        showline=False, tickfont=dict(color="#61758A", size=11), ticks="",
    )
    fig.update_traces(
        selector=dict(type="bar"), opacity=.97,
        marker_line_color="rgba(7,27,51,.10)", marker_line_width=.7,
        textfont=dict(size=10, family="Arial"),
    )
    fig.update_traces(
        selector=dict(type="scatter"),
        marker=dict(size=8, line=dict(color="white", width=2)),
    )
    return fig


def hide_value_axis(fig: go.Figure, axis: str = "y") -> go.Figure:
    args = dict(showticklabels=False, showgrid=False, zeroline=False, ticks="", showline=False)
    if axis == "x":
        fig.update_xaxes(**args)
    else:
        fig.update_yaxes(**args)
    return fig


def add_point_labels(
    fig: go.Figure,
    x_values,
    y_values,
    labels,
    *,
    color: str = NAVY,
    alternate: bool = True,
) -> go.Figure:
    """Adiciona rótulos compactos, com deslocamento alternado e leitura limpa."""
    for idx, (x, y, label) in enumerate(zip(x_values, y_values, labels)):
        if pd.isna(y):
            continue
        y_num = float(y)
        if y_num < 0:
            yshift = -22
            xshift = 0
        elif alternate:
            yshift = 22 if idx % 2 == 0 else 34
            xshift = -4 if idx % 2 == 0 else 4
        else:
            yshift = 22
            xshift = 0
        fig.add_annotation(
            x=x, y=y_num, text=str(label), showarrow=False,
            yshift=yshift, xshift=xshift,
            bgcolor="rgba(255,255,255,.98)", bordercolor="#D6E3ED",
            borderwidth=1, borderpad=3, opacity=.98,
            font=dict(size=9, color=color, family="Arial"),
        )
    return fig


def short_label(value: object, limit: int = 34) -> str:
    text = str(value or "Não informado").strip()
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"


def clean_table(df: pd.DataFrame, height: int = 360, max_rows: int = 250) -> None:
    """Tabela executiva sem linhas de grade, com cabeçalho fixo e rolagem."""
    if df is None or df.empty:
        st.info("Não há registros para os filtros selecionados.")
        return
    view = df.head(max_rows).copy()
    view = view.where(pd.notna(view), "")
    html_table = view.to_html(index=False, escape=True, classes="clean-table", border=0)
    st.markdown(
        f"<div class='clean-table-wrap' style='max-height:{height}px'>{html_table}</div>",
        unsafe_allow_html=True,
    )
    if len(df) > max_rows:
        st.markdown(
            f"<div class='clean-table-note'>Exibindo os {max_rows} primeiros registros de {len(df):,}.</div>".replace(",", "."),
            unsafe_allow_html=True,
        )


def line_result_lollipop(df: pd.DataFrame, title: str) -> go.Figure:
    """Resultado por linha em barras verticais para leitura mais limpa de valores negativos."""
    p = df.sort_values("Resultado Direto de Caixa", ascending=False).copy()
    p["Cor"] = np.where(p["Resultado Direto de Caixa"] >= 0, BLUE, RED)
    p["Texto"] = p["Resultado Direto de Caixa"].map(compact_money)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=p["Linha"],
        y=p["Resultado Direto de Caixa"],
        marker=dict(color=p["Cor"], line=dict(color="rgba(7,27,51,.10)", width=0.8)),
        text=p["Texto"],
        textposition="outside",
        textfont=dict(size=10, color=NAVY),
        cliponaxis=False,
        customdata=p[["Linha"]],
        hovertemplate="%{customdata[0]}<br>Resultado: R$ %{y:,.2f}<extra></extra>",
        showlegend=False,
    ))
    fig.add_hline(y=0, line_color="#AFC4D6", line_width=1.4)
    fig.update_layout(title=title, showlegend=False, bargap=.42)
    hide_value_axis(fig, "y")
    fig.update_xaxes(showgrid=False, ticks="", tickfont=dict(color="#61758A", size=11))
    return fig


def line_revenue_cost_dumbbell(df: pd.DataFrame, title: str) -> go.Figure:
    """Receitas e despesas por linha em barras verticais agrupadas."""
    p = df.sort_values("Receitas Recebidas", ascending=False).copy()
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=p["Linha"],
        y=p["Receitas Recebidas"],
        name="Receitas",
        marker=dict(color=BLUE, line=dict(color="rgba(7,27,51,.10)", width=0.8)),
        text=p["Receitas Recebidas"].map(compact_money),
        textposition="outside",
        textfont=dict(size=10, color=NAVY),
        cliponaxis=False,
        hovertemplate="%{x}<br>Receitas: R$ %{y:,.2f}<extra></extra>",
    ))
    fig.add_trace(go.Bar(
        x=p["Linha"],
        y=p["Custos Diretos Pagos"],
        name="Despesas",
        marker=dict(color=RED, line=dict(color="rgba(7,27,51,.10)", width=0.8)),
        text=p["Custos Diretos Pagos"].map(compact_money),
        textposition="outside",
        textfont=dict(size=10, color=RED),
        cliponaxis=False,
        hovertemplate="%{x}<br>Despesas: R$ %{y:,.2f}<extra></extra>",
    ))
    fig.update_layout(title=title, barmode="group", bargap=.35, bargroupgap=.14)
    hide_value_axis(fig, "y")
    fig.update_xaxes(showgrid=False, ticks="", tickfont=dict(color="#61758A", size=11))
    return fig


def dataframe_download(df: pd.DataFrame, name: str) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=name[:31])
        workbook = writer.book
        worksheet = writer.sheets[name[:31]]
        header_fmt = workbook.add_format({"bold": True, "font_color": "white", "bg_color": NAVY})
        money_fmt = workbook.add_format({"num_format": 'R$ #,##0.00;[Red]-R$ #,##0.00'})
        pct_fmt = workbook.add_format({"num_format": "0.0%"})
        for col_num, value in enumerate(df.columns):
            worksheet.write(0, col_num, value, header_fmt)
            worksheet.set_column(col_num, col_num, min(max(len(str(value)) + 3, 13), 42))
        for idx, col in enumerate(df.columns):
            n = norm(col)
            if any(x in n for x in ["VALOR", "RECEITA", "CUSTO", "DESPESA", "RESULTADO", "FATURAMENTO", "INADIMPLENCIA"]):
                worksheet.set_column(idx, idx, 19, money_fmt)
            if any(x in n for x in ["MARGEM", "CONVERSAO", "COBERTURA", "PARTICIPACAO", "PERFORMANCE"]):
                worksheet.set_column(idx, idx, 15, pct_fmt)
    return output.getvalue()


# =========================================================
# LEITURA DAS BASES
# =========================================================
def first_existing(candidates: Iterable[str]) -> Path | None:
    here = Path(__file__).resolve().parent
    for name in candidates:
        p = here / name
        if p.exists():
            return p
    return None


def repository_files(kind: str) -> list[Path]:
    """Localiza fontes do repositório e evita que uma cópia antiga fique invisível."""
    here = Path(__file__).resolve().parent
    found: list[Path] = []
    for path in here.glob("*.xls*"):
        if path.name.startswith("~$"):
            continue
        name = norm(path.stem)
        if kind == "base" and "BASE" in name and "BI" in name and "REV2026" not in name:
            found.append(path)
        elif kind == "rev" and "REV2026" in name and "BASE" in name and "BI" in name:
            found.append(path)
    # Mantém nomes canônicos primeiro e remove duplicidades.
    preferred = {
        "base": ["BASE BI.xlsx", "BASE BI.xlsm", "BASE BI(1).xlsx", "base_bi.xlsx"],
        "rev": ["rev2026 Base bi.xlsx", "rev2026 Base bi.xlsm", "rev2026 Base bi(1).xlsx", "REV2026.xlsx"],
    }[kind]
    order = {name.casefold(): idx for idx, name in enumerate(preferred)}
    unique = {str(p.resolve()): p for p in found}
    return sorted(unique.values(), key=lambda p: (order.get(p.name.casefold(), 999), p.name.casefold()))


def preferred_repository_file(files: list[Path], canonical_name: str) -> Path | None:
    for path in files:
        if path.name.casefold() == canonical_name.casefold():
            return path
    return files[0] if files else None


def source_bytes(uploaded, default_path: Path | None) -> tuple[bytes | None, str]:
    if uploaded is not None:
        return uploaded.getvalue(), uploaded.name
    if default_path and default_path.exists():
        return default_path.read_bytes(), default_path.name
    return None, "Não localizada"


def file_fingerprint(file_bytes: bytes | None) -> str:
    if not file_bytes:
        return "sem arquivo"
    return hashlib.sha256(file_bytes).hexdigest()[:10].upper()


@st.cache_data(show_spinner=False, ttl=300)
def workbook_sheet_states(file_bytes: bytes) -> dict[str, str]:
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    return {ws.title: ws.sheet_state for ws in wb.worksheets}


@st.cache_data(show_spinner=False, ttl=300)
def read_excel_sheet(file_bytes: bytes, sheet_name: str, usecols=None) -> pd.DataFrame:
    return clean_columns(pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, usecols=usecols, engine="openpyxl"))


@st.cache_data(show_spinner=False, ttl=300)
def read_excel_raw(file_bytes: bytes, sheet_name: str, usecols=None) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, usecols=usecols, header=None, engine="openpyxl")


def require_col(df: pd.DataFrame, candidates: list[str], context: str) -> str:
    lookup = {norm(c): c for c in df.columns}
    for cand in candidates:
        if norm(cand) in lookup:
            return lookup[norm(cand)]
    raise ValueError(f"Não encontrei a coluna necessária em {context}: {', '.join(candidates)}")


def optional_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    lookup = {norm(c): c for c in df.columns}
    for cand in candidates:
        if norm(cand) in lookup:
            return lookup[norm(cand)]
    return None


def nature_key(value: object) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    text = str(value).strip().replace(",", ".")
    try:
        return f"{float(text):.10f}".rstrip("0").rstrip(".")
    except ValueError:
        return norm(text)


def classify_billing_line(row: pd.Series) -> str:
    """Classificação de contingência quando o gerente não pertence ao mapa oficial.

    O segmento tem prioridade sobre palavras encontradas na descrição do produto. Isso
    evita, por exemplo, que uma assistência técnica de Terapia Intensiva seja enviada
    para Endoscopia apenas porque o equipamento possui alguma referência endoscópica.
    """
    segment = norm(row.get("SEGMENTO", ""))
    nova = norm(row.get("NOVA", ""))
    fields = ["FORNECEDOR", "LINHA DE PRODUTO", "PRODUTO", "DESCRIÇÃO", "Endoscopia?"]
    text = " ".join(norm(row.get(c, "")) for c in fields)

    if "MICROTECH" in segment or "MICRO TECH" in segment or "MICROTECH" in text or "MICRO TECH" in text:
        return "MICROTECH"
    if "LOCACAO" in segment or "LOCACAO" in nova:
        return "LOCACAO"
    if "GASTROENDOSCOPIA" in segment or segment.startswith("ENDOSCOPIA"):
        return "ENDOSCOPIA"
    if "VENDAS" in segment or nova in {"VENDA", "SERVICO"}:
        return "VENDAS"
    if norm(row.get("Endoscopia?")) == "SIM" or "ENDOSCOPIA" in text or "GASTROENDOSCOPIA" in text:
        return "ENDOSCOPIA"
    return "VENDAS"


@st.cache_data(show_spinner=False, ttl=300)
def load_base_bi(file_bytes: bytes) -> dict[str, pd.DataFrame | dict[str, str]]:
    states = workbook_sheet_states(file_bytes)
    needed = ["BANCO DE DADOS FATURAMENTO", "Metas", "Metas Gerentes"]
    missing = [s for s in needed if s not in states]
    if missing:
        raise ValueError(f"Abas ausentes na BASE BI: {', '.join(missing)}")

    fat = read_excel_sheet(file_bytes, "BANCO DE DADOS FATURAMENTO")
    metas = read_excel_sheet(file_bytes, "Metas")
    metas_g = read_excel_sheet(file_bytes, "Metas Gerentes")
    users = read_excel_sheet(file_bytes, "Usuarios") if "Usuarios" in states else pd.DataFrame()

    month_col = require_col(fat, ["MÊS"], "BANCO DE DADOS FATURAMENTO")
    date_col = optional_col(fat, ["DT Emissao", "DT EMISSÃO", "DATA EMISSÃO"])
    value_col = require_col(fat, ["VALOR BRUTO", "VALOR ", "VALOR"], "BANCO DE DADOS FATURAMENTO")
    # Regra oficial da BASE BI: a competência comercial é definida pela coluna MÊS.
    # A DT Emissão é mantida para auditoria e usada somente como contingência quando
    # a competência estiver vazia. Isso preserva os lançamentos que a própria base
    # alocou em 2026, ainda que a emissão tenha ocorrido no fechamento de 2025.
    fat["_DATA"] = to_datetime_mixed(fat[date_col]) if date_col else pd.NaT
    official_month = to_datetime_mixed(fat[month_col]).dt.to_period("M")
    emission_month = fat["_DATA"].dt.to_period("M")
    fat["_MES"] = official_month.where(official_month.notna(), emission_month)
    fat["_ORIGEM_MES"] = np.where(official_month.notna(), "MÊS da BASE BI", "DT Emissão (contingência)")
    fat["_VALOR"] = to_number(fat[value_col])
    for col in [
        "GERENTE", "VENDEDOR / REPRESENTANTE", "VENDEDOR", "SEGMENTO", "EMPRESA", "NOVA",
        "NOME DO CLIENTE", "CLIENTE", "FORNECEDOR", "LINHA DE PRODUTO", "PRODUTO", "Nota Fiscal", "CATEGORIA",
        "DESCRIÇÃO", "Endoscopia?",
    ]:
        if col in fat.columns:
            fat[col] = fat[col].fillna("Não informado").astype(str).str.strip()
    # A visão de cada gestor deve refletir exclusivamente o gerente oficial da BASE BI.
    # O segmento/produto é usado somente quando o gerente estiver realmente ausente.
    # Gerentes informados fora do mapa oficial (ex.: Gabriel) permanecem no consolidado,
    # mas não são incorporados silenciosamente à linha Vendas ou a qualquer outro gestor.
    fallback_line = fat.apply(classify_billing_line, axis=1)
    fat["_LINHA_FALLBACK"] = fallback_line
    if "GERENTE" in fat.columns:
        fat["_GERENTE_N"] = fat["GERENTE"].map(norm)
        manager_line = fat["_GERENTE_N"].map(MANAGER_LINE_MAP)
        manager_missing = fat["_GERENTE_N"].isin(["", "NAO INFORMADO", "NAN", "NONE"])
        fat["_LINHA_GERENTE"] = manager_line.fillna("NAO CLASSIFICADA")
        fat["_LINHA"] = np.select(
            [manager_line.notna(), manager_missing],
            [manager_line, fallback_line],
            default="NAO CLASSIFICADA",
        )
        fat["_CRITERIO_LINHA"] = np.select(
            [manager_line.notna(), manager_missing],
            ["Gerente oficial", "Segmento / produto (gerente ausente)"],
            default="Gerente fora do mapa oficial",
        )
    else:
        fat["_GERENTE_N"] = "NAO INFORMADO"
        fat["_LINHA_GERENTE"] = "NAO CLASSIFICADA"
        fat["_LINHA"] = fallback_line
        fat["_CRITERIO_LINHA"] = "Segmento / produto (gerente ausente)"
    client_col = "NOME DO CLIENTE" if "NOME DO CLIENTE" in fat.columns else "CLIENTE"
    fat["_CLIENT_KEY"] = fat[client_col].map(client_key)
    fat["_TIPO_CAIXA"] = np.where(
        fat.get("NOVA", "").astype(str).map(norm).str.contains("LOCACAO") |
        fat.get("SEGMENTO", "").astype(str).map(norm).str.contains("LOCACAO"),
        "LOCACAO", "VENDA_SERVICO"
    )

    for df, label in [(metas, "Metas"), (metas_g, "Metas Gerentes")]:
        mcol = require_col(df, ["MÊS"], label)
        vcol = require_col(df, ["META MENSAL"], label)
        annual_col = optional_col(df, ["METAL ANUAL", "META ANUAL"])
        realized_col = optional_col(df, ["ATINGIMENTO", "REALIZADO", "FATURAMENTO REALIZADO"])
        df["_MES"] = pd.to_datetime(df[mcol], errors="coerce").dt.to_period("M")
        df["_META"] = to_number(df[vcol])
        df["_META_ANUAL"] = to_number(df[annual_col]) if annual_col else 0.0
        # Na BASE BI, a coluna ATINGIMENTO guarda o valor realizado mensal,
        # enquanto a coluna % contém o percentual de atingimento.
        df["_REALIZADO_META"] = to_number(df[realized_col]) if realized_col else 0.0
        for c in ["GERENTE", "VENDEDOR"]:
            if c in df.columns:
                df[c] = df[c].fillna("Não informado").astype(str).str.strip()

    # Metadados para a página de validação. Não alteram nenhum cálculo.
    valid_date = fat["_MES"].notna()
    in_analysis_year = valid_date & fat["_MES"].map(lambda value: value.year if isinstance(value, pd.Period) else np.nan).eq(ANALYSIS_YEAR)
    # Não eliminamos linhas repetidas: em locação, várias unidades podem ter dados
    # comerciais idênticos. A validação destaca apenas ausências e valores atípicos.
    zero_mask = fat["_VALOR"].eq(0)
    negative_mask = fat["_VALOR"].lt(0)
    note_col_audit = optional_col(fat, ["Nota Fiscal", "NOTA FISCAL", "NF"])
    missing_note_mask = (
        fat[note_col_audit].fillna("").astype(str).map(norm).isin(["", "NAO INFORMADO", "NAN"])
        if note_col_audit else pd.Series(True, index=fat.index)
    )
    emission_year = fat["_DATA"].dt.year if "_DATA" in fat.columns else pd.Series(np.nan, index=fat.index)
    allocated_from_other_year = in_analysis_year & emission_year.notna() & emission_year.ne(ANALYSIS_YEAR)
    month_fallback_mask = in_analysis_year & fat["_ORIGEM_MES"].eq("DT Emissão (contingência)")
    billing_audit = {
        "date_column": date_col or "Não disponível",
        "month_column": month_col,
        "value_column": value_col,
        "rows_total": int(len(fat)),
        "rows_without_month": int((~valid_date).sum()),
        "rows_2026": int(in_analysis_year.sum()),
        "value_2026": float(fat.loc[in_analysis_year, "_VALOR"].sum()),
        "rows_emission_outside_2026": int(allocated_from_other_year.sum()),
        "value_emission_outside_2026": float(fat.loc[allocated_from_other_year, "_VALOR"].sum()),
        "rows_month_fallback_2026": int(month_fallback_mask.sum()),
        "zero_value_rows_2026": int((zero_mask & in_analysis_year).sum()),
        "negative_value_rows_2026": int((negative_mask & in_analysis_year).sum()),
        "negative_value_2026": float(fat.loc[negative_mask & in_analysis_year, "_VALOR"].sum()),
        "missing_note_rows_2026": int((missing_note_mask & in_analysis_year).sum()),
    }

    return {
        "faturamento": fat, "metas": metas, "metas_gerentes": metas_g,
        "usuarios": users, "sheet_states": states, "billing_audit": billing_audit,
    }


@st.cache_data(show_spinner=False)
def load_price_table(file_bytes: bytes) -> dict[str, object]:
    states = workbook_sheet_states(file_bytes)
    if "Tabela_UF" not in states:
        raise ValueError("A tabela de preços precisa conter a aba Tabela_UF.")
    table = read_excel_sheet(file_bytes, "Tabela_UF")
    product_col = require_col(table, ["Produto", "PRODUTO", "Código", "CODIGO"], "Tabela_UF")
    type_col = require_col(table, ["TIPO_PRECO", "TIPO PRECO"], "Tabela_UF")
    table = table[table[type_col].astype(str).map(norm) == "VENDA DIRETA"].copy()
    table["_PRODUCT_KEY"] = table[product_col].map(product_key)
    table["_PRODUCT_BASE"] = table[product_col].map(product_base_key)
    description_col = optional_col(table, ["Descrição", "DESCRICAO", "DESCRIÇÃO"])
    if description_col:
        table["_PRICE_DESCRIPTION"] = table[description_col].fillna("").astype(str).str.strip()
    else:
        table["_PRICE_DESCRIPTION"] = ""
    update_col = optional_col(table, ["ATUALIZADO EM", "ATUALIZADO EM ", "DATA ATUALIZAÇÃO"] )
    table["_UPDATED"] = to_datetime_mixed(table[update_col]) if update_col else pd.NaT

    preferred = ["SP", "Consumidor Final", "4%", "RJ", "MG", "PR", "SC", "RS", "ES", "BA", "PE", "CE", "PA", "AM", "MT", "GO", "DF", "AC", "AL", "MA", "PB", "MS", "PI", "SE", "RN", "RR", "RO", "TO", "AP"]
    reference_cols = [c for c in preferred if c in table.columns]
    for col in reference_cols:
        table[col] = to_number(table[col])
    table = table[table["_PRODUCT_KEY"] != ""].sort_values("_UPDATED").drop_duplicates("_PRODUCT_KEY", keep="last")
    return {"table": table, "references": reference_cols, "sheet_states": states}


def _fallback_stock_line(row: pd.Series) -> str:
    text = " ".join(norm(row.get(c, "")) for c in ["Linha", "Descrição", "Descrição_Estoque", "Grupo", "Grupo_Estoque"])
    if "MICROTECH" in text or "MICRO TECH" in text:
        return "MICROTECH"
    if "ENDOSCOPIA" in text or "GASTROENDOSCOPIA" in text:
        return "ENDOSCOPIA"
    if "LOCACAO" in text:
        return "LOCACAO"
    return "NAO CLASSIFICADA"


def _map_stock_lines(df: pd.DataFrame, line_map: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.copy()
    product_col = require_col(df, ["Produto", "PRODUTO", "Código", "CODIGO"], "base de estoque")
    out = df.copy()
    out["_PRODUCT_KEY"] = out[product_col].map(product_key)
    out["_PRODUCT_BASE"] = out[product_col].map(product_base_key)
    out = out.merge(line_map[["_PRODUCT_BASE", "_LINHA_HIST"]], on="_PRODUCT_BASE", how="left")
    fallback = out.apply(_fallback_stock_line, axis=1)
    out["_LINHA_ESTOQUE"] = out["_LINHA_HIST"].where(out["_LINHA_HIST"].notna(), fallback)
    out["_MATCH_LINHA"] = np.where(out["_LINHA_HIST"].notna(), "Histórico de faturamento", np.where(fallback != "NAO CLASSIFICADA", "Linha do estoque", "Não classificada"))
    return out


@st.cache_data(show_spinner=False)
def load_stock_base(file_bytes: bytes, fat: pd.DataFrame) -> dict[str, object]:
    states = workbook_sheet_states(file_bytes)
    if "Radar Executivo" not in states:
        raise ValueError("A base de estoque precisa conter a aba Radar Executivo.")
    radar = read_excel_sheet(file_bytes, "Radar Executivo")
    capital = read_excel_sheet(file_bytes, "Capital Parado") if "Capital Parado" in states else pd.DataFrame()
    armz = read_excel_sheet(file_bytes, "ARMZ") if "ARMZ" in states else pd.DataFrame()

    product_col = optional_col(fat, ["PRODUTO", "ITEM", "CÓDIGO PRODUTO"])
    if product_col:
        mapping = fat[[product_col, "_LINHA", "_VALOR"]].copy()
        mapping["_PRODUCT_BASE"] = mapping[product_col].map(product_base_key)
        mapping = (mapping[mapping["_PRODUCT_BASE"] != ""]
                   .groupby(["_PRODUCT_BASE", "_LINHA"], as_index=False)["_VALOR"].sum()
                   .sort_values("_VALOR")
                   .drop_duplicates("_PRODUCT_BASE", keep="last")
                   .rename(columns={"_LINHA": "_LINHA_HIST"}))
    else:
        mapping = pd.DataFrame(columns=["_PRODUCT_BASE", "_LINHA_HIST"])

    radar = _map_stock_lines(radar, mapping)
    capital = _map_stock_lines(capital, mapping) if not capital.empty else capital
    armz = _map_stock_lines(armz, mapping) if not armz.empty else armz

    numeric_cols = [
        "Estoque_Disponível", "Pedido_Aberto_Qtd", "Estoque_Projetado", "Valor_Estoque", "Qtd_30d", "Qtd_180d",
        "Forecast_30d", "Estoque_Segurança", "Cobertura_Dias", "Cobertura_Meses", "Cobertura_Projetada_Dias",
        "Excesso_Estoque", "Excesso_R$", "Comprar_Qtd", "Comprar_R$", "Comprar_Líquido_Qtd", "Comprar_Líquido_R$",
        "Score_Oportunidade", "Estoque_Total", "Dias_Sem_Giro", "Valor_ARMZ", "Estoque_ARMZ", "Disponível_ARMZ"
    ]
    for df in [radar, capital, armz]:
        if df.empty:
            continue
        for col in numeric_cols:
            if col in df.columns:
                df[col] = to_number(df[col]).replace([np.inf, -np.inf], np.nan)
        for col in ["Status", "Ação", "Classe ABC Receita", "Descrição", "Descrição_Estoque", "Linha", "ARMZ"]:
            if col in df.columns:
                df[col] = df[col].fillna("Não informado").astype(str).str.strip()
    return {"radar": radar, "capital": capital, "armz": armz, "sheet_states": states}


def build_price_analysis(fat_scope: pd.DataFrame, price_table: pd.DataFrame, reference: str, embedded_margin: float) -> pd.DataFrame:
    if fat_scope.empty or price_table.empty or reference not in price_table.columns:
        return pd.DataFrame()
    product_col = optional_col(fat_scope, ["PRODUTO", "ITEM", "CÓDIGO PRODUTO"] )
    if product_col is None:
        return pd.DataFrame()
    direct = fat_scope.copy()
    rental_mask = direct["_LINHA"].eq("LOCACAO")
    for col in ["NOVA", "SEGMENTO", "CATEGORIA"]:
        if col in direct.columns:
            rental_mask = rental_mask | direct[col].astype(str).map(norm).str.contains("LOCACAO")
    direct = direct[~rental_mask].copy()
    if direct.empty:
        return direct

    qty_col = optional_col(direct, ["QUANTIDADE", "QTD", "QTDE", "QTD FATURADA", "QUANTIDADE FATURADA"])
    unit_col = optional_col(direct, ["PREÇO UNITÁRIO", "PRECO UNITARIO", "VALOR UNITÁRIO", "VALOR UNITARIO"])
    direct["_QTD_VENDA"] = to_number(direct[qty_col]) if qty_col else 0.0
    direct["_PRECO_REAL_UNIT"] = to_number(direct[unit_col]) if unit_col else 0.0
    infer_mask = direct["_QTD_VENDA"] <= 0
    direct.loc[infer_mask & (direct["_PRECO_REAL_UNIT"] > 0), "_QTD_VENDA"] = (
        direct.loc[infer_mask & (direct["_PRECO_REAL_UNIT"] > 0), "_VALOR"] /
        direct.loc[infer_mask & (direct["_PRECO_REAL_UNIT"] > 0), "_PRECO_REAL_UNIT"]
    )
    direct.loc[direct["_QTD_VENDA"] <= 0, "_QTD_VENDA"] = 1.0
    direct.loc[direct["_PRECO_REAL_UNIT"] <= 0, "_PRECO_REAL_UNIT"] = (
        direct.loc[direct["_PRECO_REAL_UNIT"] <= 0, "_VALOR"] / direct.loc[direct["_PRECO_REAL_UNIT"] <= 0, "_QTD_VENDA"]
    )
    direct["_PRODUCT_KEY"] = direct[product_col].map(product_key)
    direct["_PRODUCT_BASE"] = direct[product_col].map(product_base_key)

    p = price_table[["_PRODUCT_KEY", "_PRODUCT_BASE", "_PRICE_DESCRIPTION", reference]].copy()
    exact_price = p.set_index("_PRODUCT_KEY")[reference].to_dict()
    exact_desc = p.set_index("_PRODUCT_KEY")["_PRICE_DESCRIPTION"].to_dict()
    base_counts = p.groupby("_PRODUCT_BASE")["_PRODUCT_KEY"].nunique()
    unique_bases = set(base_counts[base_counts == 1].index)
    pbase = p[p["_PRODUCT_BASE"].isin(unique_bases)].drop_duplicates("_PRODUCT_BASE")
    base_price = pbase.set_index("_PRODUCT_BASE")[reference].to_dict()
    base_desc = pbase.set_index("_PRODUCT_BASE")["_PRICE_DESCRIPTION"].to_dict()

    direct["_PRECO_TABELA_UNIT"] = direct["_PRODUCT_KEY"].map(exact_price)
    direct["_DESC_TABELA"] = direct["_PRODUCT_KEY"].map(exact_desc)
    direct["_MATCH_PRECO"] = np.where(direct["_PRECO_TABELA_UNIT"].notna(), "Exato", "Não localizado")
    missing = direct["_PRECO_TABELA_UNIT"].isna()
    direct.loc[missing, "_PRECO_TABELA_UNIT"] = direct.loc[missing, "_PRODUCT_BASE"].map(base_price)
    direct.loc[missing, "_DESC_TABELA"] = direct.loc[missing, "_PRODUCT_BASE"].map(base_desc)
    direct.loc[missing & direct["_PRECO_TABELA_UNIT"].notna(), "_MATCH_PRECO"] = "Código equivalente"
    direct["_PRECO_TABELA_UNIT"] = pd.to_numeric(direct["_PRECO_TABELA_UNIT"], errors="coerce")
    direct["_VALOR_TABELA"] = direct["_QTD_VENDA"] * direct["_PRECO_TABELA_UNIT"]
    direct["_ADERENCIA_PRECO"] = np.where(direct["_VALOR_TABELA"] > 0, direct["_VALOR"] / direct["_VALOR_TABELA"], np.nan)
    direct["_DESCONTO_TABELA"] = 1 - direct["_ADERENCIA_PRECO"]
    direct["_CUSTO_ESTIMADO"] = direct["_VALOR_TABELA"] * (1 - embedded_margin)
    direct["_LUCRO_ESTIMADO"] = direct["_VALOR"] - direct["_CUSTO_ESTIMADO"]
    direct["_MARGEM_ESTIMADA"] = np.where(direct["_VALOR"] != 0, direct["_LUCRO_ESTIMADO"] / direct["_VALOR"], np.nan)
    return direct


@st.cache_data(show_spinner=False, ttl=300)
def load_rev(file_bytes: bytes) -> dict[str, pd.DataFrame | dict[str, str]]:
    states = workbook_sheet_states(file_bytes)
    visible = {name for name, state in states.items() if state == "visible"}
    required = ["Resumo Receitas", "Resumo Despesas", "Centro de Custos", "Caixa Operacional", "Performance Recebimento"]
    missing_visible = [s for s in required if s not in visible]
    if missing_visible:
        raise ValueError("Abas obrigatórias não estão visíveis na REV2026: " + ", ".join(missing_visible))

    receitas = read_excel_sheet(file_bytes, "Resumo Receitas")
    despesas = read_excel_sheet(file_bytes, "Resumo Despesas", usecols="A:N")
    mapping_raw = read_excel_raw(file_bytes, "Resumo Despesas", usecols="O:S")
    custos = read_excel_sheet(file_bytes, "Centro de Custos")
    caixa = read_excel_sheet(file_bytes, "Caixa Operacional")
    receb = read_excel_sheet(file_bytes, "Performance Recebimento")

    receitas["_MES"] = pd.to_datetime(receitas[require_col(receitas, ["MÊS"], "Resumo Receitas")], errors="coerce").dt.to_period("M")
    receitas["_VALOR"] = to_number(receitas[require_col(receitas, ["Valor Pago"], "Resumo Receitas")])
    for col in ["Cliente", "Historico", "Banco", "SEGMENTO"]:
        if col in receitas.columns:
            receitas[col] = receitas[col].fillna("Não informado").astype(str).str.strip()
    receitas["_CLIENT_KEY"] = receitas["Cliente"].map(client_key)
    receitas["_SEGMENTO_N"] = receitas["SEGMENTO"].map(norm)
    receitas["_TIPO_CAIXA"] = np.where(receitas["_SEGMENTO_N"].str.contains("LOCACAO"), "LOCACAO", "VENDA_SERVICO")
    receitas["_NAO_OPERACIONAL"] = receitas["_SEGMENTO_N"].isin(["CAPITAL DE GIRO", "OUTRAS RECEITAS"])

    mapping = mapping_raw.iloc[2:].copy()
    mapping.columns = ["Natureza_map", "Categoria_map", "PAI_map", "Classificacao_map", "Grupo_map"]
    mapping["_NATUREZA_KEY"] = mapping["Natureza_map"].map(nature_key)
    mapping = mapping[mapping["_NATUREZA_KEY"] != ""].drop_duplicates("_NATUREZA_KEY", keep="last")
    despesas["_NATUREZA_KEY"] = despesas[require_col(despesas, ["Natureza"], "Resumo Despesas")].map(nature_key)
    despesas = despesas.merge(mapping[["_NATUREZA_KEY", "PAI_map"]], on="_NATUREZA_KEY", how="left")
    despesas["PAI"] = despesas["PAI_map"].where(despesas["PAI_map"].notna(), despesas[require_col(despesas, ["PAI"], "Resumo Despesas")])
    despesas["PAI"] = despesas["PAI"].fillna("Não informado").astype(str).str.strip()
    despesas["_MES"] = pd.to_datetime(despesas[require_col(despesas, ["Vencto Real"], "Resumo Despesas")], errors="coerce").dt.to_period("M")
    despesas["_VALOR"] = to_number(despesas[require_col(despesas, ["Valor"], "Resumo Despesas")])
    for col in ["EMPRESA", "GRUPO", "SUBGRUPO", "Categoria", "Codigo-Nome do Fornecedor"]:
        if col in despesas.columns:
            despesas[col] = despesas[col].fillna("Não informado").astype(str).str.strip()
    despesas["_GRUPO_N"] = despesas["GRUPO"].map(norm)
    despesas = despesas[despesas["_MES"].notna() & despesas["_GRUPO_N"].isin(["SAIDAS OPERACIONAIS", "SAIDAS NAO OPERACIONAIS"])].copy()

    custos["_MES"] = pd.to_datetime(custos[require_col(custos, ["MÊS", "MES"], "Centro de Custos")], errors="coerce").dt.to_period("M")
    custos["_VALOR"] = to_number(custos[require_col(custos, ["Valor", "VALOR"], "Centro de Custos")])

    direct_cc_col = require_col(
        custos, ["CENTRO DE CUSTOS", "CENTRO DE CUSTO", "CENTRO CUSTOS", "CC DIRETO"], "Centro de Custos"
    )
    if direct_cc_col != "CENTRO DE CUSTOS":
        custos["CENTRO DE CUSTOS"] = custos[direct_cc_col]
    rateio_col = optional_col(
        custos, [
            "CENTRO DE CUSTOS RATEAO", "CENTRO DE CUSTOS RATEADO",
            "CENTRO DE CUSTOS RATEIO", "CENTRO CUSTOS RATEIO",
            "CENTRO CUSTOS RATEADO", "CC RATEIO", "CC RATEADO",
        ]
    )

    text_columns = [
        "EMPRESA", "GRUPO", "SUBGRUPO", "PAI", "Categoria",
        "CENTRO DE CUSTOS", "Codigo-Nome do Fornecedor",
    ]
    if rateio_col:
        text_columns.append(rateio_col)
    for col in text_columns:
        if col in custos.columns:
            custos[col] = custos[col].fillna("Não informado").astype(str).str.strip()
    custos["_EMPRESA_N"] = custos["EMPRESA"].map(norm)
    custos["_GRUPO_N"] = custos["GRUPO"].map(norm)
    custos["_CC_N"] = custos["CENTRO DE CUSTOS"].map(norm)

    # O rateio é preservado em uma base separada e exclusivamente informativa.
    # Ele nunca substitui o Centro de Custos direto nos cálculos de resultado.
    if rateio_col:
        custos["_CC_RATEIO_ORIGINAL"] = custos[rateio_col]
        custos["_CC_RATEIO_N"] = custos[rateio_col].map(norm)
        custos["_LINHA_RATEIO"] = custos[rateio_col].map(rateio_line_key)
    else:
        custos["_CC_RATEIO_ORIGINAL"] = "Não informado"
        custos["_CC_RATEIO_N"] = ""
        custos["_LINHA_RATEIO"] = "NAO CLASSIFICADA"
    rateio = custos[
        (custos["_EMPRESA_N"] == "DESPESAS") &
        custos["_LINHA_RATEIO"].isin(LINES) &
        (~custos["_CC_N"].isin(LINES))
    ].copy()

    # Cenário alternativo: resultado integral por linha após o rateio administrativo.
    # Aqui receitas e despesas são atribuídas pela coluna Centro de Custos Rateado.
    # Esta base não altera o resultado direto, que continua usando Centro de Custos.
    custos_rateados = custos[
        custos["_EMPRESA_N"].isin(["DESPESAS", "RECEITA"]) &
        custos["_LINHA_RATEIO"].isin(LINES)
    ].copy()
    rateio_meta = {
        "column": rateio_col or "Não localizada",
        "rows": int(len(rateio)),
        "total": float(rateio["_VALOR"].sum()) if not rateio.empty else 0.0,
        "unclassified_rows": int((custos["_LINHA_RATEIO"] == "NAO CLASSIFICADA").sum()),
    }

    # Regra de gestão: toda visão departamental e todo resultado usam exclusivamente CENTRO DE CUSTOS.
    custos = custos.drop(
        columns=[rateio_col] if rateio_col else [], errors="ignore"
    ).drop(columns=["_CC_RATEIO_ORIGINAL", "_CC_RATEIO_N", "_LINHA_RATEIO"], errors="ignore")
    custos["_LINHA_DIRETA"] = custos["_CC_N"].where(custos["_CC_N"].isin(LINES), "NAO CLASSIFICADA")
    custos = custos[custos["_EMPRESA_N"].isin(["DESPESAS", "RECEITA"])].copy()

    caixa["_MES"] = pd.to_datetime(caixa[require_col(caixa, ["Mês", "MÊS"], "Caixa Operacional")], errors="coerce").dt.to_period("M")
    caixa["_RECEITAS"] = to_number(caixa[require_col(caixa, ["Receitas Realizadas"], "Caixa Operacional")])
    caixa["_DESPESAS"] = to_number(caixa[require_col(caixa, ["Despesas Realizadas"], "Caixa Operacional")])
    caixa = caixa[caixa["_MES"].notna()].copy()

    receb["_MES"] = pd.to_datetime(receb[require_col(receb, ["Mês", "MÊS"], "Performance Recebimento")], errors="coerce").dt.to_period("M")
    receb["_PREVISTO"] = to_number(receb[require_col(receb, ["Recebimento Previsto"], "Performance Recebimento")])
    receb["_REALIZADO"] = to_number(receb[require_col(receb, ["Recebimento Realizado"], "Performance Recebimento")])
    receb = receb[receb["_MES"].notna()].copy()

    return {
        "receitas": receitas, "despesas": despesas, "custos": custos,
        "rateio": rateio, "custos_rateados": custos_rateados,
        "rateio_meta": rateio_meta, "caixa": caixa,
        "recebimento": receb, "sheet_states": states,
    }


def correct_microtech_receipts_in_vendas(
    custos: pd.DataFrame, fat: pd.DataFrame,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """Corrige recebimentos Microtech lançados no CC Vendas.

    A correção é restrita a receitas operacionais e exige evidência na BASE BI:
    1) mesma nota + mesmo cliente com GERENTE = CELSO; ou
    2) cliente com 100% do faturamento de 2026 atribuído ao CELSO, quando a nota não casa.

    Custos/despesas nunca são alterados e CENTRO DE CUSTOS RATEAO não é utilizado.
    """
    out = custos.copy()
    stats: dict[str, object] = {"count": 0, "value": 0.0, "details": pd.DataFrame()}
    required_cost = {"_EMPRESA_N", "_GRUPO_N", "_CC_N", "CENTRO DE CUSTOS", "Codigo-Nome do Fornecedor", "Prf-Numero Parcela", "_VALOR"}
    required_fat = {"_CLIENT_KEY", "_VALOR", "_LINHA", "GERENTE", "Nota Fiscal"}
    if not required_cost.issubset(out.columns) or not required_fat.issubset(fat.columns):
        return out, stats

    out["_CENTRO_DE_CUSTOS_ORIGINAL"] = out["CENTRO DE CUSTOS"]
    out["_CC_AJUSTADO"] = False
    out["_CRITERIO_CC"] = "Centro de Custos original"
    out["_CLIENT_KEY_CC"] = out["Codigo-Nome do Fornecedor"].map(client_key)
    out["_DOC_KEY_CC"] = out["Prf-Numero Parcela"].map(document_key)

    billing = fat.copy()
    billing["_GERENTE_N"] = billing["GERENTE"].map(norm)
    billing["_DOC_KEY_FAT"] = billing["Nota Fiscal"].map(document_key)
    billing = billing[(billing["_GERENTE_N"] == "CELSO") & (billing["_LINHA"] == "MICROTECH")].copy()
    if billing.empty:
        return out, stats

    exact_pairs = set(
        billing.loc[(billing["_CLIENT_KEY"] != "") & (billing["_DOC_KEY_FAT"] != ""), ["_CLIENT_KEY", "_DOC_KEY_FAT"]]
        .drop_duplicates().itertuples(index=False, name=None)
    )

    client_all = (
        fat.groupby(["_CLIENT_KEY", "_LINHA"], as_index=False)["_VALOR"].sum()
        .sort_values("_VALOR", ascending=False)
    )
    client_all["_TOTAL_CLIENTE"] = client_all.groupby("_CLIENT_KEY")["_VALOR"].transform("sum")
    client_all["_SHARE"] = np.where(client_all["_TOTAL_CLIENTE"] != 0, client_all["_VALOR"] / client_all["_TOTAL_CLIENTE"], 0)
    exclusive_clients = set(
        client_all.loc[(client_all["_LINHA"] == "MICROTECH") & (client_all["_SHARE"] >= .999999), "_CLIENT_KEY"]
    )

    revenue_mask = (
        (out["_EMPRESA_N"] == "RECEITA") &
        (out["_GRUPO_N"] == "RECEITAS OPERACIONAIS") &
        (out["_CC_N"] == "VENDAS")
    )
    exact_mask = pd.Series(
        [(ck, dk) in exact_pairs for ck, dk in zip(out["_CLIENT_KEY_CC"], out["_DOC_KEY_CC"])],
        index=out.index,
    )
    fallback_mask = (out["_DOC_KEY_CC"] == "") | (~exact_mask)
    client_mask = out["_CLIENT_KEY_CC"].isin(exclusive_clients) & fallback_mask
    adjust_mask = revenue_mask & (exact_mask | client_mask)

    out.loc[adjust_mask, "CENTRO DE CUSTOS"] = "MICROTECH"
    out.loc[adjust_mask, "_CC_N"] = "MICROTECH"
    out.loc[adjust_mask, "_LINHA_DIRETA"] = "MICROTECH"
    out.loc[adjust_mask, "_CC_AJUSTADO"] = True
    out.loc[adjust_mask & exact_mask, "_CRITERIO_CC"] = "Nota + cliente + gerente Celso"
    out.loc[adjust_mask & ~exact_mask, "_CRITERIO_CC"] = "Cliente exclusivo Microtech"

    details_cols = [
        "MÊS", "Codigo-Nome do Fornecedor", "Prf-Numero Parcela", "_VALOR",
        "_CENTRO_DE_CUSTOS_ORIGINAL", "CENTRO DE CUSTOS", "_CRITERIO_CC",
    ]
    details = out.loc[adjust_mask, [c for c in details_cols if c in out.columns]].copy()
    stats = {
        "count": int(adjust_mask.sum()),
        "value": float(out.loc[adjust_mask, "_VALOR"].sum()),
        "details": details,
    }
    return out, stats


VARIABLE_COSTS = {
    "FECHAMENTOS DE CAMBIO", "DESEMBARACOS", "PECAS/ACESSORIOS E EMBALAGENS", "LOCACAO DE EQPTOS.",
    "COFINS", "COMISSOES DE REPRESENTANTES", "COMISSOES ESPECIAIS", "CARTOES DE CREDITO",
    "MAQUINAS/EQPTOS", "PIS", "OUTROS EQUIPAMENTOS", "TRANSPORTADORAS", "ICMS", "ISS", "FRETES",
}


def cost_class(pai: object) -> str:
    n = norm(pai)
    if "IRPJ" in n or "CSLL" in n:
        return "IMPOSTO_RESULTADO"
    if n in VARIABLE_COSTS or any(x in n for x in ["COMIS", "FRETE", "CAMBIO", "DESEMBARACO", "PECA"]):
        return "VARIAVEL"
    return "FIXO"


@st.cache_data(show_spinner=False)
def assign_receipt_lines(receitas: pd.DataFrame, fat: pd.DataFrame, threshold: float = 88.0) -> tuple[pd.DataFrame, dict[str, float]]:
    rec = receitas.copy()
    billing_map = (
        fat.groupby(["_CLIENT_KEY", "_TIPO_CAIXA", "_LINHA"], as_index=False)["_VALOR"].sum()
        .sort_values("_VALOR", ascending=False)
        .drop_duplicates(["_CLIENT_KEY", "_TIPO_CAIXA"])
    )
    line_map = {(r["_CLIENT_KEY"], r["_TIPO_CAIXA"]): r["_LINHA"] for _, r in billing_map.iterrows()}
    choices = {
        tipo: billing_map.loc[billing_map["_TIPO_CAIXA"] == tipo, "_CLIENT_KEY"].drop_duplicates().tolist()
        for tipo in billing_map["_TIPO_CAIXA"].dropna().unique()
    }

    cache: dict[tuple[str, str], tuple[str | None, float, str]] = {}
    for tipo in rec["_TIPO_CAIXA"].dropna().unique():
        tipo_choices = choices.get(tipo, [])
        for key in rec.loc[rec["_TIPO_CAIXA"] == tipo, "_CLIENT_KEY"].drop_duplicates():
            if (key, tipo) in line_map:
                cache[(key, tipo)] = (key, 100.0, "Exata")
            elif key and tipo_choices:
                matched, score = best_match(key, tipo_choices, cutoff=threshold)
                cache[(key, tipo)] = (matched, float(score), "Similar") if matched else (None, 0.0, "Fallback")
            else:
                cache[(key, tipo)] = (None, 0.0, "Fallback")

    matched_keys, scores, methods, lines = [], [], [], []
    for key, tipo, is_non_op in zip(rec["_CLIENT_KEY"], rec["_TIPO_CAIXA"], rec["_NAO_OPERACIONAL"]):
        if is_non_op:
            matched_keys.append(None); scores.append(100.0); methods.append("Não operacional"); lines.append("NAO OPERACIONAL")
            continue
        mk, score, method = cache.get((key, tipo), (None, 0.0, "Fallback"))
        line = line_map.get((mk, tipo)) if mk else None
        if line is None:
            line = "LOCACAO" if tipo == "LOCACAO" else "VENDAS"
        matched_keys.append(mk); scores.append(score); methods.append(method); lines.append(line)

    rec["_MATCH_KEY"] = matched_keys
    rec["_MATCH_SCORE"] = scores
    rec["_MATCH_METHOD"] = methods
    rec["_LINHA"] = lines
    op = rec[~rec["_NAO_OPERACIONAL"]]
    secure = op[op["_MATCH_METHOD"].isin(["Exata", "Similar"])]
    stats = {
        "coverage_rows": safe_div(len(secure), len(op)),
        "coverage_value": safe_div(float(secure["_VALOR"].sum()), float(op["_VALOR"].sum())),
        "fallback_value": float(op.loc[op["_MATCH_METHOD"] == "Fallback", "_VALOR"].sum()),
    }
    return rec, stats


@st.cache_data(show_spinner=False)
def read_optional_table(file_bytes: bytes, filename: str) -> tuple[pd.DataFrame, dict[str, object]]:
    metadata: dict[str, object] = {"origem": filename, "sheet": ""}
    if filename.lower().endswith(".csv"):
        try:
            df = clean_columns(pd.read_csv(io.BytesIO(file_bytes), sep=None, engine="python", encoding="utf-8-sig"))
        except UnicodeDecodeError:
            df = clean_columns(pd.read_csv(io.BytesIO(file_bytes), sep=None, engine="python", encoding="latin1"))
        metadata["sheet"] = "CSV"
        return df, metadata

    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    normalized = {norm(s): s for s in xls.sheet_names}
    detail_sheet = normalized.get("TITULOS DETALHADOS") or normalized.get("TITULOS") or xls.sheet_names[0]
    df = clean_columns(pd.read_excel(io.BytesIO(file_bytes), sheet_name=detail_sheet, engine="openpyxl"))
    metadata["sheet"] = detail_sheet

    resumo_sheet = normalized.get("RESUMO")
    if resumo_sheet:
        resumo = clean_columns(pd.read_excel(io.BytesIO(file_bytes), sheet_name=resumo_sheet, engine="openpyxl"))
        gerente_col = optional_col(resumo, ["Gerente", "Gestor"])
        valor_col = optional_col(resumo, ["Valor", "Saldo Vencido", "Vencidos Corrigidos"])
        clientes_col = optional_col(resumo, ["Clientes"])
        titulos_col = optional_col(resumo, ["Titulos", "Títulos"])
        atraso_col = optional_col(resumo, ["Maior atraso", "Maior Atraso"])
        if gerente_col and valor_col:
            resumo_out = pd.DataFrame({
                "Gerente": resumo[gerente_col].fillna("Sem gerente").astype(str).str.strip(),
                "Valor": to_number(resumo[valor_col]),
                "Clientes": to_number(resumo[clientes_col]).astype(int) if clientes_col else 0,
                "Títulos": to_number(resumo[titulos_col]).astype(int) if titulos_col else 0,
                "Maior atraso": to_number(resumo[atraso_col]).astype(int) if atraso_col else 0,
            })
            resumo_out["Linha"] = resumo_out["Gerente"].map(norm).map(MANAGER_LINE_MAP).fillna("NAO CLASSIFICADA")
            metadata["resumo_gerentes"] = resumo_out.to_dict("records")
    return df, metadata


@st.cache_data(show_spinner=False)
def prepare_inadimplencia(file_bytes: bytes, filename: str, fat: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, object]]:
    df, source_meta = read_optional_table(file_bytes, filename)
    client_col = optional_col(df, [
        "Nome do Cliente", "Nome Cliente", "nome_cliente", "Cliente", "Razão Social", "Razao Social",
        "Nome", "N Fantasia", "Codigo-Lj-Nome do Cliente"
    ])
    value_col = optional_col(df, [
        "Saldo atual", "saldo_atual", "Tit Vencidos Valor Corrigido", "Vencidos Corrigidos",
        "Tit Vencidos Valor Atual", "Vencidos", "Saldo Vencido", "Valor Vencido",
        "Valor Original", "Saldo", "Valor"
    ])
    due_col = optional_col(df, ["Vencto Real", "Vencimento", "vencimento", "Vencto Titulo", "Vencto", "Data Vencimento", "Venc. Real"])
    issue_col = optional_col(df, ["Data Emissão", "Data Emissao", "DT Emissão", "DT Emissao", "dt_emissao", "Emissão", "Emissao"])
    days_col = optional_col(df, ["Dias Atraso", "Dias em Atraso", "Atraso"])
    line_col = optional_col(df, ["Linha", "Linha de Negócio", "Segmento", "Centro de Custos"])
    manager_col = optional_col(df, ["Gerente", "gerente", "Gestor", "Responsável", "Responsavel"])
    seller_col = optional_col(df, ["Vendedor", "vendedor", "Vendedor / Representante", "Representante"])
    title_col = optional_col(df, [
        "numero_titulo", "Número Título", "Numero Titulo", "Prf-Numero Parcela", "Prf-Numero-Parcela",
        "Título", "Titulo", "Documento", "Nota Fiscal", "NF"
    ])
    prefix_col = optional_col(df, ["prefixo", "Prefixo"])
    installment_col = optional_col(df, ["parcela", "Parcela"])
    customer_code_col = optional_col(df, ["cliente_codigo", "Código Cliente", "Codigo Cliente", "Cod Cliente"])
    store_col = optional_col(df, ["loja", "Loja"])

    if client_col is None or value_col is None:
        raise ValueError("A base de inadimplência precisa conter pelo menos Cliente e Saldo atual/Valor vencido.")

    x = df.copy()
    x["_CLIENTE"] = x[client_col].fillna("Não informado").astype(str).str.strip()
    x["_CLIENT_KEY"] = x["_CLIENTE"].map(client_key)
    x["_VALOR_VENCIDO"] = to_number(x[value_col])
    x["_VENCIMENTO"] = to_datetime_mixed(x[due_col]) if due_col else pd.NaT
    x["_EMISSAO"] = to_datetime_mixed(x[issue_col]) if issue_col else pd.NaT
    x["_VENDEDOR"] = x[seller_col].fillna("Sem vendedor").astype(str).str.strip() if seller_col else "Sem vendedor"
    x["_COD_CLIENTE"] = x[customer_code_col].fillna("").astype(str).str.replace(r"\.0$", "", regex=True) if customer_code_col else ""
    x["_LOJA"] = x[store_col].fillna("").astype(str).str.replace(r"\.0$", "", regex=True) if store_col else ""

    if days_col:
        x["_DIAS_ATRASO"] = to_number(x[days_col]).astype(int)
    elif due_col:
        x["_DIAS_ATRASO"] = (pd.Timestamp.today().normalize() - x["_VENCIMENTO"]).dt.days.fillna(0).clip(lower=0).astype(int)
    else:
        x["_DIAS_ATRASO"] = 0
    x["_MES"] = x["_VENCIMENTO"].dt.to_period("M") if due_col else pd.Period(pd.Timestamp.today(), freq="M")

    def _id_text(series: pd.Series) -> pd.Series:
        return series.fillna("").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()

    if title_col:
        title_txt = _id_text(x[title_col])
        if prefix_col:
            prefix_txt = _id_text(x[prefix_col])
            title_txt = np.where(prefix_txt != "", prefix_txt + "-" + title_txt, title_txt)
            title_txt = pd.Series(title_txt, index=x.index)
        if installment_col:
            installment_txt = _id_text(x[installment_col])
            title_txt = pd.Series(np.where(installment_txt != "", title_txt.astype(str) + "/" + installment_txt, title_txt), index=x.index)
        x["_TITULO"] = title_txt.astype(str)
    else:
        x["_TITULO"] = ""

    manager_map = (
        fat.groupby(["_CLIENT_KEY", "GERENTE", "_LINHA"], as_index=False)["_VALOR"].sum()
        .sort_values("_VALOR", ascending=False).drop_duplicates("_CLIENT_KEY")
        if "GERENTE" in fat.columns else pd.DataFrame(columns=["_CLIENT_KEY", "GERENTE", "_LINHA"])
    )
    choices = manager_map["_CLIENT_KEY"].drop_duplicates().tolist()
    manager_by_key = dict(zip(manager_map["_CLIENT_KEY"], manager_map["GERENTE"]))
    line_by_key = dict(zip(manager_map["_CLIENT_KEY"], manager_map["_LINHA"]))

    if manager_col:
        x["_GERENTE"] = x[manager_col].fillna("Sem gerente").astype(str).str.strip()
        x["_LINHA"] = x["_GERENTE"].map(norm).map(MANAGER_LINE_MAP).fillna("NAO CLASSIFICADA")
        x["_MATCH_GERENTE"] = "Direto no CRM"
    elif line_col:
        raw_line = x[line_col].map(norm)
        x["_LINHA"] = np.select(
            [raw_line.str.contains("MICROTECH"), raw_line.str.contains("ENDOSCOPIA"), raw_line.str.contains("LOCACAO"), raw_line.str.contains("VENDA")],
            ["MICROTECH", "ENDOSCOPIA", "LOCACAO", "VENDAS"], default="NAO CLASSIFICADA"
        )
        x["_GERENTE"] = x["_LINHA"].map(LINE_MANAGER_MAP).fillna("Sem gerente")
        x["_MATCH_GERENTE"] = "Direto na linha"
    else:
        cache: dict[str, tuple[str, str, str]] = {}
        for key in x["_CLIENT_KEY"].drop_duplicates():
            if key in line_by_key:
                cache[key] = (manager_by_key.get(key, "Sem gerente"), line_by_key.get(key, "NAO CLASSIFICADA"), "Exata")
            elif key:
                matched, _score = best_match(key, choices, cutoff=88)
                cache[key] = (
                    manager_by_key.get(matched, "Sem gerente") if matched else "Sem gerente",
                    line_by_key.get(matched, "NAO CLASSIFICADA") if matched else "NAO CLASSIFICADA",
                    "Similar" if matched else "Não classificada",
                )
            else:
                cache[key] = ("Sem gerente", "NAO CLASSIFICADA", "Não classificada")
        x["_GERENTE"] = x["_CLIENT_KEY"].map(lambda k: cache.get(k, ("Sem gerente", "NAO CLASSIFICADA", "Não classificada"))[0])
        x["_LINHA"] = x["_CLIENT_KEY"].map(lambda k: cache.get(k, ("Sem gerente", "NAO CLASSIFICADA", "Não classificada"))[1])
        x["_MATCH_GERENTE"] = x["_CLIENT_KEY"].map(lambda k: cache.get(k, ("Sem gerente", "NAO CLASSIFICADA", "Não classificada"))[2])

    x = x[(x["_VALOR_VENCIDO"] > 0) & ((x["_DIAS_ATRASO"] > 0) | x["_VENCIMENTO"].isna())].copy()
    x["_FAIXA"] = pd.cut(
        x["_DIAS_ATRASO"], bins=[-1, 30, 60, 90, 10**9],
        labels=["Até 30 dias", "31 a 60 dias", "61 a 90 dias", "Acima de 90 dias"]
    ).astype(str)

    calculated_summary = (
        x.groupby(["_GERENTE", "_LINHA"], as_index=False)
        .agg(
            Valor=("_VALOR_VENCIDO", "sum"),
            Clientes=("_CLIENTE", "nunique"),
            Títulos=("_TITULO", "size"),
            **{"Maior atraso": ("_DIAS_ATRASO", "max")},
        )
        .rename(columns={"_GERENTE": "Gerente", "_LINHA": "Linha"})
    )

    meta: dict[str, object] = {
        "cliente": client_col,
        "valor": value_col,
        "vencimento": due_col or "Não disponível",
        "linha": line_col or ("Definida pelo gerente do CRM" if manager_col else "Mapeada pelo gerente dominante da BASE BI"),
        "gerente": manager_col or "Mapeado pelo cliente na BASE BI",
        "vendedor": seller_col or "Não disponível",
        "sheet": source_meta.get("sheet", ""),
        "resumo_gerentes": source_meta.get("resumo_gerentes") or calculated_summary.to_dict("records"),
        "formato": "CRM CSV" if filename.lower().endswith(".csv") else "Relatório Excel",
    }
    return x, meta


# =========================================================
# AUTENTICAÇÃO E PERFIS
# =========================================================
def password_hash(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


DEFAULT_USERS = {
    "diretoria": {
        "nome": "Diretoria", "usuario": "diretoria", "email": "",
        "senha_hash": "cbae32987728a10b19e2528695dbf676c9b8de0f539bab08b5789c2e9f0d8599",
        "perfil": "DIRETORIA", "linha": "CONSOLIDADO",
    },
    "paula": {
        "nome": "Paula", "usuario": "paula", "email": "paulamayara10@gmail.com",
        "senha_hash": "ad27a22ae449782700c94a3acfa95e19cb5b3a3907b7da088d11569d32bfe800",
        "perfil": "CONTROLADORIA", "linha": "CONSOLIDADO",
    },
    "celso": {
        "nome": "Celso", "usuario": "celso", "email": "",
        "senha_hash": "35a4526e8b23c095c14d926e1f4e5614f70c20069200d3cf4f8ef4da80cb240a",
        "perfil": "GESTOR", "linha": "MICROTECH",
    },
    "renato": {
        "nome": "Renato", "usuario": "renato", "email": "",
        "senha_hash": "353b16f690d6dadaa2a675f73963bdd60ddc36979922100340761dd3f712341a",
        "perfil": "GESTOR", "linha": "VENDAS",
    },
    "amauri": {
        "nome": "Amauri", "usuario": "amauri", "email": "",
        "senha_hash": "5bda933352db6ef4569709d432a7d1ac793977c82fb982c5bae65651015493c5",
        "perfil": "GESTOR", "linha": "LOCACAO",
    },
    "ronaldo": {
        "nome": "Ronaldo", "usuario": "ronaldo", "email": "",
        "senha_hash": "21ce9c880661ca01dc3af90a1803b32244e1743ba8fef495e83ef10204546fc2",
        "perfil": "GESTOR", "linha": "ENDOSCOPIA",
    },
}


def secret_users() -> dict[str, dict]:
    users = {key: dict(value) for key, value in DEFAULT_USERS.items()}
    try:
        raw = st.secrets.get("usuarios", {})
        for key, value in raw.items():
            users[str(key)] = dict(value)
    except Exception:
        pass
    return users


def authenticate() -> dict[str, str]:
    users = secret_users()
    if "auth_user" not in st.session_state:
        profile_order = ["diretoria", "paula", "celso", "renato", "amauri", "ronaldo"]
        available = [key for key in profile_order if key in users]
        profile_labels = {
            "diretoria": "Diretoria",
            "paula": "Paula · Controladoria",
            "celso": "Celso · Microtech",
            "renato": "Renato · Vendas",
            "amauri": "Amauri · Locação",
            "ronaldo": "Ronaldo · Endoscopia",
        }
        left, center, right = st.columns([1, 1.05, 1])
        with center:
            st.markdown(
                """
                <div class='login-brand'>
                  <div class='logo'>FIRST <span>INTELLIGENCE</span></div>
                  <p>Business Performance por perfil</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            with st.form("login_form", clear_on_submit=False):
                selected = st.selectbox(
                    "Perfil de acesso",
                    available,
                    format_func=lambda key: profile_labels.get(key, users[key].get("nome", key)),
                )
                password = st.text_input("Senha", type="password")
                submitted = st.form_submit_button("Acessar painel", width="stretch")
            if submitted:
                cfg = users[selected]
                expected = str(cfg.get("senha_hash", ""))
                valid = bool(expected) and password_hash(password) == expected
                if not valid and cfg.get("senha") is not None:
                    valid = password == str(cfg.get("senha"))
                if valid:
                    st.session_state["auth_user"] = {
                        "nome": str(cfg.get("nome", selected)),
                        "email": str(cfg.get("email", "")),
                        "perfil": norm(cfg.get("perfil", "GESTOR")),
                        "linha": norm(cfg.get("linha", "VENDAS")),
                        "secure": "Sim",
                    }
                    st.rerun()
                st.error("Senha incorreta para o perfil selecionado.")
        st.stop()
    return st.session_state["auth_user"]


# =========================================================
# CÁLCULOS DE CAIXA
# =========================================================
def period_filter(df: pd.DataFrame, start: pd.Period, end: pd.Period) -> pd.DataFrame:
    return df[df["_MES"].between(start, end)].copy()


def allocated_cost_center_result(
    custos_rateados: pd.DataFrame, start: pd.Period, end: pd.Period,
) -> pd.DataFrame:
    """Apura o resultado de cada linha pela coluna Centro de Custos Rateado."""
    columns = [
        "Código", "Linha", "Receitas após rateio", "Despesas após rateio",
        "Resultado após rateio", "Margem após rateio",
    ]
    if custos_rateados is None or custos_rateados.empty:
        return pd.DataFrame([
            {
                "Código": line, "Linha": line_label(line),
                "Receitas após rateio": 0.0, "Despesas após rateio": 0.0,
                "Resultado após rateio": 0.0, "Margem após rateio": 0.0,
            }
            for line in LINES
        ], columns=columns)

    period = period_filter(custos_rateados, start, end).copy()
    rows = []
    for line in LINES:
        part = period[period["_LINHA_RATEIO"] == line]
        revenue = float(part.loc[part["_EMPRESA_N"] == "RECEITA", "_VALOR"].sum())
        expense = float(part.loc[part["_EMPRESA_N"] == "DESPESAS", "_VALOR"].sum())
        result = revenue - expense
        rows.append({
            "Código": line,
            "Linha": line_label(line),
            "Receitas após rateio": revenue,
            "Despesas após rateio": expense,
            "Resultado após rateio": result,
            "Margem após rateio": safe_div(result, revenue),
        })
    return pd.DataFrame(rows, columns=columns)


def billing_by_line(fat: pd.DataFrame, line: str) -> pd.DataFrame:
    """Filtra pela classificação canônica criada uma única vez na BASE BI.

    Gerentes oficiais continuam soberanos (Celso, Renato, Amauri e Ronaldo).
    Apenas gerente ausente usa segmento/produto como contingência. Um gerente informado
    fora do mapa permanece no consolidado, mas não é atribuído à linha de outro gestor.
    """
    if line == "CONSOLIDADO":
        return fat.copy()
    return fat[fat["_LINHA"] == line].copy()


def billing_reconciliation(
    fat: pd.DataFrame, start: pd.Period, end: pd.Period,
) -> tuple[dict[str, float | int], pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Concilia faturamento consolidado, linhas, gerentes e meses para auditoria."""
    period = period_filter(fat, start, end).copy()
    total = float(period["_VALOR"].sum())

    line_rows = []
    for line in LINES:
        part = period[period["_LINHA"] == line].copy()
        note_col = optional_col(part, ["Nota Fiscal", "NOTA FISCAL", "NF"])
        line_rows.append({
            "Linha": line_label(line),
            "Código": line,
            "Gerente oficial": LINE_MANAGER_MAP.get(line, ""),
            "Faturamento": float(part["_VALOR"].sum()),
            "Linhas da base": int(len(part)),
            "Notas fiscais": int(part[note_col].nunique()) if note_col and not part.empty else 0,
            "Via contingência": float(part.loc[part["_CRITERIO_LINHA"] != "Gerente oficial", "_VALOR"].sum()) if "_CRITERIO_LINHA" in part.columns else 0.0,
        })
    by_line = pd.DataFrame(line_rows)
    sum_lines = float(by_line["Faturamento"].sum())
    unclassified = period[~period["_LINHA"].isin(LINES)].copy()

    manager_col = optional_col(period, ["GERENTE"])
    if manager_col:
        by_manager = (
            period.assign(Gerente=period[manager_col].fillna("Não informado").astype(str).str.strip())
            .groupby("Gerente", as_index=False)
            .agg(Faturamento=("_VALOR", "sum"), Linhas=("_VALOR", "size"))
            .sort_values("Faturamento", ascending=False)
        )
    else:
        by_manager = pd.DataFrame(columns=["Gerente", "Faturamento", "Linhas"])

    monthly_total = period.groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_VALOR": "Consolidado"})
    monthly_line = period[period["_LINHA"].isin(LINES)].groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_VALOR": "Quatro linhas oficiais"})
    monthly_other = period[~period["_LINHA"].isin(LINES)].groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_VALOR": "Outras gerências"})
    monthly = monthly_total.merge(monthly_line, on="_MES", how="outer").merge(monthly_other, on="_MES", how="outer").fillna(0)
    monthly["Diferença técnica"] = monthly["Consolidado"] - monthly["Quatro linhas oficiais"] - monthly["Outras gerências"]
    monthly["Mês"] = monthly["_MES"].map(month_label)
    monthly = monthly[["Mês", "Consolidado", "Quatro linhas oficiais", "Outras gerências", "Diferença técnica"]]

    unclassified_value = float(unclassified["_VALOR"].sum()) if not unclassified.empty else 0.0
    technical_difference = total - sum_lines - unclassified_value
    summary = {
        "total": total, "sum_lines": sum_lines, "difference": total - sum_lines,
        "technical_difference": technical_difference,
        "rows": int(len(period)), "unclassified_rows": int(len(unclassified)),
        "unclassified_value": unclassified_value,
        "fallback_value": float(period.loc[period.get("_CRITERIO_LINHA", "") == "Segmento / produto (gerente ausente)", "_VALOR"].sum()) if "_CRITERIO_LINHA" in period.columns else 0.0,
        "unmapped_manager_value": float(period.loc[period.get("_CRITERIO_LINHA", "") == "Gerente fora do mapa oficial", "_VALOR"].sum()) if "_CRITERIO_LINHA" in period.columns else 0.0,
    }
    return summary, by_line, by_manager, monthly


def financial_crosscheck(
    receitas: pd.DataFrame, despesas: pd.DataFrame, custos: pd.DataFrame,
    start: pd.Period, end: pd.Period,
) -> pd.DataFrame:
    """Compara os resumos financeiros com a aba Centro de Custos no mesmo período."""
    rec = period_filter(receitas, start, end)
    exp = period_filter(despesas, start, end)
    cc = period_filter(custos, start, end)

    resumo_rec = float(rec.loc[~rec["_NAO_OPERACIONAL"], "_VALOR"].sum())
    cc_rec_all = float(cc.loc[(cc["_EMPRESA_N"] == "RECEITA") & (cc["_GRUPO_N"] == "RECEITAS OPERACIONAIS"), "_VALOR"].sum())
    resumo_exp = float(exp.loc[exp["_GRUPO_N"] == "SAIDAS OPERACIONAIS", "_VALOR"].sum())
    cc_exp_all = float(cc.loc[(cc["_EMPRESA_N"] == "DESPESAS") & (cc["_GRUPO_N"] == "SAIDAS OPERACIONAIS"), "_VALOR"].sum())

    cc_rec_direct = float(cc.loc[
        (cc["_EMPRESA_N"] == "RECEITA") & (cc["_GRUPO_N"] == "RECEITAS OPERACIONAIS") & cc["_LINHA_DIRETA"].isin(LINES), "_VALOR"
    ].sum())
    cc_exp_direct = float(cc.loc[
        (cc["_EMPRESA_N"] == "DESPESAS") & (cc["_GRUPO_N"] == "SAIDAS OPERACIONAIS") & cc["_LINHA_DIRETA"].isin(LINES), "_VALOR"
    ].sum())

    return pd.DataFrame([
        {"Indicador": "Receitas operacionais", "Resumo financeiro": resumo_rec, "Centro de Custos": cc_rec_all, "Diferença": resumo_rec - cc_rec_all, "Direto nas quatro linhas": cc_rec_direct, "Geral / não classificado": cc_rec_all - cc_rec_direct},
        {"Indicador": "Saídas operacionais", "Resumo financeiro": resumo_exp, "Centro de Custos": cc_exp_all, "Diferença": resumo_exp - cc_exp_all, "Direto nas quatro linhas": cc_exp_direct, "Geral / não classificado": cc_exp_all - cc_exp_direct},
    ])


def company_cash_monthly(
    fat: pd.DataFrame, metas: pd.DataFrame, receitas: pd.DataFrame, despesas: pd.DataFrame,
    performance: pd.DataFrame, start: pd.Period, end: pd.Period,
) -> pd.DataFrame:
    months = pd.period_range(start, end, freq="M")
    out = pd.DataFrame({"Mês": months})

    f = period_filter(fat, start, end).groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_MES": "Mês", "_VALOR": "Faturamento"})
    m = period_filter(metas, start, end).groupby("_MES", as_index=False)["_META"].sum().rename(columns={"_MES": "Mês", "_META": "Meta"})
    r = period_filter(receitas, start, end)
    op = r[~r["_NAO_OPERACIONAL"]].groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_MES": "Mês", "_VALOR": "Receitas Operacionais"})
    non = r[r["_NAO_OPERACIONAL"]].groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_MES": "Mês", "_VALOR": "Entradas Não Operacionais"})

    e = period_filter(despesas, start, end).copy()
    e["_CLASSE_CAIXA"] = e["PAI"].map(cost_class)
    op_e = e[e["_GRUPO_N"] == "SAIDAS OPERACIONAIS"]
    exp = op_e.groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_MES": "Mês", "_VALOR": "Saídas Operacionais"})
    var = op_e[op_e["_CLASSE_CAIXA"] == "VARIAVEL"].groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_MES": "Mês", "_VALOR": "Saídas Variáveis"})
    fix = op_e[op_e["_CLASSE_CAIXA"] == "FIXO"].groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_MES": "Mês", "_VALOR": "Saídas Fixas"})
    taxes = op_e[op_e["_CLASSE_CAIXA"] == "IMPOSTO_RESULTADO"].groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_MES": "Mês", "_VALOR": "IRPJ e CSLL"})
    p = period_filter(performance, start, end).groupby("_MES", as_index=False)[["_PREVISTO", "_REALIZADO"]].sum().rename(columns={"_MES": "Mês", "_PREVISTO": "Recebimento Previsto", "_REALIZADO": "Recebimento Realizado"})

    for df in [f, m, op, non, exp, var, fix, taxes, p]:
        out = out.merge(df, on="Mês", how="left")
    out = out.fillna(0)
    out["Resultado Operacional de Caixa"] = out["Receitas Operacionais"] - out["Saídas Operacionais"]
    out["EBITDA Gerencial de Caixa"] = out["Resultado Operacional de Caixa"] + out["IRPJ e CSLL"]
    out["Margem de Caixa"] = np.where(out["Receitas Operacionais"] != 0, out["Resultado Operacional de Caixa"] / out["Receitas Operacionais"], 0)
    out["Margem EBITDA Caixa"] = np.where(out["Receitas Operacionais"] != 0, out["EBITDA Gerencial de Caixa"] / out["Receitas Operacionais"], 0)
    out["Contribuição de Caixa"] = out["Receitas Operacionais"] - out["Saídas Variáveis"]
    out["Margem de Contribuição Caixa"] = np.where(out["Receitas Operacionais"] != 0, out["Contribuição de Caixa"] / out["Receitas Operacionais"], 0)
    out["Conversão em Caixa"] = np.where(out["Faturamento"] != 0, out["Receitas Operacionais"] / out["Faturamento"], 0)
    out["Performance de Recebimento"] = np.where(out["Recebimento Previsto"] != 0, out["Recebimento Realizado"] / out["Recebimento Previsto"], 0)
    out["Atingimento da Meta"] = np.where(out["Meta"] != 0, out["Faturamento"] / out["Meta"], 0)
    out["Mês Texto"] = out["Mês"].map(month_label)
    return out


def commercial_performance_monthly(
    fat: pd.DataFrame, metas_g: pd.DataFrame, start: pd.Period, end: pd.Period,
    line: str = "CONSOLIDADO",
) -> pd.DataFrame:
    """Faturamento e metas pelo gerente responsável por cada linha."""
    months = pd.period_range(start, end, freq="M")
    out = pd.DataFrame({"Mês": months})

    fbase = fat.copy()
    mbase = metas_g.copy()
    if line != "CONSOLIDADO":
        fbase = billing_by_line(fbase, line)
        manager = LINE_MANAGER_MAP.get(line, "")
        if "GERENTE" in mbase.columns:
            mbase = mbase[mbase["GERENTE"].map(norm) == manager]

    actual = (
        period_filter(fbase, start, end).groupby("_MES", as_index=False)["_VALOR"].sum()
        .rename(columns={"_MES": "Mês", "_VALOR": "Faturamento"})
    )
    target = (
        period_filter(mbase, start, end).groupby("_MES", as_index=False)["_META"].sum()
        .rename(columns={"_MES": "Mês", "_META": "Meta"})
    )
    out = out.merge(actual, on="Mês", how="left").merge(target, on="Mês", how="left").fillna(0)
    out["Desvio"] = out["Faturamento"] - out["Meta"]
    out["Atingimento"] = np.where(out["Meta"] != 0, out["Faturamento"] / out["Meta"], 0)
    out["Mês Texto"] = out["Mês"].map(month_label)
    return out


def commercial_performance_totals(
    monthly: pd.DataFrame, metas_g: pd.DataFrame, line: str, reference_year: int,
) -> dict[str, float]:
    actual = float(monthly["Faturamento"].sum())
    target = float(monthly["Meta"].sum())
    month_count = max(int(monthly["Mês"].nunique()), 1)

    annual = metas_g[metas_g["_MES"].dt.year == reference_year].copy()
    if line != "CONSOLIDADO" and "GERENTE" in annual.columns:
        annual = annual[annual["GERENTE"].map(norm) == LINE_MANAGER_MAP.get(line, "")]
    if "_META_ANUAL" in annual.columns and float(annual["_META_ANUAL"].sum()) > 0:
        if "GERENTE" in annual.columns:
            annual_target = float(annual.assign(_GERENTE_KEY=annual["GERENTE"].map(norm)).groupby("_GERENTE_KEY")["_META_ANUAL"].max().sum())
        else:
            annual_target = float(annual["_META_ANUAL"].max())
    else:
        annual_target = float(annual["_META"].sum())
    monthly_average = actual / month_count
    annual_projection = monthly_average * 12
    return {
        "Faturamento": actual,
        "Meta": target,
        "Desvio": actual - target,
        "Atingimento": safe_div(actual, target),
        "Média Mensal": monthly_average,
        "Meta Anual": annual_target,
        "Projeção Anual": annual_projection,
        "Atingimento Projetado": safe_div(annual_projection, annual_target),
    }


def seller_performance(
    fat: pd.DataFrame, metas: pd.DataFrame, start: pd.Period, end: pd.Period,
    line: str = "CONSOLIDADO",
) -> pd.DataFrame:
    """Apura equipe, meta e realizado diretamente pela aba Metas.

    A aba Metas é a fonte oficial tanto para a composição da equipe quanto para
    o valor realizado individual. Na estrutura atual da BASE BI, a coluna
    ATINGIMENTO contém o realizado mensal e a coluna % contém o percentual.

    O Banco de Dados de Faturamento permanece como fonte do faturamento total
    da empresa e das linhas, mas não é usado para redistribuir o resultado entre
    vendedores, evitando zerar nomes que possuem cadastros diferentes nas bases.
    """
    fbase = period_filter(fat, start, end).copy()
    mbase = period_filter(metas, start, end).copy()

    if line != "CONSOLIDADO":
        fbase = billing_by_line(fbase, line)
        manager = LINE_MANAGER_MAP.get(line, "")
        if "GERENTE" in mbase.columns:
            mbase = mbase[mbase["GERENTE"].map(norm) == manager]

    if mbase.empty or "VENDEDOR" not in mbase.columns:
        empty = pd.DataFrame(columns=["Vendedor", "Faturamento", "Meta", "Desvio", "Atingimento", "Status"])
        empty.attrs["billing_total"] = float(fbase["_VALOR"].sum()) if "_VALOR" in fbase.columns else 0.0
        empty.attrs["meta_actual_total"] = 0.0
        empty.attrs["reconciliation_difference"] = empty.attrs["billing_total"]
        return empty

    mbase["_ROSTER_KEY"] = mbase["VENDEDOR"].map(seller_key)
    mbase = mbase[~mbase["_ROSTER_KEY"].isin(["", "NAO INFORMADO"])].copy()
    if "_REALIZADO_META" not in mbase.columns:
        mbase["_REALIZADO_META"] = 0.0

    out = (
        mbase.groupby("_ROSTER_KEY", as_index=False)
        .agg(
            Vendedor=("VENDEDOR", "first"),
            Faturamento=("_REALIZADO_META", "sum"),
            Meta=("_META", "sum"),
        )
    )

    out["Faturamento"] = out["Faturamento"].fillna(0.0)
    out["Meta"] = out["Meta"].fillna(0.0)
    out["Desvio"] = out["Faturamento"] - out["Meta"]
    out["Atingimento"] = np.where(out["Meta"] > 0, out["Faturamento"] / out["Meta"], 0.0)
    out["Status"] = np.select(
        [out["Meta"] <= 0, out["Atingimento"] >= 1, out["Atingimento"] >= .9],
        ["Meta não cadastrada", "Meta atingida", "Próximo da meta"],
        default="Abaixo da meta",
    )

    out = out.sort_values("Faturamento", ascending=False).reset_index(drop=True)
    billing_total = float(fbase["_VALOR"].sum()) if "_VALOR" in fbase.columns else 0.0
    meta_actual_total = float(out["Faturamento"].sum()) if not out.empty else 0.0
    out.attrs["billing_total"] = billing_total
    out.attrs["meta_actual_total"] = meta_actual_total
    out.attrs["reconciliation_difference"] = billing_total - meta_actual_total
    out.attrs["source"] = "Equipe, meta e realizado definidos pela aba Metas"
    return out


def line_cash_monthly(
    line: str, fat: pd.DataFrame, receitas: pd.DataFrame, custos: pd.DataFrame,
    start: pd.Period, end: pd.Period,
) -> pd.DataFrame:
    """Resultado direto por linha, integralmente apurado na aba Centro de Custos.

    Receita: EMPRESA = RECEITA, GRUPO = Receitas Operacionais e centro de custo direto da linha.
    Custo: EMPRESA = DESPESAS, GRUPO = Saídas Operacionais e centro de custo direto da linha.
    O campo CENTRO DE CUSTOS RATEAO não é utilizado.
    """
    months = pd.period_range(start, end, freq="M")
    out = pd.DataFrame({"Mês": months})

    f = (
        period_filter(billing_by_line(fat, line), start, end)
        .groupby("_MES", as_index=False)["_VALOR"].sum()
        .rename(columns={"_MES": "Mês", "_VALOR": "Faturamento"})
    )

    direct = custos[custos["_LINHA_DIRETA"] == line].copy()
    rbase = direct[
        (direct["_EMPRESA_N"] == "RECEITA") &
        (direct["_GRUPO_N"] == "RECEITAS OPERACIONAIS")
    ].copy()
    cbase = direct[
        (direct["_EMPRESA_N"] == "DESPESAS") &
        (direct["_GRUPO_N"] == "SAIDAS OPERACIONAIS")
    ].copy()
    cbase["_CLASSE_CAIXA"] = cbase["PAI"].map(cost_class)

    r = (
        period_filter(rbase, start, end).groupby("_MES", as_index=False)["_VALOR"].sum()
        .rename(columns={"_MES": "Mês", "_VALOR": "Receitas Recebidas"})
    )
    c = (
        period_filter(cbase, start, end).groupby("_MES", as_index=False)["_VALOR"].sum()
        .rename(columns={"_MES": "Mês", "_VALOR": "Custos Diretos Pagos"})
    )
    v = (
        period_filter(cbase[cbase["_CLASSE_CAIXA"] == "VARIAVEL"], start, end)
        .groupby("_MES", as_index=False)["_VALOR"].sum()
        .rename(columns={"_MES": "Mês", "_VALOR": "Custos Diretos Variáveis"})
    )

    for df in [f, r, c, v]:
        out = out.merge(df, on="Mês", how="left")
    out = out.fillna(0)
    out["Resultado Direto de Caixa"] = out["Receitas Recebidas"] - out["Custos Diretos Pagos"]
    out["Margem Direta de Caixa"] = np.where(
        out["Receitas Recebidas"] != 0,
        out["Resultado Direto de Caixa"] / out["Receitas Recebidas"], 0,
    )
    out["Contribuição Direta de Caixa"] = out["Receitas Recebidas"] - out["Custos Diretos Variáveis"]
    out["Margem de Contribuição Direta"] = np.where(
        out["Receitas Recebidas"] != 0,
        out["Contribuição Direta de Caixa"] / out["Receitas Recebidas"], 0,
    )
    out["Conversão em Caixa"] = np.where(
        out["Faturamento"] != 0, out["Receitas Recebidas"] / out["Faturamento"], 0,
    )
    out["Mês Texto"] = out["Mês"].map(month_label)
    return out


def totals_from_monthly(monthly: pd.DataFrame, line_mode: bool = False) -> dict[str, float]:
    sums = monthly.select_dtypes(include=[np.number]).sum().to_dict()
    if line_mode:
        sums["Margem Direta de Caixa"] = safe_div(sums.get("Resultado Direto de Caixa", 0), sums.get("Receitas Recebidas", 0))
        sums["Margem de Contribuição Direta"] = safe_div(sums.get("Contribuição Direta de Caixa", 0), sums.get("Receitas Recebidas", 0))
        sums["Conversão em Caixa"] = safe_div(sums.get("Receitas Recebidas", 0), sums.get("Faturamento", 0))
    else:
        sums["Margem de Caixa"] = safe_div(sums.get("Resultado Operacional de Caixa", 0), sums.get("Receitas Operacionais", 0))
        sums["Margem EBITDA Caixa"] = safe_div(sums.get("EBITDA Gerencial de Caixa", 0), sums.get("Receitas Operacionais", 0))
        sums["Margem de Contribuição Caixa"] = safe_div(sums.get("Contribuição de Caixa", 0), sums.get("Receitas Operacionais", 0))
        sums["Conversão em Caixa"] = safe_div(sums.get("Receitas Operacionais", 0), sums.get("Faturamento", 0))
        sums["Performance de Recebimento"] = safe_div(sums.get("Recebimento Realizado", 0), sums.get("Recebimento Previsto", 0))
        sums["Atingimento da Meta"] = safe_div(sums.get("Faturamento", 0), sums.get("Meta", 0))
    return sums


def line_summary(fat: pd.DataFrame, metas_g: pd.DataFrame, receitas: pd.DataFrame, custos: pd.DataFrame, start: pd.Period, end: pd.Period, inad: pd.DataFrame | None) -> pd.DataFrame:
    rows = []
    for line in LINES:
        monthly = line_cash_monthly(line, fat, receitas, custos, start, end)
        t = totals_from_monthly(monthly, True)
        perf_monthly = commercial_performance_monthly(fat, metas_g, start, end, line)
        perf = commercial_performance_totals(perf_monthly, metas_g, line, end.year)
        direct = period_filter(custos[custos["_LINHA_DIRETA"] == line], start, end)
        revenue_rows = direct[
            (direct["_EMPRESA_N"] == "RECEITA") &
            (direct["_GRUPO_N"] == "RECEITAS OPERACIONAIS")
        ]
        cost_rows = direct[
            (direct["_EMPRESA_N"] == "DESPESAS") &
            (direct["_GRUPO_N"] == "SAIDAS OPERACIONAIS")
        ]
        overdue = 0.0
        if inad is not None and not inad.empty:
            overdue = float(inad[inad["_LINHA"] == line]["_VALOR_VENCIDO"].sum())
        rows.append({
            "Linha": line_label(line), "Código": line,
            "Receitas Recebidas": t.get("Receitas Recebidas", 0),
            "Custos Diretos Pagos": t.get("Custos Diretos Pagos", 0),
            "Resultado Direto de Caixa": t.get("Resultado Direto de Caixa", 0),
            "Margem Direta de Caixa": t.get("Margem Direta de Caixa", 0),
            "Margem de Contribuição Direta": t.get("Margem de Contribuição Direta", 0),
            "Faturamento": perf.get("Faturamento", t.get("Faturamento", 0)),
            "Meta": perf.get("Meta", 0),
            "Atingimento da Meta": perf.get("Atingimento", 0),
            "Desvio da Meta": perf.get("Desvio", 0),
            "Projeção Anual": perf.get("Projeção Anual", 0),
            "Atingimento Projetado": perf.get("Atingimento Projetado", 0),
            "Conversão em Caixa": t.get("Conversão em Caixa", 0),
            "Lançamentos de Receita": int(len(revenue_rows)),
            "Lançamentos de Custo": int(len(cost_rows)),
            "Inadimplência": overdue,
        })
    return pd.DataFrame(rows)


def scoped_data(scope: str, fat: pd.DataFrame, receitas: pd.DataFrame, custos: pd.DataFrame, inad: pd.DataFrame | None, start: pd.Period, end: pd.Period):
    if scope == "CONSOLIDADO":
        f = period_filter(fat, start, end)
        r = period_filter(receitas[~receitas["_NAO_OPERACIONAL"]], start, end)
        c = period_filter(custos[(custos["_GRUPO_N"] == "SAIDAS OPERACIONAIS") & (custos["_LINHA_DIRETA"].isin(LINES))], start, end)
        i = inad.copy() if inad is not None else None
    else:
        f = period_filter(billing_by_line(fat, scope), start, end)
        r = period_filter(receitas[(receitas["_LINHA"] == scope) & (~receitas["_NAO_OPERACIONAL"])], start, end)
        c = period_filter(custos[(custos["_LINHA_DIRETA"] == scope) & (custos["_GRUPO_N"] == "SAIDAS OPERACIONAIS")], start, end)
        i = inad[inad["_LINHA"] == scope].copy() if inad is not None else None
    return f, r, c, i

def build_business_performance_pdf(
    *,
    scope_choice: str,
    user: dict[str, str],
    start_month: pd.Period,
    end_month: pd.Period,
    fat: pd.DataFrame,
    metas: pd.DataFrame,
    metas_g: pd.DataFrame,
    receitas: pd.DataFrame,
    despesas: pd.DataFrame,
    custos: pd.DataFrame,
    performance: pd.DataFrame,
    inad: pd.DataFrame | None,
    company_monthly: pd.DataFrame,
    company_totals: dict[str, float],
    line_monthly: pd.DataFrame | None,
    line_totals: dict[str, float] | None,
    commercial_monthly: pd.DataFrame,
    commercial_totals: dict[str, float],
    lines_table: pd.DataFrame,
    fat_scope: pd.DataFrame,
    rec_scope: pd.DataFrame,
    cost_scope: pd.DataFrame,
    inad_scope: pd.DataFrame | None,
) -> bytes:
    """Gera um relatório PDF paginado com o compilado de todas as áreas do app."""
    try:
        from reportlab.graphics.shapes import Drawing, Rect, String, Line
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.lib.units import mm
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle,
            KeepTogether, HRFlowable,
        )
    except ImportError as exc:
        raise RuntimeError("A biblioteca reportlab não está instalada. Inclua reportlab no requirements.txt.") from exc

    from datetime import datetime

    navy = colors.HexColor(NAVY)
    navy_2 = colors.HexColor(NAVY_2)
    blue = colors.HexColor(BLUE)
    cyan = colors.HexColor(CYAN)
    red = colors.HexColor(RED)
    gray = colors.HexColor(GRAY)
    light = colors.HexColor("#F4F7FA")
    border = colors.HexColor("#DCE6EF")
    white = colors.white

    def txt(value: object) -> str:
        value = "" if value is None else str(value)
        return (
            value.replace("–", "-").replace("—", "-").replace("•", "-")
            .replace(" ", " ").replace(" ", " ")
        )

    def money(value: float) -> str:
        return brl(float(value or 0))

    def percent(value: float) -> str:
        return pct(float(value or 0))

    def number(value: float, decimals: int = 0) -> str:
        return f"{float(value or 0):,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")

    output = io.BytesIO()
    page_size = landscape(A4)
    doc = SimpleDocTemplate(
        output,
        pagesize=page_size,
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=15 * mm,
        bottomMargin=13 * mm,
        title="First Intelligence | Business Performance",
        author="First Medical - Controladoria",
        subject=f"Relatório gerencial de {month_label(start_month)} a {month_label(end_month)}",
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="ReportTitle", parent=styles["Title"], fontName="Helvetica-Bold",
        fontSize=25, leading=29, textColor=navy, spaceAfter=8,
    ))
    styles.add(ParagraphStyle(
        name="ReportSubtitle", parent=styles["Normal"], fontName="Helvetica",
        fontSize=10, leading=14, textColor=gray, spaceAfter=5,
    ))
    styles.add(ParagraphStyle(
        name="SectionTitle", parent=styles["Heading1"], fontName="Helvetica-Bold",
        fontSize=16, leading=19, textColor=navy, spaceAfter=6,
    ))
    styles.add(ParagraphStyle(
        name="SectionNote", parent=styles["Normal"], fontName="Helvetica",
        fontSize=8.5, leading=12, textColor=gray, spaceAfter=9,
    ))
    styles.add(ParagraphStyle(
        name="BodySmall", parent=styles["Normal"], fontName="Helvetica",
        fontSize=8, leading=11, textColor=colors.HexColor("#344054"),
    ))
    styles.add(ParagraphStyle(
        name="BodyTiny", parent=styles["Normal"], fontName="Helvetica",
        fontSize=7, leading=9, textColor=colors.HexColor("#344054"),
    ))
    styles.add(ParagraphStyle(
        name="MetricLabel", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=6.7, leading=8, textColor=gray, spaceAfter=4,
    ))
    styles.add(ParagraphStyle(
        name="MetricValue", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=13.2, leading=15, textColor=navy,
    ))
    styles.add(ParagraphStyle(
        name="TableHead", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=7, leading=8.5, textColor=white, alignment=TA_LEFT,
    ))
    styles.add(ParagraphStyle(
        name="TableCell", parent=styles["Normal"], fontName="Helvetica",
        fontSize=6.8, leading=8.4, textColor=colors.HexColor("#344054"),
    ))
    styles.add(ParagraphStyle(
        name="TableCellRight", parent=styles["TableCell"], alignment=TA_RIGHT,
    ))

    def P(value: object, style: str = "BodySmall") -> Paragraph:
        safe = html.escape(txt(value)).replace("\n", "<br/>")
        return Paragraph(safe, styles[style])

    def page_header_footer(canvas, _doc):
        canvas.saveState()
        w, h = page_size
        canvas.setFillColor(navy)
        canvas.setFont("Helvetica-Bold", 8)
        canvas.drawString(doc.leftMargin, h - 9 * mm, "FIRST INTELLIGENCE | BUSINESS PERFORMANCE")
        canvas.setStrokeColor(border)
        canvas.setLineWidth(.5)
        canvas.line(doc.leftMargin, h - 10.5 * mm, w - doc.rightMargin, h - 10.5 * mm)
        canvas.setFillColor(gray)
        canvas.setFont("Helvetica", 7)
        footer = f"{line_label(scope_choice)} | {month_label(start_month)} a {month_label(end_month)}"
        canvas.drawString(doc.leftMargin, 7 * mm, footer)
        canvas.drawRightString(w - doc.rightMargin, 7 * mm, f"Página {canvas.getPageNumber()}")
        canvas.restoreState()

    def section(title: str, note: str = "") -> list:
        elements = [P(title, "SectionTitle")]
        if note:
            elements.append(P(note, "SectionNote"))
        elements.append(HRFlowable(width="100%", thickness=.7, color=border, spaceAfter=8))
        return elements

    def metric_grid(items: list[tuple[str, str]], columns: int = 4) -> Table:
        cells = []
        for label, value in items:
            cells.append([P(label.upper(), "MetricLabel"), Spacer(1, 2), P(value, "MetricValue")])
        while len(cells) % columns:
            cells.append([P("", "MetricLabel"), P("", "MetricValue")])
        rows = [cells[i:i + columns] for i in range(0, len(cells), columns)]
        table = Table(rows, colWidths=[doc.width / columns] * columns, hAlign="LEFT")
        style = [
            ("BACKGROUND", (0, 0), (-1, -1), white),
            ("BOX", (0, 0), (-1, -1), .65, border),
            ("INNERGRID", (0, 0), (-1, -1), .35, border),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 9),
            ("RIGHTPADDING", (0, 0), (-1, -1), 9),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ]
        table.setStyle(TableStyle(style))
        return table

    def data_table(
        headers: list[str], rows: list[list[object]], widths: list[float] | None = None,
        right_cols: set[int] | None = None, max_rows: int | None = None,
    ) -> Table:
        right_cols = right_cols or set()
        if max_rows is not None:
            rows = rows[:max_rows]
        if widths is None:
            widths = [doc.width / max(len(headers), 1)] * len(headers)
        total_width = sum(widths)
        if total_width > doc.width:
            scale = doc.width / total_width
            widths = [w * scale for w in widths]
        content = [[P(h, "TableHead") for h in headers]]
        for row in rows:
            formatted = []
            for idx, value in enumerate(row):
                formatted.append(P(value, "TableCellRight" if idx in right_cols else "TableCell"))
            content.append(formatted)
        table = Table(content, colWidths=widths, repeatRows=1, hAlign="LEFT", splitByRow=True)
        commands = [
            ("BACKGROUND", (0, 0), (-1, 0), navy_2),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, 0), 6),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
            ("TOPPADDING", (0, 1), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
            ("LINEBELOW", (0, 0), (-1, 0), .7, navy),
        ]
        for row_idx in range(1, len(content)):
            if row_idx % 2 == 0:
                commands.append(("BACKGROUND", (0, row_idx), (-1, row_idx), light))
            commands.append(("LINEBELOW", (0, row_idx), (-1, row_idx), .25, border))
        for col_idx in right_cols:
            commands.append(("ALIGN", (col_idx, 1), (col_idx, -1), "RIGHT"))
        table.setStyle(TableStyle(commands))
        return table

    def vbar_chart(
        categories: list[str], series: list[tuple[str, list[float], object]], title: str,
        width: float | None = None, height: float = 205,
    ) -> Drawing:
        width = float(width or doc.width)
        d = Drawing(width, height)
        d.add(String(0, height - 13, txt(title), fontName="Helvetica-Bold", fontSize=10.5, fillColor=navy))
        if not categories or not series:
            d.add(String(0, height / 2, "Sem dados para o período.", fontName="Helvetica", fontSize=8, fillColor=gray))
            return d
        values = [max(float(v), 0.0) for _, vals, _ in series for v in vals]
        max_value = max(values) if values else 0
        if max_value <= 0:
            d.add(String(0, height / 2, "Sem valores positivos para exibição.", fontName="Helvetica", fontSize=8, fillColor=gray))
            return d
        left = 18
        bottom = 34
        chart_width = width - 28
        chart_height = height - 67
        group_width = chart_width / max(len(categories), 1)
        series_count = len(series)
        bar_width = min(24, max(7, group_width / (series_count + 1.5)))
        for idx, category in enumerate(categories):
            center = left + group_width * (idx + .5)
            d.add(String(center, 9, txt(short_label(category, 15)), textAnchor="middle", fontName="Helvetica", fontSize=6.5, fillColor=gray))
            for s_idx, (_name, vals, color) in enumerate(series):
                value = max(float(vals[idx]), 0.0) if idx < len(vals) else 0.0
                bar_h = chart_height * value / max_value
                x = center - (series_count * bar_width) / 2 + s_idx * bar_width
                d.add(Rect(x, bottom, bar_width - 2, bar_h, fillColor=color, strokeColor=None))
                if value > 0:
                    d.add(String(x + (bar_width - 2) / 2, bottom + bar_h + 3, compact_money(value), textAnchor="middle", fontName="Helvetica", fontSize=5.8, fillColor=navy))
        legend_x = max(0, width - 105 * series_count)
        for idx, (name, _vals, color) in enumerate(series):
            lx = legend_x + idx * 105
            d.add(Rect(lx, height - 14, 7, 7, fillColor=color, strokeColor=None))
            d.add(String(lx + 11, height - 13, txt(name), fontName="Helvetica", fontSize=6.6, fillColor=gray))
        return d

    def hbar_chart(
        categories: list[str], series: list[tuple[str, list[float], object]], title: str,
        width: float | None = None, row_height: float = 26,
    ) -> Drawing:
        width = float(width or doc.width)
        count = max(len(categories), 1)
        height = 46 + count * row_height
        d = Drawing(width, height)
        d.add(String(0, height - 13, txt(title), fontName="Helvetica-Bold", fontSize=10.5, fillColor=navy))
        if not categories or not series:
            d.add(String(0, height / 2, "Sem dados para o período.", fontName="Helvetica", fontSize=8, fillColor=gray))
            return d
        max_value = max([max(float(v), 0.0) for _n, vals, _c in series for v in vals] or [0])
        if max_value <= 0:
            return d
        label_width = 150
        plot_width = width - label_width - 60
        series_count = len(series)
        bar_h = min(8, (row_height - 5) / max(series_count, 1))
        for idx, category in enumerate(categories):
            y_top = height - 34 - idx * row_height
            d.add(String(label_width - 7, y_top - 3, txt(short_label(category, 31)), textAnchor="end", fontName="Helvetica", fontSize=6.8, fillColor=gray))
            for s_idx, (_name, vals, color) in enumerate(series):
                value = max(float(vals[idx]), 0.0) if idx < len(vals) else 0.0
                y = y_top - 7 - s_idx * (bar_h + 3)
                bar_w = plot_width * value / max_value
                d.add(Rect(label_width, y, bar_w, bar_h, fillColor=color, strokeColor=None))
                if value > 0:
                    d.add(String(label_width + bar_w + 4, y + 1, compact_money(value), fontName="Helvetica", fontSize=6.2, fillColor=navy if color != red else red))
        legend_x = max(0, width - 105 * series_count)
        for idx, (name, _vals, color) in enumerate(series):
            lx = legend_x + idx * 105
            d.add(Rect(lx, height - 14, 7, 7, fillColor=color, strokeColor=None))
            d.add(String(lx + 11, height - 13, txt(name), fontName="Helvetica", fontSize=6.6, fillColor=gray))
        return d

    story: list = []
    period_text = f"{month_label(start_month)} a {month_label(end_month)}"
    scope_text = line_label(scope_choice)
    generated_at = datetime.now().strftime("%d/%m/%Y %H:%M")

    # Capa e visão executiva
    story.extend([
        Spacer(1, 10 * mm),
        P("FIRST INTELLIGENCE", "ReportSubtitle"),
        P("Business Performance", "ReportTitle"),
        P(f"Relatório gerencial completo | {scope_text} | {period_text}", "ReportSubtitle"),
        P(f"Perfil: {user.get('nome', 'Usuário')} | Gerado em {generated_at}", "ReportSubtitle"),
        Spacer(1, 7 * mm),
    ])

    overdue_total = float(inad_scope["_VALOR_VENCIDO"].sum()) if inad_scope is not None and not inad_scope.empty else 0.0
    if scope_choice == "CONSOLIDADO":
        cover_metrics = [
            ("Receitas operacionais recebidas", money(company_totals.get("Receitas Operacionais", 0))),
            ("Saídas operacionais pagas", money(company_totals.get("Saídas Operacionais", 0))),
            ("Resultado operacional de caixa", money(company_totals.get("Resultado Operacional de Caixa", 0))),
            ("Margem operacional de caixa", percent(company_totals.get("Margem de Caixa", 0))),
            ("Faturamento realizado", money(commercial_totals.get("Faturamento", 0))),
            ("Atingimento da meta", percent(commercial_totals.get("Atingimento", 0))),
            ("EBITDA gerencial de caixa", money(company_totals.get("EBITDA Gerencial de Caixa", 0))),
            ("Inadimplência", money(overdue_total)),
        ]
    else:
        lt = line_totals or {}
        cover_metrics = [
            ("Receitas diretas recebidas", money(lt.get("Receitas Recebidas", 0))),
            ("Custos diretos pagos", money(lt.get("Custos Diretos Pagos", 0))),
            ("Resultado direto de caixa", money(lt.get("Resultado Direto de Caixa", 0))),
            ("Margem direta de caixa", percent(lt.get("Margem Direta de Caixa", 0))),
            ("Faturamento realizado", money(commercial_totals.get("Faturamento", 0))),
            ("Atingimento da meta", percent(commercial_totals.get("Atingimento", 0))),
            ("Conversão em caixa", percent(lt.get("Conversão em Caixa", 0))),
            ("Inadimplência", money(overdue_total)),
        ]
    story.append(metric_grid(cover_metrics, 4))
    story.append(Spacer(1, 8 * mm))
    story.append(P(
        "O relatório reúne, em um único arquivo, Dashboard, Desempenho e Metas, Linhas de Negócio, "
        "Recebimentos e Inadimplência, Clientes, Produtos e Centro de Custos. Os cálculos respeitam o "
        "período, o escopo e as permissões do perfil selecionado no aplicativo.",
        "SectionNote",
    ))

    # 1 - Dashboard
    story.append(PageBreak())
    story.extend(section("1. Dashboard", "Visão integrada dos principais indicadores e da evolução mensal."))
    monthly = company_monthly if scope_choice == "CONSOLIDADO" else (line_monthly if line_monthly is not None else pd.DataFrame())
    if scope_choice == "CONSOLIDADO":
        rev_col, exp_col, result_col = "Receitas Operacionais", "Saídas Operacionais", "Resultado Operacional de Caixa"
    else:
        rev_col, exp_col, result_col = "Receitas Recebidas", "Custos Diretos Pagos", "Resultado Direto de Caixa"
    if not monthly.empty:
        story.append(vbar_chart(
            monthly["Mês Texto"].tolist(),
            [("Receitas", monthly[rev_col].tolist(), blue), ("Despesas", monthly[exp_col].tolist(), red)],
            "Receitas e despesas por mês",
        ))
        dashboard_rows = []
        for _, row in monthly.iterrows():
            dashboard_rows.append([
                row["Mês Texto"], money(row.get(rev_col, 0)), money(row.get(exp_col, 0)),
                money(row.get(result_col, 0)), percent(row.get("Conversão em Caixa", 0)),
            ])
        story.append(data_table(
            ["Mês", "Receitas", "Despesas", "Resultado", "Conversão"], dashboard_rows,
            [65, 125, 125, 125, 90], {1, 2, 3, 4},
        ))

    # 2 - Desempenho e metas
    story.append(PageBreak())
    story.extend(section("2. Desempenho e metas", "Realizado comercial, meta, desvio e projeção anual."))
    performance_metrics = [
        ("Faturamento realizado", money(commercial_totals.get("Faturamento", 0))),
        ("Meta acumulada", money(commercial_totals.get("Meta", 0))),
        ("Atingimento", percent(commercial_totals.get("Atingimento", 0))),
        ("Desvio", money(commercial_totals.get("Desvio", 0))),
        ("Média mensal", money(commercial_totals.get("Média Mensal", 0))),
        ("Meta anual", money(commercial_totals.get("Meta Anual", 0))),
        ("Projeção anual", money(commercial_totals.get("Projeção Anual", 0))),
        ("Atingimento projetado", percent(commercial_totals.get("Atingimento Projetado", 0))),
    ]
    story.append(metric_grid(performance_metrics, 4))
    story.append(Spacer(1, 5 * mm))
    if not commercial_monthly.empty:
        story.append(vbar_chart(
            commercial_monthly["Mês Texto"].tolist(),
            [("Faturamento", commercial_monthly["Faturamento"].tolist(), blue), ("Meta", commercial_monthly["Meta"].tolist(), navy)],
            "Faturamento x meta mensal",
        ))
        perf_rows = [[
            r["Mês Texto"], money(r["Faturamento"]), money(r["Meta"]), money(r["Desvio"]), percent(r["Atingimento"])
        ] for _, r in commercial_monthly.iterrows()]
        story.append(data_table(
            ["Mês", "Faturamento", "Meta", "Desvio", "Atingimento"], perf_rows,
            [65, 130, 130, 130, 90], {1, 2, 3, 4},
        ))

    sellers = seller_performance(fat, metas, start_month, end_month, scope_choice).head(15)
    if not sellers.empty:
        story.append(Spacer(1, 5 * mm))
        story.append(P("Desempenho da equipe", "SectionTitle"))
        seller_rows = [[
            r["Vendedor"], money(r["Faturamento"]), money(r["Meta"]), percent(r["Atingimento"]), money(r["Desvio"]), r["Status"]
        ] for _, r in sellers.iterrows()]
        story.append(data_table(
            ["Vendedor / representante", "Faturamento", "Meta", "Atingimento", "Desvio", "Status"], seller_rows,
            [200, 105, 105, 80, 105, 100], {1, 2, 3, 4},
        ))

    # 3 - Linhas de negócio / Resultado da linha
    story.append(PageBreak())
    if scope_choice == "CONSOLIDADO":
        story.extend(section("3. Linhas de negócio", "Receitas e custos diretamente identificados na coluna Centro de Custos, sem rateio."))
        line_rows = []
        for _, r in lines_table.iterrows():
            line_rows.append([
                r["Linha"], money(r["Receitas Recebidas"]), money(r["Custos Diretos Pagos"]),
                money(r["Resultado Direto de Caixa"]), percent(r["Margem Direta de Caixa"]),
                money(r["Faturamento"]), percent(r["Atingimento da Meta"]), money(r["Inadimplência"]),
            ])
        story.append(vbar_chart(
            lines_table["Linha"].tolist(),
            [("Receitas", lines_table["Receitas Recebidas"].tolist(), blue), ("Despesas", lines_table["Custos Diretos Pagos"].tolist(), red)],
            "Receitas x despesas por linha",
        ))
        story.append(data_table(
            ["Linha", "Receitas", "Despesas", "Resultado", "Margem", "Faturamento", "Meta %", "Inadimplência"],
            line_rows, [82, 100, 100, 100, 60, 100, 60, 100], {1, 2, 3, 4, 5, 6, 7},
        ))
    else:
        story.extend(section("3. Resultado da linha", "Visão mensal restrita à linha do perfil selecionado."))
        if line_monthly is not None and not line_monthly.empty:
            line_rows = [[
                r["Mês Texto"], money(r["Faturamento"]), money(r["Receitas Recebidas"]), money(r["Custos Diretos Pagos"]),
                money(r["Resultado Direto de Caixa"]), percent(r["Margem Direta de Caixa"]), percent(r["Conversão em Caixa"]),
            ] for _, r in line_monthly.iterrows()]
            story.append(data_table(
                ["Mês", "Faturamento", "Receitas", "Custos", "Resultado", "Margem", "Conversão"], line_rows,
                [60, 105, 105, 105, 105, 72, 72], {1, 2, 3, 4, 5, 6},
            ))

    # 4 - Recebimentos e inadimplência
    story.append(PageBreak())
    story.extend(section("4. Recebimentos e inadimplência", "Conversão financeira, aging e maiores saldos vencidos."))
    if scope_choice == "CONSOLIDADO":
        receipt_metrics = [
            ("Recebimento previsto", money(company_totals.get("Recebimento Previsto", 0))),
            ("Recebimento realizado", money(company_totals.get("Recebimento Realizado", 0))),
            ("Performance de recebimento", percent(company_totals.get("Performance de Recebimento", 0))),
            ("Inadimplência", money(overdue_total)),
        ]
    else:
        receipt_metrics = [
            ("Receitas recebidas", money((line_totals or {}).get("Receitas Recebidas", 0))),
            ("Faturamento", money(commercial_totals.get("Faturamento", 0))),
            ("Conversão em caixa", percent((line_totals or {}).get("Conversão em Caixa", 0))),
            ("Inadimplência", money(overdue_total)),
        ]
    story.append(metric_grid(receipt_metrics, 4))
    story.append(Spacer(1, 5 * mm))
    if inad_scope is not None and not inad_scope.empty:
        aging_order = ["Até 30 dias", "31 a 60 dias", "61 a 90 dias", "Acima de 90 dias"]
        aging = inad_scope.groupby("_FAIXA", observed=False)["_VALOR_VENCIDO"].sum().reindex(aging_order, fill_value=0)
        story.append(vbar_chart(
            aging_order, [("Saldo vencido", aging.tolist(), blue)], "Aging da inadimplência", height=190,
        ))
        aging_rows = [[label, money(value), percent(safe_div(value, overdue_total))] for label, value in aging.items()]
        story.append(data_table(["Faixa", "Saldo vencido", "Participação"], aging_rows, [180, 140, 100], {1, 2}))
        top_debtors = (
            inad_scope.groupby("_CLIENTE", as_index=False)["_VALOR_VENCIDO"].sum()
            .sort_values("_VALOR_VENCIDO", ascending=False).head(15)
        )
        debtor_rows = [[r["_CLIENTE"], money(r["_VALOR_VENCIDO"]), percent(safe_div(r["_VALOR_VENCIDO"], overdue_total))] for _, r in top_debtors.iterrows()]
        story.append(Spacer(1, 5 * mm))
        story.append(P("Maiores saldos vencidos", "SectionTitle"))
        story.append(data_table(["Cliente", "Saldo vencido", "Participação"], debtor_rows, [350, 130, 100], {1, 2}))
    else:
        story.append(P("A base de inadimplência não possui registros para o período e escopo selecionados.", "SectionNote"))

    # 5 - Clientes
    story.append(PageBreak())
    story.extend(section("5. Clientes", "Principais clientes por faturamento e recebimentos."))
    billing_client_col = "NOME DO CLIENTE" if "NOME DO CLIENTE" in fat_scope.columns else optional_col(fat_scope, ["CLIENTE", "Cliente", "RAZÃO SOCIAL", "Razao Social"])
    receipt_client_col = "Cliente" if "Cliente" in rec_scope.columns else optional_col(rec_scope, ["CLIENTE", "Nome do Cliente"])
    if billing_client_col and not fat_scope.empty:
        top_billing = (
            fat_scope.groupby(billing_client_col, as_index=False)["_VALOR"].sum()
            .sort_values("_VALOR", ascending=False).head(15)
        )
        total_billing = float(fat_scope["_VALOR"].sum())
        rows = [[r[billing_client_col], money(r["_VALOR"]), percent(safe_div(r["_VALOR"], total_billing))] for _, r in top_billing.iterrows()]
        story.append(P("Clientes por faturamento", "SectionTitle"))
        story.append(data_table(["Cliente", "Faturamento", "Participação"], rows, [350, 130, 100], {1, 2}))
    if receipt_client_col and not rec_scope.empty:
        story.append(Spacer(1, 5 * mm))
        top_receipts = (
            rec_scope.groupby(receipt_client_col, as_index=False)["_VALOR"].sum()
            .sort_values("_VALOR", ascending=False).head(15)
        )
        total_receipts = float(rec_scope["_VALOR"].sum())
        rows = [[r[receipt_client_col], money(r["_VALOR"]), percent(safe_div(r["_VALOR"], total_receipts))] for _, r in top_receipts.iterrows()]
        story.append(P("Clientes por recebimento", "SectionTitle"))
        story.append(data_table(["Cliente", "Recebimentos", "Participação"], rows, [350, 130, 100], {1, 2}))

    # 6 - Produtos
    story.append(PageBreak())
    story.extend(section("6. Produtos", "Faturamento, volume, preço médio e concentração."))
    product_col = optional_col(fat_scope, ["PRODUTO", "DESCRIÇÃO", "LINHA DE PRODUTO", "ITEM"])
    qty_col = optional_col(fat_scope, ["QUANTIDADE", "QTD", "QTDE", "QTD FATURADA", "QUANTIDADE FATURADA"])
    product_client_col = billing_client_col
    if product_col and not fat_scope.empty:
        prod = fat_scope.copy()
        prod["_PRODUTO_REL"] = prod[product_col].fillna("Não informado").astype(str).str.strip()
        prod["_QTD_REL"] = to_number(prod[qty_col]) if qty_col else 1.0
        prod.loc[prod["_QTD_REL"] <= 0, "_QTD_REL"] = 1.0
        agg = {"Faturamento": ("_VALOR", "sum"), "Quantidade": ("_QTD_REL", "sum")}
        if product_client_col:
            agg["Clientes"] = (product_client_col, "nunique")
        products = prod.groupby("_PRODUTO_REL", as_index=False).agg(**agg)
        if "Clientes" not in products:
            products["Clientes"] = 0
        products["Preço médio"] = np.where(products["Quantidade"] != 0, products["Faturamento"] / products["Quantidade"], 0)
        total_products = float(products["Faturamento"].sum())
        products["Participação"] = np.where(total_products != 0, products["Faturamento"] / total_products, 0)
        products = products.sort_values("Faturamento", ascending=False).head(20)
        product_rows = [[
            r["_PRODUTO_REL"], money(r["Faturamento"]), percent(r["Participação"]), number(r["Quantidade"]), money(r["Preço médio"]), number(r["Clientes"])
        ] for _, r in products.iterrows()]
        story.append(hbar_chart(
            products.head(10)["_PRODUTO_REL"].tolist(),
            [("Faturamento", products.head(10)["Faturamento"].tolist(), blue)],
            "Top produtos por faturamento",
        ))
        story.append(data_table(
            ["Produto", "Faturamento", "Participação", "Quantidade", "Preço médio", "Clientes"], product_rows,
            [275, 105, 75, 70, 100, 55], {1, 2, 3, 4, 5},
        ))
    else:
        story.append(P("A base de faturamento não possui uma coluna de produto reconhecida para este escopo.", "SectionNote"))

    # 7 - Centro de custos
    story.append(PageBreak())
    story.extend(section("7. Centro de custos", "Receitas e despesas operacionais e não operacionais, sempre pela coluna Centro de Custos."))
    cc = period_filter(custos, start_month, end_month).copy()
    if scope_choice != "CONSOLIDADO":
        cc = cc[cc["_CC_N"] == scope_choice].copy()
    cc["Movimento"] = np.where(cc["_EMPRESA_N"] == "RECEITA", "Receita", "Despesa")
    cc["Classificação"] = np.where(cc["_GRUPO_N"].str.contains("NAO OPERACIONAIS", na=False), "Não operacional", "Operacional")
    cc["Departamento"] = cc["CENTRO DE CUSTOS"].map(cost_center_label)
    if not cc.empty:
        cc_revenue = float(cc.loc[cc["Movimento"] == "Receita", "_VALOR"].sum())
        cc_expense = float(cc.loc[cc["Movimento"] == "Despesa", "_VALOR"].sum())
        cc_metrics = [
            ("Receitas", money(cc_revenue)),
            ("Despesas", money(cc_expense)),
            ("Saldo", money(cc_revenue - cc_expense)),
            ("Centros de custos", number(cc["Departamento"].nunique())),
        ]
        story.append(metric_grid(cc_metrics, 4))
        story.append(Spacer(1, 5 * mm))
        dep = (
            cc.groupby(["Departamento", "Movimento"], as_index=False)["_VALOR"].sum()
            .pivot(index="Departamento", columns="Movimento", values="_VALOR").fillna(0).reset_index()
        )
        for col in ["Receita", "Despesa"]:
            if col not in dep:
                dep[col] = 0.0
        dep["Saldo"] = dep["Receita"] - dep["Despesa"]
        dep["Movimentação"] = dep["Receita"] + dep["Despesa"]
        dep = dep.sort_values("Movimentação", ascending=False).head(15)
        story.append(hbar_chart(
            dep["Departamento"].tolist(),
            [("Receitas", dep["Receita"].tolist(), blue), ("Despesas", dep["Despesa"].tolist(), red)],
            "Receitas e despesas por departamento",
        ))
        dep_rows = [[r["Departamento"], money(r["Receita"]), money(r["Despesa"]), money(r["Saldo"])] for _, r in dep.iterrows()]
        story.append(data_table(["Departamento", "Receitas", "Despesas", "Saldo"], dep_rows, [300, 125, 125, 125], {1, 2, 3}))

        story.append(Spacer(1, 5 * mm))
        exp_nature = (
            cc[cc["Movimento"] == "Despesa"].groupby("PAI", as_index=False)["_VALOR"].sum()
            .sort_values("_VALOR", ascending=False).head(20)
        )
        if not exp_nature.empty:
            story.append(P("Principais naturezas de despesa", "SectionTitle"))
            nature_rows = [[r["PAI"], money(r["_VALOR"]), percent(safe_div(r["_VALOR"], cc_expense))] for _, r in exp_nature.iterrows()]
            story.append(data_table(["Natureza", "Valor", "Participação"], nature_rows, [350, 130, 100], {1, 2}))
    else:
        story.append(P("Não há lançamentos de Centro de Custos para o período e escopo selecionados.", "SectionNote"))

    # Notas metodológicas
    story.append(PageBreak())
    story.extend(section("Notas metodológicas", "Critérios usados na elaboração do relatório."))
    notes = [
        ["Regime de análise", "Os indicadores financeiros principais são gerenciais em regime de caixa."],
        ["Linhas de negócio", "Receitas, despesas, margens e resultados diretos são determinados exclusivamente pela coluna Centro de Custos. A coluna de rateio não compõe nem é exibida nos indicadores do painel."],
        ["Faturamento e metas", "São indicadores comerciais por competência. O faturamento usa a coluna MÊS da BASE BI e a DT Emissão permanece como referência de auditoria."],
        ["Inadimplência", "O saldo é obtido do relatório ativo do CRM de cobrança, conforme o perfil e a linha atribuídos."],
        ["Margem de contribuição de caixa", "Receitas operacionais recebidas menos saídas variáveis pagas, dividido pelas receitas operacionais recebidas."],
        ["Limitação", "O relatório gerencial não substitui demonstrações contábeis oficiais ou conciliações do balancete."],
    ]
    story.append(data_table(["Tema", "Critério"], notes, [175, 555], set()))

    doc.build(story, onFirstPage=page_header_footer, onLaterPages=page_header_footer)
    return output.getvalue()




def build_business_performance_print_html(
    *,
    scope_choice: str,
    user: dict[str, str],
    start_month: pd.Period,
    end_month: pd.Period,
    fat: pd.DataFrame,
    metas: pd.DataFrame,
    metas_g: pd.DataFrame,
    receitas: pd.DataFrame,
    despesas: pd.DataFrame,
    custos: pd.DataFrame,
    performance: pd.DataFrame,
    inad: pd.DataFrame | None,
    company_monthly: pd.DataFrame,
    company_totals: dict[str, float],
    line_monthly: pd.DataFrame | None,
    line_totals: dict[str, float] | None,
    commercial_monthly: pd.DataFrame,
    commercial_totals: dict[str, float],
    lines_table: pd.DataFrame,
    fat_scope: pd.DataFrame,
    rec_scope: pd.DataFrame,
    cost_scope: pd.DataFrame,
    inad_scope: pd.DataFrame | None,
) -> str:
    """Relatório HTML visual, com a mesma identidade do app e impressão pelo navegador."""
    from datetime import datetime

    plotly_loaded = False

    def esc(value: object) -> str:
        return html.escape("" if value is None else str(value))

    def plot_html(fig: go.Figure, height: int = 320, legend_bottom: bool = True) -> str:
        nonlocal plotly_loaded
        fig = plot_layout(fig, height, legend_bottom)
        fig.update_layout(
            paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF",
            margin=dict(l=18, r=45, t=62, b=54 if legend_bottom else 30),
            font=dict(family="Arial, sans-serif", color="#344054", size=11),
        )
        include_js = "inline" if not plotly_loaded else False
        plotly_loaded = True
        return fig.to_html(
            full_html=False,
            include_plotlyjs=include_js,
            config={"displayModeBar": False, "responsive": True, "staticPlot": True},
            default_height=f"{height}px",
        )

    def metric_cards(items: list[tuple[str, str, str, str]]) -> str:
        cards = []
        for label, value, note, tone in items:
            cards.append(
                f"<div class='metric-card' style='--tone:{tone}'>"
                f"<div class='metric-label'>{esc(label)}</div>"
                f"<div class='metric-value'>{esc(value)}</div>"
                f"<div class='metric-note'>{esc(note)}</div></div>"
            )
        return "<div class='metric-grid'>" + "".join(cards) + "</div>"

    def table_html(df: pd.DataFrame, max_rows: int = 15) -> str:
        if df is None or df.empty:
            return "<div class='empty'>Sem registros para o período selecionado.</div>"
        view = df.head(max_rows).copy().where(lambda x: pd.notna(x), "")
        return "<div class='table-card'>" + view.to_html(index=False, border=0, classes="report-table", escape=True) + "</div>"

    def section(title: str, content: str, first: bool = False) -> str:
        cls = "report-section first" if first else "report-section"
        return f"<section class='{cls}'><div class='section-title'>{esc(title)}</div>{content}</section>"

    period_text = f"{month_label(start_month)} a {month_label(end_month)}"
    scope_text = line_label(scope_choice)
    generated = datetime.now().strftime("%d/%m/%Y %H:%M")
    overdue_total = float(inad_scope["_VALOR_VENCIDO"].sum()) if inad_scope is not None and not inad_scope.empty else 0.0

    body: list[str] = []
    hero = f"""
    <div class='report-hero'>
      <div><div class='brand'>FIRST INTELLIGENCE</div><h1>Business Performance</h1>
      <p>{esc(scope_text)} · {esc(period_text)} · Gerado em {generated}</p></div>
      <div class='hero-badge'>{esc(user.get('nome', 'Usuário'))}</div>
    </div>
    """

    if scope_choice == "CONSOLIDADO":
        cover_metrics = [
            ("Receitas operacionais recebidas", brl(company_totals.get("Receitas Operacionais", 0)), "Entradas realizadas da operação", BLUE),
            ("Saídas operacionais pagas", brl(company_totals.get("Saídas Operacionais", 0)), "Pagamentos operacionais", RED),
            ("Resultado operacional de caixa", brl(company_totals.get("Resultado Operacional de Caixa", 0)), "Receitas menos saídas", BLUE),
            ("Margem operacional de caixa", pct(company_totals.get("Margem de Caixa", 0)), "Resultado sobre receitas", NAVY),
            ("Faturamento", brl(commercial_totals.get("Faturamento", 0)), "Vendas emitidas", BLUE),
            ("Atingimento da meta", pct(commercial_totals.get("Atingimento", 0)), "Realizado sobre meta", NAVY),
            ("EBITDA gerencial de caixa", brl(company_totals.get("EBITDA Gerencial de Caixa", 0)), "Antes de IRPJ e CSLL pagos", BLUE),
            ("Inadimplência", brl(overdue_total), "Saldo vencido no CRM", RED),
        ]
        monthly = company_monthly
        rev_col, exp_col, res_col = "Receitas Operacionais", "Saídas Operacionais", "Resultado Operacional de Caixa"
    else:
        lt = line_totals or {}
        cover_metrics = [
            ("Receitas operacionais", brl(lt.get("Receitas Recebidas", 0)), "Recebimentos diretos", BLUE),
            ("Despesas operacionais", brl(lt.get("Custos Diretos Pagos", 0)), "Pagamentos diretos", RED),
            ("Resultado direto de caixa", brl(lt.get("Resultado Direto de Caixa", 0)), "Receitas menos despesas", BLUE if lt.get("Resultado Direto de Caixa", 0) >= 0 else RED),
            ("Margem direta de caixa", pct(lt.get("Margem Direta de Caixa", 0)), "Resultado sobre receitas", NAVY),
            ("Faturamento", brl(commercial_totals.get("Faturamento", 0)), "Vendas emitidas", BLUE),
            ("Atingimento da meta", pct(commercial_totals.get("Atingimento", 0)), "Realizado sobre meta", NAVY),
            ("Conversão em caixa", pct(lt.get("Conversão em Caixa", 0)), "Recebido sobre faturado", BLUE),
            ("Inadimplência", brl(overdue_total), "Saldo vencido da linha", RED),
        ]
        monthly = line_monthly if line_monthly is not None else pd.DataFrame()
        rev_col, exp_col, res_col = "Receitas Recebidas", "Custos Diretos Pagos", "Resultado Direto de Caixa"

    dash_parts = [hero, metric_cards(cover_metrics)]
    if monthly is not None and not monthly.empty:
        fig = go.Figure()
        fig.add_bar(x=monthly["Mês Texto"], y=monthly[rev_col], name="Receitas", marker_color=BLUE,
                    text=monthly[rev_col].map(compact_money), textposition="outside", cliponaxis=False)
        fig.add_bar(x=monthly["Mês Texto"], y=monthly[exp_col], name="Despesas", marker_color=RED,
                    text=monthly[exp_col].map(compact_money), textposition="inside", insidetextanchor="end",
                    textfont=dict(color=WHITE), cliponaxis=False)
        fig.add_scatter(x=monthly["Mês Texto"], y=monthly[res_col], name="Resultado", mode="lines+markers",
                        line=dict(color=NAVY, width=3), marker=dict(size=7))
        add_point_labels(fig, monthly["Mês Texto"], monthly[res_col], monthly[res_col].map(compact_money), color=NAVY)
        fig.update_layout(title="Movimento mensal de caixa", barmode="group")
        hide_value_axis(fig, "y")
        dash_parts.append("<div class='chart-card'>" + plot_html(fig, 330) + "</div>")

    if scope_choice == "CONSOLIDADO" and not lines_table.empty:
        f1 = line_result_lollipop(lines_table, "Resultado direto por linha")
        f2 = line_revenue_cost_dumbbell(lines_table, "Receitas x despesas por linha")
        dash_parts.append("<div class='two-col'><div class='chart-card'>" + plot_html(f1, 310, False) + "</div><div class='chart-card'>" + plot_html(f2, 310) + "</div></div>")
    body.append(section("Visão executiva", "".join(dash_parts), True))

    perf_metrics = metric_cards([
        ("Faturamento realizado", brl(commercial_totals.get("Faturamento", 0)), "Período selecionado", BLUE),
        ("Meta acumulada", brl(commercial_totals.get("Meta", 0)), "Objetivo do período", NAVY),
        ("Atingimento", pct(commercial_totals.get("Atingimento", 0)), "Realizado sobre meta", BLUE),
        ("Desvio", brl(commercial_totals.get("Desvio", 0)), "Realizado menos meta", BLUE if commercial_totals.get("Desvio", 0) >= 0 else RED),
        ("Meta anual", brl(commercial_totals.get("Meta Anual", 0)), str(end_month.year), NAVY),
        ("Projeção anual", brl(commercial_totals.get("Projeção Anual", 0)), "Ritmo atual anualizado", BLUE),
        ("Atingimento projetado", pct(commercial_totals.get("Atingimento Projetado", 0)), "Projeção sobre meta anual", NAVY),
        ("Média mensal", brl(commercial_totals.get("Média Mensal", 0)), "Faturamento médio", BLUE),
    ])
    perf_fig = go.Figure()
    perf_fig.add_bar(x=commercial_monthly["Mês Texto"], y=commercial_monthly["Faturamento"], name="Faturamento", marker_color=BLUE,
                     text=commercial_monthly["Faturamento"].map(compact_money), textposition="outside", cliponaxis=False)
    perf_fig.add_scatter(x=commercial_monthly["Mês Texto"], y=commercial_monthly["Meta"], name="Meta", mode="lines+markers",
                         line=dict(color=NAVY, width=3, dash="dot"), marker=dict(size=7))
    add_point_labels(perf_fig, commercial_monthly["Mês Texto"], commercial_monthly["Meta"], commercial_monthly["Meta"].map(compact_money), color=NAVY)
    perf_fig.update_layout(title="Faturamento realizado x meta")
    hide_value_axis(perf_fig, "y")
    sellers = seller_performance(fat, metas, start_month, end_month, scope_choice).head(15)
    seller_view = sellers[["Vendedor", "Faturamento", "Meta", "Atingimento", "Status"]].copy() if not sellers.empty else sellers
    if not seller_view.empty:
        seller_view["Faturamento"] = seller_view["Faturamento"].map(brl)
        seller_view["Meta"] = seller_view["Meta"].map(brl)
        seller_view["Atingimento"] = seller_view["Atingimento"].map(pct)
    body.append(section("Desempenho e metas", perf_metrics + "<div class='chart-card'>" + plot_html(perf_fig, 325) + "</div><h3>Desempenho da equipe</h3>" + table_html(seller_view, 15)))

    receipt_parts = []
    if scope_choice == "CONSOLIDADO":
        rec_metrics = [
            ("Recebimento previsto", brl(company_totals.get("Recebimento Previsto", 0)), "Carteira prevista", NAVY),
            ("Recebimento realizado", brl(company_totals.get("Recebimento Realizado", 0)), "Entradas realizadas", BLUE),
            ("Performance de recebimento", pct(company_totals.get("Performance de Recebimento", 0)), "Realizado sobre previsto", BLUE),
            ("Inadimplência", brl(overdue_total), "Saldo vencido", RED),
        ]
        rfig = go.Figure()
        rfig.add_bar(x=company_monthly["Mês Texto"], y=company_monthly["Recebimento Previsto"], name="Previsto", marker_color="#8AB8DC",
                     text=company_monthly["Recebimento Previsto"].map(compact_money), textposition="outside", cliponaxis=False)
        rfig.add_scatter(x=company_monthly["Mês Texto"], y=company_monthly["Recebimento Realizado"], name="Realizado", mode="lines+markers",
                         line=dict(color=BLUE, width=3), marker=dict(size=7))
        add_point_labels(rfig, company_monthly["Mês Texto"], company_monthly["Recebimento Realizado"], company_monthly["Recebimento Realizado"].map(compact_money), color=NAVY)
        rfig.update_layout(title="Recebimento previsto x realizado")
        hide_value_axis(rfig, "y")
        receipt_parts.append(metric_cards(rec_metrics) + "<div class='chart-card'>" + plot_html(rfig, 315) + "</div>")
    else:
        lt = line_totals or {}
        receipt_parts.append(metric_cards([
            ("Receitas recebidas", brl(lt.get("Receitas Recebidas", 0)), "Entradas da linha", BLUE),
            ("Faturamento", brl(commercial_totals.get("Faturamento", 0)), "Vendas emitidas", NAVY),
            ("Conversão em caixa", pct(lt.get("Conversão em Caixa", 0)), "Recebido sobre faturado", BLUE),
            ("Inadimplência", brl(overdue_total), "Saldo vencido", RED),
        ]))
    if inad_scope is not None and not inad_scope.empty:
        aging_order = ["Até 30 dias", "31 a 60 dias", "61 a 90 dias", "Acima de 90 dias"]
        aging = inad_scope.groupby("_FAIXA", observed=False)["_VALOR_VENCIDO"].sum().reindex(aging_order, fill_value=0).reset_index()
        aging.columns = ["Faixa", "Saldo"]
        af = px.bar(aging, x="Faixa", y="Saldo", title="Aging da inadimplência")
        af.update_traces(marker_color=BLUE, text=aging["Saldo"].map(compact_money), textposition="outside", cliponaxis=False)
        hide_value_axis(af, "y")
        debt = inad_scope.groupby("_CLIENTE", as_index=False)["_VALOR_VENCIDO"].sum().sort_values("_VALOR_VENCIDO", ascending=False).head(15)
        debt.columns = ["Cliente", "Saldo vencido"]
        debt["Saldo vencido"] = debt["Saldo vencido"].map(brl)
        receipt_parts.append("<div class='chart-card'>" + plot_html(af, 285, False) + "</div><h3>Maiores saldos vencidos</h3>" + table_html(debt, 15))
    body.append(section("Recebimentos e inadimplência", "".join(receipt_parts)))

    client_col = "NOME DO CLIENTE" if "NOME DO CLIENTE" in fat_scope.columns else optional_col(fat_scope, ["CLIENTE", "Cliente"])
    client_parts = []
    if client_col and not fat_scope.empty:
        billed = fat_scope.groupby(client_col, as_index=False)["_VALOR"].sum().sort_values("_VALOR", ascending=False).head(15)
        total_billed = float(fat_scope["_VALOR"].sum())
        billed["Participação"] = np.where(total_billed != 0, billed["_VALOR"] / total_billed, 0)
        billed.columns = ["Cliente", "Faturamento", "Participação"]
        chart_data = billed.head(10).sort_values("Faturamento").copy()
        chart_data["Nome"] = chart_data["Cliente"].map(lambda x: short_label(x, 35))
        cf = px.bar(chart_data, x="Faturamento", y="Nome", orientation="h", title="Principais clientes por faturamento")
        cf.update_traces(marker_color=BLUE, text=chart_data["Faturamento"].map(compact_money), textposition="outside", cliponaxis=False)
        hide_value_axis(cf, "x")
        billed["Faturamento"] = billed["Faturamento"].map(brl)
        billed["Participação"] = billed["Participação"].map(pct)
        client_parts.append("<div class='chart-card'>" + plot_html(cf, 330, False) + "</div>" + table_html(billed, 15))
    body.append(section("Clientes", "".join(client_parts) if client_parts else "<div class='empty'>Sem dados de clientes.</div>"))

    product_col = optional_col(fat_scope, ["PRODUTO", "DESCRIÇÃO", "LINHA DE PRODUTO", "ITEM"])
    product_parts = []
    if product_col and not fat_scope.empty:
        qty_col = optional_col(fat_scope, ["QUANTIDADE", "QTD", "QTDE", "QTD FATURADA", "QUANTIDADE FATURADA"])
        prod = fat_scope.copy()
        prod["_PRODUTO_REL"] = prod[product_col].fillna("Não informado").astype(str).str.strip()
        prod["_QTD_REL"] = to_number(prod[qty_col]) if qty_col else 1.0
        prod.loc[prod["_QTD_REL"] <= 0, "_QTD_REL"] = 1.0
        products = prod.groupby("_PRODUTO_REL", as_index=False).agg(Faturamento=("_VALOR", "sum"), Quantidade=("_QTD_REL", "sum"))
        products["Preço médio"] = np.where(products["Quantidade"] != 0, products["Faturamento"] / products["Quantidade"], 0)
        products = products.sort_values("Faturamento", ascending=False).head(20)
        cp = products.head(10).sort_values("Faturamento").copy()
        cp["Produto"] = cp["_PRODUTO_REL"].map(lambda x: short_label(x, 38))
        pf = px.bar(cp, x="Faturamento", y="Produto", orientation="h", title="Produtos por faturamento")
        pf.update_traces(marker_color=BLUE, text=cp["Faturamento"].map(compact_money), textposition="outside", cliponaxis=False)
        hide_value_axis(pf, "x")
        pv = products.rename(columns={"_PRODUTO_REL": "Produto"}).copy()
        pv["Faturamento"] = pv["Faturamento"].map(brl)
        pv["Quantidade"] = pv["Quantidade"].map(lambda v: f"{v:,.0f}".replace(",", "."))
        pv["Preço médio"] = pv["Preço médio"].map(brl)
        product_parts.append("<div class='chart-card'>" + plot_html(pf, 340, False) + "</div>" + table_html(pv, 20))
    body.append(section("Produtos", "".join(product_parts) if product_parts else "<div class='empty'>Sem dados de produtos.</div>"))

    cc = period_filter(custos, start_month, end_month).copy()
    if scope_choice != "CONSOLIDADO":
        cc = cc[cc["_CC_N"] == scope_choice].copy()
    cc["Movimento"] = np.where(cc["_EMPRESA_N"] == "RECEITA", "Receita", "Despesa")
    cc["Classificação"] = np.where(cc["_GRUPO_N"].str.contains("NAO OPERACIONAIS", na=False), "Não operacional", "Operacional")
    cc["Departamento"] = cc["CENTRO DE CUSTOS"].map(cost_center_label)
    cc_parts = []
    if not cc.empty:
        total_rev = float(cc.loc[cc["Movimento"] == "Receita", "_VALOR"].sum())
        total_exp = float(cc.loc[cc["Movimento"] == "Despesa", "_VALOR"].sum())
        cc_parts.append(metric_cards([
            ("Receitas", brl(total_rev), "Movimentos de entrada", BLUE),
            ("Despesas", brl(total_exp), "Movimentos de saída", RED),
            ("Saldo", brl(total_rev-total_exp), "Receitas menos despesas", BLUE if total_rev-total_exp >= 0 else RED),
            ("Centros de custos", f"{cc['Departamento'].nunique():,}".replace(",", "."), "Departamentos", NAVY),
        ]))
        dep = cc.groupby(["Departamento", "Movimento"], as_index=False)["_VALOR"].sum().pivot(index="Departamento", columns="Movimento", values="_VALOR").fillna(0).reset_index()
        for col in ["Receita", "Despesa"]:
            if col not in dep: dep[col] = 0.0
        dep["Movimentação"] = dep["Receita"] + dep["Despesa"]
        dep = dep.sort_values("Movimentação", ascending=False).head(15).sort_values("Movimentação")
        dep["Nome"] = dep["Departamento"].map(lambda x: short_label(x, 30))
        ccf = go.Figure()
        ccf.add_bar(x=dep["Receita"], y=dep["Nome"], orientation="h", name="Receitas", marker_color=BLUE,
                    text=[compact_money(v) if v > 0 else "" for v in dep["Receita"]], textposition="outside", cliponaxis=False)
        ccf.add_bar(x=dep["Despesa"], y=dep["Nome"], orientation="h", name="Despesas", marker_color=RED,
                    text=[compact_money(v) if v > 0 else "" for v in dep["Despesa"]], textposition="outside", cliponaxis=False)
        ccf.update_layout(title="Receitas e despesas por departamento", barmode="group")
        hide_value_axis(ccf, "x")
        dep_view = dep[["Departamento", "Receita", "Despesa"]].copy()
        dep_view["Saldo"] = dep_view["Receita"] - dep_view["Despesa"]
        for col in ["Receita", "Despesa", "Saldo"]: dep_view[col] = dep_view[col].map(brl)
        cc_parts.append("<div class='chart-card'>" + plot_html(ccf, 390) + "</div>" + table_html(dep_view.sort_values("Departamento"), 20))
    body.append(section("Centro de custos", "".join(cc_parts) if cc_parts else "<div class='empty'>Sem movimentos no centro de custos.</div>"))

    css = f"""
    :root {{--navy:{NAVY};--blue:{BLUE};--red:{RED};--light:#F3F6FA;--border:#DCE6EF;}}
    *{{box-sizing:border-box;-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}}
    body{{margin:0;background:#EEF3F8;color:#24364B;font-family:Arial,sans-serif}}
    .print-toolbar{{position:sticky;top:0;z-index:50;display:flex;align-items:center;justify-content:space-between;padding:10px 16px;background:#071B33;color:white;box-shadow:0 4px 16px rgba(7,27,51,.22)}}
    .print-toolbar button{{border:0;border-radius:10px;padding:10px 18px;background:#1261A0;color:white;font-weight:800;cursor:pointer}}
    .report{{max-width:1320px;margin:18px auto;padding:0 12px 30px}}
    .report-section{{background:white;border:1px solid var(--border);border-radius:22px;padding:22px;margin:0 0 18px;box-shadow:0 14px 36px rgba(7,27,51,.07);break-before:page}}
    .report-section.first{{break-before:auto}}
    .report-hero{{display:flex;align-items:center;justify-content:space-between;gap:20px;background:linear-gradient(120deg,#071B33,#0E3762 58%,#1B5D97);color:white;border-radius:20px;padding:24px 28px;margin-bottom:16px}}
    .brand{{font-size:11px;letter-spacing:.28em;font-weight:900;color:#B9D9F0}}
    .report-hero h1{{font-size:30px;margin:7px 0 5px;letter-spacing:-.03em}}
    .report-hero p{{margin:0;color:#DCEAF5;font-size:12px}}
    .hero-badge{{border:1px solid rgba(255,255,255,.25);background:rgba(255,255,255,.10);border-radius:999px;padding:8px 13px;font-size:12px;font-weight:800}}
    .section-title{{font-size:20px;font-weight:900;color:#071B33;border-bottom:1px solid var(--border);padding-bottom:10px;margin-bottom:14px}}
    .metric-grid{{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin-bottom:12px}}
    .metric-card{{position:relative;overflow:hidden;border:1px solid var(--border);border-radius:16px;padding:14px;background:linear-gradient(145deg,#fff,#F9FCFF);min-height:105px;break-inside:avoid}}
    .metric-card:after{{content:'';position:absolute;width:58px;height:58px;border-radius:50%;right:-22px;top:-24px;background:var(--tone);opacity:.10}}
    .metric-label{{font-size:9px;font-weight:900;letter-spacing:.055em;text-transform:uppercase;color:#667085}}
    .metric-value{{font-size:18px;font-weight:950;color:#071B33;margin-top:8px;white-space:nowrap}}
    .metric-note{{font-size:9px;color:#667085;margin-top:9px;line-height:1.3}}
    .two-col{{display:grid;grid-template-columns:1fr 1fr;gap:12px}}
    .chart-card,.table-card{{border:1px solid var(--border);border-radius:16px;padding:8px;background:white;margin:10px 0;break-inside:avoid}}
    h3{{font-size:15px;color:#071B33;margin:15px 0 8px}}
    .report-table{{width:100%;border-collapse:collapse;font-size:10px}}
    .report-table th{{background:#0B2F55;color:white;text-align:left;padding:8px 9px}}
    .report-table td{{padding:7px 9px;border-bottom:1px solid #E8EEF4}}
    .report-table tr:nth-child(even) td{{background:#F6F9FC}}
    .empty{{padding:20px;border:1px dashed #C7D7E5;border-radius:14px;color:#667085;text-align:center}}
    @page{{size:A4 landscape;margin:9mm}}
    @media print{{
      body{{background:white}}
      .print-toolbar{{display:none!important}}
      .report{{max-width:none;margin:0;padding:0}}
      .report-section{{box-shadow:none;border:0;border-radius:0;margin:0;padding:0;min-height:auto;break-before:page}}
      .report-section.first{{break-before:auto}}
      .report-hero{{border-radius:14px}}
      .chart-card,.table-card,.metric-card,.two-col{{break-inside:avoid-page}}
      .metric-grid{{gap:7px}}
    }}
    @media(max-width:900px){{.metric-grid{{grid-template-columns:repeat(2,1fr)}}.two-col{{grid-template-columns:1fr}}}}
    """

    return f"""<!doctype html><html lang='pt-BR'><head><meta charset='utf-8'><title>First Intelligence | Business Performance</title><style>{css}</style></head>
    <body><div class='print-toolbar'><span>Visualização de impressão · {esc(scope_text)} · {esc(period_text)}</span><button onclick='window.print()'>Imprimir / Salvar em PDF</button></div>
    <main class='report'>{''.join(body)}</main></body></html>"""

# =========================================================
# FONTES, USUÁRIO E FILTROS
# =========================================================
user = authenticate()
is_controladoria = user["perfil"] in {"CONTROLADORIA", "ADMIN"}
is_director = user["perfil"] in {"DIRETORIA", "ADMIN", "CONTROLADORIA"}
can_validate = is_controladoria
can_manage_sources = is_controladoria
access_label = (
    "Controladoria" if is_controladoria
    else "Diretoria" if user["perfil"] == "DIRETORIA"
    else line_label(user["linha"])
)
access_note = "Acesso de validação" if is_controladoria else ("Acesso executivo" if is_director else "Acesso por linha")

with st.sidebar:
    st.markdown(
        """
        <div class='first-sidebar'>
          <div class='brand'>FIRST<small>MEDICAL INTELLIGENCE</small></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div class='user-pill'><div class='left'><span class='avatar'></span><div><b>{user['nome']}</b><small>{access_note}</small></div></div><span>{access_label}</span></div>",
        unsafe_allow_html=True,
    )
    if secret_users() and st.button("Sair", width="stretch"):
        st.session_state.pop("auth_user", None)
        st.rerun()

    base_repo_files = repository_files("base")
    rev_repo_files = repository_files("rev")
    default_base = preferred_repository_file(base_repo_files, "BASE BI.xlsx")
    default_rev = preferred_repository_file(rev_repo_files, "rev2026 Base bi.xlsx")
    default_inad = first_existing([
        "base_crm_cobranca.csv",
        "relatorio_cobranca_gerente.xlsx", "relatorio_cobranca_gerente_2026-07-13.xlsx",
        "inadimplencia.xlsx", "Inadimplencia.xlsx", "base_inadimplencia.xlsx", "inadimplencia.csv"
    ])

    up_base = up_rev = up_inad = None
    use_session_uploads = False
    if can_manage_sources:
        with st.expander("Fontes de dados", expanded=False):
            if len(base_repo_files) > 1:
                selected_base_name = st.selectbox(
                    "BASE BI do repositório",
                    [p.name for p in base_repo_files],
                    index=[p.name for p in base_repo_files].index(default_base.name),
                    key="repository_base_choice",
                    help="Quando houver cópias com nomes diferentes, selecione explicitamente a base que deve ser usada.",
                )
                default_base = next(p for p in base_repo_files if p.name == selected_base_name)
                st.warning("Há mais de uma BASE BI no repositório. Exclua a cópia antiga para evitar atualizações na base errada.")
            elif default_base:
                st.caption(f"BASE BI do repositório: `{default_base.name}`")

            if len(rev_repo_files) > 1:
                selected_rev_name = st.selectbox(
                    "REV2026 do repositório",
                    [p.name for p in rev_repo_files],
                    index=[p.name for p in rev_repo_files].index(default_rev.name),
                    key="repository_rev_choice",
                )
                default_rev = next(p for p in rev_repo_files if p.name == selected_rev_name)
                st.warning("Há mais de uma REV2026 no repositório. Exclua a cópia antiga para evitar leituras incorretas.")
            elif default_rev:
                st.caption(f"REV2026 do repositório: `{default_rev.name}`")

            if st.button("↻ Recarregar BASE BI e REV2026", width="stretch", key="reload_repository_files"):
                for upload_key in ["up_base_final", "up_rev_final", "up_inad_final"]:
                    st.session_state.pop(upload_key, None)
                st.session_state["use_session_uploads"] = False
                st.session_state.pop("_active_source_ids", None)
                st.cache_data.clear()
                st.rerun()

            use_session_uploads = st.toggle(
                "Usar arquivos enviados nesta sessão",
                value=False,
                key="use_session_uploads",
                help="Desative para usar sempre os arquivos atualizados no GitHub.",
            )
            if use_session_uploads:
                up_base = st.file_uploader("Substituir BASE BI", type=["xlsx", "xlsm"], key="up_base_final")
                up_rev = st.file_uploader("Substituir REV2026", type=["xlsx", "xlsm"], key="up_rev_final")
                up_inad = st.file_uploader("Relatório do CRM de cobrança", type=["xlsx", "xlsm", "csv"], key="up_inad_final")
            else:
                st.caption("Fonte ativa: arquivos versionados no repositório.")

base_bytes, base_name = source_bytes(up_base if use_session_uploads else None, default_base)
rev_bytes, rev_name = source_bytes(up_rev if use_session_uploads else None, default_rev)
inad_bytes, inad_name = source_bytes(up_inad if use_session_uploads else None, default_inad)

base_id = file_fingerprint(base_bytes)
rev_id = file_fingerprint(rev_bytes)
inad_id = file_fingerprint(inad_bytes)

# Se qualquer arquivo versionado mudar, invalida automaticamente todas as leituras derivadas.
active_source_ids = {"base": base_id, "rev": rev_id, "inad": inad_id}
if st.session_state.get("_active_source_ids") != active_source_ids:
    st.cache_data.clear()
    st.session_state["_active_source_ids"] = active_source_ids

if not base_bytes or not rev_bytes:
    st.error("Inclua `BASE BI.xlsx` e `rev2026 Base bi.xlsx` no repositório ou carregue os arquivos na barra lateral.")
    st.stop()

try:
    with st.spinner("Organizando visão de caixa e linhas de negócio..."):
        base = load_base_bi(base_bytes)
        rev = load_rev(rev_bytes)
        billing_audit = base.get("billing_audit", {})
        rateio_meta = rev.get("rateio_meta", {})

        # O painel é fechado exclusivamente no ano-base de 2026.
        fat = filter_analysis_year(base["faturamento"])
        metas = filter_analysis_year(base["metas"])
        metas_g = filter_analysis_year(base["metas_gerentes"])
        receitas_base = filter_analysis_year(rev["receitas"])
        despesas = filter_analysis_year(rev["despesas"])
        custos = filter_analysis_year(rev["custos"])
        rateio = filter_analysis_year(rev.get("rateio", pd.DataFrame()))
        custos_rateados = filter_analysis_year(rev.get("custos_rateados", pd.DataFrame()))
        performance = filter_analysis_year(rev["recebimento"])
        # Ajuste pontual: recebimentos Microtech comprovados na BASE BI não permanecem em Vendas.
        # Custos continuam integralmente classificados pelo CENTRO DE CUSTOS original.
        custos, cc_microtech_stats = correct_microtech_receipts_in_vendas(custos, fat)
        receitas, receipt_match_stats = assign_receipt_lines(receitas_base, fat)
        inad = None
        inad_meta = None
        if inad_bytes:
            inad, inad_meta = prepare_inadimplencia(inad_bytes, inad_name, fat)
except Exception as exc:
    st.error(f"Não foi possível carregar as bases: {exc}")
    st.stop()

all_periods = pd.concat([fat["_MES"], receitas["_MES"], despesas["_MES"], custos["_MES"], performance["_MES"]]).dropna()
all_periods = all_periods[all_periods.map(lambda value: value.year if isinstance(value, pd.Period) else pd.Period(value, freq="M").year).eq(ANALYSIS_YEAR)]
if all_periods.empty:
    st.error(f"Não foram encontrados movimentos de {ANALYSIS_YEAR} nas bases carregadas.")
    st.stop()
min_month, max_month = all_periods.min(), all_periods.max()
month_options = list(pd.period_range(min_month, max_month, freq="M"))

with st.sidebar:
    st.markdown(f"#### Período · {ANALYSIS_YEAR}")
    start_month = st.selectbox("Mês inicial", month_options, index=0, format_func=month_label)
    end_month = st.selectbox("Mês final", month_options, index=len(month_options) - 1, format_func=month_label)
    if start_month > end_month:
        start_month = end_month

    if is_director:
        scope_choice = st.selectbox("Escopo do painel", ["CONSOLIDADO"] + LINES, format_func=line_label)
    else:
        scope_choice = user["linha"] if user["linha"] in LINES else "VENDAS"

    st.markdown("#### Navegação")
    pages = ["Dashboard", "Desempenho & metas", "Recebimentos & inadimplência", "Clientes", "Produtos", "Centro de custos", "Relatório para impressão"]
    if is_director:
        pages.insert(1, "Linhas de negócio")
    if can_validate:
        pages.insert(-1, "Validação dos dados")
    if "nav_page" not in st.session_state or st.session_state.get("nav_page") not in pages:
        st.session_state["nav_page"] = pages[0]
    if st.button("🖨️ Abrir impressão visual", width="stretch", key="open_visual_print"):
        st.session_state["nav_page"] = "Relatório para impressão"
    page = st.radio("Navegação", pages, label_visibility="collapsed", key="nav_page")


period_text = f"{month_label(start_month)} a {month_label(end_month)}"
scope_text = line_label(scope_choice)

st.markdown(
    f"""
    <div class='dashboard-hero'>
      <div class='hero-grid'>
        <div>
          <div class='hero-brand'>First Medical · Controladoria</div>
          <h1>{'Business Performance' if scope_choice == 'CONSOLIDADO' else 'Business Performance · ' + scope_text}</h1>
        </div>
        <div class='hero-chips'>
          <span class='hero-chip'>{period_text}</span>
          <span class='hero-chip'>{scope_text}</span>
          <span class='hero-chip'>{'Acesso controladoria' if is_controladoria else ('Acesso diretoria' if is_director else 'Acesso restrito')}</span>
        </div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

company_monthly = company_cash_monthly(fat, metas, receitas, despesas, performance, start_month, end_month)
company_totals = totals_from_monthly(company_monthly, False)
line_monthly = line_cash_monthly(scope_choice, fat, receitas, custos, start_month, end_month) if scope_choice != "CONSOLIDADO" else None
line_totals = totals_from_monthly(line_monthly, True) if line_monthly is not None else None
commercial_monthly = commercial_performance_monthly(fat, metas_g, start_month, end_month, scope_choice)
commercial_totals = commercial_performance_totals(commercial_monthly, metas_g, scope_choice, end_month.year)
lines_table = line_summary(fat, metas_g, receitas, custos, start_month, end_month, inad)
fat_scope, rec_scope, cost_scope, inad_scope = scoped_data(scope_choice, fat, receitas, custos, inad, start_month, end_month)


# =========================================================
# RELATÓRIO VISUAL PARA IMPRESSÃO
# =========================================================
if page == "Relatório para impressão":
    st.markdown("<div class='section-head'><div><h3>Relatório visual completo</h3></div><span class='section-badge'>Ctrl + P</span></div>", unsafe_allow_html=True)
    st.caption("A visualização abaixo reúne todos os menus com as mesmas cores do dashboard. Use o botão interno para imprimir ou salvar em PDF.")
    print_html = build_business_performance_print_html(
        scope_choice=scope_choice, user=user, start_month=start_month, end_month=end_month,
        fat=fat, metas=metas, metas_g=metas_g, receitas=receitas, despesas=despesas,
        custos=custos, performance=performance, inad=inad,
        company_monthly=company_monthly, company_totals=company_totals,
        line_monthly=line_monthly, line_totals=line_totals,
        commercial_monthly=commercial_monthly, commercial_totals=commercial_totals,
        lines_table=lines_table, fat_scope=fat_scope, rec_scope=rec_scope,
        cost_scope=cost_scope, inad_scope=inad_scope,
    )
    components.html(print_html, height=1250, scrolling=True)

# =========================================================
# DASHBOARD
# =========================================================
elif page == "Dashboard":
    if scope_choice == "CONSOLIDADO":
        k1, k2, k3, k4 = st.columns(4)
        with k1: card("Receitas operacionais recebidas", brl(company_totals["Receitas Operacionais"]), "Entradas efetivamente recebidas da operação", BLUE)
        with k2: card("Saídas operacionais pagas", brl(company_totals["Saídas Operacionais"]), "Pagamentos operacionais efetivados", RED)
        with k3: card("Resultado operacional de caixa", brl(company_totals["Resultado Operacional de Caixa"]), "Receitas recebidas menos saídas pagas", BLUE if company_totals["Resultado Operacional de Caixa"] >= 0 else CYAN)
        with k4: card("Margem operacional de caixa", pct(company_totals["Margem de Caixa"]), "Resultado operacional ÷ receitas recebidas", TEAL)

        k5, k6, k7 = st.columns(3)
        with k5: card("EBITDA gerencial de caixa", brl(company_totals["EBITDA Gerencial de Caixa"]), "Resultado antes de IRPJ e CSLL pagos", NAVY)
        with k6: card("Margem de contribuição de caixa", pct(company_totals["Margem de Contribuição Caixa"]), "Receitas recebidas menos saídas variáveis pagas", CYAN)
        with k7: card("Entradas não operacionais", brl(company_totals["Entradas Não Operacionais"]), "Capital de giro e outras fontes não operacionais", CYAN)

        c1, c2 = st.columns([1.35, 1])
        with c1:
            fig = go.Figure()
            fig.add_bar(x=company_monthly["Mês Texto"], y=company_monthly["Receitas Operacionais"], name="Receitas diretas", marker_color=BLUE,
                        text=company_monthly["Receitas Operacionais"].map(compact_money), textposition="outside", cliponaxis=False, textfont=dict(size=10, color=NAVY))
            fig.add_bar(x=company_monthly["Mês Texto"], y=company_monthly["Saídas Operacionais"], name="Saídas pagas", marker_color=RED,
                        text=company_monthly["Saídas Operacionais"].map(compact_money), textposition="inside", insidetextanchor="end", cliponaxis=False, textfont=dict(size=10, color=WHITE))
            fig.add_scatter(x=company_monthly["Mês Texto"], y=company_monthly["Resultado Operacional de Caixa"], name="Resultado", mode="lines+markers",
                            line=dict(color=TEAL, width=3), marker=dict(size=7))
            add_point_labels(fig, company_monthly["Mês Texto"], company_monthly["Resultado Operacional de Caixa"],
                             company_monthly["Resultado Operacional de Caixa"].map(compact_money), color=NAVY)
            fig.update_layout(title="Movimento operacional de caixa", barmode="group")
            hide_value_axis(fig, "y")
            st.plotly_chart(plot_layout(fig, 365), width="stretch", config={"displayModeBar": False})
        with c2:
            bridge_values = [
                company_totals["Receitas Operacionais"], -company_totals["Saídas Variáveis"],
                -company_totals["Saídas Fixas"], -company_totals["IRPJ e CSLL"],
                company_totals["Resultado Operacional de Caixa"],
            ]
            bridge = go.Figure(go.Waterfall(
                orientation="v", measure=["absolute", "relative", "relative", "relative", "total"],
                x=["Receitas<br>recebidas", "Saídas<br>variáveis", "Saídas<br>fixas", "IRPJ +<br>CSLL", "Resultado<br>de caixa"],
                y=bridge_values, text=[compact_money(v) for v in bridge_values], textposition="outside", cliponaxis=False,
                connector={"line": {"color": "#AAB5C1"}}, increasing={"marker": {"color": BLUE}},
                decreasing={"marker": {"color": RED}}, totals={"marker": {"color": NAVY}},
            ))
            bridge.update_layout(title="Ponte do resultado operacional de caixa")
            hide_value_axis(bridge, "y")
            st.plotly_chart(plot_layout(bridge, 365, False), width="stretch", config={"displayModeBar": False})

        section_header("Resultado por linha de negócio", "Comparativo direto", "Sem rateio")
        dash_selected_codes = st.multiselect(
            "Linhas exibidas", LINES, default=LINES, format_func=line_label, key="filter_dashboard_lines"
        )
        dashboard_lines = lines_table[lines_table["Código"].isin(dash_selected_codes)].copy() if dash_selected_codes else lines_table.copy()
        l1, l2 = st.columns([1.05, 1.25])
        with l1:
            fig = line_result_lollipop(dashboard_lines, "Resultado direto por linha")
            st.plotly_chart(plot_layout(fig, 350, False), width="stretch", config={"displayModeBar": False})
        with l2:
            fig = line_revenue_cost_dumbbell(dashboard_lines, "Receitas x despesas diretas")
            st.plotly_chart(plot_layout(fig, 350), width="stretch", config={"displayModeBar": False})

        section_header("Desempenho comercial", "Realizado, meta e projeção")
        s1, s2, s3, s4 = st.columns(4)
        with s1: card("Faturamento realizado", brl(commercial_totals["Faturamento"]), "Vendas emitidas no período", BLUE)
        with s2: card("Meta acumulada", brl(commercial_totals["Meta"]), "Meta comercial do mesmo período", NAVY)
        with s3: card("Atingimento da meta", pct(commercial_totals["Atingimento"]), "Faturamento ÷ meta", BLUE if commercial_totals["Atingimento"] >= 1 else CYAN)
        with s4: card("Desvio da meta", brl(commercial_totals["Desvio"]), "Realizado menos meta", BLUE if commercial_totals["Desvio"] >= 0 else CYAN)

        s5, s6, s7, s8 = st.columns(4)
        with s5: card("Meta anual", brl(commercial_totals["Meta Anual"]), f"Meta cadastrada para {end_month.year}", NAVY)
        with s6: card("Projeção anual", brl(commercial_totals["Projeção Anual"]), "Ritmo médio do período anualizado", CYAN)
        with s7: card("Atingimento projetado", pct(commercial_totals["Atingimento Projetado"]), "Projeção anual ÷ meta anual", BLUE if commercial_totals["Atingimento Projetado"] >= 1 else CYAN)
        with s8: card("Conversão em caixa", pct(company_totals["Conversão em Caixa"]), "Receita operacional recebida ÷ faturamento", TEAL)

        perf_fig = go.Figure()
        perf_fig.add_bar(
            x=commercial_monthly["Mês Texto"], y=commercial_monthly["Faturamento"], name="Faturamento", marker_color=BLUE,
            text=commercial_monthly["Faturamento"].map(compact_money), textposition="outside", cliponaxis=False,
        )
        perf_fig.add_scatter(
            x=commercial_monthly["Mês Texto"], y=commercial_monthly["Meta"], name="Meta", mode="lines+markers",
            line=dict(color=NAVY, width=3, dash="dot"), marker=dict(size=7),
        )
        add_point_labels(perf_fig, commercial_monthly["Mês Texto"], commercial_monthly["Meta"], commercial_monthly["Meta"].map(compact_money), color=NAVY)
        perf_fig.update_layout(title="Faturamento realizado x meta mensal")
        hide_value_axis(perf_fig, "y")
        st.plotly_chart(plot_layout(perf_fig, 350), width="stretch", config={"displayModeBar": False})

        section_header("Clientes e custos que formam o resultado", "Abertura restrita à linha selecionada")
        c1, c2 = st.columns(2)
        with c1:
            client_col = "NOME DO CLIENTE" if "NOME DO CLIENTE" in fat_scope.columns else "CLIENTE"
            top_clients = fat_scope.groupby(client_col, as_index=False)["_VALOR"].sum().sort_values("_VALOR", ascending=False).head(10)
            top_clients["Nome"] = top_clients[client_col].map(lambda x: short_label(x, 38))
            p = top_clients.sort_values("_VALOR")
            fig = px.bar(p, x="_VALOR", y="Nome", orientation="h", title="Principais clientes por faturamento", custom_data=[client_col])
            fig.update_traces(marker_color=TEAL, text=p["_VALOR"].map(compact_money), textposition="outside", cliponaxis=False,
                              hovertemplate="%{customdata[0]}<br>Faturamento: R$ %{x:,.2f}<extra></extra>")
            hide_value_axis(fig, "x")
            st.plotly_chart(plot_layout(fig, 370, False), width="stretch", config={"displayModeBar": False})
        with c2:
            top_cost = cost_scope.groupby("PAI", as_index=False)["_VALOR"].sum().sort_values("_VALOR", ascending=False).head(10)
            top_cost["Natureza"] = top_cost["PAI"].map(lambda x: short_label(x, 36))
            p = top_cost.sort_values("_VALOR")
            fig = px.bar(p, x="_VALOR", y="Natureza", orientation="h", title="Principais custos diretos pagos", custom_data=["PAI"])
            fig.update_traces(marker_color=BLUE, text=p["_VALOR"].map(compact_money), textposition="outside", cliponaxis=False,
                              hovertemplate="%{customdata[0]}<br>Valor: R$ %{x:,.2f}<extra></extra>")
            hide_value_axis(fig, "x")
            st.plotly_chart(plot_layout(fig, 370, False), width="stretch", config={"displayModeBar": False})

    else:
        t = line_totals
        overdue = float(inad_scope["_VALOR_VENCIDO"].sum()) if inad_scope is not None and not inad_scope.empty else 0.0

        k1, k2, k3, k4 = st.columns(4)
        with k1:
            card("Receitas operacionais", brl(t["Receitas Recebidas"]), "Recebimentos diretos da linha", BLUE)
        with k2:
            card("Despesas operacionais", brl(t["Custos Diretos Pagos"]), "Pagamentos diretos da linha", RED)
        with k3:
            card(
                "Resultado direto de caixa", brl(t["Resultado Direto de Caixa"]),
                "Receitas menos despesas diretas", BLUE if t["Resultado Direto de Caixa"] >= 0 else RED,
            )
        with k4:
            card(
                "Margem direta de caixa", pct(t["Margem Direta de Caixa"]),
                "Resultado direto sobre receitas", BLUE if t["Margem Direta de Caixa"] >= 0 else RED,
            )

        k5, k6, k7, k8 = st.columns(4)
        with k5:
            card("Faturamento", brl(commercial_totals["Faturamento"]), "Vendas emitidas no período", NAVY)
        with k6:
            card("Atingimento da meta", pct(commercial_totals["Atingimento"]), "Faturamento sobre meta", BLUE)
        with k7:
            card("Conversão em caixa", pct(t["Conversão em Caixa"]), "Receita recebida sobre faturamento", TEAL)
        with k8:
            card("Inadimplência", brl(overdue) if inad is not None else "Base não carregada", "Saldo vencido da linha", RED if overdue > 0 else BLUE)

        c1, c2 = st.columns([1.18, 1])
        with c1:
            fig = go.Figure()
            fig.add_bar(
                x=line_monthly["Mês Texto"], y=line_monthly["Receitas Recebidas"],
                name="Receitas", marker_color=BLUE,
                text=line_monthly["Receitas Recebidas"].map(compact_money),
                textposition="outside", cliponaxis=False,
            )
            fig.add_bar(
                x=line_monthly["Mês Texto"], y=line_monthly["Custos Diretos Pagos"],
                name="Despesas", marker_color=RED,
                text=line_monthly["Custos Diretos Pagos"].map(compact_money),
                textposition="inside", insidetextanchor="end", cliponaxis=False,
                textfont=dict(color=WHITE, size=10),
            )
            fig.add_scatter(
                x=line_monthly["Mês Texto"], y=line_monthly["Resultado Direto de Caixa"],
                name="Resultado", mode="lines+markers",
                line=dict(color=NAVY, width=3), marker=dict(size=7),
            )
            add_point_labels(
                fig, line_monthly["Mês Texto"], line_monthly["Resultado Direto de Caixa"],
                line_monthly["Resultado Direto de Caixa"].map(compact_money), color=NAVY,
            )
            fig.update_layout(title="Resultado mensal de caixa", barmode="group")
            hide_value_axis(fig, "y")
            st.plotly_chart(plot_layout(fig, 365), width="stretch", config={"displayModeBar": False})

        with c2:
            fig = go.Figure()
            fig.add_bar(
                x=commercial_monthly["Mês Texto"], y=commercial_monthly["Faturamento"],
                name="Faturamento", marker_color=BLUE,
                text=commercial_monthly["Faturamento"].map(compact_money),
                textposition="outside", cliponaxis=False,
            )
            fig.add_scatter(
                x=commercial_monthly["Mês Texto"], y=commercial_monthly["Meta"],
                name="Meta", mode="lines+markers",
                line=dict(color=NAVY, width=3, dash="dot"), marker=dict(size=7),
            )
            add_point_labels(
                fig, commercial_monthly["Mês Texto"], commercial_monthly["Meta"],
                commercial_monthly["Meta"].map(compact_money), color=NAVY,
            )
            fig.update_layout(title="Faturamento x meta")
            hide_value_axis(fig, "y")
            st.plotly_chart(plot_layout(fig, 365), width="stretch", config={"displayModeBar": False})

        section_header("Desempenho comercial", badge=scope_text)
        s1, s2, s3, s4 = st.columns(4)
        with s1:
            card("Meta acumulada", brl(commercial_totals["Meta"]), "Objetivo do período", NAVY)
        with s2:
            card(
                "Desvio da meta", brl(commercial_totals["Desvio"]),
                "Realizado menos meta", BLUE if commercial_totals["Desvio"] >= 0 else RED,
            )
        with s3:
            card("Projeção anual", brl(commercial_totals["Projeção Anual"]), "Ritmo atual anualizado", BLUE)
        with s4:
            card(
                "Atingimento projetado", pct(commercial_totals["Atingimento Projetado"]),
                "Projeção sobre meta anual", BLUE if commercial_totals["Atingimento Projetado"] >= 1 else NAVY,
            )

        section_header("Clientes e despesas", badge=scope_text)
        c3, c4 = st.columns(2)
        with c3:
            client_col = "NOME DO CLIENTE" if "NOME DO CLIENTE" in fat_scope.columns else "CLIENTE"
            top_clients = (
                fat_scope.groupby(client_col, as_index=False)["_VALOR"].sum()
                .sort_values("_VALOR", ascending=False).head(10)
            )
            if top_clients.empty:
                st.info("Não há clientes com faturamento no período selecionado.")
            else:
                top_clients["Nome"] = top_clients[client_col].map(lambda x: short_label(x, 38))
                p = top_clients.sort_values("_VALOR")
                fig = px.bar(
                    p, x="_VALOR", y="Nome", orientation="h",
                    title="Principais clientes", custom_data=[client_col],
                )
                fig.update_traces(
                    marker_color=BLUE, text=p["_VALOR"].map(compact_money),
                    textposition="outside", cliponaxis=False,
                    hovertemplate="%{customdata[0]}<br>Faturamento: R$ %{x:,.2f}<extra></extra>",
                )
                hide_value_axis(fig, "x")
                st.plotly_chart(plot_layout(fig, 370, False), width="stretch", config={"displayModeBar": False})

        with c4:
            top_cost = (
                cost_scope.groupby("PAI", as_index=False)["_VALOR"].sum()
                .sort_values("_VALOR", ascending=False).head(10)
            )
            if top_cost.empty:
                st.info("Não há despesas diretas no período selecionado.")
            else:
                top_cost["Natureza"] = top_cost["PAI"].map(lambda x: short_label(x, 36))
                p = top_cost.sort_values("_VALOR")
                fig = px.bar(
                    p, x="_VALOR", y="Natureza", orientation="h",
                    title="Principais despesas", custom_data=["PAI"],
                )
                fig.update_traces(
                    marker_color=RED, text=p["_VALOR"].map(compact_money),
                    textposition="outside", cliponaxis=False,
                    hovertemplate="%{customdata[0]}<br>Valor: R$ %{x:,.2f}<extra></extra>",
                )
                hide_value_axis(fig, "x")
                st.plotly_chart(plot_layout(fig, 370, False), width="stretch", config={"displayModeBar": False})


# =========================================================
# DESEMPENHO E METAS
# =========================================================
elif page == "Desempenho & metas":
    section_header("Desempenho e metas", "Acompanhamento comercial do período", scope_text)

    selected_scope = scope_choice
    selected_lines = LINES
    if is_director:
        selected_lines = st.multiselect(
            "Linhas consideradas", LINES, default=LINES, format_func=line_label, key="performance_lines_filter"
        )

    if is_director and scope_choice == "CONSOLIDADO" and selected_lines and set(selected_lines) != set(LINES):
        perf_frames = [commercial_performance_monthly(fat, metas_g, start_month, end_month, line) for line in selected_lines]
        perf_monthly = pd.concat(perf_frames).groupby("Mês", as_index=False)[["Faturamento", "Meta", "Desvio"]].sum()
        perf_monthly["Atingimento"] = np.where(perf_monthly["Meta"] != 0, perf_monthly["Faturamento"] / perf_monthly["Meta"], 0)
        perf_monthly["Mês Texto"] = perf_monthly["Mês"].map(month_label)
        actual = float(perf_monthly["Faturamento"].sum())
        target = float(perf_monthly["Meta"].sum())
        annual_target = 0.0
        for line in selected_lines:
            line_full = commercial_performance_totals(
                commercial_performance_monthly(fat, metas_g, start_month, end_month, line), metas_g, line, end_month.year
            )
            annual_target += line_full["Meta Anual"]
        projection = safe_div(actual, max(len(perf_monthly), 1)) * 12
        perf_totals = {
            "Faturamento": actual, "Meta": target, "Desvio": actual - target, "Atingimento": safe_div(actual, target),
            "Média Mensal": safe_div(actual, max(len(perf_monthly), 1)), "Meta Anual": annual_target,
            "Projeção Anual": projection, "Atingimento Projetado": safe_div(projection, annual_target),
        }
    else:
        perf_monthly = commercial_monthly.copy()
        perf_totals = commercial_totals.copy()

    p1, p2, p3, p4 = st.columns(4)
    with p1: card("Faturamento realizado", brl(perf_totals["Faturamento"]), "Resultado comercial do período", BLUE)
    with p2: card("Meta acumulada", brl(perf_totals["Meta"]), "Meta correspondente ao período", NAVY)
    with p3: card("Atingimento", pct(perf_totals["Atingimento"]), "Faturamento ÷ meta", BLUE if perf_totals["Atingimento"] >= 1 else CYAN)
    with p4: card("Desvio", brl(perf_totals["Desvio"]), "Realizado menos meta", BLUE if perf_totals["Desvio"] >= 0 else CYAN)

    p5, p6, p7, p8 = st.columns(4)
    with p5: card("Média mensal", brl(perf_totals["Média Mensal"]), "Média do período selecionado", CYAN)
    with p6: card("Meta anual", brl(perf_totals["Meta Anual"]), f"Objetivo comercial de {end_month.year}", NAVY)
    with p7: card("Projeção anual", brl(perf_totals["Projeção Anual"]), "Ritmo médio anualizado", BLUE)
    with p8: card("Atingimento projetado", pct(perf_totals["Atingimento Projetado"]), "Projeção ÷ meta anual", BLUE if perf_totals["Atingimento Projetado"] >= 1 else CYAN)

    c1, c2 = st.columns([1.35, 1])
    with c1:
        fig = go.Figure()
        fig.add_bar(
            x=perf_monthly["Mês Texto"], y=perf_monthly["Faturamento"], name="Faturamento", marker_color=BLUE,
            text=perf_monthly["Faturamento"].map(compact_money), textposition="outside", cliponaxis=False,
        )
        fig.add_scatter(
            x=perf_monthly["Mês Texto"], y=perf_monthly["Meta"], name="Meta", mode="lines+markers",
            line=dict(color=NAVY, width=3, dash="dot"), marker=dict(size=7),
        )
        add_point_labels(fig, perf_monthly["Mês Texto"], perf_monthly["Meta"], perf_monthly["Meta"].map(compact_money), color=NAVY)
        fig.update_layout(title="Faturamento x meta mensal")
        hide_value_axis(fig, "y")
        st.plotly_chart(plot_layout(fig, 365), width="stretch", config={"displayModeBar": False})
    with c2:
        att = perf_monthly[perf_monthly["Atingimento"] > 0].copy()
        colors = [BLUE if value >= 1 else CYAN for value in att["Atingimento"]]
        fig = go.Figure(go.Bar(
            x=att["Mês Texto"], y=att["Atingimento"], marker_color=colors,
            text=att["Atingimento"].map(pct), textposition="outside", cliponaxis=False,
            hovertemplate="%{x}<br>Atingimento: %{y:.1%}<extra></extra>",
        ))
        fig.add_hline(y=1, line_dash="dot", line_color=NAVY, line_width=2)
        fig.update_layout(title="Atingimento mensal")
        hide_value_axis(fig, "y")
        st.plotly_chart(plot_layout(fig, 365, False), width="stretch", config={"displayModeBar": False})

    if is_director and scope_choice == "CONSOLIDADO":
        section_header("Desempenho por linha", "Comparação comercial e financeira", "Diretoria")
        line_perf = lines_table[lines_table["Código"].isin(selected_lines)].copy() if selected_lines else lines_table.copy()
        score = line_perf[[
            "Linha", "Faturamento", "Meta", "Atingimento da Meta", "Desvio da Meta",
            "Receitas Recebidas", "Conversão em Caixa", "Resultado Direto de Caixa",
            "Margem Direta de Caixa", "Inadimplência"
        ]].copy()
        score["Inadimplência / faturamento"] = np.where(score["Faturamento"] != 0, score["Inadimplência"] / score["Faturamento"], 0)
        score = score[score["Atingimento da Meta"] > 0].copy()
        score["Status"] = np.select(
            [score["Atingimento da Meta"] >= 1, score["Atingimento da Meta"] >= .9],
            ["Meta atingida", "Próximo da meta"], default="Abaixo da meta"
        )

        chart = line_perf[line_perf["Atingimento da Meta"] > 0].sort_values("Atingimento da Meta")
        fig = go.Figure(go.Bar(
            x=chart["Atingimento da Meta"], y=chart["Linha"], orientation="h",
            marker_color=[BLUE if value >= 1 else CYAN for value in chart["Atingimento da Meta"]],
            text=chart["Atingimento da Meta"].map(pct), textposition="outside", cliponaxis=False,
        ))
        fig.add_vline(x=1, line_dash="dot", line_color=NAVY, line_width=2)
        fig.update_layout(title="Atingimento da meta por linha", showlegend=False)
        hide_value_axis(fig, "x")
        st.plotly_chart(plot_layout(fig, 335, False), width="stretch", config={"displayModeBar": False})

        score_view = score.copy()
        for col in ["Faturamento", "Meta", "Desvio da Meta", "Receitas Recebidas", "Resultado Direto de Caixa", "Inadimplência"]:
            score_view[col] = score_view[col].map(brl)
        for col in ["Atingimento da Meta", "Conversão em Caixa", "Margem Direta de Caixa", "Inadimplência / faturamento"]:
            score_view[col] = score_view[col].map(pct)
        clean_table(score_view, height=255)

    section_header(
        "Desempenho da equipe",
        "Equipe, metas e resultados individuais oficiais da aba Metas",
    )
    seller_perf = seller_performance(fat, metas, start_month, end_month, scope_choice)
    seller_reconciliation = float(seller_perf.attrs.get("reconciliation_difference", 0.0))
    seller_billing_total = float(seller_perf.attrs.get("billing_total", 0.0))
    seller_meta_actual = float(seller_perf.attrs.get("meta_actual_total", 0.0))
    if is_controladoria and abs(seller_reconciliation) > 0.01:
        st.info(
            "Conciliação das fontes: o faturamento total da linha vem do Banco de Dados de Faturamento "
            f"({brl(seller_billing_total)}), enquanto os resultados individuais vêm da coluna ATINGIMENTO "
            f"da aba Metas ({brl(seller_meta_actual)}). Diferença entre as fontes: {brl(seller_reconciliation)}."
        )
    sf1, sf2, sf3 = st.columns([1.4, 1, .7])
    seller_search = sf1.text_input("Buscar vendedor ou representante", key="seller_performance_search")
    seller_status = sf2.multiselect(
        "Status", ["Meta atingida", "Próximo da meta", "Abaixo da meta", "Meta não cadastrada"], key="seller_performance_status"
    )
    seller_top = sf3.selectbox("Exibir", [10, 15, 20, 30], index=1, key="seller_performance_top")
    if seller_search:
        seller_perf = seller_perf[seller_perf["Vendedor"].astype(str).str.contains(seller_search, case=False, na=False)]
    if seller_status:
        seller_perf = seller_perf[seller_perf["Status"].isin(seller_status)]
    seller_perf = seller_perf.head(seller_top)

    if seller_perf.empty:
        st.info("Nenhum integrante da aba Metas com faturamento foi encontrado para os filtros selecionados.")
    else:
        seller_plot = seller_perf.sort_values("Atingimento")
        seller_plot["Nome"] = seller_plot["Vendedor"].map(lambda x: short_label(x, 34))
        fig = go.Figure(go.Bar(
            x=seller_plot["Atingimento"], y=seller_plot["Nome"], orientation="h",
            marker_color=[BLUE if value >= 1 else CYAN for value in seller_plot["Atingimento"]],
            text=seller_plot["Atingimento"].map(pct), textposition="outside", cliponaxis=False,
            customdata=seller_plot[["Vendedor", "Faturamento", "Meta", "Desvio"]],
            hovertemplate="%{customdata[0]}<br>Faturamento: R$ %{customdata[1]:,.2f}<br>Meta: R$ %{customdata[2]:,.2f}<br>Desvio: R$ %{customdata[3]:,.2f}<extra></extra>",
        ))
        fig.add_vline(x=1, line_dash="dot", line_color=NAVY, line_width=2)
        fig.update_layout(title="Atingimento por vendedor", showlegend=False)
        hide_value_axis(fig, "x")
        st.plotly_chart(plot_layout(fig, min(max(320, 34 * len(seller_plot)), 470), False), width="stretch", config={"displayModeBar": False})

        seller_view = seller_perf[["Vendedor", "Faturamento", "Meta", "Atingimento", "Desvio", "Status"]].copy()
        seller_export = seller_view.copy()
        seller_view["Faturamento"] = seller_view["Faturamento"].map(brl)
        seller_view["Meta"] = seller_view["Meta"].map(brl)
        seller_view["Atingimento"] = seller_view["Atingimento"].map(pct)
        seller_view["Desvio"] = seller_view["Desvio"].map(brl)
        clean_table(seller_view, height=360)
        st.download_button(
            "Exportar desempenho da equipe", dataframe_download(seller_export, "Desempenho"),
            file_name=f"desempenho_metas_{scope_choice.lower()}_{start_month}_{end_month}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# =========================================================
# LINHAS DE NEGÓCIO — SOMENTE DIRETORIA
# =========================================================
elif page == "Linhas de negócio" and is_director:
    section_header("Linhas de negócio", "Comparativo direto de caixa", "Diretoria")
    selected_codes = st.multiselect(
        "Linhas exibidas",
        options=LINES,
        default=LINES,
        format_func=line_label,
        key="filter_lines_business",
    )
    lines_view = lines_table[lines_table["Código"].isin(selected_codes)].copy() if selected_codes else lines_table.copy()

    selected_revenue = float(lines_view["Receitas Recebidas"].sum())
    selected_cost = float(lines_view["Custos Diretos Pagos"].sum())
    selected_result = float(lines_view["Resultado Direto de Caixa"].sum())
    selected_margin = safe_div(selected_result, selected_revenue)

    r1, r2, r3, r4 = st.columns(4)
    with r1: card("Receitas diretas", brl(selected_revenue), "Linhas selecionadas", BLUE)
    with r2: card("Custos diretos", brl(selected_cost), "Pagamentos diretos", RED)
    with r3: card("Resultado direto", brl(selected_result), "Receitas menos custos", BLUE if selected_result >= 0 else CYAN)
    with r4: card("Margem direta", pct(selected_margin), "Resultado ÷ receitas", NAVY)

    cols = st.columns(min(max(len(lines_view), 1), 4))
    for idx, (_, row) in enumerate(lines_view.iterrows()):
        with cols[idx % len(cols)]:
            card(row["Linha"], brl(row["Resultado Direto de Caixa"]), f"Margem: {pct(row['Margem Direta de Caixa'])}", LINE_COLORS[row["Código"]] if row["Resultado Direto de Caixa"] >= 0 else CYAN)

    c1, c2 = st.columns([1.2, 1])
    with c1:
        fig = line_revenue_cost_dumbbell(lines_view, "Receitas x despesas por linha")
        st.plotly_chart(plot_layout(fig, 365), width="stretch", config={"displayModeBar": False})
    with c2:
        fig = line_result_lollipop(lines_view, "Resultado direto por linha")
        st.plotly_chart(plot_layout(fig, 365, False), width="stretch", config={"displayModeBar": False})
    export_lines = lines_view.drop(columns=["Código"]).rename(columns={
        "Receitas Recebidas": "Receitas Operacionais Diretas",
        "Custos Diretos Pagos": "Custos Operacionais Diretos",
    })
    view = export_lines.copy()
    for c in ["Receitas Operacionais Diretas", "Custos Operacionais Diretos", "Resultado Direto de Caixa", "Faturamento", "Meta", "Desvio da Meta", "Projeção Anual", "Inadimplência"]:
        if c in view.columns:
            view[c] = view[c].map(brl)
    for c in ["Margem Direta de Caixa", "Margem de Contribuição Direta", "Atingimento da Meta", "Atingimento Projetado", "Conversão em Caixa"]:
        if c in view.columns:
            view[c] = view[c].map(pct)
    clean_table(view, height=280)
    st.download_button(
        "Exportar comparativo",
        dataframe_download(export_lines, "Linhas"),
        file_name=f"resultado_caixa_por_linha_{start_month}_{end_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# =========================================================
# RECEBIMENTOS E INADIMPLÊNCIA
# =========================================================
elif page == "Recebimentos & inadimplência":
    section_header("Conversão e risco de recebimento", "Entradas realizadas, previsão e títulos vencidos", scope_text)

    if scope_choice == "CONSOLIDADO":
        r1, r2, r3, r4 = st.columns(4)
        with r1: card("Recebimento realizado", brl(company_totals["Recebimento Realizado"]), "Base de performance de recebimento", BLUE)
        with r2: card("Recebimento previsto", brl(company_totals["Recebimento Previsto"]), "Valor programado para o período", NAVY)
        with r3: card("Performance", pct(company_totals["Performance de Recebimento"]), "Realizado ÷ previsto", GREEN if company_totals["Performance de Recebimento"] >= .9 else CYAN)
        gap = company_totals["Recebimento Realizado"] - company_totals["Recebimento Previsto"]
        with r4: card("Diferença", brl(gap), "Realizado menos previsto", BLUE if gap >= 0 else CYAN)
        fig = go.Figure()
        fig.add_bar(x=company_monthly["Mês Texto"], y=company_monthly["Recebimento Previsto"], name="Previsto", marker_color="#A8C6DF",
                    text=company_monthly["Recebimento Previsto"].map(compact_money), textposition="outside", cliponaxis=False)
        fig.add_scatter(x=company_monthly["Mês Texto"], y=company_monthly["Recebimento Realizado"], name="Realizado", mode="lines+markers",
                        line=dict(color=TEAL, width=3), marker=dict(size=7))
        add_point_labels(fig, company_monthly["Mês Texto"], company_monthly["Recebimento Realizado"],
                         company_monthly["Recebimento Realizado"].map(compact_money), color=NAVY)
        fig.update_layout(title="Recebimento previsto x realizado")
        hide_value_axis(fig, "y")
        st.plotly_chart(plot_layout(fig, 355), width="stretch", config={"displayModeBar": False})
    else:
        r1, r2, r3, r4 = st.columns(4)
        with r1: card("Receita operacional direta", brl(line_totals["Receitas Recebidas"]), "Centro de Custos · RECEITA", BLUE)
        with r2: card("Faturamento emitido", brl(line_totals["Faturamento"]), "Base comercial do período", NAVY)
        with r3: card("Conversão em caixa", pct(line_totals["Conversão em Caixa"]), "Receita direta ÷ faturamento", TEAL)
        with r4: card("Resultado direto", brl(line_totals["Resultado Direto de Caixa"]), "Receita direta menos custo direto", BLUE if line_totals["Resultado Direto de Caixa"] >= 0 else CYAN)

    section_header("Inadimplência", "Relatório do CRM de Cobrança")
    inad_page = None if inad_scope is None else inad_scope.copy()
    if inad_page is not None and not inad_page.empty:
        fi1, fi2, fi3, fi4, fi5 = st.columns([1.35, 1.05, .9, 1.05, .58])
        inad_client_search = fi1.text_input("Buscar cliente", placeholder="Nome do cliente", key="filter_inad_client")
        faixa_order = ["Até 30 dias", "31 a 60 dias", "61 a 90 dias", "Acima de 90 dias"]
        faixa_options = [x for x in faixa_order if x in inad_page["_FAIXA"].astype(str).unique()]
        selected_faixas = fi2.multiselect("Faixa de atraso", faixa_options, key="filter_inad_faixa")
        max_days = max(int(pd.to_numeric(inad_page["_DIAS_ATRASO"], errors="coerce").fillna(0).max()), 1)
        min_days = fi3.number_input("Atraso mínimo", min_value=0, max_value=max_days, value=0, step=5, key="filter_inad_days")
        seller_options = sorted(v for v in inad_page["_VENDEDOR"].dropna().astype(str).unique() if norm(v) not in {"", "SEM VENDEDOR"})
        selected_sellers = fi4.multiselect("Vendedor", seller_options, key="filter_inad_seller")
        inad_top_n = fi5.selectbox("Exibir", [10, 12, 15, 20], index=1, key="filter_inad_top")
        if inad_client_search:
            inad_page = inad_page[inad_page["_CLIENTE"].astype(str).str.contains(inad_client_search, case=False, na=False)]
        if selected_faixas:
            inad_page = inad_page[inad_page["_FAIXA"].astype(str).isin(selected_faixas)]
        if selected_sellers:
            inad_page = inad_page[inad_page["_VENDEDOR"].astype(str).isin(selected_sellers)]
        inad_page = inad_page[pd.to_numeric(inad_page["_DIAS_ATRASO"], errors="coerce").fillna(0) >= min_days]

    if inad_page is None:
        st.info("O relatório do CRM de cobrança ainda não foi carregado. O novo CSV reconhece cliente, saldo atual, vencimento, vendedor e gerente; o formato Excel anterior também continua compatível.")
    elif inad_page.empty:
        st.success("Não foram encontrados títulos vencidos para o escopo selecionado.")
    else:
        overdue = float(inad_page["_VALOR_VENCIDO"].sum())
        clients_overdue = int(inad_page["_CLIENTE"].nunique())
        titles_overdue = len(inad_page)
        avg_days = float(np.average(inad_page["_DIAS_ATRASO"], weights=np.maximum(inad_page["_VALOR_VENCIDO"], .01)))
        i1, i2, i3, i4 = st.columns(4)
        with i1: card("Saldo vencido", brl(overdue), "Valor vencido identificado", NAVY)
        with i2: card("Clientes inadimplentes", f"{clients_overdue:,}".replace(",", "."), "Clientes únicos com saldo vencido", BLUE)
        with i3: card("Títulos vencidos", f"{titles_overdue:,}".replace(",", "."), "Quantidade de registros vencidos", NAVY)
        with i4: card("Atraso médio ponderado", f"{avg_days:,.0f} dias".replace(",", "."), "Ponderado pelo valor vencido", TEAL)

        c1, c2 = st.columns(2)
        with c1:
            aging = inad_page.groupby("_FAIXA", as_index=False, observed=False)["_VALOR_VENCIDO"].sum()
            order = ["Até 30 dias", "31 a 60 dias", "61 a 90 dias", "Acima de 90 dias"]
            aging["_FAIXA"] = pd.Categorical(aging["_FAIXA"], order, ordered=True)
            aging = aging.sort_values("_FAIXA")
            fig = px.bar(aging, x="_FAIXA", y="_VALOR_VENCIDO", title="Aging da inadimplência")
            fig.update_traces(marker_color=BLUE, text=aging["_VALOR_VENCIDO"].map(compact_money), textposition="outside", cliponaxis=False)
            hide_value_axis(fig, "y")
            st.plotly_chart(plot_layout(fig, 345, False), width="stretch", config={"displayModeBar": False})
        with c2:
            cli = inad_page.groupby("_CLIENTE", as_index=False)["_VALOR_VENCIDO"].sum().sort_values("_VALOR_VENCIDO", ascending=False).head(inad_top_n)
            cli["Nome"] = cli["_CLIENTE"].map(lambda x: short_label(x, 38))
            p = cli.sort_values("_VALOR_VENCIDO")
            fig = px.bar(p, x="_VALOR_VENCIDO", y="Nome", orientation="h", title="Maiores saldos vencidos", custom_data=["_CLIENTE"])
            fig.update_traces(marker_color=NAVY, text=p["_VALOR_VENCIDO"].map(compact_money), textposition="outside", cliponaxis=False,
                              hovertemplate="%{customdata[0]}<br>Vencido: R$ %{x:,.2f}<extra></extra>")
            hide_value_axis(fig, "x")
            st.plotly_chart(plot_layout(fig, 345, False), width="stretch", config={"displayModeBar": False})

        if is_director and inad_meta and inad_meta.get("resumo_gerentes"):
            crm_summary = pd.DataFrame(inad_meta["resumo_gerentes"])
            crm_summary = crm_summary[crm_summary["Gerente"].map(norm).isin(set(MANAGER_LINE_MAP) | {"GABRIEL", "SEM GERENTE"})]
            crm_summary["Linha"] = crm_summary["Linha"].map(line_label)
            crm_view = crm_summary[["Gerente", "Linha", "Clientes", "Títulos", "Valor", "Maior atraso"]].copy()
            crm_view["Valor"] = crm_view["Valor"].map(brl)
            st.markdown("**Resumo original exportado pelo CRM de cobrança**")
            clean_table(crm_view, height=260)

        detail = inad_page[["_CLIENTE", "_TITULO", "_EMISSAO", "_VENCIMENTO", "_DIAS_ATRASO", "_FAIXA", "_VALOR_VENCIDO", "_VENDEDOR", "_LINHA", "_GERENTE"]].copy()
        detail.columns = ["Cliente", "Título", "Emissão", "Vencimento", "Dias de atraso", "Faixa", "Valor vencido", "Vendedor", "Linha", "Gerente"]
        detail["Linha"] = detail["Linha"].map(line_label)
        detail_view = detail.sort_values("Valor vencido", ascending=False).copy()
        detail_view["Valor vencido"] = detail_view["Valor vencido"].map(brl)
        detail_view["Emissão"] = pd.to_datetime(detail_view["Emissão"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
        detail_view["Vencimento"] = pd.to_datetime(detail_view["Vencimento"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
        clean_table(detail_view, height=370)
        st.download_button("Exportar inadimplência do escopo", dataframe_download(detail, "Inadimplência"),
                           file_name=f"inadimplencia_{scope_choice}_{start_month}_{end_month}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# =========================================================
# CLIENTES
# =========================================================
elif page == "Clientes":
    section_header("Clientes", "Faturamento e recebimento", scope_text)
    client_col = "NOME DO CLIENTE" if "NOME DO CLIENTE" in fat_scope.columns else "CLIENTE"
    client_base = fat_scope.copy()
    received_base = rec_scope.copy()
    segment_col = optional_col(client_base, ["SEGMENTO"])

    f1, f2, f3 = st.columns([1.35, 1, .65])
    client_search = f1.text_input("Buscar cliente", placeholder="Nome ou parte do nome", key="filter_client_search")
    segment_values = sorted(client_base[segment_col].dropna().astype(str).unique()) if segment_col else []
    selected_segments = f2.multiselect("Segmento", segment_values, key="filter_client_segment") if segment_values else []
    top_n = f3.selectbox("Exibir", [10, 15, 20, 30], index=1, key="filter_client_top")

    if client_search:
        client_base = client_base[client_base[client_col].astype(str).str.contains(client_search, case=False, na=False)]
        if "Cliente" in received_base.columns:
            received_base = received_base[received_base["Cliente"].astype(str).str.contains(client_search, case=False, na=False)]
    if selected_segments and segment_col:
        client_base = client_base[client_base[segment_col].astype(str).isin(selected_segments)]

    if client_base.empty and received_base.empty:
        st.info("Nenhum cliente encontrado com os filtros selecionados.")
    else:
        billed = client_base.groupby(client_col, as_index=False).agg(
            Faturamento=("_VALOR", "sum"),
            Notas=("Nota Fiscal", "nunique") if "Nota Fiscal" in client_base.columns else ("_VALOR", "size"),
        )
        received = received_base.groupby("Cliente", as_index=False)["_VALOR"].sum().rename(columns={"Cliente": "Cliente Recebimento", "_VALOR": "Recebido"}) if not received_base.empty else pd.DataFrame(columns=["Cliente Recebimento", "Recebido"])
        total_billed = float(billed["Faturamento"].sum()) if not billed.empty else 0.0
        billed["Participação"] = np.where(total_billed != 0, billed["Faturamento"] / total_billed, 0)
        billed = billed.sort_values("Faturamento", ascending=False)

        c1, c2, c3, c4 = st.columns(4)
        with c1: card("Clientes faturados", f"{len(billed):,}".replace(",", "."), "No filtro atual", BLUE)
        with c2: card("Concentração Top 5", pct(float(billed.head(5)["Participação"].sum()) if not billed.empty else 0), "Participação no faturamento", NAVY)
        with c3: card("Ticket médio por nota", brl(safe_div(total_billed, float(billed["Notas"].sum()) if not billed.empty else 0)), "Faturamento ÷ notas", TEAL)
        with c4: card("Clientes com recebimento", f"{received_base['Cliente'].nunique():,}".replace(",", ".") if "Cliente" in received_base.columns else "0", "Na base de caixa", BLUE)

        p1, p2 = st.columns(2)
        with p1:
            top = billed.head(top_n).copy()
            top["Nome"] = top[client_col].map(lambda x: short_label(x, 38))
            p = top.sort_values("Faturamento")
            fig = px.bar(p, x="Faturamento", y="Nome", orientation="h", title="Clientes por faturamento", custom_data=[client_col])
            fig.update_traces(marker_color=BLUE, text=p["Faturamento"].map(compact_money), textposition="outside", cliponaxis=False,
                              hovertemplate="%{customdata[0]}<br>Faturamento: R$ %{x:,.2f}<extra></extra>")
            hide_value_axis(fig, "x")
            st.plotly_chart(plot_layout(fig, 410, False), width="stretch", config={"displayModeBar": False})
        with p2:
            top = received.sort_values("Recebido", ascending=False).head(top_n).copy()
            top["Nome"] = top["Cliente Recebimento"].map(lambda x: short_label(x, 38)) if not top.empty else []
            p = top.sort_values("Recebido")
            fig = px.bar(p, x="Recebido", y="Nome", orientation="h", title="Clientes por recebimento", custom_data=["Cliente Recebimento"])
            fig.update_traces(marker_color=TEAL, text=p["Recebido"].map(compact_money), textposition="outside", cliponaxis=False,
                              hovertemplate="%{customdata[0]}<br>Recebido: R$ %{x:,.2f}<extra></extra>")
            hide_value_axis(fig, "x")
            st.plotly_chart(plot_layout(fig, 410, False), width="stretch", config={"displayModeBar": False})

        view = billed.head(100).rename(columns={client_col: "Cliente"}).copy()
        view["Faturamento"] = view["Faturamento"].map(brl)
        view["Participação"] = view["Participação"].map(pct)
        clean_table(view, height=430)

# =========================================================
# PRODUTOS
# =========================================================
elif page == "Produtos":
    section_header("Produtos", "Faturamento, volume e carteira", scope_text)
    if fat_scope.empty:
        st.info("Não há faturamento no período e escopo selecionados.")
    else:
        product_col = optional_col(fat_scope, ["PRODUTO", "DESCRIÇÃO", "LINHA DE PRODUTO", "ITEM"])
        qty_col = optional_col(fat_scope, ["QUANTIDADE", "QTD", "QTDE", "QTD FATURADA", "QUANTIDADE FATURADA"])
        client_col = "NOME DO CLIENTE" if "NOME DO CLIENTE" in fat_scope.columns else "CLIENTE"
        segment_col = optional_col(fat_scope, ["SEGMENTO"])
        if product_col is None:
            st.info("A base de faturamento não possui uma coluna de produto reconhecida.")
        else:
            prod_base = fat_scope.copy()
            prod_base["_PRODUTO"] = prod_base[product_col].fillna("Não informado").astype(str).str.strip()
            prod_base["_QTD"] = to_number(prod_base[qty_col]) if qty_col else 1.0
            prod_base.loc[prod_base["_QTD"] <= 0, "_QTD"] = 1.0

            f1, f2, f3, f4 = st.columns([1.25, 1, 1.15, .6])
            product_search = f1.text_input("Buscar produto", placeholder="Código ou descrição", key="filter_product_search")
            segment_values = sorted(prod_base[segment_col].dropna().astype(str).unique()) if segment_col else []
            selected_segments = f2.multiselect("Segmento", segment_values, key="filter_product_segment") if segment_values else []
            client_search = f3.text_input("Buscar cliente", placeholder="Nome do cliente", key="filter_product_client")
            top_n = f4.selectbox("Exibir", [10, 15, 20, 30], index=1, key="filter_product_top")

            if product_search:
                prod_base = prod_base[prod_base["_PRODUTO"].str.contains(product_search, case=False, na=False)]
            if selected_segments and segment_col:
                prod_base = prod_base[prod_base[segment_col].astype(str).isin(selected_segments)]
            if client_search:
                prod_base = prod_base[prod_base[client_col].astype(str).str.contains(client_search, case=False, na=False)]

            if prod_base.empty:
                st.info("Nenhum produto encontrado com os filtros selecionados.")
            else:
                nf_col = "Nota Fiscal" if "Nota Fiscal" in prod_base.columns else None
                agg_map = {
                    "Faturamento": ("_VALOR", "sum"),
                    "Quantidade": ("_QTD", "sum"),
                    "Clientes": (client_col, "nunique"),
                    "Notas": (nf_col, "nunique") if nf_col else ("_VALOR", "size"),
                }
                products = prod_base.groupby("_PRODUTO", as_index=False).agg(**agg_map)
                total_products = float(products["Faturamento"].sum())
                products["Participação"] = np.where(total_products != 0, products["Faturamento"] / total_products, 0)
                products["Preço médio"] = np.where(products["Quantidade"] != 0, products["Faturamento"] / products["Quantidade"], 0)
                products = products.sort_values("Faturamento", ascending=False)
                top_product = products.iloc[0]
                top10_share = float(products.head(10)["Faturamento"].sum() / total_products) if total_products else 0

                k1, k2, k3, k4, k5 = st.columns(5)
                with k1: card("Faturamento", brl(total_products), "Produtos filtrados", BLUE)
                with k2: card("Produtos ativos", f"{len(products):,}".replace(",", "."), "Itens com faturamento", NAVY)
                with k3: card("Quantidade", f"{products['Quantidade'].sum():,.0f}".replace(",", "."), qty_col or "Linhas faturadas", TEAL)
                with k4: card("Produto líder", short_label(top_product["_PRODUTO"], 24), brl(top_product["Faturamento"]), BLUE)
                with k5: card("Concentração Top 10", pct(top10_share), "Participação no faturamento", NAVY)

                c1, c2 = st.columns([1.25, 1])
                with c1:
                    top = products.head(top_n).sort_values("Faturamento")
                    top["Produto"] = top["_PRODUTO"].map(lambda x: short_label(x, 42))
                    fig = px.bar(top, x="Faturamento", y="Produto", orientation="h", title="Produtos por faturamento", custom_data=["_PRODUTO"])
                    fig.update_traces(marker_color=BLUE, text=top["Faturamento"].map(compact_money), textposition="outside", cliponaxis=False,
                                      hovertemplate="%{customdata[0]}<br>Faturamento: R$ %{x:,.2f}<extra></extra>")
                    hide_value_axis(fig, "x")
                    st.plotly_chart(plot_layout(fig, 345, False), width="stretch", config={"displayModeBar": False})
                with c2:
                    qty_top = products.sort_values("Quantidade", ascending=False).head(top_n).sort_values("Quantidade")
                    qty_top["Produto"] = qty_top["_PRODUTO"].map(lambda x: short_label(x, 34))
                    fig = px.bar(qty_top, x="Quantidade", y="Produto", orientation="h", title="Produtos por volume")
                    fig.update_traces(marker_color=TEAL, text=qty_top["Quantidade"].map(lambda v: f"{v:,.0f}".replace(",", ".")), textposition="outside", cliponaxis=False)
                    hide_value_axis(fig, "x")
                    st.plotly_chart(plot_layout(fig, 345, False), width="stretch", config={"displayModeBar": False})

                table = products.rename(columns={"_PRODUTO": "Produto"})[["Produto", "Faturamento", "Participação", "Quantidade", "Preço médio", "Clientes", "Notas"]].copy()
                table_view = table.copy()
                table_view["Faturamento"] = table_view["Faturamento"].map(brl)
                table_view["Participação"] = table_view["Participação"].map(pct)
                table_view["Quantidade"] = table_view["Quantidade"].map(lambda v: f"{v:,.0f}".replace(",", "."))
                table_view["Preço médio"] = table_view["Preço médio"].map(brl)
                clean_table(table_view, height=470)
                st.download_button(
                    "Exportar produtos",
                    dataframe_download(table, "Produtos"),
                    file_name=f"produtos_{scope_choice}_{start_month}_{end_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

# =========================================================
# VALIDAÇÃO DOS DADOS — SOMENTE CONTROLADORIA
# =========================================================
elif page == "Validação dos dados" and can_validate:
    section_header("Validação dos dados", "Conciliação técnica", "Faturamento, linhas, datas e cruzamentos sem alterar os indicadores")

    validation, validation_lines, validation_managers, validation_monthly = billing_reconciliation(
        fat, start_month, end_month
    )
    tolerance = 0.01
    reconciled = abs(float(validation["technical_difference"])) <= tolerance

    a1, a2, a3, a4 = st.columns(4)
    with a1: card("Faturamento consolidado", brl(validation["total"]), "Soma de VALOR BRUTO no período", BLUE)
    with a2: card("Quatro linhas oficiais", brl(validation["sum_lines"]), "Somente Celso, Amauri, Renato e Ronaldo", NAVY)
    with a3: card("Outras gerências", brl(validation["unclassified_value"]), "Não incorporadas à linha de outro gestor", CYAN)
    with a4: card("Diferença técnica", brl(validation["technical_difference"]), "Consolidado menos linhas oficiais e outras gerências", BLUE if reconciled else RED)

    if reconciled:
        st.success("Faturamento conciliado: consolidado = quatro linhas oficiais + outras gerências.")
    else:
        st.error(
            f"Permanece uma diferença técnica de {brl(validation['technical_difference'])}. "
            "Revise valores, competência e classificação das linhas."
        )

    v1, v2 = st.columns([1.15, 1])
    with v1:
        section_header("Faturamento por linha oficial", badge="2026")
        view = validation_lines.copy()
        for col in ["Faturamento", "Via contingência"]:
            view[col] = view[col].map(brl)
        clean_table(view.drop(columns=["Código"]), height=285)
    with v2:
        section_header("Faturamento por gerente informado", badge="BASE BI")
        manager_view = validation_managers.copy()
        if not manager_view.empty:
            manager_view["Faturamento"] = manager_view["Faturamento"].map(brl)
            clean_table(manager_view, height=285)
        else:
            st.info("A BASE BI não possui coluna de gerente reconhecida.")

    section_header("Conciliação mensal", badge=period_text)
    month_view = validation_monthly.copy()
    for col in ["Consolidado", "Quatro linhas oficiais", "Outras gerências", "Diferença técnica"]:
        month_view[col] = month_view[col].map(brl)
    clean_table(month_view, height=300)

    section_header("Cruzamento financeiro da REV2026", badge=period_text)
    financial_validation = financial_crosscheck(receitas, despesas, custos, start_month, end_month)
    financial_view = financial_validation.copy()
    for col in ["Resumo financeiro", "Centro de Custos", "Diferença", "Direto nas quatro linhas", "Geral / não classificado"]:
        financial_view[col] = financial_view[col].map(brl)
    clean_table(financial_view, height=210)
    st.caption(
        "A diferença entre os resumos e o Centro de Custos é exibida para controle da origem. "
        "Ela não é compensada automaticamente pelo app, e o resultado das linhas permanece baseado no Centro de Custos direto."
    )

    m1, m2, m3 = st.columns(3)
    with m1: card("Cobertura dos recebimentos", pct(float(receipt_match_stats.get("coverage_value", 0))), "Percentual do valor operacional ligado com segurança à BASE BI", BLUE)
    with m2: card("Recebimentos em fallback", brl(float(receipt_match_stats.get("fallback_value", 0))), "Classificados por tipo quando o cliente não foi localizado", CYAN)
    with m3: card("Ajustes Microtech", brl(float(cc_microtech_stats.get("value", 0))), f"{int(cc_microtech_stats.get('count', 0))} recebimento(s) comprovado(s) realocado(s)", NAVY)

    section_header("Integridade da BASE BI", badge=base_name)
    b1, b2, b3, b4 = st.columns(4)
    with b1: card("Faturamento validado", brl(float(billing_audit.get('value_2026', 0))), f"{int(billing_audit.get('rows_2026', len(fat))):,} linhas na competência 2026".replace(",", "."), NAVY)
    with b2: card("Emissões de outro ano", brl(float(billing_audit.get('value_emission_outside_2026', 0))), f"{int(billing_audit.get('rows_emission_outside_2026', 0)):,} linhas mantidas em 2026 pela coluna MÊS".replace(",", "."), BLUE)
    with b3: card("Linhas com valor zero", f"{int(billing_audit.get('zero_value_rows_2026', 0)):,}".replace(",", "."), "Permanecem na base, mas não alteram o total", CYAN)
    with b4: card("Linhas sem nota fiscal", f"{int(billing_audit.get('missing_note_rows_2026', 0)):,}".replace(",", "."), "Podem exigir conferência na origem", CYAN)

    st.caption(
        f"Competência oficial: {billing_audit.get('month_column', 'MÊS')} · "
        f"Data de auditoria: {billing_audit.get('date_column', 'não identificada')} · "
        f"Coluna de faturamento: {billing_audit.get('value_column', 'não identificada')} · "
        "o painel soma VALOR BRUTO pela coluna MÊS da BASE BI e considera somente as competências de 2026. "
        f"A DT Emissão é usada apenas quando MÊS estiver vazio ({int(billing_audit.get('rows_month_fallback_2026', 0))} linha(s) em contingência). "
        f"Valores negativos: {int(billing_audit.get('negative_value_rows_2026', 0))} linha(s), "
        f"total de {brl(float(billing_audit.get('negative_value_2026', 0)))}. "
        f"Linhas sem competência: {int(billing_audit.get('rows_without_month', 0))}."
    )


# =========================================================
# CENTRO DE CUSTOS
# =========================================================
elif page == "Centro de custos":
    section_header("Centro de custos", badge=scope_text)

    # Resultado alternativo, calculado pela distribuição administrativa da própria REV2026.
    # Não substitui nem altera o resultado direto exibido nas demais páginas.
    allocated_result = allocated_cost_center_result(custos_rateados, start_month, end_month)
    allocated_view = (
        allocated_result.copy()
        if scope_choice == "CONSOLIDADO"
        else allocated_result[allocated_result["Código"] == scope_choice].copy()
    )

    section_header(
        "Resultado após rateio administrativo",
        badge="Centro de Custos Rateado",
    )
    st.caption(
        "Cenário calculado por receitas rateadas menos despesas rateadas. "
        "O resultado direto continua sendo apurado exclusivamente pelo Centro de Custos original."
    )

    if allocated_view.empty:
        st.info("Não foram encontrados lançamentos rateados para o período selecionado.")
    else:
        card_columns = st.columns(len(allocated_view))
        for col, (_, row) in zip(card_columns, allocated_view.iterrows()):
            value = float(row["Resultado após rateio"])
            with col:
                card(
                    row["Linha"], brl(value),
                    f"Margem após rateio: {pct(float(row['Margem após rateio']))}",
                    BLUE if value >= 0 else RED,
                )

        chart_allocated = allocated_view.sort_values("Resultado após rateio", ascending=False).copy()
        chart_allocated["Cor"] = np.where(chart_allocated["Resultado após rateio"] >= 0, BLUE, RED)
        chart_allocated["Texto"] = chart_allocated["Resultado após rateio"].map(compact_money)
        fig_allocated = go.Figure(go.Bar(
            x=chart_allocated["Linha"],
            y=chart_allocated["Resultado após rateio"],
            marker_color=chart_allocated["Cor"],
            text=chart_allocated["Texto"],
            textposition="outside",
            cliponaxis=False,
            customdata=chart_allocated[["Receitas após rateio", "Despesas após rateio"]],
            hovertemplate=(
                "%{x}<br>Receitas rateadas: R$ %{customdata[0]:,.2f}"
                "<br>Despesas rateadas: R$ %{customdata[1]:,.2f}"
                "<br>Resultado: R$ %{y:,.2f}<extra></extra>"
            ),
        ))
        fig_allocated.add_hline(y=0, line_color="#9FB3C8", line_width=1)
        fig_allocated.update_layout(title="Resultado das linhas após o rateio administrativo")
        hide_value_axis(fig_allocated, "y")
        st.plotly_chart(
            plot_layout(fig_allocated, 360, False),
            width="stretch", config={"displayModeBar": False},
        )

        allocated_table = allocated_view.drop(columns=["Código"]).copy()
        for column in ["Receitas após rateio", "Despesas após rateio", "Resultado após rateio"]:
            allocated_table[column] = allocated_table[column].map(brl)
        allocated_table["Margem após rateio"] = allocated_table["Margem após rateio"].map(pct)
        clean_table(allocated_table, height=min(240, 54 + 36 * len(allocated_table)), max_rows=10)

    st.markdown("---")
    section_header("Movimentação pelo centro de custos direto", badge="Resultado direto")
    cc_base = period_filter(custos, start_month, end_month).copy()
    cc_base["Movimento"] = np.where(cc_base["_EMPRESA_N"] == "RECEITA", "Receita", "Despesa")
    cc_base["Classificação"] = np.where(
        cc_base["_GRUPO_N"].str.contains("NAO OPERACIONAIS", na=False),
        "Não operacional", "Operacional",
    )
    cc_base["Departamento"] = cc_base["CENTRO DE CUSTOS"].map(cost_center_label)

    # Para gestores, o departamento é sempre o centro de custo direto da própria linha.
    # A coluna de rateio não participa dos cálculos nem da exibição desta página.
    if scope_choice != "CONSOLIDADO":
        cc_base = cc_base[cc_base["_CC_N"] == scope_choice].copy()

    f1, f2, f3, f4 = st.columns([1.15, 1.1, 1.15, .75])
    movement_filter = f1.multiselect(
        "Movimento", ["Receita", "Despesa"], default=["Receita", "Despesa"], key="cc_movement"
    )
    class_filter = f2.multiselect(
        "Classificação", ["Operacional", "Não operacional"],
        default=["Operacional", "Não operacional"], key="cc_classification"
    )
    department_options = sorted(cc_base["Departamento"].dropna().unique())
    if is_director:
        department_filter = f3.multiselect("Departamento", department_options, key="cc_department")
    else:
        department_filter = [line_label(scope_choice)]
        f3.text_input("Departamento", value=line_label(scope_choice), disabled=True, key="cc_department_locked")
    top_n = f4.selectbox("Exibir", [8, 10, 15, 20, 30], index=2, key="cc_top")

    f5, f6 = st.columns([1.3, 1.2])
    party_search = f5.text_input("Buscar fornecedor ou cliente", placeholder="Nome ou código", key="cc_party")
    nature_options = sorted(cc_base["PAI"].dropna().astype(str).unique())
    nature_filter = f6.multiselect("Natureza", nature_options, key="cc_nature")

    if movement_filter:
        cc_base = cc_base[cc_base["Movimento"].isin(movement_filter)]
    if class_filter:
        cc_base = cc_base[cc_base["Classificação"].isin(class_filter)]
    if department_filter and is_director:
        cc_base = cc_base[cc_base["Departamento"].isin(department_filter)]
    if nature_filter:
        cc_base = cc_base[cc_base["PAI"].astype(str).isin(nature_filter)]

    party_col = optional_col(cc_base, ["Codigo-Nome do Fornecedor", "FORNECEDOR", "NOME DO FORNECEDOR"]) or "Codigo-Nome do Fornecedor"
    if party_search:
        cc_base = cc_base[cc_base[party_col].astype(str).str.contains(party_search, case=False, na=False)]

    if cc_base.empty:
        st.info("Nenhum lançamento encontrado para os filtros selecionados.")
    else:
        total_revenue = float(cc_base.loc[cc_base["Movimento"] == "Receita", "_VALOR"].sum())
        total_expense = float(cc_base.loc[cc_base["Movimento"] == "Despesa", "_VALOR"].sum())
        net_result = total_revenue - total_expense
        departments = int(cc_base["Departamento"].nunique())

        k1, k2, k3, k4 = st.columns(4)
        with k1: card("Receitas", brl(total_revenue), "Lançamentos recebidos", BLUE)
        with k2: card("Despesas", brl(total_expense), "Pagamentos realizados", RED)
        with k3: card("Saldo", brl(net_result), "Receitas menos despesas", BLUE if net_result >= 0 else RED)
        with k4: card("Centros de custos", f"{departments:,}".replace(",", "."), "Departamentos no filtro", NAVY)

        monthly = (
            cc_base.groupby(["_MES", "Movimento"], as_index=False)["_VALOR"].sum()
            .pivot(index="_MES", columns="Movimento", values="_VALOR").fillna(0).reset_index()
        )
        for col in ["Receita", "Despesa"]:
            if col not in monthly:
                monthly[col] = 0.0
        monthly["Saldo"] = monthly["Receita"] - monthly["Despesa"]
        monthly["Mês"] = monthly["_MES"].map(month_label)

        c1, c2 = st.columns([1.25, 1])
        with c1:
            fig = go.Figure()
            fig.add_bar(
                x=monthly["Mês"], y=monthly["Receita"], name="Receitas", marker_color=BLUE,
                text=monthly["Receita"].map(compact_money), textposition="outside", cliponaxis=False,
            )
            fig.add_bar(
                x=monthly["Mês"], y=monthly["Despesa"], name="Despesas", marker_color=RED,
                text=monthly["Despesa"].map(compact_money), textposition="inside", insidetextanchor="end",
                textfont=dict(color=WHITE), cliponaxis=False,
            )
            fig.add_scatter(
                x=monthly["Mês"], y=monthly["Saldo"], name="Saldo", mode="lines+markers",
                line=dict(color=NAVY, width=3), marker=dict(size=8),
            )
            add_point_labels(fig, monthly["Mês"], monthly["Saldo"], monthly["Saldo"].map(compact_money), color=NAVY)
            fig.update_layout(title="Movimento mensal por centro de custos", barmode="group")
            hide_value_axis(fig, "y")
            st.plotly_chart(plot_layout(fig, 365), width="stretch", config={"displayModeBar": False})

        with c2:
            by_department = (
                cc_base.groupby(["Departamento", "Movimento"], as_index=False)["_VALOR"].sum()
                .pivot(index="Departamento", columns="Movimento", values="_VALOR").fillna(0).reset_index()
            )
            for col in ["Receita", "Despesa"]:
                if col not in by_department:
                    by_department[col] = 0.0
            by_department["Movimentação"] = by_department["Receita"] + by_department["Despesa"]
            by_department = (
                by_department.sort_values("Movimentação", ascending=False)
                .head(top_n)
                .sort_values("Movimentação")
            )
            by_department["Departamento gráfico"] = by_department["Departamento"].map(lambda x: short_label(x, 30))
            revenue_labels = [compact_money(v) if float(v) > 0 else "" for v in by_department["Receita"]]
            expense_labels = [compact_money(v) if float(v) > 0 else "" for v in by_department["Despesa"]]

            fig = go.Figure()
            fig.add_bar(
                x=by_department["Receita"], y=by_department["Departamento gráfico"], orientation="h",
                name="Receitas", marker_color=BLUE, text=revenue_labels,
                textposition="outside", cliponaxis=False,
                customdata=by_department[["Departamento"]],
                hovertemplate="%{customdata[0]}<br>Receitas: R$ %{x:,.2f}<extra></extra>",
            )
            fig.add_bar(
                x=by_department["Despesa"], y=by_department["Departamento gráfico"], orientation="h",
                name="Despesas", marker_color=RED, text=expense_labels,
                textposition="outside", cliponaxis=False,
                customdata=by_department[["Departamento"]],
                hovertemplate="%{customdata[0]}<br>Despesas: R$ %{x:,.2f}<extra></extra>",
            )
            fig.update_layout(
                title="Receitas e despesas por departamento",
                barmode="group",
                bargap=.30,
                bargroupgap=.12,
            )
            hide_value_axis(fig, "x")
            fig.update_yaxes(categoryorder="array", categoryarray=by_department["Departamento gráfico"].tolist())
            st.plotly_chart(
                plot_layout(fig, min(max(320, 40 * len(by_department)), 500), True),
                width="stretch",
                config={"displayModeBar": False},
            )

        p1, p2 = st.columns(2)
        with p1:
            revenues = (
                cc_base[cc_base["Movimento"] == "Receita"].groupby("PAI", as_index=False)["_VALOR"].sum()
                .sort_values("_VALOR", ascending=False).head(top_n).sort_values("_VALOR")
            )
            if not revenues.empty:
                revenues["Natureza"] = revenues["PAI"].map(lambda x: short_label(x, 38))
                fig = px.bar(revenues, x="_VALOR", y="Natureza", orientation="h", title="Principais receitas", custom_data=["PAI"])
                fig.update_traces(marker_color=BLUE, text=revenues["_VALOR"].map(compact_money), textposition="outside", cliponaxis=False,
                                  hovertemplate="%{customdata[0]}<br>Valor: R$ %{x:,.2f}<extra></extra>")
                hide_value_axis(fig, "x")
                st.plotly_chart(plot_layout(fig, 380, False), width="stretch", config={"displayModeBar": False})
        with p2:
            expenses = (
                cc_base[cc_base["Movimento"] == "Despesa"].groupby("PAI", as_index=False)["_VALOR"].sum()
                .sort_values("_VALOR", ascending=False).head(top_n).sort_values("_VALOR")
            )
            if not expenses.empty:
                expenses["Natureza"] = expenses["PAI"].map(lambda x: short_label(x, 38))
                fig = px.bar(expenses, x="_VALOR", y="Natureza", orientation="h", title="Principais despesas", custom_data=["PAI"])
                fig.update_traces(marker_color=RED, text=expenses["_VALOR"].map(compact_money), textposition="outside", cliponaxis=False,
                                  hovertemplate="%{customdata[0]}<br>Valor: R$ %{x:,.2f}<extra></extra>")
                hide_value_axis(fig, "x")
                st.plotly_chart(plot_layout(fig, 380, False), width="stretch", config={"displayModeBar": False})

        detail = cc_base[["_MES", "Departamento", "Movimento", "Classificação", "GRUPO", "PAI", party_col, "_VALOR"]].copy()
        detail["Mês"] = detail["_MES"].map(month_label)
        detail = detail.drop(columns=["_MES"]).rename(columns={party_col: "Fornecedor / cliente", "_VALOR": "Valor"})
        detail = detail[["Mês", "Departamento", "Movimento", "Classificação", "GRUPO", "PAI", "Fornecedor / cliente", "Valor"]]
        detail = detail.sort_values("Valor", ascending=False)
        detail_view = detail.copy()
        detail_view["Valor"] = detail_view["Valor"].map(brl)
        clean_table(detail_view, height=470, max_rows=300)

        st.download_button(
            "Exportar centro de custos",
            dataframe_download(detail, "Centro de Custos"),
            file_name=f"centro_de_custos_{scope_choice}_{start_month}_{end_month}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
