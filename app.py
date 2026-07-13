from __future__ import annotations

import hashlib
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
from openpyxl import load_workbook


# =========================================================
# CONFIGURAÇÃO VISUAL
# =========================================================
st.set_page_config(
    page_title="First Intelligence | KPIs de Caixa",
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
    .block-container {{ max-width:1540px; padding-top:.85rem; padding-bottom:2.3rem; }}
    [data-testid="stSidebar"] {{ background:linear-gradient(180deg,#071B33 0%,#0B2F55 100%); border-right:0; }}
    [data-testid="stSidebar"] .block-container {{ padding-top:1rem; }}
    [data-testid="stSidebar"] label, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span {{ color:#EAF4FB; }}
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] h4 {{ color:#FFFFFF; }}
    [data-testid="stSidebar"] [data-baseweb="select"] > div,
    [data-testid="stSidebar"] input {{ background:#FFFFFF; color:#071B33; border-radius:10px; }}
    [data-testid="stSidebar"] [data-baseweb="select"] span,
    [data-testid="stSidebar"] input {{ color:#071B33 !important; }}
    [data-testid="stSidebar"] [role="radiogroup"] label {{ background:rgba(255,255,255,.045); margin-bottom:4px; border:1px solid rgba(255,255,255,.06); }}
    [data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) {{ background:rgba(79,154,209,.22); border-color:rgba(185,217,240,.42); }}
    [data-testid="stPlotlyChart"] {{
        background:linear-gradient(180deg,#FFFFFF 0%,#FCFEFF 100%); border:1px solid #D6E4EF;
        border-radius:22px; padding:8px 10px; box-shadow:0 16px 34px rgba(7,27,51,.07); overflow:hidden;
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
        color:white; border-radius:24px; padding:24px 28px; margin-bottom:16px;
        box-shadow:0 18px 38px rgba(7,27,51,.18);
    }}
    .dashboard-hero::after {{ content:""; position:absolute; width:320px; height:320px; border-radius:50%; right:-130px; top:-185px; background:rgba(255,255,255,.055); }}
    .hero-grid {{ position:relative; z-index:2; display:flex; align-items:center; justify-content:space-between; gap:20px; }}
    .hero-brand {{ font-size:.72rem; font-weight:900; letter-spacing:.28em; color:#B9D9F0; text-transform:uppercase; }}
    .dashboard-hero h1 {{ font-size:clamp(1.42rem,2.2vw,2rem); line-height:1.15; margin:7px 0 6px 0; letter-spacing:-.025em; }}
    .dashboard-hero p {{ margin:0; opacity:.82; font-size:.88rem; }}
    .hero-chips {{ display:flex; flex-wrap:wrap; gap:7px; justify-content:flex-end; }}
    .hero-chip {{ background:rgba(255,255,255,.11); border:1px solid rgba(255,255,255,.20); border-radius:999px; padding:7px 11px; font-size:.71rem; font-weight:800; white-space:nowrap; }}

    .kpi-card {{
        background:linear-gradient(145deg,#FFFFFF 0%,#FBFDFF 100%); border:1px solid #DDE7F0; border-radius:19px; padding:16px 17px 15px 17px;
        min-height:138px; height:100%; position:relative; overflow:hidden;
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

    .section-head {{ display:flex; align-items:center; justify-content:space-between; gap:15px; margin:22px 0 13px 0; padding:0 0 10px 2px; border-bottom:1px solid #DDE7F0; }}
    .section-head h3 {{ color:{NAVY}; font-size:1.13rem; font-weight:900; margin:0; letter-spacing:-.015em; }}
    .section-head p {{ color:{GRAY}; font-size:.78rem; margin:4px 0 0 0; display:none; }}
    .section-badge {{ display:inline-flex; align-items:center; border-radius:999px; padding:5px 9px; background:#EAF3FA; color:{BLUE}; font-size:.67rem; font-weight:900; text-transform:uppercase; letter-spacing:.04em; white-space:nowrap; }}

    .insight-grid {{ display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:10px; margin-top:4px; }}
    .insight-card {{ background:white; border:1px solid {BORDER}; border-left:4px solid {BLUE}; border-radius:14px; padding:12px 14px; box-shadow:0 3px 12px rgba(7,27,51,.035); }}
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
    .user-pill {{ display:flex; align-items:center; justify-content:space-between; gap:10px; background:linear-gradient(180deg,#F9FBFD 0%,#F2F6FA 100%); border:1px solid {BORDER}; border-radius:14px; padding:10px 12px; margin:8px 0 12px; box-shadow:0 6px 16px rgba(7,27,51,.05); }}
    .user-pill .left {{ display:flex; align-items:center; gap:10px; }}
    .user-pill .avatar {{ width:11px; height:11px; border-radius:50%; background:{BLUE}; box-shadow:0 0 0 5px rgba(18,97,160,.12); }}
    .user-pill b {{ color:{NAVY}; font-size:.82rem; display:block; }}
    .user-pill small {{ color:{GRAY}; font-size:.67rem; display:block; margin-top:2px; }}
    .user-pill span {{ color:{BLUE}; font-size:.69rem; font-weight:800; }}

    [data-testid="stSidebar"] [role="radiogroup"] label {{ border-radius:10px; padding:3px 4px; }}
    [data-testid="stSidebar"] hr {{ margin:.85rem 0; }}
    [data-testid="stTabs"] button {{ font-weight:800; color:#475467; }}
    [data-testid="stTabs"] button[aria-selected="true"] {{ color:{NAVY}; border-bottom-color:{BLUE}; }}

    @media(max-width:900px) {{
      .block-container {{ padding-left:.75rem; padding-right:.75rem; }}
      .dashboard-hero {{ padding:18px; }} .hero-grid {{ align-items:flex-start; flex-direction:column; }} .hero-chips {{ justify-content:flex-start; }}
      .kpi-card {{ min-height:125px; }} .insight-grid {{ grid-template-columns:1fr; }}
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
    return f"{float(value or 0) * 100:,.1f}%".replace(",", "X").replace(".", ",").replace("X", ".")


def safe_div(a: float, b: float) -> float:
    return float(a / b) if b not in (0, None) and not pd.isna(b) else 0.0


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


def plot_layout(fig: go.Figure, height: int = 390, legend_bottom: bool = True) -> go.Figure:
    legend = (
        dict(
            orientation="h", yanchor="top", y=-0.13, xanchor="left", x=0,
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
        margin=dict(l=20, r=70, t=92, b=70 if legend_bottom else 30),
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


def source_bytes(uploaded, default_path: Path | None) -> tuple[bytes | None, str]:
    if uploaded is not None:
        return uploaded.getvalue(), uploaded.name
    if default_path and default_path.exists():
        return default_path.read_bytes(), default_path.name
    return None, "Não localizada"


@st.cache_data(show_spinner=False)
def workbook_sheet_states(file_bytes: bytes) -> dict[str, str]:
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    return {ws.title: ws.sheet_state for ws in wb.worksheets}


@st.cache_data(show_spinner=False)
def read_excel_sheet(file_bytes: bytes, sheet_name: str, usecols=None) -> pd.DataFrame:
    return clean_columns(pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, usecols=usecols, engine="openpyxl"))


@st.cache_data(show_spinner=False)
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
    fields = ["FORNECEDOR", "SEGMENTO", "LINHA DE PRODUTO", "PRODUTO", "DESCRIÇÃO", "Endoscopia?", "NOVA"]
    text = " ".join(norm(row.get(c, "")) for c in fields)
    if "MICROTECH" in text or "MICRO TECH" in text:
        return "MICROTECH"
    if norm(row.get("Endoscopia?")) == "SIM" or "ENDOSCOPIA" in text or "GASTROENDOSCOPIA" in text:
        return "ENDOSCOPIA"
    if "LOCACAO" in norm(row.get("NOVA")) or "LOCACAO" in norm(row.get("SEGMENTO")):
        return "LOCACAO"
    return "VENDAS"


@st.cache_data(show_spinner=False)
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

    date_col = require_col(fat, ["DT Emissao", "MÊS"], "BANCO DE DADOS FATURAMENTO")
    month_col = require_col(fat, ["MÊS", "DT Emissao"], "BANCO DE DADOS FATURAMENTO")
    value_col = require_col(fat, ["VALOR BRUTO", "VALOR ", "VALOR"], "BANCO DE DADOS FATURAMENTO")
    fat["_DATA"] = pd.to_datetime(fat[date_col], errors="coerce")
    fat["_MES"] = pd.to_datetime(fat[month_col], errors="coerce").dt.to_period("M")
    fat["_VALOR"] = to_number(fat[value_col])
    for col in [
        "GERENTE", "VENDEDOR / REPRESENTANTE", "VENDEDOR", "SEGMENTO", "EMPRESA", "NOVA",
        "NOME DO CLIENTE", "CLIENTE", "FORNECEDOR", "LINHA DE PRODUTO", "PRODUTO", "Nota Fiscal", "CATEGORIA",
        "DESCRIÇÃO", "Endoscopia?",
    ]:
        if col in fat.columns:
            fat[col] = fat[col].fillna("Não informado").astype(str).str.strip()
    fallback_line = fat.apply(classify_billing_line, axis=1)
    if "GERENTE" in fat.columns:
        manager_line = fat["GERENTE"].map(norm).map(MANAGER_LINE_MAP)
        fat["_LINHA"] = manager_line.where(manager_line.notna(), fallback_line)
    else:
        fat["_LINHA"] = fallback_line
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
        df["_MES"] = pd.to_datetime(df[mcol], errors="coerce").dt.to_period("M")
        df["_META"] = to_number(df[vcol])
        df["_META_ANUAL"] = to_number(df[annual_col]) if annual_col else 0.0
        for c in ["GERENTE", "VENDEDOR"]:
            if c in df.columns:
                df[c] = df[c].fillna("Não informado").astype(str).str.strip()

    return {"faturamento": fat, "metas": metas, "metas_gerentes": metas_g, "usuarios": users, "sheet_states": states}


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


@st.cache_data(show_spinner=False)
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
    custos = read_excel_sheet(file_bytes, "Centro de Custos", usecols="A:P")
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

    custos["_MES"] = pd.to_datetime(custos[require_col(custos, ["MÊS"], "Centro de Custos")], errors="coerce").dt.to_period("M")
    custos["_VALOR"] = to_number(custos[require_col(custos, ["Valor"], "Centro de Custos")])
    for col in ["EMPRESA", "GRUPO", "SUBGRUPO", "PAI", "Categoria", "CENTRO DE CUSTOS", "CENTRO DE CUSTOS RATEAO", "Codigo-Nome do Fornecedor"]:
        if col in custos.columns:
            custos[col] = custos[col].fillna("Não informado").astype(str).str.strip()
    custos["_EMPRESA_N"] = custos["EMPRESA"].map(norm)
    custos["_GRUPO_N"] = custos["GRUPO"].map(norm)
    custos["_CC_N"] = custos["CENTRO DE CUSTOS"].map(norm)
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

    return {"receitas": receitas, "despesas": despesas, "custos": custos, "caixa": caixa, "recebimento": receb, "sheet_states": states}


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
    client_col = optional_col(df, ["Nome do Cliente", "Cliente", "Razão Social", "Razao Social", "Nome", "N Fantasia", "Codigo-Lj-Nome do Cliente"])
    value_col = optional_col(df, [
        "Tit Vencidos Valor Corrigido", "Vencidos Corrigidos", "Tit Vencidos Valor Atual",
        "Vencidos", "Saldo Vencido", "Valor Vencido", "Valor Original", "Saldo", "Valor"
    ])
    due_col = optional_col(df, ["Vencto Real", "Vencimento", "Vencto Titulo", "Vencto", "Data Vencimento", "Venc. Real"])
    days_col = optional_col(df, ["Dias Atraso", "Dias em Atraso", "Atraso"])
    line_col = optional_col(df, ["Linha", "Linha de Negócio", "Segmento", "Centro de Custos"])
    manager_col = optional_col(df, ["Gerente", "Gestor", "Responsável", "Responsavel"])
    title_col = optional_col(df, ["Prf-Numero Parcela", "Prf-Numero-Parcela", "Título", "Titulo", "Documento", "Nota Fiscal", "NF"])
    if client_col is None or value_col is None:
        raise ValueError("A base de inadimplência precisa conter pelo menos Cliente e Valor vencido.")

    x = df.copy()
    x["_CLIENTE"] = x[client_col].fillna("Não informado").astype(str).str.strip()
    x["_CLIENT_KEY"] = x["_CLIENTE"].map(client_key)
    x["_VALOR_VENCIDO"] = to_number(x[value_col])
    x["_VENCIMENTO"] = to_datetime_mixed(x[due_col]) if due_col else pd.NaT
    if days_col:
        x["_DIAS_ATRASO"] = to_number(x[days_col]).astype(int)
    elif due_col:
        x["_DIAS_ATRASO"] = (pd.Timestamp.today().normalize() - x["_VENCIMENTO"]).dt.days.fillna(0).clip(lower=0).astype(int)
    else:
        x["_DIAS_ATRASO"] = 0
    x["_MES"] = x["_VENCIMENTO"].dt.to_period("M") if due_col else pd.Period(pd.Timestamp.today(), freq="M")
    x["_TITULO"] = x[title_col].fillna("").astype(str) if title_col else ""

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
    elif line_col:
        raw_line = x[line_col].map(norm)
        x["_LINHA"] = np.select(
            [raw_line.str.contains("MICROTECH"), raw_line.str.contains("ENDOSCOPIA"), raw_line.str.contains("LOCACAO"), raw_line.str.contains("VENDA")],
            ["MICROTECH", "ENDOSCOPIA", "LOCACAO", "VENDAS"], default="NAO CLASSIFICADA"
        )
        x["_GERENTE"] = x["_LINHA"].map(LINE_MANAGER_MAP).fillna("Sem gerente")
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
    meta: dict[str, object] = {
        "cliente": client_col, "valor": value_col, "vencimento": due_col or "Não disponível",
        "linha": line_col or "Mapeada pelo gerente dominante da BASE BI",
        "gerente": manager_col or "Mapeado pelo cliente na BASE BI",
        "sheet": source_meta.get("sheet", ""),
        "resumo_gerentes": source_meta.get("resumo_gerentes", []),
    }
    return x, meta


# =========================================================
# AUTENTICAÇÃO E PERFIS
# =========================================================
def password_hash(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


DEFAULT_USERS = {
    "paula": {
        "nome": "Diretoria", "usuario": "paula", "email": "paulamayara10@gmail.com",
        "senha_hash": "cbae32987728a10b19e2528695dbf676c9b8de0f539bab08b5789c2e9f0d8599",
        "perfil": "DIRETORIA", "linha": "CONSOLIDADO",
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
        profile_order = ["paula", "celso", "renato", "amauri", "ronaldo"]
        available = [key for key in profile_order if key in users]
        profile_labels = {
            "paula": "Diretoria",
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
                  <p>Painel gerencial de caixa por perfil</p>
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
    out["Qualidade das Entradas"] = np.where(
        out["Receitas Operacionais"] + out["Entradas Não Operacionais"] != 0,
        out["Receitas Operacionais"] / (out["Receitas Operacionais"] + out["Entradas Não Operacionais"]), 0
    )
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
        fbase = fbase[fbase["_LINHA"] == line]
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
    fbase = period_filter(fat, start, end).copy()
    mbase = period_filter(metas, start, end).copy()
    if line != "CONSOLIDADO":
        fbase = fbase[fbase["_LINHA"] == line]
        manager = LINE_MANAGER_MAP.get(line, "")
        if "GERENTE" in mbase.columns:
            mbase = mbase[mbase["GERENTE"].map(norm) == manager]

    seller_col = "VENDEDOR / REPRESENTANTE" if "VENDEDOR / REPRESENTANTE" in fbase.columns else "VENDEDOR"
    fbase["_SELLER_KEY"] = fbase[seller_col].map(norm)
    mbase["_SELLER_KEY"] = mbase["VENDEDOR"].map(norm) if "VENDEDOR" in mbase.columns else "NAO INFORMADO"

    actual = fbase.groupby("_SELLER_KEY", as_index=False)["_VALOR"].sum().rename(columns={"_VALOR": "Faturamento"})
    actual_names = fbase.groupby("_SELLER_KEY", as_index=False)[seller_col].first().rename(columns={seller_col: "Vendedor"})
    goals = mbase.groupby("_SELLER_KEY", as_index=False)["_META"].sum().rename(columns={"_META": "Meta"})
    goal_names = mbase.groupby("_SELLER_KEY", as_index=False)["VENDEDOR"].first().rename(columns={"VENDEDOR": "Vendedor Meta"}) if "VENDEDOR" in mbase.columns else pd.DataFrame(columns=["_SELLER_KEY", "Vendedor Meta"])

    out = actual.merge(goals, on="_SELLER_KEY", how="outer").merge(actual_names, on="_SELLER_KEY", how="left")
    if not goal_names.empty:
        out = out.merge(goal_names, on="_SELLER_KEY", how="left")
        out["Vendedor"] = out["Vendedor"].where(out["Vendedor"].notna(), out["Vendedor Meta"])
        out = out.drop(columns=["Vendedor Meta"])
    out[["Faturamento", "Meta"]] = out[["Faturamento", "Meta"]].fillna(0)
    out["Vendedor"] = out["Vendedor"].fillna(out["_SELLER_KEY"].str.title())
    # Exibe somente quem teve participação efetiva na linha no período.
    # Vendedores presentes apenas na meta, mas sem faturamento, não aparecem no ranking.
    out = out[out["Faturamento"] > 0].copy()
    out["Desvio"] = out["Faturamento"] - out["Meta"]
    out["Atingimento"] = np.where(out["Meta"] != 0, out["Faturamento"] / out["Meta"], 0)
    out["Status"] = np.select(
        [out["Atingimento"] >= 1, out["Atingimento"] >= .9],
        ["Meta atingida", "Próximo da meta"],
        default="Abaixo da meta",
    )
    return out.sort_values("Faturamento", ascending=False).reset_index(drop=True)


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
        period_filter(fat[fat["_LINHA"] == line], start, end)
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
        sums["Qualidade das Entradas"] = safe_div(sums.get("Receitas Operacionais", 0), sums.get("Receitas Operacionais", 0) + sums.get("Entradas Não Operacionais", 0))
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
        f = period_filter(fat[fat["_LINHA"] == scope], start, end)
        r = period_filter(receitas[(receitas["_LINHA"] == scope) & (~receitas["_NAO_OPERACIONAL"])], start, end)
        c = period_filter(custos[(custos["_LINHA_DIRETA"] == scope) & (custos["_GRUPO_N"] == "SAIDAS OPERACIONAIS")], start, end)
        i = inad[inad["_LINHA"] == scope].copy() if inad is not None else None
    return f, r, c, i


# =========================================================
# FONTES, USUÁRIO E FILTROS
# =========================================================
user = authenticate()
is_director = user["perfil"] in {"DIRETORIA", "ADMIN", "CONTROLADORIA"}

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
        f"<div class='user-pill'><div class='left'><span class='avatar'></span><div><b>{user['nome']}</b><small>{'Acesso executivo' if is_director else 'Acesso por linha'}</small></div></div><span>{'Diretoria' if is_director else line_label(user['linha'])}</span></div>",
        unsafe_allow_html=True,
    )
    if secret_users() and st.button("Sair", width="stretch"):
        st.session_state.pop("auth_user", None)
        st.rerun()

    default_base = first_existing(["BASE BI.xlsx", "BASE BI(1).xlsx", "base_bi.xlsx"])
    default_rev = first_existing(["rev2026 Base bi.xlsx", "rev2026 Base bi(1).xlsx", "REV2026.xlsx"])
    default_inad = first_existing([
        "relatorio_cobranca_gerente.xlsx", "relatorio_cobranca_gerente_2026-07-13.xlsx",
        "inadimplencia.xlsx", "Inadimplencia.xlsx", "base_inadimplencia.xlsx", "inadimplencia.csv"
    ])

    up_base = up_rev = up_inad = None
    if is_director:
        with st.expander("Fontes de dados", expanded=False):
            up_base = st.file_uploader("Substituir BASE BI", type=["xlsx", "xlsm"], key="up_base_final")
            up_rev = st.file_uploader("Substituir REV2026", type=["xlsx", "xlsm"], key="up_rev_final")
            up_inad = st.file_uploader("Base de inadimplência", type=["xlsx", "xlsm", "csv"], key="up_inad_final")

base_bytes, base_name = source_bytes(up_base, default_base)
rev_bytes, rev_name = source_bytes(up_rev, default_rev)
inad_bytes, inad_name = source_bytes(up_inad, default_inad)

if not base_bytes or not rev_bytes:
    st.error("Inclua `BASE BI.xlsx` e `rev2026 Base bi.xlsx` no repositório ou carregue os arquivos na barra lateral.")
    st.stop()

try:
    with st.spinner("Organizando visão de caixa e linhas de negócio..."):
        base = load_base_bi(base_bytes)
        rev = load_rev(rev_bytes)
        fat = base["faturamento"]
        metas = base["metas"]
        metas_g = base["metas_gerentes"]
        receitas, receipt_match_stats = assign_receipt_lines(rev["receitas"], fat)
        despesas = rev["despesas"]
        custos = rev["custos"]
        performance = rev["recebimento"]
        inad = None
        inad_meta = None
        if inad_bytes:
            inad, inad_meta = prepare_inadimplencia(inad_bytes, inad_name, fat)
except Exception as exc:
    st.error(f"Não foi possível carregar as bases: {exc}")
    st.stop()

all_periods = pd.concat([fat["_MES"], receitas["_MES"], despesas["_MES"], custos["_MES"], performance["_MES"]]).dropna()
min_month, max_month = all_periods.min(), all_periods.max()
month_options = list(pd.period_range(min_month, max_month, freq="M"))

with st.sidebar:
    st.markdown("#### Período")
    start_month = st.selectbox("Mês inicial", month_options, index=0, format_func=month_label)
    end_month = st.selectbox("Mês final", month_options, index=len(month_options) - 1, format_func=month_label)
    if start_month > end_month:
        start_month = end_month

    if is_director:
        scope_choice = st.selectbox("Escopo do painel", ["CONSOLIDADO"] + LINES, format_func=line_label)
    else:
        scope_choice = user["linha"] if user["linha"] in LINES else "VENDAS"

    st.markdown("#### Navegação")
    pages = ["Dashboard", "Desempenho & metas", "Recebimentos & inadimplência", "Clientes", "Produtos", "Custos diretos"]
    if is_director:
        pages.insert(1, "Linhas de negócio")
    page = st.radio("Navegação", pages, label_visibility="collapsed")


period_text = f"{month_label(start_month)} a {month_label(end_month)}"
scope_text = line_label(scope_choice)

st.markdown(
    f"""
    <div class='dashboard-hero'>
      <div class='hero-grid'>
        <div>
          <div class='hero-brand'>First Medical · Controladoria</div>
          <h1>{'Painel Executivo de Caixa' if scope_choice == 'CONSOLIDADO' else 'Resultado de Caixa · ' + scope_text}</h1>
        </div>
        <div class='hero-chips'>
          <span class='hero-chip'>{period_text}</span>
          <span class='hero-chip'>{scope_text}</span>
          <span class='hero-chip'>{'Acesso diretoria' if is_director else 'Acesso restrito'}</span>
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
# DASHBOARD
# =========================================================
if page == "Dashboard":
    if scope_choice == "CONSOLIDADO":
        k1, k2, k3, k4 = st.columns(4)
        with k1: card("Receitas operacionais recebidas", brl(company_totals["Receitas Operacionais"]), "Entradas efetivamente recebidas da operação", BLUE)
        with k2: card("Saídas operacionais pagas", brl(company_totals["Saídas Operacionais"]), "Pagamentos operacionais efetivados", RED)
        with k3: card("Resultado operacional de caixa", brl(company_totals["Resultado Operacional de Caixa"]), "Receitas recebidas menos saídas pagas", BLUE if company_totals["Resultado Operacional de Caixa"] >= 0 else CYAN)
        with k4: card("Margem operacional de caixa", pct(company_totals["Margem de Caixa"]), "Resultado operacional ÷ receitas recebidas", TEAL)

        k5, k6, k7, k8 = st.columns(4)
        with k5: card("EBITDA gerencial de caixa", brl(company_totals["EBITDA Gerencial de Caixa"]), "Resultado antes de IRPJ e CSLL pagos", NAVY)
        with k6: card("Margem de contribuição de caixa", pct(company_totals["Margem de Contribuição Caixa"]), "Receitas recebidas menos saídas variáveis pagas", CYAN)
        with k7: card("Entradas não operacionais", brl(company_totals["Entradas Não Operacionais"]), "Capital de giro e outras fontes não operacionais", CYAN)
        with k8: card("Qualidade das entradas", pct(company_totals["Qualidade das Entradas"]), "Participação operacional nas entradas classificadas", TEAL)

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
            st.plotly_chart(plot_layout(fig, 430), width="stretch", config={"displayModeBar": False})
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
            st.plotly_chart(plot_layout(bridge, 430, False), width="stretch", config={"displayModeBar": False})

        section_header("Resultado por linha de negócio", "Comparativo direto", "Sem rateio")
        dash_selected_codes = st.multiselect(
            "Linhas exibidas", LINES, default=LINES, format_func=line_label, key="filter_dashboard_lines"
        )
        dashboard_lines = lines_table[lines_table["Código"].isin(dash_selected_codes)].copy() if dash_selected_codes else lines_table.copy()
        l1, l2 = st.columns([1.25, 1])
        with l1:
            chart_df = dashboard_lines.sort_values("Resultado Direto de Caixa")
            fig = go.Figure(go.Bar(
                x=chart_df["Resultado Direto de Caixa"],
                y=chart_df["Linha"],
                orientation="h",
                marker_color=[LINE_COLORS.get(code, BLUE) if value >= 0 else CYAN for code, value in zip(chart_df["Código"], chart_df["Resultado Direto de Caixa"])],
                text=chart_df["Resultado Direto de Caixa"].map(compact_money),
                textposition=["outside" if value >= 0 else "inside" for value in chart_df["Resultado Direto de Caixa"]],
                textfont_color=[DARK if value >= 0 else WHITE for value in chart_df["Resultado Direto de Caixa"]],
                insidetextanchor="middle",
                cliponaxis=False,
                customdata=chart_df[["Linha"]],
                hovertemplate="%{customdata[0]}<br>Resultado: R$ %{x:,.2f}<extra></extra>",
            ))
            fig.update_layout(title="Resultado direto de caixa por linha", showlegend=False)
            hide_value_axis(fig, "x")
            st.plotly_chart(plot_layout(fig, 410, False), width="stretch", config={"displayModeBar": False})
        with l2:
            fig = go.Figure()
            fig.add_bar(x=dashboard_lines["Linha"], y=dashboard_lines["Receitas Recebidas"], name="Receitas diretas", marker_color=BLUE,
                        text=dashboard_lines["Receitas Recebidas"].map(compact_money), textposition="outside", cliponaxis=False, textfont=dict(size=10, color=NAVY))
            fig.add_bar(x=dashboard_lines["Linha"], y=dashboard_lines["Custos Diretos Pagos"], name="Custos diretos", marker_color=RED,
                        text=dashboard_lines["Custos Diretos Pagos"].map(compact_money), textposition="inside", insidetextanchor="end", cliponaxis=False, textfont=dict(size=10, color=WHITE))
            fig.update_layout(title="Receitas diretas x custos diretos", barmode="group")
            hide_value_axis(fig, "y")
            st.plotly_chart(plot_layout(fig, 410), width="stretch", config={"displayModeBar": False})

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
        st.plotly_chart(plot_layout(perf_fig, 410), width="stretch", config={"displayModeBar": False})

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
            st.plotly_chart(plot_layout(fig, 440, False), width="stretch", config={"displayModeBar": False})
        with c2:
            top_cost = cost_scope.groupby("PAI", as_index=False)["_VALOR"].sum().sort_values("_VALOR", ascending=False).head(10)
            top_cost["Natureza"] = top_cost["PAI"].map(lambda x: short_label(x, 36))
            p = top_cost.sort_values("_VALOR")
            fig = px.bar(p, x="_VALOR", y="Natureza", orientation="h", title="Principais custos diretos pagos", custom_data=["PAI"])
            fig.update_traces(marker_color=BLUE, text=p["_VALOR"].map(compact_money), textposition="outside", cliponaxis=False,
                              hovertemplate="%{customdata[0]}<br>Valor: R$ %{x:,.2f}<extra></extra>")
            hide_value_axis(fig, "x")
            st.plotly_chart(plot_layout(fig, 440, False), width="stretch", config={"displayModeBar": False})


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
        st.plotly_chart(plot_layout(fig, 430), width="stretch", config={"displayModeBar": False})
    with c2:
        att = perf_monthly.copy()
        colors = [BLUE if value >= 1 else CYAN for value in att["Atingimento"]]
        fig = go.Figure(go.Bar(
            x=att["Mês Texto"], y=att["Atingimento"], marker_color=colors,
            text=att["Atingimento"].map(pct), textposition="outside", cliponaxis=False,
            hovertemplate="%{x}<br>Atingimento: %{y:.1%}<extra></extra>",
        ))
        fig.add_hline(y=1, line_dash="dot", line_color=NAVY, line_width=2)
        fig.update_layout(title="Atingimento mensal")
        hide_value_axis(fig, "y")
        st.plotly_chart(plot_layout(fig, 430, False), width="stretch", config={"displayModeBar": False})

    if is_director and scope_choice == "CONSOLIDADO":
        section_header("Desempenho por linha", "Comparação comercial e financeira", "Diretoria")
        line_perf = lines_table[lines_table["Código"].isin(selected_lines)].copy() if selected_lines else lines_table.copy()
        score = line_perf[[
            "Linha", "Faturamento", "Meta", "Atingimento da Meta", "Desvio da Meta",
            "Receitas Recebidas", "Conversão em Caixa", "Resultado Direto de Caixa",
            "Margem Direta de Caixa", "Inadimplência"
        ]].copy()
        score["Inadimplência / faturamento"] = np.where(score["Faturamento"] != 0, score["Inadimplência"] / score["Faturamento"], 0)
        score["Status"] = np.select(
            [score["Atingimento da Meta"] >= 1, score["Atingimento da Meta"] >= .9],
            ["Meta atingida", "Próximo da meta"], default="Abaixo da meta"
        )

        chart = line_perf.sort_values("Atingimento da Meta")
        fig = go.Figure(go.Bar(
            x=chart["Atingimento da Meta"], y=chart["Linha"], orientation="h",
            marker_color=[BLUE if value >= 1 else CYAN for value in chart["Atingimento da Meta"]],
            text=chart["Atingimento da Meta"].map(pct), textposition="outside", cliponaxis=False,
        ))
        fig.add_vline(x=1, line_dash="dot", line_color=NAVY, line_width=2)
        fig.update_layout(title="Atingimento da meta por linha", showlegend=False)
        hide_value_axis(fig, "x")
        st.plotly_chart(plot_layout(fig, 390, False), width="stretch", config={"displayModeBar": False})

        score_view = score.copy()
        for col in ["Faturamento", "Meta", "Desvio da Meta", "Receitas Recebidas", "Resultado Direto de Caixa", "Inadimplência"]:
            score_view[col] = score_view[col].map(brl)
        for col in ["Atingimento da Meta", "Conversão em Caixa", "Margem Direta de Caixa", "Inadimplência / faturamento"]:
            score_view[col] = score_view[col].map(pct)
        st.dataframe(score_view, width="stretch", hide_index=True, height=255)

    section_header("Desempenho da equipe", "Faturamento individual frente à meta")
    seller_perf = seller_performance(fat, metas, start_month, end_month, scope_choice)
    sf1, sf2, sf3 = st.columns([1.4, 1, .7])
    seller_search = sf1.text_input("Buscar vendedor ou representante", key="seller_performance_search")
    seller_status = sf2.multiselect(
        "Status", ["Meta atingida", "Próximo da meta", "Abaixo da meta"], key="seller_performance_status"
    )
    seller_top = sf3.selectbox("Exibir", [10, 15, 20, 30], index=1, key="seller_performance_top")
    if seller_search:
        seller_perf = seller_perf[seller_perf["Vendedor"].astype(str).str.contains(seller_search, case=False, na=False)]
    if seller_status:
        seller_perf = seller_perf[seller_perf["Status"].isin(seller_status)]
    seller_perf = seller_perf.head(seller_top)

    if seller_perf.empty:
        st.info("Nenhum vendedor encontrado para os filtros selecionados.")
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
        st.plotly_chart(plot_layout(fig, max(390, 42 * len(seller_plot)), False), width="stretch", config={"displayModeBar": False})

        seller_view = seller_perf[["Vendedor", "Faturamento", "Meta", "Atingimento", "Desvio", "Status"]].copy()
        seller_export = seller_view.copy()
        seller_view["Faturamento"] = seller_view["Faturamento"].map(brl)
        seller_view["Meta"] = seller_view["Meta"].map(brl)
        seller_view["Atingimento"] = seller_view["Atingimento"].map(pct)
        seller_view["Desvio"] = seller_view["Desvio"].map(brl)
        st.dataframe(seller_view, width="stretch", hide_index=True, height=360)
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

    c1, c2 = st.columns(2)
    with c1:
        p = lines_view.sort_values("Receitas Recebidas")
        fig = go.Figure(go.Bar(
            x=p["Receitas Recebidas"], y=p["Linha"], orientation="h",
            marker_color=[LINE_COLORS.get(code, BLUE) for code in p["Código"]],
            text=p["Receitas Recebidas"].map(compact_money), textposition="outside", cliponaxis=False,
            hovertemplate="%{y}<br>Receita direta: R$ %{x:,.2f}<extra></extra>",
        ))
        fig.update_layout(title="Receitas diretas por linha", showlegend=False)
        hide_value_axis(fig, "x")
        st.plotly_chart(plot_layout(fig, 440, False), width="stretch", config={"displayModeBar": False})
    with c2:
        p = lines_view.sort_values("Resultado Direto de Caixa")
        fig = go.Figure(go.Bar(
            x=p["Resultado Direto de Caixa"], y=p["Linha"], orientation="h",
            marker_color=[LINE_COLORS.get(code, BLUE) if value >= 0 else CYAN for code, value in zip(p["Código"], p["Resultado Direto de Caixa"])],
            text=p["Resultado Direto de Caixa"].map(compact_money),
            textposition=["outside" if value >= 0 else "inside" for value in p["Resultado Direto de Caixa"]],
            textfont_color=[NAVY if value >= 0 else WHITE for value in p["Resultado Direto de Caixa"]],
            insidetextanchor="middle", cliponaxis=False,
            hovertemplate="%{y}<br>Resultado: R$ %{x:,.2f}<extra></extra>",
        ))
        fig.update_layout(title="Resultado direto por linha", showlegend=False)
        hide_value_axis(fig, "x")
        st.plotly_chart(plot_layout(fig, 440, False), width="stretch", config={"displayModeBar": False})

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
    st.dataframe(view, width="stretch", hide_index=True, height=280)
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
        st.plotly_chart(plot_layout(fig, 420), width="stretch", config={"displayModeBar": False})
    else:
        r1, r2, r3, r4 = st.columns(4)
        with r1: card("Receita operacional direta", brl(line_totals["Receitas Recebidas"]), "Centro de Custos · RECEITA", BLUE)
        with r2: card("Faturamento emitido", brl(line_totals["Faturamento"]), "Base comercial do período", NAVY)
        with r3: card("Conversão em caixa", pct(line_totals["Conversão em Caixa"]), "Receita direta ÷ faturamento", TEAL)
        with r4: card("Resultado direto", brl(line_totals["Resultado Direto de Caixa"]), "Receita direta menos custo direto", BLUE if line_totals["Resultado Direto de Caixa"] >= 0 else CYAN)

    section_header("Inadimplência", "Relatório do CRM de Cobrança")
    inad_page = None if inad_scope is None else inad_scope.copy()
    if inad_page is not None and not inad_page.empty:
        fi1, fi2, fi3, fi4 = st.columns([1.3, 1.15, 1, .65])
        inad_client_search = fi1.text_input("Buscar cliente", placeholder="Nome do cliente", key="filter_inad_client")
        faixa_order = ["Até 30 dias", "31 a 60 dias", "61 a 90 dias", "Acima de 90 dias"]
        faixa_options = [x for x in faixa_order if x in inad_page["_FAIXA"].astype(str).unique()]
        selected_faixas = fi2.multiselect("Faixa de atraso", faixa_options, key="filter_inad_faixa")
        max_days = max(int(pd.to_numeric(inad_page["_DIAS_ATRASO"], errors="coerce").fillna(0).max()), 1)
        min_days = fi3.number_input("Atraso mínimo", min_value=0, max_value=max_days, value=0, step=5, key="filter_inad_days")
        inad_top_n = fi4.selectbox("Exibir", [10, 12, 15, 20], index=1, key="filter_inad_top")
        if inad_client_search:
            inad_page = inad_page[inad_page["_CLIENTE"].astype(str).str.contains(inad_client_search, case=False, na=False)]
        if selected_faixas:
            inad_page = inad_page[inad_page["_FAIXA"].astype(str).isin(selected_faixas)]
        inad_page = inad_page[pd.to_numeric(inad_page["_DIAS_ATRASO"], errors="coerce").fillna(0) >= min_days]

    if inad_page is None:
        st.info("A base de inadimplência ainda não foi carregada. Campos aceitos: Cliente, Valor vencido, Vencimento e Dias de atraso. Também são reconhecidos `Vencidos Corrigidos`, `Vencidos` e `Valor Original`.")
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
            st.plotly_chart(plot_layout(fig, 400, False), width="stretch", config={"displayModeBar": False})
        with c2:
            cli = inad_page.groupby("_CLIENTE", as_index=False)["_VALOR_VENCIDO"].sum().sort_values("_VALOR_VENCIDO", ascending=False).head(inad_top_n)
            cli["Nome"] = cli["_CLIENTE"].map(lambda x: short_label(x, 38))
            p = cli.sort_values("_VALOR_VENCIDO")
            fig = px.bar(p, x="_VALOR_VENCIDO", y="Nome", orientation="h", title="Maiores saldos vencidos", custom_data=["_CLIENTE"])
            fig.update_traces(marker_color=NAVY, text=p["_VALOR_VENCIDO"].map(compact_money), textposition="outside", cliponaxis=False,
                              hovertemplate="%{customdata[0]}<br>Vencido: R$ %{x:,.2f}<extra></extra>")
            hide_value_axis(fig, "x")
            st.plotly_chart(plot_layout(fig, 400, False), width="stretch", config={"displayModeBar": False})

        if is_director and inad_meta and inad_meta.get("resumo_gerentes"):
            crm_summary = pd.DataFrame(inad_meta["resumo_gerentes"])
            crm_summary = crm_summary[crm_summary["Gerente"].map(norm).isin(set(MANAGER_LINE_MAP) | {"GABRIEL", "SEM GERENTE"})]
            crm_summary["Linha"] = crm_summary["Linha"].map(line_label)
            crm_view = crm_summary[["Gerente", "Linha", "Clientes", "Títulos", "Valor", "Maior atraso"]].copy()
            crm_view["Valor"] = crm_view["Valor"].map(brl)
            st.markdown("**Resumo original exportado pelo CRM de cobrança**")
            st.dataframe(crm_view, width="stretch", hide_index=True)

        detail = inad_page[["_CLIENTE", "_TITULO", "_VENCIMENTO", "_DIAS_ATRASO", "_FAIXA", "_VALOR_VENCIDO", "_LINHA", "_GERENTE"]].copy()
        detail.columns = ["Cliente", "Título", "Vencimento", "Dias de atraso", "Faixa", "Valor vencido", "Linha", "Gerente"]
        detail["Linha"] = detail["Linha"].map(line_label)
        detail_view = detail.sort_values("Valor vencido", ascending=False).copy()
        detail_view["Valor vencido"] = detail_view["Valor vencido"].map(brl)
        detail_view["Vencimento"] = pd.to_datetime(detail_view["Vencimento"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
        st.dataframe(detail_view, width="stretch", hide_index=True, height=370)
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
            st.plotly_chart(plot_layout(fig, 540, False), width="stretch", config={"displayModeBar": False})
        with p2:
            top = received.sort_values("Recebido", ascending=False).head(top_n).copy()
            top["Nome"] = top["Cliente Recebimento"].map(lambda x: short_label(x, 38)) if not top.empty else []
            p = top.sort_values("Recebido")
            fig = px.bar(p, x="Recebido", y="Nome", orientation="h", title="Clientes por recebimento", custom_data=["Cliente Recebimento"])
            fig.update_traces(marker_color=TEAL, text=p["Recebido"].map(compact_money), textposition="outside", cliponaxis=False,
                              hovertemplate="%{customdata[0]}<br>Recebido: R$ %{x:,.2f}<extra></extra>")
            hide_value_axis(fig, "x")
            st.plotly_chart(plot_layout(fig, 540, False), width="stretch", config={"displayModeBar": False})

        view = billed.head(100).rename(columns={client_col: "Cliente"}).copy()
        view["Faturamento"] = view["Faturamento"].map(brl)
        view["Participação"] = view["Participação"].map(pct)
        st.dataframe(view, width="stretch", hide_index=True, height=430)

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
                    st.plotly_chart(plot_layout(fig, 520, False), width="stretch", config={"displayModeBar": False})
                with c2:
                    qty_top = products.sort_values("Quantidade", ascending=False).head(top_n).sort_values("Quantidade")
                    qty_top["Produto"] = qty_top["_PRODUTO"].map(lambda x: short_label(x, 34))
                    fig = px.bar(qty_top, x="Quantidade", y="Produto", orientation="h", title="Produtos por volume")
                    fig.update_traces(marker_color=TEAL, text=qty_top["Quantidade"].map(lambda v: f"{v:,.0f}".replace(",", ".")), textposition="outside", cliponaxis=False)
                    hide_value_axis(fig, "x")
                    st.plotly_chart(plot_layout(fig, 520, False), width="stretch", config={"displayModeBar": False})

                table = products.rename(columns={"_PRODUTO": "Produto"})[["Produto", "Faturamento", "Participação", "Quantidade", "Preço médio", "Clientes", "Notas"]].copy()
                table_view = table.copy()
                table_view["Faturamento"] = table_view["Faturamento"].map(brl)
                table_view["Participação"] = table_view["Participação"].map(pct)
                table_view["Quantidade"] = table_view["Quantidade"].map(lambda v: f"{v:,.0f}".replace(",", "."))
                table_view["Preço médio"] = table_view["Preço médio"].map(brl)
                st.dataframe(table_view, width="stretch", hide_index=True, height=470)
                st.download_button(
                    "Exportar produtos",
                    dataframe_download(table, "Produtos"),
                    file_name=f"produtos_{scope_choice}_{start_month}_{end_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

# =========================================================
# PREÇOS E RENTABILIDADE ESTIMADA
# =========================================================
elif page == "Custos diretos":
    section_header("Custos diretos", "Fornecedores e naturezas", scope_text)
    if cost_scope.empty:
        st.info("Não foram encontrados custos diretos para o escopo selecionado.")
    else:
        supplier_col = optional_col(cost_scope, ["Codigo-Nome do Fornecedor", "FORNECEDOR", "NOME DO FORNECEDOR"]) or "Codigo-Nome do Fornecedor"
        cost_base = cost_scope.copy()
        f1, f2, f3 = st.columns([1.3, 1.2, .6])
        supplier_search = f1.text_input("Buscar fornecedor", placeholder="Nome ou código", key="filter_cost_supplier")
        nature_values = sorted(cost_base["PAI"].dropna().astype(str).unique())
        selected_natures = f2.multiselect("Natureza", nature_values, key="filter_cost_nature")
        top_n = f3.selectbox("Exibir", [10, 15, 20, 30], index=1, key="filter_cost_top")
        if supplier_search:
            cost_base = cost_base[cost_base[supplier_col].astype(str).str.contains(supplier_search, case=False, na=False)]
        if selected_natures:
            cost_base = cost_base[cost_base["PAI"].astype(str).isin(selected_natures)]

        if cost_base.empty:
            st.info("Nenhum custo encontrado com os filtros selecionados.")
        else:
            total_cost = float(cost_base["_VALOR"].sum())
            suppliers = cost_base.groupby(supplier_col, as_index=False).agg(Valor=("_VALOR", "sum"), Lançamentos=("_VALOR", "size")).sort_values("Valor", ascending=False)
            nature = cost_base.groupby("PAI", as_index=False)["_VALOR"].sum().sort_values("_VALOR", ascending=False)
            months = cost_base.groupby("_MES", as_index=False)["_VALOR"].sum()
            months["Mês"] = months["_MES"].map(month_label)

            c1, c2, c3, c4 = st.columns(4)
            with c1: card("Custos pagos", brl(total_cost), "Filtro atual", RED)
            with c2: card("Fornecedores", f"{len(suppliers):,}".replace(",", "."), "Com pagamento direto", BLUE)
            with c3: card("Maior fornecedor", brl(float(suppliers.iloc[0]["Valor"])), short_label(suppliers.iloc[0][supplier_col], 34), NAVY)
            with c4: card("Ticket médio", brl(safe_div(total_cost, float(suppliers["Lançamentos"].sum()))), "Por lançamento", TEAL)

            p1, p2 = st.columns(2)
            with p1:
                p = nature.head(top_n).sort_values("_VALOR").copy()
                p["Natureza"] = p["PAI"].map(lambda x: short_label(x, 36))
                fig = px.bar(p, x="_VALOR", y="Natureza", orientation="h", title="Naturezas de custo", custom_data=["PAI"])
                fig.update_traces(marker_color=RED, text=p["_VALOR"].map(compact_money), textposition="outside", cliponaxis=False,
                                  hovertemplate="%{customdata[0]}<br>Valor: R$ %{x:,.2f}<extra></extra>")
                hide_value_axis(fig, "x")
                st.plotly_chart(plot_layout(fig, 520, False), width="stretch", config={"displayModeBar": False})
            with p2:
                p = suppliers.head(top_n).sort_values("Valor").copy()
                p["Fornecedor"] = p[supplier_col].map(lambda x: short_label(x, 38))
                fig = px.bar(p, x="Valor", y="Fornecedor", orientation="h", title="Fornecedores por pagamento", custom_data=[supplier_col])
                fig.update_traces(marker_color=RED, text=p["Valor"].map(compact_money), textposition="outside", cliponaxis=False,
                                  hovertemplate="%{customdata[0]}<br>Valor: R$ %{x:,.2f}<extra></extra>")
                hide_value_axis(fig, "x")
                st.plotly_chart(plot_layout(fig, 520, False), width="stretch", config={"displayModeBar": False})

            fig = px.bar(months, x="Mês", y="_VALOR", title="Custos diretos por mês")
            fig.update_traces(marker_color=RED, text=months["_VALOR"].map(compact_money), textposition="outside", cliponaxis=False)
            hide_value_axis(fig, "y")
            st.plotly_chart(plot_layout(fig, 380, False), width="stretch", config={"displayModeBar": False})

            table = suppliers.rename(columns={supplier_col: "Fornecedor"}).copy()
            table["Participação"] = np.where(total_cost != 0, table["Valor"] / total_cost, 0)
            table_view = table.copy()
            table_view["Valor"] = table_view["Valor"].map(brl)
            table_view["Participação"] = table_view["Participação"].map(pct)
            st.dataframe(table_view, width="stretch", hide_index=True, height=420)
            st.download_button(
                "Exportar custos diretos",
                dataframe_download(table, "Custos Diretos"),
                file_name=f"custos_diretos_{scope_choice}_{start_month}_{end_month}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
