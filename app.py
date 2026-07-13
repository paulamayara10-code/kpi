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
CYAN = "#00A7B5"
TEAL = "#008C95"
GREEN = "#1B7F5A"
RED = "#B42318"
ORANGE = "#C06A12"
YELLOW = "#F2B824"
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
    "MICROTECH": "#00A7B5",
    "LOCACAO": "#1261A0",
    "VENDAS": "#1B7F5A",
    "ENDOSCOPIA": "#6E55A3",
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
    [data-testid="stSidebar"] {{ background:#FFFFFF; border-right:1px solid {BORDER}; }}
    [data-testid="stSidebar"] .block-container {{ padding-top:1rem; }}
    [data-testid="stPlotlyChart"] {{ background:#FFFFFF; border:1px solid {BORDER}; border-radius:18px; padding:4px 7px; box-shadow:0 5px 18px rgba(7,27,51,.045); overflow:hidden; }}
    div[data-testid="stDataFrame"] {{ border:1px solid {BORDER}; border-radius:16px; overflow:hidden; background:white; }}
    div[data-testid="stDownloadButton"] button, div[data-testid="stButton"] button {{ border-radius:11px; font-weight:750; }}

    .first-sidebar {{
        background:linear-gradient(140deg,{NAVY},{NAVY_2}); color:white; border-radius:18px;
        padding:17px 18px; margin:0 0 12px 0; box-shadow:0 10px 24px rgba(7,27,51,.16);
    }}
    .first-sidebar .brand {{ font-size:1.30rem; font-weight:950; letter-spacing:.16em; line-height:1; }}
    .first-sidebar .brand small {{ display:block; color:#75E3E8; font-size:.58rem; letter-spacing:.30em; margin-top:6px; }}
    .first-sidebar p {{ margin:12px 0 0 0; opacity:.78; font-size:.74rem; line-height:1.4; }}

    .dashboard-hero {{
        position:relative; overflow:hidden; background:linear-gradient(118deg,{NAVY} 0%,#103E6A 66%,{TEAL} 132%);
        color:white; border-radius:22px; padding:22px 27px; margin-bottom:14px;
        box-shadow:0 13px 30px rgba(7,27,51,.16);
    }}
    .dashboard-hero::after {{ content:""; position:absolute; width:320px; height:320px; border-radius:50%; right:-130px; top:-185px; background:rgba(255,255,255,.055); }}
    .hero-grid {{ position:relative; z-index:2; display:flex; align-items:center; justify-content:space-between; gap:20px; }}
    .hero-brand {{ font-size:.72rem; font-weight:900; letter-spacing:.28em; color:#76E4E9; text-transform:uppercase; }}
    .dashboard-hero h1 {{ font-size:clamp(1.42rem,2.2vw,2rem); line-height:1.15; margin:7px 0 6px 0; letter-spacing:-.025em; }}
    .dashboard-hero p {{ margin:0; opacity:.82; font-size:.88rem; }}
    .hero-chips {{ display:flex; flex-wrap:wrap; gap:7px; justify-content:flex-end; }}
    .hero-chip {{ background:rgba(255,255,255,.11); border:1px solid rgba(255,255,255,.20); border-radius:999px; padding:7px 11px; font-size:.71rem; font-weight:800; white-space:nowrap; }}

    .kpi-card {{
        background:#FFFFFF; border:1px solid {BORDER}; border-radius:17px; padding:15px 16px 14px 16px;
        min-height:136px; height:100%; position:relative; overflow:hidden;
        box-shadow:0 5px 18px rgba(7,27,51,.045); display:flex; flex-direction:column; justify-content:space-between;
    }}
    .kpi-card::after {{ content:""; position:absolute; width:70px; height:70px; border-radius:50%; right:-28px; top:-30px; background:var(--tone,#00A7B5); opacity:.10; }}
    .kpi-top {{ display:flex; align-items:center; gap:9px; }}
    .kpi-dot {{ width:9px; height:9px; border-radius:50%; background:var(--tone,#00A7B5); box-shadow:0 0 0 5px color-mix(in srgb, var(--tone,#00A7B5) 13%, transparent); }}
    .kpi-label {{ color:{GRAY}; font-size:.70rem; font-weight:850; letter-spacing:.055em; text-transform:uppercase; line-height:1.3; }}
    .kpi-value {{ color:{NAVY}; font-size:clamp(1.08rem,1.45vw,1.55rem); font-weight:950; line-height:1.07; margin:11px 0 0 0; letter-spacing:-.035em; white-space:nowrap; }}
    .kpi-note {{ color:{GRAY}; font-size:.72rem; line-height:1.35; margin-top:10px; overflow-wrap:anywhere; }}
    .kpi-delta {{ display:inline-flex; width:max-content; margin-top:8px; border-radius:999px; padding:3px 7px; font-size:.67rem; font-weight:850; background:#EEF7F4; color:{GREEN}; }}
    .kpi-delta.bad {{ background:#FDECEC; color:{RED}; }}

    .section-head {{ display:flex; align-items:flex-end; justify-content:space-between; gap:15px; margin:19px 0 10px 0; padding-bottom:8px; border-bottom:1px solid {BORDER}; }}
    .section-head h3 {{ color:{NAVY}; font-size:1.13rem; font-weight:900; margin:0; letter-spacing:-.015em; }}
    .section-head p {{ color:{GRAY}; font-size:.78rem; margin:4px 0 0 0; }}
    .section-badge {{ display:inline-flex; align-items:center; border-radius:999px; padding:5px 9px; background:#EAF7F8; color:{TEAL}; font-size:.67rem; font-weight:900; text-transform:uppercase; letter-spacing:.04em; white-space:nowrap; }}

    .insight-grid {{ display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:10px; margin-top:4px; }}
    .insight-card {{ background:white; border:1px solid {BORDER}; border-left:4px solid {ORANGE}; border-radius:14px; padding:12px 14px; box-shadow:0 3px 12px rgba(7,27,51,.035); }}
    .insight-card.good {{ border-left-color:{GREEN}; }}
    .insight-card.critical {{ border-left-color:{RED}; }}
    .insight-card strong {{ color:{NAVY}; font-size:.84rem; display:block; margin-bottom:4px; }}
    .insight-card span {{ color:#475467; font-size:.76rem; line-height:1.42; }}

    .scope-note {{ background:#EDF7F8; border:1px solid #CDECEF; border-radius:13px; padding:11px 13px; color:#155C64; font-size:.77rem; line-height:1.45; }}
    .warning-note {{ background:#FFF7E8; border:1px solid #F1D59B; border-radius:13px; padding:11px 13px; color:#684900; font-size:.77rem; line-height:1.45; }}
    .secure-note {{ background:#EDF3FA; border:1px solid #D4E2F0; border-radius:12px; padding:10px 12px; color:#234B70; font-size:.73rem; line-height:1.4; }}
    .user-pill {{ display:flex; align-items:center; justify-content:space-between; gap:8px; background:#F4F7FA; border:1px solid {BORDER}; border-radius:12px; padding:9px 10px; margin:8px 0; }}
    .user-pill b {{ color:{NAVY}; font-size:.78rem; }} .user-pill span {{ color:{GRAY}; font-size:.68rem; }}

    [data-testid="stSidebar"] [role="radiogroup"] label {{ border-radius:10px; padding:3px 4px; }}
    [data-testid="stSidebar"] hr {{ margin:.85rem 0; }}
    [data-testid="stTabs"] button {{ font-weight:800; color:#475467; }}
    [data-testid="stTabs"] button[aria-selected="true"] {{ color:{NAVY}; border-bottom-color:{CYAN}; }}

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
          <div><h3>{title}</h3><p>{subtitle}</p></div>{badge_html}
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
        dict(orientation="h", yanchor="top", y=-0.14, xanchor="left", x=0)
        if legend_bottom
        else dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.01)
    )
    fig.update_layout(
        height=height,
        margin=dict(l=18, r=34, t=78, b=66 if legend_bottom else 25),
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(family="Arial", color="#344054", size=12),
        legend=legend,
        title=dict(font=dict(color=NAVY, size=16), x=.015, xanchor="left", y=.975),
        hoverlabel=dict(bgcolor="white", namelength=-1),
        hovermode="closest",
        separators=",.",
        uniformtext_minsize=9,
        uniformtext_mode="hide",
    )
    fig.update_xaxes(title_text="", showgrid=False, automargin=True, showline=False)
    fig.update_yaxes(title_text="", gridcolor="#EDF1F5", zeroline=False, automargin=True, showline=False)
    return fig


def hide_value_axis(fig: go.Figure, axis: str = "y") -> go.Figure:
    args = dict(showticklabels=False, showgrid=False, zeroline=False, ticks="", showline=False)
    if axis == "x":
        fig.update_xaxes(**args)
    else:
        fig.update_yaxes(**args)
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
        df["_MES"] = pd.to_datetime(df[mcol], errors="coerce").dt.to_period("M")
        df["_META"] = to_number(df[vcol])
        for c in ["GERENTE", "VENDEDOR"]:
            if c in df.columns:
                df[c] = df[c].fillna("Não informado").astype(str).str.strip()

    return {"faturamento": fat, "metas": metas, "metas_gerentes": metas_g, "usuarios": users, "sheet_states": states}


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


def secret_users() -> dict[str, dict]:
    try:
        raw = st.secrets.get("usuarios", {})
        return {str(k): dict(v) for k, v in raw.items()}
    except Exception:
        return {}


def authenticate() -> dict[str, str]:
    users = secret_users()
    if users:
        if "auth_user" not in st.session_state:
            st.markdown(
                f"""
                <div class='dashboard-hero' style='max-width:760px;margin:4.5rem auto 1.3rem auto'>
                  <div class='hero-brand'>First Medical · Controladoria</div>
                  <h1>Acesso ao painel executivo</h1>
                  <p>Entre com seu usuário para acessar somente as informações autorizadas.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            with st.form("login_form", clear_on_submit=False):
                email = st.text_input("E-mail ou usuário")
                password = st.text_input("Senha", type="password")
                submitted = st.form_submit_button("Entrar", width="stretch")
            if submitted:
                found_key = None
                for key, cfg in users.items():
                    identifiers = {norm(key), norm(cfg.get("email", "")), norm(cfg.get("usuario", ""))}
                    if norm(email) in identifiers:
                        found_key = key
                        break
                if found_key:
                    cfg = users[found_key]
                    expected = str(cfg.get("senha_hash", ""))
                    valid = expected and password_hash(password) == expected
                    if not valid and cfg.get("senha") is not None:
                        valid = password == str(cfg.get("senha"))
                    if valid:
                        st.session_state["auth_user"] = {
                            "nome": str(cfg.get("nome", found_key)),
                            "email": str(cfg.get("email", email)),
                            "perfil": norm(cfg.get("perfil", "GESTOR")),
                            "linha": norm(cfg.get("linha", "VENDAS")),
                            "secure": "Sim",
                        }
                        st.rerun()
                st.error("Usuário ou senha inválidos.")
            st.stop()
        return st.session_state["auth_user"]

    # Modo demonstração: facilita validação antes da criação dos segredos.
    with st.sidebar:
        st.warning("Modo demonstração: configure `.streamlit/secrets.toml` antes de publicar.")
        demo = st.selectbox("Perfil de demonstração", ["Diretoria", "Microtech", "Locação", "Vendas", "Endoscopia"])
    if demo == "Diretoria":
        return {"nome": "Diretoria", "email": "modo demonstração", "perfil": "DIRETORIA", "linha": "CONSOLIDADO", "secure": "Não"}
    line = {"Microtech": "MICROTECH", "Locação": "LOCACAO", "Vendas": "VENDAS", "Endoscopia": "ENDOSCOPIA"}[demo]
    return {"nome": f"Gestor {demo}", "email": "modo demonstração", "perfil": "GESTOR", "linha": line, "secure": "Não"}


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


def line_cash_monthly(
    line: str, fat: pd.DataFrame, receitas: pd.DataFrame, custos: pd.DataFrame,
    start: pd.Period, end: pd.Period,
) -> pd.DataFrame:
    months = pd.period_range(start, end, freq="M")
    out = pd.DataFrame({"Mês": months})
    f = period_filter(fat[fat["_LINHA"] == line], start, end).groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_MES": "Mês", "_VALOR": "Faturamento"})
    rbase = receitas[(receitas["_LINHA"] == line) & (~receitas["_NAO_OPERACIONAL"])]
    r = period_filter(rbase, start, end).groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_MES": "Mês", "_VALOR": "Receitas Recebidas"})
    cbase = custos[(custos["_LINHA_DIRETA"] == line) & (custos["_GRUPO_N"] == "SAIDAS OPERACIONAIS")].copy()
    cbase["_CLASSE_CAIXA"] = cbase["PAI"].map(cost_class)
    c = period_filter(cbase, start, end).groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_MES": "Mês", "_VALOR": "Custos Diretos Pagos"})
    v = period_filter(cbase[cbase["_CLASSE_CAIXA"] == "VARIAVEL"], start, end).groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_MES": "Mês", "_VALOR": "Custos Diretos Variáveis"})
    for df in [f, r, c, v]:
        out = out.merge(df, on="Mês", how="left")
    out = out.fillna(0)
    out["Resultado Direto de Caixa"] = out["Receitas Recebidas"] - out["Custos Diretos Pagos"]
    out["Margem Direta de Caixa"] = np.where(out["Receitas Recebidas"] != 0, out["Resultado Direto de Caixa"] / out["Receitas Recebidas"], 0)
    out["Contribuição Direta de Caixa"] = out["Receitas Recebidas"] - out["Custos Diretos Variáveis"]
    out["Margem de Contribuição Direta"] = np.where(out["Receitas Recebidas"] != 0, out["Contribuição Direta de Caixa"] / out["Receitas Recebidas"], 0)
    out["Conversão em Caixa"] = np.where(out["Faturamento"] != 0, out["Receitas Recebidas"] / out["Faturamento"], 0)
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


def line_summary(fat: pd.DataFrame, receitas: pd.DataFrame, custos: pd.DataFrame, start: pd.Period, end: pd.Period, inad: pd.DataFrame | None) -> pd.DataFrame:
    rows = []
    for line in LINES:
        monthly = line_cash_monthly(line, fat, receitas, custos, start, end)
        t = totals_from_monthly(monthly, True)
        r_scope = period_filter(receitas[(receitas["_LINHA"] == line) & (~receitas["_NAO_OPERACIONAL"])], start, end)
        secure = r_scope[r_scope["_MATCH_METHOD"].isin(["Exata", "Similar"])]
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
            "Faturamento": t.get("Faturamento", 0),
            "Conversão em Caixa": t.get("Conversão em Caixa", 0),
            "Cobertura do Mapeamento": safe_div(float(secure["_VALOR"].sum()), float(r_scope["_VALOR"].sum())),
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
          <p>Performance de caixa, rentabilidade direta e gestão por linha de negócio.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div class='user-pill'><div><b>{user['nome']}</b><br><span>{user['email']}</span></div><span>{'Diretoria' if is_director else line_label(user['linha'])}</span></div>",
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
            up_base = st.file_uploader("Substituir BASE BI", type=["xlsx", "xlsm"], key="up_base_v6")
            up_rev = st.file_uploader("Substituir REV2026", type=["xlsx", "xlsm"], key="up_rev_v6")
            up_inad = st.file_uploader("Base de inadimplência (opcional)", type=["xlsx", "xlsm", "csv"], key="up_inad_v6")
            st.caption("Na REV2026, somente abas visíveis são processadas.")

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
        st.markdown(f"<div class='secure-note'><b>Visão restrita:</b> {line_label(scope_choice)}. Outras linhas e custos compartilhados não são exibidos.</div>", unsafe_allow_html=True)

    st.markdown("#### Navegação")
    pages = ["Dashboard", "Recebimentos & inadimplência", "Clientes", "Produtos", "Custos diretos", "Qualidade & metodologia"]
    if is_director:
        pages.insert(1, "Linhas de negócio")
    page = st.radio("Navegação", pages, label_visibility="collapsed")

    with st.expander("Bases ativas", expanded=False):
        st.caption(f"BASE BI: {base_name}")
        st.caption(f"REV2026: {rev_name}")
        st.caption(f"Inadimplência: {inad_name if inad_bytes else 'não carregada'}")

period_text = f"{month_label(start_month)} a {month_label(end_month)}"
scope_text = line_label(scope_choice)

st.markdown(
    f"""
    <div class='dashboard-hero'>
      <div class='hero-grid'>
        <div>
          <div class='hero-brand'>First Medical · Controladoria</div>
          <h1>{'Painel Executivo de Caixa' if scope_choice == 'CONSOLIDADO' else 'Resultado de Caixa · ' + scope_text}</h1>
          <p>Indicadores realizados por recebimento e pagamento, com visão comercial apenas como apoio.</p>
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
lines_table = line_summary(fat, receitas, custos, start_month, end_month, inad)
fat_scope, rec_scope, cost_scope, inad_scope = scoped_data(scope_choice, fat, receitas, custos, inad, start_month, end_month)


# =========================================================
# DASHBOARD
# =========================================================
if page == "Dashboard":
    if scope_choice == "CONSOLIDADO":
        k1, k2, k3, k4 = st.columns(4)
        with k1: card("Receitas operacionais recebidas", brl(company_totals["Receitas Operacionais"]), "Entradas efetivamente recebidas da operação", BLUE)
        with k2: card("Saídas operacionais pagas", brl(company_totals["Saídas Operacionais"]), "Pagamentos operacionais efetivados", RED)
        with k3: card("Resultado operacional de caixa", brl(company_totals["Resultado Operacional de Caixa"]), "Receitas recebidas menos saídas pagas", GREEN if company_totals["Resultado Operacional de Caixa"] >= 0 else RED)
        with k4: card("Margem operacional de caixa", pct(company_totals["Margem de Caixa"]), "Resultado operacional ÷ receitas recebidas", TEAL)

        k5, k6, k7, k8 = st.columns(4)
        with k5: card("EBITDA gerencial de caixa", brl(company_totals["EBITDA Gerencial de Caixa"]), "Resultado antes de IRPJ e CSLL pagos", NAVY)
        with k6: card("Margem de contribuição de caixa", pct(company_totals["Margem de Contribuição Caixa"]), "Receitas recebidas menos saídas variáveis pagas", CYAN)
        with k7: card("Entradas não operacionais", brl(company_totals["Entradas Não Operacionais"]), "Capital de giro e outras fontes não operacionais", ORANGE)
        with k8: card("Qualidade das entradas", pct(company_totals["Qualidade das Entradas"]), "Participação operacional nas entradas classificadas", TEAL)

        c1, c2 = st.columns([1.35, 1])
        with c1:
            fig = go.Figure()
            fig.add_bar(x=company_monthly["Mês Texto"], y=company_monthly["Receitas Operacionais"], name="Receitas recebidas", marker_color=BLUE,
                        text=company_monthly["Receitas Operacionais"].map(compact_money), textposition="outside", cliponaxis=False)
            fig.add_bar(x=company_monthly["Mês Texto"], y=company_monthly["Saídas Operacionais"], name="Saídas pagas", marker_color="#D26A62",
                        text=company_monthly["Saídas Operacionais"].map(compact_money), textposition="outside", cliponaxis=False)
            fig.add_scatter(x=company_monthly["Mês Texto"], y=company_monthly["Resultado Operacional de Caixa"], name="Resultado", mode="lines+markers+text",
                            line=dict(color=GREEN, width=3), text=company_monthly["Resultado Operacional de Caixa"].map(compact_money), textposition="top center")
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
                connector={"line": {"color": "#AAB5C1"}}, increasing={"marker": {"color": GREEN}},
                decreasing={"marker": {"color": RED}}, totals={"marker": {"color": NAVY}},
            ))
            bridge.update_layout(title="Ponte do resultado operacional de caixa")
            hide_value_axis(bridge, "y")
            st.plotly_chart(plot_layout(bridge, 430, False), width="stretch", config={"displayModeBar": False})

        section_header("Visão por linha de negócio", "Receitas recebidas e custos pagos diretamente vinculados a cada linha", "Sem rateio")
        l1, l2 = st.columns([1.25, 1])
        with l1:
            chart_df = lines_table.sort_values("Resultado Direto de Caixa")
            fig = px.bar(chart_df, x="Resultado Direto de Caixa", y="Linha", orientation="h", color="Código",
                         color_discrete_map=LINE_COLORS, title="Resultado direto de caixa por linha")
            fig.update_traces(text=chart_df["Resultado Direto de Caixa"].map(compact_money), textposition="outside", cliponaxis=False,
                              hovertemplate="%{y}<br>Resultado: R$ %{x:,.2f}<extra></extra>")
            fig.update_layout(showlegend=False)
            hide_value_axis(fig, "x")
            st.plotly_chart(plot_layout(fig, 410, False), width="stretch", config={"displayModeBar": False})
        with l2:
            fig = go.Figure()
            fig.add_bar(x=lines_table["Linha"], y=lines_table["Receitas Recebidas"], name="Receitas recebidas", marker_color=BLUE,
                        text=lines_table["Receitas Recebidas"].map(compact_money), textposition="outside", cliponaxis=False)
            fig.add_bar(x=lines_table["Linha"], y=lines_table["Custos Diretos Pagos"], name="Custos diretos", marker_color="#D26A62",
                        text=lines_table["Custos Diretos Pagos"].map(compact_money), textposition="outside", cliponaxis=False)
            fig.update_layout(title="Receitas x custos diretos", barmode="group")
            hide_value_axis(fig, "y")
            st.plotly_chart(plot_layout(fig, 410), width="stretch", config={"displayModeBar": False})

        section_header("Indicadores comerciais de apoio", "Não compõem o resultado de caixa; ajudam a explicar a conversão futura")
        s1, s2, s3, s4 = st.columns(4)
        with s1: card("Faturamento emitido", brl(company_totals["Faturamento"]), "Competência comercial, apresentado apenas como apoio", BLUE)
        with s2: card("Conversão em caixa", pct(company_totals["Conversão em Caixa"]), "Receita operacional recebida ÷ faturamento emitido", TEAL)
        with s3: card("Performance de recebimento", pct(company_totals["Performance de Recebimento"]), "Realizado ÷ previsto", GREEN if company_totals["Performance de Recebimento"] >= .9 else ORANGE)
        with s4: card("Atingimento da meta", pct(company_totals["Atingimento da Meta"]), "Indicador comercial por competência", CYAN)

        insights = []
        neg_months = company_monthly[company_monthly["Resultado Operacional de Caixa"] < 0]
        if not neg_months.empty:
            insights.append({"title": "Meses com caixa operacional negativo", "text": ", ".join(neg_months["Mês Texto"].tolist()), "class": "critical"})
        if company_totals["Qualidade das Entradas"] < .8:
            insights.append({"title": "Dependência de entradas não operacionais", "text": f"Apenas {pct(company_totals['Qualidade das Entradas'])} das entradas classificadas veio da operação."})
        if company_totals["Performance de Recebimento"] < .9:
            gap = company_totals["Recebimento Previsto"] - company_totals["Recebimento Realizado"]
            insights.append({"title": "Recebimentos abaixo do previsto", "text": f"Diferença acumulada de {brl(gap)} no período."})
        if inad is not None and not inad.empty:
            overdue = float(inad["_VALOR_VENCIDO"].sum())
            insights.append({"title": "Inadimplência vencida", "text": f"Saldo vencido identificado de {brl(overdue)}."})
        else:
            insights.append({"title": "Inadimplência ainda não integrada", "text": "Inclua a base para completar a leitura de conversão, atraso e risco de caixa."})
        if receipt_match_stats["coverage_value"] < .95:
            insights.append({"title": "Classificação das receitas por linha", "text": f"{pct(receipt_match_stats['coverage_value'])} do valor recebido foi associado por correspondência exata ou similar; o restante usa regra de fallback."})
        section_header("Pontos de atenção", "Leituras automáticas sem plano de ação ou recomendações prescritivas", "Diagnóstico")
        insight_cards(insights[:6])

    else:
        t = line_totals
        k1, k2, k3, k4 = st.columns(4)
        with k1: card("Receitas recebidas", brl(t["Receitas Recebidas"]), "Receitas de caixa atribuídas à linha", BLUE)
        with k2: card("Custos diretos pagos", brl(t["Custos Diretos Pagos"]), "Somente centro de custo direto da linha", RED)
        with k3: card("Resultado direto de caixa", brl(t["Resultado Direto de Caixa"]), "Receitas recebidas menos custos diretos pagos", GREEN if t["Resultado Direto de Caixa"] >= 0 else RED)
        with k4: card("Margem direta de caixa", pct(t["Margem Direta de Caixa"]), "Resultado direto ÷ receitas recebidas", TEAL)

        k5, k6, k7, k8 = st.columns(4)
        with k5: card("Margem de contribuição direta", pct(t["Margem de Contribuição Direta"]), "Após saídas variáveis diretas pagas", CYAN)
        with k6: card("Faturamento emitido", brl(t["Faturamento"]), "Indicador comercial de apoio", NAVY)
        with k7: card("Conversão em caixa", pct(t["Conversão em Caixa"]), "Receitas recebidas ÷ faturamento emitido", TEAL)
        overdue = float(inad_scope["_VALOR_VENCIDO"].sum()) if inad_scope is not None and not inad_scope.empty else 0
        with k8: card("Inadimplência", brl(overdue) if inad is not None else "Base não carregada", "Saldo vencido atribuído à linha", ORANGE)

        c1, c2 = st.columns([1.35, 1])
        with c1:
            fig = go.Figure()
            fig.add_bar(x=line_monthly["Mês Texto"], y=line_monthly["Receitas Recebidas"], name="Receitas recebidas", marker_color=BLUE,
                        text=line_monthly["Receitas Recebidas"].map(compact_money), textposition="outside", cliponaxis=False)
            fig.add_bar(x=line_monthly["Mês Texto"], y=line_monthly["Custos Diretos Pagos"], name="Custos diretos", marker_color="#D26A62",
                        text=line_monthly["Custos Diretos Pagos"].map(compact_money), textposition="outside", cliponaxis=False)
            fig.add_scatter(x=line_monthly["Mês Texto"], y=line_monthly["Resultado Direto de Caixa"], name="Resultado direto", mode="lines+markers+text",
                            line=dict(color=GREEN, width=3), text=line_monthly["Resultado Direto de Caixa"].map(compact_money), textposition="top center")
            fig.update_layout(title=f"Movimento de caixa · {scope_text}", barmode="group")
            hide_value_axis(fig, "y")
            st.plotly_chart(plot_layout(fig, 430), width="stretch", config={"displayModeBar": False})
        with c2:
            fig = go.Figure()
            fig.add_bar(x=line_monthly["Mês Texto"], y=line_monthly["Faturamento"], name="Faturamento", marker_color=NAVY,
                        text=line_monthly["Faturamento"].map(compact_money), textposition="outside", cliponaxis=False)
            fig.add_scatter(x=line_monthly["Mês Texto"], y=line_monthly["Receitas Recebidas"], name="Recebido", mode="lines+markers+text",
                            line=dict(color=CYAN, width=3), text=line_monthly["Receitas Recebidas"].map(compact_money), textposition="top center")
            fig.update_layout(title="Faturamento x conversão em caixa")
            hide_value_axis(fig, "y")
            st.plotly_chart(plot_layout(fig, 430), width="stretch", config={"displayModeBar": False})

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

        r_scope = rec_scope.copy()
        secure_value = float(r_scope.loc[r_scope["_MATCH_METHOD"].isin(["Exata", "Similar"]), "_VALOR"].sum())
        coverage = safe_div(secure_value, float(r_scope["_VALOR"].sum()))
        st.markdown(
            f"<div class='scope-note'><b>Qualidade da classificação:</b> {pct(coverage)} das receitas desta linha foram associadas por correspondência exata ou similar entre cliente e faturamento. Custos administrativos compartilhados e rateios não são apresentados.</div>",
            unsafe_allow_html=True,
        )


# =========================================================
# LINHAS DE NEGÓCIO — SOMENTE DIRETORIA
# =========================================================
elif page == "Linhas de negócio" and is_director:
    section_header("Comparativo das linhas de negócio", "Resultado direto por caixa; custos compartilhados permanecem apenas no consolidado", "Diretoria")
    cols = st.columns(4)
    for idx, row in lines_table.iterrows():
        with cols[idx % 4]:
            card(row["Linha"], brl(row["Resultado Direto de Caixa"]), f"Margem direta: {pct(row['Margem Direta de Caixa'])}", LINE_COLORS[row["Código"]])

    c1, c2 = st.columns(2)
    with c1:
        p = lines_table.sort_values("Receitas Recebidas")
        fig = px.bar(p, x="Receitas Recebidas", y="Linha", orientation="h", color="Código", color_discrete_map=LINE_COLORS,
                     title="Receitas recebidas por linha")
        fig.update_traces(text=p["Receitas Recebidas"].map(compact_money), textposition="outside", cliponaxis=False)
        fig.update_layout(showlegend=False); hide_value_axis(fig, "x")
        st.plotly_chart(plot_layout(fig, 440, False), width="stretch", config={"displayModeBar": False})
    with c2:
        p = lines_table.sort_values("Margem Direta de Caixa")
        fig = px.bar(p, x="Margem Direta de Caixa", y="Linha", orientation="h", color="Código", color_discrete_map=LINE_COLORS,
                     title="Margem direta de caixa")
        fig.update_traces(text=p["Margem Direta de Caixa"].map(pct), textposition="outside", cliponaxis=False)
        fig.update_layout(showlegend=False); hide_value_axis(fig, "x")
        st.plotly_chart(plot_layout(fig, 440, False), width="stretch", config={"displayModeBar": False})

    view = lines_table.drop(columns=["Código"]).copy()
    for c in ["Receitas Recebidas", "Custos Diretos Pagos", "Resultado Direto de Caixa", "Faturamento", "Inadimplência"]:
        view[c] = view[c].map(brl)
    for c in ["Margem Direta de Caixa", "Margem de Contribuição Direta", "Conversão em Caixa", "Cobertura do Mapeamento"]:
        view[c] = view[c].map(pct)
    st.dataframe(view, width="stretch", hide_index=True, height=280)
    st.download_button("Exportar comparativo das linhas", dataframe_download(lines_table.drop(columns=["Código"]), "Linhas"),
                       file_name=f"resultado_caixa_por_linha_{start_month}_{end_month}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.markdown("<div class='warning-note'><b>Leitura:</b> o resultado direto de cada linha não inclui áreas administrativas, diretoria, financeiro, fiscal, tecnologia ou outros centros compartilhados. Esses valores permanecem no consolidado da diretoria.</div>", unsafe_allow_html=True)


# =========================================================
# RECEBIMENTOS E INADIMPLÊNCIA
# =========================================================
elif page == "Recebimentos & inadimplência":
    section_header("Conversão e risco de recebimento", "Entradas realizadas, previsão e títulos vencidos", scope_text)

    if scope_choice == "CONSOLIDADO":
        r1, r2, r3, r4 = st.columns(4)
        with r1: card("Recebimento realizado", brl(company_totals["Recebimento Realizado"]), "Base de performance de recebimento", BLUE)
        with r2: card("Recebimento previsto", brl(company_totals["Recebimento Previsto"]), "Valor programado para o período", NAVY)
        with r3: card("Performance", pct(company_totals["Performance de Recebimento"]), "Realizado ÷ previsto", GREEN if company_totals["Performance de Recebimento"] >= .9 else ORANGE)
        gap = company_totals["Recebimento Realizado"] - company_totals["Recebimento Previsto"]
        with r4: card("Diferença", brl(gap), "Realizado menos previsto", GREEN if gap >= 0 else RED)
        fig = go.Figure()
        fig.add_bar(x=company_monthly["Mês Texto"], y=company_monthly["Recebimento Previsto"], name="Previsto", marker_color="#A8C6DF",
                    text=company_monthly["Recebimento Previsto"].map(compact_money), textposition="outside", cliponaxis=False)
        fig.add_scatter(x=company_monthly["Mês Texto"], y=company_monthly["Recebimento Realizado"], name="Realizado", mode="lines+markers+text",
                        line=dict(color=TEAL, width=3), text=company_monthly["Recebimento Realizado"].map(compact_money), textposition="top center")
        fig.update_layout(title="Recebimento previsto x realizado")
        hide_value_axis(fig, "y")
        st.plotly_chart(plot_layout(fig, 420), width="stretch", config={"displayModeBar": False})
    else:
        r1, r2, r3, r4 = st.columns(4)
        with r1: card("Receitas recebidas", brl(line_totals["Receitas Recebidas"]), "Receitas atribuídas à linha", BLUE)
        with r2: card("Faturamento emitido", brl(line_totals["Faturamento"]), "Base comercial do período", NAVY)
        with r3: card("Conversão em caixa", pct(line_totals["Conversão em Caixa"]), "Recebido ÷ faturado", TEAL)
        coverage = safe_div(float(rec_scope.loc[rec_scope["_MATCH_METHOD"].isin(["Exata", "Similar"]), "_VALOR"].sum()), float(rec_scope["_VALOR"].sum()))
        with r4: card("Cobertura do vínculo", pct(coverage), "Correspondência exata ou similar de clientes", CYAN)

    section_header("Inadimplência", "A base é opcional e pode ser carregada pela diretoria")
    if inad_scope is None:
        st.info("A base de inadimplência ainda não foi carregada. Campos aceitos: Cliente, Valor vencido, Vencimento e Dias de atraso. Também são reconhecidos `Vencidos Corrigidos`, `Vencidos` e `Valor Original`.")
    elif inad_scope.empty:
        st.success("Não foram encontrados títulos vencidos para o escopo selecionado.")
    else:
        overdue = float(inad_scope["_VALOR_VENCIDO"].sum())
        clients_overdue = int(inad_scope["_CLIENTE"].nunique())
        titles_overdue = len(inad_scope)
        avg_days = float(np.average(inad_scope["_DIAS_ATRASO"], weights=np.maximum(inad_scope["_VALOR_VENCIDO"], .01)))
        i1, i2, i3, i4 = st.columns(4)
        with i1: card("Saldo vencido", brl(overdue), "Valor vencido identificado", RED)
        with i2: card("Clientes inadimplentes", f"{clients_overdue:,}".replace(",", "."), "Clientes únicos com saldo vencido", ORANGE)
        with i3: card("Títulos vencidos", f"{titles_overdue:,}".replace(",", "."), "Quantidade de registros vencidos", NAVY)
        with i4: card("Atraso médio ponderado", f"{avg_days:,.0f} dias".replace(",", "."), "Ponderado pelo valor vencido", TEAL)

        c1, c2 = st.columns(2)
        with c1:
            aging = inad_scope.groupby("_FAIXA", as_index=False, observed=False)["_VALOR_VENCIDO"].sum()
            order = ["Até 30 dias", "31 a 60 dias", "61 a 90 dias", "Acima de 90 dias"]
            aging["_FAIXA"] = pd.Categorical(aging["_FAIXA"], order, ordered=True)
            aging = aging.sort_values("_FAIXA")
            fig = px.bar(aging, x="_FAIXA", y="_VALOR_VENCIDO", title="Aging da inadimplência")
            fig.update_traces(marker_color=ORANGE, text=aging["_VALOR_VENCIDO"].map(compact_money), textposition="outside", cliponaxis=False)
            hide_value_axis(fig, "y")
            st.plotly_chart(plot_layout(fig, 400, False), width="stretch", config={"displayModeBar": False})
        with c2:
            cli = inad_scope.groupby("_CLIENTE", as_index=False)["_VALOR_VENCIDO"].sum().sort_values("_VALOR_VENCIDO", ascending=False).head(12)
            cli["Nome"] = cli["_CLIENTE"].map(lambda x: short_label(x, 38))
            p = cli.sort_values("_VALOR_VENCIDO")
            fig = px.bar(p, x="_VALOR_VENCIDO", y="Nome", orientation="h", title="Maiores saldos vencidos", custom_data=["_CLIENTE"])
            fig.update_traces(marker_color=RED, text=p["_VALOR_VENCIDO"].map(compact_money), textposition="outside", cliponaxis=False,
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
            st.caption("O detalhamento do arquivo não contém a coluna Gerente. Para restringir os títulos por gestor, o app relaciona cada cliente ao gerente dominante da BASE BI.")

        detail = inad_scope[["_CLIENTE", "_TITULO", "_VENCIMENTO", "_DIAS_ATRASO", "_FAIXA", "_VALOR_VENCIDO", "_LINHA", "_GERENTE"]].copy()
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
    section_header("Carteira de clientes", "Faturamento, recebimento e concentração no escopo autorizado", scope_text)
    client_col = "NOME DO CLIENTE" if "NOME DO CLIENTE" in fat_scope.columns else "CLIENTE"
    billed = fat_scope.groupby(client_col, as_index=False).agg(Faturamento=("_VALOR", "sum"), Notas=("Nota Fiscal", "nunique") if "Nota Fiscal" in fat_scope.columns else ("_VALOR", "size"))
    received = rec_scope.groupby("Cliente", as_index=False)["_VALOR"].sum().rename(columns={"Cliente": "Cliente Recebimento", "_VALOR": "Recebido"})
    total_billed = float(billed["Faturamento"].sum())
    billed["Participação"] = np.where(total_billed != 0, billed["Faturamento"] / total_billed, 0)
    billed = billed.sort_values("Faturamento", ascending=False)

    c1, c2, c3, c4 = st.columns(4)
    with c1: card("Clientes faturados", f"{len(billed):,}".replace(",", "."), "Clientes com emissão no período", BLUE)
    with c2: card("Concentração Top 5", pct(float(billed.head(5)["Participação"].sum())), "Participação dos cinco maiores", ORANGE)
    with c3: card("Ticket médio por nota", brl(safe_div(total_billed, float(billed["Notas"].sum()))), "Faturamento ÷ notas", TEAL)
    with c4: card("Clientes com recebimento", f"{rec_scope['Cliente'].nunique():,}".replace(",", "."), "Nomes presentes na base de caixa", GREEN)

    p1, p2 = st.columns(2)
    with p1:
        top = billed.head(15).copy(); top["Nome"] = top[client_col].map(lambda x: short_label(x, 38)); p = top.sort_values("Faturamento")
        fig = px.bar(p, x="Faturamento", y="Nome", orientation="h", title="Principais clientes por faturamento", custom_data=[client_col])
        fig.update_traces(marker_color=TEAL, text=p["Faturamento"].map(compact_money), textposition="outside", cliponaxis=False,
                          hovertemplate="%{customdata[0]}<br>Faturamento: R$ %{x:,.2f}<extra></extra>")
        hide_value_axis(fig, "x")
        st.plotly_chart(plot_layout(fig, 560, False), width="stretch", config={"displayModeBar": False})
    with p2:
        top = received.sort_values("Recebido", ascending=False).head(15).copy(); top["Nome"] = top["Cliente Recebimento"].map(lambda x: short_label(x, 38)); p = top.sort_values("Recebido")
        fig = px.bar(p, x="Recebido", y="Nome", orientation="h", title="Principais clientes por recebimento", custom_data=["Cliente Recebimento"])
        fig.update_traces(marker_color=BLUE, text=p["Recebido"].map(compact_money), textposition="outside", cliponaxis=False,
                          hovertemplate="%{customdata[0]}<br>Recebido: R$ %{x:,.2f}<extra></extra>")
        hide_value_axis(fig, "x")
        st.plotly_chart(plot_layout(fig, 560, False), width="stretch", config={"displayModeBar": False})

    view = billed.head(50).rename(columns={client_col: "Cliente"}).copy()
    view["Faturamento"] = view["Faturamento"].map(brl)
    view["Participação"] = view["Participação"].map(pct)
    st.dataframe(view, width="stretch", hide_index=True, height=430)
    st.caption("Faturamento e recebimentos são apresentados separadamente porque as bases podem usar grafias diferentes para o mesmo cliente.")


# =========================================================
# PRODUTOS
# =========================================================
elif page == "Produtos":
    section_header("Performance de produtos", "Faturamento, volume, concentração e carteira no escopo autorizado", scope_text)
    if fat_scope.empty:
        st.info("Não há faturamento no período e escopo selecionados.")
    else:
        product_col = optional_col(fat_scope, ["PRODUTO", "DESCRIÇÃO", "LINHA DE PRODUTO", "ITEM"])
        qty_col = optional_col(fat_scope, ["QUANTIDADE", "QTD", "QTDE", "QTD FATURADA", "QUANTIDADE FATURADA"])
        client_col = "NOME DO CLIENTE" if "NOME DO CLIENTE" in fat_scope.columns else "CLIENTE"
        if product_col is None:
            st.info("A base de faturamento não possui uma coluna de produto reconhecida.")
        else:
            prod_base = fat_scope.copy()
            prod_base["_PRODUTO"] = prod_base[product_col].fillna("Não informado").astype(str).str.strip()
            prod_base["_QTD"] = to_number(prod_base[qty_col]) if qty_col else 1.0
            prod_base.loc[prod_base["_QTD"] <= 0, "_QTD"] = 1.0
            nf_col = "Nota Fiscal" if "Nota Fiscal" in prod_base.columns else None
            agg_map = {
                "Faturamento": ("_VALOR", "sum"),
                "Quantidade": ("_QTD", "sum"),
                "Clientes": (client_col, "nunique"),
            }
            if nf_col:
                agg_map["Notas"] = (nf_col, "nunique")
            else:
                agg_map["Notas"] = ("_VALOR", "size")
            products = prod_base.groupby("_PRODUTO", as_index=False).agg(**agg_map)
            total_products = float(products["Faturamento"].sum())
            products["Participação"] = np.where(total_products != 0, products["Faturamento"] / total_products, 0)
            products["Preço médio"] = np.where(products["Quantidade"] != 0, products["Faturamento"] / products["Quantidade"], 0)
            products = products.sort_values("Faturamento", ascending=False)
            top_product = products.iloc[0]
            top10_share = float(products.head(10)["Faturamento"].sum() / total_products) if total_products else 0

            k1, k2, k3, k4, k5 = st.columns(5)
            with k1: card("Faturamento de produtos", brl(total_products), "Valor emitido no período", BLUE)
            with k2: card("Produtos ativos", f"{len(products):,}".replace(",", "."), "Itens com faturamento", NAVY)
            with k3: card("Quantidade", f"{products['Quantidade'].sum():,.0f}".replace(",", "."), f"Fonte: {qty_col or 'linhas faturadas'}", TEAL)
            with k4: card("Produto líder", short_label(top_product["_PRODUTO"], 24), brl(top_product["Faturamento"]), GREEN)
            with k5: card("Concentração Top 10", pct(top10_share), "Participação dos dez maiores produtos", ORANGE if top10_share > .70 else CYAN)

            c1, c2 = st.columns([1.25, 1])
            with c1:
                top = products.head(15).sort_values("Faturamento")
                top["Produto"] = top["_PRODUTO"].map(lambda x: short_label(x, 42))
                fig = px.bar(top, x="Faturamento", y="Produto", orientation="h", title="Produtos com maior faturamento", custom_data=["_PRODUTO"])
                fig.update_traces(marker_color=BLUE, text=top["Faturamento"].map(compact_money), textposition="outside", cliponaxis=False,
                                  hovertemplate="%{customdata[0]}<br>Faturamento: R$ %{x:,.2f}<extra></extra>")
                hide_value_axis(fig, "x")
                st.plotly_chart(plot_layout(fig, 520, False), width="stretch", config={"displayModeBar": False})
            with c2:
                qty_top = products.sort_values("Quantidade", ascending=False).head(12).sort_values("Quantidade")
                qty_top["Produto"] = qty_top["_PRODUTO"].map(lambda x: short_label(x, 34))
                fig = px.bar(qty_top, x="Quantidade", y="Produto", orientation="h", title="Produtos com maior volume")
                fig.update_traces(marker_color=TEAL, text=qty_top["Quantidade"].map(lambda v: f"{v:,.0f}".replace(",", ".")), textposition="outside", cliponaxis=False)
                hide_value_axis(fig, "x")
                st.plotly_chart(plot_layout(fig, 520, False), width="stretch", config={"displayModeBar": False})

            section_header("Detalhamento dos produtos", "Receita, volume, clientes e preço médio")
            table = products.rename(columns={"_PRODUTO": "Produto"})[["Produto", "Faturamento", "Participação", "Quantidade", "Preço médio", "Clientes", "Notas"]].copy()
            table_view = table.copy()
            table_view["Faturamento"] = table_view["Faturamento"].map(brl)
            table_view["Participação"] = table_view["Participação"].map(pct)
            table_view["Quantidade"] = table_view["Quantidade"].map(lambda v: f"{v:,.0f}".replace(",", "."))
            table_view["Preço médio"] = table_view["Preço médio"].map(brl)
            st.dataframe(table_view, width="stretch", hide_index=True, height=470)
            st.download_button("Exportar análise de produtos", dataframe_download(table, "Produtos"),
                               file_name=f"produtos_{scope_choice}_{start_month}_{end_month}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# =========================================================
# CUSTOS DIRETOS
# =========================================================
elif page == "Custos diretos":
    section_header("Custos diretos pagos", "Somente centros de custo diretamente identificados com a linha; rateios não são exibidos", scope_text)
    if cost_scope.empty:
        st.info("Não foram encontrados custos diretos para o escopo selecionado.")
    else:
        supplier_col = "Codigo-Nome do Fornecedor"
        total_cost = float(cost_scope["_VALOR"].sum())
        suppliers = cost_scope.groupby(supplier_col, as_index=False).agg(Valor=("_VALOR", "sum"), Lançamentos=("_VALOR", "size")).sort_values("Valor", ascending=False)
        nature = cost_scope.groupby("PAI", as_index=False)["_VALOR"].sum().sort_values("_VALOR", ascending=False)
        months = cost_scope.groupby("_MES", as_index=False)["_VALOR"].sum(); months["Mês"] = months["_MES"].map(month_label)
        c1, c2, c3, c4 = st.columns(4)
        with c1: card("Custos diretos pagos", brl(total_cost), "Total do período no escopo", RED)
        with c2: card("Fornecedores", f"{len(suppliers):,}".replace(",", "."), "Contrapartes com pagamento direto", BLUE)
        with c3: card("Maior fornecedor", brl(float(suppliers.iloc[0]["Valor"])), short_label(suppliers.iloc[0][supplier_col], 34), ORANGE)
        with c4: card("Ticket médio", brl(safe_div(total_cost, float(suppliers["Lançamentos"].sum()))), "Valor médio por lançamento", TEAL)

        p1, p2 = st.columns(2)
        with p1:
            p = nature.head(15).sort_values("_VALOR").copy(); p["Natureza"] = p["PAI"].map(lambda x: short_label(x, 36))
            fig = px.bar(p, x="_VALOR", y="Natureza", orientation="h", title="Principais naturezas de custo", custom_data=["PAI"])
            fig.update_traces(marker_color=BLUE, text=p["_VALOR"].map(compact_money), textposition="outside", cliponaxis=False,
                              hovertemplate="%{customdata[0]}<br>Valor: R$ %{x:,.2f}<extra></extra>")
            hide_value_axis(fig, "x")
            st.plotly_chart(plot_layout(fig, 520, False), width="stretch", config={"displayModeBar": False})
        with p2:
            p = suppliers.head(15).sort_values("Valor").copy(); p["Fornecedor"] = p[supplier_col].map(lambda x: short_label(x, 38))
            fig = px.bar(p, x="Valor", y="Fornecedor", orientation="h", title="Principais fornecedores", custom_data=[supplier_col])
            fig.update_traces(marker_color=TEAL, text=p["Valor"].map(compact_money), textposition="outside", cliponaxis=False,
                              hovertemplate="%{customdata[0]}<br>Valor: R$ %{x:,.2f}<extra></extra>")
            hide_value_axis(fig, "x")
            st.plotly_chart(plot_layout(fig, 520, False), width="stretch", config={"displayModeBar": False})

        fig = px.bar(months, x="Mês", y="_VALOR", title="Evolução mensal dos custos diretos")
        fig.update_traces(marker_color=NAVY, text=months["_VALOR"].map(compact_money), textposition="outside", cliponaxis=False)
        hide_value_axis(fig, "y")
        st.plotly_chart(plot_layout(fig, 380, False), width="stretch", config={"displayModeBar": False})

        table = suppliers.rename(columns={supplier_col: "Fornecedor"}).copy()
        table["Participação"] = np.where(total_cost != 0, table["Valor"] / total_cost, 0)
        table_view = table.copy()
        table_view["Valor"] = table_view["Valor"].map(brl)
        table_view["Participação"] = table_view["Participação"].map(pct)
        st.dataframe(table_view, width="stretch", hide_index=True, height=420)
        st.download_button("Exportar custos diretos", dataframe_download(table, "Custos Diretos"),
                           file_name=f"custos_diretos_{scope_choice}_{start_month}_{end_month}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# =========================================================
# QUALIDADE E METODOLOGIA
# =========================================================
elif page == "Qualidade & metodologia":
    section_header("Qualidade, segurança e metodologia", "Transparência sobre fontes, vínculos e limitações", "Governança")
    q1, q2, q3, q4 = st.columns(4)
    with q1: card("Cobertura das receitas por linha", pct(receipt_match_stats["coverage_value"]), "Valor associado por correspondência exata ou similar", TEAL)
    with q2: card("Receitas por fallback", brl(receipt_match_stats["fallback_value"]), "Classificação inferida pelo tipo de recebimento", ORANGE)
    with q3: card("Segurança ativa", "Sim" if user["secure"] == "Sim" else "Modo demonstração", "Controle de perfil por Streamlit Secrets", GREEN if user["secure"] == "Sim" else ORANGE)
    with q4: card("Inadimplência integrada", "Sim" if inad is not None else "Não", inad_name if inad is not None else "Base opcional ainda não carregada", BLUE)

    c1, c2 = st.columns(2)
    with c1:
        visible = [s for s, state in rev["sheet_states"].items() if state == "visible"]
        st.markdown("#### Abas REV2026 processadas")
        st.success("Somente abas visíveis são consideradas.")
        st.dataframe(pd.DataFrame({"Aba visível": visible}), width="stretch", hide_index=True)
    with c2:
        hidden = [s for s, state in rev["sheet_states"].items() if state != "visible"]
        st.markdown("#### Abas REV2026 ignoradas")
        st.warning("As abas abaixo estavam ocultas e foram desprezadas.")
        st.dataframe(pd.DataFrame({"Aba oculta": hidden}), width="stretch", hide_index=True)

    methodology = pd.DataFrame([
        ["Receitas operacionais recebidas", "Soma de Valor Pago em Venda, Locação e Serviço.", "Caixa"],
        ["Entradas não operacionais", "Capital de giro e outras receitas que aumentam o caixa, mas não representam operação.", "Caixa / financiamento"],
        ["Saídas operacionais pagas", "Pagamentos realizados classificados como saídas operacionais.", "Caixa"],
        ["Resultado operacional de caixa", "Receitas operacionais recebidas menos saídas operacionais pagas.", "Caixa"],
        ["EBITDA gerencial de caixa", "Resultado operacional de caixa com IRPJ e CSLL adicionados de volta.", "Caixa gerencial"],
        ["Margem de contribuição de caixa", "Receitas operacionais recebidas menos saídas variáveis pagas.", "Caixa gerencial"],
        ["Resultado direto da linha", "Receitas recebidas atribuídas à linha menos custos pagos no centro de custo direto da linha.", "Caixa direto"],
        ["Custos compartilhados", "Não são exibidos aos gestores e permanecem apenas no consolidado da diretoria.", "Governança"],
        ["Faturamento e meta", "Exibidos somente como indicadores comerciais de apoio; não formam o resultado de caixa.", "Competência"],
        ["Inadimplência", "Saldo vencido da base opcional, classificado por linha quando possível.", "Carteira"],
    ], columns=["Indicador", "Definição", "Regime"])
    st.dataframe(methodology, width="stretch", hide_index=True, height=390)

    st.markdown(
        "<div class='scope-note'><b>Acesso por gestor:</b> os perfis de Microtech, Locação, Vendas e Endoscopia visualizam apenas receitas recebidas, faturamento, produtos, clientes, inadimplência e custos diretos da própria linha. Não há comparação com outras linhas nem exposição de áreas compartilhadas.</div>",
        unsafe_allow_html=True,
    )
    if user["secure"] != "Sim":
        st.markdown("<div class='warning-note'><b>Antes da publicação:</b> copie `secrets.toml.example` para os segredos do Streamlit e cadastre um usuário por gestor. O modo demonstração não deve ser usado como controle de acesso em produção.</div>", unsafe_allow_html=True)

st.caption("FIRST MEDICAL · Intelligence Dashboard · Controladoria · Indicadores de caixa e acesso por linha")
