from __future__ import annotations

import io
import os
import re
import unicodedata
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl import load_workbook


# =========================================================
# CONFIGURAÇÃO
# =========================================================
st.set_page_config(
    page_title="KPIs Executivos | First Medical",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

NAVY = "#0B1F3A"
BLUE = "#1261A0"
TEAL = "#008C95"
GREEN = "#1F7A4D"
RED = "#B42318"
ORANGE = "#B76E00"
LIGHT = "#F4F7FA"
GRAY = "#667085"

st.markdown(
    f"""
    <style>
    .stApp {{ background: {LIGHT}; }}
    [data-testid="stSidebar"] {{ background: #FFFFFF; border-right: 1px solid #E6EAF0; }}
    .block-container {{ padding-top: 1.1rem; padding-bottom: 2.5rem; max-width: 1500px; }}
    .hero {{
        background: linear-gradient(115deg, {NAVY} 0%, #113B68 68%, {TEAL} 140%);
        color: white; padding: 22px 26px; border-radius: 18px; margin-bottom: 16px;
        box-shadow: 0 8px 24px rgba(11,31,58,.15); overflow: hidden;
    }}
    .hero h1 {{ margin: 0; font-size: clamp(1.45rem, 2.5vw, 1.90rem); letter-spacing: -.02em; line-height: 1.15; }}
    .hero p {{ margin: 7px 0 0 0; opacity: .88; line-height: 1.4; }}
    .kpi-card {{
        background: white; border: 1px solid #E7EBF0; border-radius: 15px;
        padding: 15px 15px 14px 15px; min-height: 142px; height: 100%;
        box-shadow: 0 3px 12px rgba(16,24,40,.05); overflow: hidden;
        display: flex; flex-direction: column; justify-content: space-between;
    }}
    .kpi-label {{ color: {GRAY}; font-size: .73rem; font-weight: 700; letter-spacing: .04em; text-transform: uppercase; line-height: 1.25; }}
    .kpi-value {{ color: {NAVY}; font-size: clamp(1.00rem, 1.35vw, 1.35rem); font-weight: 800; margin-top: 8px; line-height: 1.12; white-space: nowrap; letter-spacing: -.02em; }}
    .kpi-note {{ color: {GRAY}; font-size: .73rem; margin-top: 9px; line-height: 1.35; overflow-wrap: anywhere; }}
    .section-title {{ color: {NAVY}; font-size: 1.15rem; font-weight: 800; margin: 8px 0 6px 0; }}
    .insight {{ background: white; border-left: 4px solid {TEAL}; border-radius: 10px; padding: 12px 14px; margin: 7px 0; }}
    .method-note {{ background: #FFF7E8; border: 1px solid #F1D59B; border-radius: 12px; padding: 13px 15px; color: #684900; }}
    div[data-testid="stMetric"] {{ background: white; border: 1px solid #E7EBF0; border-radius: 14px; padding: 12px 14px; }}
    div[data-testid="stDataFrame"] {{ border-radius: 12px; overflow: hidden; }}
    .small-muted {{ color: {GRAY}; font-size: .79rem; }}
    .kpi-help {{ background: #FFFFFF; border: 1px solid #E7EBF0; border-radius: 12px; padding: 13px 15px; margin-bottom: 10px; line-height: 1.45; }}
    .kpi-help b {{ color: {NAVY}; }}
    @media (max-width: 900px) {{
        .block-container {{ padding-left: .8rem; padding-right: .8rem; }}
        .hero {{ padding: 18px 18px; }}
        .kpi-card {{ min-height: 128px; }}
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
    text = str(value).strip()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", text).upper()


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    seen: dict[str, int] = {}
    cols = []
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
    txt = series.astype(str).str.strip()
    txt = txt.str.replace(r"R\$\s*", "", regex=True).str.replace(" ", "", regex=False)
    both = txt.str.contains(",", na=False) & txt.str.contains(r"\.", na=False)
    txt.loc[both] = txt.loc[both].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    only_comma = txt.str.contains(",", na=False) & ~txt.str.contains(r"\.", na=False)
    txt.loc[only_comma] = txt.loc[only_comma].str.replace(",", ".", regex=False)
    return pd.to_numeric(txt, errors="coerce").fillna(0.0)


def brl(value: float, compact: bool = False) -> str:
    value = float(value or 0)
    if compact:
        av = abs(value)
        if av >= 1_000_000_000:
            return f"R$ {value / 1_000_000_000:,.2f} bi".replace(",", "X").replace(".", ",").replace("X", ".")
        if av >= 1_000_000:
            return f"R$ {value / 1_000_000:,.2f} mi".replace(",", "X").replace(".", ",").replace("X", ".")
        if av >= 1_000:
            return f"R$ {value / 1_000:,.1f} mil".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pct(value: float) -> str:
    return f"{float(value or 0) * 100:,.1f}%".replace(",", "X").replace(".", ",").replace("X", ".")


def safe_div(a: float, b: float) -> float:
    return float(a / b) if b not in (0, None) and not pd.isna(b) else 0.0


def month_label(period: pd.Period) -> str:
    names = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
    return f"{names[period.month - 1]}/{str(period.year)[2:]}"


def card(label: str, value: str, note: str = "") -> None:
    st.markdown(
        f"""
        <div class="kpi-card">
          <div class="kpi-label">{label}</div>
          <div class="kpi-value">{value}</div>
          <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def plot_layout(fig: go.Figure, height: int = 390, legend_bottom: bool = True) -> go.Figure:
    legend = (
        dict(orientation="h", yanchor="top", y=-0.16, xanchor="left", x=0)
        if legend_bottom
        else dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.02)
    )
    fig.update_layout(
        height=height,
        margin=dict(l=18, r=32, t=88, b=70 if legend_bottom else 30),
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(family="Arial", color="#344054", size=12),
        legend=legend,
        title=dict(font=dict(color=NAVY, size=16), x=0.01, xanchor="left", y=0.98),
        hoverlabel=dict(bgcolor="white", namelength=-1),
        hovermode="closest",
        separators=",.",
    )
    fig.update_xaxes(title_text="", showgrid=False, automargin=True, showline=False)
    fig.update_yaxes(title_text="", gridcolor="#EDF0F4", zeroline=False, automargin=True, showline=False)
    return fig


def apply_brl_axis(fig: go.Figure, axis: str = "y") -> go.Figure:
    # Mantém o formato monetário no hover e para eventuais eixos visíveis,
    # mesmo quando escondemos os ticks para priorizar rótulos nas barras.
    axis_update = dict(
        tickprefix="R$ ",
        tickformat=",.2f",
        separatethousands=True,
        exponentformat="none",
        showexponent="none",
        automargin=True,
    )
    if axis == "x":
        fig.update_xaxes(**axis_update)
    else:
        fig.update_yaxes(**axis_update)
    return fig


def compact_currency_label(value: float) -> str:
    value = float(value or 0)
    av = abs(value)
    if av >= 1_000_000_000:
        return f"{value / 1_000_000_000:,.1f} bi".replace(",", "X").replace(".", ",").replace("X", ".")
    if av >= 1_000_000:
        return f"{value / 1_000_000:,.1f} mi".replace(",", "X").replace(".", ",").replace("X", ".")
    if av >= 1_000:
        return f"{value / 1_000:,.1f} mil".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{value:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")


def compact_percent_label(value: float) -> str:
    return f"{float(value or 0) * 100:,.1f}%".replace(",", "X").replace(".", ",").replace("X", ".")


def hide_value_axis(fig: go.Figure, axis: str = "y") -> go.Figure:
    axis_update = dict(showticklabels=False, showgrid=False, zeroline=False, ticks="", showline=False)
    if axis == "x":
        fig.update_xaxes(**axis_update)
    else:
        fig.update_yaxes(**axis_update)
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
        header_fmt = workbook.add_format({"bold": True, "font_color": "white", "bg_color": NAVY, "border": 0})
        money_fmt = workbook.add_format({"num_format": 'R$ #,##0.00;[Red]-R$ #,##0.00'})
        pct_fmt = workbook.add_format({"num_format": '0.0%'})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_fmt)
            width = min(max(len(str(value)) + 2, 13), 38)
            worksheet.set_column(col_num, col_num, width)
        for idx, col in enumerate(df.columns):
            c = str(col).upper()
            if any(x in c for x in ["VALOR", "FATURAMENTO", "META", "EBITDA", "CUSTO", "DESPESA", "RECEITA"]):
                worksheet.set_column(idx, idx, 18, money_fmt)
            if "%" in c or "MARGEM" in c or "ATINGIMENTO" in c or "PERFORMANCE" in c:
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


def nature_key(value: object) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    if isinstance(value, (int, float, np.number)):
        return f"{float(value):.10f}".rstrip("0").rstrip(".")
    text = str(value).strip().replace(",", ".")
    try:
        return f"{float(text):.10f}".rstrip("0").rstrip(".")
    except ValueError:
        return norm(text)


def require_col(df: pd.DataFrame, candidates: list[str], context: str) -> str:
    lookup = {norm(c): c for c in df.columns}
    for cand in candidates:
        if norm(cand) in lookup:
            return lookup[norm(cand)]
    raise ValueError(f"Não encontrei a coluna necessária em {context}: {', '.join(candidates)}")


@st.cache_data(show_spinner=False)
def load_base_bi(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    states = workbook_sheet_states(file_bytes)
    needed = ["BANCO DE DADOS FATURAMENTO", "Metas", "Metas Gerentes"]
    missing = [s for s in needed if s not in states]
    if missing:
        raise ValueError(f"Abas ausentes na BASE BI: {', '.join(missing)}")

    fat = read_excel_sheet(file_bytes, "BANCO DE DADOS FATURAMENTO")
    metas = read_excel_sheet(file_bytes, "Metas")
    metas_g = read_excel_sheet(file_bytes, "Metas Gerentes")

    date_col = require_col(fat, ["DT Emissao", "MÊS"], "BANCO DE DADOS FATURAMENTO")
    month_col = require_col(fat, ["MÊS", "DT Emissao"], "BANCO DE DADOS FATURAMENTO")
    value_col = require_col(fat, ["VALOR BRUTO", "VALOR ", "VALOR"], "BANCO DE DADOS FATURAMENTO")

    fat["_DATA"] = pd.to_datetime(fat[date_col], errors="coerce")
    fat["_MES"] = pd.to_datetime(fat[month_col], errors="coerce").dt.to_period("M")
    fat["_VALOR"] = to_number(fat[value_col])

    text_cols = [
        "GERENTE", "VENDEDOR / REPRESENTANTE", "VENDEDOR", "SEGMENTO", "EMPRESA", "NOVA",
        "NOME DO CLIENTE", "CLIENTE", "LINHA DE PRODUTO", "PRODUTO", "Nota Fiscal", "CATEGORIA",
    ]
    for col in text_cols:
        if col in fat.columns:
            fat[col] = fat[col].fillna("Não informado").astype(str).str.strip()

    for df, label in [(metas, "Metas"), (metas_g, "Metas Gerentes")]:
        mcol = require_col(df, ["MÊS"], label)
        vcol = require_col(df, ["META MENSAL"], label)
        df["_MES"] = pd.to_datetime(df[mcol], errors="coerce").dt.to_period("M")
        df["_META"] = to_number(df[vcol])
        for c in ["GERENTE", "VENDEDOR"]:
            if c in df.columns:
                df[c] = df[c].fillna("Não informado").astype(str).str.strip()

    return {"faturamento": fat, "metas": metas, "metas_gerentes": metas_g, "sheet_states": states}


@st.cache_data(show_spinner=False)
def load_rev(file_bytes: bytes) -> dict[str, pd.DataFrame | dict[str, str]]:
    states = workbook_sheet_states(file_bytes)
    visible = {name for name, state in states.items() if state == "visible"}
    required = [
        "Resumo Receitas", "Resumo Despesas", "Centro de Custos",
        "Caixa Operacional", "Performance Recebimento"
    ]
    missing_visible = [s for s in required if s not in visible]
    if missing_visible:
        raise ValueError(
            "As seguintes abas não estão disponíveis como visíveis na REV2026: " + ", ".join(missing_visible)
        )

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

    # A aba Resumo Despesas é a fonte oficial do total de saídas.
    # O mapa auxiliar O:S transforma a natureza contábil na natureza gerencial detalhada.
    mapping = mapping_raw.iloc[2:].copy()
    mapping.columns = ["Natureza_map", "Categoria_map", "PAI_map", "Classificacao_map", "Grupo_map"]
    mapping["_NATUREZA_KEY"] = mapping["Natureza_map"].map(nature_key)
    mapping = mapping[mapping["_NATUREZA_KEY"] != ""].drop_duplicates("_NATUREZA_KEY", keep="last")

    despesas["_NATUREZA_KEY"] = despesas[require_col(despesas, ["Natureza"], "Resumo Despesas")].map(nature_key)
    despesas = despesas.merge(
        mapping[["_NATUREZA_KEY", "Categoria_map", "PAI_map", "Classificacao_map", "Grupo_map"]],
        on="_NATUREZA_KEY", how="left"
    )
    despesas["PAI_ORIGINAL"] = despesas[require_col(despesas, ["PAI"], "Resumo Despesas")]
    despesas["PAI"] = despesas["PAI_map"].where(despesas["PAI_map"].notna(), despesas["PAI_ORIGINAL"])
    despesas["PAI"] = despesas["PAI"].fillna("Não informado").astype(str).str.strip()
    despesas["_MES"] = pd.to_datetime(despesas[require_col(despesas, ["Vencto Real"], "Resumo Despesas")], errors="coerce").dt.to_period("M")
    despesas["_VALOR"] = to_number(despesas[require_col(despesas, ["Valor"], "Resumo Despesas")])
    for col in ["EMPRESA", "GRUPO", "SUBGRUPO", "Categoria", "Codigo-Nome do Fornecedor"]:
        if col in despesas.columns:
            despesas[col] = despesas[col].fillna("Não informado").astype(str).str.strip()
    despesas["_GRUPO_N"] = despesas["GRUPO"].map(norm)
    despesas = despesas[
        despesas["_MES"].notna() & despesas["_GRUPO_N"].isin(["SAIDAS OPERACIONAIS", "SAIDAS NAO OPERACIONAIS"])
    ].copy()

    # Centro de Custos é usado para abertura por área/rateio e para receitas operacionais.
    custos["_MES"] = pd.to_datetime(custos[require_col(custos, ["MÊS"], "Centro de Custos")], errors="coerce").dt.to_period("M")
    custos["_VALOR"] = to_number(custos[require_col(custos, ["Valor"], "Centro de Custos")])
    for col in ["EMPRESA", "GRUPO", "SUBGRUPO", "PAI", "Categoria", "CENTRO DE CUSTOS", "CENTRO DE CUSTOS RATEAO", "Codigo-Nome do Fornecedor"]:
        if col in custos.columns:
            custos[col] = custos[col].fillna("Não informado").astype(str).str.strip()
    custos["_EMPRESA_N"] = custos["EMPRESA"].map(norm)
    custos["_GRUPO_N"] = custos["GRUPO"].map(norm)
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
        "receitas": receitas,
        "despesas": despesas,
        "custos": custos,
        "caixa": caixa,
        "recebimento": receb,
        "sheet_states": states,
    }


# =========================================================
# CLASSIFICAÇÃO DOS CUSTOS
# =========================================================
VARIABLE_COSTS = {
    "FECHAMENTOS DE CAMBIO", "DESEMBARACOS", "PECAS/ACESSORIOS E EMBALAGENS", "LOCACAO DE EQPTOS.",
    "COFINS", "COMISSOES DE REPRESENTANTES", "COMISSOES ESPECIAIS", "CARTOES DE CREDITO",
    "MAQUINAS/EQPTOS", "PIS", "OUTROS EQUIPAMENTOS", "TRANSPORTADORAS", "ICMS", "ISS",
}


def default_classification(pai: str) -> str:
    n = norm(pai)
    if "IRPJ" in n or "CSLL" in n:
        return "Excluir EBITDA"
    if n in VARIABLE_COSTS:
        return "Variável"
    return "Fixo/Operacional"


def init_classifications(costs: pd.DataFrame) -> None:
    pais = sorted(costs.loc[costs["_GRUPO_N"] == "SAIDAS OPERACIONAIS", "PAI"].dropna().astype(str).unique())
    default_map = {p: default_classification(p) for p in pais}
    if "cost_classifications" not in st.session_state:
        st.session_state["cost_classifications"] = default_map
    else:
        current = st.session_state["cost_classifications"]
        for p, cls in default_map.items():
            current.setdefault(p, cls)


# =========================================================
# CÁLCULOS
# =========================================================
def period_filter(df: pd.DataFrame, start: pd.Period, end: pd.Period) -> pd.DataFrame:
    return df[df["_MES"].between(start, end)].copy()


def filtered_meta(
    metas: pd.DataFrame,
    metas_g: pd.DataFrame,
    start: pd.Period,
    end: pd.Period,
    managers: list[str] | None = None,
    sellers: list[str] | None = None,
) -> float:
    managers = managers or []
    sellers = sellers or []
    if sellers and "VENDEDOR" in metas.columns:
        x = period_filter(metas, start, end)
        return float(x[x["VENDEDOR"].isin(sellers)]["_META"].sum())
    if managers and "GERENTE" in metas_g.columns:
        x = period_filter(metas_g, start, end)
        return float(x[x["GERENTE"].isin(managers)]["_META"].sum())
    return float(period_filter(metas, start, end)["_META"].sum())


def calculate_monthly(
    fat: pd.DataFrame,
    metas: pd.DataFrame,
    receitas: pd.DataFrame,
    ledger: pd.DataFrame,
    expenses: pd.DataFrame,
    performance: pd.DataFrame,
    classifications: dict[str, str],
    start: pd.Period,
    end: pd.Period,
) -> pd.DataFrame:
    months = pd.period_range(start, end, freq="M")
    out = pd.DataFrame({"Mês": months})

    f = period_filter(fat, start, end).groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_MES": "Mês", "_VALOR": "Faturamento"})
    m = period_filter(metas, start, end).groupby("_MES", as_index=False)["_META"].sum().rename(columns={"_MES": "Mês", "_META": "Meta"})
    r = period_filter(receitas, start, end).groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_MES": "Mês", "_VALOR": "Recebimentos Totais"})

    l = period_filter(ledger, start, end).copy()
    e = period_filter(expenses, start, end).copy()
    e["Classificação"] = e["PAI"].map(classifications).fillna("Fixo/Operacional")

    op_rev = l[l["_GRUPO_N"] == "RECEITAS OPERACIONAIS"].groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_MES": "Mês", "_VALOR": "Receita Operacional"})
    non_op_rev = l[l["_GRUPO_N"] == "RECEITAS NAO OPERACIONAIS"].groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_MES": "Mês", "_VALOR": "Receita Não Operacional"})
    op_exp = e[e["_GRUPO_N"] == "SAIDAS OPERACIONAIS"].groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_MES": "Mês", "_VALOR": "Despesas Operacionais"})
    var = e[(e["_GRUPO_N"] == "SAIDAS OPERACIONAIS") & (e["Classificação"] == "Variável")].groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_MES": "Mês", "_VALOR": "Custos Variáveis"})
    fix = e[(e["_GRUPO_N"] == "SAIDAS OPERACIONAIS") & (e["Classificação"] == "Fixo/Operacional")].groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_MES": "Mês", "_VALOR": "Custos Fixos"})
    add = e[(e["_GRUPO_N"] == "SAIDAS OPERACIONAIS") & (e["Classificação"] == "Excluir EBITDA")].groupby("_MES", as_index=False)["_VALOR"].sum().rename(columns={"_MES": "Mês", "_VALOR": "IRPJ/CSLL add-back"})

    p = period_filter(performance, start, end).groupby("_MES", as_index=False)[["_PREVISTO", "_REALIZADO"]].sum().rename(columns={"_MES": "Mês", "_PREVISTO": "Recebimento Previsto", "_REALIZADO": "Recebimento Realizado"})

    for df in [f, m, r, op_rev, non_op_rev, op_exp, var, fix, add, p]:
        out = out.merge(df, on="Mês", how="left")
    out = out.fillna(0)
    out["Atingimento"] = np.where(out["Meta"] != 0, out["Faturamento"] / out["Meta"], 0)
    out["Margem de Contribuição"] = out["Faturamento"] - out["Custos Variáveis"]
    out["MC %"] = np.where(out["Faturamento"] != 0, out["Margem de Contribuição"] / out["Faturamento"], 0)
    out["EBITDA Gerencial"] = out["Receita Operacional"] - out["Despesas Operacionais"] + out["IRPJ/CSLL add-back"]
    out["Margem EBITDA"] = np.where(out["Receita Operacional"] != 0, out["EBITDA Gerencial"] / out["Receita Operacional"], 0)
    out["Performance Recebimento"] = np.where(out["Recebimento Previsto"] != 0, out["Recebimento Realizado"] / out["Recebimento Previsto"], 0)
    out["Entradas Classificadas"] = out["Receita Operacional"] + out["Receita Não Operacional"]
    out["Qualidade das Entradas"] = np.where(out["Entradas Classificadas"] != 0, out["Receita Operacional"] / out["Entradas Classificadas"], 0)
    out["Conversão do Faturamento"] = np.where(out["Faturamento"] != 0, out["Receita Operacional"] / out["Faturamento"], 0)
    out["Mês Texto"] = out["Mês"].map(month_label)
    return out


def executive_totals(monthly: pd.DataFrame, caixa: pd.DataFrame, start: pd.Period, end: pd.Period) -> dict[str, float]:
    sums = monthly.select_dtypes(include=[np.number]).sum().to_dict()
    caixa_p = period_filter(caixa, start, end)
    sums["Caixa Receitas"] = float(caixa_p["_RECEITAS"].sum())
    sums["Caixa Despesas"] = float(caixa_p["_DESPESAS"].sum())
    sums["Geração Operacional"] = sums["Caixa Receitas"] - sums["Caixa Despesas"]
    sums["Atingimento"] = safe_div(sums.get("Faturamento", 0), sums.get("Meta", 0))
    sums["MC %"] = safe_div(sums.get("Margem de Contribuição", 0), sums.get("Faturamento", 0))
    sums["Margem EBITDA"] = safe_div(sums.get("EBITDA Gerencial", 0), sums.get("Receita Operacional", 0))
    sums["Performance Recebimento"] = safe_div(sums.get("Recebimento Realizado", 0), sums.get("Recebimento Previsto", 0))
    sums["Entradas Classificadas"] = sums.get("Receita Operacional", 0) + sums.get("Receita Não Operacional", 0)
    sums["Qualidade das Entradas"] = safe_div(sums.get("Receita Operacional", 0), sums.get("Entradas Classificadas", 0))
    sums["Conversão do Faturamento"] = safe_div(sums.get("Receita Operacional", 0), sums.get("Faturamento", 0))
    sums["Dependência Não Operacional"] = safe_div(sums.get("Receita Não Operacional", 0), sums.get("Entradas Classificadas", 0))
    sums["Ponto de Equilíbrio"] = safe_div(sums.get("Custos Fixos", 0), sums.get("MC %", 0))
    return sums


# =========================================================
# FONTES
# =========================================================
default_base = first_existing(["BASE BI.xlsx", "BASE BI(1).xlsx", "base_bi.xlsx"])
default_rev = first_existing(["rev2026 Base bi.xlsx", "rev2026 Base bi(1).xlsx", "REV2026.xlsx"])

with st.sidebar:
    st.markdown(f"<h2 style='color:{NAVY};margin-bottom:0'>📊 KPIs Executivos</h2>", unsafe_allow_html=True)
    st.caption("Painel gerencial integrado")
    with st.expander("📁 Fontes de dados", expanded=False):
        up_base = st.file_uploader("Substituir BASE BI", type=["xlsx", "xlsm"], key="up_base")
        up_rev = st.file_uploader("Substituir REV2026", type=["xlsx", "xlsm"], key="up_rev")
        st.caption("Sem upload, o app utiliza os arquivos presentes no repositório.")

base_bytes, base_name = source_bytes(up_base, default_base)
rev_bytes, rev_name = source_bytes(up_rev, default_rev)

if not base_bytes or not rev_bytes:
    st.error("Não localizei as duas bases. Inclua `BASE BI.xlsx` e `rev2026 Base bi.xlsx` no repositório ou carregue os arquivos na barra lateral.")
    st.stop()

try:
    with st.spinner("Lendo e validando as bases..."):
        base = load_base_bi(base_bytes)
        rev = load_rev(rev_bytes)
except Exception as exc:
    st.error(f"Não foi possível carregar as bases: {exc}")
    st.stop()

fat = base["faturamento"]
metas = base["metas"]
metas_g = base["metas_gerentes"]
receitas = rev["receitas"]
expenses = rev["despesas"]
costs = rev["custos"]
caixa = rev["caixa"]
performance = rev["recebimento"]
init_classifications(expenses)

all_periods = pd.concat([
    fat["_MES"], metas["_MES"], receitas["_MES"], expenses["_MES"], costs["_MES"], performance["_MES"]
]).dropna()
min_month, max_month = all_periods.min(), all_periods.max()
month_options = list(pd.period_range(min_month, max_month, freq="M"))

with st.sidebar:
    st.divider()
    st.markdown("#### Período executivo")
    start_month = st.selectbox("Mês inicial", month_options, index=0, format_func=month_label)
    end_month = st.selectbox("Mês final", month_options, index=len(month_options) - 1, format_func=month_label)
    if start_month > end_month:
        st.warning("O mês inicial foi ajustado para não ficar após o mês final.")
        start_month = end_month
    st.divider()
    st.markdown("#### Bases ativas")
    st.success(f"BASE BI: {base_name}")
    st.success(f"REV2026: {rev_name}")
    st.caption("Na REV2026, apenas abas visíveis são lidas. Abas ocultas são ignoradas automaticamente.")

monthly = calculate_monthly(
    fat, metas, receitas, costs, expenses, performance,
    st.session_state["cost_classifications"], start_month, end_month,
)
totals = executive_totals(monthly, caixa, start_month, end_month)

period_text = f"{month_label(start_month)} a {month_label(end_month)}"
st.markdown(
    f"""
    <div class="hero">
      <h1>Painel Executivo de KPIs</h1>
      <p>{period_text} · resultado comercial, margem, caixa, recebimentos e estrutura de custos</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# KPIs principais — quatro cartões por linha para preservar leitura e evitar sobreposição.
k1, k2, k3, k4 = st.columns(4)
with k1: card("Faturamento bruto", brl(totals["Faturamento"]), f"Vendas emitidas · Meta: {brl(totals['Meta'])}")
with k2: card("Atingimento da meta", pct(totals["Atingimento"]), "Faturamento bruto ÷ meta comercial")
with k3: card("Margem de contribuição", pct(totals["MC %"]), f"Valor: {brl(totals['Margem de Contribuição'])}")
with k4: card("EBITDA gerencial de caixa", brl(totals["EBITDA Gerencial"]), "Operação antes de IRPJ e CSLL; visão gerencial")

k5, k6, k7, k8 = st.columns(4)
with k5: card("Margem EBITDA", pct(totals["Margem EBITDA"]), "EBITDA gerencial ÷ receita operacional recebida")
with k6: card("Receita operacional recebida", brl(totals["Receita Operacional"]), "Entradas ligadas à atividade principal")
with k7: card("Entradas não operacionais", brl(totals["Receita Não Operacional"]), "Pode incluir capital de giro, empréstimos e outras fontes")
with k8: card("Qualidade das entradas", pct(totals["Qualidade das Entradas"]), "Receita operacional ÷ entradas classificadas")

k9, k10, k11, k12 = st.columns(4)
with k9: card("Recebimentos totais", brl(totals["Recebimentos Totais"]), "Total pago na base de recebimentos")
with k10: card("Performance de recebimento", pct(totals["Performance Recebimento"]), "Recebimento realizado ÷ previsto")
with k11: card("Despesas operacionais", brl(totals["Despesas Operacionais"]), f"{pct(safe_div(totals['Despesas Operacionais'], totals['Receita Operacional']))} da receita operacional")
with k12: card("Ponto de equilíbrio", brl(totals["Ponto de Equilíbrio"]), "Custos fixos ÷ margem de contribuição %")

with st.expander("ℹ️ Como interpretar cada KPI", expanded=False):
    h1, h2 = st.columns(2)
    with h1:
        st.markdown(
            """
            <div class='kpi-help'><b>Faturamento bruto</b><br>Valor das vendas emitidas no período. Não significa, necessariamente, que todo o valor já foi recebido.</div>
            <div class='kpi-help'><b>Atingimento da meta</b><br>Compara o faturamento bruto com a meta comercial acumulada do mesmo período.</div>
            <div class='kpi-help'><b>Margem de contribuição</b><br>Faturamento menos custos variáveis. Mostra quanto sobra para pagar a estrutura fixa e formar resultado.</div>
            <div class='kpi-help'><b>EBITDA gerencial de caixa</b><br>Receita operacional recebida menos despesas operacionais, antes de IRPJ e CSLL. É uma visão gerencial, não o EBITDA contábil oficial.</div>
            <div class='kpi-help'><b>Margem EBITDA</b><br>Percentual da receita operacional recebida que permaneceu como EBITDA gerencial.</div>
            <div class='kpi-help'><b>Qualidade das entradas</b><br>Percentual das entradas classificadas que veio da operação. Quanto menor, maior a participação de fontes não operacionais.</div>
            """,
            unsafe_allow_html=True,
        )
    with h2:
        st.markdown(
            """
            <div class='kpi-help'><b>Receita operacional recebida</b><br>Entradas de caixa relacionadas à atividade principal da empresa.</div>
            <div class='kpi-help'><b>Entradas não operacionais</b><br>Valores que aumentam o caixa, mas não representam receita da operação, como capital de giro, empréstimos e outros recebimentos.</div>
            <div class='kpi-help'><b>Recebimentos totais</b><br>Total de títulos pagos na base de recebimentos, sem assumir que todas as entradas representam receita operacional do período.</div>
            <div class='kpi-help'><b>Performance de recebimento</b><br>Compara o que efetivamente entrou com o valor previsto para recebimento.</div>
            <div class='kpi-help'><b>Despesas operacionais</b><br>Saídas associadas à manutenção e execução das atividades da empresa no período.</div>
            <div class='kpi-help'><b>Ponto de equilíbrio</b><br>Faturamento necessário para cobrir os custos fixos, considerando a margem de contribuição estimada.</div>
            """,
            unsafe_allow_html=True,
        )

st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

tab_exec, tab_result, tab_com, tab_cash, tab_cost, tab_quality = st.tabs([
    "Visão Executiva", "Resultado Mensal", "Comercial", "Caixa & Recebimentos", "Custos & Premissas", "Qualidade da Base"
])

# =========================================================
# VISÃO EXECUTIVA
# =========================================================
with tab_exec:
    c1, c2 = st.columns([1.45, 1])
    with c1:
        fig = go.Figure()
        fig.add_bar(
            x=monthly["Mês Texto"], y=monthly["Faturamento"], name="Faturamento", marker_color=BLUE,
            text=monthly["Faturamento"].map(compact_currency_label), textposition="outside", cliponaxis=False
        )
        fig.add_scatter(
            x=monthly["Mês Texto"], y=monthly["Meta"], name="Meta", mode="lines+markers+text",
            line=dict(color=ORANGE, width=3), text=monthly["Meta"].map(compact_currency_label),
            textposition="top center"
        )
        fig.update_layout(title="Faturamento x Meta")
        apply_brl_axis(fig, "y")
        hide_value_axis(fig, "y")
        st.plotly_chart(plot_layout(fig), width="stretch", config={"displayModeBar": False})
    with c2:
        waterfall_values = [
            totals["Receita Operacional"], -totals["Custos Variáveis"], -totals["Custos Fixos"],
            totals["IRPJ/CSLL add-back"], totals["EBITDA Gerencial"]
        ]
        waterfall = go.Figure(go.Waterfall(
            name="Resultado", orientation="v", measure=["absolute", "relative", "relative", "relative", "total"],
            x=["Receita<br>operacional", "Custos<br>variáveis", "Custos<br>fixos", "IRPJ/CSLL<br>add-back", "EBITDA"],
            y=waterfall_values,
            connector={"line": {"color": "#98A2B3"}},
            increasing={"marker": {"color": GREEN}}, decreasing={"marker": {"color": RED}}, totals={"marker": {"color": NAVY}},
            text=[compact_currency_label(v) for v in waterfall_values], textposition="outside", cliponaxis=False,
        ))
        waterfall.update_layout(title="Ponte do EBITDA Gerencial")
        apply_brl_axis(waterfall, "y")
        hide_value_axis(waterfall, "y")
        st.plotly_chart(plot_layout(waterfall, 420, legend_bottom=False), width="stretch", config={"displayModeBar": False})

    c3, c4 = st.columns(2)
    with c3:
        fig = go.Figure()
        fig.add_scatter(
            x=monthly["Mês Texto"], y=monthly["MC %"], name="Margem de contribuição", mode="lines+markers+text",
            line=dict(color=TEAL, width=3), text=monthly["MC %"].map(compact_percent_label), textposition="top center"
        )
        fig.add_scatter(
            x=monthly["Mês Texto"], y=monthly["Margem EBITDA"], name="Margem EBITDA", mode="lines+markers+text",
            line=dict(color=NAVY, width=3), text=monthly["Margem EBITDA"].map(compact_percent_label), textposition="bottom center"
        )
        fig.update_yaxes(tickformat=".0%")
        fig.update_layout(title="Evolução das Margens")
        hide_value_axis(fig, "y")
        st.plotly_chart(plot_layout(fig), width="stretch", config={"displayModeBar": False})
    with c4:
        exec_costs = period_filter(expenses, start_month, end_month)
        exec_costs = exec_costs[exec_costs["_GRUPO_N"] == "SAIDAS OPERACIONAIS"]
        top_cost = exec_costs.groupby("PAI", as_index=False)["_VALOR"].sum().sort_values("_VALOR", ascending=False).head(8)
        top_cost["Natureza"] = top_cost["PAI"].map(lambda x: short_label(x, 34))
        fig = px.bar(top_cost.sort_values("_VALOR"), x="_VALOR", y="Natureza", orientation="h", title="Principais Despesas Operacionais", custom_data=["PAI"])
        fig.update_traces(marker_color=BLUE, hovertemplate="%{customdata[0]}<br>Valor: R$ %{x:,.2f}<extra></extra>", text=top_cost.sort_values("_VALOR")["_VALOR"].map(compact_currency_label), textposition="outside", cliponaxis=False)
        apply_brl_axis(fig, "x")
        hide_value_axis(fig, "x")
        st.plotly_chart(plot_layout(fig, legend_bottom=False), width="stretch", config={"displayModeBar": False})

    st.markdown("<div class='section-title'>Leitura executiva automática</div>", unsafe_allow_html=True)
    fat_month = monthly.loc[monthly["Faturamento"].idxmax()]
    weak_ebitda = monthly[monthly["EBITDA Gerencial"] < 0]
    top_pai = top_cost.iloc[0] if not top_cost.empty else None
    top_manager = period_filter(fat, start_month, end_month).groupby("GERENTE", as_index=False)["_VALOR"].sum().sort_values("_VALOR", ascending=False)
    top_manager_row = top_manager.iloc[0] if not top_manager.empty else None
    insights = [
        f"O faturamento acumulado atingiu <b>{pct(totals['Atingimento'])}</b> da meta no período.",
        f"O melhor mês de faturamento foi <b>{fat_month['Mês Texto']}</b>, com <b>{brl(fat_month['Faturamento'])}</b>.",
        f"A qualidade das entradas foi de <b>{pct(totals['Qualidade das Entradas'])}</b>; as entradas não operacionais somaram <b>{brl(totals['Receita Não Operacional'])}</b> e não foram tratadas como receita do negócio.",
    ]
    if top_manager_row is not None:
        insights.append(f"O gerente com maior faturamento foi <b>{top_manager_row['GERENTE']}</b>, com <b>{brl(top_manager_row['_VALOR'])}</b>.")
    if top_pai is not None:
        insights.append(f"A maior natureza de despesa operacional foi <b>{top_pai['PAI']}</b>, totalizando <b>{brl(top_pai['_VALOR'])}</b>.")
    if weak_ebitda.empty:
        insights.append("Todos os meses do período apresentaram EBITDA gerencial de caixa positivo.")
    else:
        months_negative = ", ".join(weak_ebitda["Mês Texto"].tolist())
        insights.append(f"Houve EBITDA gerencial negativo em <b>{months_negative}</b>; vale revisar o descasamento entre recebimentos e pagamentos desses meses.")
    for item in insights:
        st.markdown(f"<div class='insight'>{item}</div>", unsafe_allow_html=True)

# =========================================================
# RESULTADO MENSAL
# =========================================================
with tab_result:
    st.markdown("<div class='section-title'>Demonstração gerencial mensal</div>", unsafe_allow_html=True)
    display_cols = [
        "Mês Texto", "Faturamento", "Meta", "Atingimento", "Receita Operacional", "Receita Não Operacional",
        "Qualidade das Entradas", "Conversão do Faturamento", "Despesas Operacionais", "Custos Variáveis", "Custos Fixos",
        "Margem de Contribuição", "MC %", "EBITDA Gerencial", "Margem EBITDA", "Recebimento Previsto",
        "Recebimento Realizado", "Performance Recebimento",
    ]
    result_display = monthly[display_cols].rename(columns={"Mês Texto": "Mês"}).copy()
    money_cols_result = [
        "Faturamento", "Meta", "Receita Operacional", "Receita Não Operacional", "Despesas Operacionais", "Custos Variáveis",
        "Custos Fixos", "Margem de Contribuição", "EBITDA Gerencial", "Recebimento Previsto", "Recebimento Realizado"
    ]
    pct_cols_result = ["Atingimento", "Qualidade das Entradas", "Conversão do Faturamento", "MC %", "Margem EBITDA", "Performance Recebimento"]
    result_for_view = result_display.copy()
    for c in money_cols_result:
        result_for_view[c] = result_for_view[c].map(brl)
    for c in pct_cols_result:
        result_for_view[c] = result_for_view[c].map(pct)
    st.dataframe(
        result_for_view,
        width="stretch", hide_index=True, height=330,
    )
    export_df = result_display.copy()
    st.download_button(
        "⬇️ Exportar resultado mensal",
        data=dataframe_download(export_df, "Resultado Mensal"),
        file_name=f"resultado_gerencial_{start_month}_{end_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.markdown(
        "<div class='method-note'><b>Importante:</b> margem de contribuição e EBITDA usam bases com regimes diferentes. "
        "O faturamento está por competência comercial; custos e recebimentos refletem caixa. Entradas não operacionais, como capital de giro, aumentam o caixa, mas não entram como receita operacional nem melhoram a qualidade das entradas. O indicador é gerencial e deve ser conciliado ao balancete para uso contábil oficial.</div>",
        unsafe_allow_html=True,
    )

# =========================================================
# COMERCIAL
# =========================================================
with tab_com:
    st.markdown("<div class='section-title'>Filtros comerciais</div>", unsafe_allow_html=True)
    base_com = period_filter(fat, start_month, end_month)
    f1, f2, f3, f4 = st.columns(4)
    managers = f1.multiselect("Gerente", sorted(base_com["GERENTE"].dropna().unique())) if "GERENTE" in base_com else []
    seller_col = "VENDEDOR / REPRESENTANTE" if "VENDEDOR / REPRESENTANTE" in base_com else "VENDEDOR"
    sellers = f2.multiselect("Vendedor / Representante", sorted(base_com[seller_col].dropna().unique())) if seller_col in base_com else []
    segments = f3.multiselect("Segmento", sorted(base_com["SEGMENTO"].dropna().unique())) if "SEGMENTO" in base_com else []
    types = f4.multiselect("Tipo", sorted(base_com["NOVA"].dropna().unique())) if "NOVA" in base_com else []
    f5, f6 = st.columns(2)
    companies = f5.multiselect("Empresa", sorted(base_com["EMPRESA"].dropna().unique())) if "EMPRESA" in base_com else []
    search_client = f6.text_input("Buscar cliente", placeholder="Digite parte do nome")

    com = base_com.copy()
    if managers: com = com[com["GERENTE"].isin(managers)]
    if sellers: com = com[com[seller_col].isin(sellers)]
    if segments: com = com[com["SEGMENTO"].isin(segments)]
    if types: com = com[com["NOVA"].isin(types)]
    if companies: com = com[com["EMPRESA"].isin(companies)]
    if search_client and "NOME DO CLIENTE" in com:
        com = com[com["NOME DO CLIENTE"].str.contains(search_client, case=False, na=False)]

    filtered_revenue = float(com["_VALOR"].sum())
    meta_filtered = filtered_meta(metas, metas_g, start_month, end_month, managers, sellers)
    invoices = com["Nota Fiscal"].nunique() if "Nota Fiscal" in com else len(com)
    clients = com["NOME DO CLIENTE"].nunique() if "NOME DO CLIENTE" in com else 0
    ticket = safe_div(filtered_revenue, invoices)
    share = safe_div(filtered_revenue, totals["Faturamento"])

    m1, m2, m3 = st.columns(3)
    with m1: card("Faturamento filtrado", brl(filtered_revenue), f"{pct(share)} do total")
    with m2: card("Meta filtrada", brl(meta_filtered), "Conforme vendedor ou gerente selecionado")
    with m3: card("Atingimento", pct(safe_div(filtered_revenue, meta_filtered)), "Faturamento ÷ meta")
    m4, m5 = st.columns(2)
    with m4: card("Clientes ativos", f"{clients:,}".replace(",", "."), f"{invoices:,} notas".replace(",", "."))
    with m5: card("Ticket médio por NF", brl(ticket), "Faturamento ÷ notas fiscais")

    c1, c2 = st.columns(2)
    with c1:
        by_month = com.groupby("_MES", as_index=False)["_VALOR"].sum()
        by_month["Mês"] = by_month["_MES"].map(month_label)
        fig = px.bar(by_month, x="Mês", y="_VALOR", title="Faturamento Mensal Filtrado")
        fig.update_traces(marker_color=BLUE, hovertemplate="%{x}<br>Valor: R$ %{y:,.2f}<extra></extra>", text=by_month["_VALOR"].map(compact_currency_label), textposition="outside", cliponaxis=False)
        apply_brl_axis(fig, "y")
        hide_value_axis(fig, "y")
        st.plotly_chart(plot_layout(fig), width="stretch", config={"displayModeBar": False})
    with c2:
        by_manager = com.groupby("GERENTE", as_index=False)["_VALOR"].sum().sort_values("_VALOR", ascending=False)
        by_manager["Gerente"] = by_manager["GERENTE"].map(lambda x: short_label(x, 30))
        by_manager_plot = by_manager.sort_values("_VALOR")
        fig = px.bar(by_manager_plot, x="_VALOR", y="Gerente", orientation="h", title="Faturamento por Gerente", custom_data=["GERENTE"])
        fig.update_traces(marker_color=TEAL, hovertemplate="%{customdata[0]}<br>Valor: R$ %{x:,.2f}<extra></extra>", text=by_manager_plot["_VALOR"].map(compact_currency_label), textposition="outside", cliponaxis=False)
        apply_brl_axis(fig, "x")
        hide_value_axis(fig, "x")
        st.plotly_chart(plot_layout(fig, legend_bottom=False), width="stretch", config={"displayModeBar": False})

    c3, c4 = st.columns(2)
    with c3:
        by_segment = com.groupby("SEGMENTO", as_index=False)["_VALOR"].sum().sort_values("_VALOR", ascending=False).head(12)
        by_segment_plot = by_segment.sort_values("_VALOR")
        fig = px.bar(by_segment_plot, x="_VALOR", y="SEGMENTO", orientation="h", title="Principais Segmentos")
        fig.update_traces(marker_color=BLUE, hovertemplate="%{y}<br>Valor: R$ %{x:,.2f}<extra></extra>", text=by_segment_plot["_VALOR"].map(compact_currency_label), textposition="outside", cliponaxis=False)
        apply_brl_axis(fig, "x")
        hide_value_axis(fig, "x")
        st.plotly_chart(plot_layout(fig, 430, legend_bottom=False), width="stretch", config={"displayModeBar": False})
    with c4:
        type_col = "NOVA" if "NOVA" in com.columns else "CATEGORIA"
        by_type = com.groupby(type_col, as_index=False)["_VALOR"].sum().sort_values("_VALOR", ascending=False).head(10)
        by_type["Tipo"] = by_type[type_col].map(lambda x: short_label(x, 32))
        by_type_plot = by_type.sort_values("_VALOR")
        fig = px.bar(by_type_plot, x="_VALOR", y="Tipo", orientation="h", title="Mix de Receita", custom_data=[type_col])
        fig.update_traces(marker_color=TEAL, hovertemplate="%{customdata[0]}<br>Valor: R$ %{x:,.2f}<extra></extra>", text=by_type_plot["_VALOR"].map(compact_currency_label), textposition="outside", cliponaxis=False)
        apply_brl_axis(fig, "x")
        hide_value_axis(fig, "x")
        st.plotly_chart(plot_layout(fig, 430, legend_bottom=False), width="stretch", config={"displayModeBar": False})

    c5, c6 = st.columns(2)
    with c5:
        by_seller = com.groupby(seller_col, as_index=False)["_VALOR"].sum().sort_values("_VALOR", ascending=False).head(15)
        st.markdown("**Ranking de vendedores / representantes**")
        seller_view = by_seller.rename(columns={seller_col: "Vendedor / Representante", "_VALOR": "Faturamento"})
        seller_view["Faturamento"] = seller_view["Faturamento"].map(brl)
        st.dataframe(seller_view, width="stretch", hide_index=True, height=440)
    with c6:
        client_col = "NOME DO CLIENTE" if "NOME DO CLIENTE" in com else "CLIENTE"
        by_client = com.groupby(client_col, as_index=False)["_VALOR"].sum().sort_values("_VALOR", ascending=False).head(15)
        st.markdown("**Principais clientes**")
        client_view = by_client.rename(columns={client_col: "Cliente", "_VALOR": "Faturamento"})
        client_view["Faturamento"] = client_view["Faturamento"].map(brl)
        st.dataframe(client_view, width="stretch", hide_index=True, height=440)

    export_cols = [c for c in ["DT Emissao", "Nota Fiscal", "NOME DO CLIENTE", "GERENTE", seller_col, "SEGMENTO", "NOVA", "LINHA DE PRODUTO", "VALOR BRUTO"] if c in com.columns]
    export_com = com[export_cols].copy() if export_cols else com.copy()
    st.download_button(
        "⬇️ Exportar base comercial filtrada",
        data=export_com.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig").encode("utf-8-sig"),
        file_name=f"comercial_filtrado_{start_month}_{end_month}.csv",
        mime="text/csv",
    )

# =========================================================
# CAIXA & RECEBIMENTOS
# =========================================================
with tab_cash:
    c1, c2 = st.columns(2)
    with c1:
        fig = go.Figure()
        fig.add_bar(
            x=monthly["Mês Texto"], y=monthly["Recebimento Previsto"], name="Previsto", marker_color="#A7C4DF",
            text=monthly["Recebimento Previsto"].map(compact_currency_label), textposition="outside", cliponaxis=False
        )
        fig.add_scatter(
            x=monthly["Mês Texto"], y=monthly["Recebimento Realizado"], name="Realizado", mode="lines+markers+text",
            line=dict(color=TEAL, width=3), text=monthly["Recebimento Realizado"].map(compact_currency_label), textposition="top center"
        )
        fig.update_layout(title="Recebimento Previsto x Realizado")
        apply_brl_axis(fig, "y")
        hide_value_axis(fig, "y")
        st.plotly_chart(plot_layout(fig), width="stretch", config={"displayModeBar": False})
    with c2:
        fig = go.Figure()
        fig.add_bar(
            x=monthly["Mês Texto"], y=monthly["Receita Operacional"], name="Receitas operacionais", marker_color=BLUE,
            text=monthly["Receita Operacional"].map(compact_currency_label), textposition="outside", cliponaxis=False
        )
        fig.add_bar(
            x=monthly["Mês Texto"], y=monthly["Despesas Operacionais"], name="Despesas operacionais", marker_color=RED,
            text=monthly["Despesas Operacionais"].map(compact_currency_label), textposition="outside", cliponaxis=False
        )
        fig.add_scatter(
            x=monthly["Mês Texto"], y=monthly["EBITDA Gerencial"], name="EBITDA", mode="lines+markers+text",
            line=dict(color=GREEN, width=3), text=monthly["EBITDA Gerencial"].map(compact_currency_label), textposition="top center"
        )
        fig.update_layout(title="Receita, Despesa e EBITDA Gerencial", barmode="group")
        apply_brl_axis(fig, "y")
        hide_value_axis(fig, "y")
        st.plotly_chart(plot_layout(fig), width="stretch", config={"displayModeBar": False})

    c3, c4, c5, c6 = st.columns(4)
    with c3: card("Receita operacional", brl(totals["Receita Operacional"]), "Entradas vinculadas à atividade principal")
    with c4: card("Conversão do faturamento", pct(totals["Conversão do Faturamento"]), "Receita operacional recebida ÷ faturamento emitido")
    with c5: card("Qualidade das entradas", pct(totals["Qualidade das Entradas"]), "Parcela operacional das entradas classificadas")
    with c6: card("Entradas não operacionais", brl(totals["Receita Não Operacional"]), f"{pct(totals['Dependência Não Operacional'])} das entradas classificadas")

    st.markdown(
        "<div class='method-note'><b>Leitura correta:</b> capital de giro e empréstimos aumentam o saldo de caixa, "
        "mas não são receita gerada pela operação. Por isso, o app separa essas entradas e calcula a qualidade das entradas "
        "somente pela participação da receita operacional.</div>",
        unsafe_allow_html=True,
    )

    rev_period = period_filter(costs, start_month, end_month)
    comp = rev_period.groupby("_GRUPO_N", as_index=False)["_VALOR"].sum()
    label_map = {
        "RECEITAS OPERACIONAIS": "Receitas Operacionais", "RECEITAS NAO OPERACIONAIS": "Receitas Não Operacionais",
        "SAIDAS OPERACIONAIS": "Saídas Operacionais", "SAIDAS NAO OPERACIONAIS": "Saídas Não Operacionais",
    }
    comp["Grupo"] = comp["_GRUPO_N"].map(label_map).fillna(comp["_GRUPO_N"].str.title())
    comp = comp[comp["_GRUPO_N"].isin(label_map)]
    fig = px.bar(comp, x="Grupo", y="_VALOR", color="Grupo", title="Composição das Entradas e Saídas")
    fig.update_traces(text=comp["_VALOR"].map(compact_currency_label), textposition="outside", cliponaxis=False)
    fig.update_layout(showlegend=False)
    apply_brl_axis(fig, "y")
    hide_value_axis(fig, "y")
    st.plotly_chart(plot_layout(fig, legend_bottom=False), width="stretch", config={"displayModeBar": False})

    non_op = rev_period[rev_period["_GRUPO_N"] == "RECEITAS NAO OPERACIONAIS"].copy()
    if not non_op.empty:
        non_op_summary = non_op.groupby("PAI", as_index=False)["_VALOR"].sum().sort_values("_VALOR", ascending=False)
        non_op_summary["Natureza"] = non_op_summary["PAI"].map(lambda x: short_label(x, 38))
        non_op_plot = non_op_summary.sort_values("_VALOR")
        fig = px.bar(
            non_op_plot, x="_VALOR", y="Natureza", orientation="h",
            title="Detalhamento das Entradas Não Operacionais",
            custom_data=["PAI"],
        )
        fig.update_traces(marker_color=ORANGE, hovertemplate="%{customdata[0]}<br>Valor: R$ %{x:,.2f}<extra></extra>", text=non_op_plot["_VALOR"].map(compact_currency_label), textposition="outside", cliponaxis=False)
        apply_brl_axis(fig, "x")
        hide_value_axis(fig, "x")
        st.plotly_chart(plot_layout(fig, 350, legend_bottom=False), width="stretch", config={"displayModeBar": False})

# =========================================================
# CUSTOS & PREMISSAS
# =========================================================
with tab_cost:
    st.markdown("<div class='section-title'>Análise de custos operacionais</div>", unsafe_allow_html=True)
    cost_period = period_filter(costs, start_month, end_month)
    cost_period = cost_period[cost_period["_GRUPO_N"] == "SAIDAS OPERACIONAIS"].copy()
    cc1, cc2 = st.columns(2)
    centers = cc1.multiselect("Centro de custos", sorted(cost_period["CENTRO DE CUSTOS"].dropna().unique())) if "CENTRO DE CUSTOS" in cost_period else []
    rateios = cc2.multiselect("Rateio", sorted(cost_period["CENTRO DE CUSTOS RATEAO"].dropna().unique())) if "CENTRO DE CUSTOS RATEAO" in cost_period else []
    if centers: cost_period = cost_period[cost_period["CENTRO DE CUSTOS"].isin(centers)]
    if rateios: cost_period = cost_period[cost_period["CENTRO DE CUSTOS RATEAO"].isin(rateios)]

    c1, c2 = st.columns(2)
    with c1:
        pareto = cost_period.groupby("PAI", as_index=False)["_VALOR"].sum().sort_values("_VALOR", ascending=False).head(15)
        pareto_plot = pareto.sort_values("_VALOR")
        fig = px.bar(pareto_plot, x="_VALOR", y="PAI", orientation="h", title="Pareto de Despesas")
        fig.update_traces(marker_color=BLUE, hovertemplate="%{y}<br>Valor: R$ %{x:,.2f}<extra></extra>", text=pareto_plot["_VALOR"].map(compact_currency_label), textposition="outside", cliponaxis=False)
        apply_brl_axis(fig, "x")
        hide_value_axis(fig, "x")
        st.plotly_chart(plot_layout(fig, 500, legend_bottom=False), width="stretch", config={"displayModeBar": False})
    with c2:
        by_cc = cost_period.groupby("CENTRO DE CUSTOS RATEAO", as_index=False)["_VALOR"].sum().sort_values("_VALOR", ascending=False).head(12)
        by_cc["Rateio"] = by_cc["CENTRO DE CUSTOS RATEAO"].map(lambda x: short_label(x, 34))
        by_cc_plot = by_cc.sort_values("_VALOR")
        fig = px.bar(by_cc_plot, x="_VALOR", y="Rateio", orientation="h", title="Despesas por Rateio", custom_data=["CENTRO DE CUSTOS RATEAO"])
        fig.update_traces(marker_color=TEAL, hovertemplate="%{customdata[0]}<br>Valor: R$ %{x:,.2f}<extra></extra>", text=by_cc_plot["_VALOR"].map(compact_currency_label), textposition="outside", cliponaxis=False)
        apply_brl_axis(fig, "x")
        hide_value_axis(fig, "x")
        st.plotly_chart(plot_layout(fig, 500, legend_bottom=False), width="stretch", config={"displayModeBar": False})

    st.divider()
    st.markdown("#### Premissas da margem de contribuição e EBITDA")
    st.caption("Altere a classificação e clique em Aplicar. Os KPIs e gráficos serão recalculados em todo o app.")
    cost_totals = period_filter(expenses, start_month, end_month)
    cost_totals = cost_totals[cost_totals["_GRUPO_N"] == "SAIDAS OPERACIONAIS"].groupby("PAI", as_index=False)["_VALOR"].sum()
    class_df = pd.DataFrame([
        {"Natureza / PAI": p, "Classificação KPI": cls}
        for p, cls in st.session_state["cost_classifications"].items()
    ]).merge(cost_totals.rename(columns={"PAI": "Natureza / PAI", "_VALOR": "Valor no período"}), on="Natureza / PAI", how="left")
    class_df["Valor no período"] = class_df["Valor no período"].fillna(0)
    class_df = class_df.sort_values("Valor no período", ascending=False)
    class_df_editor = class_df.copy()
    class_df_editor["Valor no período"] = class_df_editor["Valor no período"].map(brl)
    edited = st.data_editor(
        class_df_editor,
        width="stretch",
        hide_index=True,
        height=480,
        column_config={
            "Natureza / PAI": st.column_config.TextColumn(disabled=True),
            "Classificação KPI": st.column_config.SelectboxColumn(options=["Variável", "Fixo/Operacional", "Excluir EBITDA"], required=True),
            "Valor no período": st.column_config.TextColumn(disabled=True),
        },
        key="classification_editor",
    )
    b1, b2, b3 = st.columns([1, 1, 3])
    if b1.button("Aplicar premissas", type="primary", width="stretch"):
        st.session_state["cost_classifications"] = dict(zip(edited["Natureza / PAI"], edited["Classificação KPI"]))
        st.rerun()
    if b2.button("Restaurar padrão", width="stretch"):
        st.session_state["cost_classifications"] = {p: default_classification(p) for p in edited["Natureza / PAI"]}
        st.rerun()
    export_premissas = class_df[["Natureza / PAI", "Valor no período"]].copy()
    export_premissas["Classificação KPI"] = export_premissas["Natureza / PAI"].map(
        dict(zip(edited["Natureza / PAI"], edited["Classificação KPI"]))
    )
    export_premissas = export_premissas[["Natureza / PAI", "Classificação KPI", "Valor no período"]]
    b3.download_button(
        "⬇️ Exportar premissas",
        data=dataframe_download(export_premissas, "Premissas Custos"),
        file_name="premissas_classificacao_custos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="content",
    )

# =========================================================
# QUALIDADE E METODOLOGIA
# =========================================================
with tab_quality:
    base_states = base["sheet_states"]
    rev_states = rev["sheet_states"]
    rev_visible = [s for s, state in rev_states.items() if state == "visible"]
    rev_hidden = [s for s, state in rev_states.items() if state != "visible"]

    q1, q2, q3, q4 = st.columns(4)
    duplicate_cols = [c for c in ["Nota Fiscal", "PRODUTO", "NOME DO CLIENTE", "_VALOR", "_MES"] if c in fat.columns]
    dup_rows = int(fat.duplicated(subset=duplicate_cols, keep=False).sum()) if duplicate_cols else 0
    missing_date = int(fat["_MES"].isna().sum())
    missing_value = int((fat["_VALOR"] == 0).sum())
    with q1: card("Linhas faturamento", f"{len(fat):,}".replace(",", "."), "BANCO DE DADOS FATURAMENTO")
    with q2: card("Linhas de despesas", f"{len(expenses):,}".replace(",", "."), "Resumo Despesas tratado")
    with q3: card("Possíveis duplicidades", f"{dup_rows:,}".replace(",", "."), "Não são removidas automaticamente")
    with q4: card("Registros sem mês / valor", f"{missing_date + missing_value:,}".replace(",", "."), "Itens para revisão da fonte")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### Abas REV2026 utilizadas")
        st.success("Somente abas visíveis entram no processamento.")
        st.dataframe(pd.DataFrame({"Aba visível": rev_visible}), width="stretch", hide_index=True)
    with c2:
        st.markdown("#### Abas REV2026 ignoradas")
        st.warning("As abas abaixo estavam ocultas e foram desprezadas.")
        st.dataframe(pd.DataFrame({"Aba oculta": rev_hidden}), width="stretch", hide_index=True)

    st.markdown("#### Metodologia")
    methodology = pd.DataFrame([
        ["Faturamento Bruto", "Soma de VALOR BRUTO da aba BANCO DE DADOS FATURAMENTO.", "Competência comercial"],
        ["Meta", "Soma de META MENSAL da aba Metas; quando há filtro de gerente, usa Metas Gerentes.", "Competência"],
        ["Margem de Contribuição", "Faturamento menos saídas operacionais da aba Resumo Despesas classificadas como variáveis.", "Estimativa gerencial mista"],
        ["EBITDA Gerencial Caixa", "Receitas operacionais do Centro de Custos menos saídas operacionais do Resumo Despesas, adicionando IRPJ e CSLL de volta.", "Caixa"],
        ["Performance de Recebimento", "Recebimento realizado dividido pelo previsto.", "Caixa"],
        ["Qualidade das Entradas", "Receita operacional recebida dividida pelas entradas operacionais e não operacionais classificadas.", "Caixa"],
        ["Conversão do Faturamento", "Receita operacional recebida dividida pelo faturamento bruto emitido.", "Regimes diferentes; observar defasagem"],
        ["Entradas Não Operacionais", "Entradas que aumentam o caixa, mas não representam receita da operação, como capital de giro e empréstimos.", "Caixa / financiamento"],
        ["Ponto de Equilíbrio", "Custos fixos divididos pela margem de contribuição percentual.", "Estimativa gerencial"],
    ], columns=["Indicador", "Cálculo", "Regime"])
    st.dataframe(methodology, width="stretch", hide_index=True)
    st.markdown(
        "<div class='method-note'><b>Próxima evolução recomendada:</b> integrar balancete ou DRE por competência, "
        "CMV/estoques, depreciação, amortização e despesas financeiras. Com isso, o app poderá exibir EBITDA contábil, "
        "lucro operacional, lucro líquido e margem por produto/cliente com maior precisão.</div>",
        unsafe_allow_html=True,
    )

st.caption("Painel Executivo de KPIs · Desenvolvido para uso gerencial da First Medical · Dados processados localmente no app")
