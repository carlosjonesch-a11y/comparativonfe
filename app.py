import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import plotly.express as px

st.set_page_config(page_title="Comparativo de Faturamento", layout="wide")

# ──────────────────────────────────────────────────────────────────────────────
# MAPEAMENTO CFOP → NATUREZA
# ──────────────────────────────────────────────────────────────────────────────
CFOP_MAP = {
    # Vendas
    6108: "Venda",
    5102: "Venda",
    6102: "Venda",
    5101: "Venda",
    6101: "Venda",
    5103: "Venda",
    6103: "Venda",
    5104: "Venda",
    6104: "Venda",
    5105: "Venda",
    6105: "Venda",
    5106: "Venda",
    6106: "Venda",
    5109: "Venda",
    6109: "Venda",
    5110: "Venda",
    6110: "Venda",
    5111: "Venda",
    6111: "Venda",
    5112: "Venda",
    6112: "Venda",
    5113: "Venda",
    6113: "Venda",
    5114: "Venda",
    6114: "Venda",
    5115: "Venda",
    6115: "Venda",
    5116: "Venda",
    6116: "Venda",
    5117: "Venda",
    6117: "Venda",
    5118: "Venda",
    6118: "Venda",
    5119: "Venda",
    6119: "Venda",
    5120: "Venda",
    6120: "Venda",
    5122: "Venda",
    6122: "Venda",
    5123: "Venda",
    6123: "Venda",
    5124: "Venda",
    6124: "Venda",
    5125: "Venda",
    6125: "Venda",
    # Devoluções de venda
    5201: "Devolução",
    6201: "Devolução",
    5202: "Devolução",
    6202: "Devolução",
    5203: "Devolução",
    6203: "Devolução",
    5204: "Devolução",
    6204: "Devolução",
    5205: "Devolução",
    6205: "Devolução",
    5206: "Devolução",
    6206: "Devolução",
    5207: "Devolução",
    6207: "Devolução",
    5208: "Devolução",
    6208: "Devolução",
    5209: "Devolução",
    6209: "Devolução",
    5210: "Devolução",
    6210: "Devolução",
    # Devoluções de compra
    1201: "Devolução de Compra",
    2201: "Devolução de Compra",
    1202: "Devolução de Compra",
    2202: "Devolução de Compra",
    1203: "Devolução de Compra",
    2203: "Devolução de Compra",
    1204: "Devolução de Compra",
    2204: "Devolução de Compra",
    1205: "Devolução de Compra",
    2205: "Devolução de Compra",
    1206: "Devolução de Compra",
    2206: "Devolução de Compra",
    1207: "Devolução de Compra",
    2207: "Devolução de Compra",
    1208: "Devolução de Compra",
    2208: "Devolução de Compra",
    # Remessas
    5901: "Remessa",
    6901: "Remessa",
    5902: "Remessa",
    6902: "Remessa",
    5903: "Remessa",
    6903: "Remessa",
    5904: "Remessa",
    6904: "Remessa",
    5905: "Remessa",
    6905: "Remessa",
    5906: "Remessa",
    6906: "Remessa",
    5907: "Remessa",
    6907: "Remessa",
    5908: "Remessa",
    6908: "Remessa",
    5909: "Remessa",
    6909: "Remessa",
    5910: "Remessa",
    6910: "Remessa",
    5911: "Remessa",
    6911: "Remessa",
    5912: "Remessa",
    6912: "Remessa",
    5913: "Remessa",
    6913: "Remessa",
    5914: "Remessa",
    6914: "Remessa",
    5915: "Remessa",
    6915: "Remessa",
    5916: "Remessa",
    6916: "Remessa",
    5917: "Remessa",
    6917: "Remessa",
    5918: "Remessa",
    6918: "Remessa",
    5919: "Remessa",
    6919: "Remessa",
    5920: "Remessa",
    6920: "Remessa",
    5921: "Remessa",
    6921: "Remessa",
    5922: "Remessa",
    6922: "Remessa",
    5923: "Remessa",
    6923: "Remessa",
    5924: "Remessa",
    6924: "Remessa",
    5925: "Remessa",
    6925: "Remessa",
    # Outras saídas
    5949: "Outras Saídas",
    6949: "Outras Saídas",
    5929: "Outras Saídas",
    6929: "Outras Saídas",
}

# Mapeamento de Operação ML → Natureza
ML_OP_MAP = {
    "Venda":                "Venda",
    "Devolução":            "Devolução",
    "Retorno de Remessa":   "Remessa",
    "Insucesso de Entrega": "Remessa",
}

# ──────────────────────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────────────────────

def cfop_to_int(cfop_val):
    """Normaliza CFOP para inteiro (ex: '6.108-000' → 6108, 6108 → 6108)."""
    try:
        s = str(cfop_val).replace(".", "").replace("-000", "").replace("-", "").strip()
        return int(s[:4])
    except Exception:
        return None


def natureza_from_cfop(cfop_val):
    key = cfop_to_int(cfop_val)
    return CFOP_MAP.get(key, "Outros")


def natureza_from_ml_op(op_val):
    return ML_OP_MAP.get(str(op_val).strip(), "Outros")


def fmt_brl(v):
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "—"


def to_excel_download(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()


# ──────────────────────────────────────────────────────────────────────────────
# LEITURA E PARSE DOS ARQUIVOS
# ──────────────────────────────────────────────────────────────────────────────

def _limpar_valor_brl(v):
    """Converte 'R$ 1.235,21' ou 1235.21 para float."""
    if pd.isna(v):
        return np.nan
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    # remover R$, espaços, pontos de milhar, trocar vírgula decimal por ponto
    s = s.replace("R$", "").replace("\xa0", "").strip()
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return np.nan


def _read_ml(data_bytes):
    """
    ML.xlsx: aba 'Invoices', linha 0 = NaN, linha 1 = cabeçalho real, dados a partir da linha 2.
    Valores de Valor/Valor Total vêm como string 'R$ 929,00'.
    """
    df_raw = pd.read_excel(BytesIO(data_bytes), sheet_name="Invoices", header=None, engine="openpyxl")
    # encontrar a linha do cabeçalho (procura 'Status' ou 'Numero da NF')
    header_row = 1
    for i in range(min(10, len(df_raw))):
        row_vals = [str(v).strip() for v in df_raw.iloc[i].tolist()]
        if "Status" in row_vals or "Numero da NF-e" in row_vals:
            header_row = i
            break
    df_raw.columns = df_raw.iloc[header_row]
    df = df_raw.iloc[header_row + 1:].copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df[df.columns[df.columns.notna()]]
    df = df.reset_index(drop=True)
    return df


def _read_vhsys(data_bytes):
    """
    vhsys.xls é na verdade HTML disfarçado de XLS — usa read_html.
    Valores de Valor Total vêm como string 'R$ 1.180,00'.
    Remove colunas Unnamed extras.
    """
    tables = pd.read_html(BytesIO(data_bytes), encoding="utf-8")
    # pegar a maior tabela (a principal)
    df = max(tables, key=lambda t: t.shape[0] * t.shape[1])
    # remover colunas completamente vazias (Unnamed)
    df = df.loc[:, ~df.columns.str.startswith("Unnamed")]
    df = df.dropna(how="all").reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def parse_ml(uploaded_file, col_nf, col_status, col_operacao, col_serie, col_valor, col_emissao, col_nome):
    data = uploaded_file.read() if hasattr(uploaded_file, "read") else uploaded_file
    df = _read_ml(data)
    df = df.rename(columns={
        col_nf:       "NF",
        col_status:   "Status",
        col_operacao: "Operacao",
        col_serie:    "Serie",
        col_valor:    "Valor",
        col_emissao:  "Data",
        col_nome:     "Cliente",
    })
    df["NF"] = pd.to_numeric(df["NF"], errors="coerce")
    df["Valor"] = df["Valor"].apply(_limpar_valor_brl)
    df = df[df["NF"].notna() & df["Valor"].notna()].copy()
    df["Natureza"] = df["Operacao"].apply(natureza_from_ml_op)
    df["Origem"] = "ML"
    return df


def parse_vhsys(uploaded_file, col_nf, col_situacao, col_cfop, col_natureza, col_valor, col_emissao, col_cliente):
    data = uploaded_file.read() if hasattr(uploaded_file, "read") else uploaded_file
    df = _read_vhsys(data)
    df = df.rename(columns={
        col_nf:       "NF",
        col_situacao: "Situacao",
        col_cfop:     "CFOP_raw",
        col_natureza: "Natureza_raw",
        col_valor:    "Valor",
        col_emissao:  "Data",
        col_cliente:  "Cliente",
    })
    df["NF"] = pd.to_numeric(df["NF"], errors="coerce")
    df["Valor"] = df["Valor"].apply(_limpar_valor_brl)
    df = df[df["NF"].notna() & df["Valor"].notna()].copy()
    df["CFOP"] = df["CFOP_raw"].apply(cfop_to_int)
    df["Natureza"] = df["CFOP"].apply(lambda c: CFOP_MAP.get(c, "Outros"))
    df["Origem"] = "VHsys"
    return df


def parse_contabilidade(uploaded_file, col_nf_ini, col_nf_fim, col_cfop, col_valor, col_data, col_serie):
    """
    Contabilidade tem cabeçalho na linha 5 (índice 5), com rodapés/quebras de página.
    """
    df_raw = pd.read_excel(uploaded_file, header=None)

    # Detectar a linha do cabeçalho real (procurar linha com 'CFOP' ou 'Doc. ini.')
    header_row = 5
    for i in range(min(20, len(df_raw))):
        row_vals = [str(v).strip().lower() for v in df_raw.iloc[i].tolist()]
        if any(x in row_vals for x in ["cfop", "doc. ini.", "chave"]):
            header_row = i
            break

    df = pd.read_excel(uploaded_file, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    df = df.rename(columns={
        col_nf_ini: "NF",
        col_cfop:   "CFOP_raw",
        col_valor:  "Valor",
        col_data:   "Data",
        col_serie:  "Serie",
    })
    if col_nf_fim and col_nf_fim != col_nf_ini:
        df = df.rename(columns={col_nf_fim: "NF_fim"})

    # Filtrar linhas válidas: NF deve ser numérica (elimina cabeçalhos repetidos de páginas)
    df["NF"] = pd.to_numeric(df["NF"], errors="coerce")
    df = df[df["NF"].notna()].copy()

    # Remover eventuais linhas onde CFOP é texto (linha de cabeçalho repetida)
    df = df[df["CFOP_raw"].apply(lambda v: str(v).strip().upper() not in ("CFOP", "NAN", ""))].copy()

    df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce")
    df = df[df["Valor"].notna()].copy()
    df["CFOP"] = df["CFOP_raw"].apply(cfop_to_int)
    df["Natureza"] = df["CFOP"].apply(lambda c: CFOP_MAP.get(c, "Outros"))
    df["Origem"] = "Contabilidade"
    return df


# ──────────────────────────────────────────────────────────────────────────────
# FILTROS DINÂMICOS POR PLANILHA
# ──────────────────────────────────────────────────────────────────────────────

def render_filtros(df, prefix, label, default_cols=None, default_values=None):
    """Renderiza filtros de inclusão por coluna com suporte a valores pré-selecionados.

    default_cols: lista de colunas a pré-selecionar no multiselect de colunas.
    default_values: dict {coluna: [valores_padrão]} para os sub-filtros.
    Os defaults só têm efeito no primeiro carregamento; após isso o estado da sessão
    prevalece, permitindo que o usuário altere sem perder suas escolhas.
    """
    _dc = default_cols or []
    _dv = default_values or {}
    st.markdown(f"**Filtros — {label}**")
    colunas = df.columns.tolist()
    valid_default_cols = [c for c in _dc if c in colunas]
    col_filtro = st.multiselect(
        f"Colunas para filtrar ({label})",
        options=colunas,
        default=valid_default_cols,
        key=f"{prefix}_cols",
    )
    mask = pd.Series([True] * len(df), index=df.index)
    for col in col_filtro:
        valores_unicos = sorted(df[col].dropna().unique().tolist(), key=str)
        preset = [v for v in _dv.get(col, []) if v in valores_unicos]
        selecionados = st.multiselect(
            f"{label} → {col}: manter apenas",
            options=valores_unicos,
            default=preset,
            key=f"{prefix}_{col}_incl",
        )
        if selecionados:
            mask = mask & (df[col].isin(selecionados))
    return df[mask].copy()


# ──────────────────────────────────────────────────────────────────────────────
# COMPARATIVO
# ──────────────────────────────────────────────────────────────────────────────

def build_comparativo(df_ml, df_vh, df_cont):
    """
    ML e VHsys são tratados como um único sistema de Vendas.
    Cruzamento principal: Vendas (ML ∪ VHsys) × Contabilidade.
    Também detecta inconsistências internas entre ML e VHsys.
    """
    ml_base = df_ml.groupby("NF").agg(
        Valor_ML=("Valor", "sum"),
        Natureza_ML=("Natureza", lambda x: x.mode().iloc[0] if len(x) > 0 else ""),
    ).reset_index()

    vh_base = df_vh.groupby("NF").agg(
        Valor_VH=("Valor", "sum"),
        Natureza_VH=("Natureza", lambda x: x.mode().iloc[0] if len(x) > 0 else ""),
        CFOP_VH=("CFOP", lambda x: x.mode().iloc[0] if len(x) > 0 else ""),
    ).reset_index()

    cont_base = df_cont.groupby("NF").agg(
        Valor_Cont=("Valor", "sum"),
        Natureza_Cont=("Natureza", lambda x: x.mode().iloc[0] if len(x) > 0 else ""),
        CFOP_Cont=("CFOP_raw", lambda x: x.mode().iloc[0] if len(x) > 0 else ""),
    ).reset_index()

    comp = pd.merge(ml_base, vh_base, on="NF", how="outer")
    comp = pd.merge(comp, cont_base, on="NF", how="outer")

    # Natureza consolidada (VHsys tem CFOP → prioridade)
    def nat_consolidada(row):
        for nat in [row.get("Natureza_VH"), row.get("Natureza_ML"), row.get("Natureza_Cont")]:
            if pd.notna(nat) and str(nat).strip() not in ("", "nan"):
                return nat
        return "Outros"

    comp["Natureza"] = comp.apply(nat_consolidada, axis=1)

    # Flags de presença
    comp["Em_ML"]    = comp["Valor_ML"].notna()
    comp["Em_VH"]    = comp["Valor_VH"].notna()
    comp["Em_Cont"]  = comp["Valor_Cont"].notna()
    comp["Em_Vendas"] = comp["Em_ML"] | comp["Em_VH"]

    # Valor de Vendas: VHsys tem prioridade (tem CFOP); fallback para ML
    comp["Valor_Vendas"] = comp["Valor_VH"].combine_first(comp["Valor_ML"])

    tol = 0.05
    # Divergência interna: mesma NF em ML e VHsys com valores diferentes
    comp["Dif_ML_VH"] = (
        comp["Em_ML"] & comp["Em_VH"] &
        ((comp["Valor_ML"] - comp["Valor_VH"]).abs() > tol)
    )
    # Divergência Vendas × Contabilidade
    comp["Dif_Vendas_Cont"] = (
        comp["Em_Vendas"] & comp["Em_Cont"] &
        ((comp["Valor_Vendas"] - comp["Valor_Cont"]).abs() > tol)
    )

    comp["Tem_Diferenca"] = (
        comp["Em_Vendas"] != comp["Em_Cont"]
    )

    def desc_dif(row):
        msgs = []
        if row["Em_Cont"] and not row["Em_Vendas"]:  msgs.append("Só na Contabilidade")
        elif row["Em_Vendas"] and not row["Em_Cont"]: msgs.append("Só nas Vendas")
        if row["Dif_ML_VH"]:                          msgs.append("ML ≠ VHsys")
        if row["Dif_Vendas_Cont"]:                    msgs.append("Vendas ≠ Contabilidade")
        return "; ".join(msgs) if msgs else "OK"

    comp["Diferenca"] = comp.apply(desc_dif, axis=1)
    comp["NF"] = comp["NF"].astype(int)
    comp = comp.sort_values("NF").reset_index(drop=True)
    return comp


# ──────────────────────────────────────────────────────────────────────────────
# RESUMO POR NATUREZA
# ──────────────────────────────────────────────────────────────────────────────

def build_resumo(df_ml, df_vh, df_cont):
    rows = []
    for nat in sorted(set(
        df_ml["Natureza"].unique().tolist() +
        df_vh["Natureza"].unique().tolist() +
        df_cont["Natureza"].unique().tolist()
    )):
        v_ml   = df_ml[df_ml["Natureza"] == nat]["Valor"].sum()
        v_vh   = df_vh[df_vh["Natureza"] == nat]["Valor"].sum()
        v_cont = df_cont[df_cont["Natureza"] == nat]["Valor"].sum()
        q_ml   = df_ml[df_ml["Natureza"] == nat]["NF"].nunique()
        q_vh   = df_vh[df_vh["Natureza"] == nat]["NF"].nunique()
        q_cont = df_cont[df_cont["Natureza"] == nat]["NF"].nunique()
        rows.append({
            "Natureza":         nat,
            "Qtd ML":           q_ml,
            "Valor ML":         v_ml if v_ml else 0,
            "Qtd VHsys":        q_vh,
            "Valor VHsys":      v_vh if v_vh else 0,
            "Qtd Contab.":      q_cont,
            "Valor Contab.":    v_cont if v_cont else 0,
        })
    return pd.DataFrame(rows)


# ──────────────────────────────────────────────────────────────────────────────
# INTERFACE PRINCIPAL
# ──────────────────────────────────────────────────────────────────────────────

st.title("📊 Comparativo de Faturamento")

# ── SIDEBAR: Upload e mapeamento de colunas ──────────────────────────────────
with st.sidebar:
    st.header("1. Importar arquivos")

    f_ml   = st.file_uploader("📥 ML (Mercado Livre)",   type=["xlsx", "xls"])
    f_vh   = st.file_uploader("📥 VHsys",                type=["xlsx", "xls"])
    f_cont = st.file_uploader("📥 Contabilidade",        type=["xlsx", "xls"])

# ── Só avança se todos os arquivos estiverem carregados ──────────────────────
if not (f_ml and f_vh and f_cont):
    st.info("⬅️ Faça o upload dos três arquivos na barra lateral para começar.")
    st.stop()

# ── Ler colunas dos arquivos ──────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def get_cols_ml(data):
    return _read_ml(data).columns.tolist()

@st.cache_data(show_spinner=False)
def get_cols_vh(data):
    return _read_vhsys(data).columns.tolist()

@st.cache_data(show_spinner=False)
def get_cols_cont(data):
    df_raw = pd.read_excel(BytesIO(data), header=None)
    header_row = 5
    for i in range(min(20, len(df_raw))):
        row_vals = [str(v).strip().lower() for v in df_raw.iloc[i].tolist()]
        if any(x in row_vals for x in ["cfop", "doc. ini.", "chave"]):
            header_row = i
            break
    return pd.read_excel(BytesIO(data), header=header_row).columns.tolist()

bytes_ml   = f_ml.read()
bytes_vh   = f_vh.read()
bytes_cont = f_cont.read()

cols_ml   = get_cols_ml(bytes_ml)
cols_vh   = get_cols_vh(bytes_vh)
cols_cont = get_cols_cont(bytes_cont)

# ── Mapeamento de colunas ────────────────────────────────────────────────────
with st.sidebar:
    st.header("2. Mapear colunas")

    with st.expander("Colunas — ML", expanded=False):
        def sc(label, cols, default):
            idx = cols.index(default) if default in cols else 0
            return st.selectbox(label, cols, index=idx, key=f"ml_{label}")

        ml_col_nf       = sc("Número da NF",  cols_ml, "Numero da NF-e")
        ml_col_status   = sc("Status",         cols_ml, "Status")
        ml_col_operacao = sc("Operação",       cols_ml, "Operação")
        ml_col_serie    = sc("Série",          cols_ml, "Série")
        ml_col_valor    = sc("Valor",          cols_ml, "Valor")
        ml_col_emissao  = sc("Data Emissão",   cols_ml, "Emissão")
        ml_col_nome     = sc("Nome/Cliente",   cols_ml, "Nome")

    with st.expander("Colunas — VHsys", expanded=False):
        def sv(label, cols, default):
            idx = cols.index(default) if default in cols else 0
            return st.selectbox(label, cols, index=idx, key=f"vh_{label}")

        vh_col_nf       = sv("Número da NF",        cols_vh, "NF-e")
        vh_col_situacao = sv("Situação",             cols_vh, "Situação")
        vh_col_cfop     = sv("CFOP",                 cols_vh, "CFOP")
        vh_col_natureza = sv("Natureza de Operação", cols_vh, "Natureza de Operação")
        vh_col_valor    = sv("Valor Total",          cols_vh, "Valor Total")
        vh_col_emissao  = sv("Data Emissão",         cols_vh, "Data Emissão")
        vh_col_cliente  = sv("Cliente",              cols_vh, "Cliente")

    with st.expander("Colunas — Contabilidade", expanded=False):
        def sct(label, cols, default):
            idx = cols.index(default) if default in cols else 0
            return st.selectbox(label, cols, index=idx, key=f"cont_{label}")

        cont_col_nf     = sct("Número NF (Doc. ini.)", cols_cont, "Doc. ini.")
        cont_col_nf_fim = sct("Número NF (Doc. fim)",  cols_cont, "Doc. fim")
        cont_col_cfop   = sct("CFOP",                  cols_cont, "CFOP")
        cont_col_valor  = sct("Valor Contábil",        cols_cont, "Valor contábil")
        cont_col_data   = sct("Data",                  cols_cont, "Data")
        cont_col_serie  = sct("Série",                 cols_cont, "Série")

# ── Processar os DataFrames ──────────────────────────────────────────────────
@st.cache_data(show_spinner="Processando arquivos…")
def load_all(bml, bvh, bcont,
             ml_nf, ml_st, ml_op, ml_ser, ml_val, ml_em, ml_nm,
             vh_nf, vh_sit, vh_cfop, vh_nat, vh_val, vh_em, vh_cli,
             ct_nf, ct_nf2, ct_cfop, ct_val, ct_dt, ct_ser):
    df_ml   = parse_ml(bml,  ml_nf, ml_st, ml_op, ml_ser, ml_val, ml_em, ml_nm)
    df_vh   = parse_vhsys(bvh, vh_nf, vh_sit, vh_cfop, vh_nat, vh_val, vh_em, vh_cli)
    df_cont = parse_contabilidade(BytesIO(bcont), ct_nf, ct_nf2, ct_cfop, ct_val, ct_dt, ct_ser)
    return df_ml, df_vh, df_cont

df_ml_raw, df_vh_raw, df_cont_raw = load_all(
    bytes_ml, bytes_vh, bytes_cont,
    ml_col_nf, ml_col_status, ml_col_operacao, ml_col_serie, ml_col_valor, ml_col_emissao, ml_col_nome,
    vh_col_nf, vh_col_situacao, vh_col_cfop, vh_col_natureza, vh_col_valor, vh_col_emissao, vh_col_cliente,
    cont_col_nf, cont_col_nf_fim, cont_col_cfop, cont_col_valor, cont_col_data, cont_col_serie,
)

# ──────────────────────────────────────────────────────────────────────────────
# ABAS PRINCIPAIS
# ──────────────────────────────────────────────────────────────────────────────
tab_filtros, tab_resumo, tab_difs = st.tabs([
    "🔧 Filtros",
    "📋 Resumo por Natureza",
    "⚠️ Notas com Diferença",
])

# ── ABA FILTROS ──────────────────────────────────────────────────────────────
with tab_filtros:
    st.subheader("Filtrar registros da análise")
    st.caption(
        "Selecione a coluna e os valores que deseja **manter** em cada planilha. "
        "Se nenhum valor for selecionado na coluna, todos os registros são mantidos."
    )

    # Primeira NF do VHsys detectada automaticamente — usada como limite inferior da Contabilidade
    _nf_vh_min = int(df_vh_raw["NF"].min()) if df_vh_raw["NF"].notna().any() else 0

    c1, c2, c3 = st.columns(3)
    with c1:
        df_ml = render_filtros(
            df_ml_raw, "ml", "ML",
            default_cols=["Status", "Operacao", "Natureza", "Serie"],
            default_values={
                "Status":   ["Autorizada"],
                "Operacao": ["Venda"],
                "Natureza": ["Venda"],
                "Serie":    ["2"],
            },
        )
        # Filtro numérico de NF (ML)
        st.markdown("**Intervalo de NF (ML)**")
        _nf1, _nf2 = st.columns(2)
        nf_ml_min = _nf1.number_input(
            "NF ≥", min_value=0, value=0, step=1, key="ml_nf_min",
            help="0 = sem limite inferior"
        )
        nf_ml_max = _nf2.number_input(
            "NF ≤", min_value=0, value=0, step=1, key="ml_nf_max",
            help="0 = sem limite superior"
        )
        if nf_ml_min > 0:
            df_ml = df_ml[df_ml["NF"] >= nf_ml_min]
        if nf_ml_max > 0:
            df_ml = df_ml[df_ml["NF"] <= nf_ml_max]
        st.caption(f"ML: {len(df_ml_raw)} → **{len(df_ml)}** registros")

    with c2:
        df_vh = render_filtros(
            df_vh_raw, "vh", "VHsys",
            default_cols=["Situacao", "Natureza"],
            default_values={
                "Situacao": ["Atendido"],
                "Natureza": ["Venda"],
            },
        )
        st.caption(f"VHsys: {len(df_vh_raw)} → **{len(df_vh)}** registros")

    with c3:
        df_cont = render_filtros(
            df_cont_raw, "cont", "Contabilidade",
            default_cols=["Natureza"],
            default_values={"Natureza": ["Venda"]},
        )
        # Filtro de NF mínima — ignora notas anteriores ao início do ML
        st.markdown("**NF mínima (Contabilidade)**")
        nf_cont_min = st.number_input(
            "NF ≥  (ignora notas de outros canais)",
            min_value=0, value=_nf_vh_min, step=1, key="cont_nf_min",
            help=f"Pré-definido com a 1ª NF do VHsys ({_nf_vh_min}). 0 = sem limite.",
        )
        if nf_cont_min > 0:
            df_cont = df_cont[df_cont["NF"] >= nf_cont_min]
        st.caption(f"Contabilidade: {len(df_cont_raw)} → **{len(df_cont)}** registros")

# DataFrames filtrados ficam disponíveis nas demais abas via session_state
st.session_state["df_ml"]   = df_ml
st.session_state["df_vh"]   = df_vh
st.session_state["df_cont"] = df_cont

# Recupera DataFrames (pode ter sido definido acima na aba de filtros, mas
# como o Streamlit re-executa tudo, já estão nas variáveis locais)

# ── ABA RESUMO ───────────────────────────────────────────────────────────────
with tab_resumo:
    st.subheader("Resumo por Natureza de Operação")

    resumo = build_resumo(df_ml, df_vh, df_cont)

    tot_ml   = resumo["Valor ML"].sum()
    tot_vh   = resumo["Valor VHsys"].sum()
    tot_vend = tot_ml + tot_vh
    tot_cont = resumo["Valor Contab."].sum()
    qtd_ml   = int(resumo["Qtd ML"].sum())
    qtd_vh   = int(resumo["Qtd VHsys"].sum())
    qtd_cont = int(resumo["Qtd Contab."].sum())
    dif_total = tot_vend - tot_cont

    # ── Linha 1: totais principais ────────────────────────────────────────────
    st.markdown("#### Totais de Venda (ML + VHsys) vs Contabilidade")
    ca, cb, cc = st.columns(3)
    ca.metric(
        "🛒 Total Vendas (ML + VHsys)",
        fmt_brl(tot_vend),
        delta=f"{qtd_ml + qtd_vh} NFs",
        delta_color="off",
    )
    cb.metric(
        "📊 Total Contabilidade",
        fmt_brl(tot_cont),
        delta=f"{qtd_cont} NFs",
        delta_color="off",
    )
    cc.metric(
        "Δ Diferença (Vendas − Contab.)",
        fmt_brl(dif_total),
        delta_color="off",
    )

    # ── Linha 2: ML e VHsys separados (informativo) ───────────────────────────
    st.caption("Detalhe por canal de venda:")
    cd, ce, _ = st.columns([1, 1, 1])
    cd.metric("ML",    fmt_brl(tot_ml), f"{qtd_ml} NFs")
    ce.metric("VHsys", fmt_brl(tot_vh), f"{qtd_vh} NFs")

    st.divider()

    # ── Tabela por natureza ───────────────────────────────────────────────────
    resumo_disp = resumo.copy()
    resumo_disp["Valor Vendas"] = resumo_disp["Valor ML"] + resumo_disp["Valor VHsys"]
    resumo_disp["Qtd Vendas"]   = resumo_disp["Qtd ML"] + resumo_disp["Qtd VHsys"]
    for col in ["Valor ML", "Valor VHsys", "Valor Vendas", "Valor Contab."]:
        resumo_disp[col] = resumo_disp[col].apply(fmt_brl)
    cols_ord = ["Natureza", "Qtd ML", "Valor ML", "Qtd VHsys", "Valor VHsys",
                "Qtd Vendas", "Valor Vendas", "Qtd Contab.", "Valor Contab."]
    st.dataframe(resumo_disp[cols_ord], use_container_width=True, hide_index=True)

    # ── Gráfico: Vendas vs Contabilidade ──────────────────────────────────────
    st.divider()
    st.markdown("**Vendas (ML + VHsys) vs Contabilidade — por Natureza**")
    resumo_graf = resumo.copy()
    resumo_graf["Valor Vendas"] = resumo_graf["Valor ML"] + resumo_graf["Valor VHsys"]
    chart_data = resumo_graf.melt(
        id_vars="Natureza",
        value_vars=["Valor Vendas", "Valor Contab."],
        var_name="Sistema",
        value_name="Valor",
    )
    fig = px.bar(
        chart_data,
        x="Natureza",
        y="Valor",
        color="Sistema",
        barmode="group",
        color_discrete_map={"Valor Vendas": "#1f77b4", "Valor Contab.": "#d62728"},
        labels={"Valor": "R$", "Natureza": ""},
    )
    fig.update_layout(legend_title_text="")
    st.plotly_chart(fig, use_container_width=True)

    # Download
    st.download_button(
        "⬇️ Baixar resumo (Excel)",
        data=to_excel_download(resumo),
        file_name="resumo_natureza.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )




# ── ABA DIFERENÇAS ────────────────────────────────────────────────────────────
with tab_difs:
    st.subheader("⚠️ Relatório de Notas com Diferença")
    st.caption(
        "ML e VHsys são tratados como um único sistema de **Vendas**. "
        "As diferenças são analisadas entre Vendas (ML + VHsys) e Contabilidade, "
        "além de inconsistências internas entre os dois canais de venda."
    )

    _comp = build_comparativo(df_ml, df_vh, df_cont)

    # ── Categorias ───────────────────────────────────────────────────────────
    df_so_cont   = _comp[_comp["Em_Cont"]   & ~_comp["Em_Vendas"]].copy()
    df_so_vendas = _comp[_comp["Em_Vendas"] & ~_comp["Em_Cont"]].copy()
    df_todas     = _comp[_comp["Tem_Diferenca"]].copy()

    if df_todas.empty:
        st.success("✅ Nenhuma diferença encontrada entre as bases!")
    else:
        # ── Cards rápidos ────────────────────────────────────────────────────
        c1, c2, c3 = st.columns(3)
        c1.metric(
            "🔴 Só na Contabilidade",
            len(df_so_cont),
            delta=fmt_brl(df_so_cont["Valor_Cont"].sum()) if len(df_so_cont) else None,
            delta_color="inverse",
        )
        c2.metric(
            "🟡 Só nas Vendas (VHsys+ML)",
            len(df_so_vendas),
            delta=fmt_brl(df_so_vendas["Valor_Vendas"].sum()) if len(df_so_vendas) else None,
            delta_color="inverse",
        )
        c3.metric(
            "📋 Total com diferença",
            len(df_todas),
        )

        st.divider()

        sec = st.radio(
            "Ver seção:",
            [
                "🔴 Só na Contabilidade — sem ML nem VHsys",
                "🟡 Só nas Vendas (VHsys+ML) — sem Contabilidade",
                "📋 Todas as diferenças",
            ],
            horizontal=False,
            key="dif_sec",
        )

        _map_sec = {
            "🔴 Só na Contabilidade — sem ML nem VHsys": df_so_cont,
            "🟡 Só nas Vendas (VHsys+ML) — sem Contabilidade": df_so_vendas,
            "📋 Todas as diferenças":                       df_todas,
        }
        df_dif_show = _map_sec[sec]

        # Filtro por natureza
        nats_dif = ["Todas"] + sorted(df_dif_show["Natureza"].dropna().unique().tolist())
        nat_dif = st.selectbox("Natureza", nats_dif, key="dif_nat")
        if nat_dif != "Todas":
            df_dif_show = df_dif_show[df_dif_show["Natureza"] == nat_dif]

        # Totais
        if not df_dif_show.empty:
            ct1, ct2, ct3, ct4 = st.columns(4)
            ct1.metric("NFs", len(df_dif_show))
            ct2.metric("Valor ML",      fmt_brl(df_dif_show["Valor_ML"].fillna(0).sum()))
            ct3.metric("Valor VHsys",   fmt_brl(df_dif_show["Valor_VH"].fillna(0).sum()))
            ct4.metric("Valor Contab.", fmt_brl(df_dif_show["Valor_Cont"].fillna(0).sum()))

        # Tabela
        _cols_disp = ["NF", "Natureza", "Em_ML", "Valor_ML", "Em_VH", "Valor_VH",
                      "Em_Cont", "Valor_Cont", "Diferenca"]
        df_dif_disp = df_dif_show[_cols_disp].copy()
        for col in ["Valor_ML", "Valor_VH", "Valor_Cont"]:
            df_dif_disp[col] = df_dif_disp[col].apply(lambda v: fmt_brl(v) if pd.notna(v) else "—")
        df_dif_disp = df_dif_disp.rename(columns={
            "Em_ML":     "✔ ML",
            "Valor_ML":  "Valor ML",
            "Em_VH":     "✔ VHsys",
            "Valor_VH":  "Valor VHsys",
            "Em_Cont":   "✔ Contab.",
            "Valor_Cont": "Valor Contab.",
            "Diferenca":  "Tipo",
        })

        st.caption(f"{len(df_dif_disp)} NFs")
        st.dataframe(df_dif_disp, use_container_width=True, hide_index=True)

        st.download_button(
            "⬇️ Baixar esta seção (Excel)",
            data=to_excel_download(df_dif_show),
            file_name="relatorio_diferencas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
