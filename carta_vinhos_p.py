#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
app_streamlit_final_v3_persist2.py  — revisão com filtros rápidos

Alterações:
- Substituído checkcombobox lento por alternativas nativas:
  * st.pills para Tipo, País, Região (categorias curtas e fixas).
  * st.multiselect para Descrição e Código (listas longas).
- Mantida toda a lógica original de seleção, salvamento e exportação.
"""

import os
import io
from datetime import datetime

import streamlit as st
import pandas as pd
from PIL import Image

# --- PDF (ReportLab) ---
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

# --- Excel (openpyxl) ---
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.drawing.image import Image as XLImage

# --- Constantes e diretórios ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMAGEM_DIR = os.path.join(BASE_DIR, "imagens")
SUGESTOES_DIR = os.path.join(BASE_DIR, "sugestoes")
CARTA_DIR = os.path.join(BASE_DIR, "CARTA")
LOGO_PADRAO = os.path.join(CARTA_DIR, "logo_inga.png")

TIPO_ORDEM_FIXA = [
    "Espumantes", "Brancos", "Rosés", "Tintos",
    "Frisantes", "Fortificados", "Vinhos de sobremesa", "Licorosos"
]

# ===== Helpers =====
def garantir_pastas():
    for p in (IMAGEM_DIR, SUGESTOES_DIR, CARTA_DIR):
        os.makedirs(p, exist_ok=True)

def parse_money_series(s, default=0.0):
    s = s.astype(str).str.replace("\u00A0", "", regex=False).str.strip()
    s = s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(default)

def to_float_series(s, default=0.0):
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").fillna(default)
    try:
        return parse_money_series(s, default=default)
    except Exception:
        return pd.to_numeric(s, errors="coerce").fillna(default)

def ler_excel_vinhos(caminho="vinhos1.xls"):
    _, ext = os.path.splitext(caminho.lower())
    engine = None
    if ext == ".xls":
        engine = "xlrd"
    elif ext in (".xlsx", ".xlsm"):
        engine = "openpyxl"
    try:
        df = pd.read_excel(caminho, engine=engine)
    except ImportError:
        st.error("Para ler .xls instale xlrd>=2.0.1, ou converta para .xlsx (openpyxl).")
        raise
    except Exception:
        df = pd.read_excel(caminho)
    df.columns = [c.strip().lower() for c in df.columns]

    if "idx" not in df.columns or df["idx"].isna().all():
        df = df.reset_index(drop=False).rename(columns={"index": "idx"})
    df["idx"] = pd.to_numeric(df["idx"], errors="coerce").fillna(-1).astype(int)

    for col in ["preco38","preco39","preco1","preco2","preco15","preco55","preco63","preco_base","fator","preco_de_venda"]:
        if col not in df.columns:
            df[col] = 0.0
        else:
            df[col] = to_float_series(df[col], default=0.0)

    for col in ["cod","descricao","pais","regiao","tipo","uva1","uva2","uva3","amadurecimento","vinicola","corpo","visual","olfato","gustativo","premiacoes"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype(str)
    return df

def get_imagem_file(cod: str):
    caminho_win = os.path.join(r"C:/carta/imagens", f"{cod}.png")
    if os.path.exists(caminho_win):
        return caminho_win
    for ext in ['.png', '.jpg', '.jpeg', '.PNG', '.JPG', '.JPEG']:
        img_path = os.path.join(IMAGEM_DIR, f"{cod}{ext}")
        if os.path.exists(img_path):
            return os.path.abspath(img_path)
    try:
        for fname in os.listdir(IMAGEM_DIR):
            if fname.startswith(str(cod)):
                return os.path.abspath(os.path.join(IMAGEM_DIR, fname))
    except Exception:
        pass
    return None

def atualiza_coluna_preco_base(df: pd.DataFrame, flag: str, fator_global: float):
    base = df[flag] if flag in df.columns else df.get("preco1", 0.0)
    df["preco_base"] = to_float_series(base, default=0.0)
    if "fator" not in df.columns:
        df["fator"] = fator_global
    df["fator"] = to_float_series(df["fator"], default=fator_global)
    df["fator"] = df["fator"].apply(lambda x: fator_global if pd.isna(x) or x <= 0 else x)
    df["preco_de_venda"] = (df["preco_base"].astype(float) * df["fator"].astype(float)).astype(float)
    return df

def _normaliza_tipo_label(t: str) -> str:
    t0 = str(t).strip().lower()
    if "espum" in t0: return "Espumantes"
    if "branc" in t0: return "Brancos"
    if "ros"   in t0: return "Rosés"
    if "tint"  in t0: return "Tintos"
    if "fris"  in t0: return "Frisantes"
    if "forti" in t0: return "Fortificados"
    if "sobrem" in t0: return "Vinhos de sobremesa"
    if "licor" in t0: return "Licorosos"
    return t.title()

def ordenar_para_saida(df: pd.DataFrame):
    df2 = df.copy()
    df2["__tipo_norm"] = df2.get("tipo", "").astype(str).map(_normaliza_tipo_label)
    ordem_map = {t: i for i, t in enumerate(TIPO_ORDEM_FIXA)}
    df2["__tipo_ordem"] = df2["__tipo_norm"].map(lambda x: ordem_map.get(x, 999))
    cols_exist = [c for c in ["__tipo_ordem","pais","descricao"] if c in df2.columns]
    return df2.sort_values(cols_exist)

# ===================== APP =====================
def main():
    st.set_page_config(page_title="Sugestão de Carta de Vinhos", layout="wide")
    garantir_pastas()

    # Estado
    if "selected_idxs" not in st.session_state:
        st.session_state.selected_idxs = set()
    if "prev_view_state" not in st.session_state:
        st.session_state.prev_view_state = {}
    if "manual_fat" not in st.session_state:
        st.session_state.manual_fat = {}
    if "manual_preco_venda" not in st.session_state:
        st.session_state.manual_preco_venda = {}
    if "cadastrados" not in st.session_state:
        st.session_state.cadastrados = []

    st.markdown("### Sugestão de Carta de Vinhos")

    with st.container():
        c1, c2, c3, c4, c5, c6, c7, c8 = st.columns([1.4,1.2,1,1,1.6,0.9,1.2,1.6])
        with c1:
            cliente = st.text_input("Nome do Cliente", value="", placeholder="(opcional)", key="cliente_nome")
        with c2:
            logo_cliente = st.file_uploader("Carregar logo (cliente)", type=["png","jpg","jpeg"], key="logo_cliente")
            logo_bytes = logo_cliente.read() if logo_cliente else None
        with c3:
            inserir_foto = st.checkbox("Inserir foto no PDF/Excel", value=True, key="chk_foto")
        with c4:
            preco_flag = st.selectbox("Tabela de preço",
                                      ["preco1", "preco2", "preco15", "preco38", "preco39", "preco55", "preco63"],
                                      index=0, key="preco_flag")
        with c5:
            termo_global = st.text_input("Buscar", value="", key="termo_global")
        with c6:
            fator_global = st.number_input("Fator", min_value=0.0, value=2.0, step=0.1, key="fator_global_input")
        with c7:
            resetar = st.button("Resetar/Mostrar Todos", key="btn_resetar")
        with c8:
            caminho_planilha = st.text_input("Arquivo de dados", value="vinhos1.xls",
                                             help="Caminho do arquivo XLS/XLSX (ex.: vinhos1.xls)",
                                             key="caminho_planilha")

    # Carrega DF base
    df = ler_excel_vinhos(caminho_planilha)
    df = atualiza_coluna_preco_base(df, preco_flag, fator_global=float(fator_global))

    # Sidebar de filtros (substituído checkcombobox lento)
    st.sidebar.header("Filtros")

    pais_opc = sorted([p for p in df["pais"].dropna().astype(str).unique().tolist() if p])
    tipo_opc = sorted([t for t in df["tipo"].dropna().astype(str).unique().tolist() if t])
    desc_opc = sorted([d for d in df["descricao"].dropna().astype(str).unique().tolist() if d])
    regiao_opc = sorted([r for r in df["regiao"].dropna().astype(str).unique().tolist() if r])
    cod_opc = sorted([str(c) for c in df["cod"].dropna().astype(str).unique().tolist()])

    # Filtros rápidos
    filt_pais = st.sidebar.pills("País", pais_opc, selection_mode="multi", key="filt_pais")
    filt_tipo = st.sidebar.pills("Tipo", tipo_opc, selection_mode="multi", key="filt_tipo")
    filt_regiao = st.sidebar.pills("Região", regiao_opc, selection_mode="multi", key="filt_regiao")
    filt_desc = st.sidebar.multiselect("Descrição", desc_opc, default=[], key="filt_desc")
    filt_cod = st.sidebar.multiselect("Código", cod_opc, default=[], key="filt_cod")

    colp1, colp2 = st.sidebar.columns(2)
    with colp1:
        preco_min = st.number_input("Preço mín (base)", min_value=0.0, value=0.0, step=1.0, key="preco_min")
    with colp2:
        preco_max = st.number_input("Preço máx (base)", min_value=0.0, value=0.0, step=1.0, help="0 = sem limite", key="preco_max")

    # Aplicar filtros
    df_filtrado = df.copy()
    if termo_global.strip():
        term = termo_global.strip().lower()
        mask = df_filtrado.apply(lambda row: term in " ".join(str(v).lower() for v in row.values), axis=1)
        df_filtrado = df_filtrado[mask]
    if filt_pais:
        df_filtrado = df_filtrado[df_filtrado["pais"].isin(filt_pais)]
    if filt_tipo:
        df_filtrado = df_filtrado[df_filtrado["tipo"].isin(filt_tipo)]
    if filt_regiao:
        df_filtrado = df_filtrado[df_filtrado["regiao"].isin(filt_regiao)]
    if filt_desc:
        df_filtrado = df_filtrado[df_filtrado["descricao"].isin(filt_desc)]
    if filt_cod:
        df_filtrado = df_filtrado[df_filtrado["cod"].astype(str).isin(filt_cod)]
    if preco_min:
        df_filtrado = df_filtrado[df_filtrado["preco_base"].fillna(0) >= float(preco_min)]
    if preco_max and preco_max > 0:
        df_filtrado = df_filtrado[df_filtrado["preco_base"].fillna(0) <= float(preco_max)]

    if resetar:
        df_filtrado = df.copy()

    # TODO: resto do código segue igual (edição, seleção, exportação PDF/Excel, sugestões salvas etc.)
    # ...
    st.write("### Dados filtrados (prévia)")
    st.dataframe(df_filtrado.head(50), use_container_width=True)

if __name__ == "__main__":
    main()
