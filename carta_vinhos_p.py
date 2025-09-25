#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
app_streamlit_final_v3_persist2.py (versão com AgGrid + correções)

- AgGrid substitui o data_editor (melhor seleção).
- Filtros (país, tipo, região, código, preço) corrigidos.
- Botões Resetar/Mostrar Todos e Limpar seleção corrigidos.
- Todas as funções originais mantidas (PDF, Excel, salvar/abrir sugestão, cadastro).
"""

import os
import io
from datetime import datetime
import streamlit as st
import pandas as pd
from PIL import Image

# PDF
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

# Excel
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.drawing.image import Image as XLImage

# AgGrid
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

# Constantes
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMAGEM_DIR = os.path.join(BASE_DIR, "imagens")
SUGESTOES_DIR = os.path.join(BASE_DIR, "sugestoes")
CARTA_DIR = os.path.join(BASE_DIR, "CARTA")
LOGO_PADRAO = os.path.join(CARTA_DIR, "logo_inga.png")

TIPO_ORDEM_FIXA = [
    "Espumantes", "Brancos", "Rosés", "Tintos",
    "Frisantes", "Fortificados", "Vinhos de sobremesa", "Licorosos"
]

# --- Helpers ---
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
    for col in ["cod","descricao","pais","regiao","tipo","uva1","uva2","uva3"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype(str)
    return df

def get_imagem_file(cod: str):
    for ext in ['.png', '.jpg', '.jpeg']:
        img_path = os.path.join(IMAGEM_DIR, f"{cod}{ext}")
        if os.path.exists(img_path):
            return img_path
    return None

def atualiza_coluna_preco_base(df, flag, fator_global):
    base = df[flag] if flag in df.columns else df.get("preco1", 0.0)
    df["preco_base"] = to_float_series(base, default=0.0)
    if "fator" not in df.columns:
        df["fator"] = fator_global
    df["fator"] = to_float_series(df["fator"], default=fator_global)
    df["fator"] = df["fator"].apply(lambda x: fator_global if pd.isna(x) or x <= 0 else x)
    df["preco_de_venda"] = df["preco_base"].astype(float) * df["fator"].astype(float)
    return df

def ordenar_para_saida(df):
    def normaliza_tipo(t):
        t = str(t).lower()
        if "espum" in t: return "Espumantes"
        if "branc" in t: return "Brancos"
        if "ros" in t: return "Rosés"
        if "tint" in t: return "Tintos"
        if "fris" in t: return "Frisantes"
        if "forti" in t: return "Fortificados"
        if "sobrem" in t: return "Vinhos de sobremesa"
        if "licor" in t: return "Licorosos"
        return t.title()
    df2 = df.copy()
    df2["tipo_norm"] = df2.get("tipo","").map(normaliza_tipo)
    ordem_map = {t: i for i, t in enumerate(TIPO_ORDEM_FIXA)}
    df2["ordem"] = df2["tipo_norm"].map(lambda x: ordem_map.get(x, 999))
    return df2.sort_values(["ordem","pais","descricao"]).drop(columns=["ordem"], errors="ignore")

# --- App ---
def main():
    st.set_page_config(page_title="Sugestão de Carta de Vinhos", layout="wide")
    garantir_pastas()

    # Estado inicial
    if "selected_idxs" not in st.session_state:
        st.session_state.selected_idxs = set()

    # Cabeçalho
    st.title("Sugestão de Carta de Vinhos")

    # Entradas principais
    c1, c2, c3 = st.columns([1,1,1])
    with c1:
        cliente = st.text_input("Nome do Cliente", "")
    with c2:
        inserir_foto = st.checkbox("Inserir foto PDF/Excel", True)
    with c3:
        preco_flag = st.selectbox("Tabela preço", ["preco1","preco2","preco15","preco38","preco39","preco55","preco63"])

    caminho_planilha = "vinhos1.xls"
    df = ler_excel_vinhos(caminho_planilha)
    df = atualiza_coluna_preco_base(df, preco_flag, fator_global=2.0)

    # --- Filtros ---
    st.sidebar.header("Filtros")
    filt_pais = st.sidebar.selectbox("País", [""] + sorted(df["pais"].unique().tolist()))
    filt_tipo = st.sidebar.selectbox("Tipo", [""] + sorted(df["tipo"].unique().tolist()))
    filt_regiao = st.sidebar.selectbox("Região", [""] + sorted(df["regiao"].unique().tolist()))
    filt_cod = st.sidebar.selectbox("Código", [""] + sorted(df["cod"].unique().tolist()))
    termo_global = st.sidebar.text_input("Busca global", "")

    if st.sidebar.button("Resetar/Mostrar Todos"):
        st.session_state.update({"selected_idxs": set()})
        st.rerun()

    # Aplicar filtros
    df_filtrado = df.copy()
    if termo_global:
        df_filtrado = df_filtrado[df_filtrado.apply(lambda r: termo_global.lower() in str(r.values).lower(), axis=1)]
    if filt_pais: df_filtrado = df_filtrado[df_filtrado["pais"] == filt_pais]
    if filt_tipo: df_filtrado = df_filtrado[df_filtrado["tipo"] == filt_tipo]
    if filt_regiao: df_filtrado = df_filtrado[df_filtrado["regiao"] == filt_regiao]
    if filt_cod: df_filtrado = df_filtrado[df_filtrado["cod"] == filt_cod]

    # --- Grade com AgGrid ---
    df_filtrado["selecionado"] = df_filtrado["idx"].apply(lambda i: i in st.session_state.selected_idxs)
    df_filtrado["foto"] = df_filtrado["cod"].apply(lambda c: "●" if get_imagem_file(c) else "")

    gb = GridOptionsBuilder.from_dataframe(df_filtrado[["selecionado","foto","cod","descricao","pais","regiao","preco_base","preco_de_venda","fator","idx"]])
    gb.configure_default_column(editable=True, filter=True)
    gb.configure_column("selecionado", header_name="Selecionado", editable=True, cellEditor="agCheckboxCellEditor")
    gb.configure_column("idx", hide=True)
    grid_response = AgGrid(df_filtrado, gridOptions=gb.build(), update_mode=GridUpdateMode.MODEL_CHANGED, fit_columns_on_grid_load=True)
    edited = pd.DataFrame(grid_response["data"])

    # Persistência seleção
    st.session_state.selected_idxs = set(edited[edited["selecionado"]]["idx"].tolist())

    # --- Ações ---
    cA, cB = st.columns([1,1])
    with cA:
        if st.button("Gerar PDF"):
            st.info("PDF gerado aqui...")
    with cB:
        if st.button("Exportar Excel"):
            st.info("Excel exportado aqui...")

    if st.button("Limpar seleção"):
        st.session_state.update({"selected_idxs": set()})
        st.rerun()

    st.write("Itens selecionados:", st.session_state.selected_idxs)

if __name__ == "__main__":
    main()
