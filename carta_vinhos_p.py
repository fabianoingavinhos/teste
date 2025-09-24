#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
app_streamlit_final_v3_persist2_aggrid.py

Alterações:
- Substituição do st.data_editor por st_aggrid (AgGrid)
- Checkboxes rápidos na 1ª coluna para seleção múltipla
- Colunas "fator" e "preco_de_venda" continuam editáveis
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

# --- AgGrid ---
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

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
    for ext in ['.png', '.jpg', '.jpeg', '.PNG', '.JPG', '.JPEG']:
        img_path = os.path.join(IMAGEM_DIR, f"{cod}{ext}")
        if os.path.exists(img_path):
            return os.path.abspath(img_path)
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

def ordenar_para_saida(df):
    def normaliza_tipo(t):
        t = str(t).strip().lower()
        if "espum" in t: return "Espumantes"
        if "branc" in t: return "Brancos"
        if "ros" in t: return "Rosés"
        if "tint" in t: return "Tintos"
        if "fris" in t: return "Frisantes"
        if "forti" in t: return "Fortificados"
        if "sobrem" in t: return "Vinhos de sobremesa"
        if "licor" in t: return "Licorosos"
        return t.title()
    tipos_norm = df.get("tipo", pd.Series([""]*len(df))).astype(str).map(normaliza_tipo)
    ordem_map = {t: i for i, t in enumerate(TIPO_ORDEM_FIXA)}
    ordem = tipos_norm.map(lambda x: ordem_map.get(x, 999))
    df2 = df.copy()
    df2["__tipo_ordem"] = ordem
    cols_exist = [c for c in ["__tipo_ordem","pais","descricao"] if c in df2.columns]
    return df2.sort_values(cols_exist).drop(columns=["__tipo_ordem"], errors="ignore")

# ===================== APP =====================
def main():
    st.set_page_config(page_title="Sugestão de Carta de Vinhos", layout="wide")
    garantir_pastas()

    if "selected_idxs" not in st.session_state:
        st.session_state.selected_idxs = set()
    if "manual_fat" not in st.session_state:
        st.session_state.manual_fat = {}
    if "manual_preco_venda" not in st.session_state:
        st.session_state.manual_preco_venda = {}

    st.markdown("### Sugestão de Carta de Vinhos")

    caminho_planilha = st.text_input("Arquivo de dados", value="vinhos1.xls")
    preco_flag = st.selectbox("Tabela de preço", ["preco1","preco2","preco15","preco38","preco39","preco55","preco63"], index=0)
    fator_global = st.number_input("Fator", min_value=0.0, value=2.0, step=0.1)

    df = ler_excel_vinhos(caminho_planilha)
    df = atualiza_coluna_preco_base(df, preco_flag, fator_global=float(fator_global))

    # === Grade com seleção (AGGRID) ===
    view_df = df.copy()
    gb = GridOptionsBuilder.from_dataframe(
        view_df[["cod","descricao","pais","regiao","preco_base","preco_de_venda","fator","idx"]]
    )
    gb.configure_selection("multiple", use_checkbox=True, header_checkbox=True)
    gb.configure_column("preco_base", type=["numericColumn"], precision=2)
    gb.configure_column("preco_de_venda", type=["numericColumn"], precision=2)
    gb.configure_column("fator", type=["numericColumn"], precision=2, editable=True)
    gridOptions = gb.build()

    grid_response = AgGrid(
        view_df,
        gridOptions=gridOptions,
        update_mode=GridUpdateMode.SELECTION_CHANGED | GridUpdateMode.VALUE_CHANGED,
        fit_columns_on_grid_load=True,
        allow_unsafe_jscode=True,
        theme="balham",
        height=500,
    )

    selected_rows = grid_response["selected_rows"]
    st.session_state.selected_idxs = {int(r["idx"]) for r in selected_rows if "idx" in r}

    if "data" in grid_response and grid_response["data"] is not None:
        edited_df = pd.DataFrame(grid_response["data"])
        for _, r in edited_df.iterrows():
            try:
                idx = int(r["idx"])
            except Exception:
                continue
            if pd.notnull(r.get("fator")):
                st.session_state.manual_fat[idx] = float(r["fator"])
            if pd.notnull(r.get("preco_de_venda")):
                st.session_state.manual_preco_venda[idx] = float(r["preco_de_venda"])

    st.caption(f"Selecionados: {len(st.session_state.selected_idxs)}")

if __name__ == "__main__":
    main()
