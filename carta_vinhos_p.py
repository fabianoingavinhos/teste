#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
app_streamlit_final_v3_persist2.py (versão com AgGrid)

Alterações:
- Substituído o uso de st.data_editor (que simulava um checkcombobox lento)
  por AgGrid com checkboxes nativos.
- Mantidas todas as funções originais (filtros, salvar/abrir sugestão, PDF, Excel, etc).
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
from openpyxl.utils import get_column_letter
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
    if "prev_view_state" not in st.session_state:
        st.session_state.prev_view_state = {}
    if "manual_fat" not in st.session_state:
        st.session_state.manual_fat = {}
    if "manual_preco_venda" not in st.session_state:
        st.session_state.manual_preco_venda = {}
    if "cadastrados" not in st.session_state:
        st.session_state.cadastrados = []

    st.markdown("### Sugestão de Carta de Vinhos")

    # --- Entrada principal ---
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
                                         help="Caminho do arquivo XLS/XLSX", key="caminho_planilha")

    # --- Carregar DF ---
    df = ler_excel_vinhos(caminho_planilha)
    df = atualiza_coluna_preco_base(df, preco_flag, fator_global=float(fator_global))

    # --- Sidebar filtros ---
    st.sidebar.header("Filtros")
    pais_opc = [""] + sorted([p for p in df["pais"].dropna().astype(str).unique() if p])
    tipo_opc = [""] + sorted([t for t in df["tipo"].dropna().astype(str).unique() if t])
    desc_opc = [""] + sorted([d for d in df["descricao"].dropna().astype(str).unique() if d])
    regiao_opc = [""] + sorted([r for r in df["regiao"].dropna().astype(str).unique() if r])
    cod_opc = [""] + sorted([str(c) for c in df["cod"].dropna().astype(str).unique()])

    filt_pais = st.sidebar.selectbox("País", pais_opc, index=0, key="filt_pais")
    filt_tipo = st.sidebar.selectbox("Tipo", tipo_opc, index=0, key="filt_tipo")
    filt_desc = st.sidebar.selectbox("Descrição", desc_opc, index=0, key="filt_desc")
    filt_regiao = st.sidebar.selectbox("Região", regiao_opc, index=0, key="filt_regiao")
    filt_cod = st.sidebar.selectbox("Código", cod_opc, index=0, key="filt_cod")

    colp1, colp2 = st.sidebar.columns(2)
    with colp1:
        preco_min = st.number_input("Preço mín (base)", min_value=0.0, value=0.0, step=1.0, key="preco_min")
    with colp2:
        preco_max = st.number_input("Preço máx (base)", min_value=0.0, value=0.0, step=1.0, help="0 = sem limite", key="preco_max")

    # --- Aplicar filtros ---
    df_filtrado = df.copy()
    if termo_global.strip():
        term = termo_global.strip().lower()
        mask = df_filtrado.apply(lambda row: term in " ".join(str(v).lower() for v in row.values), axis=1)
        df_filtrado = df_filtrado[mask]
    if filt_pais: df_filtrado = df_filtrado[df_filtrado["pais"] == filt_pais]
    if filt_tipo: df_filtrado = df_filtrado[df_filtrado["tipo"] == filt_tipo]
    if filt_desc: df_filtrado = df_filtrado[df_filtrado["descricao"] == filt_desc]
    if filt_regiao: df_filtrado = df_filtrado[df_filtrado["regiao"] == filt_regiao]
    if filt_cod: df_filtrado = df_filtrado[df_filtrado["cod"].astype(str) == filt_cod]
    if preco_min:
        df_filtrado = df_filtrado[df_filtrado["preco_base"].fillna(0) >= float(preco_min)]
    if preco_max and preco_max > 0:
        df_filtrado = df_filtrado[df_filtrado["preco_base"].fillna(0) <= float(preco_max)]
    if resetar: df_filtrado = df.copy()

    # === Grade com seleção via AgGrid ===
    view_df = df_filtrado.copy()
    view_df["selecionado"] = view_df["idx"].apply(lambda i: i in st.session_state.selected_idxs)
    view_df["foto"] = view_df["cod"].apply(lambda c: "●" if get_imagem_file(str(c)) else "")

    gb = GridOptionsBuilder.from_dataframe(
        view_df[["selecionado","foto","cod","descricao","pais","regiao","preco_base","preco_de_venda","fator","idx"]]
    )
    gb.configure_default_column(editable=True, filter=True, resizable=True)
    gb.configure_column("selecionado", header_name="Selecionado", editable=True, cellEditor="agCheckboxCellEditor")
    gb.configure_column("foto", editable=False)
    gb.configure_column("idx", editable=False, hide=True)
    gb.configure_column("preco_base", type=["numericColumn"], valueFormatter="x.toFixed(2)")
    gb.configure_column("preco_de_venda", type=["numericColumn"], valueFormatter="x.toFixed(2)")
    gb.configure_column("fator", type=["numericColumn"], valueFormatter="x.toFixed(2)")
    grid_options = gb.build()

    grid_response = AgGrid(
        view_df,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        fit_columns_on_grid_load=True,
        enable_enterprise_modules=False,
        use_container_width=True,
        height=500,
        key="aggrid_main"
    )
    edited = pd.DataFrame(grid_response["data"])

    # --- Persistência incremental das seleções ---
    curr_state = {}
    if isinstance(edited, pd.DataFrame) and not edited.empty:
        for _, row in edited.iterrows():
            try:
                idx_i = int(row["idx"])
            except Exception:
                continue
            sel = bool(row.get("selecionado", False))
            curr_state[idx_i] = sel

    prev_state = st.session_state.get("prev_view_state", {})
    global_sel = set(st.session_state.selected_idxs)
    to_add = {i for i, s in curr_state.items() if s and prev_state.get(i) is not True}
    to_remove = {i for i, s in curr_state.items() if (prev_state.get(i) is True) and not s}
    global_sel |= to_add
    global_sel -= to_remove
    st.session_state.selected_idxs = global_sel
    st.session_state.prev_view_state = curr_state

    # --- Ajustes manuais de fator e preço venda ---
    if isinstance(edited, pd.DataFrame) and not edited.empty:
        for _, r in edited.iterrows():
            try:
                idx = int(r["idx"])
            except Exception:
                continue
            if pd.notnull(r.get("fator")):
                st.session_state.manual_fat[idx] = float(r["fator"])
            if pd.notnull(r.get("preco_de_venda")):
                st.session_state.manual_preco_venda[idx] = float(r["preco_de_venda"])

    for idx, fat in st.session_state.manual_fat.items():
        df.loc[df["idx"]==idx, "fator"] = float(fat)
    df["fator"] = to_float_series(df["fator"], default=float(fator_global))
    df["fator"] = df["fator"].apply(lambda x: float(fator_global) if pd.isna(x) or x <= 0 else float(x))
    df["preco_de_venda"] = (df["preco_base"].astype(float) * df["fator"].astype(float)).astype(float)
    for idx, pv in st.session_state.manual_preco_venda.items():
        df.loc[df["idx"]==idx, "preco_de_venda"] = float(pv)

    # --- daqui em diante permanece igual ao original (botões, PDF, Excel, salvar sugestão, tabs, etc.) ---

    # (COLE O RESTANTE DO SEU CÓDIGO ORIGINAL AQUI, sem mudanças)

if __name__ == "__main__":
    main()
