#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
app_streamlit_final_v3_persist2.py

Novidades:
- "Sugestões Salvas": ao selecionar uma sugestão, ela é CARREGADA automaticamente,
  a relação de itens aparece abaixo e você pode incluir novos itens e salvar MESCLANDO.
- "Preço de venda" agora é garantido: preco_de_venda = preco_base * fator
  com parsing robusto (vírgula decimal) e fator zerado/NaN substituído pelo fator global.
- Coluna "Selecionado" corrigida (Solução 2 com session_state):
  clique é imediato, sem lentidão ou desmarcar outro item.
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
    if "selecionados" not in st.session_state:
        st.session_state.selecionados = {}

    st.markdown("### Sugestão de Carta de Vinhos")

    caminho_planilha = "vinhos1.xls"
    df = ler_excel_vinhos(caminho_planilha)
    df = atualiza_coluna_preco_base(df, "preco1", fator_global=2.0)

    # === Grade com seleção ===
    view_df = df.copy()
    view_df["idx"] = pd.to_numeric(view_df["idx"], errors="coerce").fillna(-1).astype(int)
    view_df["cod"] = view_df.get("cod", "").astype(str)

    # Preenche coluna Selecionado com estado salvo
    view_df["Selecionado"] = view_df["idx"].apply(
        lambda i: st.session_state.selecionados.get(i, False)
    )
    view_df["foto"] = view_df["cod"].apply(
        lambda c: "●" if get_imagem_file(str(c)) else ""
    )

    edited = st.data_editor(
        view_df[["Selecionado","foto","cod","descricao","pais","regiao",
                 "preco_base","preco_de_venda","fator","idx"]],
        hide_index=True,
        column_config={
            "Selecionado": st.column_config.CheckboxColumn("SELECIONADO"),
            "foto": st.column_config.TextColumn("FOTO"),
            "cod": st.column_config.TextColumn("COD"),
            "descricao": st.column_config.TextColumn("DESCRICAO"),
            "pais": st.column_config.TextColumn("PAIS"),
            "regiao": st.column_config.TextColumn("REGIAO"),
            "preco_base": st.column_config.NumberColumn("PRECO_BASE", format="R$ %.2f"),
            "preco_de_venda": st.column_config.NumberColumn("PRECO_VENDA", format="R$ %.2f"),
            "fator": st.column_config.NumberColumn("FATOR", format="%.2f"),
            "idx": st.column_config.NumberColumn("IDX"),
        },
        use_container_width=True,
        num_rows="dynamic",
        key="editor_main",
    )

    # Atualiza session_state a cada clique
    if isinstance(edited, pd.DataFrame) and not edited.empty:
        for _, row in edited.iterrows():
            try:
                idx = int(row["idx"])
            except Exception:
                continue
            st.session_state.selecionados[idx] = bool(row.get("Selecionado", False))

    st.session_state.selected_idxs = {
        i for i, marcado in st.session_state.selecionados.items() if marcado
    }

    st.write("Selecionados:", st.session_state.selected_idxs)

if __name__ == "__main__":
    main()
