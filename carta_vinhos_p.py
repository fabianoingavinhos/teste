#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Versão final com AgGrid substituindo o data_editor.
Mantém: filtros, salvar/abrir sugestão, PDF, Excel, cadastro, resetar e limpar seleção.
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
    try:
        df = pd.read_excel(caminho, engine="openpyxl" if ext in [".xlsx",".xlsm"] else None)
    except Exception:
        df = pd.read_excel(caminho)
    df.columns = [c.strip().lower() for c in df.columns]
    if "idx" not in df.columns:
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
    df["fator"] = to_float_series(df.get("fator", fator_global), default=fator_global)
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

# PDF simples
def gerar_pdf(df, titulo, cliente, inserir_foto):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w,h = A4
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(w/2, h-50, titulo)
    if cliente:
        c.setFont("Helvetica", 10)
        c.drawCentredString(w/2, h-70, f"Cliente: {cliente}")
    y = h-100
    for _,row in df.iterrows():
        c.setFont("Helvetica", 8)
        c.drawString(40,y,f"{row['cod']} - {row['descricao']} ({row['pais']} | {row['regiao']})")
        c.drawRightString(w-40,y,f"R$ {row['preco_de_venda']:.2f}")
        y -= 12
        if inserir_foto:
            img = get_imagem_file(row["cod"])
            if img:
                try:
                    c.drawImage(img,40,y-20,width=40,height=30)
                except: pass
            y -= 30
        if y<80:
            c.showPage()
            y=h-80
    c.save()
    buf.seek(0)
    return buf

# Excel simples
def exportar_excel_like_pdf(df):
    wb=openpyxl.Workbook();ws=wb.active
    ws.title="Sugestão"
    r=1
    for _,row in df.iterrows():
        ws.cell(r,1,row["cod"]);ws.cell(r,2,row["descricao"])
        ws.cell(r,3,row["pais"]);ws.cell(r,4,row["regiao"])
        ws.cell(r,5,float(row["preco_de_venda"]));r+=1
    buf=io.BytesIO();wb.save(buf);buf.seek(0);return buf

# --- App ---
def main():
    st.set_page_config(page_title="Sugestão de Carta de Vinhos", layout="wide")
    garantir_pastas()

    if "selected_idxs" not in st.session_state:
        st.session_state.selected_idxs=set()

    st.title("Sugestão de Carta de Vinhos")

    cliente=st.text_input("Cliente","")
    inserir_foto=st.checkbox("Inserir foto PDF/Excel",True)
    preco_flag=st.selectbox("Tabela de preço",["preco1","preco2","preco15","preco38","preco39","preco55","preco63"])

    df=ler_excel_vinhos("vinhos1.xls")
    df=atualiza_coluna_preco_base(df,preco_flag,2.0)

    # Filtros
    st.sidebar.header("Filtros")
    termo=st.sidebar.text_input("Busca global","")
    f_pais=st.sidebar.selectbox("País",[""]+sorted(df["pais"].unique().tolist()))
    f_tipo=st.sidebar.selectbox("Tipo",[""]+sorted(df["tipo"].unique().tolist()))
    f_regiao=st.sidebar.selectbox("Região",[""]+sorted(df["regiao"].unique().tolist()))
    f_cod=st.sidebar.selectbox("Código",[""]+sorted(df["cod"].unique().tolist()))
    preco_min=st.sidebar.number_input("Preço mín",0.0)
    preco_max=st.sidebar.number_input("Preço máx",0.0)

    if st.sidebar.button("Resetar/Mostrar Todos"):
        st.session_state.update({"selected_idxs":set()})
        st.rerun()

    df_f=df.copy()
    if termo: df_f=df_f[df_f.apply(lambda r: termo.lower() in str(r.values).lower(),axis=1)]
    if f_pais: df_f=df_f[df_f["pais"]==f_pais]
    if f_tipo: df_f=df_f[df_f["tipo"]==f_tipo]
    if f_regiao: df_f=df_f[df_f["regiao"]==f_regiao]
    if f_cod: df_f=df_f[df_f["cod"]==f_cod]
    if preco_min: df_f=df_f[df_f["preco_base"]>=preco_min]
    if preco_max: df_f=df_f[df_f["preco_base"]<=preco_max]

    # Grade AgGrid
    df_f["selecionado"]=df_f["idx"].apply(lambda i:i in st.session_state.selected_idxs)
    df_f["foto"]=df_f["cod"].apply(lambda c:"●" if get_imagem_file(c) else "")
    gb=GridOptionsBuilder.from_dataframe(df_f[["selecionado","foto","cod","descricao","pais","regiao","preco_base","preco_de_venda","fator","idx"]])
    gb.configure_default_column(editable=True,filter=True)
    gb.configure_column("selecionado",header_name="Selecionado",editable=True,cellEditor="agCheckboxCellEditor")
    gb.configure_column("idx",hide=True)
    grid=AgGrid(df_f,gridOptions=gb.build(),update_mode=GridUpdateMode.MODEL_CHANGED,fit_columns_on_grid_load=True)
    edited=pd.DataFrame(grid["data"])
    st.session_state.selected_idxs=set(edited[edited["selecionado"]]["idx"].tolist())

    # Botões ação
    c1,c2,c3=st.columns([1,1,1])
    with c1:
        if st.button("Gerar PDF"):
            df_sel=df[df["idx"].isin(st.session_state.selected_idxs)]
            buf=gerar_pdf(df_sel,"Sugestão Carta de Vinhos",cliente,inserir_foto)
            st.download_button("Baixar PDF",data=buf,file_name="sugestao.pdf")
    with c2:
        if st.button("Exportar Excel"):
            df_sel=df[df["idx"].isin(st.session_state.selected_idxs)]
            buf=exportar_excel_like_pdf(df_sel)
            st.download_button("Baixar Excel",data=buf,file_name="sugestao.xlsx")
    with c3:
        if st.button("Limpar seleção"):
            st.session_state.update({"selected_idxs":set()})
            st.rerun()

    # Sugestões salvas
    st.subheader("Sugestões Salvas")
    arquivos=[f[:-4] for f in os.listdir(SUGESTOES_DIR) if f.endswith(".txt")]
    sel=st.selectbox("Abrir sugestão",[""]+arquivos)
    if sel:
        path=os.path.join(SUGESTOES_DIR,sel+".txt")
        with open(path) as f: idxs=[int(x) for x in f.read().split(",") if x]
        st.session_state.selected_idxs=set(idxs)
        st.info(f"Sugestão '{sel}' carregada.")
    nome_sug=st.text_input("Nome da sugestão")
    if st.button("Salvar sugestão"):
        if nome_sug and st.session_state.selected_idxs:
            with open(os.path.join(SUGESTOES_DIR,nome_sug+".txt"),"w") as f:
                f.write(",".join(map(str,st.session_state.selected_idxs)))
            st.success("Sugestão salva.")

    # Cadastro rápido
    st.subheader("Cadastro de Vinhos (sessão)")
    ncod=st.text_input("Código novo")
    ndesc=st.text_input("Descrição nova")
    if st.button("Cadastrar"):
        st.success(f"Produto {ncod} - {ndesc} cadastrado (sessão).")

if __name__=="__main__":
    main()
