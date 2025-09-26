#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
app_streamlit_final_v3_persist2.py

Novidades:
- "Sugestões Salvas": ao selecionar uma sugestão, ela é CARREGADA automaticamente,
  a relação de itens aparece abaixo e você pode incluir novos itens e salvar MESCLANDO.
- "Preço de venda" agora é garantido: preco_de_venda = preco_base * fator
  com parsing robusto (vírgula decimal) e fator zerado/NaN substituído pelo fator global.
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

def add_pdf_footer(c, contagem, total_rotulos, fator_geral):
    width, height = A4
    y_rodape = 35
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    c.setLineWidth(0.4)
    c.line(30, y_rodape+32, width-30, y_rodape+32)
    c.setFont("Helvetica", 5)
    c.drawString(32, y_rodape+20, f"Gerado em: {now}")
    try:
        fator_str = f"{float(fator_geral):.2f}"
    except Exception:
        fator_str = str(fator_geral)
    c.setFont("Helvetica-Bold", 6)
    c.drawString(32, y_rodape+7,
        f"Brancos: {contagem.get('Brancos',0)} | Tintos: {contagem.get('Tintos',0)} | "
        f"Rosés: {contagem.get('Rosés',0)} | Espumantes: {contagem.get('Espumantes',0)} | "
        f"Total: {int(total_rotulos)} | Fator: {fator_str}")
    c.setFont("Helvetica", 5)
    c.drawString(32, y_rodape-5, "Ingá Distribuidora Ltda | CNPJ 05.390.477/0002-25 Rod BR 232, KM 18,5 - S/N- Manassu - CEP 54130-340 Jaboatão")
    c.setFont("Helvetica-Bold", 6)
    c.drawString(width-190, y_rodape-5, "b2b.ingavinhos.com.br")

def gerar_pdf(df, titulo, cliente, inserir_foto, logo_cliente_bytes=None):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    if logo_cliente_bytes:
        try:
            c.drawImage(ImageReader(io.BytesIO(logo_cliente_bytes)), 40, height-60, width=120, height=40, mask='auto')
        except Exception:
            pass
    if os.path.exists(LOGO_PADRAO):
        try:
            c.drawImage(LOGO_PADRAO, width-80, height-40, width=48, height=24, mask='auto')
        except Exception:
            pass

    x_texto = 90
    y = height - 40
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(width/2, y, titulo)
    y -= 20
    if cliente:
        c.setFont("Helvetica", 10)
        c.drawCentredString(width/2, y, f"Cliente: {cliente}")
        y -= 20

    ordem_geral = 1
    contagem = {'Brancos':0, 'Tintos':0, 'Rosés':0, 'Espumantes':0, 'outros':0}
    df_sorted = ordenar_para_saida(df)

    for tipo in df_sorted['tipo'].fillna("").astype(str).unique():
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x_texto, y, str(tipo).upper()); y -= 14
        for pais in df_sorted[df_sorted['tipo']==tipo]['pais'].dropna().unique():
            c.setFont("Helvetica-Bold", 8)
            c.drawString(x_texto, y, str(pais).upper()); y -= 12
            grupo = df_sorted[(df_sorted['tipo']==tipo) & (df_sorted['pais']==pais)]
            for _, row in grupo.iterrows():
                t = str(tipo).lower()
                if "branc" in t: contagem["Brancos"] += 1
                elif "tint" in t: contagem["Tintos"] += 1
                elif "ros" in t: contagem["Rosés"] += 1
                elif "espum" in t: contagem["Espumantes"] += 1
                else: contagem["outros"] += 1

                c.setFont("Helvetica", 6)
                try: codtxt = f"{int(row['cod'])}"
                except Exception: codtxt = str(row.get('cod',""))
                c.drawString(x_texto, y, f"{ordem_geral:02d} ({codtxt})")
                c.setFont("Helvetica-Bold", 7)
                c.drawString(x_texto+55, y, str(row['descricao']))
                uvas = [str(row.get(f"uva{i}", "")).strip() for i in range(1,4)]
                uvas = [u for u in uvas if u and u.lower() != "nan"]
                regiao_str = f"{row.get('pais','')} | {row.get('regiao','')}"
                if uvas: regiao_str += f" | {', '.join(uvas)}"
                c.setFont("Helvetica", 5); c.drawString(x_texto+55, y-10, regiao_str)

                amad = str(row.get("amadurecimento", ""))
                if amad and amad.lower() != "nan":
                    c.setFont("Helvetica", 7); c.drawString(220, y-7, "🛢️")

                c.setFont("Helvetica", 5)
                try: c.drawRightString(width-120, y, f"(R$ {float(row['preco_base']):.2f})")
                except Exception: c.drawRightString(width-120, y, "(R$ -)")
                c.setFont("Helvetica-Bold", 7)
                try: c.drawRightString(width-40, y, f"R$ {float(row['preco_de_venda']):.2f}")
                except Exception: c.drawRightString(width-40, y, "R$ -")

                if inserir_foto:
                    imgfile = get_imagem_file(str(row.get('cod','')))
                    if imgfile:
                        try:
                            c.drawImage(imgfile, x_texto+340, y-2, width=40, height=30, mask='auto'); y -= 28
                        except Exception: y -= 20
                    else:
                        y -= 20
                else:
                    y -= 20

                ordem_geral += 1

                if y < 100:
                    add_pdf_footer(c, contagem, ordem_geral-1, fator_geral=df.get('fator', pd.Series([0])).median())
                    c.showPage()
                    y = height - 40
                    if logo_cliente_bytes:
                        try: c.drawImage(ImageReader(io.BytesIO(logo_cliente_bytes)), 40, height-60, width=120, height=40, mask='auto')
                        except Exception: pass
                    if os.path.exists(LOGO_PADRAO):
                        try: c.drawImage(LOGO_PADRAO, width-80, height-40, width=48, height=24, mask='auto')
                        except Exception: pass
                    c.setFont("Helvetica-Bold", 16); c.drawCentredString(width/2, y, titulo); y -= 20
                    if cliente: c.setFont("Helvetica", 10); c.drawCentredString(width/2, y, f"Cliente: {cliente}"); y -= 20

    add_pdf_footer(c, contagem, ordem_geral-1, fator_geral=df.get('fator', pd.Series([0])).median())
    c.save(); buffer.seek(0)
    return buffer

def exportar_excel_like_pdf(df, inserir_foto=True):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sugestão"
    row_num = 1; ordem_geral = 1
    df_sorted = ordenar_para_saida(df)
    for tipo in df_sorted['tipo'].fillna("").unique():
        ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=8)
        cell = ws.cell(row=row_num, column=1, value=str(tipo).upper()); cell.font = Font(bold=True, size=18); row_num += 1
        for pais in df_sorted[df_sorted['tipo'] == tipo]['pais'].dropna().unique():
            ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=8)
            cell = ws.cell(row=row_num, column=1, value=str(pais).upper()); cell.font = Font(bold=True, size=14); row_num += 1
            grupo = df_sorted[(df_sorted['tipo'] == tipo) & (df_sorted['pais'] == pais)]
            for _, row in grupo.iterrows():
                ws.cell(row=row_num, column=1, value=f"{ordem_geral:02d} ({int(row['cod']) if str(row['cod']).isdigit() else ''})").font = Font(size=11)
                ws.cell(row=row_num, column=2, value=str(row['descricao'])).font = Font(bold=True, size=12)
                if inserir_foto:
                    imgfile = get_imagem_file(str(row.get('cod','')))
                    if imgfile and os.path.exists(imgfile):
                        try:
                            img = XLImage(imgfile); img.width, img.height = 32, 24; ws.add_image(img, f"C{row_num}")
                        except Exception: pass
                try:
                    base_val = float(row['preco_base']); pv_val = float(row['preco_de_venda'])
                    base_str = f"(R$ {base_val:.2f})"; pv_str = f"R$ {pv_val:.2f}"
                except Exception:
                    base_str = "(R$ -)"; pv_str = "R$ -"
                ws.cell(row=row_num, column=7, value=base_str).alignment = Alignment(horizontal='right'); ws.cell(row=row_num, column=7).font = Font(size=10)
                ws.cell(row=row_num, column=8, value=pv_str).font = Font(bold=True, size=13); ws.cell(row=row_num, column=8).alignment = Alignment(horizontal='right')
                uvas = [str(row.get(f"uva{i}", "")).strip() for i in range(1,4)]; uvas = [u for u in uvas if u and u.lower() != "nan"]
                regiao_str = f"{row.get('pais','')} | {row.get('regiao','')}"; 
                if uvas: regiao_str += f" | {', '.join(uvas)}"
                ws.cell(row=row_num+1, column=2, value=regiao_str).font = Font(size=10)
                amad = str(row.get("amadurecimento", ""))
                if amad and amad.lower() != "nan":
                    ws.cell(row=row_num+1, column=3, value="🛢️").font = Font(size=10)
                row_num += 2; ordem_geral += 1
    stream = io.BytesIO(); wb.save(stream); stream.seek(0); return stream

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
    if "cadastr
