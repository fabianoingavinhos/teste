import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog, Toplevel
import pandas as pd
from PIL import Image, ImageTk
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from datetime import datetime
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMAGEM_DIR = os.path.join(BASE_DIR, "imagens")
SUGESTOES_DIR = "sugestoes"
CARTA_DIR = "CARTA"
LOGO_PADRAO = os.path.join(CARTA_DIR, "logo_inga.png")

# Campos obrigatórios, incluindo 'tipo'
CAMPOS_NOVOS = [
    "cod", "descricao", "visual", "olfato", "gustativo", "premiacoes", "amadurecimento",
    "regiao", "pais", "vinicola", "corpo", "tipo",
    "uva1", "uva2", "uva3",
    "preco38", "preco39", "preco1", "preco2", "preco15", "preco55", "preco63"
]

class WineMenuApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sugestão de Carta de Vinhos")
        self.root.geometry("1200x600")

        if not os.path.isdir(SUGESTOES_DIR):
            os.makedirs(SUGESTOES_DIR)
        if not os.path.isdir(CARTA_DIR):
            os.makedirs(CARTA_DIR)

        # Lê vinhos1.xls e garante todos os campos obrigatórios
        self.data = pd.read_excel('vinhos1.xls')
        self.data.columns = [c.strip().lower() for c in self.data.columns]
        for col in CAMPOS_NOVOS:
            if col not in self.data.columns:
                self.data[col] = ""
        # Garante tipos corretos para os campos de preço
        for col in ["preco38", "preco39", "preco1", "preco2", "preco15", "preco55", "preco63"]:
            self.data[col] = pd.to_numeric(self.data[col], errors='coerce').fillna(0.0)

        # Flag para escolher tabela de preço
        self.preco_flag_var = tk.StringVar(value="preco1")
        self.atualiza_coluna_preco_base()

        if "fator" not in self.data.columns:
            self.data["fator"] = 2.0
        self.data["fator"] = pd.to_numeric(self.data["fator"], errors='coerce').fillna(2.0)
        self.data["preco_de_venda"] = self.data["preco_base"] * self.data["fator"]
        if "cadastrado_manual" not in self.data.columns:
            self.data["cadastrado_manual"] = False

        self.filtered_data = self.data.copy()
        self.selected_indices = set()
        self.logo_path = None
        self.manual_prices = {}
        self.manual_fat = {}

        self.sugestoes_dict = {}
        self.sugestao_nome_var = tk.StringVar()

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True)
        self.frame_sugestao = ttk.Frame(self.notebook)
        self.notebook.add(self.frame_sugestao, text="Sugestão")
        self.frame_salvas = ttk.Frame(self.notebook)
        self.notebook.add(self.frame_salvas, text="Sugestões Salvas")
        self.frame_cadastro = ttk.Frame(self.notebook)
        self.notebook.add(self.frame_cadastro, text="Cadastro de Vinhos")

        self.build_sugestao_tab()
        self.build_salvas_tab()
        self.build_cadastro_tab()

    def atualiza_coluna_preco_base(self):
        flag = getattr(self, "preco_flag_var", None)
        flag = flag.get() if flag else "preco1"
        if flag not in self.data.columns:
            self.data["preco_base"] = self.data["preco1"] if "preco1" in self.data else 0.0
        else:
            self.data["preco_base"] = self.data[flag].fillna(0.0)
        self.data["preco_de_venda"] = self.data["preco_base"] * self.data.get("fator", 2.0)
        self.filtered_data = self.data.copy()

    def muda_tabela_preco(self, event=None):
        self.atualiza_coluna_preco_base()
        self.refresh_table()
        self.update_label_contagem()

    def build_sugestao_tab(self):
        f = self.frame_sugestao
        titulo_principal = tk.Label(f, text="Sugestão de Carta de Vinhos", font=("Arial", 14, "bold"))
        titulo_principal.pack(pady=5)

        super_top_frame = ttk.Frame(f)
        super_top_frame.pack(fill='x', padx=5, pady=2)

        ttk.Label(super_top_frame, text="Nome do Cliente:", font=("Arial", 10)).pack(side='left', padx=(0,2))
        self.title_var = tk.StringVar(value="")
        ttk.Entry(super_top_frame, textvariable=self.title_var, width=30, font=("Arial", 10)).pack(side='left', padx=2)
        ttk.Button(super_top_frame, text="Carregar logo", command=self.load_logo).pack(side='left', padx=5)
        self.logo_label = ttk.Label(super_top_frame)
        self.logo_label.pack(side='left', padx=5)

        self.inserir_foto_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(super_top_frame, text="Inserir foto no PDF/Excel", variable=self.inserir_foto_var).pack(side='left', padx=10)

        ttk.Label(super_top_frame, text="Tabela de preço:", font=("Arial", 10)).pack(side='left', padx=(10,2))
        opcoes_precos = ["preco1", "preco2", "preco15", "preco38", "preco39", "preco55", "preco63"]
        preco_menu = ttk.Combobox(super_top_frame, textvariable=self.preco_flag_var, values=opcoes_precos, width=8)
        preco_menu.pack(side='left', padx=2)
        preco_menu.bind("<<ComboboxSelected>>", self.muda_tabela_preco)

        ttk.Label(super_top_frame, text="  ").pack(side='left')
        ttk.Label(super_top_frame, text="Buscar:", font=("Arial", 10)).pack(side='left', padx=(0,2))
        self.global_filter_var = tk.StringVar()
        global_entry = ttk.Entry(super_top_frame, textvariable=self.global_filter_var, width=30, font=("Arial", 10))
        global_entry.pack(side='left', padx=2)
        global_entry.bind('<KeyRelease>', self.apply_global_filter)
        ttk.Label(super_top_frame, text="Fator:", font=("Arial", 10)).pack(side="left", padx=(10,2))
        self.fator_var = tk.StringVar(value="2.0")
        fator_entry = ttk.Entry(super_top_frame, textvariable=self.fator_var, width=5, font=("Arial",10))
        fator_entry.pack(side="left")
        fator_entry.bind('<KeyRelease>', lambda e: self.atualiza_fator_geral())
        ttk.Button(super_top_frame, text="Resetar/Mostrar Todos", command=self.resetar_tudo).pack(side='right', padx=5)

        filter_frame = ttk.LabelFrame(f, text="Filtros")
        filter_frame.pack(fill='x', padx=5, pady=2)
        self.pais_var = tk.StringVar()
        self.tipo_var = tk.StringVar()
        self.desc_var = tk.StringVar()
        self.regiao_var = tk.StringVar()
        self.cod_var = tk.StringVar()
        self.preco_min_var = tk.StringVar()
        self.preco_max_var = tk.StringVar()
        paises = [""] + sorted(self.data['pais'].dropna().unique().tolist())
        tipos = [""] + sorted(self.data['tipo'].dropna().unique().tolist())
        descrs = [""] + sorted(self.data['descricao'].dropna().unique().tolist())
        regioes = [""] + sorted(self.data['regiao'].dropna().unique().tolist())
        codigos = [""] + sorted(map(str, self.data['cod'].dropna().unique().tolist()))

        ttk.Label(filter_frame, text="País:", font=("Arial", 8)).grid(row=0, column=0, padx=2)
        pais_cb = ttk.Combobox(filter_frame, textvariable=self.pais_var, values=paises, width=12, font=("Arial", 8))
        pais_cb.grid(row=0, column=1, padx=2)
        pais_cb.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())
        ttk.Label(filter_frame, text="Tipo:", font=("Arial", 8)).grid(row=0, column=2, padx=2)
        tipo_cb = ttk.Combobox(filter_frame, textvariable=self.tipo_var, values=tipos, width=12, font=("Arial", 8))
        tipo_cb.grid(row=0, column=3, padx=2)
        tipo_cb.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())
        ttk.Label(filter_frame, text="Descrição:", font=("Arial", 8)).grid(row=0, column=4, padx=2)
        desc_cb = ttk.Combobox(filter_frame, textvariable=self.desc_var, values=descrs, width=15, font=("Arial", 8))
        desc_cb.grid(row=0, column=5, padx=2)
        desc_cb.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())
        ttk.Label(filter_frame, text="Região:", font=("Arial", 8)).grid(row=0, column=6, padx=2)
        regiao_cb = ttk.Combobox(filter_frame, textvariable=self.regiao_var, values=regioes, width=12, font=("Arial", 8))
        regiao_cb.grid(row=0, column=7, padx=2)
        regiao_cb.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())
        ttk.Label(filter_frame, text="Código:", font=("Arial", 8)).grid(row=0, column=8, padx=2)
        cod_cb = ttk.Combobox(filter_frame, textvariable=self.cod_var, values=codigos, width=8, font=("Arial", 8))
        cod_cb.grid(row=0, column=9, padx=2)
        cod_cb.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())
        ttk.Label(filter_frame, text="Preço Mín:", font=("Arial", 8)).grid(row=0, column=10, padx=2)
        preco_min_entry = ttk.Entry(filter_frame, textvariable=self.preco_min_var, width=6, font=("Arial", 8))
        preco_min_entry.grid(row=0, column=11, padx=2)
        preco_min_entry.bind('<KeyRelease>', lambda e: self.apply_filters())
        ttk.Label(filter_frame, text="Preço Máx:", font=("Arial", 8)).grid(row=0, column=12, padx=2)
        preco_max_entry = ttk.Entry(filter_frame, textvariable=self.preco_max_var, width=6, font=("Arial", 8))
        preco_max_entry.grid(row=0, column=13, padx=2)
        preco_max_entry.bind('<KeyRelease>', lambda e: self.apply_filters())
        self.check_var_select_all = tk.BooleanVar()
        self.check_var_select_none = tk.BooleanVar()
        ttk.Checkbutton(filter_frame, text="Selecionar todos", variable=self.check_var_select_all, command=self.select_all_filtered).grid(row=0, column=14, padx=2)
        ttk.Checkbutton(filter_frame, text="Desmarcar todos", variable=self.check_var_select_none, command=self.deselect_all_filtered).grid(row=0, column=15, padx=2)
        clear_btn = ttk.Button(filter_frame, text="Limpar filtros", command=self.clear_filters)
        clear_btn.grid(row=0, column=16, padx=5)

        style = ttk.Style()
        style.configure("Treeview.Heading", font=('Arial', 8, 'bold'))
        style.configure("Treeview", font=('Arial', 8))
        style.map("Treeview", background=[("selected", "#e0ecef")])
        style.configure("Treeview.even", background="#e6e6e6")
        style.configure("Treeview.odd", background="#ffffff")
        self.tree = ttk.Treeview(
            f,
            columns=("selecionado", "foto", "cod", "descricao", "pais", "regiao", "preco_base", "preco_de_venda", "fator"),
            show="headings",
            selectmode="none",
            height=12
        )
        for col, w in zip(
            ("selecionado", "foto", "cod", "descricao", "pais", "regiao", "preco_base", "preco_de_venda", "fator"),
            [30, 30, 50, 200, 70, 70, 80, 90, 60]
        ):
            self.tree.heading(col, text=col.upper())
            self.tree.column(col, width=w)
        self.tree.pack(fill="both", expand=True, padx=5, pady=2)
        self.tree.bind('<ButtonRelease-1>', self.on_tree_click)
        self.tree.bind('<Double-1>', self.edit_price_or_fator)
        self.tree.tag_configure('manual', font=('Arial', 8))
        self.tree.tag_configure('even', background="#e6e6e6")
        self.tree.tag_configure('odd', background="#ffffff")
        self.label_contagem = ttk.Label(f, text="", font=("Arial", 8), anchor='e')
        self.label_contagem.pack(fill="x", padx=5, pady=(0,5))

        btn_frame = ttk.Frame(f)
        btn_frame.pack(pady=5)
        ttk.Button(btn_frame, text="Visualizar Sugestão", command=self.visualizar_pdf).pack(side='left', padx=3)
        ttk.Button(btn_frame, text="Visualizar Itens Marcados", command=self.visualizar_marcados).pack(side='left', padx=3)
        ttk.Button(btn_frame, text="Gerar PDF", command=self.generate_pdf_layout).pack(side='left', padx=3)
        ttk.Button(btn_frame, text="Exportar para Excel", command=self.export_to_excel_layout).pack(side='left', padx=3)
        ttk.Button(btn_frame, text="Salvar Sugestão", command=self.salvar_sugestao_dialog).pack(side='left', padx=3)
        self.update_label_contagem()
        self.refresh_table()

    def update_label_contagem(self):
        df = self.filtered_data
        contagem = {'Brancos': 0, 'Tintos': 0, 'Rosés': 0, 'Espumantes': 0, 'outros': 0}
        tipo_map = {'branco': 'Brancos', 'tinto': 'Tintos', 'rose': 'Rosés', 'rosé': 'Rosés', 'espumante': 'Espumantes'}
        for tipo in df['tipo'].dropna().unique():
            tipo_label = next((lbl for k, lbl in tipo_map.items() if k in str(tipo).lower()), 'outros')
            contagem[tipo_label] = len(df[df['tipo'] == tipo])
        total = len(df)
        selecionados = len(self.selected_indices)
        fator_geral = self.fator_var.get()
        texto = f"Brancos: {contagem.get('Brancos', 0)} | Tintos: {contagem.get('Tintos', 0)} | Rosés: {contagem.get('Rosés', 0)} | Espumantes: {contagem.get('Espumantes', 0)} | Total: {total} | Selecionados: {selecionados} | Fator: {fator_geral}"
        self.label_contagem.config(text=texto)

    def atualiza_fator_geral(self):
        try:
            fat = float(self.fator_var.get())
            self.data['fator'] = fat
            self.data['preco_de_venda'] = self.data['preco_base'] * self.data['fator']
            for idx in self.manual_fat:
                self.data.at[idx, 'fator'] = self.manual_fat[idx]
                self.data.at[idx, 'preco_de_venda'] = self.data.at[idx, 'preco_base'] * self.manual_fat[idx]
            self.filtered_data = self.data.copy()
            self.refresh_table()
            self.update_label_contagem()
        except Exception:
            pass

    def refresh_table(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        ordem = 1
        for idx_filt, row in self.filtered_data.iterrows():
            idx = row.name
            preco = row['preco_base'] if pd.notnull(row['preco_base']) else 0.0
            fator = self.manual_fat.get(idx, row.get('fator', 2.0))
            preco_venda = self.manual_prices.get(idx, preco * fator)
            tag = 'manual' if idx in self.manual_fat else ''
            tag += ' even' if ordem % 2 == 0 else ' odd'
            imgfile = self.get_imagem_file(str(row['cod']))
            
            x_texto = 90
            tem_foto = "●" if imgfile else ""
            self.tree.insert('', 'end', iid=idx, values=(
                "✔" if idx in self.selected_indices else "",
                tem_foto,
                row.get('cod', ''),
                row.get('descricao', ''),
                row.get('pais', ''),
                row.get('regiao', ''),
                f"R$ {preco:.2f}",
                f"R$ {preco_venda:.2f}",
                f"{fator:.2f}"
            ), tags=(tag,))
            ordem += 1
        self.update_label_contagem()

    def on_tree_click(self, event):
        row_id = self.tree.identify_row(event.y)
        if not row_id: return
        idx = int(row_id)
        if idx in self.selected_indices:
            self.selected_indices.remove(idx)
        else:
            self.selected_indices.add(idx)
        self.refresh_table()

    def edit_price_or_fator(self, event):
        col = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if not row_id: return
        idx = int(row_id)
        row = self.data.loc[idx]
        preco = row['preco_base']
        if col == "#9":
            fat_antigo = self.manual_fat.get(idx, row.get('fator', 2.0))
            novo_fat = simpledialog.askfloat("Fator", f"Novo fator para {row['descricao']}:", initialvalue=fat_antigo)
            if novo_fat is not None:
                self.manual_fat[idx] = novo_fat
                self.data.at[idx, 'fator'] = novo_fat
                self.data.at[idx, 'preco_de_venda'] = self.data.at[idx, 'preco_base'] * novo_fat
                self.filtered_data = self.data.copy()
                self.refresh_table()
        elif col == "#8":
            preco_venda_antigo = self.manual_prices.get(idx, row.get('preco_de_venda', preco * row.get('fator', 2.0)))
            novo_preco_venda = simpledialog.askfloat("Ajustar Preço de Venda", f"Novo preço de venda para {row['descricao']}:", initialvalue=preco_venda_antigo)
            if novo_preco_venda is not None:
                self.manual_prices[idx] = novo_preco_venda
                self.data.at[idx, 'preco_de_venda'] = novo_preco_venda
                self.filtered_data = self.data.copy()
                self.refresh_table()

    def apply_global_filter(self, event=None):
        term = self.global_filter_var.get().strip().lower()
        if not term:
            self.filtered_data = self.data.copy()
        else:
            df = self.data.copy()
            mask = df.apply(lambda row: term in " ".join(str(v).lower() for v in row.values), axis=1)
            self.filtered_data = df[mask]
        self.refresh_table()

    def apply_filters(self, event=None):
        df = self.data.copy()
        term = self.global_filter_var.get().strip().lower()
        if term:
            mask = df.apply(lambda row: term in " ".join(str(v).lower() for v in row.values), axis=1)
            df = df[mask]
        if self.pais_var.get():
            df = df[df['pais'] == self.pais_var.get()]
        if self.tipo_var.get():
            df = df[df['tipo'] == self.tipo_var.get()]
        if self.desc_var.get():
            df = df[df['descricao'] == self.desc_var.get()]
        if self.regiao_var.get():
            df = df[df['regiao'] == self.regiao_var.get()]
        if self.cod_var.get():
            df = df[df['cod'].astype(str) == self.cod_var.get()]
        try:
            if self.preco_min_var.get():
                min_price = float(self.preco_min_var.get())
                df = df[df['preco_base'] >= min_price]
        except ValueError:
            pass
        try:
            if self.preco_max_var.get():
                max_price = float(self.preco_max_var.get())
                df = df[df['preco_base'] <= max_price]
        except ValueError:
            pass
        self.filtered_data = df
        self.refresh_table()
        if self.filtered_data.empty:
            messagebox.showinfo("Filtros", "Nenhum vinho encontrado com os filtros aplicados.")

    def select_all_filtered(self):
        self.selected_indices |= set(self.filtered_data.index)
        self.refresh_table()
        self.check_var_select_all.set(False)

    def deselect_all_filtered(self):
        self.selected_indices -= set(self.filtered_data.index)
        self.refresh_table()
        self.check_var_select_none.set(False)

    def clear_filters(self):
        self.pais_var.set("")
        self.tipo_var.set("")
        self.desc_var.set("")
        self.regiao_var.set("")
        self.cod_var.set("")
        self.preco_min_var.set("")
        self.preco_max_var.set("")
        self.global_filter_var.set("")
        self.filtered_data = self.data.copy()
        self.selected_indices.clear()
        self.refresh_table()

    def resetar_tudo(self):
        self.filtered_data = self.data.copy()
        self.selected_indices.clear()
        self.refresh_table()

    def load_logo(self):
        filepath = filedialog.askopenfilename(title="Selecione a logo", filetypes=[("Imagens", "*.jpg *.jpeg *.png")])
        if filepath:
            self.logo_path = filepath
            img = Image.open(filepath)
            img.thumbnail((80, 40))
            img_tk = ImageTk.PhotoImage(img)
            self.logo_label.img_tk = img_tk
            self.logo_label.configure(image=img_tk)

    def visualizar_pdf(self):
        preview = self.gerar_preview_texto()
        win = Toplevel(self.root)
        win.title("Pré-visualização da Sugestão")
        text = tk.Text(win, wrap="word", font=("Arial", 10))
        text.pack(fill="both", expand=True)
        text.insert("1.0", preview)
        text.config(state="disabled")

    def visualizar_marcados(self):
        if not self.selected_indices:
            messagebox.showinfo("Info", "Nenhum item marcado.")
            return
        win = Toplevel(self.root)
        win.title("Itens Marcados")
        text = tk.Text(win, wrap="word", font=("Arial", 10))
        text.pack(fill="both", expand=True)
        for idx in self.selected_indices:
            row = self.data.loc[idx]
            text.insert("end", f"{row['cod']} - {row['descricao']} | {row['pais']} | {row['regiao']} | R$ {row['preco_base']:.2f}\n")
        text.config(state="disabled")

    def gerar_preview_texto(self):
        if not self.selected_indices:
            return "Nenhum item selecionado."
        df = self.data.loc[list(self.selected_indices)].copy()
        df['preco_de_venda'] = [self.manual_prices.get(idx, row.get('preco_de_venda', row['preco_base']*row.get('fator',2))) for idx, row in df.iterrows()]
        df['fator'] = [self.manual_fat.get(idx, row.get('fator',2.0)) for idx, row in df.iterrows()]
        df = df.sort_values(['tipo', 'pais', 'descricao'])
        preview = []
        preview.append("Sugestão Carta de Vinhos")
        if self.title_var.get():
            preview.append(f"Cliente: {self.title_var.get()}")
        preview.append("=" * 70)
        ordem_geral = 1
        contagem = {'Brancos':0, 'Tintos':0, 'Rosés':0, 'Espumantes':0, 'outros':0}
        tipo_map = {'branco':'Brancos', 'tinto':'Tintos', 'rose':'Rosés', 'rosé':'Rosés', 'espumante':'Espumantes'}
        for tipo in df['tipo'].dropna().unique():
            tipo_label = next((lbl for k,lbl in tipo_map.items() if k in str(tipo).lower()), 'outros')
            preview.append(f"\n{tipo.upper()}".ljust(60," "))
            for pais in df[df['tipo']==tipo]['pais'].dropna().unique():
                preview.append(f"  {pais.upper()}")
                grupo = df[(df['tipo'] == tipo) & (df['pais'] == pais)]
                for i, row in grupo.iterrows():
                    contagem[tipo_label] = contagem.get(tipo_label,0) + 1
                    desc = row['descricao']
                    preco = f"R$ {row['preco_base']:.2f}"
                    pvenda = f"R$ {row['preco_de_venda']:.2f}"
                    regiao = row['regiao']
                    cod = int(row['cod'])
                    preview.append(f"    {ordem_geral:02d} ({cod}) {desc}")
                    uvas = [str(row.get(f"uva{i}", "")).strip() for i in range(1,4)]
                    uvas_str = ", ".join([u for u in uvas if u.lower() != "nan" and u])
                    preview.append(f"      {row['pais']} | {regiao}" + (f" | {uvas_str}" if uvas_str else ""))
                    preview.append(f"      ({preco})  {pvenda}")
                    amad = str(row.get("amadurecimento", ""))
                    if amad and amad.lower() != "nan":
                        preview[-1] += " [🛢️]"
                    imgfile = self.get_imagem_file(str(row['cod']))
                    x_texto = 90
                    if imgfile:
                        preview.append("      [COM FOTO]")
                    ordem_geral += 1
        preview.append("\n" + "="*70)
        now = datetime.now().strftime("%d/%m/%Y %H:%M")
        preview.append(f"Gerado em: {now}")
        preview.append(
            f"Brancos: {contagem.get('Brancos',0)} | Tintos: {contagem.get('Tintos',0)} | "
            f"Rosés: {contagem.get('Rosés',0)} | Espumantes: {contagem.get('Espumantes',0)} | "
            f"Total: {ordem_geral-1} | Fator: {self.fator_var.get()}"
        )
        preview.append("Ingá Distribuidora Ltda | CNPJ 05.390.477/0002-25 Rod BR 232, KM 18,5 - S/N- Manassu - CEP 54130-340 Jaboatão - www.ingavinhos.com.br b2b.ingavinhos.com.br")
        return "\n".join(preview)

    def get_imagem_file(self, cod):
        img_path = os.path.join(r"C:/carta/imagens", f"{cod}.png")
        if os.path.exists(img_path):
            return img_path
        for ext in ['.png', '.jpg', '.jpeg', '.PNG', '.JPG', '.JPEG']:
            img_path = os.path.join(IMAGEM_DIR, f"{cod}{ext}")
            if os.path.exists(img_path):
                return os.path.abspath(img_path)
        for fname in os.listdir(IMAGEM_DIR):
            if fname.startswith(str(cod)):
                return os.path.abspath(os.path.join(IMAGEM_DIR, fname))
        return None

    def generate_pdf_layout(self):
        if not self.selected_indices:
            messagebox.showinfo("Atenção", "Selecione ao menos um vinho")
            return
        df = self.data.loc[list(self.selected_indices)].copy()
        df['preco_de_venda'] = [self.manual_prices.get(idx, row.get('preco_de_venda', row['preco_base']*row.get('fator',2))) for idx, row in df.iterrows()]
        df['fator'] = [self.manual_fat.get(idx, row.get('fator',2.0)) for idx, row in df.iterrows()]
        df = df.sort_values(['tipo', 'pais', 'descricao'])
        filename = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if filename:
            self.export_pdf(df, filename, "Sugestão Carta de Vinhos", self.title_var.get(), self.logo_path)
            messagebox.showinfo("PDF", "PDF gerado com sucesso!")

    def export_pdf(self, df, filename, titulo, cliente, logo_path=None):
        c = canvas.Canvas(filename, pagesize=A4)
        width, height = A4

        # Logo cliente (se houver)
        if logo_path and os.path.exists(logo_path):
            c.drawImage(ImageReader(logo_path), 40, height-60, width=120, height=40, mask='auto')
        # Logo Ingá SEMPRE no topo direito
        logo_inga_path = os.path.join(CARTA_DIR, "logo_inga.png")
        if os.path.exists(logo_inga_path):
            c.drawImage(logo_inga_path, width-80, height-40, width=48, height=24, mask='auto')

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
        tipo_map = {'branco':'Brancos', 'tinto':'Tintos', 'rose':'Rosés', 'rosé':'Rosés', 'espumante':'Espumantes'}
        for tipo in df['tipo'].dropna().unique():
            tipo_label = next((lbl for k,lbl in tipo_map.items() if k in str(tipo).lower()), 'outros')
            c.setFont("Helvetica-Bold", 10)
            c.drawString(x_texto, y, tipo.upper())
            y -= 14
            for pais in df[df['tipo']==tipo]['pais'].dropna().unique():
                c.setFont("Helvetica-Bold", 8)
                c.drawString(x_texto, y, pais.upper())
                y -= 12
                grupo = df[(df['tipo'] == tipo) & (df['pais'] == pais)]
                for i, row in grupo.iterrows():
                    contagem[tipo_label] = contagem.get(tipo_label,0) + 1
                    c.setFont("Helvetica", 6)
                    c.drawString(x_texto, y, f"{ordem_geral:02d} ({int(row['cod'])})")
                    c.setFont("Helvetica-Bold", 7)
                    c.drawString(x_texto+55, y, row['descricao'])
                    uvas = [str(row.get(f"uva{i}", "")).strip() for i in range(1,4)]
                    uvas_str = ", ".join([u for u in uvas if u.lower() != "nan" and u])
                    regiao_str = f"{row['pais']} | {row['regiao']}"
                    if uvas_str:
                        regiao_str += f" | {uvas_str}"
                    c.setFont("Helvetica", 5)
                    c.drawString(x_texto+55, y-10, regiao_str)
                    amad = str(row.get("amadurecimento", ""))
                    if amad and amad.lower() != "nan":
                        c.setFont("Helvetica", 7)
                        c.drawString(220, y-7, "🛢️")
                    c.setFont("Helvetica", 5)
                    c.drawRightString(width-120, y, f"(R$ {row['preco_base']:.2f})")
                    c.setFont("Helvetica-Bold", 7)
                    c.drawRightString(width-40, y, f"R$ {row['preco_de_venda']:.2f}")

                    imgfile = self.get_imagem_file(str(row['cod']))
                    if self.inserir_foto_var.get() and imgfile:
                        try:
                            c.drawImage(imgfile, x_texto+340, y-2, width=40, height=30, mask='auto')
                        except Exception:
                            pass
                        y -= 28
                    else:
                        y -= 20

                    ordem_geral += 1
                    if y < 100:
                        self.add_pdf_footer(c, contagem, ordem_geral-1, self.fator_var.get())
                        c.showPage()
                        y = height - 40
                        # Repete os cabeçalhos a cada nova página
                        if logo_path and os.path.exists(logo_path):
                            c.drawImage(ImageReader(logo_path), 40, height-60, width=120, height=40, mask='auto')
                        if os.path.exists(logo_inga_path):
                            c.drawImage(logo_inga_path, width-80, height-40, width=48, height=24, mask='auto')
                        c.setFont("Helvetica-Bold", 16)
                        c.drawCentredString(width/2, y, titulo)
                        y -= 20
                        if cliente:
                            c.setFont("Helvetica", 10)
                            c.drawCentredString(width/2, y, f"Cliente: {cliente}")
                            y -= 20
        self.add_pdf_footer(c, contagem, ordem_geral-1, self.fator_var.get())
        c.save()

    def add_pdf_footer(self, c, contagem, total_rotulos, fator_geral):
        width, height = A4
        y_rodape = 35
        now = datetime.now().strftime("%d/%m/%Y %H:%M")
        c.setLineWidth(0.4)
        c.line(30, y_rodape+32, width-30, y_rodape+32)
        c.setFont("Helvetica", 5)
        c.drawString(32, y_rodape+20, f"Gerado em: {now}")
        c.setFont("Helvetica-Bold", 6)
        c.drawString(32, y_rodape+7,
            f"Brancos: {contagem.get('Brancos',0)} | Tintos: {contagem.get('Tintos',0)} | "
            f"Rosés: {contagem.get('Rosés',0)} | Espumantes: {contagem.get('Espumantes',0)} | "
            f"Total: {total_rotulos} | Fator: {fator_geral}")
        c.setFont("Helvetica", 5)
        c.drawString(32, y_rodape-5, "Ingá Distribuidora Ltda | CNPJ 05.390.477/0002-25 Rod BR 232, KM 18,5 - S/N- Manassu - CEP 54130-340 Jaboatão")
        c.setFont("Helvetica-Bold", 6)
        c.drawString(width-190, y_rodape-5, "b2b.ingavinhos.com.br")

    def export_to_excel_layout(self):
        if not self.selected_indices:
            messagebox.showinfo("Atenção", "Selecione ao menos um vinho para exportar.")
            return
        df = self.data.loc[list(self.selected_indices)].copy()
        df['preco_de_venda'] = [self.manual_prices.get(idx, row.get('preco_de_venda', row['preco_base']*row.get('fator',2))) for idx, row in df.iterrows()]
        df['fator'] = [self.manual_fat.get(idx, row.get('fator',2.0)) for idx, row in df.iterrows()]
        df = df.sort_values(['tipo', 'pais', 'descricao'])

        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.export_to_excel_like_pdf_layout(df, filename)
            messagebox.showinfo("Excel", "Arquivo Excel criado com sucesso!")

    def export_to_excel_like_pdf_layout(self, df, filename):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sugestão"

        row_num = 1
        ordem_geral = 1

        tipos = df['tipo'].dropna().unique()
        for tipo in tipos:
            ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=8)
            cell = ws.cell(row=row_num, column=1, value=tipo.upper())
            cell.font = Font(bold=True, size=18)
            row_num += 1

            paises = df[df['tipo'] == tipo]['pais'].dropna().unique()
            for pais in paises:
                ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=8)
                cell = ws.cell(row=row_num, column=1, value=pais.upper())
                cell.font = Font(bold=True, size=14)
                row_num += 1

                grupo = df[(df['tipo'] == tipo) & (df['pais'] == pais)]
                for i, row in grupo.iterrows():
                    ws.cell(row=row_num, column=1, value=f"{ordem_geral:02d} ({int(row['cod'])})").font = Font(size=11)
                    ws.cell(row=row_num, column=2, value=row['descricao']).font = Font(bold=True, size=12)
                    imgfile = self.get_imagem_file(str(row['cod']))
                    x_texto = 90
                    if self.inserir_foto_var.get() and imgfile:
                        try:
                            img = XLImage(imgfile)
                            img.width, img.height = 32, 24
                            cell_ref = f"C{row_num}"
                            ws.add_image(img, cell_ref)
                        except Exception:
                            pass
                    ws.cell(row=row_num, column=7, value=f"(R$ {row['preco_base']:.2f})").alignment = Alignment(horizontal='right')
                    ws.cell(row=row_num, column=7).font = Font(size=10)
                    ws.cell(row=row_num, column=8, value=f"R$ {row['preco_de_venda']:.2f}").font = Font(bold=True, size=13)
                    ws.cell(row=row_num, column=8).alignment = Alignment(horizontal='right')
                    uvas = [str(row.get(f"uva{i}", "")).strip() for i in range(1,4)]
                    uvas_str = ", ".join([u for u in uvas if u.lower() != "nan" and u])
                    regiao_str = f"{row['pais']} | {row['regiao']}"
                    if uvas_str:
                        regiao_str += f" | {uvas_str}"
                    ws.cell(row=row_num+1, column=2, value=regiao_str).font = Font(size=10)
                    amad = str(row.get("amadurecimento", ""))
                    if amad and amad.lower() != "nan":
                        ws.cell(row=row_num+1, column=3, value="🛢️").font = Font(size=10)
                    row_num += 2
                    ordem_geral += 1

        ws.column_dimensions[get_column_letter(1)].width = 13
        ws.column_dimensions[get_column_letter(2)].width = 45
        ws.column_dimensions[get_column_letter(3)].width = 8
        ws.column_dimensions[get_column_letter(7)].width = 16
        ws.column_dimensions[get_column_letter(8)].width = 16

        wb.save(filename)

    # ========== ABA 2: SUGESTÕES SALVAS ==========
    def build_salvas_tab(self):
        f = self.frame_salvas
        self.listbox_salvas = tk.Listbox(f, width=40, font=("Arial", 10))
        self.listbox_salvas.pack(side='left', fill='y', padx=5, pady=5)
        self.listbox_salvas.bind('<<ListboxSelect>>', self.carregar_sugestao_lista)
        self.atualiza_listbox_salvas()
        btn_frame = ttk.Frame(f)
        btn_frame.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        ttk.Button(btn_frame, text="Excluir Sugestão", command=self.excluir_sugestao_lista).pack(pady=2)
        ttk.Button(btn_frame, text="Editar Itens", command=self.editar_sugestao_lista).pack(pady=2)
        ttk.Button(btn_frame, text="Adicionar Outros Itens", command=self.trazer_todos_apos_salva).pack(pady=2)
        ttk.Label(btn_frame, text="Ao editar, vá para aba Sugestão para salvar.", font=("Arial", 8)).pack(pady=3)

    def salvar_sugestao_dialog(self):
        nome = simpledialog.askstring("Salvar sugestão", "Nome da sugestão:")
        if not nome:
            return
        if not self.selected_indices:
            messagebox.showinfo("Sugestão", "Selecione produtos para salvar.")
            return
        indices = list(self.selected_indices)
        path = os.path.join(SUGESTOES_DIR, f"{nome}.txt")
        with open(path, "w") as f:
            f.write(",".join(map(str, indices)))
        messagebox.showinfo("Sugestão", f"Sugestão '{nome}' salva.")
        self.atualiza_listbox_salvas()

    def carregar_sugestao_lista(self, event=None):
        sel = self.listbox_salvas.curselection()
        if not sel:
            return
        nome = self.listbox_salvas.get(sel[0])
        path = os.path.join(SUGESTOES_DIR, f"{nome}.txt")
        if not os.path.exists(path):
            messagebox.showerror("Erro", "Arquivo da sugestão não encontrado.")
            return
        with open(path) as f:
            indices = [int(x) for x in f.read().strip().split(",") if x]
        self.filtered_data = self.data.loc[indices].copy()
        self.selected_indices = set(indices)
        self.notebook.select(self.frame_sugestao)
        self.refresh_table()

    def trazer_todos_apos_salva(self):
        self.filtered_data = self.data.copy()
        self.refresh_table()

    def excluir_sugestao_lista(self):
        sel = self.listbox_salvas.curselection()
        if not sel:
            return
        nome = self.listbox_salvas.get(sel[0])
        path = os.path.join(SUGESTOES_DIR, f"{nome}.txt")
        if os.path.exists(path):
            os.remove(path)
        self.atualiza_listbox_salvas()
        messagebox.showinfo("Sugestão", f"Sugestão '{nome}' excluída.")

    def editar_sugestao_lista(self):
        self.carregar_sugestao_lista()

    def atualiza_listbox_salvas(self):
        self.listbox_salvas.delete(0, tk.END)
        files = [f for f in os.listdir(SUGESTOES_DIR) if f.endswith(".txt")]
        for f in files:
            self.listbox_salvas.insert(tk.END, f[:-4])

    # ========== ABA 3: CADASTRO DE VINHOS ==========
    def build_cadastro_tab(self):
        f = self.frame_cadastro
        entry_frame = ttk.LabelFrame(f, text="Cadastrar Novo Produto")
        entry_frame.pack(fill='x', padx=5, pady=5)
        self.new_cod_var = tk.StringVar()
        self.new_desc_var = tk.StringVar()
        self.new_preco_var = tk.StringVar()
        self.new_fat_var = tk.StringVar(value="2.0")
        self.new_preco_venda_var = tk.StringVar()
        self.new_pais_var = tk.StringVar()
        self.new_regiao_var = tk.StringVar()
        ttk.Label(entry_frame, text="Código:", font=("Arial", 8)).grid(row=0, column=0, padx=2)
        ttk.Entry(entry_frame, textvariable=self.new_cod_var, width=8, font=("Arial", 8)).grid(row=0, column=1, padx=2)
        ttk.Label(entry_frame, text="Descrição:", font=("Arial", 8)).grid(row=0, column=2, padx=2)
        ttk.Entry(entry_frame, textvariable=self.new_desc_var, width=20, font=("Arial", 8)).grid(row=0, column=3, padx=2)
        ttk.Label(entry_frame, text="Preço:", font=("Arial", 8)).grid(row=0, column=4, padx=2)
        ttk.Entry(entry_frame, textvariable=self.new_preco_var, width=8, font=("Arial", 8)).grid(row=0, column=5, padx=2)
        ttk.Label(entry_frame, text="Fator:", font=("Arial", 8)).grid(row=0, column=6, padx=2)
        ttk.Entry(entry_frame, textvariable=self.new_fat_var, width=6, font=("Arial", 8)).grid(row=0, column=7, padx=2)
        ttk.Label(entry_frame, text="Preço Venda:", font=("Arial", 8)).grid(row=0, column=8, padx=2)
        ttk.Entry(entry_frame, textvariable=self.new_preco_venda_var, width=8, font=("Arial", 8)).grid(row=0, column=9, padx=2)
        ttk.Label(entry_frame, text="País:", font=("Arial", 8)).grid(row=0, column=10, padx=2)
        ttk.Entry(entry_frame, textvariable=self.new_pais_var, width=8, font=("Arial", 8)).grid(row=0, column=11, padx=2)
        ttk.Label(entry_frame, text="Região:", font=("Arial", 8)).grid(row=0, column=12, padx=2)
        ttk.Entry(entry_frame, textvariable=self.new_regiao_var, width=10, font=("Arial", 8)).grid(row=0, column=13, padx=2)
        ttk.Button(entry_frame, text="Cadastrar", command=self.cadastrar_produto).grid(row=0, column=14, padx=5)

        self.tree_cadastro = ttk.Treeview(f, columns=("cod","descricao","preco_base","fator","preco_de_venda","pais","regiao"), show="headings", height=10)
        for col, w in zip(("cod","descricao","preco_base","fator","preco_de_venda","pais","regiao"), [50,220,70,70,80,80,80]):
            self.tree_cadastro.heading(col, text=col.upper())
            self.tree_cadastro.column(col, width=w)
        self.tree_cadastro.pack(fill="both", expand=True, padx=5, pady=5)
        self.tree_cadastro.tag_configure('manual', font=('Arial', 8))
        ttk.Button(f, text="Excluir Produto Selecionado", command=self.excluir_produto).pack(pady=2)
        self.refresh_tree_cadastro()

    def refresh_tree_cadastro(self):
        for i in self.tree_cadastro.get_children():
            self.tree_cadastro.delete(i)
        for idx, row in self.data.iterrows():
            if row.get('cadastrado_manual', False):
                self.tree_cadastro.insert('', 'end', iid=idx, values=(
                    row.get('cod', ''),
                    row.get('descricao', ''),
                    row.get('preco_base', ''),
                    row.get('fator', ''),
                    row.get('preco_de_venda', ''),
                    row.get('pais', ''),
                    row.get('regiao', ''),
                ), tags=('manual',))

    def cadastrar_produto(self):
        try:
            cod = int(self.new_cod_var.get())
            desc = self.new_desc_var.get()
            preco = float(self.new_preco_var.get())
            fat = float(self.new_fat_var.get())
            preco_venda = float(self.new_preco_venda_var.get()) if self.new_preco_venda_var.get() else preco * fat
            pais = self.new_pais_var.get()
            regiao = self.new_regiao_var.get()
            if not desc or not pais or not regiao:
                raise Exception("Preencha todos os campos obrigatórios.")
            novo = pd.DataFrame([{
                "cod": cod,
                "descricao": desc,
                "preco_base": preco,
                "fator": fat,
                "preco_de_venda": preco_venda,
                "pais": pais,
                "regiao": regiao,
                "tipo": "",
                "cadastrado_manual": True
            }])
            self.data = pd.concat([self.data, novo], ignore_index=True)
            self.filtered_data = self.data.copy()
            self.refresh_tree_cadastro()
            self.refresh_table()
            messagebox.showinfo("Cadastro", "Produto cadastrado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao cadastrar: {e}")

    def excluir_produto(self):
        sel = self.tree_cadastro.selection()
        if not sel:
            messagebox.showinfo("Atenção", "Selecione um produto para excluir.")
            return
        idx = int(sel[0])
        cod_to_del = self.data.loc[idx, 'cod']
        self.data = self.data[self.data['cod'] != cod_to_del].copy()
        self.filtered_data = self.data.copy()
        self.refresh_tree_cadastro()
        self.refresh_table()
        messagebox.showinfo("Exclusão", "Produto removido com sucesso.")

if __name__ == "__main__":
    root = tk.Tk()
    app = WineMenuApp(root)
    root.mainloop()