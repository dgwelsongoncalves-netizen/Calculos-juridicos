import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side
from datetime import datetime
import subprocess
import sys

class NashDataEntryGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Nash System - Preenchimento (v3.0)")
        self.root.geometry("850x600")

        # Frame principal
        main_frame = ttk.Frame(root, padding=10)
        main_frame.pack(fill="both", expand=True)

        # Sistema de Separadores
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill="both", expand=True, pady=(0, 10))

        # Frames de cada separador
        self.tab_parametros = ttk.Frame(self.notebook, padding=10)
        self.tab_danos = ttk.Frame(self.notebook, padding=10)
        self.tab_custas = ttk.Frame(self.notebook, padding=10)

        # Anexar os Frames
        self.notebook.add(self.tab_parametros, text=" ⚙️ Parâmetros Gerais ")
        self.notebook.add(self.tab_danos, text=" 💰 Danos ")
        self.notebook.add(self.tab_custas, text=" ⚖️ Custas Processuais ")

        # Área de Ações Fixa
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill="x", side="bottom")

        btn_importar = ttk.Button(action_frame, text="📂 Importar Planilha Antiga")
        btn_importar.pack(side="left")

        btn_gerar = ttk.Button(action_frame, text="🚀 Salvar Template e Gerar Cálculo", command=self.salvar_e_gerar)
        btn_gerar.pack(side="right")

        self.montar_aba_parametros()
        self.montar_aba_danos()
        self.montar_aba_custas()

    def montar_aba_parametros(self):
        self.vars_param = {}
        campos = [
            ("Processo", "Entry", "5000000-00.2026.8.13.0024"),
            ("Atuação", "Combo", ["AUTOR", "RÉU"]),
            ("Data do Trânsito", "Entry", "30/08/2023"),
            ("Data da Sentença", "Entry", "18/07/2023"),
            ("Termo Inicial Juros", "Combo", ["Trânsito em julgado", "Desembolso", "Citação", "Evento"]),
            ("Data da Citação", "Entry", "19/05/2022"),
            ("Data do Evento", "Entry", "05/09/2016"),
            ("Justiça Gratuita", "Combo", ["Não", "Sim"]),
            ("Pagamento Voluntário 15d", "Combo", ["Sim", "Não", "Não se aplica"]),
            ("Fazenda Pública", "Combo", ["Não", "Sim"]),
            ("Honorários Sucumbência (%)", "Entry", "10"),
            ("Honorários Fixos (R$)", "Entry", ""),
            ("Base Honorários", "Combo", ["CONDENAÇÃO", "VALOR DA CAUSA", "PROVEITO ECONÔMICO"]),
            ("Valor Causa Original", "Entry", ""),
            ("Data Propositura", "Entry", ""),
            ("Proporção Honorários (%)", "Entry", "100%"),
            ("Proporção Custas (%)", "Entry", "100%")
        ]

        row = 0
        col = 0
        
        for nome, tipo, default in campos:
            var = tk.StringVar()
            self.vars_param[nome] = var
            
            lbl = ttk.Label(self.tab_parametros, text=f"{nome}:")
            lbl.grid(row=row, column=col, sticky="w", padx=10, pady=5)
            
            if tipo == "Entry":
                widget = ttk.Entry(self.tab_parametros, textvariable=var, width=25)
                if isinstance(default, str):
                    widget.insert(0, default)
            elif tipo == "Combo":
                widget = ttk.Combobox(self.tab_parametros, textvariable=var, values=default, state="readonly", width=22)
                widget.current(0)
                
            widget.grid(row=row, column=col+1, sticky="w", padx=10, pady=5)
            
            col += 2
            if col > 2:
                col = 0
                row += 1

    # --- ABA DANOS ---
    def montar_aba_danos(self):
        frame_botoes = ttk.Frame(self.tab_danos)
        frame_botoes.pack(fill="x", pady=(0, 10))

        ttk.Button(frame_botoes, text="➕ Adicionar Nova Linha", command=self.popup_adicionar_dano).pack(side="left", padx=5)
        ttk.Button(frame_botoes, text="🗑️ Remover Selecionada", command=self.remover_dano).pack(side="left", padx=5)
        ttk.Button(frame_botoes, text="🧹 Limpar Tudo", command=self.limpar_danos).pack(side="left", padx=5)

        colunas = ('id', 'descricao', 'data_desembolso', 'valor', 'regra', 'data_juros', 'valor_pedido', 'data_pedido')
        self.tree_danos = ttk.Treeview(self.tab_danos, columns=colunas, show='headings', height=15)

        titulos = ['ID / Folha', 'Descrição', 'Data Desembolso', 'Valor Histórico', 'Regra', 'Data Juros', 'Valor Ped. Inicial', 'Data do Pedido']
        for col, titulo in zip(colunas, titulos):
            self.tree_danos.heading(col, text=titulo)
            self.tree_danos.column(col, width=110, anchor="center")
        
        self.tree_danos.column('descricao', width=250, anchor="w")

        scroll_y = ttk.Scrollbar(self.tab_danos, orient="vertical", command=self.tree_danos.yview)
        scroll_x = ttk.Scrollbar(self.tab_danos, orient="horizontal", command=self.tree_danos.xview)
        self.tree_danos.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        self.tree_danos.pack(fill="both", expand=True)

    def popup_adicionar_dano(self):
        popup = tk.Toplevel(self.root)
        popup.title("Adicionar Dano")
        popup.geometry("380x350")
        popup.transient(self.root)
        popup.grab_set()

        campos = ['ID / Folha', 'Descrição', 'Data Desembolso', 'Valor Histórico', 'Regra', 'Data Juros', 'Valor Ped. Inicial', 'Data do Pedido']
        entradas = {}

        for i, campo in enumerate(campos):
            ttk.Label(popup, text=f"{campo}:").grid(row=i, column=0, padx=10, pady=5, sticky="w")
            ent = ttk.Entry(popup, width=30)
            ent.grid(row=i, column=1, padx=10, pady=5)
            entradas[campo] = ent

        def salvar_dano():
            valores = [entradas[c].get() for c in campos]
            self.tree_danos.insert('', 'end', values=valores)
            popup.destroy()

        ttk.Button(popup, text="Salvar Linha", command=salvar_dano).grid(row=len(campos), column=0, columnspan=2, pady=15)

    def remover_dano(self):
        for item in self.tree_danos.selection():
            self.tree_danos.delete(item)

    def limpar_danos(self):
        for item in self.tree_danos.get_children():
            self.tree_danos.delete(item)

    # --- ABA CUSTAS ---
    def montar_aba_custas(self):
        frame_botoes = ttk.Frame(self.tab_custas)
        frame_botoes.pack(fill="x", pady=(0, 10))

        ttk.Button(frame_botoes, text="➕ Adicionar Nova Custa", command=self.popup_adicionar_custa).pack(side="left", padx=5)
        ttk.Button(frame_botoes, text="🗑️ Remover Selecionada", command=self.remover_custa).pack(side="left", padx=5)
        ttk.Button(frame_botoes, text="🧹 Limpar Tudo", command=self.limpar_custas).pack(side="left", padx=5)

        colunas = ('id', 'descricao', 'data_desembolso', 'valor')
        self.tree_custas = ttk.Treeview(self.tab_custas, columns=colunas, show='headings', height=15)

        titulos = ['ID / Folha', 'Descrição', 'Data Desembolso', 'Valor Histórico']
        for col, titulo in zip(colunas, titulos):
            self.tree_custas.heading(col, text=titulo)
            self.tree_custas.column(col, width=150, anchor="center")
        
        self.tree_custas.column('descricao', width=350, anchor="w")

        scroll_y = ttk.Scrollbar(self.tab_custas, orient="vertical", command=self.tree_custas.yview)
        scroll_x = ttk.Scrollbar(self.tab_custas, orient="horizontal", command=self.tree_custas.xview)
        self.tree_custas.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        self.tree_custas.pack(fill="both", expand=True)

    def popup_adicionar_custa(self):
        popup = tk.Toplevel(self.root)
        popup.title("Adicionar Custa")
        popup.geometry("380x200")
        popup.transient(self.root)
        popup.grab_set()

        campos = ['ID / Folha', 'Descrição', 'Data Desembolso', 'Valor Histórico']
        entradas = {}

        for i, campo in enumerate(campos):
            ttk.Label(popup, text=f"{campo}:").grid(row=i, column=0, padx=10, pady=5, sticky="w")
            ent = ttk.Entry(popup, width=30)
            ent.grid(row=i, column=1, padx=10, pady=5)
            entradas[campo] = ent

        def salvar_custa():
            valores = [entradas[c].get() for c in campos]
            self.tree_custas.insert('', 'end', values=valores)
            popup.destroy()

        ttk.Button(popup, text="Salvar Custa", command=salvar_custa).grid(row=len(campos), column=0, columnspan=2, pady=15)

    def remover_custa(self):
        for item in self.tree_custas.selection():
            self.tree_custas.delete(item)

    def limpar_custas(self):
        for item in self.tree_custas.get_children():
            self.tree_custas.delete(item)

    # --- GERADOR E INTEGRADOR FINAL ---
    def salvar_e_gerar(self):
        try:
            # 1. Padronização dos Nomes dos Arquivos
            processo_bruto = self.vars_param["Processo"].get().strip()
            # Remove caracteres que o Windows proíbe em nomes de arquivos
            processo_seguro = "".join(c for c in processo_bruto if c not in r'\/:*?"<>|')
            
            data_atual = datetime.now().strftime("%d.%m.%Y")
            
            # Aqui estão as variáveis que tinham ficado de fora!
            nome_template = f"template {processo_seguro} {data_atual}.xlsx"
            nome_laudo_xls = f"laudo {processo_seguro} {data_atual}.xlsx"
            nome_laudo_pdf = f"laudo {processo_seguro} {data_atual}.pdf"

            # 2. Criação do Excel
            wb = Workbook()
            f_bold = Font(bold=True)
            borda = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            fill_cinza = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

            # ABA PARAMETROS
            ws_param = wb.active
            ws_param.title = "Parametros"
            
            for r_idx, (chave, var_tk) in enumerate(self.vars_param.items(), 1):
                valor = var_tk.get()
                c1 = ws_param.cell(row=r_idx, column=1, value=chave)
                c1.font = f_bold; c1.fill = fill_cinza; c1.border = borda
                
                c2 = ws_param.cell(row=r_idx, column=2, value=valor)
                c2.border = borda

            ws_param.column_dimensions['A'].width = 25
            ws_param.column_dimensions['B'].width = 30

            # ABA DANOS
            ws_danos = wb.create_sheet('Danos')
            headers_danos = ['ID / Folha', 'Descrição', 'Data Desembolso', 'Valor Histórico', 'Regra', 'Data Juros', 'Valor Pedido Inicial', 'Data do Pedido']
            for col_idx, h in enumerate(headers_danos, 1):
                cell = ws_danos.cell(row=1, column=col_idx, value=h)
                cell.font = f_bold; cell.fill = fill_cinza; cell.border = borda
                ws_danos.column_dimensions[chr(64+col_idx)].width = 20
            
            for row_idx, item in enumerate(self.tree_danos.get_children(), 2):
                valores = self.tree_danos.item(item, 'values')
                for col_idx, val in enumerate(valores, 1):
                    ws_danos.cell(row=row_idx, column=col_idx, value=val).border = borda

            # ABA CUSTAS
            ws_custas = wb.create_sheet('Custas')
            headers_custas = ['ID / Folha', 'Descrição', 'Data Desembolso', 'Valor Histórico']
            for col_idx, h in enumerate(headers_custas, 1):
                cell = ws_custas.cell(row=1, column=col_idx, value=h)
                cell.font = f_bold; cell.fill = fill_cinza; cell.border = borda
                ws_custas.column_dimensions[chr(64+col_idx)].width = 22
                
            for row_idx, item in enumerate(self.tree_custas.get_children(), 2):
                valores = self.tree_custas.item(item, 'values')
                for col_idx, val in enumerate(valores, 1):
                    ws_custas.cell(row=row_idx, column=col_idx, value=val).border = borda

            # ABA DEDUCOES
            ws_deducoes = wb.create_sheet('Deducoes')
            headers_deducoes = ['ID / Folha', 'Data bloqueio/deposito', 'Valor']
            for col_idx, h in enumerate(headers_deducoes, 1):
                cell = ws_deducoes.cell(row=1, column=col_idx, value=h)
                cell.font = f_bold; cell.fill = fill_cinza; cell.border = borda
                ws_deducoes.column_dimensions[chr(64+col_idx)].width = 25

            # Salva o arquivo final com o nome dinâmico
            wb.save(nome_template)
            
            # Chama o motor matemático capturando TUDO o que acontecer lá dentro
            resultado = subprocess.run(
                [sys.executable, "nash.py", nome_template, nome_laudo_xls, nome_laudo_pdf],
                capture_output=True, text=True
            )
            
            # Se o nash.py craschar (retornar código diferente de zero), ele exibe o erro real na tela
            if resultado.returncode != 0:
                messagebox.showerror(
                    "Erro no Motor Matemático", 
                    f"O nash.py quebrou nos bastidores!\n\nDetalhes do erro:\n{resultado.stderr}"
                )
            else:
                messagebox.showinfo(
                    "Cálculo Concluído", 
                    f"Processamento finalizado com sucesso!\n\nArquivos gerados na pasta:\n- {nome_laudo_xls}\n- {nome_laudo_pdf}"
                )
                
        except Exception as e:
            messagebox.showerror("Erro na Interface", f"Falha ao gerar o Excel da interface:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    
    style = ttk.Style()
    if 'clam' in style.theme_names():
        style.theme_use('clam')
        
    app = NashDataEntryGUI(root)
    root.mainloop()