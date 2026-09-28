import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side
from datetime import datetime
import subprocess
import sys
from pathlib import Path
import threading
import nash

class NashDataEntryGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Nash System - Preenchimento (v4.0 - Beta Acordos)")
        self.root.geometry("900x650")

        main_frame = ttk.Frame(root, padding=10)
        main_frame.pack(fill="both", expand=True)

        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill="both", expand=True, pady=(0, 10))

        # Criação das Abas
        self.tab_parametros = ttk.Frame(self.notebook, padding=10)
        self.tab_danos = ttk.Frame(self.notebook, padding=10)
        self.tab_custas = ttk.Frame(self.notebook, padding=10)
        self.tab_deducoes = ttk.Frame(self.notebook, padding=10)
        self.tab_acordos = ttk.Frame(self.notebook, padding=10)

        self.notebook.add(self.tab_parametros, text=" ⚙️ Parâmetros Gerais ")
        self.notebook.add(self.tab_danos, text=" 💰 Danos Principais ")
        self.notebook.add(self.tab_custas, text=" ⚖️ Custas ")
        self.notebook.add(self.tab_deducoes, text=" 📉 Deduções (Conta Gráfica) ")
        self.notebook.add(self.tab_acordos, text=" 🤝 Acordo Não Cumprido ")

        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill="x", side="bottom")

        btn_importar = ttk.Button(action_frame, text="📂 Importar Planilha Antiga", command=self.importar_planilha)
        btn_importar.pack(side="left")

        self.btn_processar = ttk.Button(action_frame, text="🚀 Salvar Template e Gerar Cálculo", command=self.salvar_e_gerar)
        self.btn_processar.pack(side="right")

        # Opções de Regras Globais
        self.opcoes_regras = [
            "R1 - TJMG + Juros de 1% a.m.",
            "R2 - Taxa Selic (critério único)",
            "R3 - TJMG + Juros 1% a.m. até 08/24; após, Selic",
            "R4 - TJMG + Juros 1% a.m. até 08/24; após, Lei 14.905/24",
            "R5 - Selic até 08/24; após, Lei 14.905/24 (Tema 1.368 STJ)",
            "R6 - Lei 14.905/24: IPCA + Taxa Legal",
            "R7 - TJMG CM + Taxa Legal Retroativa"
        ]

        self.montar_aba_parametros()
        self.montar_aba_danos()
        self.montar_aba_custas()
        self.montar_aba_deducoes()
        self.montar_aba_acordos()

    # ================= ABA 1: PARÂMETROS =================
    def montar_aba_parametros(self):
        self.vars_param = {}
        campos = [
            ("Processo", "Entry", "5101247-58.2022.8.13.0024"),
            ("Atuação", "Combo", ["RÉU", "AUTOR"]),
            ("Data do Trânsito", "Entry", ""),
            ("Data da Sentença", "Entry", "19/03/2026"),
            ("Termo Inicial Juros", "Combo", ["Propositura", "Trânsito em julgado", "Desembolso", "Citação", "Evento"]),
            ("Data da Citação", "Entry", ""),
            ("Data do Evento", "Entry", ""),
            ("Justiça Gratuita", "Combo", ["Não", "Sim"]),
            ("Pagamento Voluntário 15d", "Combo", ["Não se aplica", "Sim", "Não"]),
            ("Fazenda Pública", "Combo", ["Não", "Sim"]),
            ("Honorários Sucumbência (%)", "Entry", "10"),
            ("Honorários Fixos (R$)", "Entry", ""),
            ("Base Honorários", "Combo", ["CONDENAÇÃO", "VALOR DA CAUSA", "PROVEITO ECONÔMICO"]),
            ("Valor Causa Original", "Entry", "34620"),
            ("Data Propositura", "Entry", "26/05/2022"),
            ("Proporção Honorários (%)", "Entry", "100%"),
            ("Proporção Custas (%)", "Entry", "100%")
        ]

        row = col = 0
        for nome, tipo, default in campos:
            var = tk.StringVar()
            self.vars_param[nome] = var
            lbl = ttk.Label(self.tab_parametros, text=f"{nome}:")
            lbl.grid(row=row, column=col, sticky="w", padx=10, pady=5)
            
            if tipo == "Entry":
                widget = ttk.Entry(self.tab_parametros, textvariable=var, width=25)
                if isinstance(default, str): widget.insert(0, default)
            elif tipo == "Combo":
                widget = ttk.Combobox(self.tab_parametros, textvariable=var, values=default, state="readonly", width=22)
                widget.current(0)
                
            widget.grid(row=row, column=col+1, sticky="w", padx=10, pady=5)
            col += 2
            if col > 2:
                col = 0
                row += 1

    # ================= ABA 2: DANOS =================
    def montar_aba_danos(self):
        frame_botoes = ttk.Frame(self.tab_danos)
        frame_botoes.pack(fill="x", pady=(0, 10))

        ttk.Button(frame_botoes, text="➕ Adicionar", command=lambda: self.popup_generico('danos')).pack(side="left", padx=5)
        ttk.Button(frame_botoes, text="✏️ Editar", command=lambda: self.editar_selecionado('danos')).pack(side="left", padx=5)
        ttk.Button(frame_botoes, text="🗑️ Remover", command=lambda: self.remover_selecionado('danos')).pack(side="left", padx=5)
        ttk.Button(frame_botoes, text="🧹 Limpar", command=lambda: self.limpar_tree('danos')).pack(side="left", padx=5)

        colunas = ('id', 'descricao', 'data_desembolso', 'valor', 'regra', 'data_juros', 'valor_pedido', 'data_pedido')
        self.tree_danos = ttk.Treeview(self.tab_danos, columns=colunas, show='headings', height=12)
        titulos = ['ID / Folha', 'Descrição', 'Data Desembolso', 'Valor Histórico', 'Regra', 'Data Juros', 'Valor Ped. Inicial', 'Data do Pedido']
        for col, titulo in zip(colunas, titulos):
            self.tree_danos.heading(col, text=titulo)
            self.tree_danos.column(col, width=100, anchor="center")
        self.tree_danos.column('descricao', width=220, anchor="w")

        scroll_y = ttk.Scrollbar(self.tab_danos, orient="vertical", command=self.tree_danos.yview)
        scroll_x = ttk.Scrollbar(self.tab_danos, orient="horizontal", command=self.tree_danos.xview)
        self.tree_danos.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        self.tree_danos.pack(fill="both", expand=True)

    # ================= ABA 3: CUSTAS =================
    def montar_aba_custas(self):
        frame_botoes = ttk.Frame(self.tab_custas)
        frame_botoes.pack(fill="x", pady=(0, 10))

        ttk.Button(frame_botoes, text="➕ Adicionar", command=lambda: self.popup_generico('custas')).pack(side="left", padx=5)
        ttk.Button(frame_botoes, text="✏️ Editar", command=lambda: self.editar_selecionado('custas')).pack(side="left", padx=5)
        ttk.Button(frame_botoes, text="🗑️ Remover", command=lambda: self.remover_selecionado('custas')).pack(side="left", padx=5)
        ttk.Button(frame_botoes, text="🧹 Limpar", command=lambda: self.limpar_tree('custas')).pack(side="left", padx=5)

        colunas = ('id', 'descricao', 'data_desembolso', 'valor')
        self.tree_custas = ttk.Treeview(self.tab_custas, columns=colunas, show='headings', height=12)
        titulos = ['ID / Folha', 'Descrição', 'Data Desembolso', 'Valor Histórico']
        for col, titulo in zip(colunas, titulos):
            self.tree_custas.heading(col, text=titulo)
            self.tree_custas.column(col, width=150, anchor="center")
        self.tree_custas.column('descricao', width=350, anchor="w")

        scroll_y = ttk.Scrollbar(self.tab_custas, orient="vertical", command=self.tree_custas.yview)
        self.tree_custas.configure(yscrollcommand=scroll_y.set)
        scroll_y.pack(side="right", fill="y")
        self.tree_custas.pack(fill="both", expand=True)

    # ================= ABA 4: DEDUÇÕES (CONTA GRÁFICA) =================
    def montar_aba_deducoes(self):
        frame_botoes = ttk.Frame(self.tab_deducoes)
        frame_botoes.pack(fill="x", pady=(0, 10))

        ttk.Button(frame_botoes, text="➕ Adicionar Dedução", command=lambda: self.popup_generico('deducoes')).pack(side="left", padx=5)
        ttk.Button(frame_botoes, text="✏️ Editar", command=lambda: self.editar_selecionado('deducoes')).pack(side="left", padx=5)
        ttk.Button(frame_botoes, text="🗑️ Remover", command=lambda: self.remover_selecionado('deducoes')).pack(side="left", padx=5)
        ttk.Button(frame_botoes, text="🧹 Limpar", command=lambda: self.limpar_tree('deducoes')).pack(side="left", padx=5)

        tk.Label(frame_botoes, text="(Lançar Bloqueios Sisbajud, Depósitos Judiciais ou Alvarás)", fg="gray", font=("Arial", 9, "italic")).pack(side="right", padx=10)

        colunas = ('id', 'data_bloqueio', 'valor')
        self.tree_deducoes = ttk.Treeview(self.tab_deducoes, columns=colunas, show='headings', height=12)
        titulos = ['ID / Evento / Origem', 'Data do Bloqueio/Depósito', 'Valor (R$)']
        for col, titulo in zip(colunas, titulos):
            self.tree_deducoes.heading(col, text=titulo)
            self.tree_deducoes.column(col, width=200, anchor="center")
        self.tree_deducoes.column('id', width=400, anchor="w")

        scroll_y = ttk.Scrollbar(self.tab_deducoes, orient="vertical", command=self.tree_deducoes.yview)
        self.tree_deducoes.configure(yscrollcommand=scroll_y.set)
        scroll_y.pack(side="right", fill="y")
        self.tree_deducoes.pack(fill="both", expand=True)

    # ================= ABA 5: ACORDO NÃO CUMPRIDO (NOVO MÓDULO) =================
    def montar_aba_acordos(self):
        # Frame Superior: Parâmetros do Acordo
        frame_params = tk.LabelFrame(self.tab_acordos, text=" 📝 Parâmetros da Quebra de Acordo ", font=("Arial", 10, "bold"), padx=10, pady=10)
        frame_params.pack(fill="x", pady=(0, 10))

        self.vars_acordo = {}
        campos_acordo = [
            ("Data do Inadimplemento", "Entry", ""),
            ("Regra de Atualização Pós-Quebra", "Combo", self.opcoes_regras),
            ("Valor Original Confessado (Sem Desconto)", "Entry", ""),
            ("Valor do Desconto (A Reverter)", "Entry", ""),
            ("Multa Moratória Contratual (%)", "Entry", "10"),
            ("Honorários de Retomada/Execução (%)", "Entry", "20")
        ]

        row = col = 0
        for nome, tipo, default in campos_acordo:
            var = tk.StringVar()
            self.vars_acordo[nome] = var
            lbl = ttk.Label(frame_params, text=f"{nome}:")
            lbl.grid(row=row, column=col, sticky="w", padx=10, pady=5)
            
            if tipo == "Entry":
                widget = ttk.Entry(frame_params, textvariable=var, width=25)
                if isinstance(default, str): widget.insert(0, default)
            elif tipo == "Combo":
                widget = ttk.Combobox(frame_params, textvariable=var, values=default, state="readonly", width=38)
                widget.current(1) # Default R2
                
            widget.grid(row=row, column=col+1, sticky="w", padx=10, pady=5)
            col += 2
            if col > 2:
                col = 0
                row += 1

        # Frame Inferior: Parcelas Pagas
        frame_pagas = tk.LabelFrame(self.tab_acordos, text=" 💸 Parcelas Pagas (A Abater do Saldo) ", font=("Arial", 10, "bold"), padx=10, pady=5)
        frame_pagas.pack(fill="both", expand=True)

        frame_botoes_ac = ttk.Frame(frame_pagas)
        frame_botoes_ac.pack(fill="x", pady=(0, 5))
        ttk.Button(frame_botoes_ac, text="➕ Adicionar Parcela Paga", command=lambda: self.popup_generico('acordo_pagas')).pack(side="left", padx=5)
        ttk.Button(frame_botoes_ac, text="✏️ Editar", command=lambda: self.editar_selecionado('acordo_pagas')).pack(side="left", padx=5)
        ttk.Button(frame_botoes_ac, text="🗑️ Remover", command=lambda: self.remover_selecionado('acordo_pagas')).pack(side="left", padx=5)

        colunas_ac = ('id', 'data_pagamento', 'valor')
        self.tree_acordo_pagas = ttk.Treeview(frame_pagas, columns=colunas_ac, show='headings', height=6)
        titulos_ac = ['Identificação / Parcela', 'Data do Pagamento', 'Valor Pago (R$)']
        for col, titulo in zip(colunas_ac, titulos_ac):
            self.tree_acordo_pagas.heading(col, text=titulo)
            self.tree_acordo_pagas.column(col, width=200, anchor="center")
        self.tree_acordo_pagas.column('id', width=400, anchor="w")

        scroll_y_ac = ttk.Scrollbar(frame_pagas, orient="vertical", command=self.tree_acordo_pagas.yview)
        self.tree_acordo_pagas.configure(yscrollcommand=scroll_y_ac.set)
        scroll_y_ac.pack(side="right", fill="y")
        self.tree_acordo_pagas.pack(fill="both", expand=True)


    # ================= LÓGICA DE POPUPS UNIFICADA =================
    def popup_generico(self, tipo_tree, item_selecionado=None):
        popup = tk.Toplevel(self.root)
        popup.transient(self.root)
        popup.grab_set()

        if tipo_tree == 'danos':
            popup.title("Editar Dano" if item_selecionado else "Adicionar Dano")
            campos = ['ID / Folha', 'Descrição', 'Data Desembolso', 'Valor Histórico', 'Regra', 'Data Juros', 'Valor Ped. Inicial', 'Data do Pedido']
            tree_alvo = self.tree_danos
        elif tipo_tree == 'custas':
            popup.title("Editar Custa" if item_selecionado else "Adicionar Custa")
            campos = ['ID / Folha', 'Descrição', 'Data Desembolso', 'Valor Histórico']
            tree_alvo = self.tree_custas
        elif tipo_tree == 'deducoes':
            popup.title("Editar Dedução" if item_selecionado else "Adicionar Depósito/Bloqueio")
            campos = ['ID / Folha', 'Data bloqueio/deposito', 'Valor']
            tree_alvo = self.tree_deducoes
        elif tipo_tree == 'acordo_pagas':
            popup.title("Editar Parcela Paga" if item_selecionado else "Adicionar Parcela Paga")
            campos = ['Identificação / Parcela', 'Data do Pagamento', 'Valor Pago (R$)']
            tree_alvo = self.tree_acordo_pagas

        entradas = {}
        valores_atuais = tree_alvo.item(item_selecionado, 'values') if item_selecionado else [''] * len(campos)

        for i, campo in enumerate(campos):
            ttk.Label(popup, text=f"{campo}:").grid(row=i, column=0, padx=10, pady=5, sticky="w")
            
            if campo == 'Regra':
                ent = ttk.Combobox(popup, values=self.opcoes_regras, state="readonly", width=37)
                val_atual = str(valores_atuais[i]) if valores_atuais[i] else ""
                encontrou = False
                for opt in self.opcoes_regras:
                    if opt.startswith(val_atual):
                        ent.set(opt)
                        encontrou = True; break
                if not encontrou: ent.current(1)
            else:
                ent = ttk.Entry(popup, width=40)
                if valores_atuais[i]: ent.insert(0, valores_atuais[i])
                
            ent.grid(row=i, column=1, padx=10, pady=5)
            entradas[campo] = ent

        def salvar():
            novos_valores = []
            for c in campos:
                val = entradas[c].get()
                if c == 'Regra' and " - " in val: val = val.split(" - ")[0]
                novos_valores.append(val)
                
            if item_selecionado: tree_alvo.item(item_selecionado, values=novos_valores)
            else: tree_alvo.insert('', 'end', values=novos_valores)
            popup.destroy()

        ttk.Button(popup, text="Salvar Alterações" if item_selecionado else "Salvar Linha", command=salvar).grid(row=len(campos), column=0, columnspan=2, pady=15)

    def editar_selecionado(self, tipo_tree):
        tree = getattr(self, f"tree_{tipo_tree}")
        selecionados = tree.selection()
        if not selecionados:
            messagebox.showwarning("Aviso", "Selecione uma linha para editar.")
            return
        self.popup_generico(tipo_tree, item_selecionado=selecionados[0])

    def remover_selecionado(self, tipo_tree):
        tree = getattr(self, f"tree_{tipo_tree}")
        for item in tree.selection(): tree.delete(item)

    def limpar_tree(self, tipo_tree):
        tree = getattr(self, f"tree_{tipo_tree}")
        for item in tree.get_children(): tree.delete(item)

    # ================= EXPORTAÇÃO E IMPORTAÇÃO =================
    def importar_planilha(self):
        caminho = filedialog.askopenfilename(title="Selecione o Template Antigo", filetypes=[("Planilha Excel", "*.xlsx")])
        if not caminho: return
            
        try:
            wb = load_workbook(caminho, data_only=True)
            
            # Importa Parâmetros
            if 'Parametros' in wb.sheetnames:
                ws = wb['Parametros']
                for row in ws.iter_rows(values_only=True):
                    if row[0]:
                        chave = str(row[0]).strip()
                        valor = row[1] if row[1] is not None else ""
                        if isinstance(valor, datetime): valor = valor.strftime("%d/%m/%Y")
                        if chave in self.vars_param: self.vars_param[chave].set(str(valor))
                            
            # Função genérica para importar trees
            def carregar_tree(sheet_name, tree_widget, cols_esperadas):
                if sheet_name in wb.sheetnames:
                    for item in tree_widget.get_children(): tree_widget.delete(item)
                    ws = wb[sheet_name]
                    headers = [cell.value for cell in ws[1]]
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        if not any(row): continue
                        row_data = dict(zip(headers, row))
                        valores = []
                        for col in cols_esperadas:
                            val = row_data.get(col, "")
                            val = val if val is not None else ""
                            if isinstance(val, datetime): val = val.strftime("%d/%m/%Y")
                            valores.append(str(val))
                        if valores[1]: tree_widget.insert('', 'end', values=valores)

            carregar_tree('Danos', self.tree_danos, ['ID / Folha', 'Descrição', 'Data Desembolso', 'Valor Histórico', 'Regra', 'Data Juros', 'Valor Pedido Inicial', 'Data do Pedido'])
            carregar_tree('Custas', self.tree_custas, ['ID / Folha', 'Descrição', 'Data Desembolso', 'Valor Histórico'])
            carregar_tree('Deducoes', self.tree_deducoes, ['ID / Folha', 'Data bloqueio/deposito', 'Valor'])
            
            # Importar Acordos
            if 'Acordo_Params' in wb.sheetnames:
                ws = wb['Acordo_Params']
                for row in ws.iter_rows(values_only=True):
                    if row[0]:
                        chave = str(row[0]).strip()
                        valor = row[1] if row[1] is not None else ""
                        if isinstance(valor, datetime): valor = valor.strftime("%d/%m/%Y")
                        if chave in self.vars_acordo: self.vars_acordo[chave].set(str(valor))
                        
            carregar_tree('Acordo_Pagas', self.tree_acordo_pagas, ['Identificação / Parcela', 'Data do Pagamento', 'Valor Pago (R$)'])
            
            messagebox.showinfo("Sucesso", "Planilha e histórico carregados com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro ao Importar", f"Ocorreu um erro ao tentar ler a planilha:\n{str(e)}")

    def salvar_e_gerar(self):
        try:
            processo_bruto = self.vars_param["Processo"].get().strip()
            processo_seguro = "".join(c for c in processo_bruto if c not in r'\/:*?"<>|')
            data_atual = datetime.now().strftime("%d.%m.%Y")
            
            caminho_destino = filedialog.asksaveasfilename(
                title="Salvar Laudo de Liquidação",
                defaultextension=".xlsx",
                initialfile=f"laudo {processo_seguro} {data_atual}.xlsx",
                filetypes=[("Planilha Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
            )
            
            if not caminho_destino: return
                
            path_destino = Path(caminho_destino)
            pasta_destino = path_destino.parent
            nome_base = path_destino.stem
            nome_template = pasta_destino / f"template {processo_seguro} {data_atual}.xlsx"
            nome_laudo_xls = pasta_destino / f"{nome_base}.xlsx"

            wb = Workbook()
            f_bold = Font(bold=True)
            borda = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            fill_cinza = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

            def salvar_params(aba, dicionario):
                ws = wb.create_sheet(aba) if aba != 'Parametros' else wb.active
                ws.title = aba
                for r_idx, (chave, var_tk) in enumerate(dicionario.items(), 1):
                    val = var_tk.get()
                    if 'Regra' in chave and ' - ' in val: val = val.split(' - ')[0]
                    c1 = ws.cell(row=r_idx, column=1, value=chave)
                    c1.font = f_bold; c1.fill = fill_cinza; c1.border = borda
                    ws.cell(row=r_idx, column=2, value=val).border = borda
                ws.column_dimensions['A'].width = 35
                ws.column_dimensions['B'].width = 30

            def salvar_tree(aba, tree_widget, headers):
                ws = wb.create_sheet(aba)
                for col_idx, h in enumerate(headers, 1):
                    cell = ws.cell(row=1, column=col_idx, value=h)
                    cell.font = f_bold; cell.fill = fill_cinza; cell.border = borda
                    ws.column_dimensions[chr(64+col_idx)].width = 20
                for row_idx, item in enumerate(tree_widget.get_children(), 2):
                    for col_idx, val in enumerate(tree_widget.item(item, 'values'), 1):
                        ws.cell(row=row_idx, column=col_idx, value=val).border = borda

            salvar_params('Parametros', self.vars_param)
            salvar_tree('Danos', self.tree_danos, ['ID / Folha', 'Descrição', 'Data Desembolso', 'Valor Histórico', 'Regra', 'Data Juros', 'Valor Pedido Inicial', 'Data do Pedido'])
            salvar_tree('Custas', self.tree_custas, ['ID / Folha', 'Descrição', 'Data Desembolso', 'Valor Histórico'])
            salvar_tree('Deducoes', self.tree_deducoes, ['ID / Folha', 'Data bloqueio/deposito', 'Valor'])
            salvar_params('Acordo_Params', self.vars_acordo)
            salvar_tree('Acordo_Pagas', self.tree_acordo_pagas, ['Identificação / Parcela', 'Data do Pagamento', 'Valor Pago (R$)'])

            wb.save(nome_template)
            
            self.btn_processar.config(state="disabled")
            
            def thread_calculo():
                try:
                    nash.executar_nash(str(nome_template), str(nome_laudo_xls))
                    def sucesso():
                        messagebox.showinfo("Cálculo Concluído", f"Processamento finalizado com sucesso!\n\nArquivos guardados em:\n{pasta_destino}")
                        self.btn_processar.config(state="normal")
                    self.root.after(0, sucesso)
                except Exception as e:
                    erro_msg = str(e)
                    def falha():
                        messagebox.showerror("Erro no Motor Matemático", f"O cálculo falhou:\n\n{erro_msg}\n\nNota: Se for um Acordo Não Cumprido, o motor matemático correspondente será ativado amanhã.")
                        self.btn_processar.config(state="normal")
                    self.root.after(0, falha)

            threading.Thread(target=thread_calculo, daemon=True).start()
            
        except Exception as e:
            messagebox.showerror("Erro Crítico", f"Falha ao gerar o Excel:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style()
    if 'clam' in style.theme_names(): style.theme_use('clam')
    app = NashDataEntryGUI(root)
    root.mainloop()