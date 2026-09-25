import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side
from datetime import datetime
import subprocess
import sys
from pathlib import Path

class NashDataEntryGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Nash System - Preenchimento (v3.1)")
        self.root.geometry("850x600")

        main_frame = ttk.Frame(root, padding=10)
        main_frame.pack(fill="both", expand=True)

        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill="both", expand=True, pady=(0, 10))

        self.tab_parametros = ttk.Frame(self.notebook, padding=10)
        self.tab_danos = ttk.Frame(self.notebook, padding=10)
        self.tab_custas = ttk.Frame(self.notebook, padding=10)

        self.notebook.add(self.tab_parametros, text=" ⚙️ Parâmetros Gerais ")
        self.notebook.add(self.tab_danos, text=" 💰 Danos ")
        self.notebook.add(self.tab_custas, text=" ⚖️ Custas Processuais ")

        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill="x", side="bottom")

        # AGORA O BOTÃO ESTÁ LIGADO À FUNÇÃO!
        btn_importar = ttk.Button(action_frame, text="📂 Importar Planilha Antiga", command=self.importar_planilha)
        btn_importar.pack(side="left")

        btn_gerar = ttk.Button(action_frame, text="🚀 Salvar Template e Gerar Cálculo", command=self.salvar_e_gerar)
        btn_gerar.pack(side="right")

        self.montar_aba_parametros()
        self.montar_aba_danos()
        self.montar_aba_custas()

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

    def montar_aba_danos(self):
        frame_botoes = ttk.Frame(self.tab_danos)
        frame_botoes.pack(fill="x", pady=(0, 10))

        ttk.Button(frame_botoes, text="➕ Adicionar Nova Linha", command=lambda: self.popup_dano()).pack(side="left", padx=5)
        ttk.Button(frame_botoes, text="✏️ Editar Selecionada", command=self.editar_dano_selecionado).pack(side="left", padx=5)
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

    def popup_dano(self, item_selecionado=None):
        popup = tk.Toplevel(self.root)
        popup.title("Editar Dano" if item_selecionado else "Adicionar Dano")
        popup.geometry("450x380")
        popup.transient(self.root)
        popup.grab_set()

        campos = ['ID / Folha', 'Descrição', 'Data Desembolso', 'Valor Histórico', 'Regra', 'Data Juros', 'Valor Ped. Inicial', 'Data do Pedido']
        entradas = {}
        
        valores_atuais = self.tree_danos.item(item_selecionado, 'values') if item_selecionado else [''] * len(campos)

        for i, campo in enumerate(campos):
            ttk.Label(popup, text=f"{campo}:").grid(row=i, column=0, padx=10, pady=5, sticky="w")
            
            if campo == 'Regra':
                opcoes_regras = [
                    "R1 - TJMG + Juros de 1% a.m.",
                    "R2 - Taxa Selic (critério único)",
                    "R3 - TJMG + Juros 1% a.m. até 08/24; após, Selic",
                    "R4 - TJMG + Juros 1% a.m. até 08/24; após, Lei 14.905/24",
                    "R5 - Selic até 08/24; após, Lei 14.905/24 (Tema 1.368 STJ)",
                    "R6 - Lei 14.905/24: IPCA + Taxa Legal",
                    "R7 - TJMG CM + Taxa Legal Retroativa"
                ]
                ent = ttk.Combobox(popup, values=opcoes_regras, state="readonly", width=37)
                
                val_atual = str(valores_atuais[i]) if valores_atuais[i] else ""
                encontrou = False
                for opt in opcoes_regras:
                    if opt.startswith(val_atual):
                        ent.set(opt)
                        encontrou = True
                        break
                if not encontrou:
                    ent.current(1)
            else:
                ent = ttk.Entry(popup, width=40)
                if valores_atuais[i]:
                    ent.insert(0, valores_atuais[i])
                
            ent.grid(row=i, column=1, padx=10, pady=5)
            entradas[campo] = ent

        def salvar_dano():
            novos_valores = []
            for c in campos:
                val = entradas[c].get()
                if c == 'Regra' and " - " in val:
                    val = val.split(" - ")[0]
                novos_valores.append(val)
                
            if item_selecionado:
                self.tree_danos.item(item_selecionado, values=novos_valores)
            else:
                self.tree_danos.insert('', 'end', values=novos_valores)
            popup.destroy()

        ttk.Button(popup, text="Salvar Alterações" if item_selecionado else "Salvar Linha", command=salvar_dano).grid(row=len(campos), column=0, columnspan=2, pady=15)

    def editar_dano_selecionado(self):
        selecionados = self.tree_danos.selection()
        if not selecionados:
            messagebox.showwarning("Aviso", "Selecione uma linha de dano para editar.")
            return
        self.popup_dano(item_selecionado=selecionados[0])

    def remover_dano(self):
        for item in self.tree_danos.selection():
            self.tree_danos.delete(item)

    def limpar_danos(self):
        for item in self.tree_danos.get_children():
            self.tree_danos.delete(item)

    def montar_aba_custas(self):
        frame_botoes = ttk.Frame(self.tab_custas)
        frame_botoes.pack(fill="x", pady=(0, 10))

        ttk.Button(frame_botoes, text="➕ Adicionar Nova Custa", command=lambda: self.popup_custa()).pack(side="left", padx=5)
        ttk.Button(frame_botoes, text="✏️ Editar Selecionada", command=self.editar_custa_selecionada).pack(side="left", padx=5)
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

    def popup_custa(self, item_selecionado=None):
        popup = tk.Toplevel(self.root)
        popup.title("Editar Custa" if item_selecionado else "Adicionar Custa")
        popup.geometry("380x200")
        popup.transient(self.root)
        popup.grab_set()

        campos = ['ID / Folha', 'Descrição', 'Data Desembolso', 'Valor Histórico']
        entradas = {}
        
        valores_atuais = self.tree_custas.item(item_selecionado, 'values') if item_selecionado else [''] * len(campos)

        for i, campo in enumerate(campos):
            ttk.Label(popup, text=f"{campo}:").grid(row=i, column=0, padx=10, pady=5, sticky="w")
            ent = ttk.Entry(popup, width=30)
            if valores_atuais[i]:
                ent.insert(0, valores_atuais[i])
            ent.grid(row=i, column=1, padx=10, pady=5)
            entradas[campo] = ent

        def salvar_custa():
            novos_valores = [entradas[c].get() for c in campos]
            if item_selecionado:
                self.tree_custas.item(item_selecionado, values=novos_valores)
            else:
                self.tree_custas.insert('', 'end', values=novos_valores)
            popup.destroy()

        ttk.Button(popup, text="Salvar Alterações" if item_selecionado else "Salvar Custa", command=salvar_custa).grid(row=len(campos), column=0, columnspan=2, pady=15)

    def editar_custa_selecionada(self):
        selecionados = self.tree_custas.selection()
        if not selecionados:
            messagebox.showwarning("Aviso", "Selecione uma linha de custas para editar.")
            return
        self.popup_custa(item_selecionado=selecionados[0])

    def remover_custa(self):
        for item in self.tree_custas.selection():
            self.tree_custas.delete(item)

    def limpar_custas(self):
        for item in self.tree_custas.get_children():
            self.tree_custas.delete(item)

    # =============== NOVA FUNÇÃO: IMPORTAR PLANILHA ===============
    def importar_planilha(self):
        caminho = filedialog.askopenfilename(
            title="Selecione o Template Antigo", 
            filetypes=[("Planilha Excel", "*.xlsx")]
        )
        if not caminho:
            return
            
        try:
            wb = load_workbook(caminho, data_only=True)
            
            # Importar Parâmetros Gerais
            if 'Parametros' in wb.sheetnames:
                ws = wb['Parametros']
                for row in ws.iter_rows(values_only=True):
                    if row[0]:
                        chave = str(row[0]).strip()
                        valor = row[1] if row[1] is not None else ""
                        if isinstance(valor, datetime):
                            valor = valor.strftime("%d/%m/%Y")
                        if chave in self.vars_param:
                            self.vars_param[chave].set(str(valor))
                            
            # Importar Danos
            if 'Danos' in wb.sheetnames:
                self.limpar_danos()
                ws = wb['Danos']
                headers = [cell.value for cell in ws[1]]
                col_esp = ['ID / Folha', 'Descrição', 'Data Desembolso', 'Valor Histórico', 'Regra', 'Data Juros', 'Valor Pedido Inicial', 'Data do Pedido']
                
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if not any(row): continue
                    row_data = dict(zip(headers, row))
                    valores = []
                    for col in col_esp:
                        val = row_data.get(col, "")
                        val = val if val is not None else ""
                        if isinstance(val, datetime):
                            val = val.strftime("%d/%m/%Y")
                        valores.append(str(val))
                    if valores[1]: # Se a descrição não estiver vazia
                        self.tree_danos.insert('', 'end', values=valores)

            # Importar Custas
            if 'Custas' in wb.sheetnames:
                self.limpar_custas()
                ws = wb['Custas']
                headers = [cell.value for cell in ws[1]]
                col_esp = ['ID / Folha', 'Descrição', 'Data Desembolso', 'Valor Histórico']
                
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if not any(row): continue
                    row_data = dict(zip(headers, row))
                    valores = []
                    for col in col_esp:
                        val = row_data.get(col, "")
                        val = val if val is not None else ""
                        if isinstance(val, datetime):
                            val = val.strftime("%d/%m/%Y")
                        valores.append(str(val))
                    if valores[1]:
                        self.tree_custas.insert('', 'end', values=valores)
            
            messagebox.showinfo("Sucesso", "Planilha carregada e preenchida com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro ao Importar", f"Ocorreu um erro ao tentar ler a planilha:\n{str(e)}")
    # ===============================================================

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
            
            if not caminho_destino:
                return
                
            path_destino = Path(caminho_destino)
            pasta_destino = path_destino.parent
            nome_base = path_destino.stem
            
            nome_template = pasta_destino / f"template {processo_seguro} {data_atual}.xlsx"
            nome_laudo_xls = pasta_destino / f"{nome_base}.xlsx"
            nome_laudo_pdf = pasta_destino / f"{nome_base}.pdf"

            wb = Workbook()
            f_bold = Font(bold=True)
            borda = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            fill_cinza = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

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

            ws_deducoes = wb.create_sheet('Deducoes')
            headers_deducoes = ['ID / Folha', 'Data bloqueio/deposito', 'Valor']
            for col_idx, h in enumerate(headers_deducoes, 1):
                cell = ws_deducoes.cell(row=1, column=col_idx, value=h)
                cell.font = f_bold; cell.fill = fill_cinza; cell.border = borda
                ws_deducoes.column_dimensions[chr(64+col_idx)].width = 25

            wb.save(nome_template)
            
            caminho_nash = Path(__file__).parent / "nash.py"
            resultado = subprocess.run(
                [sys.executable, str(caminho_nash), str(nome_template), str(nome_laudo_xls), str(nome_laudo_pdf)],
                capture_output=True, text=True
            )
            
            if resultado.returncode != 0:
                messagebox.showerror(
                    "Erro no Motor Matemático", 
                    f"O nash.py quebrou nos bastidores!\n\nDetalhes do erro:\n{resultado.stderr}"
                )
            else:
                messagebox.showinfo(
                    "Cálculo Concluído", 
                    f"Processamento finalizado com sucesso!\n\nArquivos guardados em:\n{pasta_destino}"
                )
            
        except Exception as e:
            messagebox.showerror("Erro Crítico", f"Falha ao gerar o Excel:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style()
    if 'clam' in style.theme_names():
        style.theme_use('clam')
    app = NashDataEntryGUI(root)
    root.mainloop()