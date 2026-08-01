import pandas as pd
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import sys
import os

# --- 1. CONFIGURAÇÕES BASE ---
if getattr(sys, 'frozen', False):
    PASTA_BASE = Path(sys.executable).parent
else:
    PASTA_BASE = Path(__file__).parent

PASTA_TABELAS = PASTA_BASE / 'Tabelas_Oficiais'
ARQUIVO_TJMG = PASTA_TABELAS / 'tabela_tjmg.xlsx'

# --- 2. MOTORES MATEMÁTICOS REAIS ---
def carregar_tjmg():
    try:
        df = pd.read_excel(ARQUIVO_TJMG, sheet_name='Plan1', skiprows=8, names=['ANO', 'MÊS', 'ÍNDICE'])
        df = df.dropna(subset=['MÊS', 'ÍNDICE'])
        meses = {'Janeiro': '01', 'Fevereiro': '02', 'Março': '03', 'Abril': '04', 'Maio': '05', 'Junho': '06', 'Julho': '07', 'Agosto': '08', 'Setembro': '09', 'Outubro': '10', 'Novembro': '11', 'Dezembro': '12'}
        df['MÊS_NUM'] = df['MÊS'].str.strip().map(meses)
        df = df.dropna(subset=['MÊS_NUM'])
        df['DATA_REF'] = pd.to_datetime(df['ANO'].astype(int).astype(str) + '-' + df['MÊS_NUM'] + '-01')
        return df[['DATA_REF', 'ÍNDICE']].set_index('DATA_REF')
    except Exception as e:
        raise Exception(f"Erro ao carregar tabela do TJMG: {e}")

def carregar_taxas_bcb():
    try:
        from bcb import sgs
        df = sgs.get({'SELIC': 4390}, start='1999-01-01')
        df = df / 100.0  
        df.index = df.index.to_period('M').to_timestamp()
        return df
    except Exception as e:
        raise Exception(f"Erro ao conectar com BCB: {e}")

def obter_indice_tjmg(df_tjmg, data):
    if df_tjmg is None: return 1.0
    data_mes = pd.to_datetime(f"{data.year}-{data.month:02d}-01")
    try:
        return float(df_tjmg.loc[data_mes, 'ÍNDICE'])
    except KeyError:
        return float(df_tjmg['ÍNDICE'].iloc[-1])

def calcular_fator_r5(df_tjmg, data_base, data_calculo):
    fator_cm = obter_indice_tjmg(df_tjmg, data_base)
    meses = (data_calculo.year - data_base.year) * 12 + (data_calculo.month - data_base.month)
    if meses < 0: meses = 0
    fator_juros = meses * 0.01
    return fator_cm * (1 + fator_juros)

def calcular_fator_r4(df_bcb, data_base, data_calculo):
    if df_bcb is None: return 1.0
    data_base_mes = pd.to_datetime(f"{data_base.year}-{data_base.month:02d}-01")
    data_calc_mes = pd.to_datetime(f"{data_calculo.year}-{data_calculo.month:02d}-01")
    mask = (df_bcb.index >= data_base_mes) & (df_bcb.index <= data_calc_mes)
    return (1 + df_bcb.loc[mask, 'SELIC']).prod()

def calcular_fator_r1(df_tjmg, df_bcb, data_base, data_calculo):
    data_corte = pd.to_datetime("2024-08-30")
    if data_base >= data_corte:
        return calcular_fator_r4(df_bcb, data_base, data_calculo)

    indice_base = obter_indice_tjmg(df_tjmg, data_base)
    indice_corte = obter_indice_tjmg(df_tjmg, data_corte)
    fator_cm = indice_base / indice_corte if indice_corte != 0 else 1.0
    meses = (data_corte.year - data_base.year) * 12 + (data_corte.month - data_base.month)
    if meses < 0: meses = 0
    fator_fase1 = fator_cm * (1 + (meses * 0.01))

    if df_bcb is not None:
        corte_mes = pd.to_datetime(f"{data_corte.year}-{data_corte.month:02d}-01")
        calc_mes = pd.to_datetime(f"{data_calculo.year}-{data_calculo.month:02d}-01")
        mask = (df_bcb.index > corte_mes) & (df_bcb.index <= calc_mes)
        fator_fase2 = (1 + df_bcb.loc[mask, 'SELIC']).prod()
    else:
        fator_fase2 = 1.0

    return fator_fase1 * fator_fase2

def calcular_fator_r2(df_bcb, data_base, data_calculo):
    data_corte = pd.to_datetime("2024-08-30")
    if data_base >= data_corte:
        return calcular_fator_r4(df_bcb, data_base, data_calculo)
    
    # Fase 1: Selic do evento até a entrada em vigor da Lei Nova (30/08/2024)
    fator_fase1 = calcular_fator_r4(df_bcb, data_base, data_corte)
    # Fase 2: Critérios da Lei Nova (30/08/2024 em diante)
    fator_fase2 = calcular_fator_r4(df_bcb, data_corte, data_calculo)
    
    return fator_fase1 * fator_fase2

def calcular_fator_r3(df_bcb, data_base, data_calculo):
    return calcular_fator_r4(df_bcb, data_base, data_calculo)

# --- 3. PROCESSAMENTO CENTRAL DO NASH ---
def executar_nash(caminho_entrada, arquivo_saida):
    tabela_tjmg = carregar_tjmg()
    df_bcb = carregar_taxas_bcb()
    
    df_param = pd.read_excel(caminho_entrada, sheet_name='Parametros', header=None, index_col=0)
    processo = str(df_param.loc['Processo', 1])
    data_transito_raw = df_param.loc['Data do Trânsito', 1]
    teve_transito = not pd.isna(data_transito_raw)
    
    jg = str(df_param.loc['Justiça Gratuita', 1]).strip().upper() == 'SIM'
    pag_voluntario = str(df_param.loc['Pagamento Voluntário 15d', 1]).strip().upper() == 'SIM'
    hon_perc = float(df_param.loc['Honorários Sucumbência', 1]) / 100

    df_danos = pd.read_excel(caminho_entrada, sheet_name='Danos')
    df_custas = pd.read_excel(caminho_entrada, sheet_name='Custas')
    
    data_calculo = pd.Timestamp.today()
    total_danos = 0.0
    
    for idx, row in df_danos.iterrows():
        data = pd.to_datetime(row['Data Desembolso'], dayfirst=True)
        valor = float(row['Valor Histórico'])
        regra = str(row.get('Regra', '')).strip().upper()
        
        if regra == 'R1':
            df_danos.at[idx, 'Desc_Regra'] = "TJMG + Juros de 1% a.m. até 08/2024; após, Selic"
            fator = calcular_fator_r1(tabela_tjmg, df_bcb, data, data_calculo)
        elif regra == 'R2':
            df_danos.at[idx, 'Desc_Regra'] = "Selic até 08/2024; após, critérios da Lei Nova (14.905/24)"
            fator = calcular_fator_r2(df_bcb, data, data_calculo)
        elif regra == 'R3':
            df_danos.at[idx, 'Desc_Regra'] = "IPCA + (Selic deduzida do IPCA)"
            fator = calcular_fator_r3(df_bcb, data, data_calculo)
        elif regra == 'R4':
            df_danos.at[idx, 'Desc_Regra'] = "Taxa Selic (critério único) durante todo o período"
            fator = calcular_fator_r4(df_bcb, data, data_calculo)
        elif regra == 'R5':
            df_danos.at[idx, 'Desc_Regra'] = "Tabela TJMG + Juros de 1% a.m. (Critério único)"
            fator = calcular_fator_r5(tabela_tjmg, data, data_calculo)
        else:
            df_danos.at[idx, 'Desc_Regra'] = "Regra não identificada. Sem correção."
            fator = 1.0

        valor_corr = valor * fator
        df_danos.at[idx, 'Valor Atualizado'] = valor_corr
        total_danos += valor_corr

    total_custas = 0.0
    for idx, row in df_custas.iterrows():
        data = pd.to_datetime(row['Data Desembolso'], dayfirst=True)
        valor = float(row['Valor Histórico'])
        fator = calcular_fator_r5(tabela_tjmg, data, data_calculo) 
        df_custas.at[idx, 'Desc_Regra'] = "Tabela TJMG + Juros de 1% a.m. (Critério único)"
        
        valor_corr = valor * fator
        exigivel = 0.0 if jg else valor_corr
        df_custas.at[idx, 'Valor Atualizado'] = valor_corr
        df_custas.at[idx, 'Exigível'] = exigivel
        total_custas += exigivel

    subtotal = total_danos + total_custas
    valor_honorarios = subtotal * hon_perc
    base_multa = subtotal + valor_honorarios
    
    aplica_multa = teve_transito and not pag_voluntario
    valor_multa = (base_multa * 0.10) if aplica_multa else 0.0
    honorarios_523 = (base_multa * 0.10) if aplica_multa else 0.0
    
    total_geral = base_multa + valor_multa + honorarios_523

    gerar_laudo_excel(processo, teve_transito, jg, df_danos, df_custas, subtotal, valor_honorarios, valor_multa, honorarios_523, total_geral, arquivo_saida)

# --- 4. GERAÇÃO DO LAUDO FORMATADO ---
def gerar_laudo_excel(processo, teve_transito, jg, df_danos, df_custas, subtotal, hon, multa, hon_523, total, arquivo_saida):
    wb = Workbook()
    ws = wb.active
    ws.title = "Laudo de Liquidação"

    f_titulo = Font(name="Arial", size=14, bold=True, color="FFFFFF")
    f_negrito = Font(name="Arial", size=11, bold=True)
    f_normal = Font(name="Arial", size=11)
    fundo_escuro = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fundo_cinza = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    borda = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    centro = Alignment(horizontal="center", vertical="center")
    moeda = 'R$ #,##0.00'

    def formatar_linha(linha, cor=None, fonte=f_normal):
        for col in range(1, 7):
            c = ws.cell(row=linha, column=col)
            c.font = fonte
            c.border = borda
            if cor: c.fill = cor

    ws.merge_cells('A1:F1')
    ws['A1'] = f"LAUDO DE CÁLCULO JUDICIAL - NASH SYSTEM"
    ws['A1'].font = f_titulo
    ws['A1'].fill = fundo_escuro
    ws['A1'].alignment = centro

    ws['A3'] = "Processo:"
    ws['B3'] = processo
    ws['A4'] = "Tipo de Título:"
    ws['B4'] = "Sentença (Com Trânsito)" if teve_transito else "Acordo / Sem Trânsito"
    ws['A5'] = "Justiça Gratuita:"
    ws['B5'] = "DEFERIDA (Custas Inexigíveis)" if jg else "NÃO REQUERIDA / INDEFERIDA"
    
    for r in range(3, 6):
        ws.cell(row=r, column=1).font = f_negrito

    linha_atual = 7

    ws.merge_cells(f'A{linha_atual}:F{linha_atual}')
    ws[f'A{linha_atual}'] = "1. DANOS MATERIAIS / NOTAS"
    ws[f'A{linha_atual}'].font = f_negrito
    ws[f'A{linha_atual}'].fill = fundo_cinza
    linha_atual += 1

    cabecalhos = ['ID / Folha', 'Descrição', 'Data Desembolso', 'Valor Histórico', 'Índice Aplicado', 'Valor Atualizado']
    for i, texto in enumerate(cabecalhos, 1):
        ws.cell(row=linha_atual, column=i, value=texto).font = f_negrito
        ws.cell(row=linha_atual, column=i).border = borda
    linha_atual += 1

    for _, row in df_danos.iterrows():
        ws.cell(row=linha_atual, column=1, value=row['ID / Folha'])
        ws.cell(row=linha_atual, column=2, value=row['Descrição'])
        ws.cell(row=linha_atual, column=3, value=row['Data Desembolso'].strftime('%d/%m/%Y'))
        ws.cell(row=linha_atual, column=4, value=row['Valor Histórico']).number_format = moeda
        ws.cell(row=linha_atual, column=5, value=row.get('Desc_Regra', ''))
        ws.cell(row=linha_atual, column=6, value=row['Valor Atualizado']).number_format = moeda
        formatar_linha(linha_atual)
        linha_atual += 1

    linha_atual += 1

    ws.merge_cells(f'A{linha_atual}:F{linha_atual}')
    ws[f'A{linha_atual}'] = "2. CUSTAS E DESPESAS PROCESSUAIS"
    ws[f'A{linha_atual}'].font = f_negrito
    ws[f'A{linha_atual}'].fill = fundo_cinza
    linha_atual += 1

    cab_custas = ['ID / Folha', 'Descrição', 'Data', 'Valor Histórico', 'Índice Aplicado', 'Atualizado Exigível']
    for i, texto in enumerate(cab_custas, 1):
        ws.cell(row=linha_atual, column=i, value=texto).font = f_negrito
        ws.cell(row=linha_atual, column=i).border = borda
    linha_atual += 1

    for _, row in df_custas.iterrows():
        ws.cell(row=linha_atual, column=1, value=row['ID / Folha'])
        ws.cell(row=linha_atual, column=2, value=row['Descrição'])
        ws.cell(row=linha_atual, column=3, value=row['Data Desembolso'].strftime('%d/%m/%Y'))
        ws.cell(row=linha_atual, column=4, value=row['Valor Histórico']).number_format = moeda
        ws.cell(row=linha_atual, column=5, value=row.get('Desc_Regra', ''))
        ws.cell(row=linha_atual, column=6, value=row['Exigível']).number_format = moeda
        formatar_linha(linha_atual)
        linha_atual += 1

    linha_atual += 2

    ws.merge_cells(f'A{linha_atual}:F{linha_atual}')
    ws[f'A{linha_atual}'] = "3. RESUMO DA LIQUIDAÇÃO"
    ws[f'A{linha_atual}'].font = f_titulo
    ws[f'A{linha_atual}'].fill = fundo_escuro
    ws[f'A{linha_atual}'].alignment = centro
    linha_atual += 1

    def add_total(desc, valor, negrito=False, destaque=False):
        nonlocal linha_atual
        ws.merge_cells(f'A{linha_atual}:E{linha_atual}')
        ws.cell(row=linha_atual, column=1, value=desc).alignment = Alignment(horizontal="right")
        c_val = ws.cell(row=linha_atual, column=6, value=valor)
        c_val.number_format = moeda
        
        fonte_usada = f_negrito if negrito else f_normal
        ws.cell(row=linha_atual, column=1).font = fonte_usada
        c_val.font = fonte_usada
        
        if destaque:
            ws.cell(row=linha_atual, column=1).fill = fundo_cinza
            c_val.fill = fundo_cinza
            
        for i in range(1, 7): ws.cell(row=linha_atual, column=i).border = borda
        linha_atual += 1

    add_total("SUBTOTAL (Principal + Custas):", subtotal, True)
    add_total("Honorários de Sucumbência:", hon)
    if multa > 0:
        add_total("Multa Art. 523 CPC (10%):", multa)
        add_total("Honorários Fase Cumprimento Art. 523 CPC (10%):", hon_523)
    add_total("TOTAL GERAL DEVIDO:", total, True, True)

    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 16
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 50 
    ws.column_dimensions['F'].width = 20

    wb.save(arquivo_saida)

# --- 5. INTERFACE GRÁFICA (Dr. Nash) ---
class NashGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Nash System - Liquidação Judicial")
        self.root.geometry("600x480")
        self.root.configure(padx=20, pady=20)
        
        icone_path = PASTA_BASE / "dr_nash.ico"
        if icone_path.exists():
            try:
                self.root.iconbitmap(str(icone_path))
            except:
                pass

        tk.Label(root, text="NASH SYSTEM", font=("Arial", 16, "bold")).pack(pady=(0, 5))
        tk.Label(root, text="Assistente de Cálculos Judiciais", font=("Arial", 10, "italic")).pack(pady=(0, 20))

        # Legenda das Regras
        frame_regras = tk.LabelFrame(root, text=" 📖 Dicionário de Regras Matemáticas ", font=("Arial", 10, "bold"), padx=10, pady=10)
        frame_regras.pack(fill="x", pady=10)

        regras = [
            ("R1", "TJMG + Juros de 1% a.m. até 08/2024; após, Taxa Selic"),
            ("R2", "Selic até 08/2024; após, critérios da Lei Nova (14.905/24)"),
            ("R3", "IPCA + (Selic deduzida do IPCA)"),
            ("R4", "Taxa Selic (critério único) durante todo o período"),
            ("R5", "Tabela TJMG + Juros de 1% a.m. (critério único, Lei Antiga)")
        ]

        for regra, desc in regras:
            linha = tk.Frame(frame_regras)
            linha.pack(anchor="w", pady=2)
            tk.Label(linha, text=f"{regra}: ", font=("Arial", 10, "bold")).pack(side="left")
            tk.Label(linha, text=desc, font=("Arial", 10)).pack(side="left")

        # Status
        self.lbl_status = tk.Label(root, text="Aguardando arquivo...", font=("Arial", 10), fg="gray")
        self.lbl_status.pack(pady=15)

        # Botão Principal
        self.btn_processar = tk.Button(root, text="📂 Selecionar Planilha e Calcular", font=("Arial", 12, "bold"), bg="#2F5597", fg="white", padx=20, pady=10, command=self.iniciar_processo)
        self.btn_processar.pack()

    def iniciar_processo(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a Planilha do Cliente",
            filetypes=[("Planilhas Excel", "*.xlsx")]
        )
        
        if not caminho:
            return
            
        caminho_entrada = Path(caminho)
        arquivo_saida = caminho_entrada.parent / f"Laudo_{caminho_entrada.name}"

        self.lbl_status.config(text="Conectando ao Banco Central e calculando...", fg="blue")
        self.btn_processar.config(state="disabled")
        self.root.update()

        threading.Thread(target=self.processar_em_background, args=(caminho_entrada, arquivo_saida)).start()

    def processar_em_background(self, caminho_entrada, arquivo_saida):
        try:
            executar_nash(caminho_entrada, arquivo_saida)
            
            def sucesso():
                self.lbl_status.config(text=f"Concluído! Salvo em:\n{arquivo_saida.name}", fg="green")
                self.btn_processar.config(state="normal")
                messagebox.showinfo("Sucesso", "Liquidação calculada e laudo gerado com sucesso!")
                
            self.root.after(0, sucesso)
            
        except Exception as e:
            def erro():
                self.lbl_status.config(text="Erro durante o cálculo.", fg="red")
                self.btn_processar.config(state="normal")
                messagebox.showerror("Erro", str(e))
                
            self.root.after(0, erro)

if __name__ == "__main__":
    app = tk.Tk()
    gui = NashGUI(app)
    app.mainloop()