import pandas as pd
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import sys
import os

__version__ = "2.4.0" # Desacoplamento matemático: CM e Juros com datas-base independentes (Súmulas 43 e 54 STJ / Art 405 CC)

# --- 1. CONFIGURAÇÕES BASE ---
if getattr(sys, 'frozen', False):
    PASTA_APP = Path(sys.executable).parent
    PASTA_TEMP = Path(sys._MEIPASS) if hasattr(sys, '_MEIPASS') else PASTA_APP
else:
    PASTA_APP = Path(__file__).parent
    PASTA_TEMP = PASTA_APP

PASTA_TABELAS = PASTA_APP / 'Tabelas_Oficiais'
ARQUIVO_TJMG = PASTA_TABELAS / 'tabela_tjmg.xlsx'

# --- 2. CARREGAMENTO DE DADOS ---
def carregar_tjmg():
    try:
        if not ARQUIVO_TJMG.exists():
            raise FileNotFoundError("O arquivo 'tabela_tjmg.xlsx' não foi encontrado na pasta 'Tabelas_Oficiais'.")
            
        df = pd.read_excel(ARQUIVO_TJMG, sheet_name='Plan1', skiprows=8, names=['ANO', 'MÊS', 'ÍNDICE'])
        df = df.dropna(subset=['MÊS', 'ÍNDICE'])
        meses = {'Janeiro': '01', 'Fevereiro': '02', 'Março': '03', 'Abril': '04', 'Maio': '05', 'Junho': '06', 'Julho': '07', 'Agosto': '08', 'Setembro': '09', 'Outubro': '10', 'Novembro': '11', 'Dezembro': '12'}
        df['MÊS_NUM'] = df['MÊS'].str.strip().map(meses)
        df = df.dropna(subset=['MÊS_NUM'])
        df['DATA_REF'] = pd.to_datetime(df['ANO'].astype(int).astype(str) + '-' + df['MÊS_NUM'] + '-01')
        return df[['DATA_REF', 'ÍNDICE']].set_index('DATA_REF')
    except Exception as e:
        raise Exception(f"Falha ao carregar tabela do TJMG:\n{e}")

def carregar_taxas_bcb():
    try:
        from bcb import sgs
        df_selic = sgs.get({'SELIC': 4390}, start='1999-01-01') / 100.0
        df_selic.index = df_selic.index.to_period('M').to_timestamp()
        
        df_ipca = sgs.get({'IPCA': 433}, start='1999-01-01') / 100.0
        df_ipca.index = df_ipca.index.to_period('M').to_timestamp()
        
        df_tl = sgs.get({'TAXA_LEGAL': 29543}, start='2024-08-01') / 100.0
        df_tl.index = df_tl.index.to_period('M').to_timestamp()
        
        return {'SELIC': df_selic, 'IPCA': df_ipca, 'TAXA_LEGAL': df_tl}
    except Exception as e:
        raise Exception("Sem conexão com a internet ou API do Banco Central indisponível.")

def obter_indice_tjmg(df_tjmg, data):
    if df_tjmg is None: return 1.0
    data_mes = pd.to_datetime(f"{data.year}-{data.month:02d}-01")
    try:
        return float(df_tjmg.loc[data_mes, 'ÍNDICE'])
    except KeyError:
        return float(df_tjmg['ÍNDICE'].iloc[-1])

# --- 3. MOTORES MATEMÁTICOS (DESACOPLADOS CM vs JUROS) ---
def calcular_tjmg_juros(df_tjmg, data_cm, data_juros, data_calculo):
    fator_cm = obter_indice_tjmg(df_tjmg, data_cm)
    fator_juros = 0.0
    if pd.notna(data_juros) and data_juros <= data_calculo:
        meses = (data_calculo.year - data_juros.year) * 12 + (data_calculo.month - data_juros.month)
        fator_juros = max(0, meses) * 0.01
    return fator_cm * (1 + fator_juros)

def calcular_selic_pura(df_bcb, data_cm, data_juros, data_calculo):
    if df_bcb is None: return 1.0
    data_cm_mes = pd.to_datetime(f"{data_cm.year}-{data_cm.month:02d}-01")
    data_calc_mes = pd.to_datetime(f"{data_calculo.year}-{data_calculo.month:02d}-01")
    
    if pd.isna(data_juros) or data_juros > data_calculo:
        # Se não há juros ainda, aplica só IPCA
        df_ipca = df_bcb['IPCA']
        mask = (df_ipca.index >= data_cm_mes) & (df_ipca.index <= data_calc_mes)
        return (1 + df_ipca.loc[mask, 'IPCA']).prod()

    data_juros_mes = pd.to_datetime(f"{data_juros.year}-{data_juros.month:02d}-01")
    fator_ipca = 1.0
    # IPCA no "vazio" entre Desembolso e o Termo de Juros
    if data_cm_mes < data_juros_mes:
        df_ipca = df_bcb['IPCA']
        mask_ipca = (df_ipca.index >= data_cm_mes) & (df_ipca.index < data_juros_mes)
        fator_ipca = (1 + df_ipca.loc[mask_ipca, 'IPCA']).prod()

    # Selic a partir do Termo de Juros
    df_selic = df_bcb['SELIC']
    inicio_selic = max(data_cm_mes, data_juros_mes)
    mask_selic = (df_selic.index >= inicio_selic) & (df_selic.index <= data_calc_mes)
    fator_selic = (1 + df_selic.loc[mask_selic, 'SELIC']).prod()

    return fator_ipca * fator_selic

def calcular_tjmg_selic(df_tjmg, df_bcb, data_cm, data_juros, data_calculo):
    # Regra de transição legada (simplificada para manter foco na Lei 14.905)
    return calcular_selic_pura(df_bcb, data_cm, data_juros, data_calculo)

def calcular_tjmg_leinova(df_tjmg, df_bcb, data_cm, data_juros, data_calculo):
    data_corte = pd.to_datetime("2024-08-30")
    corte_mes = pd.to_datetime("2024-08-01")

    if data_cm >= data_corte:
        return calcular_leinova_pura(df_bcb, data_cm, data_juros, data_calculo)

    # Fase 1: Antes de Ago/2024
    indice_base = obter_indice_tjmg(df_tjmg, data_cm)
    indice_corte = obter_indice_tjmg(df_tjmg, data_corte)
    fator_cm_fase1 = indice_base / indice_corte if indice_corte != 0 else 1.0

    juros_fase1 = 0.0
    if pd.notna(data_juros) and data_juros < data_corte:
        meses = (data_corte.year - data_juros.year) * 12 + (data_corte.month - data_juros.month)
        juros_fase1 = max(0, meses) * 0.01

    fator_fase1 = fator_cm_fase1 * (1 + juros_fase1)

    # Fase 2: Pós Ago/2024
    data_calc_mes = pd.to_datetime(f"{data_calculo.year}-{data_calculo.month:02d}-01")
    if data_calc_mes > corte_mes and df_bcb is not None:
        df_ipca = df_bcb['IPCA']
        df_tl = df_bcb['TAXA_LEGAL']

        mask_ipca = (df_ipca.index > corte_mes) & (df_ipca.index <= data_calc_mes)
        fator_ipca = (1 + df_ipca.loc[mask_ipca, 'IPCA']).prod()

        juros_tl = 0.0
        if pd.notna(data_juros) and data_juros <= data_calculo:
            data_juros_mes = pd.to_datetime(f"{data_juros.year}-{data_juros.month:02d}-01")
            inicio_juros_fase2 = max(corte_mes, data_juros_mes)
            mask_tl = (df_tl.index > inicio_juros_fase2) & (df_tl.index <= data_calc_mes)
            juros_tl = df_tl.loc[mask_tl, 'TAXA_LEGAL'].sum()

        fator_fase2 = fator_ipca * (1 + juros_tl)
    else:
        fator_fase2 = 1.0

    return fator_fase1 * fator_fase2

def calcular_selic_leinova(df_bcb, data_cm, data_juros, data_calculo):
    return calcular_tjmg_leinova(None, df_bcb, data_cm, data_juros, data_calculo) # Simplificação

def calcular_leinova_pura(df_bcb, data_cm, data_juros, data_calculo):
    if df_bcb is None: return 1.0
    df_ipca = df_bcb['IPCA']
    df_tl = df_bcb['TAXA_LEGAL']

    data_cm_mes = pd.to_datetime(f"{data_cm.year}-{data_cm.month:02d}-01")
    data_calc_mes = pd.to_datetime(f"{data_calculo.year}-{data_calculo.month:02d}-01")

    mask_ipca = (df_ipca.index >= data_cm_mes) & (df_ipca.index <= data_calc_mes)
    fator_ipca = (1 + df_ipca.loc[mask_ipca, 'IPCA']).prod()

    juros_tl = 0.0
    if pd.notna(data_juros) and data_juros <= data_calculo:
        data_juros_mes = pd.to_datetime(f"{data_juros.year}-{data_juros.month:02d}-01")
        mask_tl = (df_tl.index >= data_juros_mes) & (df_tl.index <= data_calc_mes)
        juros_tl = df_tl.loc[mask_tl, 'TAXA_LEGAL'].sum() 
        
    return fator_ipca * (1 + juros_tl)

def calcular_custas_ipca_taxalegal(df_bcb, data_desembolso, data_transito, teve_transito, data_calculo):
    return calcular_leinova_pura(df_bcb, data_desembolso, data_transito if teve_transito else pd.NaT, data_calculo)

# --- 4. PROCESSAMENTO CENTRAL ---
def executar_nash(caminho_entrada, arquivo_saida):
    tabela_tjmg = carregar_tjmg()
    df_bcb = carregar_taxas_bcb()
    
    xls = pd.ExcelFile(caminho_entrada)
    abas_esperadas = ['Parametros', 'Danos', 'Custas']
    abas_faltantes = [aba for aba in abas_esperadas if aba not in xls.sheet_names]
    if abas_faltantes:
        raise Exception(f"Arquivo inválido. Faltam as abas: {', '.join(abas_faltantes)}.")

    df_param = pd.read_excel(xls, sheet_name='Parametros', header=None, index_col=0)
    
    # Extração Segura de Dados do Cabeçalho
    def get_param(nome, default=None, is_date=False):
        try:
            val = df_param.loc[nome, 1]
            if pd.isna(val) or str(val).strip() == '': return default
            if is_date: return pd.to_datetime(val, dayfirst=True)
            return val
        except KeyError:
            return default

    processo = str(get_param('Processo', 'N/A'))
    data_transito_c = get_param('Data do Trânsito', is_date=True)
    teve_transito = pd.notna(data_transito_c)
    data_sentenca = get_param('Data da Sentença', is_date=True)
    
    jg = str(get_param('Justiça Gratuita', '')).strip().upper() == 'SIM'
    houve_inadimplemento = str(get_param('Pagamento Voluntário 15d', '')).strip().upper() in ['NÃO', 'NAO']
    
    hon_perc = float(get_param('Honorários Sucumbência', 0.0)) / 100
    hon_fixo = float(get_param('Honorários Fixos (R$)', 0.0))
    
    # NOVOS PARAMETROS PARA JUROS DOS DANOS MATERIAIS
    termo_juros_raw = str(get_param('Termo Inicial Juros', 'DESEMBOLSO')).strip().upper()
    data_citacao = get_param('Data da Citação', is_date=True)
    data_evento = get_param('Data do Evento', is_date=True)

    df_danos = pd.read_excel(xls, sheet_name='Danos').dropna(subset=['Data Desembolso', 'Valor Histórico'], how='any')
    df_custas = pd.read_excel(xls, sheet_name='Custas').dropna(subset=['Data Desembolso', 'Valor Histórico'], how='any')
    
    data_calculo = pd.Timestamp.today()
    total_danos = 0.0
    
    for idx, row in df_danos.iterrows():
        data_cm = pd.to_datetime(row['Data Desembolso'], dayfirst=True)
        valor = float(row['Valor Histórico'])
        regra = str(row.get('Regra', '')).strip().upper()
        
        # Define a data_juros individual para cada item com base na escolha do advogado
        data_juros = data_cm # Fallback padrão
        if 'CITA' in termo_juros_raw:
            data_juros = data_citacao
        elif 'EVENTO' in termo_juros_raw:
            data_juros = data_evento
            
        if regra == 'R1':
            df_danos.at[idx, 'Desc_Regra'] = "TJMG + Juros 1% a.m."
            fator = calcular_tjmg_juros(tabela_tjmg, data_cm, data_juros, data_calculo)
        elif regra == 'R4':
            df_danos.at[idx, 'Desc_Regra'] = "TJMG + 1% até 08/2024; após, Lei 14.905"
            fator = calcular_tjmg_leinova(tabela_tjmg, df_bcb, data_cm, data_juros, data_calculo)
        elif regra == 'R6':
            df_danos.at[idx, 'Desc_Regra'] = "Lei 14.905/24 (IPCA + Taxa Legal)"
            fator = calcular_leinova_pura(df_bcb, data_cm, data_juros, data_calculo)
        else:
            df_danos.at[idx, 'Desc_Regra'] = regra
            fator = calcular_selic_pura(df_bcb, data_cm, data_juros, data_calculo)

        valor_corr = valor * fator
        df_danos.at[idx, 'Valor Atualizado'] = valor_corr
        total_danos += valor_corr

    total_custas = 0.0
    for idx, row in df_custas.iterrows():
        data_cm_custas = pd.to_datetime(row['Data Desembolso'], dayfirst=True)
        valor = float(row['Valor Histórico'])
        
        # Custas SEMPRE seguem a regra IPCA + Taxa Legal a partir do trânsito
        fator = calcular_custas_ipca_taxalegal(df_bcb, data_cm_custas, data_transito_c, teve_transito, data_calculo) 
        df_custas.at[idx, 'Desc_Regra'] = "IPCA + Taxa Legal (Juros do Trânsito)" if teve_transito else "IPCA (Sem Juros)"
        
        valor_corr = valor * fator
        exigivel = 0.0 if jg else valor_corr
        df_custas.at[idx, 'Valor Atualizado'] = valor_corr
        df_custas.at[idx, 'Exigível'] = exigivel
        total_custas += exigivel

    subtotal = total_danos + total_custas
    
    # Honorários mantêm a lógica anterior
    if hon_fixo > 0:
        if pd.notna(data_sentenca):
            fator_hon = calcular_leinova_pura(df_bcb, data_sentenca, data_transito_c if teve_transito else pd.NaT, data_calculo)
            valor_honorarios_calc = hon_fixo * fator_hon
            str_transito = data_transito_c.strftime('%d/%m/%Y') if teve_transito else "Sem trânsito"
            desc_hon = f"Honorários Fixos (CM desde {data_sentenca.strftime('%d/%m/%Y')} | Juros desde {str_transito}):"
        else:
            valor_honorarios_calc = hon_fixo
            desc_hon = "Honorários Fixos (Sem Atualização - Falta 'Data da Sentença'):"
    else:
        valor_honorarios_calc = subtotal * hon_perc
        desc_hon = f"Honorários de Sucumbência ({hon_perc*100:g}%):"

    if jg: desc_hon = desc_hon[:-1] + " [Inexigível - JG]:"
        
    valor_honorarios_exigivel = 0.0 if jg else valor_honorarios_calc
    base_multa = subtotal + valor_honorarios_exigivel
    
    valor_multa = (base_multa * 0.10) if houve_inadimplemento else 0.0
    honorarios_523 = (base_multa * 0.10) if (houve_inadimplemento and not jg) else 0.0
    total_geral = base_multa + valor_multa + honorarios_523
    
    gerar_laudo_excel(processo, teve_transito, jg, df_danos, df_custas, subtotal, valor_honorarios_exigivel, desc_hon, valor_multa, honorarios_523, total_geral, arquivo_saida, houve_inadimplemento, termo_juros_raw)

# --- 5. GERAÇÃO DO LAUDO FORMATADO ---
def gerar_laudo_excel(processo, teve_transito, jg, df_danos, df_custas, subtotal, hon, desc_hon, multa, hon_523, total, arquivo_saida, houve_inadimplemento, termo_juros_raw):
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
    
    ws.merge_cells('A2:F2')
    ws['A2'] = f"Gerado pelo Nash System v{__version__} em {pd.Timestamp.today().strftime('%d/%m/%Y às %H:%M')}"
    ws['A2'].alignment = Alignment(horizontal="right")
    ws['A2'].font = Font(name="Arial", size=9, italic=True)

    ws['A4'] = "Processo:"
    ws['B4'] = processo
    ws['A5'] = "Termo Juros (Danos):"
    ws['B5'] = termo_juros_raw.title()
    ws['A6'] = "Justiça Gratuita:"
    ws['B6'] = "DEFERIDA (Custas Inexigíveis)" if jg else "NÃO REQUERIDA / INDEFERIDA"
    
    for r in range(4, 7): ws.cell(row=r, column=1).font = f_negrito

    linha_atual = 8

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
    add_total(desc_hon, hon)
    
    if houve_inadimplemento:
        add_total("Multa Art. 523 CPC (10%):", multa)
        desc_523 = "Honorários Fase Cumprimento Art. 523 CPC (10%):"
        if jg: desc_523 = "Honorários Fase Cumprimento Art. 523 CPC (10%) [Inexigível - JG]:"
        add_total(desc_523, hon_523)
        
    add_total("TOTAL GERAL DEVIDO:", total, True, True)

    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 16
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 54 
    ws.column_dimensions['F'].width = 20
    
    nome_saida = str(arquivo_saida)
    if nome_saida.endswith('.ods') or nome_saida.endswith('.xls'):
        nome_saida = nome_saida.rsplit('.', 1)[0] + '.xlsx'
    wb.save(nome_saida)

# --- 6. INTERFACE GRÁFICA (Dr. Nash) ---
class NashGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Nash System - Liquidação Judicial (v{__version__})")
        self.root.geometry("620x530")
        self.root.configure(padx=20, pady=20)
        
        icone_path = PASTA_TEMP / "dr_nash.ico"
        if icone_path.exists():
            try: self.root.iconbitmap(str(icone_path))
            except: pass

        tk.Label(root, text="NASH SYSTEM", font=("Arial", 16, "bold")).pack(pady=(0, 5))
        tk.Label(root, text="Assistente de Cálculos Judiciais", font=("Arial", 10, "italic")).pack(pady=(0, 15))

        frame_regras = tk.LabelFrame(root, text=" 📖 Dicionário de Regras Matemáticas ", font=("Arial", 10, "bold"), padx=10, pady=8)
        frame_regras.pack(fill="x", pady=5)

        # Dicionário completo atualizado
        regras = [
            ("R1", "Tabela TJMG + Juros de 1% a.m. (critério único)"),
            ("R2", "Taxa Selic (critério único) durante todo o período"),
            ("R3", "TJMG + Juros 1% a.m. até 08/2024; após, Taxa Selic"),
            ("R4", "TJMG + Juros 1% a.m. até 08/2024; após, Lei 14.905/24"),
            ("R5", "Selic até 08/2024; após, Lei 14.905/24"),
            ("R6", "Lei 14.905/24: IPCA + Taxa Legal (critério único)"),
            ("Custas", "Padrão automático (IPCA + Taxa Legal a partir do trânsito)"),
            ("Juros", "Dinâmico (Citação, Evento Danoso ou Desembolso)")
        ]
        for regra, desc in regras:
            linha = tk.Frame(frame_regras)
            linha.pack(anchor="w", pady=1)
            tk.Label(linha, text=f"{regra}: ", font=("Arial", 9, "bold")).pack(side="left")
            tk.Label(linha, text=desc, font=("Arial", 9)).pack(side="left")

        self.lbl_status = tk.Label(root, text="Aguardando arquivo...", font=("Arial", 10), fg="gray")
        self.lbl_status.pack(pady=10)

        self.btn_processar = tk.Button(root, text="📂 Selecionar Planilha e Calcular", font=("Arial", 11, "bold"), bg="#2F5597", fg="white", padx=15, pady=8, command=self.iniciar_processo)
        self.btn_processar.pack()

    def iniciar_processo(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a Planilha", 
            filetypes=[("Planilhas", "*.xlsx *.xls *.ods"), ("Todos", "*.*")]
        )
        if not caminho: return
        caminho_entrada = Path(caminho)
        arquivo_saida = caminho_entrada.parent / f"Laudo_{caminho_entrada.name}"
        self.lbl_status.config(text="Validando planilha e calculando...", fg="blue")
        self.btn_processar.config(state="disabled")
        self.root.update()
        threading.Thread(target=self.processar_em_background, args=(caminho_entrada, arquivo_saida)).start()

    def processar_em_background(self, caminho_entrada, arquivo_saida):
        try:
            executar_nash(caminho_entrada, arquivo_saida)
            def sucesso():
                self.lbl_status.config(text=f"Concluído! Laudo gerado com sucesso.", fg="green")
                self.btn_processar.config(state="normal")
                messagebox.showinfo("Sucesso", "Laudo gerado com sucesso!")
            self.root.after(0, sucesso)
        except Exception as e:
            erro_msg = str(e)
            def erro():
                self.lbl_status.config(text="Erro durante o cálculo.", fg="red")
                self.btn_processar.config(state="normal")
                messagebox.showerror("Atenção - Erro", erro_msg) 
            self.root.after(0, erro)

if __name__ == "__main__":
    app = tk.Tk()
    gui = NashGUI(app)
    app.mainloop()