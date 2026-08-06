import pandas as pd
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import sys
import os

__version__ = "2.3.7" # Suporte multi-formato (ODS/XLS/XLSX) e Validação amigável de abas

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
        
        ultima_data = df['DATA_REF'].max()
        hoje = pd.Timestamp.today()
        diferenca_meses = (hoje.year - ultima_data.year) * 12 + (hoje.month - ultima_data.month)
        if diferenca_meses >= 3:
            raise ValueError(f"ATENÇÃO: A tabela do TJMG está desatualizada (último mês: {ultima_data.strftime('%m/%Y')}).\n\nBaixe a versão mais recente no site do Tribunal e substitua o arquivo na pasta 'Tabelas_Oficiais'.")
            
        return df[['DATA_REF', 'ÍNDICE']].set_index('DATA_REF')
    except ValueError as ve:
        raise ve
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
        raise Exception("Sem conexão com a internet ou a API do Banco Central está fora do ar.\nVerifique sua rede e tente novamente.")

def obter_indice_tjmg(df_tjmg, data):
    if df_tjmg is None: return 1.0
    data_mes = pd.to_datetime(f"{data.year}-{data.month:02d}-01")
    try:
        return float(df_tjmg.loc[data_mes, 'ÍNDICE'])
    except KeyError:
        return float(df_tjmg['ÍNDICE'].iloc[-1])

# --- 3. MOTORES MATEMÁTICOS PRINCIPAIS ---
def calcular_tjmg_juros(df_tjmg, data_base, data_calculo):
    fator_cm = obter_indice_tjmg(df_tjmg, data_base)
    meses = (data_calculo.year - data_base.year) * 12 + (data_calculo.month - data_base.month)
    if meses < 0: meses = 0
    fator_juros = meses * 0.01
    return fator_cm * (1 + fator_juros)

def calcular_selic_pura(df_bcb, data_base, data_calculo):
    if df_bcb is None: return 1.0
    df_selic = df_bcb['SELIC']
    data_base_mes = pd.to_datetime(f"{data_base.year}-{data_base.month:02d}-01")
    data_calc_mes = pd.to_datetime(f"{data_calculo.year}-{data_calculo.month:02d}-01")
    mask = (df_selic.index >= data_base_mes) & (df_selic.index <= data_calc_mes)
    return (1 + df_selic.loc[mask, 'SELIC']).prod()

def calcular_tjmg_selic(df_tjmg, df_bcb, data_base, data_calculo):
    data_corte = pd.to_datetime("2024-08-30")
    if data_base >= data_corte: return calcular_selic_pura(df_bcb, data_base, data_calculo)
    indice_base = obter_indice_tjmg(df_tjmg, data_base)
    indice_corte = obter_indice_tjmg(df_tjmg, data_corte)
    fator_cm = indice_base / indice_corte if indice_corte != 0 else 1.0
    meses = max(0, (data_corte.year - data_base.year) * 12 + (data_corte.month - data_base.month))
    fator_fase1 = fator_cm * (1 + (meses * 0.01))
    if df_bcb is not None:
        corte_mes = pd.to_datetime(f"{data_corte.year}-{data_corte.month:02d}-01")
        calc_mes = pd.to_datetime(f"{data_calculo.year}-{data_calculo.month:02d}-01")
        df_selic = df_bcb['SELIC']
        mask = (df_selic.index > corte_mes) & (df_selic.index <= calc_mes)
        fator_fase2 = (1 + df_selic.loc[mask, 'SELIC']).prod()
    else: fator_fase2 = 1.0
    return fator_fase1 * fator_fase2

def calcular_tjmg_leinova(df_tjmg, df_bcb, data_base, data_calculo):
    data_corte = pd.to_datetime("2024-08-30")
    corte_mes = pd.to_datetime("2024-08-01")
    if data_base >= data_corte: return calcular_leinova_pura(df_bcb, data_base, data_calculo)
    indice_base = obter_indice_tjmg(df_tjmg, data_base)
    indice_corte = obter_indice_tjmg(df_tjmg, data_corte)
    fator_cm_fase1 = indice_base / indice_corte if indice_corte != 0 else 1.0
    meses = max(0, (data_corte.year - data_base.year) * 12 + (data_corte.month - data_base.month))
    fator_fase1 = fator_cm_fase1 * (1 + (meses * 0.01))
    data_calc_mes = pd.to_datetime(f"{data_calculo.year}-{data_calculo.month:02d}-01")
    if data_calc_mes > corte_mes and df_bcb is not None:
        df_ipca = df_bcb['IPCA']
        df_tl = df_bcb['TAXA_LEGAL']
        mask_ipca = (df_ipca.index > corte_mes) & (df_ipca.index <= data_calc_mes)
        fator_ipca = (1 + df_ipca.loc[mask_ipca, 'IPCA']).prod()
        mask_tl = (df_tl.index > corte_mes) & (df_tl.index <= data_calc_mes)
        juros_tl = df_tl.loc[mask_tl, 'TAXA_LEGAL'].sum()
        fator_fase2 = fator_ipca * (1 + juros_tl)
    else: fator_fase2 = 1.0
    return fator_fase1 * fator_fase2

def calcular_selic_leinova(df_bcb, data_base, data_calculo):
    data_corte = pd.to_datetime("2024-08-30")
    corte_mes = pd.to_datetime("2024-08-01")
    if data_base >= data_corte: return calcular_leinova_pura(df_bcb, data_base, data_calculo)
    df_selic = df_bcb['SELIC']
    data_base_mes = pd.to_datetime(f"{data_base.year}-{data_base.month:02d}-01")
    mask_fase1 = (df_selic.index >= data_base_mes) & (df_selic.index <= corte_mes)
    fator_fase1 = (1 + df_selic.loc[mask_fase1, 'SELIC']).prod()
    data_calc_mes = pd.to_datetime(f"{data_calculo.year}-{data_calculo.month:02d}-01")
    if data_calc_mes > corte_mes:
        df_ipca = df_bcb['IPCA']
        df_tl = df_bcb['TAXA_LEGAL']
        mask_ipca = (df_ipca.index > corte_mes) & (df_ipca.index <= data_calc_mes)
        fator_ipca = (1 + df_ipca.loc[mask_ipca, 'IPCA']).prod()
        mask_tl = (df_tl.index > corte_mes) & (df_tl.index <= data_calc_mes)
        juros_tl = df_tl.loc[mask_tl, 'TAXA_LEGAL'].sum()
        fator_fase2 = fator_ipca * (1 + juros_tl)
    else: fator_fase2 = 1.0
    return fator_fase1 * fator_fase2

def calcular_leinova_pura(df_bcb, data_base, data_calculo):
    if df_bcb is None: return 1.0
    df_ipca = df_bcb['IPCA']
    df_tl = df_bcb['TAXA_LEGAL']
    data_base_mes = pd.to_datetime(f"{data_base.year}-{data_base.month:02d}-01")
    data_calc_mes = pd.to_datetime(f"{data_calculo.year}-{data_calculo.month:02d}-01")
    mask_ipca = (df_ipca.index >= data_base_mes) & (df_ipca.index <= data_calc_mes)
    fator_ipca = (1 + df_ipca.loc[mask_ipca, 'IPCA']).prod()
    juros_tl = df_tl.loc[(df_tl.index >= data_base_mes) & (df_tl.index <= data_calc_mes), 'TAXA_LEGAL'].sum() 
    return fator_ipca * (1 + juros_tl)

# --- MOTOR EXCLUSIVO: CUSTAS PROCESSUAIS (NOVA REGRA IPCA + TAXA LEGAL) ---
def calcular_custas_ipca_taxalegal(df_bcb, data_desembolso, data_transito, teve_transito, data_calculo):
    if df_bcb is None: return 1.0
    df_ipca = df_bcb['IPCA']
    df_tl = df_bcb['TAXA_LEGAL']

    # 1. Correção Monetária: IPCA desde o desembolso
    data_base_mes = pd.to_datetime(f"{data_desembolso.year}-{data_desembolso.month:02d}-01")
    data_calc_mes = pd.to_datetime(f"{data_calculo.year}-{data_calculo.month:02d}-01")
    mask_ipca = (df_ipca.index >= data_base_mes) & (df_ipca.index <= data_calc_mes)
    fator_ipca = (1 + df_ipca.loc[mask_ipca, 'IPCA']).prod()

    # 2. Juros de Mora: Apenas se houver trânsito
    juros_total = 0.0
    if teve_transito and pd.notna(data_transito):
        transito_mes = pd.to_datetime(f"{data_transito.year}-{data_transito.month:02d}-01")
        corte_juros = pd.to_datetime("2024-08-01")

        if transito_mes < corte_juros:
            meses_1pct = (corte_juros.year - transito_mes.year) * 12 + (corte_juros.month - transito_mes.month)
            juros_fase1 = max(0, meses_1pct) * 0.01
            inicio_fase2 = corte_juros
        else:
            juros_fase1 = 0.0
            inicio_fase2 = transito_mes

        if data_calc_mes > corte_juros:
            mask_tl = (df_tl.index > inicio_fase2) & (df_tl.index <= data_calc_mes)
            juros_fase2 = df_tl.loc[mask_tl, 'TAXA_LEGAL'].sum()
        else:
            juros_fase2 = 0.0

        juros_total = juros_fase1 + juros_fase2

    return fator_ipca * (1 + juros_total)

# --- MOTOR EXCLUSIVO: HONORÁRIOS EQUITATIVOS (DATAS DIVIDIDAS) ---
def atualizar_honorarios_fixos(valor, data_sentenca, data_transito, df_tjmg, df_bcb, data_calculo):
    data_corte = pd.to_datetime("2024-08-30")
    corte_mes = pd.to_datetime("2024-08-01")
    calc_mes = pd.to_datetime(f"{data_calculo.year}-{data_calculo.month:02d}-01")
    sentenca_mes = pd.to_datetime(f"{data_sentenca.year}-{data_sentenca.month:02d}-01")

    if data_sentenca < data_corte:
        indice_sent = obter_indice_tjmg(df_tjmg, data_sentenca)
        indice_corte = obter_indice_tjmg(df_tjmg, data_corte)
        fator_cm_fase1 = indice_sent / indice_corte if indice_corte != 0 else 1.0
    else:
        fator_cm_fase1 = 1.0

    fator_cm_fase2 = 1.0
    if calc_mes > corte_mes and df_bcb is not None:
        inicio_cm_fase2 = corte_mes if data_sentenca < data_corte else sentenca_mes
        df_ipca = df_bcb['IPCA']
        mask_ipca = (df_ipca.index > inicio_cm_fase2) & (df_ipca.index <= calc_mes)
        fator_cm_fase2 = (1 + df_ipca.loc[mask_ipca, 'IPCA']).prod()

    valor_corrigido = valor * fator_cm_fase1 * fator_cm_fase2

    juros_total = 0.0
    if pd.notna(data_transito):
        transito_mes = pd.to_datetime(f"{data_transito.year}-{data_transito.month:02d}-01")
        if data_transito < data_corte:
            meses = (data_corte.year - data_transito.year) * 12 + (data_corte.month - data_transito.month)
            juros_fase1 = max(0, meses) * 0.01
        else:
            juros_fase1 = 0.0

        juros_fase2 = 0.0
        if calc_mes > corte_mes and df_bcb is not None:
            inicio_juros_fase2 = corte_mes if data_transito < data_corte else transito_mes
            df_tl = df_bcb['TAXA_LEGAL']
            mask_tl = (df_tl.index > inicio_juros_fase2) & (df_tl.index <= calc_mes)
            juros_fase2 = df_tl.loc[mask_tl, 'TAXA_LEGAL'].sum()

        juros_total = juros_fase1 + juros_fase2

    return valor_corrigido * (1 + juros_total)

# --- 4. PROCESSAMENTO CENTRAL ---
def executar_nash(caminho_entrada, arquivo_saida):
    tabela_tjmg = carregar_tjmg()
    df_bcb = carregar_taxas_bcb()
    
    # -------------------------------------------------------------
    # BLINDAGEM DE FORMATOS E ABAS: Leitura única na memória
    # -------------------------------------------------------------
    try:
        xls = pd.ExcelFile(caminho_entrada)
    except Exception as e:
        raise Exception(f"Não foi possível abrir o arquivo.\nFormato inválido ou arquivo corrompido.\n\nDetalhe Técnico: {e}")

    abas_esperadas = ['Parametros', 'Danos', 'Custas']
    abas_faltantes = [aba for aba in abas_esperadas if aba not in xls.sheet_names]
    
    if abas_faltantes:
        raise Exception(
            f"O arquivo selecionado NÃO é o template correto do Nash System.\n\n"
            f"Faltam as seguintes abas: {', '.join(abas_faltantes)}.\n\n"
            f"DICA: Verifique se você não abriu uma planilha vazia, um arquivo CSV puro "
            f"ou se alguém apagou abas acidentalmente."
        )

    # A partir daqui o arquivo é validado e seguro
    df_param = pd.read_excel(xls, sheet_name='Parametros', header=None, index_col=0)
    processo = str(df_param.loc['Processo', 1])
    
    data_transito_raw = df_param.loc['Data do Trânsito', 1]
    teve_transito = not pd.isna(data_transito_raw)
    data_transito_c = pd.to_datetime(data_transito_raw, dayfirst=True) if teve_transito else None
    
    jg = str(df_param.loc['Justiça Gratuita', 1]).strip().upper() == 'SIM'
    
    pag_voluntario_str = ""
    try:
        val_pv = df_param.loc['Pagamento Voluntário 15d', 1]
        if pd.notna(val_pv):
            pag_voluntario_str = str(val_pv).strip().upper()
    except KeyError: pass
    
    houve_inadimplemento = (pag_voluntario_str in ['NÃO', 'NAO'])
    
    hon_perc = 0.0
    try:
        val_perc = df_param.loc['Honorários Sucumbência', 1]
        if pd.notna(val_perc) and str(val_perc).strip() != '':
            hon_perc = float(val_perc) / 100
    except KeyError: pass
        
    hon_fixo = 0.0
    try:
        val_fixo = df_param.loc['Honorários Fixos (R$)', 1]
        if pd.notna(val_fixo) and str(val_fixo).strip() != '':
            hon_fixo = float(val_fixo)
    except KeyError: pass

    data_sentenca_raw = None
    try:
        data_sentenca_raw = df_param.loc['Data da Sentença', 1]
    except KeyError: pass

    # Lendo abas validadas direto do objeto 'xls' na memória (mais rápido e seguro)
    df_danos = pd.read_excel(xls, sheet_name='Danos')
    df_danos = df_danos.dropna(subset=['Data Desembolso', 'Valor Histórico'], how='any')

    df_custas = pd.read_excel(xls, sheet_name='Custas')
    df_custas = df_custas.dropna(subset=['Data Desembolso', 'Valor Histórico'], how='any')
    
    data_calculo = pd.Timestamp.today()
    total_danos = 0.0
    
    for idx, row in df_danos.iterrows():
        data = pd.to_datetime(row['Data Desembolso'], dayfirst=True)
        valor = float(row['Valor Histórico'])
        regra = str(row.get('Regra', '')).strip().upper()
        
        if regra == 'R1':
            df_danos.at[idx, 'Desc_Regra'] = "TJMG + Juros 1% a.m."
            fator = calcular_tjmg_juros(tabela_tjmg, data, data_calculo)
        elif regra == 'R2':
            df_danos.at[idx, 'Desc_Regra'] = "Selic Pura"
            fator = calcular_selic_pura(df_bcb, data, data_calculo)
        elif regra == 'R3':
            df_danos.at[idx, 'Desc_Regra'] = "TJMG + 1% até 08/2024; após, Selic"
            fator = calcular_tjmg_selic(tabela_tjmg, df_bcb, data, data_calculo)
        elif regra == 'R4':
            df_danos.at[idx, 'Desc_Regra'] = "TJMG + 1% até 08/2024; após, Lei 14.905"
            fator = calcular_tjmg_leinova(tabela_tjmg, df_bcb, data, data_calculo)
        elif regra == 'R5':
            df_danos.at[idx, 'Desc_Regra'] = "Selic até 08/2024; após, Lei 14.905"
            fator = calcular_selic_leinova(df_bcb, data, data_calculo)
        elif regra == 'R6':
            df_danos.at[idx, 'Desc_Regra'] = "Lei 14.905/24 (IPCA + Taxa Legal)"
            fator = calcular_leinova_pura(df_bcb, data, data_calculo)
        else:
            df_danos.at[idx, 'Desc_Regra'] = "Regra não identificada."
            fator = 1.0

        valor_corr = valor * fator
        df_danos.at[idx, 'Valor Atualizado'] = valor_corr
        total_danos += valor_corr

    total_custas = 0.0
    for idx, row in df_custas.iterrows():
        data = pd.to_datetime(row['Data Desembolso'], dayfirst=True)
        valor = float(row['Valor Histórico'])
        
        fator = calcular_custas_ipca_taxalegal(df_bcb, data, data_transito_c, teve_transito, data_calculo) 
        
        if teve_transito:
            df_custas.at[idx, 'Desc_Regra'] = "IPCA + Taxa Legal (Juros do Trânsito)"
        else:
            df_custas.at[idx, 'Desc_Regra'] = "IPCA (Sem Juros)"
        
        valor_corr = valor * fator
        exigivel = 0.0 if jg else valor_corr
        df_custas.at[idx, 'Valor Atualizado'] = valor_corr
        df_custas.at[idx, 'Exigível'] = exigivel
        total_custas += exigivel

    subtotal = total_danos + total_custas
    
    if hon_fixo > 0:
        if pd.notna(data_sentenca_raw):
            data_sentenca = pd.to_datetime(data_sentenca_raw, dayfirst=True)
            
            valor_honorarios_calc = atualizar_honorarios_fixos(hon_fixo, data_sentenca, data_transito_c, tabela_tjmg, df_bcb, data_calculo)
            str_transito = data_transito_c.strftime('%d/%m/%Y') if teve_transito else "Sem trânsito"
            desc_hon = f"Honorários Fixos (CM desde {data_sentenca.strftime('%d/%m/%Y')} | Juros desde {str_transito}):"
        else:
            valor_honorarios_calc = hon_fixo
            desc_hon = "Honorários Fixos (Sem Atualização - Falta 'Data da Sentença'):"
    else:
        valor_honorarios_calc = subtotal * hon_perc
        desc_hon = f"Honorários de Sucumbência ({hon_perc*100:g}%):"

    if jg:
        desc_hon = desc_hon[:-1] + " [Inexigível - JG]:"
        
    valor_honorarios_exigivel = 0.0 if jg else valor_honorarios_calc
    base_multa = subtotal + valor_honorarios_exigivel
    
    valor_multa = (base_multa * 0.10) if houve_inadimplemento else 0.0
    
    honorarios_523 = 0.0
    if houve_inadimplemento and not jg:
        honorarios_523 = (base_multa * 0.10)
    
    total_geral = base_multa + valor_multa + honorarios_523
    
    # Atualiza o arquivo final mas salvando forçadamente em formato XLSX moderno
    gerar_laudo_excel(processo, teve_transito, jg, df_danos, df_custas, subtotal, valor_honorarios_exigivel, desc_hon, valor_multa, honorarios_523, total_geral, arquivo_saida, houve_inadimplemento)

# --- 5. GERAÇÃO DO LAUDO FORMATADO ---
def gerar_laudo_excel(processo, teve_transito, jg, df_danos, df_custas, subtotal, hon, desc_hon, multa, hon_523, total, arquivo_saida, houve_inadimplemento):
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
    ws['A5'] = "Tipo de Título:"
    ws['B5'] = "Sentença (Com Trânsito)" if teve_transito else "Acordo / Sem Trânsito"
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
        if jg:
            desc_523 = "Honorários Fase Cumprimento Art. 523 CPC (10%) [Inexigível - JG]:"
        
        add_total(desc_523, hon_523)
        
    add_total("TOTAL GERAL DEVIDO:", total, True, True)

    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 16
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 54 
    ws.column_dimensions['F'].width = 20
    
    # Salva sempre como XLSX final garantindo retrocompatibilidade
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

        regras = [
            ("R1", "Tabela TJMG + Juros de 1% a.m. (critério único)"),
            ("R2", "Taxa Selic (critério único) durante todo o período"),
            ("R3", "TJMG + Juros 1% a.m. até 08/2024; após, Taxa Selic"),
            ("R4", "TJMG + Juros 1% a.m. até 08/2024; após, Lei 14.905/24"),
            ("R5", "Selic até 08/2024; após, Lei 14.905/24"),
            ("R6", "Lei 14.905/24: IPCA + Taxa Legal (critério único)"),
            ("Custas", "Padrão automático (IPCA + Taxa Legal)")
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
            title="Selecione a Planilha do Cliente", 
            filetypes=[
                ("Planilhas (Excel/LibreOffice)", "*.xlsx *.xls *.ods"),
                ("Todos os Arquivos", "*.*")
            ]
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
                messagebox.showinfo("Sucesso", "Liquidação calculada e laudo gerado com sucesso!")
            self.root.after(0, sucesso)
        except Exception as e:
            erro_msg = str(e)
            def erro():
                self.lbl_status.config(text="Erro durante o cálculo.", fg="red")
                self.btn_processar.config(state="normal")
                messagebox.showerror("Atenção - Erro na Planilha", erro_msg) 
            self.root.after(0, erro)

if __name__ == "__main__":
    app = tk.Tk()
    gui = NashGUI(app)
    app.mainloop()