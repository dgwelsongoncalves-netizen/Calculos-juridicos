import pandas as pd
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import sys
import os
import logging
import traceback

__version__ = "2.6.3" # Auditoria: Fazenda Termo Juros + Alerta Excesso + Validação TJMG + Timeline Conta Gráfica

# --- 1. CONFIGURAÇÕES BASE ---
if getattr(sys, 'frozen', False):
    PASTA_APP = Path(sys.executable).parent
    PASTA_TEMP = Path(sys._MEIPASS) if hasattr(sys, '_MEIPASS') else PASTA_APP
else:
    PASTA_APP = Path(__file__).parent
    PASTA_TEMP = PASTA_APP

PASTA_TABELAS = PASTA_APP / 'Tabelas_Oficiais'
ARQUIVO_TJMG = PASTA_TABELAS / 'tabela_tjmg.xlsx'

# --- CONFIGURAÇÃO DO LOG ---
ARQUIVO_LOG = PASTA_APP / "nash_system.log"
logging.basicConfig(
    filename=ARQUIVO_LOG,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%d/%m/%Y %H:%M:%S",
    encoding="utf-8"
)
logging.info(f"--- Nash System v{__version__} Iniciado ---")

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

def carregar_taxas_bcb(data_minima):
    try:
        from bcb import sgs
        
        if pd.isna(data_minima):
            str_data_inicio = '2010-01-01' 
        else:
            str_data_inicio = f"{data_minima.year}-{data_minima.month:02d}-01"
            
        logging.info(f"Baixando taxas do BCB a partir de: {str_data_inicio}")

        df_selic = sgs.get({'SELIC': 4390}, start=str_data_inicio)
        df_selic = df_selic.apply(pd.to_numeric, errors='coerce').fillna(0) / 100.0
        df_selic.index = df_selic.index.to_period('M').to_timestamp()
        
        df_ipca = sgs.get({'IPCA': 433}, start=str_data_inicio)
        df_ipca = df_ipca.apply(pd.to_numeric, errors='coerce').fillna(0) / 100.0
        df_ipca.index = df_ipca.index.to_period('M').to_timestamp()
        
        inicio_tl = max(pd.to_datetime('2024-08-01'), pd.to_datetime(str_data_inicio))
        df_tl = sgs.get({'TAXA_LEGAL': 29543}, start=inicio_tl.strftime('%Y-%m-%d'))
        df_tl = df_tl.apply(pd.to_numeric, errors='coerce').fillna(0) / 100.0
        df_tl.index = df_tl.index.to_period('M').to_timestamp()

        df_ipca_e = sgs.get({'IPCA_E': 10764}, start=str_data_inicio)
        df_ipca_e = df_ipca_e.apply(pd.to_numeric, errors='coerce').fillna(0) / 100.0
        df_ipca_e.index = df_ipca_e.index.to_period('M').to_timestamp()

        df_poup = sgs.get({'POUPANCA': 195}, start=str_data_inicio)
        df_poup = df_poup.resample('MS').first() 
        df_poupanca = df_poup.apply(pd.to_numeric, errors='coerce').fillna(0) / 100.0
        df_poupanca.index = df_poupanca.index.to_period('M').to_timestamp()
        
        return {
            'SELIC': df_selic, 
            'IPCA': df_ipca, 
            'TAXA_LEGAL': df_tl,
            'IPCA_E': df_ipca_e,
            'POUPANCA': df_poupanca
        }
    except Exception as e:
        raise Exception(f"Sem conexão com a internet ou API do Banco Central indisponível.\nDetalhe: {e}")
    
def obter_fator_tjmg(df_tjmg, data_inicio, data_fim):
    if df_tjmg is None: return 1.0
    d_ini = pd.to_datetime(f"{data_inicio.year}-{data_inicio.month:02d}-01")
    d_fim = pd.to_datetime(f"{data_fim.year}-{data_fim.month:02d}-01")
    
    ultimo_mes_tabela = df_tjmg.index[-1]
    if d_fim > ultimo_mes_tabela + pd.DateOffset(months=2):
        raise ValueError(f"A Tabela TJMG está defasada! O cálculo atual exige a competência de {d_fim.strftime('%m/%Y')}, mas o arquivo 'tabela_tjmg.xlsx' disponível na pasta só vai até {ultimo_mes_tabela.strftime('%m/%Y')}. Atualize a tabela para evitar prejuízos no cálculo.")

    try:
        idx_ini = float(df_tjmg.loc[d_ini, 'ÍNDICE'])
    except KeyError:
        idx_ini = float(df_tjmg['ÍNDICE'].iloc[-1])
    try:
        idx_fim = float(df_tjmg.loc[d_fim, 'ÍNDICE'])
    except KeyError:
        idx_fim = float(df_tjmg['ÍNDICE'].iloc[-1])
    return idx_ini / idx_fim if idx_fim != 0 else 1.0

# --- 3. MOTORES MATEMÁTICOS SEPARADOS ---
def calc_tjmg_juros(df_tjmg, data_cm, data_juros, data_calculo):
    fator_cm = obter_fator_tjmg(df_tjmg, data_cm, data_calculo)
    fator_juros = 0.0
    if pd.notna(data_juros) and data_juros <= data_calculo:
        meses = (data_calculo.year - data_juros.year) * 12 + (data_calculo.month - data_juros.month)
        fator_juros = max(0, meses) * 0.01
    return fator_cm, fator_juros

def calc_selic_pura(df_bcb, data_cm, data_juros, data_calculo):
    if df_bcb is None: return 1.0, 0.0
    data_cm_mes = pd.to_datetime(f"{data_cm.year}-{data_cm.month:02d}-01")
    data_calc_mes = pd.to_datetime(f"{data_calculo.year}-{data_calculo.month:02d}-01")
    df_ipca = df_bcb['IPCA']
    df_selic = df_bcb['SELIC']
    
    if pd.isna(data_juros) or data_juros > data_calculo:
        mask = (df_ipca.index >= data_cm_mes) & (df_ipca.index <= data_calc_mes)
        return (1 + df_ipca.loc[mask, 'IPCA']).prod(), 0.0

    data_juros_mes = pd.to_datetime(f"{data_juros.year}-{data_juros.month:02d}-01")
    fator_ipca = 1.0
    if data_cm_mes < data_juros_mes:
        mask_ipca = (df_ipca.index >= data_cm_mes) & (df_ipca.index < data_juros_mes)
        fator_ipca = (1 + df_ipca.loc[mask_ipca, 'IPCA']).prod()

    inicio_selic = max(data_cm_mes, data_juros_mes)
    mask_selic = (df_selic.index >= inicio_selic) & (df_selic.index <= data_calc_mes)
    fator_selic = (1 + df_selic.loc[mask_selic, 'SELIC']).prod()
    
    fator_cm = fator_ipca
    juros_efetivo = fator_selic - 1.0
    return fator_cm, juros_efetivo

def calc_tjmg_leinova(df_tjmg, df_bcb, data_cm, data_juros, data_calculo):
    data_corte = pd.to_datetime("2024-08-30")
    corte_mes = pd.to_datetime("2024-08-01")

    if data_cm >= data_corte:
        return calc_leinova_pura(df_bcb, data_cm, data_juros, data_calculo)

    fator_cm_fase1 = obter_fator_tjmg(df_tjmg, data_cm, data_corte)
    juros_fase1 = 0.0
    if pd.notna(data_juros) and data_juros < data_corte:
        meses = (data_corte.year - data_juros.year) * 12 + (data_corte.month - data_juros.month)
        juros_fase1 = max(0, meses) * 0.01

    data_calc_mes = pd.to_datetime(f"{data_calculo.year}-{data_calculo.month:02d}-01")
    fator_cm_fase2 = 1.0
    juros_fase2 = 0.0
    
    if data_calc_mes > corte_mes and df_bcb is not None:
        df_ipca = df_bcb['IPCA']
        df_tl = df_bcb['TAXA_LEGAL']
        mask_ipca = (df_ipca.index > corte_mes) & (df_ipca.index <= data_calc_mes)
        fator_cm_fase2 = (1 + df_ipca.loc[mask_ipca, 'IPCA']).prod()

        if pd.notna(data_juros) and data_juros <= data_calculo:
            data_juros_mes = pd.to_datetime(f"{data_juros.year}-{data_juros.month:02d}-01")
            inicio_juros_fase2 = max(corte_mes, data_juros_mes)
            mask_tl = (df_tl.index > inicio_juros_fase2) & (df_tl.index <= data_calc_mes)
            juros_fase2 = df_tl.loc[mask_tl, 'TAXA_LEGAL'].sum()

    fator_cm_final = fator_cm_fase1 * fator_cm_fase2
    juros_final = juros_fase1 + juros_fase2
    return fator_cm_final, juros_final

def calc_fazenda_publica(df_bcb, data_cm, data_juros, data_calculo):
    """
    Motor Fazenda Pública (Restrito a Responsabilidade Civil/Não-Tributário).
    Até 30/11/2021: IPCA-E + Poupança (Tema 810 STF)
    A partir de 01/12/2021: Taxa Selic Pura segmentada (EC 113/2021)
    """
    if df_bcb is None: return 1.0, 0.0
    
    data_corte_ec113 = pd.to_datetime("2021-11-30")
    mes_corte_ec113 = pd.to_datetime("2021-12-01")
    
    df_ipca_e = df_bcb['IPCA_E']
    df_poupanca = df_bcb['POUPANCA']
    df_selic = df_bcb['SELIC']

    # FASE 1: Antes da EC 113/2021 (IPCA-E + Poupança)
    fator_cm_fase1 = 1.0
    juros_fase1 = 0.0
    
    if data_cm <= data_corte_ec113:
        data_cm_mes = pd.to_datetime(f"{data_cm.year}-{data_cm.month:02d}-01")
        fim_fase1 = min(pd.to_datetime(f"{data_calculo.year}-{data_calculo.month:02d}-01"), mes_corte_ec113)
        
        mask_ipca_e = (df_ipca_e.index >= data_cm_mes) & (df_ipca_e.index < fim_fase1)
        fator_cm_fase1 = (1 + df_ipca_e.loc[mask_ipca_e, 'IPCA_E']).prod()
        
        if pd.notna(data_juros) and data_juros <= data_corte_ec113:
            data_juros_mes = pd.to_datetime(f"{data_juros.year}-{data_juros.month:02d}-01")
            mask_poup = (df_poupanca.index >= data_juros_mes) & (df_poupanca.index < fim_fase1)
            juros_fase1 = df_poupanca.loc[mask_poup, 'POUPANCA'].sum()

    # FASE 2: Após a EC 113/2021 (Selic Pura Segmentada)
    data_calc_mes = pd.to_datetime(f"{data_calculo.year}-{data_calculo.month:02d}-01")
    fator_cm_fase2 = 1.0
    
    if data_calc_mes >= mes_corte_ec113:
        inicio_fase2 = max(mes_corte_ec113, pd.to_datetime(f"{data_cm.year}-{data_cm.month:02d}-01"))
        
        if pd.isna(data_juros) or data_juros <= mes_corte_ec113:
            # Mora já consolidada: aplica Selic direta
            mask_selic = (df_selic.index >= inicio_fase2) & (df_selic.index <= data_calc_mes)
            fator_cm_fase2 = (1 + df_selic.loc[mask_selic, 'SELIC']).prod()
        else:
            # Evento/Citação ocorreu durante a fase Selic. 
            data_juros_mes = pd.to_datetime(f"{data_juros.year}-{data_juros.month:02d}-01")
            
            # 1. Correção pura por IPCA-E até o início dos juros
            if inicio_fase2 < data_juros_mes:
                mask_ipca_e_f2 = (df_ipca_e.index >= inicio_fase2) & (df_ipca_e.index < data_juros_mes)
                fator_cm_fase2 = (1 + df_ipca_e.loc[mask_ipca_e_f2, 'IPCA_E']).prod()
            
            # 2. Selic (Juros+Correção) a partir da mora
            inicio_selic = max(inicio_fase2, data_juros_mes)
            mask_selic = (df_selic.index >= inicio_selic) & (df_selic.index <= data_calc_mes)
            fator_selic = (1 + df_selic.loc[mask_selic, 'SELIC']).prod()
            
            fator_cm_fase2 = fator_cm_fase2 * fator_selic

    fator_cm_final = fator_cm_fase1 * fator_cm_fase2
    return fator_cm_final, juros_fase1

def calc_leinova_pura(df_bcb, data_cm, data_juros, data_calculo):
    if df_bcb is None: return 1.0, 0.0
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
        
    return fator_ipca, juros_tl

# --- 4. PROCESSAMENTO CENTRAL (Com Suporte a Conta Gráfica) ---
def executar_nash(caminho_entrada, arquivo_saida):
    tabela_tjmg = carregar_tjmg()
    
    xls = pd.ExcelFile(caminho_entrada)
    abas_esperadas = ['Parametros', 'Danos', 'Custas']
    if not all(aba in xls.sheet_names for aba in abas_esperadas):
        raise Exception("Arquivo inválido. Verifique se o arquivo possui as abas obrigatórias: 'Parametros', 'Danos' e 'Custas'.")

    df_param = pd.read_excel(xls, sheet_name='Parametros', header=None, index_col=0)
    
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
    
    is_fazenda = str(get_param('Fazenda Pública', 'NÃO')).strip().upper() == 'SIM'
    if is_fazenda:
        houve_inadimplemento = False
        
    hon_perc = float(get_param('Honorários Sucumbência', 0.0)) / 100
    hon_fixo = float(get_param('Honorários Fixos (R$)', 0.0))
    
    termo_juros_raw = str(get_param('Termo Inicial Juros', 'DESEMBOLSO')).strip().upper()
    data_citacao = get_param('Data da Citação', is_date=True)
    data_evento = get_param('Data do Evento', is_date=True)

    df_danos = pd.read_excel(xls, sheet_name='Danos').dropna(subset=['Data Desembolso', 'Valor Histórico'], how='any')
    df_danos['Data Desembolso'] = pd.to_datetime(df_danos['Data Desembolso'], format='mixed', dayfirst=True, errors='coerce')
    
    df_custas = pd.read_excel(xls, sheet_name='Custas').dropna(subset=['Data Desembolso', 'Valor Histórico'], how='any')
    df_custas['Data Desembolso'] = pd.to_datetime(df_custas['Data Desembolso'], format='mixed', dayfirst=True, errors='coerce')
    
    todas_as_datas = pd.Series(dtype='datetime64[ns]')
    if not df_danos.empty: todas_as_datas = pd.concat([todas_as_datas, df_danos['Data Desembolso']])
    if not df_custas.empty: todas_as_datas = pd.concat([todas_as_datas, df_custas['Data Desembolso']])
    if pd.notna(data_citacao): todas_as_datas = pd.concat([todas_as_datas, pd.Series([data_citacao])])
    if pd.notna(data_evento): todas_as_datas = pd.concat([todas_as_datas, pd.Series([data_evento])])
    
    data_minima_processo = todas_as_datas.min() if not todas_as_datas.empty else pd.NaT
    df_bcb = carregar_taxas_bcb(data_minima_processo)
    
    for idx, row in df_danos.iterrows():
        regra_txt = str(row.get('Regra', '')).strip().upper()
        if is_fazenda: df_danos.at[idx, 'Desc_Regra'] = "Fazenda Pública (Tema 810 / EC 113)"
        elif regra_txt == 'R1': df_danos.at[idx, 'Desc_Regra'] = "TJMG + Juros 1% a.m."
        elif regra_txt == 'R2': df_danos.at[idx, 'Desc_Regra'] = "Taxa Selic"
        elif regra_txt == 'R3': df_danos.at[idx, 'Desc_Regra'] = "TJMG + 1% até 08/24; após, Taxa Selic"
        elif regra_txt == 'R4': df_danos.at[idx, 'Desc_Regra'] = "TJMG + 1% até 08/24; após, Lei 14.905"
        elif regra_txt == 'R5': df_danos.at[idx, 'Desc_Regra'] = "Selic até 08/24; após, Lei 14.905"
        elif regra_txt == 'R6': df_danos.at[idx, 'Desc_Regra'] = "Lei 14.905/24 (IPCA + Taxa Legal)"
        else: df_danos.at[idx, 'Desc_Regra'] = regra_txt if regra_txt else "Selic"
        
    for idx, row in df_custas.iterrows():
        df_custas.at[idx, 'Desc_Regra'] = "IPCA + Taxa Legal (Juros do Trânsito)" if teve_transito else "IPCA (Sem Juros)"

    try:
        df_deducoes = pd.read_excel(xls, sheet_name='Deducoes').dropna(subset=['Data bloqueio/deposito', 'Valor'], how='any')
        df_deducoes['Data bloqueio/deposito'] = pd.to_datetime(df_deducoes['Data bloqueio/deposito'], format='mixed', dayfirst=True, errors='coerce')
        df_deducoes = df_deducoes.sort_values('Data bloqueio/deposito')
        tem_deducao = not df_deducoes.empty
    except Exception:
        tem_deducao = False
        df_deducoes = pd.DataFrame()

    data_calculo = pd.Timestamp.today()
    
    # MODO 1: SEM DEDUÇÕES
    if not tem_deducao:
        total_princ_danos = 0.0
        total_juros_danos = 0.0
        
        for idx, row in df_danos.iterrows():
            data_cm = pd.to_datetime(row['Data Desembolso'], dayfirst=True)
            valor = float(row['Valor Histórico'])
            regra = str(row.get('Regra', '')).strip().upper()
            
            data_juros = data_cm
            if 'CITA' in termo_juros_raw: data_juros = data_citacao
            elif 'EVENTO' in termo_juros_raw: data_juros = data_evento
                
            if is_fazenda: f_cm, f_jur = calc_fazenda_publica(df_bcb, data_cm, data_juros, data_calculo)
            elif regra == 'R1': f_cm, f_jur = calc_tjmg_juros(tabela_tjmg, data_cm, data_juros, data_calculo)
            elif regra == 'R4': f_cm, f_jur = calc_tjmg_leinova(tabela_tjmg, df_bcb, data_cm, data_juros, data_calculo)
            elif regra == 'R6': f_cm, f_jur = calc_leinova_pura(df_bcb, data_cm, data_juros, data_calculo)
            else: f_cm, f_jur = calc_selic_pura(df_bcb, data_cm, data_juros, data_calculo)

            val_princ = valor * f_cm
            val_jur = val_princ * f_jur
            df_danos.at[idx, 'Valor Atualizado'] = val_princ + val_jur
            total_princ_danos += val_princ
            total_juros_danos += val_jur

        total_princ_custas = 0.0
        total_juros_custas = 0.0
        for idx, row in df_custas.iterrows():
            data_cm_custas = pd.to_datetime(row['Data Desembolso'], dayfirst=True)
            valor = float(row['Valor Histórico'])
            if is_fazenda: f_cm, f_jur = calc_fazenda_publica(df_bcb, data_cm_custas, data_transito_c if teve_transito else pd.NaT, data_calculo)
            else: f_cm, f_jur = calc_leinova_pura(df_bcb, data_cm_custas, data_transito_c if teve_transito else pd.NaT, data_calculo) 
            val_princ = valor * f_cm
            val_jur = val_princ * f_jur
            exigivel = 0.0 if jg else (val_princ + val_jur)
            df_custas.at[idx, 'Exigível'] = exigivel
            total_princ_custas += (valor * f_cm) if not jg else 0.0
            total_juros_custas += val_jur if not jg else 0.0

        subtotal_princ = total_princ_danos + total_princ_custas
        subtotal_juros = total_juros_danos + total_juros_custas
        
        hon_calc_princ = 0.0
        hon_calc_juros = 0.0
        if hon_fixo > 0 and pd.notna(data_sentenca):
            if is_fazenda: f_cm, f_jur = calc_fazenda_publica(df_bcb, data_sentenca, data_transito_c if teve_transito else pd.NaT, data_calculo)
            else: f_cm, f_jur = calc_leinova_pura(df_bcb, data_sentenca, data_transito_c if teve_transito else pd.NaT, data_calculo)
            hon_calc_princ = hon_fixo * f_cm
            hon_calc_juros = hon_calc_princ * f_jur
        elif hon_fixo > 0:
            hon_calc_princ = hon_fixo
        else:
            hon_calc_princ = (subtotal_princ + subtotal_juros) * hon_perc
            
        hon_exigivel_princ = 0.0 if jg else hon_calc_princ
        hon_exigivel_juros = 0.0 if jg else hon_calc_juros

        base_multa = subtotal_princ + subtotal_juros + hon_exigivel_princ + hon_exigivel_juros
        valor_multa = (base_multa * 0.10) if houve_inadimplemento else 0.0
        hon_523 = (base_multa * 0.10) if (houve_inadimplemento and not jg) else 0.0
        
        total_geral = base_multa + valor_multa + hon_523
        
        gerar_laudo_excel(processo, teve_transito, jg, df_danos, df_custas, (subtotal_princ+subtotal_juros), (hon_exigivel_princ+hon_exigivel_juros), valor_multa, hon_523, total_geral, arquivo_saida, houve_inadimplemento, termo_juros_raw, [])

    # MODO 2: COM DEDUÇÕES (Amortização em Escada Art. 354 CC)
    else:
        df_danos['Valor Atualizado'] = df_danos['Valor Histórico']
        df_custas['Exigível'] = df_custas['Valor Histórico'] if not jg else 0.0
        
        historico_conta_grafica = []
        saldo_principal = 0.0
        saldo_juros = 0.0
        
        data_corte = df_deducoes.iloc[0]['Data bloqueio/deposito']
        
        # Processa apenas o que ocorreu até o 1º bloqueio
        for _, row in df_danos.iterrows():
            data_cm = pd.to_datetime(row['Data Desembolso'], dayfirst=True)
            if data_cm > data_corte: continue
            valor = float(row['Valor Histórico'])
            regra = str(row.get('Regra', '')).strip().upper()
            data_juros = data_cm
            if 'CITA' in termo_juros_raw: data_juros = data_citacao
            elif 'EVENTO' in termo_juros_raw: data_juros = data_evento
                
            if is_fazenda: f_cm, f_jur = calc_fazenda_publica(df_bcb, data_cm, data_juros, data_corte)
            elif regra == 'R1': f_cm, f_jur = calc_tjmg_juros(tabela_tjmg, data_cm, data_juros, data_corte)
            elif regra == 'R4': f_cm, f_jur = calc_tjmg_leinova(tabela_tjmg, df_bcb, data_cm, data_juros, data_corte)
            elif regra == 'R6': f_cm, f_jur = calc_leinova_pura(df_bcb, data_cm, data_juros, data_corte)
            else: f_cm, f_jur = calc_selic_pura(df_bcb, data_cm, data_juros, data_corte)
            val_princ = valor * f_cm
            saldo_principal += val_princ
            saldo_juros += val_princ * f_jur
            
        for _, row in df_custas.iterrows():
            data_cm_custas = pd.to_datetime(row['Data Desembolso'], dayfirst=True)
            if data_cm_custas > data_corte or jg: continue
            if is_fazenda: f_cm, f_jur = calc_fazenda_publica(df_bcb, data_cm_custas, data_transito_c if teve_transito else pd.NaT, data_corte)
            else: f_cm, f_jur = calc_leinova_pura(df_bcb, data_cm_custas, data_transito_c if teve_transito else pd.NaT, data_corte) 
            val_princ = row['Valor Histórico'] * f_cm
            saldo_principal += val_princ
            saldo_juros += val_princ * f_jur
            
        historico_conta_grafica.append((data_corte, "Subtotal (Principal + Juros)", saldo_principal, saldo_juros, 0.0, saldo_principal+saldo_juros))
        
        # Honorários / Multas na consolidação
        hon_suc_princ = 0.0
        hon_suc_jur = 0.0
        if hon_fixo > 0 and pd.notna(data_sentenca) and not jg:
            if is_fazenda: f_cm, f_jur = calc_fazenda_publica(df_bcb, data_sentenca, data_transito_c if teve_transito else pd.NaT, data_corte)
            else: f_cm, f_jur = calc_leinova_pura(df_bcb, data_sentenca, data_transito_c if teve_transito else pd.NaT, data_corte)
            hon_suc_princ = hon_fixo * f_cm
            hon_suc_jur = hon_suc_princ * f_jur
        elif hon_fixo > 0 and not jg:
            hon_suc_princ = hon_fixo
        elif not jg:
            hon_suc_princ = (saldo_principal + saldo_juros) * hon_perc
            
        if hon_suc_princ > 0 or hon_suc_jur > 0:
            saldo_principal += hon_suc_princ
            saldo_juros += hon_suc_jur
            historico_conta_grafica.append((data_corte, "(+) Honorários Sucumbenciais", hon_suc_princ, hon_suc_jur, 0.0, saldo_principal+saldo_juros))
            
        if houve_inadimplemento:
            base_multa = saldo_principal + saldo_juros
            multa_523 = base_multa * 0.10
            saldo_principal += multa_523
            historico_conta_grafica.append((data_corte, "(+) Multa Art. 523 CPC (10%)", multa_523, 0.0, 0.0, saldo_principal+saldo_juros))
            
            if not jg:
                hon_523 = base_multa * 0.10
                saldo_principal += hon_523
                historico_conta_grafica.append((data_corte, "(+) Honorários Art. 523 CPC (10%)", hon_523, 0.0, 0.0, saldo_principal+saldo_juros))
                
        historico_conta_grafica.append((data_corte, "DÍVIDA CONSOLIDADA (Pré-Bloqueio)", saldo_principal, saldo_juros, 0.0, saldo_principal+saldo_juros))
        
        ultima_data = data_corte
        for _, row in df_deducoes.iterrows():
            data_ded = pd.to_datetime(row['Data bloqueio/deposito'])
            valor_ded = float(row['Valor'])
            
            # Auditoria 3: Injeção de despesas posteriores ao corte antes de debitar o bloqueio
            despesas_tardias = df_danos[(df_danos['Data Desembolso'] > ultima_data) & (df_danos['Data Desembolso'] <= data_ded)]
            for _, nd in despesas_tardias.iterrows():
                d_dano = pd.to_datetime(nd['Data Desembolso'])
                if is_fazenda: f_cm, f_jur = calc_fazenda_publica(df_bcb, ultima_data, ultima_data, d_dano)
                else: f_cm, f_jur = calc_leinova_pura(df_bcb, ultima_data, ultima_data, d_dano) 
                saldo_principal = saldo_principal * f_cm
                saldo_juros += saldo_principal * f_jur
                v_add = float(nd['Valor Histórico'])
                saldo_principal += v_add
                historico_conta_grafica.append((d_dano, f"(+) Nova Inclusão: {nd['Descrição']}", saldo_principal, saldo_juros, 0.0, saldo_principal+saldo_juros))
                ultima_data = d_dano

            if data_ded > ultima_data:
                if is_fazenda: f_cm, f_jur = calc_fazenda_publica(df_bcb, ultima_data, ultima_data, data_ded)
                else: f_cm, f_jur = calc_leinova_pura(df_bcb, ultima_data, ultima_data, data_ded) 
                saldo_principal = saldo_principal * f_cm
                saldo_juros += saldo_principal * f_jur
                historico_conta_grafica.append((data_ded, "Atualização do Saldo", saldo_principal, saldo_juros, 0.0, saldo_principal+saldo_juros))
                ultima_data = data_ded
            
            abate_juros = min(valor_ded, saldo_juros)
            saldo_juros -= abate_juros
            abate_princ = valor_ded - abate_juros
            
            # Auditoria 4: Restituição de Excedente de Bloqueio
            if abate_princ > saldo_principal:
                excesso = abate_princ - saldo_principal
                saldo_principal = 0.0
                historico_conta_grafica.append((data_ded, "(-) Bloqueio Judicial", 0.0, saldo_juros, valor_ded, 0.0))
                historico_conta_grafica.append((data_ded, "(!) EXCESSO DE EXECUÇÃO / RESTITUIR", 0.0, 0.0, excesso, 0.0))
            else:
                saldo_principal -= abate_princ
                historico_conta_grafica.append((data_ded, "(-) Bloqueio Judicial", saldo_principal, saldo_juros, valor_ded, saldo_principal+saldo_juros))
        
        # Traz as últimas pendências até hoje
        despesas_finais = df_danos[df_danos['Data Desembolso'] > ultima_data]
        for _, nd in despesas_finais.iterrows():
            d_dano = pd.to_datetime(nd['Data Desembolso'])
            if is_fazenda: f_cm, f_jur = calc_fazenda_publica(df_bcb, ultima_data, ultima_data, d_dano)
            else: f_cm, f_jur = calc_leinova_pura(df_bcb, ultima_data, ultima_data, d_dano) 
            saldo_principal = saldo_principal * f_cm
            saldo_juros += saldo_principal * f_jur
            v_add = float(nd['Valor Histórico'])
            saldo_principal += v_add
            historico_conta_grafica.append((d_dano, f"(+) Nova Inclusão: {nd['Descrição']}", saldo_principal, saldo_juros, 0.0, saldo_principal+saldo_juros))
            ultima_data = d_dano

        if ultima_data < data_calculo:
            if is_fazenda: f_cm, f_jur = calc_fazenda_publica(df_bcb, ultima_data, ultima_data, data_calculo)
            else: f_cm, f_jur = calc_leinova_pura(df_bcb, ultima_data, ultima_data, data_calculo)
            saldo_principal = saldo_principal * f_cm
            saldo_juros += saldo_principal * f_jur
            historico_conta_grafica.append((data_calculo, "Atualização Final (Hoje)", saldo_principal, saldo_juros, 0.0, saldo_principal+saldo_juros))
            
        gerar_laudo_excel(processo, teve_transito, jg, df_danos, df_custas, 0, 0, 0, 0, (saldo_principal+saldo_juros), arquivo_saida, houve_inadimplemento, termo_juros_raw, historico_conta_grafica)

# --- 5. GERAÇÃO DO LAUDO FORMATADO ---
def gerar_laudo_excel(processo, teve_transito, jg, df_danos, df_custas, subtotal, hon, multa, hon_523, total, arquivo_saida, houve_inadimplemento, termo_juros_raw, historico):
    wb = Workbook()
    ws = wb.active
    ws.title = "Laudo de Liquidação"

    f_titulo = Font(name="Arial", size=14, bold=True, color="FFFFFF")
    f_negrito = Font(name="Arial", size=11, bold=True)
    f_normal = Font(name="Arial", size=11)
    fundo_escuro = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fundo_cinza = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    borda = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    moeda = 'R$ #,##0.00'

    ws.merge_cells('A1:F1')
    ws['A1'] = f"LAUDO DE CÁLCULO JUDICIAL - NASH SYSTEM"
    ws['A1'].font = f_titulo; ws['A1'].fill = fundo_escuro; ws['A1'].alignment = Alignment(horizontal="center")
    
    ws.merge_cells('A2:F2')
    ws['A2'] = f"Gerado pelo Nash System v{__version__} em {pd.Timestamp.today().strftime('%d/%m/%Y às %H:%M')}"
    ws['A2'].alignment = Alignment(horizontal="right"); ws['A2'].font = Font(name="Arial", size=9, italic=True)

    ws['A4'] = "Processo:"; ws['B4'] = processo
    ws['A5'] = "Termo Juros (Danos):"; ws['B5'] = termo_juros_raw.title()
    ws['A6'] = "Justiça Gratuita:"; ws['B6'] = "DEFERIDA (Custas Inexigíveis)" if jg else "NÃO REQUERIDA / INDEFERIDA"
    for r in range(4, 7): ws.cell(row=r, column=1).font = f_negrito

    linha = 8
    
    ws.merge_cells(f'A{linha}:F{linha}')
    ws[f'A{linha}'] = "1. DANOS MATERIAIS / NOTAS"
    ws[f'A{linha}'].font = f_negrito; ws[f'A{linha}'].fill = fundo_cinza; linha += 1

    cabs = ['ID / Folha', 'Descrição', 'Data Desemb.', 'Valor Histórico', 'Índice Aplicado', 'Valor Atualizado']
    for i, t in enumerate(cabs, 1):
        ws.cell(row=linha, column=i, value=t).font = f_negrito; ws.cell(row=linha, column=i).border = borda
    linha += 1

    for _, r in df_danos.iterrows():
        ws.cell(row=linha, column=1, value=r['ID / Folha']).border = borda
        ws.cell(row=linha, column=2, value=r['Descrição']).border = borda
        ws.cell(row=linha, column=3, value=r['Data Desembolso'].strftime('%d/%m/%Y')).border = borda
        ws.cell(row=linha, column=4, value=r['Valor Histórico']).number_format = moeda; ws.cell(row=linha, column=4).border = borda
        ws.cell(row=linha, column=5, value=r.get('Desc_Regra', '')).border = borda
        val_display = r['Valor Histórico'] if historico else r['Valor Atualizado']
        ws.cell(row=linha, column=6, value=val_display).number_format = moeda; ws.cell(row=linha, column=6).border = borda
        linha += 1

    linha += 1

    ws.merge_cells(f'A{linha}:F{linha}')
    ws[f'A{linha}'] = "2. CUSTAS E DESPESAS PROCESSUAIS"
    ws[f'A{linha}'].font = f_negrito; ws[f'A{linha}'].fill = fundo_cinza; linha += 1
    
    cabs_custas = ['ID / Folha', 'Descrição', 'Data', 'Valor Histórico', 'Índice Aplicado', 'Atualizado Exigível']
    for i, t in enumerate(cabs_custas, 1):
        ws.cell(row=linha, column=i, value=t).font = f_negrito; ws.cell(row=linha, column=i).border = borda
    linha += 1
    
    for _, r in df_custas.iterrows():
        ws.cell(row=linha, column=1, value=r['ID / Folha']).border = borda
        ws.cell(row=linha, column=2, value=r['Descrição']).border = borda
        ws.cell(row=linha, column=3, value=r['Data Desembolso'].strftime('%d/%m/%Y')).border = borda
        ws.cell(row=linha, column=4, value=r['Valor Histórico']).number_format = moeda; ws.cell(row=linha, column=4).border = borda
        ws.cell(row=linha, column=5, value=r.get('Desc_Regra', '')).border = borda
        val_display_c = r['Valor Histórico'] if (historico and not jg) else r['Exigível']
        if jg: val_display_c = 0.0
        ws.cell(row=linha, column=6, value=val_display_c).number_format = moeda; ws.cell(row=linha, column=6).border = borda
        linha += 1

    linha += 2

    if historico:
        ws.merge_cells(f'A{linha}:F{linha}')
        ws[f'A{linha}'] = "3. EVOLUÇÃO DA DÍVIDA (AMORTIZAÇÃO CONTA GRÁFICA - Art. 354 CC)"
        ws[f'A{linha}'].font = f_titulo; ws[f'A{linha}'].fill = fundo_escuro; ws[f'A{linha}'].alignment = Alignment(horizontal="center"); linha += 1
        
        cabs_cg = ['Data', 'Evento', 'Principal Corrigido', 'Juros Acumulados', 'Valor Pago/Bloqueio', 'Saldo Devedor Total']
        for i, t in enumerate(cabs_cg, 1):
            ws.cell(row=linha, column=i, value=t).font = f_negrito; ws.cell(row=linha, column=i).border = borda; ws.cell(row=linha, column=i).fill = fundo_cinza
        linha += 1
        
        for data, evento, princ, jur, pago, saldo in historico:
            ws.cell(row=linha, column=1, value=data.strftime('%d/%m/%Y')).border = borda
            ws.cell(row=linha, column=2, value=evento).border = borda
            if "Restituir" in str(evento).title():
                ws.cell(row=linha, column=2).font = Font(name="Arial", size=11, bold=True, color="FF0000") # Destaque vermelho
            elif "Bloqueio" in str(evento):
                ws.cell(row=linha, column=2).font = f_negrito
            else:
                ws.cell(row=linha, column=2).font = f_normal
                
            ws.cell(row=linha, column=3, value=princ).number_format = moeda; ws.cell(row=linha, column=3).border = borda
            ws.cell(row=linha, column=4, value=jur).number_format = moeda; ws.cell(row=linha, column=4).border = borda
            ws.cell(row=linha, column=5, value=pago if pago > 0 else "-").number_format = moeda if pago > 0 else 'General'; ws.cell(row=linha, column=5).border = borda
            ws.cell(row=linha, column=6, value=saldo).number_format = moeda; ws.cell(row=linha, column=6).border = borda; ws.cell(row=linha, column=6).font = f_negrito
            linha += 1
            
        # --- RODAPÉ DE DESTAQUE ---
        linha += 1
        ws.merge_cells(f'A{linha}:E{linha}')
        ws.cell(row=linha, column=1, value="SALDO DEVEDOR FINAL DA EXECUÇÃO:").alignment = Alignment(horizontal="right")
        ws.cell(row=linha, column=1).font = f_titulo; ws.cell(row=linha, column=1).fill = fundo_escuro
        
        saldo_final_val = historico[-1][5] if historico else 0.0
        ws.cell(row=linha, column=6, value=saldo_final_val).number_format = moeda
        ws.cell(row=linha, column=6).font = f_titulo; ws.cell(row=linha, column=6).fill = fundo_escuro
        for i in range(1, 7): ws.cell(row=linha, column=i).border = borda
            
    else:
        ws.merge_cells(f'A{linha}:F{linha}')
        ws[f'A{linha}'] = "3. RESUMO DA LIQUIDAÇÃO"
        ws[f'A{linha}'].font = f_titulo; ws[f'A{linha}'].fill = fundo_escuro; ws[f'A{linha}'].alignment = Alignment(horizontal="center"); linha += 1

        def add_total(desc, val, negrito=False, dest=False):
            nonlocal linha
            ws.merge_cells(f'A{linha}:E{linha}')
            ws.cell(row=linha, column=1, value=desc).alignment = Alignment(horizontal="right")
            ws.cell(row=linha, column=6, value=val).number_format = moeda
            fonte = f_negrito if negrito else f_normal
            ws.cell(row=linha, column=1).font = fonte; ws.cell(row=linha, column=6).font = fonte
            if dest: ws.cell(row=linha, column=1).fill = fundo_cinza; ws.cell(row=linha, column=6).fill = fundo_cinza
            for i in range(1, 7): ws.cell(row=linha, column=i).border = borda
            linha += 1

        add_total("SUBTOTAL (Principal + Custas Atualizados):", subtotal, True)
        add_total("Honorários Acumulados:", hon)
        if houve_inadimplemento:
            add_total("Multa Art. 523 CPC (10%):", multa)
            add_total("Honorários Art. 523 CPC (10%):", hon_523)
        add_total("TOTAL GERAL DEVIDO:", total, True, True)

    # --- AUTO-AJUSTE INTELIGENTE DE COLUNAS ---
    larguras_minimas = {'A': 15, 'B': 30, 'C': 18, 'D': 18, 'E': 40, 'F': 22}
    for letra_col, larg_min in larguras_minimas.items():
        max_len = larg_min
        for row in range(8, linha + 1):
            valor_celula = ws[f'{letra_col}{row}'].value
            if valor_celula:
                tamanho_atual = len(str(valor_celula))
                if tamanho_atual > max_len:
                    max_len = tamanho_atual
        ws.column_dimensions[letra_col].width = min(max_len + 2, 65)
    
    # --- ÁREA DE IMPRESSÃO PARA PDF ---
    ws.print_area = f'A1:F{linha}' 
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    
    nome_saida = str(arquivo_saida)
    if nome_saida.endswith(('.ods', '.xls')): nome_saida = nome_saida.rsplit('.', 1)[0] + '.xlsx'
    
    path_final = Path(nome_saida)
    if path_final.exists():
        pasta = path_final.parent
        nome_base = path_final.stem
        ext = path_final.suffix
        contador = 1
        while True:
            novo_caminho = pasta / f"{nome_base} ({contador}){ext}"
            if not novo_caminho.exists():
                path_final = novo_caminho
                break
            contador += 1
            
    wb.save(str(path_final))

class NashGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Nash System - Liquidação (v{__version__})")
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
            ("R1", "TJMG + Juros de 1% a.m."),
            ("R2", "Taxa Selic (critério único)"),
            ("R3", "TJMG + Juros 1% a.m. até 08/24; após, Taxa Selic"),
            ("R4", "TJMG + Juros 1% a.m. até 08/24; após, Lei 14.905/24"),
            ("R5", "Selic até 08/24; após, Lei 14.905/24"),
            ("R6", "Lei 14.905/24: IPCA + Taxa Legal"),
            ("Juros", "Dinâmico (Citação, Evento Danoso ou Desembolso)"),
            ("Fazenda", "Tema 810 STF e EC 113/2021 (Automático se 'SIM' na aba)"),
            ("NOVO", "Amortização Conta Gráfica Automática (Aba 'Deducoes')")
        ]
        for regra, desc in regras:
            linha_ui = tk.Frame(frame_regras)
            linha_ui.pack(anchor="w", pady=1)
            tk.Label(linha_ui, text=f"{regra}: ", font=("Arial", 9, "bold")).pack(side="left")
            tk.Label(linha_ui, text=desc, font=("Arial", 9)).pack(side="left")

        self.lbl_status = tk.Label(root, text="Aguardando arquivo...", font=("Arial", 10), fg="gray")
        self.lbl_status.pack(pady=10)
        self.btn_processar = tk.Button(root, text="📂 Selecionar Planilha e Calcular", font=("Arial", 11, "bold"), bg="#2F5597", fg="white", padx=15, pady=8, command=self.iniciar_processo)
        self.btn_processar.pack()

    def iniciar_processo(self):
        caminho = filedialog.askopenfilename(title="Selecione a Planilha", filetypes=[("Planilhas", "*.xlsx *.xls *.ods")])
        if not caminho: return
        caminho_entrada = Path(caminho)
        arquivo_saida = caminho_entrada.parent / f"Laudo_{caminho_entrada.name}"
        self.lbl_status.config(text="Processando cálculo...", fg="blue")
        self.btn_processar.config(state="disabled")
        self.root.update()
        threading.Thread(target=self.processar_em_background, args=(caminho_entrada, arquivo_saida)).start()

    def processar_em_background(self, caminho_entrada, arquivo_saida):
        logging.info(f"Iniciando cálculo para: {caminho_entrada.name}")
        try:
            executar_nash(caminho_entrada, arquivo_saida)
            def sucesso():
                self.lbl_status.config(text=f"Laudo gerado com sucesso.", fg="green")
                self.btn_processar.config(state="normal")
                logging.info(f"Cálculo concluído com sucesso. Salvo como: {arquivo_saida.name}")
                messagebox.showinfo("Sucesso", "Cálculo processado!")
            self.root.after(0, sucesso)
        except Exception as e:
            tb_str = traceback.format_exc()
            logging.error(f"Erro ao processar {caminho_entrada.name}:\n{tb_str}")
            
            erro_msg = str(e)
            
            def erro():
                self.lbl_status.config(text="Erro crítico (Veja o log).", fg="red")
                self.btn_processar.config(state="normal")
                mensagem = f"O cálculo foi interrompido pelo seguinte erro:\n\n{erro_msg}\n\nUm registro detalhado foi salvo no arquivo 'nash_system.log' na pasta do programa."
                messagebox.showerror("Atenção - Erro no Processamento", mensagem) 
            self.root.after(0, erro)

if __name__ == "__main__":
    app = tk.Tk()
    gui = NashGUI(app)
    app.mainloop()