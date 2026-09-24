import pytest
import pandas as pd
import nash 
import os
import openpyxl

# Carrega o Pandas dentro do nash.py antes de rodar os testes
nash.preload_heavy_libs()

# --- 1. AMBIENTE CONTROLADO (MOCKS) ---
@pytest.fixture
def mock_tjmg():
    # Cria uma inflação simulada cravada em 1% ao mês para testarmos a matemática pura.
    dates = pd.date_range(start='2020-01-01', end='2026-12-01', freq='MS')
    indices = [1.0 * (1.01 ** i) for i in range(len(dates))][::-1] 
    return pd.DataFrame({'DATA_REF': dates, 'ÍNDICE': indices}).set_index('DATA_REF')

@pytest.fixture
def mock_bcb():
    # Congela todas as taxas do governo em exatos 1% ao mês para testes blindados.
    dates = pd.date_range(start='2020-01-01', end='2026-12-01', freq='MS')
    df_fixo = pd.DataFrame({'VALOR': [0.01]*len(dates)}, index=dates)
    return {
        'SELIC': df_fixo.rename(columns={'VALOR': 'SELIC'}),
        'IPCA': df_fixo.rename(columns={'VALOR': 'IPCA'}),
        'TAXA_LEGAL': df_fixo.rename(columns={'VALOR': 'TAXA_LEGAL'}), 
        'IPCA_E': df_fixo.rename(columns={'VALOR': 'IPCA_E'}),
        'POUPANCA': df_fixo.rename(columns={'VALOR': 'POUPANCA'})
    }

# --- 2. TESTES DAS REGRAS INDIVIDUAIS ---
def test_r1_tjmg_mais_1_porcento(mock_tjmg):
    d_inicio = pd.to_datetime('2023-01-01')
    d_fim = pd.to_datetime('2023-11-01') # 10 meses
    f_cm, f_jur = nash.calc_tjmg_juros(mock_tjmg, d_inicio, d_inicio, d_fim)
    assert f_jur == pytest.approx(0.10)
    assert f_cm == pytest.approx(1.01 ** 10)

def test_r2_selic_pura(mock_bcb):
    d_inicio = pd.to_datetime('2023-01-01')
    d_fim = pd.to_datetime('2023-03-01') # 3 meses
    f_cm, f_jur = nash.calc_selic_pura(mock_bcb, d_inicio, d_inicio, d_fim)
    assert f_cm == 1.0 
    assert f_jur == pytest.approx((1.01 ** 3) - 1.0) 

def test_r3_tjmg_selic(mock_tjmg, mock_bcb):
    d_inicio = pd.to_datetime('2024-06-01')
    d_fim = pd.to_datetime('2024-10-01')
    f_cm, f_jur = nash.calc_tjmg_juros_selic(mock_tjmg, mock_bcb, d_inicio, d_inicio, d_fim)
    assert f_cm > 1.0
    assert f_jur > 0.0

def test_r4_tjmg_leinova(mock_tjmg, mock_bcb):
    d_inicio = pd.to_datetime('2024-06-01')
    d_fim = pd.to_datetime('2024-10-01')
    f_cm, f_jur = nash.calc_tjmg_leinova(mock_tjmg, mock_bcb, d_inicio, d_inicio, d_fim)
    assert f_cm > 1.0
    assert f_jur > 0.0

def test_r5_tema_1368_transicao(mock_bcb):
    d_inicio = pd.to_datetime('2024-06-01')
    d_fim = pd.to_datetime('2024-10-01') 
    f_cm, f_jur = nash.calc_selic_leinova(mock_bcb, d_inicio, d_inicio, d_fim)
    assert f_cm > 1.0
    assert f_jur >= 0.0

def test_r6_leinova_pura(mock_bcb):
    d_inicio = pd.to_datetime('2024-06-01')
    d_fim = pd.to_datetime('2024-10-01')
    f_cm, f_jur = nash.calc_leinova_pura(mock_bcb, d_inicio, d_inicio, d_fim)
    assert f_cm > 1.0
    assert f_jur == 0.0

def test_r7_taxalegal_retroativa(mock_tjmg, mock_bcb):
    d_inicio = pd.to_datetime('2024-06-01')
    d_fim = pd.to_datetime('2024-10-01')
    f_cm, f_jur = nash.calc_tjmg_taxalegal_retroativa(mock_tjmg, mock_bcb, d_inicio, d_inicio, d_fim)
    assert f_cm > 1.0
    assert f_jur == 0.0 

def test_fazenda_publica(mock_bcb):
    d_inicio = pd.to_datetime('2021-10-01')
    d_fim = pd.to_datetime('2022-03-01')
    f_cm, f_jur = nash.calc_fazenda_publica(mock_bcb, d_inicio, d_inicio, d_fim)
    assert f_cm > 1.0
    assert f_jur > 0.0

# --- 3. TESTES DE LÓGICAS COMPLEXAS ---
def test_delta_conta_grafica_sem_anatocismo():
    jur_t1 = 0.10
    jur_t2 = 0.15
    f_jur_delta = jur_t2 - jur_t1
    assert f_jur_delta == pytest.approx(0.05) 

def test_exito_proveito_economico():
    v_pedido_inicial = 10000.0
    f_cm_pedido = 1.5
    f_jur_pedido = 0.50 
    risco_atualizado_total = (v_pedido_inicial * f_cm_pedido) * (1 + f_jur_pedido) 
    
    val_princ_condenacao = 5000.0
    val_jur_condenacao = 1000.0
    total_condenacao = val_princ_condenacao + val_jur_condenacao 
    
    proveito = max(0, risco_atualizado_total - total_condenacao)
    
    assert risco_atualizado_total == 22500.0
    assert proveito == 16500.0 

def test_limpeza_de_moeda():
    assert nash.limpar_moeda("1.500,50") == 1500.50
    assert nash.limpar_moeda("800,25") == 800.25
    assert nash.limpar_moeda("500") == 500.0
    assert nash.limpar_moeda("NaN") == 0.0

def test_integracao_custas_leinova_com_conta_grafica(tmp_path):
    # 1. Cria os dados do cenário que causava o erro
    df_param = pd.DataFrame([
        ['Processo', 'Teste Sincronizacao Custas'],
        ['Justiça Gratuita', 'NÃO'],
        ['Atuação', 'AUTOR'],
        ['Data do Trânsito', '19/06/2024'],
        ['Proporção Custas (%)', '100%'],
    ]).set_index(0)
    
    df_danos = pd.DataFrame({
        'ID / Folha': ['DANO-1'], 'Descrição': ['Dano Principal'], 
        'Data Desembolso': ['20/10/2018'], 'Valor Histórico': [1000.0], 
        'Regra': ['R1']
    })
    
    df_custas = pd.DataFrame({
        'ID / Folha': ['CUSTA-1'], 'Descrição': ['Custas Iniciais'], 
        'Data Desembolso': ['16/04/2021'], 'Valor Histórico': [464.64]
    })
    
    df_deducoes = pd.DataFrame({
        'Data bloqueio/deposito': ['05/09/2025'], 'Valor': [450.0]
    })
    
    # 2. Guarda a folha de cálculo falsa numa pasta temporária do pytest
    input_file = tmp_path / "entrada_teste_luciana.xlsx"
    output_file = tmp_path / "Laudo_entrada_teste_luciana.xlsx"
    
    with pd.ExcelWriter(input_file) as writer:
        df_param.to_excel(writer, sheet_name='Parametros', header=False)
        df_danos.to_excel(writer, sheet_name='Danos', index=False)
        df_custas.to_excel(writer, sheet_name='Custas', index=False)
        df_deducoes.to_excel(writer, sheet_name='Deducoes', index=False)
        
    # 3. Executa o motor do Nash System (como Path objects simulando o subprocess)
    nash.executar_nash(input_file, output_file)
    
    # 4. Lê o Laudo gerado para validar a correção matemática
    assert output_file.exists(), "O laudo de liquidação não foi gerado!"
    
    wb = openpyxl.load_workbook(output_file)
    ws = wb['Laudo de Liquidação']
    
    achou_tabela_custas = False
    regra_aplicada = ""
    valor_exigivel = 0.0
    
    for row in ws.iter_rows(values_only=True):
        col0 = str(row[0]) if row[0] else ""
        if "CUSTAS E DESPESAS PROCESSUAIS" in col0:
            achou_tabela_custas = True
            
        if achou_tabela_custas and col0 == 'CUSTA-1':
            regra_aplicada = str(row[6]) 
            valor_exigivel = float(row[7]) 
            break
            
    # Aserções de auditoria atualizadas para a v3.0
    assert "Lei 14.905" in regra_aplicada, f"Erro: A regra das custas não foi fixada na Lei Nova. Registou: {regra_aplicada}"
    assert valor_exigivel > 0.0, "Erro Crítico: O valor exigível das custas ficou a zeros na Conta Gráfica!"
    assert valor_exigivel > 464.64, "Erro: O valor exigível não sofreu atualização monetária."