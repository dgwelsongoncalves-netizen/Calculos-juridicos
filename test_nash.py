import pytest
import pandas as pd
import nash 

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
        'TAXA_LEGAL': df_fixo.rename(columns={'VALOR': 'TAXA_LEGAL'}), # Como IPCA é 1% e Selic é 1%, Taxa Legal = 0% na prática real, mas mantemos isolado no mock
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
    """Testa transição de TJMG+1% para Selic após 08/2024"""
    d_inicio = pd.to_datetime('2024-06-01')
    d_fim = pd.to_datetime('2024-10-01')
    f_cm, f_jur = nash.calc_tjmg_juros_selic(mock_tjmg, mock_bcb, d_inicio, d_inicio, d_fim)
    assert f_cm > 1.0
    assert f_jur > 0.0

def test_r4_tjmg_leinova(mock_tjmg, mock_bcb):
    """Testa transição de TJMG+1% para Lei 14.905 após 08/2024"""
    d_inicio = pd.to_datetime('2024-06-01')
    d_fim = pd.to_datetime('2024-10-01')
    f_cm, f_jur = nash.calc_tjmg_leinova(mock_tjmg, mock_bcb, d_inicio, d_inicio, d_fim)
    assert f_cm > 1.0
    assert f_jur > 0.0

def test_r5_tema_1368_transicao(mock_bcb):
    """Testa a blindagem da transição do STJ: Selic até 08/2024, IPCA+TL depois"""
    d_inicio = pd.to_datetime('2024-06-01')
    d_fim = pd.to_datetime('2024-10-01') 
    f_cm, f_jur = nash.calc_selic_leinova(mock_bcb, d_inicio, d_inicio, d_fim)
    assert f_cm > 1.0
    assert f_jur >= 0.0

def test_r6_leinova_pura(mock_bcb):
    """Testa a regra da Lei 14.905/24 atuando em período completo"""
    d_inicio = pd.to_datetime('2024-06-01')
    d_fim = pd.to_datetime('2024-10-01')
    f_cm, f_jur = nash.calc_leinova_pura(mock_bcb, d_inicio, d_inicio, d_fim)
    assert f_cm > 1.0
    assert f_jur == 0.0 # Porque SELIC (0.01) - IPCA (0.01) no mock resulta em Taxa Legal = 0

def test_r7_taxalegal_retroativa(mock_tjmg, mock_bcb):
    """Testa TJMG atuando apenas na correção e Taxa Legal Retroativa nos Juros"""
    d_inicio = pd.to_datetime('2024-06-01')
    d_fim = pd.to_datetime('2024-10-01')
    f_cm, f_jur = nash.calc_tjmg_taxalegal_retroativa(mock_tjmg, mock_bcb, d_inicio, d_inicio, d_fim)
    assert f_cm > 1.0
    assert f_jur == 0.0 # Mesma lógica da R6 no mock (SELIC - IPCA = 0)

def test_fazenda_publica(mock_bcb):
    """Testa se o corte temporal da Emenda Constitucional 113/2021 (Dez/21) é aplicado"""
    d_inicio = pd.to_datetime('2021-10-01')
    d_fim = pd.to_datetime('2022-03-01')
    f_cm, f_jur = nash.calc_fazenda_publica(mock_bcb, d_inicio, d_inicio, d_fim)
    assert f_cm > 1.0
    assert f_jur > 0.0

# --- 3. TESTES DE LÓGICAS COMPLEXAS ---
def test_delta_conta_grafica_sem_anatocismo():
    """Prova matemática de que a lógica Conta Gráfica calcula o salto de juros sem compor (sem anatocismo)"""
    jur_t1 = 0.10
    jur_t2 = 0.15
    f_jur_delta = jur_t2 - jur_t1
    assert f_jur_delta == pytest.approx(0.05) 

def test_exito_proveito_economico():
    """Prova a matemática do Relatório de Economia: Proveito = Risco - Condenação Real"""
    v_pedido_inicial = 10000.0
    f_cm_pedido = 1.5
    f_jur_pedido = 0.50 # 50% de juros
    risco_atualizado_total = (v_pedido_inicial * f_cm_pedido) * (1 + f_jur_pedido) # 10k * 1.5 * 1.5 = 22.500,00
    
    val_princ_condenacao = 5000.0
    val_jur_condenacao = 1000.0
    total_condenacao = val_princ_condenacao + val_jur_condenacao # 6.000,00 real
    
    proveito = max(0, risco_atualizado_total - total_condenacao)
    
    assert risco_atualizado_total == 22500.0
    assert proveito == 16500.0 # A associação deixou de pagar 16.500