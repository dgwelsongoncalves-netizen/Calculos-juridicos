import pandas as pd
from pathlib import Path

# Configurando os caminhos das pastas
PASTA_BASE = Path(__file__).parent
PASTA_TABELAS = PASTA_BASE / 'Tabelas_Oficiais'
ARQUIVO_TJMG = PASTA_TABELAS / 'tabela_tjmg.xlsx'

def carregar_tabela_tjmg():
    print("Iniciando a leitura da tabela oficial do TJMG...")
    try:
        # Pula as 8 linhas de cabeçalho do tribunal e força os nomes corretos
        df_tjmg = pd.read_excel(ARQUIVO_TJMG, sheet_name='Plan1', skiprows=8, names=['ANO', 'MÊS', 'ÍNDICE'])
        df_tjmg = df_tjmg.dropna(subset=['MÊS', 'ÍNDICE'])
        
        # Dicionário de tradução (com remoção de espaços em branco por segurança)
        df_tjmg['MÊS'] = df_tjmg['MÊS'].astype(str).str.strip()
        meses = {
            'Janeiro': '01', 'Fevereiro': '02', 'Março': '03', 'Abril': '04',
            'Maio': '05', 'Junho': '06', 'Julho': '07', 'Agosto': '08',
            'Setembro': '09', 'Outubro': '10', 'Novembro': '11', 'Dezembro': '12'
        }
        df_tjmg['MÊS_NUM'] = df_tjmg['MÊS'].map(meses)
        
        # Remove qualquer linha de rodapé do TJMG que não seja um mês válido
        df_tjmg = df_tjmg.dropna(subset=['MÊS_NUM'])
        
        # Cria a data real e a define como o "índice de busca" da tabela
        df_tjmg['DATA_REF'] = pd.to_datetime(
            df_tjmg['ANO'].astype(int).astype(str) + '-' + df_tjmg['MÊS_NUM'] + '-01'
        )
        df_tjmg = df_tjmg[['DATA_REF', 'ÍNDICE']].set_index('DATA_REF')
        
        print("✓ Tabela do TJMG carregada e limpa com sucesso!")
        return df_tjmg

    except FileNotFoundError:
        print(f"ERRO: Não encontrei o arquivo em {ARQUIVO_TJMG}")
    except Exception as e:
        print(f"Erro inesperado ao ler a tabela: {e}")

def calcular_fator_tjmg(df_tjmg, data_nota_str, data_limite_str="2024-08-01"):
    """
    Calcula o multiplicador exato travando a atualização na data limite.
    """
    data_nota = pd.to_datetime(data_nota_str)
    data_limite = pd.to_datetime(data_limite_str)
    
    # Extrai os dois índices diretamente da tabela do tribunal
    indice_nota = df_tjmg.loc[data_nota, 'ÍNDICE']
    indice_limite = df_tjmg.loc[data_limite, 'ÍNDICE']
    
    # A mágica matemática da divisão
    return indice_nota / indice_limite

# --- TESTANDO O MOTOR ---
if __name__ == "__main__":
    tabela = carregar_tabela_tjmg()
    
    if tabela is not None:
        # Exemplo prático: Uma nota fiscal de Janeiro de 2020
        data_exemplo = '2020-01-01'
        fator_trava = calcular_fator_tjmg(tabela, data_exemplo)
        
        print(f"\n--- TESTE DE TRAVA DE DATA (LEI NOVA) ---")
        print(f"Data da despesa: {data_exemplo} | Marco da Lei: 2024-08-01")
        print(f"Índice bruto da despesa: {tabela.loc['2020-01-01', 'ÍNDICE']:.6f}")
        print(f"Índice bruto do limite:  {tabela.loc['2024-08-01', 'ÍNDICE']:.6f}")
        print(f"-> Fator multiplicador final: {fator_trava:.6f}")
        
        valor_historico = 1000.00
        valor_corrigido = valor_historico * fator_trava
        print(f"-> Exemplo de cálculo: R$ {valor_historico:.2f} corrigidos pelo TJMG param em R$ {valor_corrigido:.2f}")