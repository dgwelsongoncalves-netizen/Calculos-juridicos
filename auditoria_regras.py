import pandas as pd
from pathlib import Path
import nash
import os

def criar_planilha_teste(caminho_saida, regra_dano, is_fazenda=False, is_acordo=False):
    # 1. Parâmetros Gerais
    parametros = {
        'Processo': [f"TESTE-AUDITORIA-{regra_dano}"],
        'Atuação': ['RÉU'],
        'Data da Sentença': ['10/05/2024'],
        'Termo Inicial Juros': ['CITAÇÃO'],
        'Data da Citação': ['15/01/2021'],
        'Justiça Gratuita': ['NÃO'],
        'Pagamento Voluntário 15d': ['NÃO'],
        'Fazenda Pública': ['SIM' if is_fazenda else 'NÃO'],
        'Honorários Sucumbência (%)': ['10'],
        'Base Honorários': ['CONDENAÇÃO'],
        'Proporção Honorários (%)': ['100'],
        'Proporção Custas (%)': ['100']
    }
    df_param = pd.DataFrame(parametros).T

    # 2. Danos (R$ 10.000,00)
    danos = {
        'ID / Folha': ['Doc 1'],
        'Descrição': [f'Indenização Principal ({regra_dano})'],
        'Data Desembolso': ['15/01/2021'],
        'Valor Histórico': [10000.00],
        'Regra': [regra_dano],
        'Data Juros': ['15/01/2021'],
        'Valor Pedido Inicial': [15000.00],
        'Data do Pedido': ['15/01/2021']
    }
    df_danos = pd.DataFrame(danos)

    # 3. Custas (R$ 1.000 e R$ 500)
    custas = {
        'ID / Folha': ['Guia 1', 'Guia 2'],
        'Descrição': ['Custas Iniciais', 'Diligência Oficial'],
        'Data Desembolso': ['15/01/2021', '10/06/2022'],
        'Valor Histórico': [1000.00, 500.00]
    }
    df_custas = pd.DataFrame(custas)

    # 4. Deduções / Conta Gráfica (Depósito de R$ 200)
    deducoes = {
        'ID / Folha': ['Depósito Judicial 1'],
        'Data bloqueio/deposito': ['20/05/2023'],
        'Valor': [200.00]
    }
    df_deducoes = pd.DataFrame(deducoes)

    # Escrever no Excel
    with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
        df_param.to_excel(writer, sheet_name='Parametros', header=False)
        df_danos.to_excel(writer, sheet_name='Danos', index=False)
        df_custas.to_excel(writer, sheet_name='Custas', index=False)
        df_deducoes.to_excel(writer, sheet_name='Deducoes', index=False)
        
        # 5. Se for teste de Acordo, injetar as abas extras
        if is_acordo:
            acordo_params = {
                'Data do Inadimplemento': ['10/01/2025'],
                'Regra de Atualização Pós-Quebra': [regra_dano],
                'Valor Original Confessado (Sem Desconto)': [13000.00],
                'Valor do Desconto (A Reverter)': [1500.00],
                'Multa Moratória Contratual (%)': ['10'],
                'Honorários de Retomada/Execução (%)': ['20']
            }
            df_acordo_p = pd.DataFrame(acordo_params).T
            df_acordo_p.to_excel(writer, sheet_name='Acordo_Params', header=False)

            acordo_pagas = {
                'Identificação / Parcela': ['Parcela 1 Paga'],
                'Data do Pagamento': ['20/05/2023'],
                'Valor Pago (R$)': [200.00] # Os mesmos R$ 200 como parcela paga
            }
            df_acordo_pagas = pd.DataFrame(acordo_pagas)
            df_acordo_pagas.to_excel(writer, sheet_name='Acordo_Pagas', index=False)

def rodar_auditoria():
    pasta_auditoria = Path("Relatorios_Auditoria")
    pasta_auditoria.mkdir(exist_ok=True)
    
    testes = [
        ("1_Teste_Regra_R1", "R1", False, False),
        ("2_Teste_Regra_R2", "R2", False, False),
        ("3_Teste_Regra_R3", "R3", False, False),
        ("4_Teste_Regra_R4", "R4", False, False),
        ("5_Teste_Regra_R5", "R5", False, False),
        ("6_Teste_Regra_R6", "R6", False, False),
        ("7_Teste_Regra_R7", "R7", False, False),
        ("8_Teste_Fazenda_Publica", "R6", True, False), # Fazenda Pública
        ("9_Teste_Quebra_de_Acordo", "R6", False, True) # Acordo Não Cumprido
    ]
    
    print("Iniciando Auditoria do Motor Matemático (Nash System)...\n")
    
    for nome_arquivo, regra, is_fazenda, is_acordo in testes:
        caminho_template = pasta_auditoria / f"template_{nome_arquivo}.xlsx"
        caminho_laudo = pasta_auditoria / f"Laudo_{nome_arquivo}.xlsx"
        
        print(f"-> Gerando cenário: {nome_arquivo.replace('_', ' ')}...")
        
        # 1. Cria a planilha fria com os dados exatos
        criar_planilha_teste(caminho_template, regra, is_fazenda, is_acordo)
        
        # 2. Pede ao motor para calcular
        try:
            nash.executar_nash(str(caminho_template), str(caminho_laudo))
            print("   [OK] Laudo PDF e XLSX gerados com sucesso!")
        except Exception as e:
            print(f"   [ERRO] Falha ao processar: {e}")
            
    print("\nAuditoria concluída! Verifique a pasta 'Relatorios_Auditoria'.")

if __name__ == "__main__":
    rodar_auditoria()