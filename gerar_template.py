import pandas as pd
from pathlib import Path

# Garante que a pasta Processos_Entrada existe
pasta_entrada = Path("Processos_Entrada")
pasta_entrada.mkdir(exist_ok=True)

caminho_arquivo = pasta_entrada / "entrada_nash.xlsx"

# Aba 1: Parâmetros (Formato chave-valor)
param_data = {
    "Chave": [
        "Processo", 
        "Data do Trânsito", 
        "Justiça Gratuita", 
        "Pagamento Voluntário 15d", 
        "Honorários Sucumbência"
    ],
    "Valor": [
        "0000000-00.2026.8.13.0024", 
        "",  # Vazio para testar a lógica de Acordo Pré-Sentença
        "Não", 
        "Não", 
        10
    ]
}
df_param = pd.DataFrame(param_data)

# Aba 2: Danos e Notas (Categoria 1)
danos_data = {
    "ID / Folha": ["Id. 123456", "Id. 123457"],
    "Descrição": ["Conserto da motocicleta", "Capacete e vestuário"],
    "Data Desembolso": ["15/01/2023", "18/01/2023"],
    "Valor Histórico": [4500.00, 850.00],
    "Regra": ["TJMG_ate_LeiNova", "TJMG_ate_LeiNova"]
}
df_danos = pd.DataFrame(danos_data)

# Aba 3: Custas e Despesas Processuais (Categoria 2)
custas_data = {
    "ID / Folha": ["Id. 234567", "Id. 234568"],
    "Descrição": ["Custas Iniciais", "Guia de Oficial de Justiça"],
    "Data Desembolso": ["20/01/2023", "25/02/2023"],
    "Valor Histórico": [250.00, 95.50]
}
df_custas = pd.DataFrame(custas_data)

# Gerando o arquivo Excel com múltiplas abas usando openpyxl
print("Fabricando a planilha modelo do Nash...")
with pd.ExcelWriter(caminho_arquivo, engine='openpyxl') as writer:
    # Salvando sem o cabeçalho na aba de Parâmetros para ficar mais limpo
    df_param.to_excel(writer, sheet_name='Parametros', index=False, header=False)
    df_danos.to_excel(writer, sheet_name='Danos', index=False)
    df_custas.to_excel(writer, sheet_name='Custas', index=False)

print(f"✓ Sucesso! Arquivo gerado em: {caminho_arquivo.absolute()}")
print("Você já pode abrir o LibreOffice Calc para conferir o resultado.")