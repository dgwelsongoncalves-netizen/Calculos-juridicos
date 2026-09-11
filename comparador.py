import nash
import pandas as pd
import argparse

# Carrega o Pandas dentro do nash.py para liberar as bibliotecas
nash.preload_heavy_libs()

def simular_regras(valor, data_inicio, data_fim):
    print("Conectando aos servidores do Tribunal e do Banco Central...")
    try:
        tjmg = nash.carregar_tjmg()
        bcb = nash.carregar_taxas_bcb(pd.to_datetime(data_inicio, dayfirst=True))
    except Exception as e:
        print(f"Erro ao baixar dados: {e}")
        return

    d_ini = pd.to_datetime(data_inicio, dayfirst=True)
    d_fim = pd.to_datetime(data_fim, dayfirst=True)
    
    print(f"\n=============================================")
    print(f"💰 SIMULADOR DE BOLSO - NASH SYSTEM")
    print(f"Valor Histórico: R$ {valor:,.2f}")
    print(f"Período: {data_inicio} até {data_fim}")
    print(f"=============================================\n")
    
    # Mapeamento universal de todas as regras matemáticas do Nash
    regras = {
        'R1 (TJMG + 1%)': lambda: nash.calc_tjmg_juros(tjmg, d_ini, d_ini, d_fim),
        'R2 (Selic Pura)': lambda: nash.calc_selic_pura(bcb, d_ini, d_ini, d_fim),
        'R3 (TJMG + 1% até 08/24; após Selic)': lambda: nash.calc_tjmg_juros_selic(tjmg, bcb, d_ini, d_ini, d_fim),
        'R4 (TJMG + 1% até 08/24; após Lei 14.905)': lambda: nash.calc_tjmg_leinova(tjmg, bcb, d_ini, d_ini, d_fim),
        'R5 (Selic até 08/24; após Lei 14.905)': lambda: nash.calc_selic_leinova(bcb, d_ini, d_ini, d_fim),
        'R6 (Lei 14.905/24 Pura)': lambda: nash.calc_leinova_pura(bcb, d_ini, d_ini, d_fim),
        'R7 (TJMG CM + Taxa Legal Retroativa)': lambda: nash.calc_tjmg_taxalegal_retroativa(tjmg, bcb, d_ini, d_ini, d_fim),
        'Fazenda Pública (EC 113)': lambda: nash.calc_fazenda_publica(bcb, d_ini, d_ini, d_fim)
    }
    
    for nome, funcao in regras.items():
        f_cm, f_jur = funcao()
        principal_corrigido = valor * f_cm
        valor_juros = principal_corrigido * f_jur
        total = principal_corrigido + valor_juros
        
        print(f"📌 {nome}")
        print(f"   Fator Correção: {f_cm:.6f}")
        print(f"   Juros Acumulados: {f_jur*100:.2f}%")
        print(f"   Principal Corrigido: R$ {principal_corrigido:,.2f}")
        print(f"   Valor dos Juros: R$ {valor_juros:,.2f}")
        print(f"   TOTAL DEVIDO: R$ {total:,.2f}\n")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Compara as regras do Nash System via Terminal")
    parser.add_argument("valor", type=float, help="Valor histórico do dano (ex: 1500.50)")
    parser.add_argument("inicio", type=str, help="Data de desembolso (ex: 15/01/2022)")
    parser.add_argument("fim", type=str, help="Data de cálculo (ex: 10/09/2026)")
    
    args = parser.parse_args()
    simular_regras(args.valor, args.inicio, args.fim)