"""
Teste de mutação do Nash System.

Para cada "mutante" (uma cópia do nash.py com UM bug injetado de propósito), roda o oráculo
(oraculo_nash.py). Se o oráculo acusa falha, o mutante foi "pego" (bom). Se passa, o detector
tem um ponto cego (ruim). Rode com os arquivos na mesma pasta do nash.py:  python mutation_nash.py
Também funciona com o Chaos Monkey:  python mutation_nash.py --monkey
"""
import json, os, shutil, subprocess, sys, tempfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
# Cada mutante = (id, descrição, trecho_original, trecho_novo). O trecho original TEM que existir exatamente 1x.
M = [
 ("M01","TJMG invertido (idx_fim/idx_ini)",
  "return idx_ini / idx_fim if idx_fim != 0 else 1.0","return idx_fim / idx_ini if idx_ini != 0 else 1.0"),
 ("M02","Multa art. 523 de 10% para 1%",
  "valor_multa = (base_multa * 0.10) if houve_inadimplemento else 0.0","valor_multa = (base_multa * 0.01) if houve_inadimplemento else 0.0"),
 ("M03","Saldo de juros deixa de ser corrigido (bug antigo)",
  "s[1] = s[1] * f_cm_delta + s[0] * f_jur_delta","s[1] = s[1] + s[0] * f_jur_delta"),
 ("M04","Bloqueio abate principal antes dos juros (viola art. 354 CC)",
  "            abate_juros = min(valor_ded, J0)\n            if J0 > 0:\n                for s in saldos.values(): s[1] -= abate_juros * (s[1] / J0)\n            abate_princ = valor_ded - abate_juros\n",
  "            abate_princ = min(valor_ded, P0)\n            abate_juros = min(valor_ded - abate_princ, J0)\n            if J0 > 0:\n                for s in saldos.values(): s[1] -= abate_juros * (s[1] / J0)\n"),
 ("M05","Taxa legal (juros) 50% maior na Lei 14.905",
  "juros = df_periodo['TAXA_LEGAL_RETROATIVA'].sum()","juros = df_periodo['TAXA_LEGAL_RETROATIVA'].sum() * 1.5"),
 ("M06","Justiça gratuita ignorada nas custas (caminho sem bloqueio)",
  "            if jg_adversa: continue\n            val_princ = (row['Valor Histórico'] * row['Fator CM']) * prop_custas","            val_princ = (row['Valor Histórico'] * row['Fator CM']) * prop_custas"),
 ("M07","Fazenda Pública passa a pagar multa do art. 523",
  "    if is_fazenda:\n        houve_inadimplemento_total = False","    if False:\n        houve_inadimplemento_total = False"),
 ("M08","Pagamento parcial não é abatido da base da multa",
  "base_multa = max(0.0, base_multa_bruta - valor_pago_parcial)","base_multa = base_multa_bruta"),
 ("M09","Última linha de Danos some silenciosamente",
  "    if 'Valor Pedido Inicial' not in df_danos.columns: df_danos['Valor Pedido Inicial'] = 0.0\n","    if len(df_danos) > 1: df_danos = df_danos.iloc[:-1]\n    if 'Valor Pedido Inicial' not in df_danos.columns: df_danos['Valor Pedido Inicial'] = 0.0\n"),
 ("M10","Honorários do art. 523 de 10% para 20%",
  "hon_523 = (base_multa * 0.10) if (houve_inadimplemento and not jg_adversa) else 0.0","hon_523 = (base_multa * 0.20) if (houve_inadimplemento and not jg_adversa) else 0.0"),
 ("M11","IPCA acumulado por soma em vez de produto (Lei 14.905)",
  "f_ipca = (1 + df_bcb['IPCA'].loc[(df_bcb['IPCA'].index >= d_cm_m) & (df_bcb['IPCA'].index <= d_calc_m), 'IPCA']).prod()",
  "f_ipca = 1 + df_bcb['IPCA'].loc[(df_bcb['IPCA'].index >= d_cm_m) & (df_bcb['IPCA'].index <= d_calc_m), 'IPCA'].sum()"),
 ("M12","Juros de 1% a.m. (R1) viram 2% a.m.",
  "(data_calculo.month - data_juros.month)) * 0.01 if pd.notna(data_juros) and data_juros <= data_calculo else 0.0","(data_calculo.month - data_juros.month)) * 0.02 if pd.notna(data_juros) and data_juros <= data_calculo else 0.0"),
 ("M13","Poupança (Fazenda, pré-EC113) pela metade",
  ", 'POUPANCA'].sum()",", 'POUPANCA'].sum() * 0.5"),
 ("M14","Valor do bloqueio/depósito aplicado pela metade",
  "valor_ded = float(row['Valor'])","valor_ded = float(row['Valor']) * 0.5"),
 ("M15","Bloqueio mais antigo some em silêncio",
  "        df_deducoes = df_ded.sort_values('Data bloqueio/deposito', kind='stable')\n","        df_deducoes = df_ded.sort_values('Data bloqueio/deposito', kind='stable')\n        if len(df_deducoes) > 1: df_deducoes = df_deducoes.iloc[1:]\n"),
 ("M16","Regra forçada: todo saldo segue a regra da 1ª linha (bug antigo)",
  "def grupo_de(regra): return 'FAZ' if is_fazenda else regra","def grupo_de(regra): return 'FAZ' if is_fazenda else resolver_regra(df_danos.iloc[0].get('Regra', ''))"),
 ("M17","Data inválida deixa de gerar erro (vira vazio)",
  "if pd.isna(dt) or dt.year < 1990 or dt.year > 2100:","if False:"),
 ("M18","Valor ilegível ('abc') vira zero em silêncio",
  "    except ValueError: raise EntradaInvalidaError(f\"{onde}: não foi possível ler '{val}' como valor em reais.\")","    except ValueError: return 0.0"),
]

AQUI = Path(__file__).resolve().parent
USAR_MONKEY = "--monkey" in sys.argv

def preparar(mid, novo_src):
    d = Path(tempfile.mkdtemp(prefix=f"nash_{mid}_"))
    (d/"nash.py").write_text(novo_src, encoding="utf-8")
    shutil.copy(AQUI/"oraculo_nash.py", d/"oraculo_nash.py")
    if (AQUI/"chaos_nash.py").exists(): shutil.copy(AQUI/"chaos_nash.py", d/"chaos_nash.py")
    os.symlink((AQUI/"Tabelas_Oficiais").resolve(), d/"Tabelas_Oficiais")
    return d

def rodar(item):
    mid, desc, old, new = item
    src = (AQUI/"nash.py").read_text(encoding="utf-8")
    if src.count(old) != 1:
        return mid, desc, "N/A", f"trecho original aparece {src.count(old)}x (nash.py mudou; ajuste o mutante)"
    d = preparar(mid, src.replace(old, new))
    try:
        r = subprocess.run([sys.executable, "oraculo_nash.py"], cwd=d, capture_output=True, text=True, timeout=240)
        j = json.loads(r.stdout.strip().splitlines()[-1])
        status = "PEGOU" if j["n_falhas"] else "passou"
        detalhe = ",".join(j["tags"])
    except Exception as e:
        status, detalhe = "ERRO", str(e)[:60]
    if USAR_MONKEY and (d/"chaos_nash.py").exists():
        (d/"run_monkey.py").write_text("import io,contextlib,json,chaos_nash as c\nm=c.ChaosMonkeyNash(iteracoes=25,seed=101)\nwith contextlib.redirect_stdout(io.StringIO()): m.executar_bateria_caos()\nprint(len([f for f in m.relatorio_falhas if 'does not exist' not in f[1]]))\n")
        r2 = subprocess.run([sys.executable, "run_monkey.py"], cwd=d, capture_output=True, text=True, timeout=240)
        detalhe += f" | monkey: {'PEGOU' if r2.stdout.strip().endswith(tuple('123456789')) else 'passou'}"
    shutil.rmtree(d, ignore_errors=True)
    return mid, desc, status, detalhe

if __name__ == "__main__":
    with ThreadPoolExecutor(max_workers=4) as ex: res = list(ex.map(rodar, M))
    pegos = 0
    for mid, desc, st, det in res:
        print(f"{mid} {st:7} {desc[:62]:62} [{det}]"); pegos += st == "PEGOU"
    print(f"\nPegos: {pegos}/{len(M)}  (os que 'passaram' são pontos cegos do detector)")