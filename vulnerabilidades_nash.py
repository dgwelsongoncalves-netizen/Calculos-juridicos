"""Vulnerabilidades conhecidas do nash.py (entrada malformada e ordem das linhas). Rode ao lado do nash.py: python vulnerabilidades_nash.py"""
import io, contextlib, tempfile, openpyxl, pandas as pd
from pathlib import Path
from unittest.mock import patch
import nash
nash.preload_heavy_libs(); nash.converter_para_pdf = lambda x: True
def rodar(danos, deducoes, regra="R6"):
    p = {"Processo":"MAL-1","Atuação":"AUTOR","Data do Trânsito":"","Data da Sentença":"01/03/2025","Termo Inicial Juros":"Desembolso",
         "Justiça Gratuita":"Não","Pagamento Voluntário 15d":"Sim","Fazenda Pública":"Não","Honorários Sucumbência (%)":10}
    with tempfile.TemporaryDirectory() as td:
        e = Path(td)/"e.xlsx"; wb = openpyxl.Workbook(); ws = wb.active; ws.title="Parametros"
        for k,v in p.items(): ws.append([k,v])
        w = wb.create_sheet("Danos"); w.append(["ID / Folha","Descrição","Data Desembolso","Valor Histórico","Regra"])
        for r in danos: w.append([r[0],r[1],r[2],r[3],regra])
        w = wb.create_sheet("Custas"); w.append(["ID / Folha","Descrição","Data Desembolso","Valor Histórico"])
        w = wb.create_sheet("Deducoes"); w.append(["ID / Folha","Data bloqueio/deposito","Valor"])
        for r in deducoes: w.append(r)
        wb.save(e)
        try:
            with patch("bcb.sgs.get", side_effect=ConnectionError("x")), contextlib.redirect_stdout(io.StringIO()):
                nash.executar_nash(str(e), str(Path(td)/"s.xlsx"))
        except Exception as ex: return f"ERRO EXPLÍCITO: {str(ex)[:70]}"
        for f in Path(td).iterdir():
            if f.name.startswith("s"):
                wb2 = openpyxl.load_workbook(f, data_only=True)
                if "Laudo de Liquidação" in wb2.sheetnames:
                    for r in wb2["Laudo de Liquidação"].iter_rows(values_only=True):
                        t = " ".join(str(c) for c in r if c is not None)
                        if "TOTAL GERAL DEVIDO:" in t or "SALDO DEVEDOR FINAL" in t:
                            return round([c for c in r if isinstance(c,(int,float))][-1], 2)
    return "sem total"
ok = [["F1","Nota fiscal","10/10/2024",5000.0]]
print("A) referência, sem bloqueio                 :", rodar(ok, []))
print("B) referência, bloqueio válido de R$1.000   :", rodar(ok, [["D1","10/01/2026",1000.0]]))
print("C) bloqueio com data inválida '31/02/2026'  :", rodar(ok, [["D1","31/02/2026",1000.0]]))
print("D) bloqueio com data em texto 'ontem'       :", rodar(ok, [["D1","ontem",1000.0]]))
print("E) dano com data inválida '31/02/2024'      :", rodar(ok+[["F2","Outra nota","31/02/2024",3000.0]], []))
print("F) dano com valor ilegível 'abc'            :", rodar(ok+[["F2","Outra nota","10/11/2024","abc"]], []))
print("G) dano com data como número do Excel 44966 :", rodar([["F1","Nota fiscal",44966,5000.0]], []))
print("H) igual a G, mas com a data correta        :", rodar([["F1","Nota fiscal","05/02/2023",5000.0]], []))

print()
import io, contextlib, tempfile, openpyxl
from pathlib import Path
from unittest.mock import patch
import nash
nash.preload_heavy_libs(); nash.converter_para_pdf = lambda x: True
def total(danos):
    p = {"Processo":"ORD","Atuação":"AUTOR","Data da Sentença":"01/03/2025","Termo Inicial Juros":"Desembolso","Justiça Gratuita":"Não","Pagamento Voluntário 15d":"Sim","Fazenda Pública":"Não","Honorários Sucumbência (%)":10}
    with tempfile.TemporaryDirectory() as td:
        e = Path(td)/"e.xlsx"; wb = openpyxl.Workbook(); ws = wb.active; ws.title="Parametros"
        for k,v in p.items(): ws.append([k,v])
        w = wb.create_sheet("Danos"); w.append(["ID / Folha","Descrição","Data Desembolso","Valor Histórico","Regra"])
        for r in danos: w.append(r)
        w = wb.create_sheet("Custas"); w.append(["ID / Folha","Descrição","Data Desembolso","Valor Histórico"])
        w = wb.create_sheet("Deducoes"); w.append(["ID / Folha","Data bloqueio/deposito","Valor"]); w.append(["D1","10/01/2026",1000.0])
        wb.save(e)
        with patch("bcb.sgs.get", side_effect=ConnectionError("x")), contextlib.redirect_stdout(io.StringIO()):
            nash.executar_nash(str(e), str(Path(td)/"s.xlsx"))
        for f in Path(td).iterdir():
            if f.name.startswith("s"):
                for r in openpyxl.load_workbook(f, data_only=True)["Laudo de Liquidação"].iter_rows(values_only=True):
                    t = " ".join(str(c) for c in r if c is not None)
                    if "SALDO DEVEDOR FINAL" in t: return round([c for c in r if isinstance(c,(int,float))][-1],2)
A = ["F1","Nota antiga","10/10/2021",5000.0,"R1"]; B = ["F2","Nota recente","10/10/2024",5000.0,"R6"]
print("Mesmas duas linhas e mesmo bloqueio, só trocando a ORDEM das linhas na aba Danos:")
print("  [R1, R6] ->", total([A,B]))
print("  [R6, R1] ->", total([B,A]))
