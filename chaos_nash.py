import random
import shutil
import tempfile
from datetime import datetime
from pathlib import Path
from unittest.mock import patch
import pandas as pd
import openpyxl
import nash

# Garante o carregamento das bibliotecas pesadas do motor
nash.preload_heavy_libs()


class ChaosMonkeyNash:
    """
    Agente de Engenharia de Caos e Fuzzing local para o Nash System.
    Injeta dados malformados, cenários extremos e falhas de rede para testar
    a resiliência contábil e estrutural do motor matemático.
    """

    def __init__(self, iteracoes=25, seed=None, salvar_todas_planilhas=False):
        self.iteracoes = iteracoes
        self.seed = seed if seed is not None else random.randint(1, 999999)
        self.salvar_todas_planilhas = salvar_todas_planilhas
        random.seed(self.seed)
        self.regras = ["R1", "R2", "R3", "R4", "R5", "R6", "R7"]
        self.relatorio_falhas = []
        self.log_execucao = []

        # Pastas de saída visíveis no projeto
        self.pasta_auditoria = Path("Relatorios_Auditoria")
        self.pasta_forense = Path("Relatorios_Chaos_Falhas")

    def _caos_moeda(self, valor_float):
        """Transforma um float limpo em formatos humanos caóticos de digitação."""
        formatos = [
            lambda v: v,  # Float nativo
            lambda v: f"{v:.2f}".replace(".", ","),  # Vírgula simples: "2792,23"
            lambda v: f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),  # Padrão BR: "2.792,23"
            lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),  # Com R$: "R$ 2.792,23"
            lambda v: f"  {v:.2f}".replace(".", ",") + "  ",  # Espaços nas bordas
            lambda v: str(int(v)) if v.is_integer() else f"{v:.2f}",  # String com ponto
        ]
        return random.choice(formatos)(valor_float)

    def _caos_percentual(self, perc_int):
        """Gera variações de preenchimento de percentuais."""
        opcoes = [
            str(perc_int),
            f"{perc_int}%",
            f" {perc_int},0% ",
            f"{perc_int / 100.0:.2f}".replace(".", ","),
        ]
        return random.choice(opcoes)

    def _caos_data(self, ano_min=2021, ano_max=2025):
        """Gera datas válidas em formatos variados de string ou datetime."""
        ano = random.randint(ano_min, ano_max)
        mes = random.randint(1, 12)
        dia = random.randint(1, 28)
        formatos = [
            f"{dia:02d}/{mes:02d}/{ano}",
            f" {dia:02d}/{mes:02d}/{ano} ",
            pd.Timestamp(year=ano, month=mes, day=dia),
        ]
        return random.choice(formatos)

    def gerar_cenario_caotico(self, caminho_xlsx, idx_cenario):
        """Cria um template Excel com combinações extremas de parâmetros e deduções."""
        pagamento_15d = random.choice(["Não se aplica", "Sim", "Não", "Parcial", " nao ", "PARCIAL"])
        tem_valor_parcial_param = random.choice([True, False])
        val_parcial = self._caos_moeda(random.uniform(100.0, 5000.0)) if tem_valor_parcial_param else ""
        atuacao = random.choice(["RÉU", "AUTOR", " réu ", "autor"])
        fazenda = random.choice(["Não", "Sim"])
        jg = random.choice(["Não", "Sim", "NÃO", "sim"])

        parametros = {
            "Processo": [f"CHAOS-TEST-{idx_cenario:03d}"],
            "Atuação": [atuacao],
            "Data do Trânsito": [self._caos_data(2024, 2026) if random.choice([True, False]) else ""],
            "Data da Sentença": [self._caos_data(2023, 2025)],
            "Termo Inicial Juros": [random.choice(["Desembolso", "Citação", "Evento", "Propositura", "Trânsito em julgado"])],
            "Data da Citação": [self._caos_data(2021, 2023)],
            "Data do Evento": [self._caos_data(2020, 2022)],
            "Justiça Gratuita": [jg],
            "Juros sobre Custas": [random.choice(["Sim (Após Trânsito)", "Não (Sem Juros)"])],
            "Pagamento Voluntário 15d": [pagamento_15d],
            "Valor Pagamento Parcial (R$)": [val_parcial],
            "Fazenda Pública": [fazenda],
            "Honorários Sucumbência (%)": [self._caos_percentual(random.choice([10, 15, 20]))],
            "Honorários Fixos (R$)": [self._caos_moeda(1500.0) if random.random() < 0.2 else ""],
            "Base Honorários": [random.choice(["CONDENAÇÃO", "VALOR DA CAUSA", "PROVEITO ECONÔMICO"])],
            "Valor Causa Original": [self._caos_moeda(random.uniform(5000.0, 50000.0))],
            "Data Propositura": [self._caos_data(2021, 2022)],
            "Proporção Honorários (%)": [self._caos_percentual(random.choice([50, 75, 80, 100]))],
            "Proporção Custas (%)": [self._caos_percentual(random.choice([50, 75, 80, 100]))],
        }
        df_param = pd.DataFrame(parametros).T

        qtd_danos = random.randint(1, 5)
        regra_escolhida = random.choice(self.regras)
        linhas_danos = []
        for i in range(qtd_danos):
            v_hist = round(random.uniform(50.0, 20000.0), 2)
            linhas_danos.append({
                "ID / Folha": f"Doc Caos {i+1}",
                "Descrição": f"Dano Randômico {i+1} ({regra_escolhida})",
                "Data Desembolso": self._caos_data(2021, 2025),
                "Valor Histórico": self._caos_moeda(v_hist),
                "Regra": regra_escolhida if random.random() > 0.2 else random.choice(self.regras),
                "Data Juros": self._caos_data(2021, 2025) if random.random() < 0.3 else "",
                "Valor Pedido Inicial": self._caos_moeda(v_hist * 1.5),
                "Data do Pedido": self._caos_data(2021, 2023),
            })
        df_danos = pd.DataFrame(linhas_danos)

        qtd_custas = random.randint(0, 3)
        linhas_custas = []
        for i in range(qtd_custas):
            linhas_custas.append({
                "ID / Folha": f"Guia {i+1}",
                "Descrição": f"Custa Caótica {i+1}",
                "Data Desembolso": self._caos_data(2021, 2025),
                "Valor Histórico": self._caos_moeda(round(random.uniform(35.0, 1500.0), 2)),
            })
        df_custas = pd.DataFrame(linhas_custas) if linhas_custas else pd.DataFrame(
            columns=["ID / Folha", "Descrição", "Data Desembolso", "Valor Histórico"]
        )

        qtd_deducoes = random.choice([0, 0, 1, 2, 3])
        linhas_ded = []
        for i in range(qtd_deducoes):
            linhas_ded.append({
                "ID / Folha": f"Depósito {i+1}",
                "Data bloqueio/deposito": self._caos_data(2023, 2025),
                "Valor": self._caos_moeda(round(random.uniform(200.0, 8000.0), 2)),
            })
        df_deducoes = pd.DataFrame(linhas_ded) if linhas_ded else pd.DataFrame(
            columns=["ID / Folha", "Data bloqueio/deposito", "Valor"]
        )

        with pd.ExcelWriter(caminho_xlsx, engine="openpyxl") as writer:
            df_param.to_excel(writer, sheet_name="Parametros", header=False)
            df_danos.to_excel(writer, sheet_name="Danos", index=False)
            df_custas.to_excel(writer, sheet_name="Custas", index=False)
            df_deducoes.to_excel(writer, sheet_name="Deducoes", index=False)

        resumo_cenario = (
            f"Regra={regra_escolhida} | Danos={qtd_danos} | Custas={qtd_custas} | "
            f"Deduções={qtd_deducoes} | 523={pagamento_15d.strip()} | Faz={fazenda} | JG={jg.strip()}"
        )
        return resumo_cenario

    def auditar_invariantes_laudo(self, caminho_laudo, idx_cenario):
        """Verifica se o Excel gerado não possui NaNs visuais, valores negativos ou quebras."""
        wb = openpyxl.load_workbook(caminho_laudo, data_only=True)
        ws = wb["Laudo de Liquidação"]

        val_final_encontrado = None
        for r_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
            texto_linha = " ".join([str(c) for c in row if c is not None])
            if "nan" in texto_linha.lower() and "id:" not in texto_linha.lower():
                raise AssertionError(f"Célula com texto 'NaN' vazou na linha {r_idx} do Laudo: {row}")

            col0 = str(row[0]) if row[0] else ""
            if "TOTAL GERAL DEVIDO:" in col0 or "SALDO DEVEDOR FINAL DA EXECUÇÃO:" in col0:
                val_final_encontrado = float(row[7])
                if val_final_encontrado < -0.01:
                    raise AssertionError(f"Saldo final negativo detectado ({val_final_encontrado}) no cenário {idx_cenario}!")

        if val_final_encontrado is None:
            raise AssertionError(f"Linha do total final não encontrada no Laudo do cenário {idx_cenario}!")
        return val_final_encontrado

    def executar_bateria_caos(self):
        self.pasta_auditoria.mkdir(exist_ok=True)
        arquivo_relatorio_txt = self.pasta_auditoria / "relatorio_chaos_monkey.txt"

        cabecalho = (
            f"=====================================================================\n"
            f"🐒 RELATÓRIO DE ENGENHARIA DE CAOS - NASH SYSTEM v{nash.__version__}\n"
            f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} | Seed: {self.seed} | Ataques: {self.iteracoes}\n"
            f"=====================================================================\n"
        )
        print(cabecalho)
        self.log_execucao.append(cabecalho)

        if not nash.ARQUIVO_CACHE_BCB.exists():
            print("📡 Inicializando cache local do Banco Central antes do estresse...")
            nash.carregar_taxas_bcb(pd.to_datetime("2020-01-01"))

        original_pdf = nash.converter_para_pdf
        nash.converter_para_pdf = lambda x: True

        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                pasta_temp = Path(tmpdir)

                with patch("bcb.sgs.get", side_effect=ConnectionError("Chaos Monkey cortou o cabo do BCB!")):
                    for i in range(1, self.iteracoes + 1):
                        arq_in = pasta_temp / f"chaos_in_{i:02d}.xlsx"
                        arq_out = pasta_temp / f"chaos_out_{i:02d}.xlsx"
                        resumo_cenario = self.gerar_cenario_caotico(arq_in, i)

                        try:
                            nash.executar_nash(str(arq_in), str(arq_out))
                            total_apurado = self.auditar_invariantes_laudo(arq_out, i)
                            msg_ok = f"  [SOBREVIVEU] Ataque #{i:02d} ({resumo_cenario}) -> Total: R$ {total_apurado:,.2f}"
                            print(msg_ok)
                            self.log_execucao.append(msg_ok)

                            # Salva o último ataque como amostra na pasta Relatorios_Auditoria (ou todos se salvar_todas_planilhas=True)
                            if self.salvar_todas_planilhas:
                                pasta_amostras = self.pasta_auditoria / "Chaos_Amostras"
                                pasta_amostras.mkdir(exist_ok=True)
                                shutil.copy(arq_in, pasta_amostras / f"chaos_template_{i:02d}.xlsx")
                                shutil.copy(arq_out, pasta_amostras / f"chaos_laudo_{i:02d}.xlsx")
                            elif i == self.iteracoes:
                                shutil.copy(arq_in, self.pasta_auditoria / "amostra_chaos_template.xlsx")
                                shutil.copy(arq_out, self.pasta_auditoria / "amostra_chaos_laudo.xlsx")

                        except Exception as e:
                            msg_erro = f"  [💥 FALHA DETECTADA] Ataque #{i:02d} ({resumo_cenario}) quebrou: {e}"
                            print(msg_erro)
                            self.log_execucao.append(msg_erro)
                            self.pasta_forense.mkdir(exist_ok=True)
                            caminho_evidencia = self.pasta_forense / f"falha_cenario_{i:02d}_seed_{self.seed}.xlsx"
                            shutil.copy(arq_in, caminho_evidencia)
                            self.relatorio_falhas.append((i, str(e), caminho_evidencia))

        finally:
            nash.converter_para_pdf = original_pdf

        rodape = ["\n" + "=" * 69]
        if not self.relatorio_falhas:
            rodape.append("✅ RESULTADO: O Nash System sobreviveu a 100% dos ataques do Chaos Monkey!")
            rodape.append(f"📄 Relatório detalhado salvo em: {arquivo_relatorio_txt}")
            rodape.append(f"📊 Planilha de amostra do último ataque salva em: {self.pasta_auditoria / 'amostra_chaos_laudo.xlsx'}")
        else:
            rodape.append(f"⚠️ RESULTADO: O Chaos Monkey encontrou {len(self.relatorio_falhas)} vulnerabilidade(s):")
            for idx, erro, evid in self.relatorio_falhas:
                rodape.append(f"   - Cenário #{idx:02d}: {erro}")
                rodape.append(f"     📁 Planilha causadora salva em: {evid}")
        rodape.append("=" * 69)

        texto_rodape = "\n".join(rodape)
        print(texto_rodape)
        self.log_execucao.append(texto_rodape)

        with open(arquivo_relatorio_txt, "w", encoding="utf-8") as f:
            f.write("\n".join(self.log_execucao))


if __name__ == "__main__":
    monkey = ChaosMonkeyNash(iteracoes=25, salvar_todas_planilhas=False)
    monkey.executar_bateria_caos()