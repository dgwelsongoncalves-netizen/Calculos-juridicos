import pandas as pd
from pathlib import Path
import urllib.request
import json
import re

GITHUB_REPO = "dgwelsongoncalves-netizen/Calculos-juridicos"

# Configurando os caminhos das pastas
PASTA_BASE = Path(__file__).parent
PASTA_TABELAS = PASTA_BASE / 'Tabelas_Oficiais'
ARQUIVO_TJMG = PASTA_TABELAS / 'tabela_tjmg.xlsx'


def extrair_tupla_versao(v_str):
    nums = [int(x) for x in re.findall(r"\d+", str(v_str))][:3]
    return tuple(nums) if nums else (0, 0, 0)


def consultar_versao_remota(repo=GITHUB_REPO):
    """
    Lê diretamente a variável __version__ do nash.py na branch main do GitHub
    e busca o instalador .exe na última Release disponível.
    """
    tag_remota = ""
    url_exe = None
    url_html = f"https://github.com/{repo}/releases"

    # 1. Lê __version__ direto do código-fonte no GitHub (sem depender de tag manual)
    url_raw = f"https://raw.githubusercontent.com/{repo}/main/nash.py"
    req_raw = urllib.request.Request(url_raw, headers={"User-Agent": "NashSystem-Updater"})
    with urllib.request.urlopen(req_raw, timeout=10) as resp_raw:
        conteudo_nash = resp_raw.read().decode("utf-8", errors="ignore")
        match_ver = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', conteudo_nash)
        if match_ver:
            tag_remota = match_ver.group(1)

    # 2. Busca o instalador .exe nos assets da última Release (se existir)
    try:
        url_api = f"https://api.github.com/repos/{repo}/releases/latest"
        req_api = urllib.request.Request(url_api, headers={"User-Agent": "NashSystem-Updater"})
        with urllib.request.urlopen(req_api, timeout=10) as resp_api:
            dados = json.loads(resp_api.read().decode("utf-8"))
            if not tag_remota:
                tag_remota = dados.get("tag_name", "")
            url_html = dados.get("html_url", url_html)
            for asset in dados.get("assets", []):
                if str(asset.get("name", "")).lower().endswith(".exe"):
                    url_exe = asset.get("browser_download_url")
                    break
    except Exception:
        pass

    return tag_remota, url_exe, url_html


def carregar_tabela_tjmg():
    print("Iniciando a leitura da tabela oficial do TJMG...")
    try:
        df_tjmg = pd.read_excel(ARQUIVO_TJMG, sheet_name='Plan1', skiprows=8, names=['ANO', 'MÊS', 'ÍNDICE'])
        df_tjmg = df_tjmg.dropna(subset=['MÊS', 'ÍNDICE'])

        df_tjmg['MÊS'] = df_tjmg['MÊS'].astype(str).str.strip()
        meses = {
            'Janeiro': '01', 'Fevereiro': '02', 'Março': '03', 'Abril': '04',
            'Maio': '05', 'Junho': '06', 'Julho': '07', 'Agosto': '08',
            'Setembro': '09', 'Outubro': '10', 'Novembro': '11', 'Dezembro': '12'
        }
        df_tjmg['MÊS_NUM'] = df_tjmg['MÊS'].map(meses)
        df_tjmg = df_tjmg.dropna(subset=['MÊS_NUM'])

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

    indice_nota = df_tjmg.loc[data_nota, 'ÍNDICE']
    indice_limite = df_tjmg.loc[data_limite, 'ÍNDICE']

    return indice_nota / indice_limite


if __name__ == "__main__":
    tabela = carregar_tabela_tjmg()

    if tabela is not None:
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