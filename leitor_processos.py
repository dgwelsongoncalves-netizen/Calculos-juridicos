import os
import sys
import time
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import json
import pandas as pd
from google import genai
from google.genai import types
from dotenv import load_dotenv

__version__ = "2.1.0"  # Adição de multi-modelos (Padrão e Profundo)

# --- CONFIGURAÇÃO DE DIRETÓRIOS ---
if getattr(sys, 'frozen', False):
    PASTA_APP = Path(sys.executable).parent
else:
    PASTA_APP = Path(__file__).parent

TEMPLATE_NASH = PASTA_APP / 'template_nash.xlsx'

load_dotenv(dotenv_path=PASTA_APP / '.env')

TIMEOUT_HTTP_MS = 15 * 60 * 1000
TIMEOUT_PROCESSAMENTO_ARQUIVO_S = 600
MAX_TENTATIVAS = 3

def _aguardar_arquivo_ativo(client, uploaded_file):
    inicio = time.time()
    arquivo = uploaded_file
    while arquivo.state and arquivo.state.name == "PROCESSING":
        if time.time() - inicio > TIMEOUT_PROCESSAMENTO_ARQUIVO_S:
            raise TimeoutError("O Google levou mais de 10 minutos para processar o PDF. Tente novamente.")
        time.sleep(5)
        arquivo = client.files.get(name=arquivo.name)

    if not arquivo.state or arquivo.state.name != "ACTIVE":
        raise Exception(f"O Google rejeitou o arquivo (estado: {arquivo.state}).")
    return arquivo

def processar_com_gemini(caminho_pdf, modelo_escolhido="gemini-2.5-flash"):
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise ValueError(
            f"A chave GEMINI_API_KEY não foi encontrada.\n"
            f"Crie o arquivo '.env' em {PASTA_APP} com sua chave de faturamento ativo."
        )

    client = genai.Client(api_key=api_key, http_options=types.HttpOptions(timeout=TIMEOUT_HTTP_MS))

    prompt = """
    Você é um assistente jurídico sênior especializado em cálculos judiciais no TJMG.
    Analise os autos completos fornecidos.
    1. Localize a Petição Inicial e extraia danos materiais, notas fiscais e suas respectivas datas de desembolso.
    2. Localize ao longo de todos os autos as guias de custas, despesas processuais e suas datas.
    3. Identifique: Processo, Data da Sentença, Data do Trânsito, Data da Citação, Data do Evento.

    Retorne APENAS um objeto JSON válido (sem blocos de código markdown, apenas o JSON puro) com a estrutura exata:
    {
      "Parametros": {
        "Processo": "Número do processo",
        "Data_Sentenca": "DD/MM/AAAA",
        "Data_Transito": "DD/MM/AAAA",
        "Justiça_Gratuita": "NÃO",
        "Pagamento_Voluntario": "NÃO",
        "Honorarios_Sucumbencia": 10.0,
        "Honorarios_Fixos": 0.0,
        "Termo_Inicial_Juros": "DESEMBOLSO",
        "Data_Citacao": "DD/MM/AAAA",
        "Data_Evento": "DD/MM/AAAA",
        "Fazenda_Pública": "NÃO"
      },
      "Danos": [{"ID": "Folha X", "Descricao": "...", "Data": "DD/MM/AAAA", "Valor": 0.00, "Regra": "R1"}],
      "Custas": [{"ID": "Folha Y", "Descricao": "...", "Data": "DD/MM/AAAA", "Valor": 0.00}]
    }
    """

    ultimo_erro = None
    for tentativa in range(1, MAX_TENTATIVAS + 1):
        uploaded_file = None
        try:
            print(f"[Tentativa {tentativa}/{MAX_TENTATIVAS}] Enviando PDF... (Modelo: {modelo_escolhido})")
            uploaded_file = client.files.upload(file=caminho_pdf)
            uploaded_file = _aguardar_arquivo_ativo(client, uploaded_file)

            response = client.models.generate_content(
                model=modelo_escolhido,
                contents=[uploaded_file, prompt],
                config=types.GenerateContentConfig(
                    temperature=0.0,
                    response_mime_type="application/json",
                ),
            )

            client.files.delete(name=uploaded_file.name)
            return json.loads(response.text)

        except Exception as e:
            ultimo_erro = e
            print(f"Falhou na tentativa {tentativa}: {e}")
            if uploaded_file is not None:
                try:
                    client.files.delete(name=uploaded_file.name)
                except Exception:
                    pass
            if tentativa < MAX_TENTATIVAS:
                time.sleep(5 * tentativa)

    raise Exception(f"Falha ao processar o PDF após {MAX_TENTATIVAS} tentativas.\nÚltimo erro: {ultimo_erro}")

def preencher_template_nash(dados_json, caminho_saida):
    if not TEMPLATE_NASH.exists():
        raise FileNotFoundError(f"Template não encontrado: {TEMPLATE_NASH}")

    xls = pd.ExcelFile(TEMPLATE_NASH)
    df_param = pd.read_excel(xls, sheet_name='Parametros', header=None)
    p = dados_json.get('Parametros', {})

    mapeamento = {
        'Processo': p.get('Processo'), 'Data da Sentença': p.get('Data_Sentenca'),
        'Data do Trânsito': p.get('Data_Transito'), 'Justiça Gratuita': p.get('Justiça_Gratuita'),
        'Pagamento Voluntário 15d': p.get('Pagamento_Voluntario'), 'Honorários Sucumbência': p.get('Honorarios_Sucumbencia'),
        'Honorários Fixos (R$)': p.get('Honorarios_Fixos'), 'Termo Inicial Juros': p.get('Termo_Inicial_Juros'),
        'Data da Citação': p.get('Data_Citacao'), 'Data do Evento': p.get('Data_Evento'), 'Fazenda Pública': p.get('Fazenda_Pública')
    }

    for idx, row in df_param.iterrows():
        if row[0] in mapeamento and mapeamento[row[0]] is not None:
            df_param.loc[idx, 1] = mapeamento[row[0]]

    df_danos = pd.DataFrame(dados_json.get('Danos', [])).rename(columns={'ID': 'ID / Folha', 'Descricao': 'Descrição', 'Data': 'Data Desembolso', 'Valor': 'Valor Histórico'})
    df_custas = pd.DataFrame(dados_json.get('Custas', [])).rename(columns={'ID': 'ID / Folha', 'Descricao': 'Descrição', 'Data': 'Data Desembolso', 'Valor': 'Valor Histórico'})

    with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
        df_param.to_excel(writer, sheet_name='Parametros', index=False, header=False)
        if not df_danos.empty: df_danos.to_excel(writer, sheet_name='Danos', index=False)
        if not df_custas.empty: df_custas.to_excel(writer, sheet_name='Custas', index=False)
        if 'Deducoes' in xls.sheet_names: pd.read_excel(xls, sheet_name='Deducoes').to_excel(writer, sheet_name='Deducoes', index=False)

class LeitorPDFGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Nash System - Extrator IA (v{__version__})")
        self.root.geometry("550x320")
        
        tk.Label(root, text="NASH SYSTEM - EXTRATOR DE AUTOS", font=("Arial", 14, "bold")).pack(pady=15)
        
        self.lbl_status = tk.Label(root, text="Aguardando arquivo PDF...", font=("Arial", 10))
        self.lbl_status.pack(pady=5)
        
        frame_botoes = tk.Frame(root)
        frame_botoes.pack(pady=15)
        
        self.btn_padrao = tk.Button(
            frame_botoes, text="📄 Extração Padrão (Rápida/Barata)", 
            font=("Arial", 10, "bold"), bg="#107C41", fg="white", 
            command=lambda: self.iniciar("gemini-2.5-flash")
        )
        self.btn_padrao.pack(fill="x", pady=5)
        
        self.btn_profundo = tk.Button(
            frame_botoes, text="🔍 Extração Profunda (Modelo Pro - Autos Complexos)", 
            font=("Arial", 10, "bold"), bg="#B22222", fg="white", 
            command=lambda: self.iniciar("gemini-2.5-pro")
        )
        self.btn_profundo.pack(fill="x", pady=5)

    def iniciar(self, modelo_escolhido):
        caminho = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not caminho: return
        
        self.btn_padrao.config(state="disabled")
        self.btn_profundo.config(state="disabled")
        self.lbl_status.config(text=f"Processando no Google Cloud ({modelo_escolhido})...", fg="blue")
        
        threading.Thread(target=self.rodar, args=(Path(caminho), modelo_escolhido)).start()

    def rodar(self, path, modelo_escolhido):
        try:
            dados = processar_com_gemini(str(path), modelo_escolhido)
            saida = path.parent / f"Planilha_{path.stem}.xlsx"
            preencher_template_nash(dados, saida)
            messagebox.showinfo("Sucesso", f"Planilha gerada com sucesso em:\n{saida.name}")
            self.lbl_status.config(text="Sucesso!", fg="green")
        except Exception as e:
            messagebox.showerror("Erro", str(e))
            self.lbl_status.config(text="Erro.", fg="red")
        
        self.btn_padrao.config(state="normal")
        self.btn_profundo.config(state="normal")

if __name__ == "__main__":
    app = tk.Tk()
    LeitorPDFGUI(app)
    app.mainloop()