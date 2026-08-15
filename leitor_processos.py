import os
import sys
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import json
import pandas as pd
import google.generativeai as genai
from dotenv import load_dotenv

# Carrega as variáveis do arquivo .env
load_dotenv()

__version__ = "1.0.6"

# Configuração de diretórios
if getattr(sys, 'frozen', False):
    PASTA_APP = Path(sys.executable).parent
else:
    PASTA_APP = Path(__file__).parent

TEMPLATE_NASH = PASTA_APP / 'template_nash.xlsx'

def processar_com_gemini(caminho_pdf):
    """Processa o PDF usando a SDK clássica e estável do Google"""
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("A chave GEMINI_API_KEY não foi encontrada no arquivo .env.")
    
    genai.configure(api_key=api_key)
    
    # Upload do arquivo usando a File API nativa do generativeai
    print("Fazendo upload do PDF para os servidores do Google...")
    uploaded_file = genai.upload_file(path=caminho_pdf)
    
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

    model = genai.GenerativeModel('gemini-1.5-flash')
    response = model.generate_content(
        [uploaded_file, prompt],
        generation_config={
            "temperature": 0.0,
            "response_mime_type": "application/json"
        }
    )
    
    # Remove o arquivo da nuvem do Google
    genai.delete_file(uploaded_file.name)
    
    return json.loads(response.text)

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
        self.root.title(f"Nash System - Extrator v{__version__}")
        self.root.geometry("500x250")
        tk.Label(root, text="NASH SYSTEM - EXTRATOR IA", font=("Arial", 12, "bold")).pack(pady=20)
        self.lbl_status = tk.Label(root, text="Aguardando arquivo PDF...", font=("Arial", 10))
        self.lbl_status.pack(pady=10)
        self.btn = tk.Button(root, text="Selecionar PDF dos Autos", command=self.iniciar)
        self.btn.pack(pady=10)

    def iniciar(self):
        caminho = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not caminho: return
        self.btn.config(state="disabled")
        self.lbl_status.config(text="Processando no Google Cloud...", fg="blue")
        threading.Thread(target=self.rodar, args=(Path(caminho),)).start()

    def rodar(self, path):
        try:
            dados = processar_com_gemini(str(path))
            saida = path.parent / f"Planilha_{path.stem}.xlsx"
            preencher_template_nash(dados, saida)
            messagebox.showinfo("Sucesso", f"Gerado em: {saida.name}")
            self.lbl_status.config(text="Sucesso!", fg="green")
        except Exception as e:
            messagebox.showerror("Erro", str(e))
            self.lbl_status.config(text="Erro.", fg="red")
        self.btn.config(state="normal")

if __name__ == "__main__":
    app = tk.Tk()
    LeitorPDFGUI(app)
    app.mainloop()