Markdown

# 🏛️ Nash System (v4.1)
**Assistente de Cálculos Judiciais, Liquidação de Sentença e Análise Estratégica**

O **Nash System** é uma ferramenta open-source desenvolvida em Python para automatizar o processamento de cálculos judiciais complexos, elaboração de laudos de liquidação de sentença e relatórios de proveito econômico (êxito). O sistema integra uma interface gráfica intuitiva para entrada de dados com um robusto motor matemático (desacoplado em background) que aplica rigorosamente as tabelas oficiais e a legislação processual civil vigente. 

🛡️ **Privacy by Design & LGPD:** Toda a manipulação de dados de processos e valores é executada e armazenada **localmente** (offline) na máquina do usuário. Nenhuma informação sensível é enviada para nuvens ou servidores externos, garantindo conformidade absoluta com a Lei Geral de Proteção de Dados (LGPD).

---

## 🚀 Principais Funcionalidades (Atualização v4.1)

- **Adequação ao Art. 523, § 2º CPC (Novo):** Cálculo inteligente de pagamento voluntário parcial. O sistema deduz matematicamente o depósito feito no prazo de 15 dias e faz incidir a multa (10%) e honorários (10%) **exclusivamente** sobre o saldo devedor remanescente.
- **Dashboard Estratégico (Viabilidade Financeira):** Renderização gráfica em tempo real alertando sobre o risco econômico da ação. Cruza o volume de custas adiantadas (*Sunk Cost*) com o pedido principal e o status de Justiça Gratuita da parte adversa para apontar inviabilidade econômica.
- **Módulo de Quebra de Acordos:** Aba dedicada para calcular inadimplementos de acordos homologados, permitindo imputar parcelas pagas, multas contratuais moratórias, honorários de retomada e a reversão de descontos confessados.
- **Conta Gráfica Dinâmica (Amortização Art. 354 do CC):** Abatimento automático de múltiplos depósitos judiciais, bloqueios (Sisbajud) ou expedições de alvarás diretamente na linha do tempo do cálculo, expurgando juros excessivos.
- **Arquitetura Desacoplada (MVC):** Interface gráfica (`interface_v3.py`) operando independentemente do motor matemático (`nash.py`), com carregamento preguiçoso (*lazy loading*) de bibliotecas pesadas para inicialização instantânea da GUI.

---

## 📖 Dicionário de Regras Matemáticas

Na aba **Danos** ou **Quebra de Acordo**, selecione o critério correspondente determinado no título judicial:

*   **`R1`**: TJMG + Juros de 1% a.m. (Padrão clássico).
*   **`R2`**: Taxa Selic (critério único de correção e juros).
*   **`R3`**: TJMG + Juros de 1% a.m. até 08/2024; após, transição para Taxa Selic.
*   **`R4`**: TJMG + Juros de 1% a.m. até 08/2024; após, transição para a **Lei 14.905/24** (IPCA + Taxa Legal).
*   **`R5`**: Taxa Selic até 08/2024; após, transição para a **Lei 14.905/24** (Tema 1.368 do STJ).
*   **`R6`**: Aplicação integral da **Lei 14.905/24** (IPCA + Taxa Legal) desde a origem.
*   **`R7`**: TJMG (Correção Monetária) + Taxa Legal Retroativa (Juros de Mora).
*   *Fazenda Pública:* O sistema aplica automaticamente o Regime da EC 113 / Tema 810 do STF se a flag correspondente for ativada nos parâmetros iniciais.

---

## 📋 Passo a Passo de Uso (Nova Interface)

### Gerando um Laudo
1. **Aba "Parâmetros Gerais":** Preencha os dados básicos. Defina o *Termo Inicial dos Juros*, Justiça Gratuita (Adversa) e os Honorários. 
    * *Nota:* Caso tenha havido depósito parcial em cumprimento de sentença, altere a opção "Pagamento Voluntário 15d" para **Parcial** e informe o valor no campo correspondente para aplicação escorreita do Art. 523, §2º.
2. **Aba "Danos Principais":** Cadastre os valores históricos. Em atuações pela defesa (RÉU), informe o *Valor Pedido Inicial* para que o sistema emita o **Relatório de Êxito** (economia gerada ao cliente).
3. **Aba "Visão Estratégica":** Acompanhe o gráfico de risco financeiro gerado automaticamente com base nos valores informados.
4. **Exportação:** Clique em **"🚀 Salvar Template e Gerar Cálculo"**. O sistema salvará o banco de dados (`.xlsx`), gerará o laudo detalhado e, de forma subjacente, converterá os relatórios para PDF.

### Importação Histórica
Para atualizar um cálculo antigo, clique em **"📂 Importar Planilha Antiga"** e selecione qualquer arquivo `.xlsx` gerado pelo Nash. O sistema reconstruirá o banco de dados inteiro na tela, pronto para você adicionar novos bloqueios ou alterar regras.

---

## 💻 Instruções para Desenvolvedores e Equipe Técnica

### Instalação de Dependências
```bash
pip install -r requirements.txt

(Dependências principais: pandas, openpyxl, python-bcb, pytest para a suíte de testes)
Execução em Ambiente de Desenvolvimento

Para operar a interface nativamente via Python (necessário ter o LibreOffice instalado nas variáveis de ambiente para conversão invisível de PDF):
Bash

python interface_v3.py

Execução via Linha de Comando (Modo Backend / Headless)

Para chamadas de automação corporativa, o motor pode ser acionado silenciosamente:
Bash

python nash.py "caminho/do/template.xlsx" "caminho/do/laudo_final.xlsx"

Rotina de Compilação (CI/CD Windows)

O projeto conta com workflows no GitHub Actions que utilizam o Inno Setup para empacotar a aplicação em um instalador Windows robusto. Caso deseje compilar manualmente via PyInstaller (para binário único):
Bash

pyinstaller --noconfirm --windowed --onefile --name "nash" --icon=dr_nash.ico --add-data "dr_nash.ico;." interface_v3.py

📂 Estrutura do Repositório (v4.1)
Plaintext

├── Tabelas_Oficiais/
│   ├── tabela_tjmg.xlsx         # Matriz oficial do TJMG (Atualizar mensalmente)
│   └── cache_bcb.csv            # Cache dinâmico (Fallback) da API do Banco Central
├── interface_v3.py              # Interface Gráfica de Usuário (Tkinter) e gestão de abas
├── nash.py                      # Motor Matemático Central (Lógica contábil e jurídica)
├── test_nash.py                 # Suíte de testes unitários (pytest)
├── dr_nash.ico                  # Ícone nativo da aplicação para executáveis
├── requirements.txt             # Dependências fixadas de ecossistema
└── README.md                    # Documentação Oficial

📄 Licença

Distribuído sob a licença GNU General Public License v3.0 (GPLv3). Veja LICENSE para mais informações. Fomentamos o compartilhamento de código limpo e o livre aprimoramento de ferramentas jurídicas.