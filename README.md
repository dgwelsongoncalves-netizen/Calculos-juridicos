Markdown

# 🏛️ Nash System (v2.8.4)
**Assistente de Cálculos Judiciais, Liquidação de Sentença e Demonstração de Êxito**

O **Nash System** é uma ferramenta interna desenvolvida em Python para automatizar o processamento de cálculos judiciais complexos, elaboração de laudos de liquidação de sentença e relatórios de proveito econômico (êxito). O sistema lê uma planilha padrão em Excel, aplica rigorosamente as tabelas oficiais e a legislação vigente, e gera arquivos auditáveis em Excel (`.xlsx`) e relatórios executivos em PDF orientados em paisagem (`landscape`).

---

## 🚀 Principais Funcionalidades

- **Motor Matemático Avançado:** Suporte a múltiplos critérios de correção monetária (Tabela TJMG, IPCA, IPCA-E, INPC) e juros de mora (1% ao mês, Taxa Selic, e transições normativas).
- **Adequação à Nova Legislação (Lei 14.905/24 & Tema 1.368 STJ):** Tratamento automatizado para o corte de agosto de 2024 (transição para IPCA + Taxa Legal do Banco Central).
- **Conta Gráfica (Amortização Art. 354 do CC):** Abatimento automático de depósitos judiciais e bloqueios (Sisbajud), separando juros e principal para estancar a correção sobre o montante garantido.
- **Governança por Atuação (Autor vs. Réu):** Configuração inteligente que gera o Laudo Oficial para o PJe e oculta automaticamente o Relatório de Êxito quando o escritório atua pela parte ativa.
- **Marcos Temporais Desacoplados:** Capacidade de fixar datas de juros independentes por verba, blindando o cálculo contra divergências sentenciais.
- **Exportação Multiplataforma para PDF:** Conversão automática dos laudos e demonstrativos em arquivos PDF perfeitamente formatados usando o LibreOffice em modo *headless* (compatível com Windows e Linux).

---

## 📖 Dicionário de Regras Matemáticas

Na aba **Danos** da planilha padrão, utilize os seguintes códigos de regra conforme determinado na sentença:

*   **`R1`**: TJMG + Juros de 1% a.m. (Padrão clássico).
*   **`R2`**: Taxa Selic (critério único de correção e juros).
*   **`R3`**: TJMG + Juros de 1% a.m. até 08/2024; após, transição para Taxa Selic.
*   **`R4`**: TJMG + Juros de 1% a.m. até 08/2024; após, transição para a **Lei 14.905/24** (IPCA + Taxa Legal).
*   **`R5`**: Taxa Selic até 08/2024; após, transição para a **Lei 14.905/24** (Tema 1.368 do STJ).
*   **`R6`**: Aplicação integral da **Lei 14.905/24** desde a origem.
*   *Nota Fazenda Pública:* Se a chave correspondente estiver ativa, o sistema aplica automaticamente o Regime da EC 113 / Tema 810 do STF.

---

## 📋 Passo a Passo de Uso (Como Preencher e Rodar)

### Passo 1: Preenchimento da Planilha Padrão (`template_nash.xlsx`)
1. **Aba "Parametros":** Insira os dados gerais do processo. Preste atenção especial na linha **Atuação** (digite `RÉU` se estivermos na defesa para gerar o relatório de êxito, ou `AUTOR` se estivermos na ativa para gerar apenas o laudo principal) e na **Base Honorários** (`CONDENAÇÃO` ou `VALOR DA CAUSA`).
2. **Aba "Danos":** Lance os valores históricos, datas de desembolso e a regra aplicável (`R1` a `R6`). Caso a sentença fixe marcos temporais diferentes para a correção e para os juros, utilize a coluna opcional **Data Juros**. Se estiver atuando pela defesa (`RÉU`), preencha também as colunas de **Valor Pedido Inicial** e **Data do Pedido** para calcular a economia gerada.
3. **Aba "Custas":** Lance as custas e despesas processuais comprovadas.
4. **Aba "Deducoes":** Caso existam depósitos judiciais ou bloqueios (Sisbajud), lance as datas e valores para que o sistema monte a Conta Gráfica (Art. 354 do CC).
5. Salve e feche a planilha (recomenda-se renomeá-la com o número do processo, ex: `Planilha_5001234.xlsx`).

---

## 💻 Instruções de Execução

### Opção A: Uso no Windows (Equipe e Advogados Associados)
1. Certifique-se de que o **LibreOffice** está instalado na máquina (para que a conversão automática para PDF funcione em segundo plano).
2. Dê dois cliques no executável do **Nash System** disponibilizado pelo escritório.
3. Clique em **"📂 Selecionar Planilha e Calcular"**.
4. Escolha a planilha preenchida do processo. O sistema processará os dados em background e gerará os arquivos `.xlsx` e `.pdf` diretamente na mesma pasta do arquivo original.

### Opção B: Uso no Linux / Desenvolvimento (Terminal ou VS Code)
1. Abra o seu terminal ou o **VS Code** na pasta raiz do projeto.
2. Certifique-se de ativar o seu ambiente virtual (caso utilize) e de que o LibreOffice está instalado (`sudo dnf install libreoffice` ou `sudo apt install libreoffice`).
3. Instale ou atualize as dependências do Python:
   ```bash
   pip install -r requirements.txt

    Execute o script principal da interface gráfica:
    Bash

    python nash.py

    Na janela que se abrir, clique no botão de seleção de planilha e aponte para o arquivo desejado.

📂 Estrutura do Repositório
Plaintext

├── Tabelas_Oficiais/
│   └── tabela_tjmg.xlsx         # Índice oficial de correção do TJMG
├── template_nash.xlsx           # Planilha modelo padrão do escritório
├── nash.py                      # Código-fonte principal (Motor + Interface GUI)
├── requirements.txt             # Dependências do projeto (pandas, openpyxl, bcb, etc.)
└── README.md                    # Documentação oficial

Desenvolvido para otimizar o fluxo de trabalho e assegurar precisão matemática absoluta em laudos judiciais.