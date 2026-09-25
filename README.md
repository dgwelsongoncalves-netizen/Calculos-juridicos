Markdown

# 🏛️ Nash System (v3.1)
**Assistente de Cálculos Judiciais, Liquidação de Sentença e Demonstração de Êxito**

O **Nash System** é uma ferramenta interna desenvolvida em Python para automatizar o processamento de cálculos judiciais complexos, elaboração de laudos de liquidação de sentença e relatórios de proveito econômico (êxito). O sistema integra uma interface gráfica intuitiva para entrada de dados com um robusto motor matemático (desacoplado em background) que aplica rigorosamente as tabelas oficiais e a legislação vigente, gerando arquivos auditáveis em Excel (`.xlsx`) e relatórios executivos em PDF orientados em paisagem (`landscape`).

---

## 🚀 Principais Funcionalidades (Nova Versão 3.1)

- **Arquitetura Desacoplada (MVC):** Interface gráfica moderna (`interface_v3.py`) separada do motor matemático (`nash.py`), garantindo maior estabilidade, testes isolados e facilidade de compilação.
- **Interface Inteligente de Entrada de Dados:** Formulários interativos com seleção em *dropdowns*, funcionalidade de edição direta nas linhas inseridas (Danos e Custas) e escolha de diretório de destino na hora de salvar.
- **Importação de Planilhas Antigas:** Capacidade de importar dados, parâmetros e histórico de verbas a partir de templates ou laudos previamente gerados, evitando retrabalho de digitação.
- **Motor Matemático Avançado:** Suporte a múltiplos critérios de correção (Tabela TJMG, IPCA, IPCA-E, INPC) e juros de mora (1% a.m., Taxa Selic). Adequação total à **Lei 14.905/24** e ao Tema 1.368 do STJ.
- **Conta Gráfica (Amortização Art. 354 do CC):** Abatimento automático de depósitos judiciais e bloqueios (Sisbajud).
- **Marcos Temporais Personalizáveis:** Flexibilidade para aplicar Juros de Mora a partir do Evento, Desembolso, Citação, Trânsito em Julgado ou **Propositura**.
- **Demonstrativo de Êxito Automático:** Geração automática do relatório de economia do cliente (quando atuando como "RÉU"), oculto inteligentemente quando atuando como "AUTOR".

---

## 📖 Dicionário de Regras Matemáticas

Na aba **Danos** da interface, selecione a regra correspondente conforme determinado na sentença:

*   **`R1`**: TJMG + Juros de 1% a.m. (Padrão clássico).
*   **`R2`**: Taxa Selic (critério único de correção e juros).
*   **`R3`**: TJMG + Juros de 1% a.m. até 08/2024; após, transição para Taxa Selic.
*   **`R4`**: TJMG + Juros de 1% a.m. até 08/2024; após, transição para a **Lei 14.905/24** (IPCA + Taxa Legal).
*   **`R5`**: Taxa Selic até 08/2024; após, transição para a **Lei 14.905/24** (Tema 1.368 do STJ).
*   **`R6`**: Aplicação integral da **Lei 14.905/24** (IPCA + Taxa Legal) desde a origem.
*   **`R7`**: TJMG (Correção) + Taxa Legal Retroativa (Juros)
*   *Nota Fazenda Pública:* Se a chave correspondente estiver ativa nos parâmetros, o sistema aplica automaticamente o Regime da EC 113 / Tema 810 do STF.

---

## 📋 Passo a Passo de Uso (Na Nova Interface)

### Como Preencher e Rodar o Cálculo
1. Abra o **Nash System**.
2. **Aba "Parâmetros Gerais":** Preencha os dados do processo. Defina a **Atuação** (`RÉU` ou `AUTOR`), o **Termo Inicial dos Juros** (Citação, Evento, Propositura, etc.), as datas base da sentença e o percentual/valor fixo de Honorários.
3. **Aba "Danos":** Clique em "Adicionar Nova Linha". Preencha o Valor Histórico e a Data de Desembolso. Selecione a Regra (ex: `R2`). Para defesas, informe o *Valor Pedido Inicial* e a *Data do Pedido* para cálculo automático do êxito. Pode editar linhas clicando em "Editar Selecionada".
4. **Aba "Custas":** Adicione as guias e despesas processuais comprovadas. As custas adotarão sempre a correção da Lei 14.905/24 automaticamente se a ação for contra particular.
5. Clique em **"🚀 Salvar Template e Gerar Cálculo"**. Escolha a pasta do cliente no seu computador onde deseja guardar os arquivos. O motor operará em background e emitirá um aviso de sucesso.

### Como Importar um Cálculo Antigo
1. Na tela principal, clique no botão **"📂 Importar Planilha Antiga"**.
2. Selecione qualquer `.xlsx` gerado anteriormente pelo Nash (seja template de entrada ou laudo final).
3. O sistema carregará todos os parâmetros, danos e custas de volta para a tela, prontos para você revisar, editar ou corrigir datas.

---

## 💻 Instruções para Desenvolvedores e Equipe Técnica

### Dependências
```bash
pip install -r requirements.txt

(Pacotes principais: pandas, openpyxl, python-bcb, pytest para a suíte de testes)
Execução em Ambiente de Desenvolvimento (Linux/Windows)

    Certifique-se de que o LibreOffice está instalado (necessário para a conversão invisível para PDF).

    Para abrir a aplicação visual, execute o ficheiro de interface:
    Bash

    python interface_v3.py

Execução via Linha de Comando (Modo Backend / CLI)

Para automações ou integração de sistemas, o motor matemático pode ser invocado via terminal passando os caminhos do template e do arquivo de saída como argumentos:
Bash

python nash.py "caminho/do/template.xlsx" "caminho/do/laudo_final.xlsx"

Rotina de Compilação para Windows (PyInstaller)

O sistema possui uma pipeline no GitHub Actions para empacotar o código nativamente. Para compilar manualmente, utilize:
Bash

pyinstaller --noconfirm --windowed --onefile --name "nash" --icon=dr_nash.ico --add-data "dr_nash.ico;." interface_v3.py

Note que o alvo do empacotador agora é o interface_v3.py, garantindo que o .exe abrirá a tela gráfica, enquanto o nash.py é incluído implicitamente.

📂 Estrutura do Repositório (v3.1)
Plaintext

├── Tabelas_Oficiais/
│   └── tabela_tjmg.xlsx         # Tabela matriz oficial de correção do TJMG (Atualizar mensalmente)
├── interface_v3.py              # Interface Gráfica de Usuário (Tkinter) e gestão de arquivos
├── nash.py                      # Motor Matemático Central (Processamento Pandas e regras jurídicas)
├── test_nash.py                 # Suíte rigorosa de testes unitários (pytest)
├── dr_nash.ico                  # Ícone da aplicação para Windows
├── requirements.txt             # Dependências fixadas
└── README.md                    # Documentação Oficial

Desenvolvido para otimizar fluxos de trabalho e assegurar precisão matemática absoluta em rotinas judiciais.