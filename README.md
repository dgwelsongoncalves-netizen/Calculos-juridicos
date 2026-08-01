# Nash System - Liquidação Judicial Automatizada ⚖️

O **Nash System** é uma esteira de automação desenvolvida em Python para padronizar, acelerar e garantir precisão matemática na liquidação de sentenças judiciais (focado nas diretrizes do TJMG e regras do STJ/STF).

O sistema processa planilhas padronizadas, conecta-se à API do Banco Central para captura da Taxa Selic temporal, cruza com as Tabelas Práticas dos Tribunais e gera um Laudo Pericial completo, detalhando a memória de cálculo.

## 📁 Estrutura de Pastas e Esteira de Produção

A arquitetura foi desenhada para manter a organização contínua dos processos do escritório:

* **`Processos_Entrada/` (A Caixa de Entrada):** Contém o arquivo `template_nash.xlsx`. É aqui que você deve depositar a cópia da planilha preenchida do cliente antes de iniciar a automação.
* **`Processos_Calculados/` (A Saída):** Onde o laudo finalizado (`.xlsx` e futuramente `.pdf`) será salvo, pronto para ser juntado aos autos.
* **`Processos_Arquivados/` (O Histórico):** Após o processamento com sucesso, a planilha original de entrada é movida automaticamente para esta pasta, mantendo a caixa de entrada limpa.
* **`Tabelas_Oficiais/`:** Repositório local com os índices de correção monetária estáticos (ex: Tabela Prática do TJMG).

## ⚙️ Dicionário de Regras Matemáticas

Ao preencher o template, utilize as seguintes chaves na coluna "Regra":

* **R1:** TJMG + Juros de 1% a.m. até 30/08/2024; após, Taxa Selic (Transição da Lei Nova).
* **R2:** Selic desde o evento até vigência da Lei Nova.
* **R3:** IPCA + (Selic deduzida do IPCA) desde o desembolso.
* **R4:** Taxa Selic (critério único) durante todo o período.
* **R5:** Tabela TJMG + Juros de 1% a.m. como critério único em todo o período (Apenas Lei Antiga).

## 🚀 Como Executar Localmente

1. Garanta que o ambiente virtual está ativo.
2. Certifique-se de que a planilha do cliente está salva em `Processos_Entrada/`.
3. No terminal, execute:
   ```bash
   python nash.py