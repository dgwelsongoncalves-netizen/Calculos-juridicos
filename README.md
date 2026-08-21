# Nash System - Assistente de Cálculos Judiciais (v2.7.0)

Sistema automatizado de liquidação e atualização de débitos judiciais, desenvolvido para cálculos precisos com base na jurisprudência, Tabela do TJMG e parâmetros da taxa Selic / Lei 14.905/2024.

## 🚀 Principais Funcionalidades
- **Múltiplas Regras Matemáticas:** Suporte a critérios exclusivos (TJMG + Juros, Selic Pura, divisões de períodos pré e pós-agosto de 2024, e o novo padrão IPCA + Taxa Legal).
- **Honorários Equitativos e Sucumbência Recíproca:** Atualização monetária conforme a Súmula 14 do STJ e rateio automático de despesas.
- **Regra Trifásica de Cumprimento de Sentença:** Tratamento automatizado para o prazo de pagamento voluntário (art. 523 do CPC) conjugado com as diretrizes de Justiça Gratuita do executado.
- **Conta Gráfica (Art. 354 CC):** Amortização automática de bloqueios e depósitos judiciais ao longo do tempo.
- **Demonstrativo de Êxito:** Geração autônoma de relatório executivo comprovando a economia gerada ao cliente, com preservação total de sigilo financeiro.
- **Integração Nativa com o Banco Central:** Consulta automática e em tempo real das séries temporais do SGS (Selic, IPCA e Taxa Legal).
- **Relatórios em Excel:** Geração de laudos formatados com grades profissionais, cores corporativas e indicativos de exigibilidade.

## 💻 Requisitos
- Python 3.10 ou superior
- Bibliotecas: `pandas`, `openpyxl`, `python-bcb`, `typing-extensions`

---

## 📘 Manual de Uso (POP) para Advogados

O Nash System lê a nossa planilha padrão (`template_nash.xlsx`) e gera os laudos automaticamente, blindando o escritório contra erros matemáticos. Siga o fluxo abaixo:

### Passo 1: A Aba "Parâmetros"
Nesta aba, você informa os dados gerais do processo e a base da condenação. Preste muita atenção aos campos de honorários e proporção:
*   **Base Honorários:** Digite `CONDENAÇÃO` se estivermos executando a parte contrária. Digite `VALOR DA CAUSA` se estivermos pela Defesa e o juiz fixou honorários sobre o valor da causa (o sistema aplicará a regra do STJ: atualização desde a propositura, sem juros de mora).
*   **Data Propositura:** Essencial preencher se estivermos pela Defesa.
*   **Proporção Honorários / Custas (%):** Se houve sucumbência recíproca (ex: ganhamos 70% e perdemos 30%), digite `70%`. Se ganhamos tudo, deixe `100%`. O sistema fará o rateio automático.

### Passo 2: O Dicionário de Regras (Aba "Danos")
Na coluna **Regra**, você deve digitar o código exato correspondente ao que o juiz determinou na sentença.
*   **R1 (Padrão Antigo TJMG):** Tabela da Corregedoria do TJMG + Juros de 1% ao mês.
*   **R2 (Selic Pura):** Apenas a Taxa Selic (engloba juros e correção). Comum em restituição de tributos.
*   **R3 (Transição Selic):** TJMG + Juros de 1% ao mês até 08/2024. A partir daí, aplica-se apenas a Taxa Selic.
*   **R4 (A Mais Comum Agora):** TJMG + Juros de 1% ao mês até 08/2024. A partir daí, aplica-se automaticamente a nova **Lei 14.905/24** (IPCA + Taxa Legal do Banco Central).
*   **R5 (Transição Mista - Tema 1.368 do STJ):** Taxa Selic até 08/2024. A partir daí, aplica-se a nova **Lei 14.905/24** (IPCA + Taxa Legal).
*   **R6 (Nova Lei Pura):** Aplica a Lei 14.905/24 desde o início.
*   *Nota sobre Fazenda Pública:* Se na aba de parâmetros a "Fazenda Pública" estiver como "Sim", o sistema forçará o Tema 810 do STF e a EC 113.

### Passo 3: Como Gerar o "Relatório de Êxito" (Defesa)
Quando atuamos na defesa, precisamos demonstrar a economia gerada.
1.  **Valor Histórico:** Coloque o que o autor *efetivamente ganhou* (se ele perdeu a verba, coloque **0,00** e preencha a data do evento para base temporal).
2.  **Valor Pedido Inicial:** Coloque o que o autor *havia pedido na inicial* (o nosso risco financeiro).
3.  **Data do Pedido:** Coloque a data do ajuizamento da ação (ou aditamento).

*O sistema gerará dois arquivos: O Laudo (para juntar no PJe) e o Relatório de Êxito (documento interno para prestação de contas com o cliente).*

### Passo 4: Depósitos e Bloqueios (Aba "Deduções")
Se houve depósitos judiciais (pagamento voluntário) ou bloqueios Bacenjud/Sisbajud na conta do executado, lance na aba "Deduções" com **Data, Valor e ID do PJe**. O sistema abaterá primeiro os juros e depois o principal, parando a fluência da correção sobre a quantia garantida.

### 🚀 Como Executar o Sistema
1.  Preencha o `template_nash.xlsx` e salve-o renomeado com o número do processo.
2.  Abra o executável do **Nash System**.
3.  Clique em **"Selecionar Planilha e Calcular"**.
4.  Selecione a sua planilha. Os laudos estarão na mesma pasta prontos para uso.

---

## 📄 Licença
Este programa é um software livre: você pode redistribuí-lo e/ou modificá-lo sob os termos da **GNU Affero General Public License (AGPL)** conforme publicada pela Free Software Foundation, na versão 3 da Licença, ou (a seu critério) qualquer versão posterior.

Consulte o arquivo `LICENSE` para obter mais detalhes.