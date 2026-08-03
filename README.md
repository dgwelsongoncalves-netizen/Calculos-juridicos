# Nash System - Assistente de Cálculos Judiciais (v2.3.3)

Sistema automatizado de liquidação e atualização de débitos judiciais, desenvolvido para cálculos precisos com base na jurisprudência, Tabela do TJMG e parâmetros da taxa Selic / Lei 14.905/2024.

## 🚀 Principais Funcionalidades
- **Múltiplas Regras Matemáticas:** Suporte a critérios exclusivos (TJMG + Juros, Selic Pura, divisões de períodos pré e pós-agosto de 2024, e o novo padrão IPCA + Taxa Legal).
- **Honorários Equitativos com Datas Divididas:** Atualização monetária desde a sentença e juros de mora aplicados a partir do trânsito em julgado.
- **Regra Trifásica de Cumprimento de Sentença:** Tratamento automatizado para o prazo de pagamento voluntário (art. 523 do CPC) conjugado com as diretrizes de Justiça Gratuita do executado.
- **Integração Nativa com o Banco Central:** Consulta automática e em tempo real das séries temporais do SGS (Selic, IPCA e Taxa Legal).
- **Relatórios Executivos em Excel:** Geração de laudos formatados com grades profissionais, cores corporativas e indicativos de exigibilidade.

## 💻 Requisitos
- Python 3.10 ou superior
- Bibliotecas: `pandas`, `openpyxl`, `python-bcb`, `typing-extensions`

## 📄 Licença
Este programa é um software livre: você pode redistribuí-lo e/ou modificá-lo sob os termos da **GNU Affero General Public License (AGPL)** conforme publicada pela Free Software Foundation, na versão 3 da Licença, ou (a seu critério) qualquer versão posterior.

Consulte o arquivo `LICENSE` para obter mais detalhes.