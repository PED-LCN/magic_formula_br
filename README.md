# 📈 Deep Value Strategy Screener

> Ferramenta automatizada para análise fundamentalista de ações brasileiras (B3) baseada na metodologia da magic formula e Deep Value com Z-Score .

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![Status](https://img.shields.io/badge/Status-Stable-green)
![License](https://img.shields.io/badge/License-MIT-lightgrey)

## ⚠️ Disclaimer (Aviso Legal) ⚠️

Este projeto foi desenvolvido para fins **de estudo** de programação aplicada a finanças.

* **Não é recomendação de investimento:** Os rankings gerados são resultados de fórmulas matemáticas aplicadas a dados históricos. Eles não levam em conta o cenário macroeconômico, governança ou fatos relevantes recentes.
* **Riscos:** O investimento em renda variável envolve riscos de perda de capital. O ranking gerado é baseado em filtros matemáticos e dados passados, o que NÃO garante rentabilidade futura.
* **Responsabilidade:** O autor deste código não se responsabiliza por quaisquer decisões de investimento tomadas com base nos dados gerados por esta ferramenta.
* **Faça sua própria análise:** Antes de investir, busque conhecimento aprofundado sobre as empresas ou consulte um profissional financeiro certificado (CNPI).

Este projeto foi desenvolvido para automatizar o processo de *screening* (filtragem) de ações, replicando e personalizando estratégias de investimento baseada na Magic Fórmula(Joel Greenblatt) com algumas métricas adicionais para montagem do índice. Com objetivo para análises em estratégias de swing trade e stock picking.

Este script conecta-se a fontes de dados públicas, processa as informações e entrega um relatório para tomada de decisão, eliminando viés emocional e erros de cálculo. O foco principal é montar um ranking para melhor direcionar uma análise individual dos ativos, por se tratar de dados públicos é notável que alguns indicadores podem não serem fornecidos corretamente para todas as ações, o que pode gerar incorformidades em um número pequeno de ativos( a analíse individual e correta para cada ação com base na extratégia escolhida é orbigatória caso decida utilizar o ranking).

É realizado dois cortes sendo o primeiro no top 20 ações e outro na top 30, pensando em extratégias que focam em uma carteira de 15 a 25 ações, mantendo uma  margem para um stock picking mais profundo.

### ⚙️ Funcionalidades Principais

* **Coleta Automática:** Busca dados fundamentalistas (P/L, PVP, EV/EBIT, etc.) via `fundamentus` e histórico de preços via `yfinance`.
* **Filtros Inteligentes:**
    * Liquidez mínima configurável (ex: R$ 6 Milhões/dia).
    * Exclusão automática de empresas com prejuízo.
    * Seleção do ticker mais líquido por empresa (ex: prefere PETR4 a PETR3).
    * **Filtro de Risco:** Cálculo de volatilidade anualizada para exclusão dos ativos mais arriscados (Decil de Risco).
* **Algoritmo de Ranking (Z-Score):**
    * Normalização estatística dos indicadores **Earnings Yield** e **Book-to-Market**.
    * Tratamento de *Outliers* (Winsorização) para evitar distorções estatísticas.
    * Lógica híbrida para Bancos (utilizando P/L) vs Indústria (utilizando EV/EBIT).
* **Output Visual:** Geração de planilha Excel com formatação condicional (Verde/Amarelo/Vermelho) pronta para uso.

---

## 🚀 Como Executar

### Pré-requisitos

Certifique-se de ter o **Python 3.x** instalado. O projeto depende das seguintes bibliotecas:

* `pandas` (Manipulação de dados)
* `fundamentus` (Dados B3)
* `yfinance` (Dados históricos/Yahoo Finance)
* `numpy` (Cálculos matemáticos)
* `openpyxl` (Geração de Excel)

### Instalação

1. Clone o repositório:
   ```bash
   git clone [https://github.com/SEU-USUARIO/deep-value-screener.git](https://github.com/SEU-USUARIO/deep-value-screener.git)