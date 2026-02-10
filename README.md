# 💰 Consolidador de Relatórios Financeiros

Uma aplicação web desenvolvida em **Python** e **Streamlit** para automatizar o tratamento, limpeza e consolidação de planilhas financeiras de um cliente específico.

O sistema processa arquivos brutos, corrige formatações quebradas, separa dados de fornecedores (CNPJ/Nome) e exporta um relatório unificado pronto para análise contábil.

## 🚀 Funcionalidades

- **Upload Múltiplo:** Aceita vários arquivos Excel (.xlsx) de uma vez.
- **Identificação Automática:** Reconhece a conta e o código contábil baseando-se no nome do arquivo.
- **Limpeza de Dados:**
  - Remove cabeçalhos e rodapés inúteis.
  - Reconstrói descrições e históricos que foram "quebrados" em múltiplas linhas pelo banco.
  - Separa automaticamente o **CNPJ** do **Nome do Fornecedor**.
- **Exportação:** Gera um único arquivo Excel consolidado e padronizado.

## 🛠️ Pré-requisitos

Para rodar o projeto localmente, você precisa ter o [Python](https://www.python.org/) instalado.

## Link web:
[ https://tratamento-de-dados---relatorio-financeiro-wh9ytxpqha6ynnmlmsa.streamlit.app/ ]
