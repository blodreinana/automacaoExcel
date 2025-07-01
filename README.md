# Automação de Cotações com Python e Excel

Este projeto é um estudo prático de **automação com Python**, integrando **web scraping com BeautifulSoup**, requisições com `requests` e manipulação de planilhas Excel usando o `pandas`.

O objetivo é buscar automaticamente as cotações atuais de ações brasileiras e salvar essas informações em uma planilha Excel organizada.

---

##  Funcionalidades

- Geração automática de uma planilha com nomes e códigos de empresas da bolsa.
- Consulta em tempo real das cotações pelo site do Google Finance.
- Armazenamento das cotações em um novo arquivo Excel (`cotacoes_atualizadas.xlsx`).
- Tratamento de erros de conexão e páginas não encontradas.
- Pausa entre as requisições para evitar bloqueios.

---

##  Tecnologias usadas

- [Python 3.13+](https://www.python.org/)
- [Pandas](https://pandas.pydata.org/)
- [BeautifulSoup4](https://www.crummy.com/software/BeautifulSoup/)
- [Requests](https://docs.python-requests.org/)
- [openpyxl](https://openpyxl.readthedocs.io/) (para leitura/escrita de arquivos Excel)

---

Estrutura do projeto
bash
Copiar
Editar
automacao-excel-cotacoes/
├── main.py                   # Script principal com toda a automação
├── empresas_info.xlsx        # Gerado automaticamente com nome/código das empresas
└── cotacoes_atualizadas.xlsx # Resultado final com cotações atualizadas

 Autoria
Feito por Allana Helena Campos Albino com durante os estudos de automação em Python e Análise de Dados.
Este projeto é de uso educacional e livre para adaptação.



