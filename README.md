# Projeto de Agendamento de Pedidos

Criei esse script em Python visando automatizar o processo de envio de e-mails para fornecedores que possuem pedidos sem agendamento, de acordo com as regras estabelecidas. O projeto utiliza as bibliotecas `win32com.client` para interação com o Outlook, `pandas` para manipulação de dados em formato tabular e `datetime` para lidar com informações de data e hora.

## Leitura da Base de Dados

```python
# Bibliotecas utilizadas no projeto
import win32com.client as win32
import pandas as pd
from datetime import datetime

# Base da carteira de pedidos
base = pd.read_excel("planilhas/BASE_DASHBOARD.xlsx")

# Data atual
data_atual = datetime.now()

# Filtrar pedidos com data de entrega maior que a data atual
base_dashboard = base[base['DT_ENTREGA'] >= data_atual]

# Leitura das bases de e-mails dos fornecedores
emails_forn = pd.read_excel("planilhas/emails_forn.xlsx")
emails_amigao = pd.read_excel("planilhas/emails_amigao.xlsx")

