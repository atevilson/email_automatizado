# Projeto email automatizado

Criei esse script em Python visando automatizar o processo de cobrança de agendamento em massa de fornecedores que possuem pedidos sem agendamento, de acordo com as regras estabelecidas. O projeto utiliza as bibliotecas `win32com.client` para interação com o Outlook, `pandas` para manipulação e tratamento de dados em formato tabular e `datetime` para lidar com informações de data e hora.

## Leitura da Base de Dados

```python


# Base da carteira de pedidos
base = pd.read_excel("planilhas/BASE_DASHBOARD_PEDIDOS_COMPRA.xlsx")
emails_forn = pd.read_excel("planilhas/emails_forn.xlsx")
emails_amigao = pd.read_excel("planilhas/emails_amigao.xlsx")
```
---

## Estrutura para base de contatos de fornecedores e dos departamentos da empresa em questão

### Tabela: Usuários

| nome_usuario | email-usuario                                     | depto          |
|--------------|---------------------------------------------------|----------------|
| usuario1     | exampler@email.com; exampler@email.com; exampler@email.com | depto_example  |
| usuario2     | exampler@email.com; exampler@email.com; exampler@email.com | depto_example  |


### Tabela: Fornecedores

| cod_fornecedor | nome_fornecedor  | email                 | email_forn_cc                                    |
|----------------|------------------|-----------------------|--------------------------------------------------|
| 1010101        | nome_fornecedor1 | email_forn@example.com | example1@email.com; example2@email.com           |
| 1010101        | nome_fornecedor2 | email_forn@example.com | example1@email.com; example2@email.com           |
| 1010101        | nome_fornecedor3 | email_forn@example.com | example1@email.com; example2@email.com           |




## Demais destaques do Código

### No dataframe base_dashboard
- Filtragem dos pedidos com data de entrega maior ou igual à data atual.

### Filtragem de Fornecedores
- Identificação de fornecedores sem agendamento e válidos de acordo com as novas regras.

### Envio de E-mails para Fornecedores "Sem agendamento de entrega"
- Utilização da biblioteca `win32com.client` para interação com o Outlook.
- Criação de um dicionário para armazenar os pedidos por fornecedor e usuário.
- Laço de repetição para agrupar os pedidos no dicionário.
- Laço de repetição para enviar e-mails aos fornecedores e usuários.
- Construção do corpo do e-mail com informações relevantes.
- Verificação e adição de cópia em CC com base nas regras de departamento.
- Tratamento de erros durante o envio de e-mails "try/exception".

### Construção do Corpo do E-mail
- Adição das informações relevantes no corpo do e-mail, como fornecedor, código do fornecedor, pedidos, departamento, etc.
- Adição da data de emissão e lead time para cada pedido.
- Verificação de correspondência do departamento entre as bases `base` e `emails_amigao`.
- Adição de cópia em CC para o usuário correspondente ao departamento.

### Destaques Adicionais
- Utilização de boas práticas de programação, como a modularização do código em trechos específicos.
- Mensagens de log para informar sobre a execução do script, destacando sucesso ou falha no envio dos e-mails.

## Execução do Script
Para executar o script, certifique-se de ter as bibliotecas necessárias instaladas. Você pode instalar as dependências usando:

```bash
pip install pandas pywin32
pip install openpyxl
```
---

[Confira o código no notebook do Jupyter](https://github.com/atevilson/email_automatizado/blob/main/email_forn_sem_agenda.ipynb)


### Autor
---

<a href="https://medium.com/@freitas.atevilson/inova%C3%A7%C3%A3o-sim-todos-podemos-inovar-18934cfb787e">
 <img style="border-radius: 50%;" src="https://avatars.githubusercontent.com/u/62858618?s=400&u=5f6e68fa29a7808de7e4954f4017bae120585572&v=4" width="100px;" alt=""/>
 <br />
 <sub><b>Atevilson Freitas</b></sub></a> <a href="https://medium.com/@freitas.atevilson/inova%C3%A7%C3%A3o-sim-todos-podemos-inovar-18934cfb787e">🚀</a>


Envio de email outlook automatizado

[![Linkedin Badge](https://img.shields.io/badge/LinkedIn-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/atevilson-freitas/) 
[![Medium Badge](https://img.shields.io/badge/Medium-12100E?style=for-the-badge&logo=medium&logoColor=white)](https://medium.com/@freitas.atevilson/inova%C3%A7%C3%A3o-sim-todos-podemos-inovar-18934cfb787e)
