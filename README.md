
Este script em Python automatiza o envio de e-mails via Microsoft Outlook para clientes com faturas em aberto e vencidas, com base em uma planilha Excel de Contas a Receber, podendo mudar para oque desejar.

---- Requisitos ----

- Python 3.x
- Microsoft Outlook instalado e configurado no computador, OBS: nao pode ser o outlook da microsoft store!!!
- Pacotes Python:
  - `pandas`
  - `pywin32` (`win32com.client`)
  - `openpyxl` (para ler arquivos .xlsx)

Instale as dependências com:

`pip install pandas pywin32 openpyxl`

