import win32com.client as client
import pandas as pd
import datetime as dt

tabela = pd.read_excel("Contas a Receber.xlsx")

hoje = dt.datetime.now()

tabela.columns = tabela.columns.str.strip()
tabela["Data Prevista para pagamento"] = pd.to_datetime(
    tabela["Data Prevista para pagamento"], errors = "coerce"
)

tab_devedores = tabela.loc[
    (tabela["Status"] == "Em aberto") &
    (tabela["Data Prevista para pagamento"] < hoje)
]

outlook = client.Dispatch("Outlook.Application")

dados = tab_devedores[["Valor em aberto", "Data Prevista para pagamento", "E-mail", "NF"]].values.tolist()

for valor, data, email_cliente, nf in dados:
    msg = outlook.CreateItem(0)
    msg.To = "latasak180@bdnets.com"
    msg.Subject = f"Cliente Devedor - NF {nf}"
    msg.Body = f"Valor: {valor}\n"
    f"Vencimento: {data.strftime("%d/%m/%y")}\n"
    f"Email: {email_cliente}"
    msg.Send()
