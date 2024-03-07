import win32com.client as win32
import pandas as pd

df = pd.read_excel('C:/Users/walla/OneDrive/Área de Trabalho/Reports PBI/Estoque_Almox/Parametros_Envio_Almox.xlsx')
for i, contato in enumerate(df['EMAIL']):
    indice = df.loc[i,"ID_ARQUIVO "]
    arquivo = df.loc[i, "ANEXO"]

    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)

    email.To = contato
    email.Subject = f'PDF Almoxarifado {indice}'
    email.HTMLBody = "Olá, segue o PDF."
    email.Attachments.Add(arquivo)
    email.Send()
print('e-mail enviado com sucesso!')
