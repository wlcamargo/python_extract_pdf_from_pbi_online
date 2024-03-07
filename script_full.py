from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from time import sleep
import parametrospbi
from pathlib import Path
import shutil
import os
from PyPDF2 import PdfFileReader, PdfFileWriter
import win32com.client as win32
import pandas as pd

login = 'Seu Login'
senha = 'Sua Senha'

#Link do Relatório
relatorio = ('https://app.powerbi.com/groups/a47d4552-e6f8-40ca-8150-644a24fab1be/reports/4f659da3-c421-440a-a7b0-d9d0e3db437c/ReportSection8d0dc371e1d3a2237c4a')

chrome_option = Options()
navegador = webdriver.Chrome()
navegador.get(relatorio)
sleep(10)
#Login do Power BI Service
navegador.find_element_by_xpath('//*[@id="email"]').send_keys(login)
sleep(5)
navegador.find_element_by_xpath('//*[@id="submitBtn"]').click()
sleep(10)
#Senha do Power BI Service
navegador.find_element_by_xpath('//*[@id="i0118"]').send_keys(senha)
sleep(5)
navegador.find_element_by_xpath('//*[@id="idSIButton9"]').click()
sleep(5)
navegador.find_element_by_xpath('//*[@id="idSIButton9"]').click()
sleep(5)
navegador.find_element_by_xpath('//*[(@id = "exportMenuBtn")]').click()
sleep(5)
navegador.find_element_by_css_selector('#mat-menu-panel-6 > div > button:nth-child(3) > span').click()
sleep(5)
navegador.find_element_by_xpath('//*[@id="okButton"]').click()
sleep(5)
erro = 0
while erro ==0:
    try:
        navegador.find_element_by_css_selector('* snack-bar-container > div > div > notification-toast > section > div > div > div > span').get_attribute("innerHTML")
        pass
    except:
        erro = 1
sleep(5)
navegador.close()
sleep(5)
print('Download do PDF Realizado com sucesso!')
print()
sleep(5)

#Caminho do completo do Download na tua máquina
origem_download = Path('Caminho do Arquivo completo.pdf')

# APONTA PARA O CAMINHO QUE O ARQUIVO SERÁ COPIADO
destino_download = Path('Caminho do Local onde o arquivo será transferido')

# COPIA DATASET DO LOCAL DE ORIGEM NO DESTINO
shutil.copy2(origem_download, destino_download)

# VERIFICA SE EXISTE PDF E SE EXISTIR APAGA
if os.path.exists(origem_download):
    os.remove(origem_download)

print("Arquivo transferido com sucesso!")
print (f'Local de origem {origem_download} movido para {destino_download}')
print()

#Novo Caminho do Arquivo PDF 
pdf_file_path = 'Novo Caminho do PDF.pdf'
file_base_name = pdf_file_path.replace('.pdf', '')
output_folder_path = os.path.join(os.getcwd(), 'Output')

pdf = PdfFileReader(pdf_file_path)

for page_num in range(pdf.numPages):
    pdfWriter = PdfFileWriter()
    pdfWriter.addPage(pdf.getPage(page_num))

    with open(os.path.join(output_folder_path, '{0}_Almox{1}.pdf'.format(file_base_name, page_num + 1)), 'wb') as f:
        pdfWriter.write(f)
        f.close()
print('PDF Cortado com sucesso!')
print()

#Caminho da planilha com os parâmetros do envio do e-mail
df = pd.read_excel('Caminho completo da planilha.xlsx')
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
print('E-mail enviado com sucesso!')
