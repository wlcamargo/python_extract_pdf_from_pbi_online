from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from time import sleep

#Digite abaixo seu login e senha
login = 'Seu Login'
senha = 'Sua senha'

#Link do Relatório
relatorio = ('Caminho do Relatório')

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
sleep(5)

