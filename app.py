import openpyxl
import webbrowser
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver import Keys
from selenium.common.exceptions import *
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as condicao_esperada
import random
import schedule
import re
from datetime import datetime

def iniciar_driver():
    chrome_options = Options()

    arguments = ['--lang=pt-BR', '--window-size=1300,800',
                 '--incognito']
    for argument in arguments:
        chrome_options.add_argument(argument)

    chrome_options.add_experimental_option("prefs", {
        'download.directory_upgrade': True,
        'download.prompt_for_download': False,
        "profile.default_content_setting_values.notifications": 2,
        "profile.default_content_setting_values.automatic_downloads": 1,
    })
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(
        driver,
        10,
        poll_frequency=1,
        ignored_exceptions=[
            NoSuchElementException,
            ElementNotVisibleException,
            ElementNotSelectableException,
        ]
    )
    return driver, wait

def digitar_naturalmente(texto, elemento):
    for letra in texto:
        elemento.send_keys(letra)
        sleep(random.randint(1, 5)/30)

def consulta_de_precos():
    driver, wait = iniciar_driver()
    print('---------------------------------------------')
    print('Iniciando automação!')
    print('---------------------------------------------')
    driver.get('https://www.magazineluiza.com.br/')
    sleep(2)
    print('Site aberto com sucesso!')
    campo_pesquisa = wait.until(condicao_esperada.element_to_be_clickable((By.ID,'input-search')))

    print('Iniciando busca pelo produto!')
    digitar_naturalmente('Samsung Galaxy S24 Ultra',campo_pesquisa)
    campo_pesquisa.send_keys(Keys.ENTER)

    sleep(10)
    produto = wait.until(condicao_esperada.visibility_of_all_elements_located((By.XPATH,"//img[@class='sc-cWSHoV iJPAvC']")))
    print('Produto encontrado!')
    sleep(1.5)
    produto[0].click()
    sleep(4)
    print('Consultando dados do produto...')
    sleep(1)
    driver.execute_script("window.scrollTo(0, 500);")
    sleep(2.5)

    precos = driver.find_elements(By.XPATH,"//div[@class='sc-dhKdcB ryZxx']/p[@class='sc-kpDqfm eCPtRw sc-bOhtcR dOwMgM']")
    preco = precos[0].text
    texto_numerico = re.sub(r'[^\d\.,]', '', preco)
    texto_numerico = texto_numerico.replace(',','.')
    preco_produto = texto_numerico
    
    sleep(1)
    nome_xpath = driver.find_element(By.XPATH,"//h1[@data-testid='heading-product-title']")
    nome_produto = nome_xpath.text
    
    data_consulta = datetime.strftime(datetime.now(),"%d/%m/%Y, %H:%M")
    
    link_produto = driver.current_url
    
    sleep(2)
    print('Dados coletados!')
    sleep(5)
    print('Criando planilha...')


    # Criar planilha
    
    workbook = openpyxl.Workbook()
    del workbook['Sheet']
    workbook.create_sheet('Consulta de Preços')
    sheet_consulta = workbook['Consulta de Preços']
    sheet_consulta.append(['Nome do Produto','Preço (R$)','Data de Consulta','Link do Produto'])
    sheet_consulta.append([f'{nome_produto}',f'{preco_produto}',f'{data_consulta}',f'{link_produto}'])
    workbook.save('consulta.xlsx')



    print('Planilha criada com sucesso!')
    sleep(2)
    print('Voltaremos em 30 minutos...')
    driver.close()

schedule.every(30).minutes.do(consulta_de_precos)

print(f'Próximo agendamento irá ocorrer as {schedule.next_run()}')

while True:
    schedule.run_pending()
    sleep(1)


    






