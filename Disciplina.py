import datetime
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common import keys
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options  
from selenium.webdriver.common.by import By

arquivo = open('Saída de Dados.txt','w')
base = pd.read_excel('Disciplinas.xlsx')
hoje = datetime.date.today().strftime("%d/%m/%Y")

def cadastro_disciplina():
    try:
        driver.find_element(By.XPATH, '//*[@id="search"]/div/ul/li/input').send_keys(disciplinas + Keys.ENTER)
        time.sleep(1)
        driver.find_element(By.XPATH, '/html/body/main/div/div/div/div[2]/div/table/tbody').text
    except:
        driver.find_element(By.XPATH, '/html/body/main/div/div/div/div[1]/div/div[2]/a/i').click()
        time.sleep(2)
        driver.find_element(By.XPATH, '//*[@id="nome"]').send_keys(disciplinas)
        driver.find_element(By.XPATH, '//input[@id="tipoDisciplina-selectized"]').send_keys(tipo_disciplina)
        time.sleep(2)
        driver.find_element(By.XPATH, '/html/body/main/div/div/div/form/div[1]/div/div[2]/div/div/div[2]/div/div[1]').click()
        driver.find_element(By.XPATH, '//input[@id="ativo-selectized"]').send_keys(ativa)
        time.sleep(2)
        driver.find_element(By.XPATH, '/html/body/main/div/div/div/form/div[1]/div/div[3]/div/div[2]/div/div[1]').click()
        driver.get('https://umc.stage.platosedu.io/erp/disciplina/index')
        
        print(disciplinas, '- foi cadastrada em', hoje, file=arquivo)
    else:
        time.sleep(3)
        driver.get('https://umc.stage.platosedu.io/erp/disciplina/index')
        print(disciplinas, '- já está cadastrada', file=arquivo)

driver = webdriver.Chrome()
driver.maximize_window()
driver.get('https://umc.stage.platosedu.io/erp/disciplina/index')
driver.find_element(By.ID, 'usuario').send_keys("bruno.pedroso")
driver.find_element(By.ID, 'password').send_keys("tubets@123" + Keys.ENTER)
time.sleep(4)

for i, disciplinas in enumerate(base['Disciplina']):
    tipo_disciplina = base.loc[i, 'Tipo']
    ativa = base.loc[i, 'Ativa']
    insti = base.loc[i, 'Instituição']
    cadastro_disciplina()