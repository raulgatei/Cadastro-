import datetime
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common import keys
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options  
from selenium.webdriver.common.by import By

#Variáveis e Arquivos
arquivo = open('Saída de Dados.txt','w')
base_curso = pd.read_excel('Curso.xlsx')
hoje = datetime.date.today().strftime("%d/%m/%Y")

def cadastro_curso():
    try:
#Verificar se o curso já está cadastrado
        driver.find_element(By.XPATH, '//*[@id="search"]/div/ul/li/input').send_keys(curso + Keys.ENTER)
        time.sleep(1)
        driver.find_element(By.XPATH, '/html/body/main/div/div/div/div[2]/div/table/tbody').text
    except:
#Tipo do Curso
        driver.find_element(By.XPATH, '/html/body/main/div/div/div/div[1]/div/div[2]/a').click()
        driver.find_element(By.XPATH, '//input[@id="tipoCurso-selectized"]').send_keys(tipo_curso)
        time.sleep(1)
#Nome do Curso
        driver.find_element(By.XPATH, '/html/body/main/div/div/form/div/div[1]/div/div[1]/div/div/div[2]/div/div[1]').click()
        driver.find_element(By.XPATH, '//*[@id="nome"]').send_keys(curso)        
#Ativo?
        driver.find_element(By.XPATH, '//input[@id="ativo-selectized"]').send_keys(ativa)
        time.sleep(1)
        driver.find_element(By.XPATH, '/html/body/main/div/div/form/div/div[1]/div/div[3]/div/div/div[2]/div/div[1]').click()
#Nome Documento e Requerimento
        driver.find_element(By.XPATH, '//input[@id="nomeDocumentos"]').send_keys(curso)
        driver.find_element(By.XPATH, '//input[@id="nomeRequerimentos"]').send_keys(curso)
#Conjunção
        driver.find_element(By.XPATH, '//input[@id="conjuncao"]').send_keys("EM")
#Titulação
        if "MBA" in curso:
            driver.find_element(By.XPATH, '//input[@id="titulacao"]').send_keys("MBA")
        else:
            driver.find_element(By.XPATH, '//input[@id="titulacao"]').send_keys("Especialista")
#Área WEB
        driver.find_element(By.XPATH, '//input[@id="area"]').send_keys(web)
#Área Conhecimento
        driver.find_element(By.XPATH, '//input[@id="areaConhecimento"]').send_keys(conhec)
        time.sleep(5)
#Atualizar a página
        driver.get('https://umc.stage.platosedu.io/erp/curso/index')
        
        print(curso, '- foi cadastrada em', hoje, file=arquivo)
    else:
        time.sleep(3)
        driver.get('https://umc.stage.platosedu.io/erp/curso/index')
        print(curso, '- já está cadastrada', file=arquivo)

driver = webdriver.Chrome()
driver.maximize_window()
driver.get('https://umc.stage.platosedu.io/erp/curso/index')
driver.find_element(By.ID, 'usuario').send_keys("bruno.pedroso")
driver.find_element(By.ID, 'password').send_keys("tubets@123" + Keys.ENTER)
time.sleep(4)

for i, curso in enumerate(base_curso['Curso']):
    tipo_curso = base_curso.loc[i, 'Tipo']
    ativa = base_curso.loc[i, 'Ativa']
    conj = base_curso.loc[i, 'Conjunção']
    titu = base_curso.loc[i, 'Titulação']
    web = base_curso.loc[i, 'WEB']
    conhec = base_curso.loc[i, 'Conhecimento']
    cadastro_curso()