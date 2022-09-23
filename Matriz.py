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
base_curso = pd.read_excel('Matriz.xlsx')
hoje = datetime.date.today().strftime("%d/%m/%Y")
base_curso['Carga'] = base_curso['Carga'].astype(object)

def cadastro_matriz():
    try:
        driver.find_element(By.XPATH, '//*[@id="search"]/div/ul/li/input').send_keys(curso + Keys.ENTER)
        time.sleep(1)
        driver.find_element(By.XPATH, '/html/body/main/div/div/div/div[2]/div/table/tbody').text
    except:
        print(curso, "- Não cadastrado", hoje, file=arquivo)
    else:
        driver.find_element(By.XPATH, '/html/body/main/div/div/div/div[2]/div/table/tbody/tr[1]/td[5]/div/button/i').click()
        driver.find_element(By.XPATH, '/html/body/main/div/div/div/div[2]/div/table/tbody/tr[1]/td[5]/div/div/a[4]').click()
        time.sleep(1)
        validacao = driver.find_element(By.XPATH, '/html/body/main/div/div/div/div[2]/div/table/tbody/tr/td[2]').text
        if validacao == matriz:
            print(matriz, '- já cadastrada', hoje, file=arquivo)
        else:
            driver.find_element(By.XPATH,'/html/body/main/div/div/div/div[1]/div/div[2]/a').click()
            time.sleep(1)
            driver.find_element(By.XPATH,'//input[@id="tipoMatriz-selectized"]').send_keys(tipo_matriz)
            driver.find_element(By.XPATH,'/html/body/main/div/div/form/div/div[1]/div[1]/div[1]/div/div/div[2]/div/div[1]').click()
            driver.find_element(By.XPATH,'//input[@id="nome"]').send_keys(matriz)
            driver.find_element(By.XPATH,'//input[@id="ch"]').send_keys(ch)
            driver.find_element(By.XPATH,'//input[@id="resolucaoDocumento"]').send_keys(resolucao)
            driver.find_element(By.XPATH,'//input[@id="percentualMestresDoutores"]').send_keys(per1)
            driver.find_element(By.XPATH,'//input[@id="percentualMestres"]').send_keys(per2)
            driver.find_element(By.XPATH,'//input[@id="percentualDoutores"]').send_keys(per3)
            driver.find_element(By.XPATH,'//*[@id="apostilamento"]').send_keys(apost)
            driver.find_element(By.XPATH,'//*[@id="descricaoSite"]').send_keys(desc)
            driver.find_element(By.XPATH,'//*[@id="resumo"]').send_keys(resumo)
            driver.find_element(By.XPATH,'//*[@id="introducao"]').send_keys(introd)
            driver.find_element(By.XPATH,'//*[@id="certificacao"]').send_keys(cert)
            driver.find_element(By.XPATH,'//*[@id="objetivo"]').send_keys(obje)
            driver.find_element(By.XPATH,'//*[@id="metodologia"]').send_keys(metodologia)
            driver.find_element(By.XPATH,'//*[@id="publico"]').send_keys(publico)
            driver.find_element(By.XPATH,'//*[@id="ementa"]').send_keys(ementa)
            driver.find_element(By.XPATH,'//*[@id="periodicidade"]').send_keys(periodicidade)
            time.sleep(4)

driver = webdriver.Chrome()
driver.maximize_window()
driver.get('https://umc.stage.platosedu.io/erp/curso/index')
driver.find_element(By.ID, 'usuario').send_keys("bruno.pedroso")
driver.find_element(By.ID, 'password').send_keys("tubets@123" + Keys.ENTER)
time.sleep(4)

for i, curso in enumerate(base_curso['Curso']):
    tipo_matriz = base_curso.loc[i, 'Tipo Matriz']
    matriz = base_curso.loc[i,'Nome']
    ch = base_curso.loc[i, 'Carga']
    resolucao = base_curso.loc[i,'Resolução']
    per1 = base_curso.loc[i,'Perc 1']
    per2 = base_curso.loc[i,'Perc 2']
    per3 = base_curso.loc[i,'Perc 3']
    apost = base_curso.loc[i,'Apostilamento']
    desc = base_curso.loc[i,'Descricão']
    resumo = base_curso.loc[i,'Resumo']
    introd = base_curso.loc[i,'Introdução']
    cert = base_curso.loc[i,'Certificação']
    obje = base_curso.loc[i,'Objetivo']
    metodologia = base_curso.loc[i,'Metodologia']
    publico = base_curso.loc[i,'Publico']
    ementa = base_curso.loc[i,'Ementa']
    periodicidade = base_curso.loc[i,'Periodicidade']
    cadastro_matriz()
driver.quit()