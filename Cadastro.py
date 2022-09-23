import datetime
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common import keys
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options  
from selenium.webdriver.common.by import By

# Arquivos de leitura e gravação
arquivo = open('Cadastro de Disciplina.txt','w')
arquivo_curso = open('Cadastro de Curso.txt','w')
arquivo_matriz = open('Cadastro de Matriz.txt','w')
base = pd.read_excel('Disciplinas.xlsx')
base_curso = pd.read_excel('Curso.xlsx')
base_matriz = pd.read_excel('Matriz.xlsx')
hoje = datetime.date.today().strftime("%d/%m/%Y")
base_matriz['Carga'] = base_matriz['Carga'].astype(object)

def cadastro_disciplina():
    try:
        driver.find_element(By.XPATH, '//*[@id="search"]/div/ul/li/input').send_keys(disciplinas + Keys.ENTER)
        time.sleep(3)
        driver.find_element(By.XPATH, '/html/body/main/div/div/div/div[2]/div/table/tbody').text
    except:
        driver.find_element(By.XPATH, '/html/body/main/div/div/div/div[1]/div/div[2]/a/i').click()
        time.sleep(3)
        driver.find_element(By.XPATH, '//*[@id="nome"]').send_keys(disciplinas)
        driver.find_element(By.XPATH, '//input[@id="tipoDisciplina-selectized"]').send_keys(tipo_disciplina)
        driver.find_element(By.XPATH, '/html/body/main/div/div/div/form/div[1]/div/div[2]/div/div/div[2]/div/div[1]').click()
        driver.find_element(By.XPATH, '//input[@id="ativo-selectized"]').send_keys(ativa)
        driver.find_element(By.XPATH, '/html/body/main/div/div/div/form/div[1]/div/div[3]/div/div[2]/div/div[1]').click()
        time.sleep(3)
        driver.get('https://umc.stage.platosedu.io/erp/disciplina/index')
        
        print(disciplinas, '- foi cadastrada em', hoje, file=arquivo)
    else:
        time.sleep(3)
        driver.get('https://umc.stage.platosedu.io/erp/disciplina/index')
        print(disciplinas, '- já está cadastrada', file=arquivo)

def cadastro_curso():
    try: 
        driver.get('https://umc.stage.platosedu.io/erp/curso/index')
#Verificar se o curso já está cadastrado
        driver.find_element(By.XPATH, '//*[@id="search"]/div/ul/li/input').send_keys(curso + Keys.ENTER)
        time.sleep(3)
        driver.find_element(By.XPATH, '/html/body/main/div/div/div/div[2]/div/table/tbody').text
    except:
#Tipo do Curso
        driver.find_element(By.XPATH, '/html/body/main/div/div/div/div[1]/div/div[2]/a').click()
        driver.find_element(By.XPATH, '//input[@id="tipoCurso-selectized"]').send_keys(tipo_curso)
#Nome do Curso
        driver.find_element(By.XPATH, '/html/body/main/div/div/form/div/div[1]/div/div[1]/div/div/div[2]/div/div[1]').click()
        driver.find_element(By.XPATH, '//*[@id="nome"]').send_keys(curso)        
#Ativo?
        driver.find_element(By.XPATH, '//input[@id="ativo-selectized"]').send_keys(ativa)
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
        time.sleep(3)
#Atualizar a página
        driver.get('https://umc.stage.platosedu.io/erp/curso/index')
        
        print(curso, '- foi cadastrada em', hoje, file=arquivo_curso)
    else:
        time.sleep(3)
        driver.get('https://umc.stage.platosedu.io/erp/curso/index')
        print(curso, '- já está cadastrada', file=arquivo_curso)

def cadastro_matriz():
    try:
        driver.find_element(By.XPATH, '//*[@id="search"]/div/ul/li/input').send_keys(curso + Keys.ENTER)
        time.sleep(3)
        driver.find_element(By.XPATH, '/html/body/main/div/div/div/div[2]/div/table/tbody').text
    except:
        print(curso, "- Não cadastrado", hoje, file=arquivo_matriz)
    else:
        driver.find_element(By.XPATH, '/html/body/main/div/div/div/div[2]/div/table/tbody/tr[1]/td[5]/div/button/i').click()
        driver.find_element(By.XPATH, '/html/body/main/div/div/div/div[2]/div/table/tbody/tr[1]/td[5]/div/div/a[4]').click()
        time.sleep(3)
        validacao = driver.find_element(By.XPATH, '/html/body/main/div/div/div/div[2]/div/table/tbody/tr/td[2]').text
        if validacao == matriz:
            print(matriz, '- já cadastrada', hoje, file=arquivo_matriz)
        else:
            driver.find_element(By.XPATH,'/html/body/main/div/div/div/div[1]/div/div[2]/a').click()
            time.sleep(3)
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
            time.sleep(3)

driver = webdriver.Chrome()
driver.maximize_window()
driver.get('https://umc.stage.platosedu.io/erp/disciplina/index')
driver.find_element(By.ID, 'usuario').send_keys("bruno.pedroso")
driver.find_element(By.ID, 'password').send_keys("tubets@123" + Keys.ENTER)
time.sleep(3)

for i, disciplinas in enumerate(base['Disciplina']):
    tipo_disciplina = base.loc[i, 'Tipo']
    ativa = base.loc[i, 'Ativa']
    insti = base.loc[i, 'Instituição']
    cadastro_disciplina()

for i, curso in enumerate(base_curso['Curso']):
    tipo_curso = base_curso.loc[i, 'Tipo']
    ativa = base_curso.loc[i, 'Ativa']
    conj = base_curso.loc[i, 'Conjunção']
    titu = base_curso.loc[i, 'Titulação']
    web = base_curso.loc[i, 'WEB']
    conhec = base_curso.loc[i, 'Conhecimento']
    cadastro_curso()
    
for i, curso in enumerate(base_matriz['Curso']):
    tipo_matriz = base_matriz.loc[i, 'Tipo Matriz']
    matriz = base_matriz.loc[i,'Nome']
    ch = base_matriz.loc[i, 'Carga']
    resolucao = base_matriz.loc[i,'Resolução']
    per1 = base_matriz.loc[i,'Perc 1']
    per2 = base_matriz.loc[i,'Perc 2']
    per3 = base_matriz.loc[i,'Perc 3']
    apost = base_matriz.loc[i,'Apostilamento']
    desc = base_matriz.loc[i,'Descricão']
    resumo = base_matriz.loc[i,'Resumo']
    introd = base_matriz.loc[i,'Introdução']
    cert = base_matriz.loc[i,'Certificação']
    obje = base_matriz.loc[i,'Objetivo']
    metodologia = base_matriz.loc[i,'Metodologia']
    publico = base_matriz.loc[i,'Publico']
    ementa = base_matriz.loc[i,'Ementa']
    periodicidade = base_matriz.loc[i,'Periodicidade']
    cadastro_matriz()

driver.quit()