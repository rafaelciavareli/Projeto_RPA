### PROJETO DESENVOLVIDO POR: RAFAEL RODRIGUES CIAVARELI ###
#### RPA CHALLENGE - TOOLS DIGITAL SERVICE - PIRACICABA ####
################# CRIADO EM 14/07/2023 #####################

# Bibliotecas de tempo
import time
import datetime 
from datetime import datetime

# Bibliotecas Tkinter para exibir a janela de upload do EXCEL e Aviso quando finalizar 
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

#Biblioteca PANDAS para leitura do arquivo EXCEL
import pandas as pd

#Bibliotecas SELENIUM para processamento dos registros na pagina WEB 
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

#Biblioteca para encerrar o Python
import sys

#Biblioteca para baixar o arquivo via URL
import requests
###################################################################################

# Iniciando o Programa #

#Configurando o Driver para abrir o Navegador [Google Chrome)
options = webdriver.ChromeOptions()
options.add_argument('--start-minimezed')
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 5)


#Definindo a URL que será aberta
url = "https://www.rpachallenge.com/"
driver.get(url)


#Baixa o arquivo automaticamente para o usuário importar. 
acessar_arquivo = requests.get("https://www.rpachallenge.com/assets/downloadFiles/challenge.xlsx")

#Verifico se a Página esta disponivel para download usando o retorno da requisição 
if acessar_arquivo.status_code == 200:
    #Pego o conteudo da pagina e armazendo em uma variavel 
    arquivo_imported = acessar_arquivo.content
    #Leio o conteudo da pagina salvo na variavel usando PANDAS
    arquivo_process = pd.read_excel(arquivo_imported)

#Inciando o Processamento dos Dados atráves dos dados recebidos via Requests.
time.sleep(0.1) #Aguarda a pagina carregar antes de clicar 
start_botton = driver.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button').click() #Clica no Botão Start para iniciar o teste

#Função para clicar e inserir os valores nos elementos da página.
def processar_campos(element, value):
    element.click()
    element.send_keys(value)

for index, row in arquivo_process.iterrows():

    date_now = datetime.now().strftime("%d/%m/%Y %H:%M:%S") #Defino a data atual em que os dados sera processado. 
    time_wait = 0.1 #Defino o tempo em que cada elemento será preenchido
    
    #Processando as ROW recebidas do "arquivo_process"
    time.sleep(time_wait)
    first_name_element = driver.find_element(By.CSS_SELECTOR, f"[ng-reflect-name='labelFirstName']") #Captura do CSS do Angular usado na página RPA CHALLENGER
    processar_campos(first_name_element,row["First Name"])

    time.sleep(time_wait)
    last_name_element = driver.find_element(By.CSS_SELECTOR,"[ng-reflect-name='labelLastName']")
    processar_campos(last_name_element,row["Last Name "])
    
    time.sleep(time_wait)
    company_name_element = driver.find_element(By.CSS_SELECTOR,"[ng-reflect-name='labelCompanyName']")
    processar_campos(company_name_element,row["Company Name"])
    
    time.sleep(time_wait)
    role_in_campany_element = driver.find_element(By.CSS_SELECTOR,"[ng-reflect-name='labelRole']")
    processar_campos(role_in_campany_element,row["Role in Company"])
    
    time.sleep(time_wait)
    address_element = driver.find_element(By.CSS_SELECTOR,"[ng-reflect-name='labelAddress']")
    processar_campos(address_element,row["Address"])
    
    time.sleep(time_wait)
    email_element = driver.find_element(By.CSS_SELECTOR,"[ng-reflect-name='labelEmail']")
    processar_campos(email_element,row["Email"])
    
    time.sleep(time_wait)
    phone_number_element = driver.find_element(By.CSS_SELECTOR,"[ng-reflect-name='labelPhone']")
    processar_campos(phone_number_element,row["Phone Number"])

    #Clicando no botão 'SUBMIT'
    time.sleep(time_wait)
    submit_botton = driver.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input').click()
    

    #Exibir a mensagem de sucesso a cada Inserção e mostra no Log a quantia inserida e o nome e sobrenome do usuário.
    first_name = row["First Name"]
    last_name = row["Last Name "]
    print("**************************************************************************")
    print(f"Usuário: {first_name} {last_name}, Registrado com sucesso em: {date_now}")
    
return_success = driver.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[2]/div[2]')
return_success_message = return_success.text
print("**************************************************************************")
print(return_success.text)
messagebox.showinfo("Congratulations!", f"{return_success_message}")

driver.quit()
sys.exit()