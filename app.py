from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time, os, pandas as pd

# Integração à planilha
studentsSpreadsheet = pd.read_excel(r"sistemaBiblioteca-EEEP-AN\BibLivre5\2D.xlsx", header = 0)
registrationUser    = studentsSpreadsheet['Matricula'].tolist()
nameUser            = studentsSpreadsheet['Nome'].tolist()
cpfUser             = studentsSpreadsheet['CPF'].tolist()
emailUser           = studentsSpreadsheet['Email'].tolist()
genderUser          = studentsSpreadsheet['Genero'].tolist()
# O "x".tolist() retorna uma lista float (???)

yearRegister = input('Insira o ano de registro:\n')

options = webdriver.EdgeOptions()
options.add_experimental_option("detach", True)
driver = webdriver.Edge(options=options)

#Abre o BibLivre e faz Login
driver.get('http://localhost/Biblivre5/')
userNameLogin = driver.find_element(By.XPATH, '/html/body/form/div[1]/div[6]/ul/li[5]/input[1]')
passWordLogin = driver.find_element(By.XPATH, '/html/body/form/div[1]/div[6]/ul/li[5]/input[2]')
userNameLogin.clear()
passWordLogin.clear()
userNameLogin.send_keys('admin')
passWordLogin.send_keys('abracadabra')
driver.find_element(By.XPATH, '/html/body/form/div[1]/div[6]/ul/li[4]/button').click()

#Abre o Cadastro de Usuários
menu_trigger = driver.find_element(By.XPATH, '/html/body/form/div[1]/div[6]/ul/li[2]')
cursor = ActionChains(driver)
cursor.move_to_element(menu_trigger).perform()
popup_option = driver.find_element(By.XPATH, '/html/body/form/div[1]/div[6]/ul/li[2]/ul/li[1]').click()

#Registro de novo usuário
cont = 0
for i in nameUser:
    driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[1]/div[4]/a').click()
    driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[1]/div[1]/div/div[1]/div[2]/input').send_keys(f'{nameUser[cont]}')
    driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[1]/div[1]/div/div[3]/div[2]/select').find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[1]/div[1]/div/div[3]/div[2]/select/option[2]').click()
    driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[2]/div/div/fieldset/div/div[1]/div[2]/input').send_keys(f'{emailUser[cont]}')
    if genderUser[cont] == 'f':
        driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[2]/div/div/fieldset/div/div[2]/div[2]/select').find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[2]/div/div/fieldset/div/div[2]/div[2]/select/option[3]').click()
    elif genderUser[cont] == 'm':
        driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[2]/div/div/fieldset/div/div[2]/div[2]/select').find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[2]/div/div/fieldset/div/div[2]/div[2]/select/option[2]').click()
    driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[2]/div/div/fieldset/div/div[8]/div[2]/input').send_keys(f'{cpfUser[cont]}')
    if genderUser[cont] == 'f':
        driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[2]/div/div/fieldset/div/div[16]/div[2]/textarea').send_keys(f'Aluna de matrícula: {registrationUser[cont]}, do ano de {yearRegister}.')
    elif genderUser[cont] == 'm':
        driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[2]/div/div/fieldset/div/div[16]/div[2]/textarea').send_keys(f'Aluno de matrícula: {registrationUser[cont]}, do ano de {yearRegister}.')
    driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[3]/div[2]/a[1]').click()
    driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[2]/a').click()
    cont += 1

driver.quit()