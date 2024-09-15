from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time, os, pandas as pd

# Integração à planilha
studentsSpreadsheet = pd.read_excel(r".gitignore/primeiro.xlsx", header = 0)
registrationUser    = studentsSpreadsheet['Matricula'].tolist()
nameUser            = studentsSpreadsheet['Nomes'].tolist()
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


driver.quit()