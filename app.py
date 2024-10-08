from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import datetime, sys, time, os, threading, pandas as pd

spreadName = input('Digite o nome da planilha:\n')
spreadName.title()
yearRegister = input('Insira o ano de ingresso dos alunos:\n')
if str(yearRegister) == ' ':
    sys.exit("Informação inserida é invalida.\nNada foi informado")
elif str(yearRegister) == 'nan':
    sys.exit("Informação inserida é invalida.\nInformação discrepante, não contém números.")
elif len(str(yearRegister)) != 4:
    sys.exit("Informação inserida é invalida.\nAno com menos de quatro dígitos.")
elif yearRegister > datetime.now().year:
    sys.exit("Informação inserida é invalida.\nAno inválido, não há como inserir anos futuros.")

cooldown = 0.5

# Integração à planilha
studentsSpreadsheet = pd.read_excel(rf"SIGE-CE_To_BibLivre5\archive\{spreadName}.xlsx", header = 0)
registrationUser    = studentsSpreadsheet['Matricula'].tolist()
nameUser            = studentsSpreadsheet['Nomes'].tolist()
cpfUser             = studentsSpreadsheet['CPF'].tolist()
emailUser           = studentsSpreadsheet['Email'].tolist()
genderUser          = studentsSpreadsheet['Gênero'].tolist()
# O "x".tolist() retorna uma lista float (???)

# Filtro de informações

## Iniciando listas pra armazenar as informações filtradas
registrationUsers = [ ]
nameUsers         = [ ]
cpfUsers          = [ ]
emailUsers        = [ ]
genderUsers       = [ ]
contShelf = 0
for i in registrationUser:
    if str(registrationUser[contShelf]) != 'nan':
        registrationStudent = int(registrationUser[contShelf])
        registrationUsers.append(str(registrationStudent))
    if str(nameUser[contShelf]).startswith('ALUNO: ') == True:
        nameStudent = str(nameUser[contShelf]).title()
        nameUsers.append(nameStudent[7:])
    if str(cpfUser[contShelf]).startswith('CPF: ') == True:
        cpfStudent = str(cpfUser[contShelf])
        cpfUsers.append(cpfStudent[5:])
    if str(emailUser[contShelf]).startswith('EMAIL: ') == True:
        emailStudent = str(emailUser[contShelf])
        emailUsers.append(emailStudent[7:])
    if str(genderUser[contShelf]) != 'nan':
        genderStudent = str(genderUser[contShelf]).upper()
        genderUsers.append(genderStudent)

    contShelf += 1

# Geração de log
def logGenerator():
    userChoose = ''
    while userChoose.upper() != "Y" and userChoose.upper() != "N":
        userChoose = input('Criar log dos alunos carregados? Y ou N\n')
        
        if userChoose.upper() == 'Y':
            informations = [ ]
            contInfo = 0
            for i in registrationUsers:
                informations.append(f'{contInfo+1}, MATRICULA: {registrationUsers[contInfo]}, NOME: {nameUsers[contInfo]}, CPF: {cpfUsers[contInfo]}, EMAIL: {emailUsers[contInfo]}, GENERO: {genderUsers[contInfo]}')
                contInfo += 1
            with open(r"SIGE-CE_To_BibLivre5\archive\log.txt", "w") as file:
                for i in informations:
                    file.write(f'{i}\n')
                    print(i)
                    time.sleep(0.2)
        elif userChoose.upper() == 'N':
            print('Log não será gerado\nMotivo: Solicitação negada')
        else:
            print('Informação inserida é inválida')
def userRegister():
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
    contRegister = 0
    for i in registrationUsers:
        driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[1]/div[4]/a').click()
        time.sleep(cooldown)
        driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[1]/div[1]/div/div[1]/div[2]/input').send_keys(f'{nameUsers[contRegister]}')
        driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[1]/div[1]/div/div[3]/div[2]/select').find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[1]/div[1]/div/div[3]/div[2]/select/option[2]').click()
        time.sleep(cooldown-0.4)
        driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[2]/div/div/fieldset/div/div[1]/div[2]/input').send_keys(f'{emailUsers[contRegister]}')
        if str(genderUsers[contRegister]).upper() == 'F':
            driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[2]/div/div/fieldset/div/div[2]/div[2]/select').find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[2]/div/div/fieldset/div/div[2]/div[2]/select/option[3]').click()
            time.sleep(cooldown-0.4)
            driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[2]/div/div/fieldset/div/div[16]/div[2]/textarea').send_keys(f'Aluna de matrícula: {registrationUsers[contRegister]}, do ano de {yearRegister}.')
            time.sleep(cooldown-0.4)
        elif str(genderUsers[contRegister]).upper() == 'M':
            driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[2]/div/div/fieldset/div/div[2]/div[2]/select').find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[2]/div/div/fieldset/div/div[2]/div[2]/select/option[2]').click()
            time.sleep(cooldown-0.4)
            driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[2]/div/div/fieldset/div/div[16]/div[2]/textarea').send_keys(f'Aluno de matrícula: {registrationUsers[contRegister]}, do ano de {yearRegister}.')
            time.sleep(cooldown-0.4)
        driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[2]/div/div/fieldset/div/div[8]/div[2]/input').send_keys(f'{cpfUsers[contRegister]}')
        time.sleep(cooldown-0.4)
        driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[4]/div[3]/div[2]/a[1]').click()
        time.sleep(cooldown)
        driver.find_element(By.XPATH, '/html/body/form/div[3]/div[1]/div[2]/div[3]/div[2]/a').click()
        time.sleep(cooldown)

        contRegister += 1
    driver.quit()

# Controle de concorrência
firstThread  = threading.Thread(target=logGenerator)
secondThread = threading.Thread(target=userRegister)

firstThread.start()
secondThread.start()