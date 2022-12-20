import pyautogui
import time
from mod.BD_update_PBI import BD_update_PBI
from mod.SmartMail import SmartMail

""" 
bd_caderno

Script que gera relatórios e atualiza o arquivo BD_Caderno Gerencial, utilizado no dashboard de Caderno Gerencial

Author: Juliano Kosloski - Automation Developer
Created: 11/11/2022 by Juliano Kosloski
"""

pyautogui.PAUSE = 1.0
pyautogui.hotkey('win', 'd') #go to desktop

#-------------Fluid - Downloading files-------------------

print("Initiating download...")

sys = "Fluid"
driver = BD_update_PBI.startDriver()
BD_update_PBI.getLogin(driver, sys)

url = "https://dsadas.com.br/26" #  ->>>> atualização cadastral
BD_update_PBI.getLink(driver, url)
initDay, endDay = BD_update_PBI.startEndPreviousMonth() #returns two days: the start and the end of the previous month
BD_update_PBI.inputDates(driver, initDay, endDay)
BD_update_PBI.downloadRelFluid(driver) 
time.sleep(30)

url = "https://dsadas.com.br/26" #  ->>>> encerramento de contas
BD_update_PBI.getLink(driver, url)
initDay, endDay = BD_update_PBI.startEndPreviousMonth() #returns two days: the start and the end of the previous month
BD_update_PBI.inputDates(driver, initDay, endDay)
BD_update_PBI.downloadRelFluid(driver) 
time.sleep(30)

url = "https://dsadas.com.br/2dsda6" #  ->>>> abertura de contas
BD_update_PBI.getLink(driver, url)
initDay, endDay = BD_update_PBI.startEndPreviousMonth() #returns two days: the start and the end of the previous month
BD_update_PBI.inputDates(driver, initDay, endDay)
BD_update_PBI.downloadRelFluid(driver) 
time.sleep(30)
    
print("Finishing downloads...")
BD_update_PBI.endSession(driver) 

# pathTo = "C:/PATH/cxca.xlsx" #DEV
pathTo = "C:/PATH/dasdas.xlsx" # --> ROBOT

#-------------Copying data from Atualização Cadastral to BD_Caderno Gerencial-------------------

print("Copying data from Atualização Cadastral to BD_Caderno Gerencial")
pattern = r"^0_Atualdasad_ewqJ_.*.xls"
pathXLS = BD_update_PBI.getFilePath(pattern) 
pathFrom = BD_update_PBI.saveCorruptedXLSAsXLSX(pathXLS)
fluidSheet = "ATUALI_CADASTRAL"
BD_update_PBI.copyFromTo(pathFrom, pathTo, fluidSheet, sys)

print("Dados copiados - Atualização Cadastral")
time.sleep(15)
#-------------Copying data from Encerramento de Contas to BD_Caderno Gerencial-------------------

print("Copying data from Encerramento de Contas to BD_Caderno Gerencial")
pattern = r"^0_Enceqwe_Contewqeqw.*.xls"
pathXLS = BD_update_PBI.getFilePath(pattern) 
pathFrom = BD_update_PBI.saveCorruptedXLSAsXLSX(pathXLS)
fluidSheet = "ENCERRAMENTO"
BD_update_PBI.copyFromTo(pathFrom, pathTo, fluidSheet, sys)

print("Dados copiados - Encerramento de Contas")
time.sleep(15)
#-------------Copying data from Abertura de Contas to BD_Caderno Gerencial-------------------

print("Copying data from Abertura de Contas to BD_Caderno Gerencial")
pattern = r"^0_PBewqeura_de_CoeqweqwqJ_.*.xls" 
pathXLS = BD_update_PBI.getFilePath(pattern) 
pathFrom = BD_update_PBI.saveCorruptedXLSAsXLSX(pathXLS)
fluidSheet = "AB_CONTAS"
BD_update_PBI.copyFromTo(pathFrom, pathTo, fluidSheet, sys)

print("Dados copiados - Abertura de Contas")
time.sleep(15)
# -------------Sending email--------------------

print('Enviando email...')
mailTo = "adasdqwe@confedeqweqewqeqeqeq.com"
subject = "Smart - BD_Caderno Gerencial atualizado"
body = "Bom dia!" + "\n" + "O arquivo BD_Caderno Gerencial foi atualizado."
SmartMail.sendMail(mailTo, subject, body) #sends email using Outlook 
  
print('Finalizando o script.')