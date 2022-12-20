import pyautogui
import time
from mod.BD_update_PBI import BD_update_PBI
from mod.SmartMail import SmartMail

""" 
bd_evolucao

Script que gera relatórios e atualiza o arquivo BD_Evolucao, utilizado no dashboard de Caderno Gerencial

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

url = "https://dsadasd.com/dsadsa" # ---> tempo por colaborador
BD_update_PBI.getLink(driver, url)
initDay, endDay = BD_update_PBI.startEndPreviousMonth() #returns two days: the start and the end of the previous month
BD_update_PBI.inputDates(driver, initDay, endDay)
BD_update_PBI.downloadRelFluid(driver) 
time.sleep(30)

url = "https://dsadasd.com/dsadsa" # ---> tipo de retorno por colaborador
BD_update_PBI.getLink(driver, url)
BD_update_PBI.inputDates(driver, initDay, endDay, alt=True)
BD_update_PBI.downloadRelFluid(driver, noAdd = True) 
time.sleep(30)

print("Finishing downloads...")
BD_update_PBI.endSession(driver)

# pathTo = "C:/PATH/foiead.xlsx" #DEV
pathTo = "C:/PATH/dsadas.xlsx" #ROBOT

#-------------Copying data from Tempo por Colaborador to BD_Evolucao-------------------

print("Copying data from Tempo por Colaborador to BD_Evolucao...")
pattern = r"^0_Analidsadasqlucao_.*.xls" 
pathXLS = BD_update_PBI.getFilePath(pattern) 
pathFrom = BD_update_PBI.saveCorruptedXLSAsXLSX(pathXLS)
fluidSheet = "5- Analise tempos - por colabor"
BD_update_PBI.copyFromTo(pathFrom, pathTo, fluidSheet, sys)

print("Dados copiados - Tempo por Colaborador")
time.sleep(15)
#-------------Copying data from QuantTipoColab to BD_Evolucao-------------------

print("Copying data from QuantTipoColab to BD_Evolucao...")
# pathXLS = "C:/PATH/dsadas.xls" #---> this one is always the same DEV DEV DEV
pathXLS = "C:/PATH/dsadas.xls" #---> ROBOT ROBOT ROBOT
pathFrom = BD_update_PBI.saveCorruptedXLSAsXLSX(pathXLS)
fluidSheet = "Devoluções - FLUID"
BD_update_PBI.copyFromTo(pathFrom, pathTo, fluidSheet, sys)

print("Dados copiados - Tipo de Retorno por Colaborador")
time.sleep(15)
# -------------Sending email--------------------

print('Enviando email...')
mailTo = "dsadasda@coqeqderadadadi.oneqweqft.com"
subject = "Smart - BD_Evolucao atualizado"
body = "Bom dia!" + "\n" + "O arquivo BD_Evolucao foi atualizado."
SmartMail.sendMail(mailTo, subject, body) #sends email using Outlook 
  
print('Finalizando o script.')