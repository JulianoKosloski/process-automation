import pyautogui
import time
from mod.BD_update_PBI import BD_update_PBI
from mod.SmartMail import SmartMail

""" 
rel_oferta_basica

Script que gera relatórios e atualiza o arquivo BD_ASSO, utilizado no dashboard de Oferta Básica

Author: Juliano Kosloski - Automation Developer
Created: 30/09/2022 by Juliano Kosloski
"""

pyautogui.PAUSE = 1.0

pyautogui.hotkey('win', 'd') #go to desktop

#-------------Fluid - Downloading files-------------------

print("Initiating download...")

sys = "Fluid"

driver = BD_update_PBI.startDriver()
BD_update_PBI.getLogin(driver, sys)
BD_update_PBI.getLink(driver, "https://url.com", 1)
initDay = BD_update_PBI._beforeDate(7)
endDay = BD_update_PBI._beforeDate(1)
BD_update_PBI.inputDates(driver, initDay, endDay)
time.sleep(5)
BD_update_PBI.downloadRelFluid(driver)
time.sleep(30)

sys = "AdmCanais"
BD_update_PBI.getLogin(driver, sys)
BD_update_PBI.getLink(driver, "https://url.com", 1)
BD_update_PBI.downloadRelAdmCanais(driver) #has its own input date logic inside the method (not the best practice and I know it)
time.sleep(20)

print("Finishing downloads...")
BD_update_PBI.endSession(driver)

# pathTo = "C:/PATH/file.xlsx"  #----DEV
pathTo = "C:/PATH/file.xlsx"

#-------------Copying Fluid data to BD_oferta_basica-------------------

pattern = r"^dsadada_.*.xls" 
pathXLS = BD_update_PBI.getFilePath(pattern)
pathFrom = BD_update_PBI.saveCorruptedXLSAsXLSX(pathXLS)
fluidSheet = "FLUID - ABERTURA CC"
BD_update_PBI.copyFromTo(pathFrom, pathTo, fluidSheet, "Fluid")

time.sleep(15)

#-------------Copying AdmCanais data to BD_oferta_basica-------------------

# pathXLS = "C:/PATH/file.xls" #dev machine
pathXLS = "C:/PATH/file.xls"
pathFrom = BD_update_PBI.saveXLSAsXLSX(pathXLS)
admSheet = "Canais"
BD_update_PBI.copyFromTo(pathFrom, pathTo, admSheet, "AdmCanais")

time.sleep(15)

# -------------Sending email--------------------

print('Enviando email...')
mailTo = "email@email.com"
subject = "Smart - BD_ASSO atualizado"
body = "Bom dia!" + "\n" + "O arquivo BD_ASSO foi atualizado."
SmartMail.sendMail(mailTo, subject, body) #sends email using Outlook 
  
print('Finalizando o script.')