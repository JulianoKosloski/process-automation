import pyautogui
import time
from mod.BD_update_PBI import BD_update_PBI
from mod.SmartMail import SmartMail

""" 
bd_seguros_ag

Script que gera relatórios e atualiza o arquivo BD_SEGUROS_REPORT_AG, utilizado no dashboard de Seguros

Author: Juliano Kosloski - Automation Developer
Created: 23/11/2022 by Juliano Kosloski
"""

pyautogui.PAUSE = 1.0
pyautogui.hotkey('win', 'd') #go to desktop

#-------------Fluid - Downloading files-------------------

print("Initiating download...")

sys = "Fluid"
driver = BD_update_PBI.startDriver()
BD_update_PBI.getLogin(driver, sys)


# -------------Sending email--------------------

print('Enviando email...')
mailTo = "dsadas@cdasdasderdsadas.com"
subject = "Smart - BD_Seguros_AG atualizado"
body = "Bom dia!" + "\n" + "O arquivo BD_Seguros_AG foi atualizado."
SmartMail.sendMail(mailTo, subject, body) #sends email using Outlook 
  
print('Finalizando o script.')