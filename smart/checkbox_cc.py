from mod.HGV import HGV
from mod.SmartMail import SmartMail
import time
import pyautogui

""" 
checkbox_cc

Atualiza três relatórios na plataforma HGV, enviando um email ao final do script.

Author: Juliano Kosloski - Automation Developer
Created: 01/12/2022 by Juliano Kosloski
"""

pyautogui.PAUSE = 1.0
pyautogui.hotkey('win', 'd') #go to desktop

driver = HGV.startDriver()
HGV.getLogin(driver)
url = "https://url.com/dasdas.php"
HGV.getLink(driver, url)
pesq = HGV.checkbox(driver)
time.sleep(30)
HGV.endSession(driver)

# -------------Sending email--------------------

print('Enviando email...')
mailTo = "eaadas@ceaadaads.com"
subject = "Importações - Denodo/HGV"
if (pesq == True):
    body = "Bom dia!" + "\n" + "Os relatórios ????? foram importados na HGV."
else:
    body = "Bom dia!" + "\n" + "Os relatórios ???? foram importados na HGV."
SmartMail.sendMail(mailTo, subject, body) #sends email using Outlook 
  
print('Finalizando o script.')