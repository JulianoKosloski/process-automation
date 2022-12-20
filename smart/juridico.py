from mod.HGV import HGV
from mod.SmartMail import SmartMail
import time
import pyautogui


####
#### UNFINISHED
####

""" 
juridico

Realiza o download de quatro relatórios que chegam por email (dois na noite anterior, dois de manhã) e os importa na HGV.

Author: Juliano Kosloski - Automation Developer
Created: 01/12/2022 by Juliano Kosloski
"""

pyautogui.PAUSE = 1.0
pyautogui.hotkey('win', 'd') #go to desktop

#Get email attachments

#HGV
print('Iniciando importações...')
driver = HGV.startDriver()
HGV.getLogin(driver)
print('Importando arquivos...')
# HGV.uploadFiles(driver)  ---> update this method so 081 is updated along with four files
HGV.endSession(driver)