from mod.SYS1 import SYS1
from mod.SYS4 import SYS4
from mod.SYS2 import SYS2
from mod.SmartMail import SmartMail
import pyautogui
import sys

""" 
import_cc

Script que gera relatórios utilizando os sistemas X, Y, Z e os importa
na página de importações W, enviando um email ao final do script.

Author: Juliano Kosloski - Automation Developer
Created: 12/07/2022 by Juliano Kosloski
"""

pyautogui.PAUSE = 2.0

print('Iniciando o script...')
#SYS1
sys_rep_list = {"AAA": ["139", "083"], "BBB": ["061", "068"]}

rep_counter = 0

for x in sys_rep_list:
    
    system = x
    
    for y in sys_rep_list[x]:
        
        report = y
        
        if report == "083":
            
            rep_timeout = 480
            
        else:
            
            rep_timeout = 300
            
        SYS1.startClient()
        print('Abrindo X...')
        SYS1.getSystem(system) 
        print('Buscando credenciais...')   
        SYS1.getLogin()
        print('Buscando relatório {}...'.format(report))
        SYS1.getReport(report)
        
        print('Checando o download...')
        checkReport = SYS1.checkReport(report, rep_timeout) #checks if the file was downloaded
        
        if checkReport == True:
            
            print('Download realizado com sucesso')
            rep_counter += 1
            if rep_counter == 4:
                print("Abrindo o SYS4...")
            
        elif checkReport == False:  
             
            print('Não foi possível baixar o arquivo {}'.format(report))
            pyautogui.alert("Erro ao baixar o relatório " + report + ". Encerrando o programa") 
            raise pyautogui.FailSafeException #end the script 
        
#---sys4

driver = SYS4.startDriver()
SYS4.getLogin(driver)
SYS4.downloadReports(driver, timeout = 15)

check = SYS4.checkReports(driver)
if check == False:
    
    print('Não foi possível baixar o relatório do SYS4...')
    raise pyautogui.FailSafeException

#SYS2
print('Iniciando importações...')
driver = SYS2.startDriver()
SYS2.getLogin(driver)
print('Importando arquivos...')
SYS2.uploadFiles(driver)
SYS2.endSession(driver)

#Email

#sets up mail variables
print('Enviando email...')
mailTo = "deassascgrea@email.com"
subject = "bom dia"
body = "importações"

SmartMail.sendMail(mailTo, subject, body) #sends email using Outlook
print('Encerrando o programa.')
    
