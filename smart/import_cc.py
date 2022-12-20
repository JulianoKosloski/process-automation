from mod.ACClient import ACClient
from mod.OpenData import OpenData
from mod.HGV import HGV
from mod.SmartMail import SmartMail
import pyautogui

""" 
import_cc

Script que gera relatórios utilizando os sistemas SIAT, SACG, OpenData e os importa
na página de importações da HGV, enviando um email ao final do script.

Author: Juliano Kosloski - Automation Developer
Created: 12/07/2022 by Juliano Kosloski
"""

pyautogui.PAUSE = 2.0

print('Iniciando o script...')
#ACClient
sys_rep_list = {"SACG": ["139", "083"], "SIAT": ["061", "068"]}

rep_counter = 0

for x in sys_rep_list:
    
    system = x
    
    for y in sys_rep_list[x]:
        
        report = y
        
        if report == "083":
            
            rep_timeout = 480
            
        else:
            
            rep_timeout = 300
            
        ACClient.startClient()
        print('Abrindo ACClient...')
        ACClient.getSystem(system) 
        print('Buscando credenciais...')   
        ACClient.getLogin()
        print('Buscando relatório {}...'.format(report))
        ACClient.getReport(report)
        
        print('Checando o download...')
        checkReport = ACClient.checkReport(report, rep_timeout) #checks if the file was downloaded
        
        if checkReport == True:
            
            print('Download realizado com sucesso')
            rep_counter += 1
            
        elif checkReport == False:  
             
            print('Não foi possível baixar o arquivo {}'.format(report))
            pyautogui.alert("Erro ao baixar o relatório " + report + ". Encerrando o programa") 
            raise pyautogui.FailSafeException #end the script 
        
#OpenData

driver = OpenData.startDriver()
OpenData.getLogin(driver)
OpenData.downloadReports(driver, timeout = 45)

check = OpenData.checkReports(driver)
if check == False:
    
    print('Não foi possível baixar o relatório do OpenData...')
    raise pyautogui.FailSafeException

#HGV
print('Iniciando importações...')
driver = HGV.startDriver()
HGV.getLogin(driver)
print('Importando arquivos...')
HGV.uploadFiles(driver)
HGV.endSession(driver)

#Email

#sets up mail variables
print('Enviando email...')
mailTo = "ema@email.com"
subject = "Smart - Importações HGV realizadas"
body = "Olá! As importações dos relatórios ?????" + " foram realizadas com sucesso."

SmartMail.sendMail(mailTo, subject, body) #sends email using Outlook
print('Encerrando o programa.')
    
