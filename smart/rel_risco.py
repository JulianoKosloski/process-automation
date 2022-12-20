import pyautogui
import time
from mod.HGV import HGV
from mod.Risco import Risco
from mod.SmartMail import SmartMail

""" 
rel_risco

Script que gera relatórios de risco utilizando o sistema HGV, processa os arquivos e envia emails
para colaboradores nas agências e no Centro Administrativo.

Author: Juliano Kosloski - Automation Developer
Created: 08/08/2022 by Juliano Kosloski
"""

pyautogui.PAUSE = 1.0

pyautogui.hotkey('win', 'd') #go to desktop

#HGV - Downloading files 

print('Baixando arquivos Risco...')
driver = HGV.startDriver()
urlRisco = "https://url.com"
xpathList = [r'//*[@id="btnExportar1"]',r'//*[@id="btnExportar"]'] #lists the buttoms 'Exportar' and 'Exp. Detalhes'

HGV.getLogin(driver)
HGV.getLink(driver, urlRisco, attempts = 2) #needs to go to the url twice to get the right page
HGV.downloadFiles(driver, xpathList)
HGV.endSession(driver)

#Risco - manipulating .csv files

print('Buscando arquivos na pasta...')
arqPath, detPath, detFileName = Risco.getFiles() #gets file paths

#save files as .xlsx

print('Transformando arquivos em .xlsx...')
arqPath = Risco.openFile(arqPath) 
arqNewPath = Risco.saveAsXLSX(arqPath)
Risco.closeExcel()

detPath = Risco.openFile(detPath) 
detNewPath = Risco.saveAsXLSX(detPath)
Risco.closeExcel()

mainPath = "F:/PATH/file.xlsx" #changing to server path might solve some issues with task scheduler

#prepares files and copies them to the main file

print('Preparando arquivos...')
arqPath, detPath = Risco.prepFiles(arqNewPath, detNewPath)
Risco.clearMainFile(mainPath)
print('Copiando dados para o arquivo principal...')
mainPath = Risco.appendMainFile(arqPath, detPath, mainPath) 

#opens main file for computing differences by agency and reason

print('Atualizando tabelas...')
Risco.updateTables(mainPath)
Risco.diffAgencia(mainPath)
time.sleep(15);
Risco.updateTables(mainPath)
Risco.diffMotivo(mainPath)
time.sleep(15);
Risco.updateTables(mainPath)

#saves files to different folders
Risco.saveFiles(mainPath)
time.sleep(10)

print('Enviando email...')
mailTo = "email@email.com"
subject = "Smart - BD_risco_previa atualizado"
body = "Bom dia!" + "\n" + "O arquivo BD_risco_previa foi atualizado."
SmartMail.sendMail(mailTo, subject, body) #sends email using Outlook   
print('Finalizando o script.')

