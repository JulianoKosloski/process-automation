import pyautogui
from mod.SYS2 import SYS2
from mod.SYS5 import SYS5
from mod.SmartMail import SmartMail
import sys

""" 
rel_y

Script que gera relatórios de risco utilizando o sistema 2, processa os arquivos e envia emails
para dasdas!.

Author: Juliano Kosloski - Automation Developer
Created: 08/08/2022 by Juliano Kosloski
"""

pyautogui.PAUSE = 1.0

pyautogui.hotkey('win', 'd') #go to desktop

#HGV - Downloading files 

print('Baixando arquivos ...')
driver = SYS2.startDriver()
urlRisco = "https://dasdasds/dasdeerq.php"
xpathList = [r'//*[@id="btnExportar1"]',r'//*[@id="btnExportar"]'] #lists the buttoms 'Exportar' and 'Exp. Detalhes'

SYS2.getLogin(driver)
SYS2.getLink(driver, urlRisco, attempts = 2) #needs to go to the url twice to get the right page
SYS2.downloadFiles(driver, xpathList)
SYS2.endSession(driver)

#Risco - manipulating .csv files

print('Buscando arquivos na pasta...')
arqPath, detPath, detFileName = SYS5.getFiles() #gets file paths

#save files as .xlsx

print('Transformando arquivos em .xlsx...')
arqPath = SYS5.openFile(arqPath) 
arqNewPath = SYS5.saveAsXLSX(arqPath)
SYS5.closeExcel()

detPath = SYS5.openFile(detPath) 
detNewPath = SYS5.saveAsXLSX(detPath)
SYS5.closeExcel()

mainPath = "F:/sdadqewq/ytry.xlsx" #changing to server path might solve some issues with task scheduler

#prepares files and copies them to the main file

print('Preparando arquivos...')
arqPath, detPath = SYS5.prepFiles(arqNewPath, detNewPath)
SYS5.clearMainFile(mainPath)
print('Copiando dados para o arquivo principal...')
mainPath = SYS5.appendMainFile(arqPath, detPath, mainPath) 

#opens main file for computing 

print('Atualizando tabelas...')
SYS5.updateTables(mainPath)
SYS5.diffAgencia(mainPath)
SYS5.updateTables(mainPath)
SYS5.diffMotivo(mainPath)

#saves files to different folders
SYS5.saveFiles(mainPath)

print('Enviando email...')
mailTo = "eqwy@eqpooke.com"
subject = "Smart - dsadaseqw atualizado"
body = "Bom dia!" + "\n" + "O arquivo rqyere foi atualizado."
SmartMail.sendMail(mailTo, subject, body) #sends email using Outlook   
print('Finalizando o script.')

