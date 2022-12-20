import pyautogui
import time
from mod.BD_update_PBI import BD_update_PBI
from mod.SmartMail import SmartMail

""" 
bd_caderno_seguros

Script que gera relatórios e atualiza o arquivo BD_Caderno Gerencial - Seguros, utilizado no dashboard de Caderno Gerencial

Author: Juliano Kosloski - Automation Developer
Created: 23/11/2022 by Juliano Kosloski
"""

pyautogui.PAUSE = 1.0
pyautogui.hotkey('win', 'd') #go to desktop

#-------------Fluid - Downloading files (8 min on average)-------------------

print("Initiating download...")

sys = "Fluid"
driver = BD_update_PBI.startDriver()
BD_update_PBI.getLogin(driver, sys)

# sheet SEG-NOVO-ENVIADO
url = "https://eqweqweasda.com/9" #  ->>>> SN Processos Recebidos
BD_update_PBI.getLink(driver, url, attempts = 2)
initDay, endDay = BD_update_PBI.startEndPreviousMonth() #returns two days: the start and the end of the previous month
BD_update_PBI.inputDates(driver, initDay, endDay)
sucess = BD_update_PBI.downloadRelFluid(driver) 
if (sucess == False):
    sn_enviado = False
else:
    sn_enviado = True
time.sleep(10)

# sheet SEG-NOVO-EFETIVADO
url = "https://eqweasdas.com" #  ->>>> SN Processos EFETIVADOS
BD_update_PBI.getLink(driver, url, attempts = 2)
initDay, endDay = BD_update_PBI.startEndPreviousMonth() #returns two days: the start and the end of the previous month
BD_update_PBI.inputDates(driver, initDay, endDay)
sucess = BD_update_PBI.downloadRelFluid(driver) 
if (sucess == False):
    sn_efetivado = False
else:
    sn_efetivado = True
time.sleep(10)

# sheet SEG-NOVO-AGENDADO ------------->>>>> outubro não tem nada e não gera relatorio, novembro só tem 2 registros até agora
url = "https://eqweasdas.com" #  ->>>> SN Agendado
BD_update_PBI.getLink(driver, url, attempts = 2)
initDay, endDay = BD_update_PBI.startEndPreviousMonth() #returns two days: the start and the end of the previous month
BD_update_PBI.inputDates(driver, initDay, endDay)
sucess = BD_update_PBI.downloadRelFluid(driver) 
if (sucess == False):
    sn_agendado = False
else:
    sn_agendado = True
time.sleep(10)

# sheet SEG-NOVO-ENVIADO GN ------------->>>>> outubro não tem nada e não gera relatorio
url = "https://eqweasdas.com" #  ->>>> SN Enviados GN
BD_update_PBI.getLink(driver, url, attempts = 2)
initDay, endDay = BD_update_PBI.startEndPreviousMonth() #returns two days: the start and the end of the previous month
BD_update_PBI.inputDates(driver, initDay, endDay)
sucess = BD_update_PBI.downloadRelFluid(driver) 
if (sucess == False):
    sn_enviado_gn = False
else:
    sn_enviado_gn = True
time.sleep(10)

# # sheet SEG-ENDOSSO-ENVIADO 
url = "https://eqweasdas.com" #  ->>>> Endosso Processos Recebidos
BD_update_PBI.getLink(driver, url, attempts = 2)
initDay, endDay = BD_update_PBI.startEndPreviousMonth() #returns two days: the start and the end of the previous month
BD_update_PBI.inputDates(driver, initDay, endDay)
sucess = BD_update_PBI.downloadRelFluid(driver) 
if (sucess == False):
    end_enviado = False
else:
    end_enviado = True
time.sleep(10)

# sheet SEG-ENDOSSO-EFETIVADO 
url = "https://eqweasdas.com" #  ->>>> Endosso Efetivado
BD_update_PBI.getLink(driver, url, attempts = 2)
initDay, endDay = BD_update_PBI.startEndPreviousMonth() #returns two days: the start and the end of the previous month
BD_update_PBI.inputDates(driver, initDay, endDay)
sucess = BD_update_PBI.downloadRelFluid(driver) 
if (sucess == False):
    end_efetivado = False
else:
    end_efetivado = True
time.sleep(10)

# sheet RECUSAS
url = "https://eqweasdas.com" #  ->>>> Recusas Seguros
BD_update_PBI.getLink(driver, url, attempts = 2)
initDay, endDay = BD_update_PBI.startEndPreviousMonth() #returns two days: the start and the end of the previous month
BD_update_PBI.inputDates(driver, initDay, endDay)
sucess = BD_update_PBI.downloadRelFluid(driver) 
if (sucess == False):
    recusas = False
else:
    recusas = True
time.sleep(10)

# sheet Perform - Processos trabalhados 
url = "https://eqweasdas.com" #  ->>>> Seguros Perfor Renovacao Proc Trabalhados
BD_update_PBI.getLink(driver, url, attempts = 2)
initDay, endDay = BD_update_PBI.startEndPreviousMonth() #returns two days: the start and the end of the previous month
BD_update_PBI.inputDates(driver, initDay, endDay)
sucess = BD_update_PBI.downloadRelFluid(driver) 
if (sucess == False):
    perform = False
else:
    perform = True
time.sleep(10)

# sheet Perform - 1. Contato 
url = "https://eqweasdas.com" #  ->>>> Seguros Perfor Renovacao 1. Contato
BD_update_PBI.getLink(driver, url, attempts = 2)
initDay, endDay = BD_update_PBI.startEndPreviousMonth() #returns two days: the start and the end of the previous month
BD_update_PBI.inputDates(driver, initDay, endDay)
sucess = BD_update_PBI.downloadRelFluid(driver) 
if (sucess == False):
    perform1 = False
else:
    perform1 = True
time.sleep(10)

# sheet Perform -minimo 3 dias envio ag 
url = "https://eqweasdas.com" #  ->>>> Seguros Perfor Renovacao Dias Envio GN
BD_update_PBI.getLink(driver, url, attempts = 2)
initDay, endDay = BD_update_PBI.startEndPreviousMonth() #returns two days: the start and the end of the previous month
BD_update_PBI.inputDates(driver, initDay, endDay)
sucess = BD_update_PBI.downloadRelFluid(driver) 
if (sucess == False):
    perform3 = False
else:
    perform3 = True
time.sleep(10)

print("Finishing downloads...")
BD_update_PBI.endSession(driver) 

# pathTo = "C:/PATH/dadas.txt" #DEV
pathTo = "C:/PATH/dsadas.xls" # --> ROBOT

#-------------Copying data from SN Processos Recebidos to BD_Caderno Gerencial - Seguros-------------------

fluidSheet = "SEG-NOVO-ENVIADO"
if (sn_enviado == True):
    print("Copying data from SN Processos Recebidos to BD_Caderno Gerencial - Seguros")
    pattern = r"^4_ewqeasdadasqs.*.xls"
    pathXLS = BD_update_PBI.getFilePath(pattern) 
    pathFrom = BD_update_PBI.saveCorruptedXLSAsXLSX(pathXLS)
    BD_update_PBI.copyFromTo(pathFrom, pathTo, fluidSheet, sys)
    time.sleep(15)
else:
    print(f"{fluidSheet} não foi atualizado por falta de registros.")

#-------------Copying data from SN Processos Efetivados to BD_Caderno Gerencial - Seguros-------------------
fluidSheet = "SEG-NOVO-EFETIVADO"
if (sn_efetivado == True):
    print("Copying data from SN Processos Efetivados to BD_Caderno Gerencial - Seguros")
    pattern = r"^4_ewqeasdadasqs.*.xls"
    pathXLS = BD_update_PBI.getFilePath(pattern) 
    pathFrom = BD_update_PBI.saveCorruptedXLSAsXLSX(pathXLS)
    BD_update_PBI.copyFromTo(pathFrom, pathTo, fluidSheet, sys)
    time.sleep(15)
else:
    print(f"{fluidSheet} não foi atualizado por falta de registros.")

#-------------Copying data from SN Agendado to BD_Caderno Gerencial - Seguros-------------------
fluidSheet = "SEG-NOVO-AGENDADO"
if (sn_agendado == True):
    print("Copying data from SN Agendado to BD_Caderno Gerencial - Seguros")
    pattern = r"^4_ewqeasdadasqs.*.xls"
    pathXLS = BD_update_PBI.getFilePath(pattern) 
    pathFrom = BD_update_PBI.saveCorruptedXLSAsXLSX(pathXLS)
    BD_update_PBI.copyFromTo(pathFrom, pathTo, fluidSheet, sys)
    time.sleep(15)
else:
    print(f"{fluidSheet} não foi atualizado por falta de registros.")
    
#-------------Copying data from SN Enviados GN  to BD_Caderno Gerencial - Seguros-------------------
fluidSheet = "SEG-NOVO-ENVIADO GN"
if (sn_enviado_gn == True):
    print("Copying data from SN Enviados GN to BD_Caderno Gerencial - Seguros")
    pattern = r"^4_ewqeasdadasqs.*.xls"
    pathXLS = BD_update_PBI.getFilePath(pattern) 
    pathFrom = BD_update_PBI.saveCorruptedXLSAsXLSX(pathXLS)
    BD_update_PBI.copyFromTo(pathFrom, pathTo, fluidSheet, sys)
    time.sleep(15)
else:
    print(f"{fluidSheet} não foi atualizado por falta de registros.")

#-------------Copying data from Endosso Processos Recebidos to BD_Caderno Gerencial - Seguros-------------------

fluidSheet = "SEG-ENDOSSO-ENVIADO"
if (end_enviado == True):
    print("Copying data from Endosso Processos Recebidos to BD_Caderno Gerencial - Seguros")
    pattern = r"^4_ewqeasdadasqs.*.xls"
    pathXLS = BD_update_PBI.getFilePath(pattern) 
    pathFrom = BD_update_PBI.saveCorruptedXLSAsXLSX(pathXLS)
    BD_update_PBI.copyFromTo(pathFrom, pathTo, fluidSheet, sys)
    time.sleep(15)
else:
    print(f"{fluidSheet} não foi atualizado por falta de registros.")
    
# #-------------Copying data from Endossso Efetivado to BD_Caderno Gerencial - Seguros-------------------

fluidSheet = "SEG-ENDOSSO-EFETIVADO"
if (end_efetivado == True):
    print("Copying data from Endosso Efetivado to BD_Caderno Gerencial - Seguros")
    pattern = r"^4_ewqeasdadasqs.*.xls"
    pathXLS = BD_update_PBI.getFilePath(pattern) 
    pathFrom = BD_update_PBI.saveCorruptedXLSAsXLSX(pathXLS)
    BD_update_PBI.copyFromTo(pathFrom, pathTo, fluidSheet, sys)
    time.sleep(15)
else:
    print(f"{fluidSheet} não foi atualizado por falta de registros.")
    
# #-------------Copying data from Recusas Seguros to BD_Caderno Gerencial - Seguros-------------------

fluidSheet = "RECUSAS"
if (recusas == True):
    print("Copying data from Recusas Seguros to BD_Caderno Gerencial - Seguros")
    pattern = r"^4_ewqeasdadasqs.*.xls"
    pathXLS = BD_update_PBI.getFilePath(pattern) 
    pathFrom = BD_update_PBI.saveCorruptedXLSAsXLSX(pathXLS)
    BD_update_PBI.copyFromTo(pathFrom, pathTo, fluidSheet, sys)
    time.sleep(15)
else:
    print(f"{fluidSheet} não foi atualizado por falta de registros.")
    
# #-------------Copying data from Seguros Perform Renovacao Proc Trabalhados to BD_Caderno Gerencial - Seguros-------------------

fluidSheet = "Perform - Processos trabalhados"
if (perform == True):
    print("Copying data from Seguros Perform Renovacao Proc Trabalhados to BD_Caderno Gerencial - Seguros")
    pattern = r"^4_ewqeasdadasqs.*.xls"
    pathXLS = BD_update_PBI.getFilePath(pattern) 
    pathFrom = BD_update_PBI.saveCorruptedXLSAsXLSX(pathXLS)
    BD_update_PBI.copyFromTo(pathFrom, pathTo, fluidSheet, sys)
    time.sleep(15)
else:
    print(f"{fluidSheet} não foi atualizado por falta de registros.")
    
# #-------------Copying data from Seguros Perform Renovacao 1. Contato to BD_Caderno Gerencial - Seguros-------------------

fluidSheet = "Perform - 1. Contato"
if (perform1 == True):
    print("Copying data from Seguros Perform Renovacao 1. Contato to BD_Caderno Gerencial - Seguros")
    pattern = r"^4_ewqeasdadasqs.*.xls"
    pathXLS = BD_update_PBI.getFilePath(pattern) 
    pathFrom = BD_update_PBI.saveCorruptedXLSAsXLSX(pathXLS)
    BD_update_PBI.copyFromTo(pathFrom, pathTo, fluidSheet, sys)
    time.sleep(15)
else:
    print(f"{fluidSheet} não foi atualizado por falta de registros.")
    
# #-------------Copying data from Seguros Perform Renovacao Dias Envio GN to BD_Caderno Gerencial - Seguros-------------------

fluidSheet = "Perform -minimo 3 dias envio ag"
if (perform3 == True):
    print("Copying data from Seguros Perform Renovacao Dias Envio GN to BD_Caderno Gerencial - Seguros")
    pattern = r"^4_ewqeasdadasqs.*.xls"
    pathXLS = BD_update_PBI.getFilePath(pattern) 
    pathFrom = BD_update_PBI.saveCorruptedXLSAsXLSX(pathXLS)
    BD_update_PBI.copyFromTo(pathFrom, pathTo, fluidSheet, sys)
    time.sleep(15)
else:
    print(f"{fluidSheet} não foi atualizado por falta de registros.")
    
print("TODOS OS DADOS COPIADOS")

# -------------Sending email--------------------

print('Enviando email...')
mailTo = "dsadadqq@eqeqwewq.com"
subject = "Smart - BD_Caderno Seguros atualizado"
body = "Bom dia!" + "\n" + "O arquivo BD_Caderno Seguros foi atualizado."
SmartMail.sendMail(mailTo, subject, body) #sends email using Outlook 
  
print('Finalizando o script.')