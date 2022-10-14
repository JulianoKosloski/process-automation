import pyautogui
import time
from mod.SYS3 import SYS3
from mod.SmartMail import SmartMail


pyautogui.PAUSE = 1.0

pyautogui.hotkey('win', 'd') #go to desktop

#-------------SYS3 - Downloading files-------------------

print("Initiating download...")

driver = SYS3.startDriver()
SYS3.getLogin(driver)
SYS3.getLink(driver, "https://dasdaslo/dsada/adseqw", 1)
SYS3.downloadRelFluid(driver)
time.sleep(15)

print("Finishing download...")
SYS3.endSession(driver)

#-------------Copying data to file-------------------

pathXLS = SYS3.getFilePath()
pathFrom = SYS3.saveAsXLSX(pathXLS)
# pathTo = "C:/Users/dsadasd/adasdas.xlsx"  #----DEV
pathTo = "C:/Users/dsadadsa/dsadasd.xlsx"
SYS3.copyFromTo(pathFrom, pathTo)

# -------------Sending email--------------------

print('Enviando email...')
mailTo = "sadas@dasdasdsa.com"
subject = "file atualizado"
body = "Bom dia!" + "\n" + "O arquivo dasdas foi atualizado."
SmartMail.sendMail(mailTo, subject, body) #sends email using Outlook 
  
print('Finalizando o script.')