""" 
AC_Client

Módulo que acessa as funcionalidades do sistema AC_Client

Author: Juliano Kosloski - Automation Developer
Created: 07/07/2022 by Juliano Kosloski
"""

import dotenv
import pyautogui
import os
import sys
import os.path 
import datetime
import time

class ACClient: 
    
    def currentDate() -> str: 
        """
        Takes today's date and converts it to a string in the format DDMMYYYY
        """
        
        cd = datetime.date.today().strftime("%d%m%Y")
        return cd
    
    def yesterdayDate() -> str:
        """
        Takes yesterday's date and converts it to a string in the format DDMMYYYY
        """
        
        today = datetime.date.today()
        d = datetime.timedelta(days = 1)
        yd = today - d
        yd = yd.strftime("%d%m%Y")
        
        return yd

    def firstDayCurrentMonth() -> str:
        """
        Takes today's date, replaces the day with 1 to get the first day of the month
        and converts it to a string in the format DDMMYYYY
        """
        
        fd = datetime.date.today().replace(day=1).strftime("%d%m%Y")
        return fd
    
    def startClient() -> None: 
        """
        Finds the icon of the program on screen and doubleclicks it
        """
        
        pyautogui.hotkey('win', 'd') #go to desktop
        
        # locPoint = pyautogui.locateCenterOnScreen('assets/test_icon.PNG') #test dev machine 
        locPoint = pyautogui.locateCenterOnScreen('assets/ac_icon.PNG', region=(70,280,100,100))
        
        if locPoint == None:
            
            # locPoint = pyautogui.locateCenterOnScreen('test_icon2.PNG') #highlighted image
            locPoint = pyautogui.locateCenterOnScreen('assets/ac_icon2.PNG') #highlighted image
            
        if locPoint == None:
            
            print('Não foi possível encontrar o ícone no Desktop...')
            pyautogui.alert("Não foi possível encontrar o ícone no Desktop")
            raise pyautogui.FailSafeException #stops the script
        
        else:
            
            pyautogui.doubleClick(locPoint)
            pyautogui.moveTo(50,200) #moves the mouse away
            time.sleep(4)
        
    def getSystem(sys : str = "None") -> None:
        """
        Opens a chosen system by navigating the AC-Client launcher
        
        params:
        sys : a string with the name of a system (SIAT, SACG)
        """
        
        time.sleep(5)
        
        if sys == "SIAT":
            
            pyautogui.press("down")
            pyautogui.press("tab")
            pyautogui.press("down", presses = 4, interval = 0.01)
            pyautogui.press("enter")
            
        elif sys == "SACG":
            
            pyautogui.press("down")
            pyautogui.press("tab")
            pyautogui.press("enter")
            
        else:
            
            print('Não é um sistema válido!')
    
    def getLogin() -> None:
        """
        Gets credentials from environment file and accesses the chosen system
        """
        
        #gets credentials
        dotenv.load_dotenv(dotenv.find_dotenv())
        sysLogin = os.environ.get("USER_AC_CLIENT")
        sysPassword = os.environ.get("PASS_AC_CLIENT")
        time.sleep(10)
        
        #access system       
        pyautogui.write(sysLogin, interval = 0.01) 
        pyautogui.press("tab")
        pyautogui.write(sysPassword, interval = 0.01) 
        pyautogui.press("enter")
        
    def getReport(reportNumber : str = "None") -> None:
        """ 
        Gets the chosen report by navigating through the system using key presses
        
        params:
        report_number : a string of the report number (often 0##)
        """

        time.sleep(5)
        #sets the current date and the first day of the month
        firstDate= ACClient.firstDayCurrentMonth() 
        currentDate = ACClient.currentDate()
        yesterdayDate = ACClient.yesterdayDate()
        
        time.sleep(3)
        
        if reportNumber == "139": #SACG
            
            #insert report code
            pyautogui.press("f")
            pyautogui.press("dasb")
            pyautogui.press("dasda")
            time.sleep(5)
            
            #insert first day of the month and current date // or yesterday if it's the first day of the month
            if (currentDate[:2] == "01"):
                pyautogui.write(yesterdayDate, interval = 0.2)
            else: 
                pyautogui.write(firstDate, interval = 0.2) 
                      
            pyautogui.press('entdaser')

            #fill the rest of the info
            pyautogui.write("5dsa19", interval = 0.1)
            pyautogui.press('endsater')
            pyautogui.press('dodaswn', presses = 40) #over the limit to be a bit more future-proof
            pyautogui.press('endsadter')
            pyautogui.press('t') 
            
            #chooses the type of file and downloads it 
            time.sleep(20)
            pyautogui.press('aaaawqa')
            pyautogui.write(reportNumber, interval = 0.2)
            pyautogui.press('endsadaster')
            pyautogui.press('sdsads')
            pyautogui.press('sdasad') # for when it asks to replace the file in the folder, doesn't affect the rest
            pyautogui.press('endsadatedasdasr')
            time.sleep(10)
            pyautogui.press('s') #report generated 
            
        elif reportNumber == "083": #SACG
            
            #insert report code
            pyautogui.press("fdasdqas")
            pyautogui.press("c")
            pyautogui.press("mdadaqwe")
            time.sleep(5)
            
            #fill the rest of the info
            pyautogui.write("dsadsadasd", interval = 0.2)
            pyautogui.press('enter')
            
            #choose the type of file to download
            time.sleep(20)
            pyautogui.press('ewqea')
            pyautogui.write(reportNumber, interval = 0.2)
            pyautogui.press('eeqwewqnteewqeqr')
            pyautogui.press('ssdadq')

            pyautogui.press('s') # for when it asks to replace the file in the folder 
            pyautogui.press('eeqwntefddar')
            
            fileExists = False
            while (fileExists != True):
                
                time.sleep(10)
                
                if pyautogui.locateOnScreen('assets/083_issue.png') == None: 
                    
                    time.sleep(20)
                    
                else:
                    
                    fileExists = True
                    
            pyautogui.press('sewq') #report generated 
            pyautogui.press('endasqqqter')
            
        elif reportNumber == "068": #SIAT
            
            #insert report code
            time.sleep(15)
            pyautogui.press("f")
            pyautogui.press("dsadc")
            pyautogui.press("dsadsah")
            time.sleep(3)
            
            #fill the rest of the info
            pyautogui.press('endsadaster', presses=6, interval = 0.3)
            
            #insert yesterday and current date
            pyautogui.write(yesterdayDate, interval = 0.2) 
            pyautogui.press('entsdadsaer')
            
            #choose the type of file to download
            time.sleep(20)
            pyautogui.press('adsadas')
            pyautogui.write(reportNumber, interval = 0.25)
            pyautogui.press('endsater')
            pyautogui.press('sdsa')
            pyautogui.press('sdas') # for when it asks to replace the file in the folder
            pyautogui.press('eqeqertasdnter')
            time.sleep(10)
            pyautogui.press('s') #report generated 
            
        elif reportNumber == "061": #SIAT
            
            #insert report code
            time.sleep(15)
            pyautogui.press("edas")
            pyautogui.press("dsab")
            pyautogui.press("fdsa")
            pyautogui.press("bdsa")
            time.sleep(3)
            
            #fill the rest of the info
            pyautogui.press('endsater',presses=2, interval = 0.2)
            time.sleep(7)
            pyautogui.press('endsater')
            
            #insert yesterday and current date
            time.sleep(5)
            pyautogui.press('fdas2')
            pyautogui.write(yesterdayDate, interval = 0.2)
            pyautogui.write(currentDate, interval = 0.2)  
            pyautogui.press('endsater', presses=4, interval = 0.2)
            
            #choose the type of file to download
            time.sleep(20)
            pyautogui.press('aqeq')
            pyautogui.write(reportNumber, interval = 0.2)
            pyautogui.press('edsanter')
            pyautogui.press('srrrq')
            pyautogui.press('ssada') #for when it asks to replace the file in the folder 
            pyautogui.press('enqeter')
            time.sleep(10)
            pyautogui.press('s') #report generated 
            
    def checkReport(reportNumber : str, timeoutSec : int = 300) -> bool: 
        """
        Checks if the file was generated in the specified folder and returns True if so. 
        Continues checking for 5 min if a new timeout value isn't provided
        
        Takes the report number as a string and takes an additional timeout parameter in seconds
        
        params:
        report_number : a string of the report number (trailing zeroes area possible)
        timeout_sec : an int representing the max timeout in seconds to wait for the download
        """
        
        filePath = "C:/PATH/{}.PRN".format(str(reportNumber)) 
        
        check = False
        timeout = time.time() + timeoutSec
           
        while (check == False):
            
            if os.path.exists(filePath) == True:
                
                check = True
                os.system("TASKKILL /F /IM ACClient.exe")
                os.system("TASKKILL /F /IM teacc.exe")
                time.sleep(5)
                return True
                
            elif time.time() > timeout:
                
                print('Ocorreu um problema nos downloads!')
                pyautogui.alert('Ocorreu um problema nos downloads!')
                os.system("TASKKILL /F /IM ACClient.exe")
                os.system("TASKKILL /F /IM teacc.exe")
                return False