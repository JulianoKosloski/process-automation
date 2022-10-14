""" 
SYS1

Módulo que acessa as funcionalidades do sistema SYS1

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

#TODO #8 add try/except blocks

class SYS1: 
    
    def currentDate() -> str: 
        """
        Takes today's date and converts it to a string in the format DDMMYYYY
        """
        
        cd = datetime.date.today().strftime("%d%m%Y")
        return cd

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
        Opens a chosen system by navigating the X launcher
        
        params:
        sys : a string with the name of a system (AAA, BBB)
        """
        
        if sys == "AAA":
            
            pyautogui.press("down")
            pyautogui.press("tab")
            pyautogui.press("down", presses = 4, interval = 0.01)
            pyautogui.press("enter")
            
        elif sys == "BBB":
            
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
        sysLogin = os.environ.get("USER_SYS1")
        sysPassword = os.environ.get("PASS_SYS1")
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
    
        #sets the current date and the first day of the month
        firstDate= SYS1.firstDayCurrentMonth() 
        currentDate = SYS1.currentDate()
        time.sleep(3)
        
        if reportNumber == "139": #AAA
            
            #insert report code
            pyautogui.press("f")
            pyautogui.press("b")
            pyautogui.press("a")
            time.sleep(3)
            
            #insert first day of the month and current date
            pyautogui.write(firstDate, interval = 0.2) 
            pyautogui.press('enter')
            
            #fill the rest of the info
            pyautogui.write("519", interval = 0.1)
            pyautogui.press('enter')
            pyautogui.press('down', presses = 40) #over the limit to be a bit more future-proof
            pyautogui.press('enter')
            pyautogui.press('t') 
            
            #chooses the type of file and downloads it 
            time.sleep(10)
            pyautogui.press('a')
            pyautogui.write(reportNumber, interval = 0.2)
            pyautogui.press('enter')
            pyautogui.press('s')
            pyautogui.press('s') # for when it asks to replace the file in the folder, doesn't affect the rest
            pyautogui.press('enter')
            time.sleep(10)
            pyautogui.press('s') #report generated 
            
        elif reportNumber == "083": #AAA
            
            #insert report code
            pyautogui.press("f")
            pyautogui.press("c")
            pyautogui.press("m")
            time.sleep(3)
            
            #fill the rest of the info
            pyautogui.write("00001-2", interval = 0.2)
            pyautogui.press('enter')
            
            #choose the type of file to download
            time.sleep(2)
            pyautogui.press('a')
            pyautogui.write(reportNumber, interval = 0.2)
            pyautogui.press('enter')
            pyautogui.press('s')

            pyautogui.press('s') # for when it asks to replace the file in the folder 
            pyautogui.press('enter')
            
            fileExists = False
            while (fileExists != True):
                
                time.sleep(10)
                
                if pyautogui.locateOnScreen('assets/083_issue.png') == None: 
                    
                    time.sleep(20)
                    
                else:
                    
                    fileExists = True
                    
            pyautogui.press('s') #report generated 
            pyautogui.press('enter')
            
        elif reportNumber == "068": #BBB
            
            #insert report code
            time.sleep(5)
            pyautogui.press("f")
            pyautogui.press("c")
            pyautogui.press("h")
            time.sleep(3)
            
            #fill the rest of the info
            pyautogui.press('enter', presses=6, interval = 0.3)
            
            #insert first day of the month and current date
            pyautogui.write(firstDate, interval = 0.2) 
            pyautogui.press('enter')
            
            #choose the type of file to download
            time.sleep(2)
            pyautogui.press('a')
            pyautogui.write(reportNumber, interval = 0.25)
            pyautogui.press('enter')
            pyautogui.press('s')
            pyautogui.press('s') # for when it asks to replace the file in the folder
            pyautogui.press('enter')
            time.sleep(10)
            pyautogui.press('s') #report generated 
            
        elif reportNumber == "061": #BBB
            
            #insert report code
            time.sleep(5)
            pyautogui.press("e")
            pyautogui.press("b")
            pyautogui.press("f")
            pyautogui.press("b")
            time.sleep(3)
            
            #fill the rest of the info
            pyautogui.press('enter',presses=2, interval = 0.2)
            time.sleep(7)
            pyautogui.press('enter')
            
            #insert first day of the month and current date
            time.sleep(2)
            pyautogui.press('f2')
            pyautogui.write(firstDate, interval = 0.2)
            pyautogui.write(currentDate, interval = 0.2)  
            pyautogui.press('enter', presses=4, interval = 0.2)
            
            #choose the type of file to download
            time.sleep(2)
            pyautogui.press('a')
            pyautogui.write(reportNumber, interval = 0.2)
            pyautogui.press('enter')
            pyautogui.press('s')
            pyautogui.press('s') #for when it asks to replace the file in the folder 
            pyautogui.press('enter')
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
        
        filePath = "C:/TEMP/{}.PRN".format(str(reportNumber)) 
        
        check = False
        timeout = time.time() + timeoutSec
           
        while (check == False):
            
            if os.path.exists(filePath) == True:
                
                check == True
                os.system("TASKKILL /F /IM SYSTEM!!!.exe")
                os.system("TASKKILL /F /IM SYSTEM???.exe")
                return True
                
            elif time.time() > timeout:
                
                print('Ocorreu um problema nos downloads!')
                pyautogui.alert('Ocorreu um problema nos downloads!')
                os.system("TASKKILL /F /IM SYSTEM!!!.exe")
                os.system("TASKKILL /F /IM SYSTEM???.exe")
                return False