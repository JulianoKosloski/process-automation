from mod.SYS1 import SYS1
from mod.SYS2 import SYS2
from mod.SYS4 import SYS4
from mod.SYS5 import SYS5
from mod.SYS3 import SYS3    
from mod.SmartMail import SmartMail 

import os
import os.path 
import datetime
import time
import sys
import logging
import dotenv 
import pyautogui
import pandas
import xlsxwriter
import re
import openpyxl
import win32com.client as win32
import requests 
import selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By


#all possible dependencies for automation scripts until now

"""
README

Keep list of imports updated
Remember to change all paths to the robot machine before compiling
Use relative paths to call images, such as 'assets/image.png'

1. On the terminal, run: 

pyinstaller requirements.py 

2. Get the build and dist folders and copy them to a folder in F:/Folder/python_automations

3. Inside the dist folder, create a 'assets' folder with any images that are used

4. Add the .env file to the dist folder

5. Use pyinstaller to compile the other scripts

6. Copy the .exe files created for the other scripts into the dist folder in /python_automations

7. When scheduling the scripts on Task Scheduler, set the action to open the .exe file and to startup in the folder just above 'dist'

8. Everything should work just fine

"""