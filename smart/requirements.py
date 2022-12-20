from mod.ACClient import ACClient
from mod.HGV import HGV
from mod.OpenData import OpenData
from mod.Risco import Risco
from mod.BD_update_PBI import BD_update_PBI    
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

2. Get the build and dist folders and copy them to a folder in F:/Smart/python_automations

3. Inside the dist folder, create a 'assets' folder with any images that are used

4. Add the .env file to the dist folder

5. Use pyinstaller to compile the other scripts

6. Copy the .exe files created for the other scripts into the dist folder in /python_automations

7. When scheduling the scripts on Task Scheduler, set the action to open the .exe file and to startup in the folder just above 'dist'

8. Everything should work just fine

"""