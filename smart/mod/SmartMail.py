""" 
SmartMail

Módulo que acessa as funcionalidades do email Outlook

Author: Juliano Kosloski - Automation Developer
Created: 11/08/2022 by Juliano Kosloski
"""

import win32com.client as win32

class SmartMail:
    
    def sendMail(mailTo : str = "", subject : str = "", body = "", attach : bool = False, filePath : str = "") -> None:
        """
        Takes mail information, passes it to outlook on the user's machine and sends an email
        
        params:
        mailTo: email string
        subject: email subject string
        body: email body string
        attach: attach a file bool (optional)
        filePath: file path string (optional)
        """
        
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = mailTo
        mail.Subject = subject
        mail.Body = body
        
        if attach == True: #checks if the user wants to attach a file
            
            print("Anexando arquivo...")
            mail.Attachments.add(filePath)
        
        print('Enviando email...')
        mail.Send() #sends the email
    
