import win32com.client as win32
import re

""" 
SmartMail

Módulo que acessa as funcionalidades do email Outlook

Author: Juliano Kosloski - Automation Developer
Created: 11/08/2022 by Juliano Kosloski
"""

class SmartMail:
    
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    
    """
    Index of Outlook items:
    3 Deleted Items
    4 Outbox 
    5 Sent Items
    6 Inbox
    16 Drafts
    """
    
    def _getJuridicoEmail() -> None:
        """
        Finds the two latest emails on the inbox and download its attachments
        
        """
        message = SmartMail.messages.GetFirst()
        pattern1 = "dsadaIDAS.*"
        pattern2 = "ORdasdasdsaq JUdsdasewS.*"
        r1 = re.compile(pattern1)
        r2 = re.compile(pattern2)
        
        while True:
            
            text = message.subject
            
            if (r1.match(text)):
                attachments = message.Attachments# return the first item in attachments
                attachment = attachments.Item(1)
                # the name of attachment file      
                attachment_name1 = str(attachment)
                print("Found a file: " + attachment_name1)
                attachment.SaveASFile('C:/PATH/' + attachment_name1)
                # -----
                attachment = attachments.Item(2)
                attachment_name1a = str(attachment)
                print("Found a file: " + attachment_name1a)
                attachment.SaveASFile('C:/PATH/' + attachment_name1a)
            
            if (r2.match(text)):
                attachments = message.Attachments# return the first item in attachments
                attachment = attachments.Item(1)
                # the name of attachment file      
                attachment_name2 = str(attachment)
                print("Found a file: " + attachment_name2)
                attachment.SaveASFile('C:/PATH/' + attachment_name2)
                # -----
                attachment = attachments.Item(2)
                attachment_name2a = str(attachment)
                print("Found a file: " + attachment_name2a)
                attachment.SaveASFile('C:/PATH/' + attachment_name2a)
                break
            
            message = SmartMail.messages.GetNext()
                
        return attachment_name1, attachment_name1a, attachment_name2, attachment_name2a
    
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
    
