from win32com.client import Dispatch
import os
import re
os.chdir("D:\\email")

outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.items
for message in messages:
    message = messages.GetNext()
    name = str(message.subject)
    name = re.sub('[^A-Za-z0-9]+', '', name)+'.msg'    
    message.SaveAs(os.getcwd()+'//'+name)
