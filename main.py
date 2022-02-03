import win32com.client
import ctypes  # for the VM_QUIT to stop PumpMessage()
import pythoncom
import re
import time
import psutil
import requests
import json


class Handler_Class(object):
    def __init__(self):

        # First action to do when using the class in the DispatchWithEvents
        inbox = self.Application.GetNamespace("MAPI").GetDefaultFolder(6).Folders.Item("FOLDER YOU WANT TO LISTEN TO")
        accounts = self.Session.Accounts
        email_account = accounts.Item(2) # get Outlook email account
        print((email_account.DisplayName))
        if email_account.DisplayName == "EMAIL ACCOUNT TO MONITOR":
            messages = inbox.Items
            # Check for unread emails when starting the event
            for message in messages:
                if message.UnRead:
                    print('Unread Mail Subject', message.Subject)  # Or whatever code you wish to execute.

    def OnQuit(self):
        # To stop PumpMessages() when Outlook Quit
        # Note: Not sure it works when disconnecting!!
        ctypes.windll.user32.PostQuitMessage(0)

    """Function that listen on new mails
        
    """
    def OnNewMailEx(self, receivedItemsIDs):
        # RecrivedItemIDs is a collection of mail IDs separated by a ",".
        # You know, sometimes more than 1 mail is received at the same moment.

        for ID in receivedItemsIDs.split(","):
            mail = self.Session.GetItemFromID(ID)
            sender_list = [] # List of sender's email addresses

            # Check if the mail is from any sender in the 
            if mail.SenderEmailAddress in sender_list:
                print(mail.SenderEmailAddress)
                print(mail.Subject)
                
                subject = mail.Subject
                body = mail.Body
                From = mail.SenderEmailAddress

                # perform a post request to Middleware
                url = "http://localhost:8080/api/REMOTE_API"
                
                payload = {'From': From, 'Body': body, 'Subject': subject}
                #headers = {'content-type': 'application/json'}
                response = requests.post(url, data=json.dumps(payload))
                print(response)
                print(response.status_code)
                print(response.text)
                    


# Function to check if outlook is open
def check_outlook_open():
    list_process = []

    for pid in psutil.pids():
        p = psutil.Process(pid) # get the process of that PID
        # Append to the list of process
        # print(p.name())
        list_process.append(p.name())
    # If outlook open then return True
    if "OUTLOOK.EXE" in list_process:
        return True
    else:
        return False


# Loop
while True:
    try:
        outlook_open = check_outlook_open()
    except:
        outlook_open = False
    # If outlook opened then it will start the DispatchWithEvents
    if outlook_open == True:
        outlook = win32com.client.DispatchWithEvents(
            "Outlook.Application", Handler_Class
        )
        
        # ProgID for an object is a short string that names the object and typically creates an instance of the object. For example, Microsoft Excel defines its ProgID as Excel.Application, Microsoft Word defines Word.Application, and so forth. Python programs use the win32com.client.Dispatch() method to create COM objects from a ProgID or CLSID.

        pythoncom.PumpMessages() # PumpMessages() is a blocking call that waits for a message to be posted to the message queue. It is used to process messages posted by the COM server. pythoncom is a module that provides a Python interface to the COM API. it enables Python programs to communicate with COM objects. 


    time.sleep(10)
