""" Creates an template for the bi-weekly velocity update """

import datetime
import win32com.client

# Create Email Contents
print("Generating Email...")
with open('velocity-template.html', 'r') as emailFormat:
    BODY = emailFormat.read().replace('\n', '')

def createemail(emailbody):
    """ Sent Email Contents to Outlook """
    olmailitem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newmail = obj.CreateItem(olmailitem)
    today = datetime.date.today()
    newmail.Subject = today.strftime("Velocity Update - %d %b %Y")
    newmail.HTMLBody = emailbody
    newmail.display()

createemail(BODY)
