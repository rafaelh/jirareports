""" Creates emails based on JIRA statistics """

import datetime
import sys
from getpass import getpass
import win32com.client
from jira import JIRA, JIRAError

# TODO: structure this properly as a module

#USERNAME = os.getlogin()
USERNAME = 'rhart'
print("JIRA Username: " + USERNAME)
PASSWORD = getpass("JIRA Password: ")

try:
    JQL = JIRA(server=('https://jira.starrez.com'), basic_auth=(USERNAME, PASSWORD))
except JIRAError as error:
    print("Error", error.status_code, "-", error.text)
    sys.exit(1)

BODY = ''

# Create an email using the assembled information
def createemail(emailbody):
    """ Sent Email Contents to Outlook """
    olmailitem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newmail = obj.CreateItem(olmailitem)
    today = datetime.date.today()
    newmail.Subject = today.strftime("Development Update - %d %b %Y")
    newmail.HTMLBody = emailbody
    newmail.display()

def create