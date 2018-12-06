""" Creates emails based on JIRA statistics """

import datetime
import sys
import webbrowser
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
    """ Create an email directly in Outlook """
    olmailitem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newmail = obj.CreateItem(olmailitem)
    today = datetime.date.today()
    newmail.Subject = today.strftime("Development Update - %d %b %Y")
    newmail.HTMLBody = emailbody
    newmail.display()

def createhtml(htmlbody):
    """ Create an HTML page to cut & paste into email """
    print("Generating HTML...")
    f = open('output.html', 'w')
    f.write(BODY)
    f.close
    webbrowser.open('output.html')