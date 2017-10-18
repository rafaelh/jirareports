# Requires pypiwin32 and jira

# To Do:
# * Increase spacing between bullet points
# * Add Error handling in case the password is wrong

import datetime
from getpass import getpass
import win32com.client
from jira import JIRA


def getpassword():
    """ Get JIRA password, so we aren't hardcoding it """
    password = getpass("JIRA Password: ")
    if not password:
        exit()
    else:
        return password

USERNAME = 'rhart'
print("Username: " + USERNAME)
PASSWORD = getpassword()
JQL = JIRA(server=('https://jira.starrez.com'), basic_auth=(USERNAME, PASSWORD))

BODY = '<html><body>'

for x in range (0, 5):
    print("Getting Month -" + str(x) + " for Web...")
    searchQuery = "project = Web and createdDate >= startOfMonth(-" + str(x) + "M) and createdDate <= endofMonth(-" + str(x) + "M) and issueFunction in linkedIssuesOf(\"project = techhelp\") and type = bug"
    bugList = JQL.search_issues(searchQuery, maxResults=200)
    BODY += "<p>%s web bugs for month %s<br>" % (len(bugList), str(x))
    bugList = None

    print("Getting Month -" + str(x) + " for PortalX...")
    searchQuery = "project = PortalX and createdDate >= startOfMonth(-" + str(x) + "M) and createdDate <= endofMonth(-" + str(x) + "M) and issueFunction in linkedIssuesOf(\"project = techhelp\") and type = bug"
    bugList = JQL.search_issues(searchQuery, maxResults=200)
    BODY += "%s PortalX bugs for month %s</p>" % (len(bugList), str(x))
    bugList = None

    print("Getting Month -" + str(x) + " for StarRez X...")
    searchQuery = "project = \"Mobile Applications\" and createdDate >= startOfMonth(-" + str(x) + "M) and createdDate <= endofMonth(-" + str(x) + "M) and issueFunction in linkedIssuesOf(\"project = techhelp\") and type = bug"
    bugList = JQL.search_issues(searchQuery, maxResults=200)
    BODY += "%s StarRez X bugs for month %s</p>" % (len(bugList), str(x))
    bugList = None

    print("Getting Month -" + str(x) + " for Cloud...")
    searchQuery = "project = \"Cloud & Framework\" and createdDate >= startOfMonth(-" + str(x) + "M) and createdDate <= endofMonth(-" + str(x) + "M) and issueFunction in linkedIssuesOf(\"project = techhelp\") and type = bug"
    bugList = JQL.search_issues(searchQuery, maxResults=200)
    BODY += "%s Cloud bugs for month %s</p>" % (len(bugList), str(x))
    bugList = None

def createemail(emailbody):
    """ Sent Email Contents to Outlook """
    olmailitem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newmail = obj.CreateItem(olmailitem)
    today = datetime.date.today()
    newmail.Subject = today.strftime("Production Bugs - %d %b %Y")
    newmail.HTMLBody = emailbody
    newmail.display()

createemail(BODY)
