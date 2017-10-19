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

for x in range (0, 2):
    print("Getting Month -" + str(x) + " for Web...")
    searchQuery = "project = Web and createdDate >= startOfMonth(-" + str(x) + "M) and createdDate <= endofMonth(-" + str(x) + "M) and issueFunction in linkedIssuesOf(\"project = techhelp\") and type = bug"
    bugList = JQL.search_issues(searchQuery, maxResults=200)
    BODY += "<p>%s web bugs for month %s<br>" % (len(bugList), str(x))
    bugList = None

    print("Getting Month -" + str(x) + " for PortalX...")
    searchQuery = "project = PortalX and createdDate >= startOfMonth(-" + str(x) + "M) and createdDate <= endofMonth(-" + str(x) + "M) and issueFunction in linkedIssuesOf(\"project = techhelp\") and type = bug"
    bugList = JQL.search_issues(searchQuery, maxResults=200)
    BODY += "%s PortalX bugs for month %s<br>" % (len(bugList), str(x))
    bugList = None

    print("Getting Month -" + str(x) + " for StarRez X...")
    searchQuery = "project = \"Mobile Applications\" and createdDate >= startOfMonth(-" + str(x) + "M) and createdDate <= endofMonth(-" + str(x) + "M) and issueFunction in linkedIssuesOf(\"project = techhelp\") and type = bug"
    bugList = JQL.search_issues(searchQuery, maxResults=200)
    BODY += "%s StarRez X bugs for month %s<br>" % (len(bugList), str(x))
    bugList = None

    print("Getting Month -" + str(x) + " for Cloud...")
    searchQuery = "project = \"Cloud & Framework\" and createdDate >= startOfMonth(-" + str(x) + "M) and createdDate <= endofMonth(-" + str(x) + "M) and issueFunction in linkedIssuesOf(\"project = techhelp\") and type = bug"
    bugList = JQL.search_issues(searchQuery, maxResults=200)
    BODY += "%s Cloud bugs for month %s<br>" % (len(bugList), str(x))
    bugList = None

    print("Getting Month -" + str(x) + " for UI/UX...")
    searchQuery = "resolved >= startOfMonth(-" + str(x) + "M) and resolved <= endofMonth(-" + str(x) + "M) and \"Epic Link\" in (PORTALX-1508, WEB-7375) or resolved >= startOfMonth(-" + str(x) + "M) and resolved <= endofMonth(-" + str(x) + "M) and Developer = mjack"
    bugList = JQL.search_issues(searchQuery, maxResults=200)
    BODY += "%s UI/UX jobs for month %s<br>" % (len(bugList), str(x))
    bugList = None

    print("Getting Month -" + str(x) + " for Documents...")
    searchQuery = "project = Documentation and resolved >= startOfMonth(-" + str(x) + "M) and resolved <= endofMonth(-" + str(x) + "M)"
    bugList = JQL.search_issues(searchQuery, maxResults=200)
    BODY += "%s Doc jobs for month %s<br>" % (len(bugList), str(x))
    bugList = None

    print("Getting Month -" + str(x) + " Velocity for Web...")
    searchQuery = "project = Web and resolved >= startOfMonth(-" + str(x) + "M) and resolved <= endofMonth(-" + str(x) + "M)"
    bugList = JQL.search_issues(searchQuery, maxResults=200)
    jobtotal = 0
    for issue in bugList:
        if issue.fields.timeoriginalestimate is None:
            result = 0
        else:
            result = issue.fields.timeoriginalestimate
        jobtotal += result / 3600
    BODY += "%s Web jobs for month %s, totalling %s hours<br>" % (len(bugList), str(x), round(jobtotal))
    bugList = None
    jobtotal = None

    print("Getting Month -" + str(x) + " Velocity for PortalX...")
    searchQuery = "project = PortalX and resolved >= startOfMonth(-" + str(x) + "M) and resolved <= endofMonth(-" + str(x) + "M)"
    bugList = JQL.search_issues(searchQuery, maxResults=200)
    jobtotal = 0
    for issue in bugList:
        if issue.fields.timeoriginalestimate is None:
            result = 0
        else:
            result = issue.fields.timeoriginalestimate
        jobtotal += result / 3600
    BODY += "%s PortalX jobs for month %s, totalling %s hours<br>" % (len(bugList), str(x), round(jobtotal))
    bugList = None
    jobtotal = None

    print("Getting Month -" + str(x) + " Velocity for Cloud...")
    searchQuery = "project = \"Cloud & Framework\" and resolved >= startOfMonth(-" + str(x) + "M) and resolved <= endofMonth(-" + str(x) + "M)"
    bugList = JQL.search_issues(searchQuery, maxResults=200)
    jobtotal = 0
    for issue in bugList:
        if issue.fields.timeoriginalestimate is None:
            result = 0
        else:
            result = issue.fields.timeoriginalestimate
        jobtotal += result / 3600
    BODY += "%s Cloud jobs for month %s, totalling %s hours<br>" % (len(bugList), str(x), round(jobtotal))
    bugList = None
    jobtotal = None

    print("Getting Month -" + str(x) + " Velocity for Integrations...")
    searchQuery = "project = \"Custom Development\" and resolved >= startOfMonth(-" + str(x) + "M) and resolved <= endofMonth(-" + str(x) + "M) and Developer is NOT EMPTY"
    bugList = JQL.search_issues(searchQuery, maxResults=200)
    jobtotal = 0
    for issue in bugList:
        if issue.fields.timespent is None:
            result = 0
        else:
            result = issue.fields.timespent
        jobtotal += result / 3600
    BODY += "%s Integration jobs for month %s, totalling %s hours<p>" % (len(bugList), str(x), round(jobtotal))
    bugList = None
    jobtotal = None

BODY += '</body></html>'

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
