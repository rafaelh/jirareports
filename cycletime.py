# Requires pypiwin32 and jira

from datetime import datetime
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

for x in range (60, 70):

    print("Month -" + str(x) + ", Cycletime for Web: ", end="")
    searchQuery = "project = Web and resolved >= startOfMonth(-" + str(x) + "M) and resolved <= endofMonth(-" + str(x) + "M) and type in (Enhancement, \"Internal Development Task\")"
    issues = JQL.search_issues(searchQuery, maxResults=200)
    cycle_time = 0
    for issue in issues:
        #print("Raw Created: %s, Raw Resolved: %s" % (issue.fields.created, issue.fields.resolutiondate))
        created_date = datetime.strptime(issue.fields.created, '%Y-%m-%dT%H:%M:%S.%f%z').date()
        resolved_date = datetime.strptime(issue.fields.resolutiondate, '%Y-%m-%dT%H:%M:%S.%f%z').date()
        difference = resolved_date - created_date
        #print("Job: %s, Created: %s, Resolved: %s, Difference: %s" % (issue.key, created_date, resolved_date, difference.days))
        cycle_time += difference.days
    print(round(cycle_time / len(issues)))

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

#createemail(BODY)
