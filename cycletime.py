# Requires jira

from datetime import datetime
from getpass import getpass
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

for x in range (0, 70):

    print("Month -" + str(x) + ", Cycletime for PortalX: ", end="")
    searchQuery = "project = \"Cloud & Framework\" and resolved >= startOfMonth(-" + str(x) + "M) and resolved <= endofMonth(-" + str(x) + "M) and type in (Enhancement, \"Internal Development Task\")"
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
