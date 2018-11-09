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

for x in range (0, 50):
    print("Week -" + str(x) + ", Bugs found in Production for all projects: ", end="")
    searchQuery = "project in (Bug, \"Cloud Adoption\", LUX, Explore, Kraken, \"Mobile Applications\", \"Value Adds\", Marketplace, \"Development Ops\") AND createdDate >= startOfWeek(-" + str(x) + "w) AND createdDate <= endofWeek(-" + str(x) + "w) AND issueFunction in linkedIssuesOf(\"project = techhelp\") AND type = bug"
    issues = JQL.search_issues(searchQuery, maxResults=200)
    print(len(issues))
