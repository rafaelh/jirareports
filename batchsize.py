# Requires pypiwin32 and jira

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

    print("Month -" + str(x) + ", Batch Size: ", end="")
    searchQuery = "project = \"Cloud & Framework\" and resolved >= startOfMonth(-" + str(x) + "M) and resolved <= endofMonth(-" + str(x) + "M) and type in (Enhancement, \"Internal Development Task\")"
    issues = JQL.search_issues(searchQuery, maxResults=200)
    batch_size = 0
    for issue in issues:
        if issue.fields.timeestimate is not None:
            batch_size += issue.fields.timeestimate / 3600
    print(round(batch_size / len(issues)), "Hours")
