""" Script for reporting out of JIRA. """

from loguru import logger
from getpass import getpass
from jira import JIRA
import datetime

@logger.catch
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
MONTHINPUT = input("How many months do you want to go back? ")
MONTHS = int(MONTHINPUT)
now = datetime.datetime.now()

for x in range (0, MONTHS):

    print("Month -" + str(x) + ", Batch Size: ", end="")

    #Commented because CLOUD doesn't assign developers
    #searchQuery = "project = \"Cloud & Framework\" and developer is not empty and resolved >= startOfMonth(-" + str(x) + "M) and resolved <= endofMonth(-" + str(x) + "M) and type in (Enhancement, \"Internal Development Task\")"
    searchQuery = "project = \"Cloud & Framework\" and resolved >= startOfMonth(-" + str(x) + "M) and resolved <= endofMonth(-" + str(x) + "M) and type in (Enhancement, \"Internal Development Task\")"
    issues = JQL.search_issues(searchQuery, maxResults=200)
    batch_size = 0
    zero_issues = 0
    for issue in issues:
        if issue.fields.timeoriginalestimate is not None:
            if  issue.fields.timeoriginalestimate != 0:
                batch_size += issue.fields.timeoriginalestimate / 3600
            else:
                zero_issues += 1
    if len(issues) - zero_issues != 0:
        print(round(batch_size / len(issues) - zero_issues), "Hours, ", len(issues), "jobs", zero_issues, "Zeros")
    else:
        print("Divide by zero")
