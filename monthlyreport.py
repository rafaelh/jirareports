#!/usr/bin/env python3
''' Script to get last month's hours and work for a set of team members '''

import sys
import jira
from getpass import getpass
from loguru import logger

@logger.catch
def main():
    # Connect to JIRA
    jiraserver = 'https://jira.starrez.com'
    jirausername = 'rhart'
    print("JIRA Username: " + jirausername)
    jirapassword = getpass("JIRA Password: ")
    try:
        JQL = jira.JIRA(server=jiraserver, basic_auth=(jirausername, jirapassword))
    except jira.JIRAError() as error:
        logger.error("Error", error.status_code, "-", error.text)
        sys.exit(1)

    # Get team member information
    team = ['rhart', 'rklemm']

    for teammember in team:
        print("====== " + teammember + "======")
        alljobs = JQL.search_issues('worklogAuthor = ' + teammember + ' and worklogDate >= startOfMonth(-1) and worklogDate <= endOfMonth(-1)', maxResults=200)

        totaltime = 0
        for issue in alljobs:
            totaltime += issue.fields.timespent
        print("Total time (hours):", round(totaltime/3600, 2), " (minutes): ", totaltime/60, " (seconds): ", totaltime)

        print("Jobs worked on: ")
        for issue in alljobs:
            print(issue, "-", issue.fields.summary, " - ", round(issue.fields.timespent/3600, 2), "hrs")


if __name__ == '__main__':
    main()