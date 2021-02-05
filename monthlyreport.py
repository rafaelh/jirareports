#!/usr/bin/env python3
''' Script to get last month's hours and work for a set of team members '''

import sys
import jira
import datetime
import xlsxwriter
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

    # Time wangling
    month = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)

    # Get team member information
    teamName = "platforms"
    team = ['rhart', 'rklemm', 'shooper']
    hoursPerDay = 7.6
    print("\n\nTeam name: " + teamName)
    print("Month: " + month.strftime("%d/%m/%Y"))
    print("Weekdays in month: ")
    print("Less Public Holidays: ")
    print("Hours per day: " + str(hoursPerDay))
    print("Headcount: " + str(len(team)))

    totalTeamTime = 0
    for teammember in team:
        #print("\n====== " + teammember + "======")
        alljobs = JQL.search_issues('worklogAuthor = ' + teammember + ' and worklogDate >= startOfMonth(-1) and worklogDate <= endOfMonth(-1)', maxResults=200)

        totaltime = 0
        for issue in alljobs:
            totaltime += issue.fields.timespent
            totalTeamTime += issue.fields.timespent
        #print("Total time (hours):", round(totaltime/3600, 2), " (minutes): ", totaltime/60, " (seconds): ", totaltime)

        #print("Jobs worked on: ")
        #for issue in alljobs:
        #    print(issue, "-", issue.fields.summary, " - ", round(issue.fields.timespent/3600, 2), "hrs")
    print("Hours Logged in Jira: " + str(totalTeamTime/3600))

    # Set output file
    workbook = xlsxwriter.Workbook(month.strftime("%Y-%m-%d" + " - Team Hours Summary.xlsx"))
    worksheet = workbook.add_worksheet()

    data = (
        ['','Metric', 'Value'],
        ['💻', 'Team name', teamName],
        ['📅', 'Month', month.strftime("%d/%m/%Y")],
        ['📅', 'Weekdays in month', '=NETWORKDAYS(DATE(YEAR($B$3),MONTH($B$3),1),$B$3)'],
        ['📅', 'Less Public Holidays', '0'],
        ['📅', 'Equals Total Working Days', '=B4-B5'],
        ['⏳', 'Hours per Day', hoursPerDay],
        ['🤼', 'Headcount (from last day of previous month)', len(team)],
        ['🕑', 'Total Working Hours', '=B6*B7*B8'],
        ['🕑', 'Hours Logged in JIRA', totalTeamTime/3600],
        ['💹', '% Hours Logged', '=B10/B9']
    )

    # Print to screen
    for emoji, metric, value in (data):
        print(emoji, metric, ": ", value)

    row = 0
    col = 0
    for emoji, metric, value in (data):
        worksheet.write(row, col, metric)
        worksheet.write(row, col + 1, value)
        row += 1
    workbook.close()


if __name__ == '__main__':
    main()
