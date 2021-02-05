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

    # Get last month
    lastmonth = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)

    # Get team member information
    teamName = "platforms"
    team = ['rhart', 'rklemm', 'shooper']
    hoursPerDay = 7.6


    totalTeamTime = 0
    for teammember in team:
        #print("\n====== " + teammember + "======")
        alljobs = JQL.search_issues('worklogAuthor = ' + teammember + ' and worklogDate >= startOfMonth(-1) and worklogDate <= endOfMonth(-1)', maxResults=200)

        teammembertotaltime = 0
        for issue in alljobs:
            teammembertotaltime += issue.fields.timespent
            totalTeamTime += issue.fields.timespent
        #print("Total time (hours):", round(totaltime/3600, 2), " (minutes): ", totaltime/60, " (seconds): ", totaltime)

        #print("Jobs worked on: ")
        #for issue in alljobs:
        #    print(issue, "-", issue.fields.summary, " - ", round(issue.fields.timespent/3600, 2), "hrs")


    # Print to screen
    print("\n💻 Team name:            ", teamName)
    print("📅 Month:                ", lastmonth.strftime("%d/%m/%Y"))
    print("🤼 Headcount:            ", len(team))
    print("🕑 Hours Logged in JIRA: ", totalTeamTime/3600, "\n")

    # Write summary to Excel file
    workbook = xlsxwriter.Workbook(lastmonth.strftime("%Y-%m-%d" + " - Team Hours Summary.xlsx"))
    bold = workbook.add_format({'bold': True})

    worksheet = workbook.add_worksheet('Summary')
    row = 0
    col = 0
    data = (
        ['Metric', 'Value'],
        ['Team name', teamName],
        ['Month', lastmonth.strftime("%d/%m/%Y")],
        ['Weekdays in month', '=NETWORKDAYS(DATE(YEAR($B$3),MONTH($B$3),1),$B$3)'],
        ['Less Public Holidays', '0'],
        ['Equals Total Working Days', '=B4-B5'],
        ['Hours per Day', hoursPerDay],
        ['Headcount (from last day of previous month)', len(team)],
        ['Total Working Hours', '=B6*B7*B8'],
        ['Hours Logged in JIRA', totalTeamTime/3600],
        ['% Hours Logged', '=B10/B9']
    )
    for metric, value in (data):
        worksheet.write(row, col, metric)
        worksheet.write(row, col + 1, value)
        row += 1

    worksheet = workbook.add_worksheet('Detail')
    row = 0
    col = 0
    for teammember in team:
        alljobs = JQL.search_issues('worklogAuthor = ' + teammember + ' and worklogDate >= startOfMonth(-1) and worklogDate <= endOfMonth(-1)', maxResults=200)
        for issue in alljobs:
            teammembertotaltime += issue.fields.timespent

        worksheet.write(row, col, teammember, bold)
        worksheet.write(row, col + 1, "Issues", bold)
        worksheet.write(row, col + 2, round(teammembertotaltime/3600, 2), bold)
        row += 1

        for issue in alljobs:
            worksheet.write(row, col, str(issue))
            worksheet.write(row, col + 1, str(issue.fields.summary))
            worksheet.write(row, col + 2, round(issue.fields.timespent/3600, 2))
            row += 1
        row += 1

    workbook.close()


if __name__ == '__main__':
    main()
