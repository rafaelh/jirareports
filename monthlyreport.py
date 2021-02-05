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

    # Manually set variables
    teamName = "Platforms"
    team = ['rhart', 'rklemm', 'shooper']
    hoursPerDay = 7.6

    # Get team member information
    lastmonth = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
    totalTeamTime = 0
    for teammember in team:
        alljobs = JQL.search_issues('worklogAuthor = ' + teammember + ' and worklogDate >= startOfMonth(-1) and worklogDate <= endOfMonth(-1)', maxResults=200)

        teammembertotaltime = 0
        for issue in alljobs:
            teammembertotaltime += issue.fields.timespent
            totalTeamTime += issue.fields.timespent

    # Print to screen
    print("\n💻 Team name:            ", teamName)
    print("📅 Month:                ", lastmonth.strftime("%d/%m/%Y"))
    print("🤼 Headcount:            ", len(team))
    print("🕑 Hours Logged in JIRA: ", totalTeamTime/3600, "\n")

    # Create an excel file
    workbook = xlsxwriter.Workbook(lastmonth.strftime("%Y-%m-%d" + " - Team Hours Summary.xlsx"))
    bold = workbook.add_format({'bold': True})

    # Summary page
    worksheet1 = workbook.add_worksheet('Summary')
    worksheet1.set_column('A:A', 40)
    worksheet1.set_column('B:B', 10)
    row = 0
    col = 0
    data = (
        ['Team name', teamName],
        ['Month', lastmonth.strftime("%d/%m/%Y")],
        ['Weekdays in month', '=NETWORKDAYS(DATE(YEAR($B$3),MONTH($B$3),1),$B$3)'],
        ['Less Public Holidays', 0],
        ['Equals Total Working Days', '=B4-B5'],
        ['Hours per Day', hoursPerDay],
        ['Headcount (from last day of previous month)', len(team)],
        ['Total Working Hours', '=B6*B7*B8'],
        ['Hours Logged in JIRA', totalTeamTime/3600],
        ['% Hours Logged', '=B10/B9']
    )
    worksheet1.write(row, col, "Metric", bold)
    worksheet1.write(row, col + 1, "Value", bold)
    row += 1
    for metric, value in (data):
        worksheet1.write(row, col, metric)
        worksheet1.write(row, col + 1, value)
        row += 1

    # Detail page
    worksheet2 = workbook.add_worksheet('Detail')
    worksheet2.set_column('A:A', 13)
    worksheet2.set_column('B:B', 80)
    worksheet2.set_column('C:C', 7)
    row = 0
    col = 0
    for teammember in team:
        alljobs = JQL.search_issues('worklogAuthor = ' + teammember + ' and worklogDate >= startOfMonth(-1) and worklogDate <= endOfMonth(-1)', maxResults=200)
        for issue in alljobs:
            teammembertotaltime += issue.fields.timespent

        worksheet2.write(row, col, teammember, bold)
        worksheet2.write(row, col + 1, "Issues", bold)
        worksheet2.write(row, col + 2, round(teammembertotaltime/3600, 2), bold)
        row += 1

        for issue in alljobs:
            worksheet2.write(row, col, str(issue))
            worksheet2.write(row, col + 1, str(issue.fields.summary))
            worksheet2.write(row, col + 2, round(issue.fields.timespent/3600, 2))
            row += 1
        row += 1

    # Write the Excel file out
    workbook.close()
    print("\nExcel file written as: " + workbook + "\n")


if __name__ == '__main__':
    main()
