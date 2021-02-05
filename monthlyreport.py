#!/usr/bin/env python3
''' Script to get last month's hours and work for a set of team members '''

import os
import sys
import jira
import datetime
import xlsxwriter
from getpass import getpass
from loguru import logger

@logger.catch
def main():
    # Manually set variables
    jirausername = 'rhart'
    teamName = "Platforms"
    team = ['rhart', 'rklemm', 'shooper']
    hoursPerDay = 7.6

    # Connect to JIRA
    jiraserver = 'https://jira.starrez.com'
    print("JIRA Username: " + jirausername)
    jirapassword = getpass("JIRA Password: ")
    try:
        JQL = jira.JIRA(server=jiraserver, basic_auth=(jirausername, jirapassword))
    except jira.JIRAError() as error:
        logger.error("Error", error.status_code, "-", error.text)
        sys.exit(1)

    # Get totals
    lastmonth = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
    totalTeamTime = 0
    totalBugTime = 0
    totalTechTime = 0
    for teammember in team:
        alljobs = JQL.search_issues('worklogAuthor = ' + teammember + ' and worklogDate >= startOfMonth(-1) and worklogDate <= endOfMonth(-1)', maxResults=200)

        teammembertotaltime = 0 # All an individual's time logged on jobs (regardless of type)
        for issue in alljobs:
            teammembertotaltime += issue.fields.timespent
            if 'BUG' in str(issue.key):
                totalBugTime += issue.fields.timespent  # Time on bugs doesn't add to team total
            elif 'TECHHELP' in str(issue.key):
                totalTechTime += issue.fields.timespent # Time on Techhelps doesn't add to total
            else:
                totalTeamTime += issue.fields.timespent # All other issues add to total time

    # Print to screen
    print("\n💻 Team name:            ", teamName)
    print("📅 Month:                ", lastmonth.strftime("%d/%m/%Y"))
    print("🤼 Headcount:            ", len(team))
    print("🕑 Hours Logged in JIRA: ", totalTeamTime/3600)
    print("🐛 Hours on BUGs:        ", totalBugTime/3600)
    print("🔧 Hours on TECHHELPs:   ", totalTechTime/3600, "\n")

    # Create an excel file
    filename = lastmonth.strftime("%Y-%m-%d" + " - Team Hours Summary.xlsx")
    if os.path.exists(filename):
        print("💥 A previous excel file exists - this can cause formatting errors")
    workbook = xlsxwriter.Workbook(filename)
    bold = workbook.add_format({'bold': True})
    percentage = workbook.add_format()
    percentage.set_num_format('0.00%')
    percentage.set_align('left')
    aligncenter = workbook.add_format()
    aligncenter.set_align('left')

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
        ['Hours Logged in JIRA (Total)', totalTeamTime/3600],
        ['Hours Logged on BUGs', totalBugTime/3600],
        ['Hours Logged on TECHHELPs', totalTechTime/3600]
    )
    worksheet1.write(row, col, "Metric", bold)
    worksheet1.write(row, col + 1, "Value", bold)
    row += 1
    for metric, value in (data):
        worksheet1.write(row, col, metric)
        worksheet1.write(row, col + 1, value, aligncenter)
        row += 1
    worksheet1.write(row, col, "% Hours Logged")
    worksheet1.write(row, col + 1, '=B10/B9', percentage)

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

    # Write the Excel file
    workbook.close()
    print("\n🧮 Excel file written as: " + filename + "\n")


if __name__ == '__main__':
    main()
