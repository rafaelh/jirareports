""" Creates emails based on JIRA statistics """

import datetime
import sys
import webbrowser
from loguru import logger
from getpass import getpass
from jira import JIRA, JIRAError
import win32com.client


@logger.catch
def createemail(emailbody):
    """ Create an email directly in Outlook """
    logger.info("Generating Email")
    olmailitem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newmail = obj.CreateItem(olmailitem)
    today = datetime.date.today()
    newmail.Subject = today.strftime("Development Update - %d %b %Y")
    newmail.HTMLBody = emailbody
    newmail.display()

@logger.catch
def createhtml(htmlbody):
    """ Create an HTML page to cut & paste into email """
    logger.info("Generating HTML")
    f = open('output.html', 'w')
    f.write(BODY)
    f.close
    webbrowser.open('output.html')


if __name__ == "__main__":
    #USERNAME = os.getlogin()
    BODY =     ''
    SERVER =   'https://jira.starrez.com'
    USERNAME = 'rhart'
    print("JIRA Username: " + USERNAME)
    PASSWORD = getpass("JIRA Password: ")
    try:
        JQL = JIRA(server=SERVER, basic_auth=(USERNAME, PASSWORD))
    except JIRAError as error:
        logger.error("Error", error.status_code, "-", error.text)
        sys.exit(1)

