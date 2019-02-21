""" Creates emails based on JIRA statistics """

import datetime
import sys
import webbrowser
from getpass import getpass
from loguru import logger
from jira import JIRA, JIRAError


@logger.catch
def createhtml(htmlbody):
    """ Create an HTML page to cut & paste into email """
    logger.info("Generating HTML")
    file = open('output.html', 'w')
    file.write(BODY)
    file.close
    webbrowser.open('output.html')


if __name__ == "__main__":
    #USERNAME = os.getlogin()
    BODY = ''
    SERVER = 'https://jira.starrez.com'
    USERNAME = 'rhart'
    print("JIRA Username: " + USERNAME)
    PASSWORD = getpass("JIRA Password: ")
    try:
        JQL = JIRA(server=SERVER, basic_auth=(USERNAME, PASSWORD))
    except JIRAError as error:
        logger.error("Error", error.status_code, "-", error.text)
        sys.exit(1)
