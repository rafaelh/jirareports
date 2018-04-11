""" Creates an update on Development items for the last week """
# Requires pypiwin32 and jira

import datetime
from getpass import getpass
import win32com.client
from jira import JIRA

#USERNAME = os.getlogin()
USERNAME = 'rhart'
print("JIRA Username: " + USERNAME)
PASSWORD = getpass("JIRA Password: ")

try:
    JQL = JIRA(server=('https://jira.starrez.com'), basic_auth=(USERNAME, PASSWORD))
except:
    print("Connection didn't work. Maybe the username or password is wrong?")

# Get data from JIRA
class PortalX:
    """ Query JIRA for information on PortalX """
    def __init__(self):
        print("Querying JIRA for PortalX issues...")
        self.techdebt = JQL.search_issues('"Epic Link" = PORTALX-1499 and ' \
        + 'resolution = Unresolved', maxResults=200)
        self.bugs = JQL.search_issues('project = PortalX AND resolution = ' \
        + 'Unresolved AND type in (Bug, "Testing Bug", "Sub-Task Bug") and component != "UITest"', maxResults=200)
        self.closedbugs1w = JQL.search_issues('project = PortalX AND resolved ' \
        + '>= -1w AND type in (Bug, "Testing Bug", "Sub-Task Bug") AND ' \
        + 'resolution not in (duplicate, "No Action Required") and component != "UITest"', maxResults=200)
        self.enhancements = JQL.search_issues('project = PortalX AND resolved ' \
        '>= -1w AND type not in (Bug, "Testing Bug", "Sub-Task Bug") AND ' \
        'resolution = Fixed ORDER BY priority DESC', maxResults=200)

class StarRezWeb:
    """ Query JIRA for information on StarRez Web """
    def __init__(self):
        print("Querying JIRA for StarRez Web issues...")
        self.techdebt = JQL.search_issues('"Epic Link" = WEB-7359 and resolution = ' \
        + 'Unresolved', maxResults=200)
        self.bugs = JQL.search_issues('project = WEB AND resolution = Unresolved ' \
        + ' AND type in (Bug, "Testing Bug", "Sub-Task Bug") and component != "UITest"', maxResults=200)
        self.closedbugs1w = JQL.search_issues('project = "StarRez Web" AND ' \
        + 'resolved >= -1w AND type in (Bug, "Testing Bug", "Sub-Task Bug") ' \
        + 'AND resolution not in (duplicate, "No Action Required") and component != "UITest"', maxResults=200)
        self.enhancements = JQL.search_issues('project = "StarRez Web" AND resolved' \
        + ' >= -1w AND type not in (Bug, "Testing Bug", "Sub-Task Bug") AND ' \
        + 'resolution = Fixed ORDER BY priority DESC', maxResults=200)

class Cloud:
    """ Query JIRA for information on Cloud """
    def __init__(self):
        print("Querying JIRA for Cloud issues...")
        self.bugs = JQL.search_issues('project = Cloud AND resolution = Unresolved ' \
        + 'AND type in (Bug, "Testing Bug", "Sub-Task Bug")', maxResults=200)
        self.closedbugs1w = JQL.search_issues('project = Cloud AND ' \
        + 'resolved >= -1w AND type in (Bug, "Testing Bug", "Sub-Task Bug") ' \
        + 'AND resolution not in (duplicate, "No Action Required")', maxResults=200)
        self.enhancements = JQL.search_issues('project = Cloud AND resolved >= -1w ' \
        + 'AND type not in (Bug, "Testing Bug", "Sub-Task Bug") AND ' \
        + 'resolution in (Done, Fixed) ORDER BY priority DESC', maxResults=200)

class Mobile:
    """ Query JIRA for information on Mobile """
    def __init__(self):
        print("Querying JIRA for Mobile issues...")
        self.bugs = JQL.search_issues('project = "Mobile Applications" AND resolution = Unresolved ' \
        + 'AND type in (Bug, "Testing Bug", "Sub-Task Bug")', maxResults=200)
        self.closedbugs1w = JQL.search_issues('project = "Mobile Applications" AND ' \
        + 'resolved >= -1w AND type in (Bug, "Testing Bug", "Sub-Task Bug") ' \
        + 'AND resolution not in (duplicate, "No Action Required")', maxResults=200)
        self.enhancements = JQL.search_issues('project = "Mobile Applications" AND resolved >= -1w ' \
        + 'AND type not in (Bug, "Testing Bug", "Sub-Task Bug") AND ' \
        + 'resolution in (Done, Fixed) ORDER BY priority DESC', maxResults=200)

class Ux:
    """ Query JIRA for information on the UX project """
    def __init__(self):
        print("Querying JIRA for UX issues...")
        self.enhancements = JQL.search_issues("project = UX AND resolved >= -1w AND type not in (Epic)", maxResults=200)

class Integrations:
    """ Query JIRA for information on the Integrations Team """
    def __init__(self):
        print("Querying JIRA for Integration issues...")
        self.enhancements = JQL.search_issues('project = "Custom Development" AND resolved >= -1w ' \
        + 'AND type not in (Bug, "Testing Bug", "Sub-Task Bug", Sub-Project) AND ' \
        + 'resolution in (Done, Fixed) and developer is not EMPTY ORDER BY priority DESC', maxResults=200)

class Techhelp:
    """ Query JIRA for information on Techhelps """
    def __init__(self):
        print("Querying JIRA for Techhelp issues...")
        self.in2weeks = JQL.search_issues('project = tech and createdDate >= -2w', maxResults=200)
        self.in3weeks = JQL.search_issues('project = TECH AND created >= -3w AND ' \
        + 'created <= -1w', maxResults=200)
        if len(self.in2weeks) > len(self.in3weeks):
            self.trend = "up"
        else:
            self.trend = "down"

class FeatureParity:
    """ Query JIRA for information on Feature Parity """
    def __init__(self):
        self.todo = JQL.search_issues('labels = Feature_Parity AND resolution = unresolved ',
                                      maxResults=200)
        self.done = JQL.search_issues('labels = Feature_Parity AND resolution = fixed ',
                                      maxResults=500)

class Documentation:
    """ Query JIRA for information on Doc jobs """
    def __init__(self):
        print("Querying JIRA for Documentation issues...")
        self.newdocs = JQL.search_issues('project = Documentation AND resolved >= -1w AND resolution = Fixed ORDER BY resolutiondate')

PORTALX = PortalX()
WEB = StarRezWeb()
CLOUD = Cloud()
MOBILE = Mobile()
INTEGRATIONS = Integrations()
TECHHELP = Techhelp()
FEATUREPARITY = FeatureParity()
DOCUMENTATION = Documentation()
UX = Ux()

# Create Email Contents
print("Generating Email...")
with open('emailheader.html', 'r') as emailFormat:
    BODY = emailFormat.read().replace('\n', '')

BODY += "<p><br><b>Product Health</b><br>"
BODY += "<br><p>**Insert Table**</p><br>"

BODY += "<p><b>Links</b></p>"
BODY += "<p>Feature Parity: <a href=\"https://jira.starrez.com/issues/?filter=20417\">%s</a> pending, %s done<br>" % (len(FEATUREPARITY.todo), len(FEATUREPARITY.done))
BODY += "Web - <a href=\"https://jira.starrez.com/issues/?filter=19937\">%s</a> open bugs, <a href=\"https://jira.starrez.com/issues/?filter=24217\">%s</a> open Tech Debt issues<br>" % (len(WEB.bugs), len(WEB.techdebt))
BODY += "PortalX - <a href=\"https://jira.starrez.com/issues/?filter=20511\">%s</a> open bugs, <a href=\"https://jira.starrez.com/issues/?filter=24218\">%s</a> open Tech Debt issues<br>" % (len(PORTALX.bugs), len(PORTALX.techdebt))
BODY += "Cloud - <a href=\"https://jira.starrez.com/issues/?filter=23239\">%s</a> open bugs<br>" % len(CLOUD.bugs)
BODY += "StarRez X - <a href=\"https://jira.starrez.com/issues/?filter=24815\">%s</a> open bugs</p>" % len(MOBILE.bugs)

BODY += "<br><p>**Insert Bug Graph**</p><br>"

BODY += "<p><b>Techhelps</b> - %s jobs in the last two weeks, %s from %s at the last check<br>" \
        % (len(TECHHELP.in2weeks), TECHHELP.trend, len(TECHHELP.in3weeks))

BODY += "<br>**Insert Techhelp Chart**</p><br>"

BODY += "<p>Done in the last week:</p><ul>"

BODY += "<li>%s Bugs (" % len(PORTALX.closedbugs1w + WEB.closedbugs1w + CLOUD.closedbugs1w)
if PORTALX.closedbugs1w:
    BODY += "<a href=\"https://jira.starrez.com/issues/?filter=22711\">%s PortalX</a>" % len(PORTALX.closedbugs1w)
if WEB.closedbugs1w:
    BODY += " / <a href=\"https://jira.starrez.com/issues/?filter=22712\">%s Web</a>" % len(WEB.closedbugs1w)
if CLOUD.closedbugs1w:
    BODY += " / <a href=\"https://jira.starrez.com/issues/?filter=24332\">%s Cloud</a>" % len(CLOUD.closedbugs1w)
if MOBILE.closedbugs1w:
    BODY += " / <a href=\"https://jira.starrez.com/issues/?filter=24823\">%s Mobile</a>" % len(MOBILE.closedbugs1w)
BODY += ")</li>"

for issue in UX.enhancements:
    BODY += "<li><a href=\"https://jira.starrez.com/browse/%s\">%s</a> - %s</li>" \
    % (issue, issue, issue.fields.summary)
for issue in MOBILE.enhancements:
    BODY += "<li><a href=\"https://jira.starrez.com/browse/%s\">%s</a> - %s</li>" \
    % (issue, issue, issue.fields.summary)
for issue in PORTALX.enhancements:
    BODY += "<li><a href=\"https://jira.starrez.com/browse/%s\">%s</a> - %s</li>" \
    % (issue, issue, issue.fields.summary)
for issue in WEB.enhancements:
    BODY += "<li><a href=\"https://jira.starrez.com/browse/%s\">%s</a> - %s</li>" \
    % (issue, issue, issue.fields.summary)
for issue in CLOUD.enhancements:
    BODY += "<li><a href=\"https://jira.starrez.com/browse/%s\">%s</a> - %s</li>" \
    % (issue, issue, issue.fields.summary)
for issue in INTEGRATIONS.enhancements:
    BODY += "<li><a href=\"https://jira.starrez.com/browse/%s\">%s</a> - %s</li>" \
    % (issue, issue, issue.fields.summary)
BODY += "</ul>"


if DOCUMENTATION.newdocs:
    BODY += "<p>New Documents:</p><ul>"
    for issue in DOCUMENTATION.newdocs:
        BODY += "<li><a href=\"https://jira.starrez.com/browse/%s\">%s</a> - %s</li>" \
        % (issue, issue, issue.fields.summary)
BODY += "</ul>"

BODY += "<p>Thanks,<br><br>Rafe<br></p></body></html>"


def createemail(emailbody):
    """ Sent Email Contents to Outlook """
    olmailitem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newmail = obj.CreateItem(olmailitem)
    today = datetime.date.today()
    newmail.Subject = today.strftime("Development Update - %d %b %Y")
    newmail.HTMLBody = emailbody
    newmail.display()

createemail(BODY)
