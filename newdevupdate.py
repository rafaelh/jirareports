""" Creates an update on Development items for the last week """

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
class Enhancements:
    """ Query JIRA for information on enhancements for each project """
    def __init__(self):
        print("Querying JIRA for PortalX enhancements...")
        self.portalx = JQL.search_issues('project = PortalX AND resolved ' \
        '>= -1w AND type not in (Bug, "Testing Bug", "Sub-Task Bug") AND ' \
        'resolution in (Fixed, Done) ORDER BY priority DESC', maxResults=200)

        print("Querying JIRA for Web enhancements...")
        self.web = JQL.search_issues('project = "StarRez Web" AND resolved' \
        + ' >= -1w AND type not in (Bug, "Testing Bug", "Sub-Task Bug") AND ' \
        + 'resolution in (Fixed, Done) ORDER BY priority DESC', maxResults=200)

        print("Querying JIRA for Deployment enhancements...")
        self.cloud = JQL.search_issues('project = Cloud AND resolved >= -1w ' \
        + 'AND type not in (Bug, "Testing Bug", "Sub-Task Bug") AND ' \
        + 'resolution in (Done, Fixed) ORDER BY priority DESC', maxResults=200)

        print("Querying JIRA for Mobile enhancements...")
        self.mobile = JQL.search_issues('project = "Mobile Applications" AND resolved >= -1w ' \
        + 'AND type not in (Bug, "Testing Bug", "Sub-Task Bug") AND ' \
        + 'resolution in (Done, Fixed) ORDER BY priority DESC', maxResults=200)

        print("Querying JIRA for UX enhancements...")
        self.ux = JQL.search_issues('project = UX AND resolved >= -1w AND type ' \
        + 'not in (Epic) and resolution not in (\"Couldn\'t Solve\", \"No Action ' \
        + 'Required\", \"Won\'t Do\")', maxResults=200)

        print("Querying JIRA for CD issues...")
        self.cd = JQL.search_issues('project = "Custom Development" AND resolved >= -1w ' \
        + 'AND type not in (Bug, "Testing Bug", "Sub-Task Bug", Sub-Project) AND ' \
        + 'resolution in (Done, Fixed) and developer is not EMPTY ORDER BY priority DESC', \
        maxResults=200)

        print("Querying JIRA for Cloud Adoption enhancements...")
        self.cloudadoption = JQL.search_issues('project = "Cloud Adoption" AND resolved >= -1w ' \
        + 'AND type not in (Bug, "Testing Bug", "Sub-Task Bug") AND ' \
        + 'resolution in (Done, Fixed) ORDER BY priority DESC', maxResults=200)

        print("Querying JIRA for Marketplace enhancements...")
        self.marketplace = JQL.search_issues('project = Marketplace AND resolved >= -1w ' \
        + 'AND type not in (Bug, "Testing Bug", "Sub-Task Bug") AND ' \
        + 'resolution in (Done, Fixed) ORDER BY priority DESC', maxResults=200)

        print("Querying JIRA for DevOps enhancements...")
        self.devops = JQL.search_issues('project = "Development Ops" AND resolved >= -1w ' \
        + 'AND type not in (Bug, "Testing Bug", "Sub-Task Bug") AND ' \
        + 'resolution in (Done, Fixed) ORDER BY priority DESC', maxResults=200)

        print("Querying JIRA for Enhance enhancements...")
        self.enhance = JQL.search_issues('project = Enhance AND resolved >= -1w ' \
        + 'AND type not in (Bug, "Testing Bug", "Sub-Task Bug") AND ' \
        + 'resolution in (Done, Fixed) ORDER BY priority DESC', maxResults=200)

        print("Querying JIRA for Explore enhancements...")
        self.explore = JQL.search_issues('project = Explore AND resolved >= -1w ' \
        + 'AND type not in (Bug, "Testing Bug", "Sub-Task Bug") AND ' \
        + 'resolution in (Done, Fixed) ORDER BY priority DESC', maxResults=200)

        print("Querying JIRA for ValueAdd enhancements...")
        self.valueadds = JQL.search_issues('project = "Value Adds" AND resolved >= -1w ' \
        + 'AND type not in (Bug, "Testing Bug", "Sub-Task Bug") AND ' \
        + 'resolution in (Done, Fixed) ORDER BY priority DESC', maxResults=200)

        print("Querying JIRA for New Markets enhancements...")
        self.newmarkets = JQL.search_issues('project = "New Markets" AND resolved >= -1w ' \
        + 'AND type not in (Bug, "Testing Bug", "Sub-Task Bug") AND ' \
        + 'resolution in (Done, Fixed) ORDER BY priority DESC', maxResults=200)


class Bugs:
    """ Query JIRA for information on Bugs in each project """
    def __init__(self):
        print("Querying JIRA for PortalX Bugs...")
        self.portalx = JQL.search_issues('project = PortalX AND resolution = ' \
        + 'Unresolved AND type in (Bug, "Testing Bug", "Sub-Task Bug") and component != "UITest"', maxResults=200)
        self.portalxclosedlastweek = JQL.search_issues('project = PortalX AND resolved ' \
        + '>= -1w AND type in (Bug, "Testing Bug", "Sub-Task Bug") AND ' \
        + 'resolution not in (duplicate, "No Action Required", "Won\'t Do")', maxResults=200)

        print("Querying JIRA for Web Bugs...")
        self.web = JQL.search_issues('project = WEB AND resolution = Unresolved ' \
        + ' AND type in (Bug, "Testing Bug", "Sub-Task Bug") and component != "UITest"', maxResults=200)
        self.webclosedlastweek = JQL.search_issues('project = "StarRez Web" AND ' \
        + 'resolved >= -1w AND type in (Bug, "Testing Bug", "Sub-Task Bug") ' \
        + 'AND resolution not in (duplicate, "No Action Required", "Won\'t Do")', maxResults=200)

        print("Querying JIRA for Cloud Bugs...")
        self.cloud = JQL.search_issues('project = Cloud AND resolution = Unresolved ' \
        + 'AND type in (Bug, "Testing Bug", "Sub-Task Bug")', maxResults=200)
        self.cloudclosedlastweek = JQL.search_issues('project = Cloud AND ' \
        + 'resolved >= -1w AND type in (Bug, "Testing Bug", "Sub-Task Bug") ' \
        + 'AND resolution not in (duplicate, "No Action Required", "Won\'t Do")', maxResults=200)
        self.cloudfilter = "https://jira.starrez.com/issues/?filter=23239"

        print("Querying JIRA for Mobile Bugs...")
        self.mobile = JQL.search_issues('project = "Mobile Applications" AND resolution = Unresolved ' \
        + 'AND type in (Bug, "Testing Bug", "Sub-Task Bug")', maxResults=200)
        self.mobileclosedlastweek = JQL.search_issues('project = "Mobile Applications" AND ' \
        + 'resolved >= -1w AND type in (Bug, "Testing Bug", "Sub-Task Bug") ' \
        + 'AND resolution not in (duplicate, "No Action Required", "Won\'t Do")', maxResults=200)
        self.mobilefilter = "https://jira.starrez.com/issues/?filter=24815"

        print("Querying JIRA for Cloud Adoption Bugs...")
        self.cloudadoption = JQL.search_issues('project = "Cloud Adoption" AND resolution = Unresolved ' \
        + 'AND type in (Bug, "Testing Bug", "Sub-Task Bug")', maxResults=200)
        self.cloudadoptionclosedlastweek = JQL.search_issues('project = "Cloud Adoption" AND ' \
        + 'resolved >= -1w AND type in (Bug, "Testing Bug", "Sub-Task Bug") ' \
        + 'AND resolution not in (duplicate, "No Action Required", "Won\'t Do")', maxResults=200)
        self.cloudadoptionfilter = "https://jira.starrez.com/issues/?filter=26355"

        print("Querying JIRA for Marketplace Bugs...")
        self.marketplace = JQL.search_issues('project = Marketplace AND resolution = Unresolved ' \
        + 'AND type in (Bug, "Testing Bug", "Sub-Task Bug")', maxResults=200)
        self.marketplaceclosedlastweek = JQL.search_issues('project = Marketplace AND ' \
        + 'resolved >= -1w AND type in (Bug, "Testing Bug", "Sub-Task Bug") ' \
        + 'AND resolution not in (duplicate, "No Action Required", "Won\'t Do")', maxResults=200)
        self.marketplacefilter = "https://jira.starrez.com/issues/?filter=26356"

        print("Querying JIRA for DevOps Bugs...")
        self.devops = JQL.search_issues('project = "Development Ops" AND resolution = Unresolved ' \
        + 'AND type in (Bug, "Testing Bug", "Sub-Task Bug")', maxResults=200)
        self.devopsclosedlastweek = JQL.search_issues('project = "Development Ops" AND ' \
        + 'resolved >= -1w AND type in (Bug, "Testing Bug", "Sub-Task Bug") ' \
        + 'AND resolution not in (duplicate, "No Action Required", "Won\'t Do")', maxResults=200)
        self.devopsfilter = "https://jira.starrez.com/issues/?filter=26358"

        print("Querying JIRA for Enhance Bugs...")
        self.enhance = JQL.search_issues('project = Enhance AND resolution = Unresolved ' \
        + 'AND type in (Bug, "Testing Bug", "Sub-Task Bug")', maxResults=200)
        self.enhanceclosedlastweek = JQL.search_issues('project = Enhance AND ' \
        + 'resolved >= -1w AND type in (Bug, "Testing Bug", "Sub-Task Bug") ' \
        + 'AND resolution not in (duplicate, "No Action Required", "Won\'t Do")', maxResults=200)
        self.enhancefilter = "https://jira.starrez.com/issues/?filter=26359"

        print("Querying JIRA for Explore Bugs...")
        self.explore = JQL.search_issues('project = Explore AND resolution = Unresolved ' \
        + 'AND type in (Bug, "Testing Bug", "Sub-Task Bug")', maxResults=200)
        self.exploreclosedlastweek = JQL.search_issues('project = Explore AND ' \
        + 'resolved >= -1w AND type in (Bug, "Testing Bug", "Sub-Task Bug") ' \
        + 'AND resolution not in (duplicate, "No Action Required", "Won\'t Do")', maxResults=200)
        self.explorefilter = "https://jira.starrez.com/issues/?filter=26362"

        print("Querying JIRA for Value Add Bugs...")
        self.valueadds = JQL.search_issues('project = "Value Adds" AND resolution = Unresolved ' \
        + 'AND type in (Bug, "Testing Bug", "Sub-Task Bug")', maxResults=200)
        self.valueaddsclosedlastweek = JQL.search_issues('project = "Value Adds" AND ' \
        + 'resolved >= -1w AND type in (Bug, "Testing Bug", "Sub-Task Bug") ' \
        + 'AND resolution not in (duplicate, "No Action Required", "Won\'t Do")', maxResults=200)
        self.valueaddsfilter = "https://jira.starrez.com/issues/?filter=26364"

        print("Querying JIRA for New Markets Bugs...")
        self.newmarkets = JQL.search_issues('project = "New Markets" AND resolution = Unresolved ' \
        + 'AND type in (Bug, "Testing Bug", "Sub-Task Bug")', maxResults=200)
        self.newmarketsclosedlastweek = JQL.search_issues('project = "New Markets" AND ' \
        + 'resolved >= -1w AND type in (Bug, "Testing Bug", "Sub-Task Bug") ' \
        + 'AND resolution not in (duplicate, "No Action Required", "Won\'t Do")', maxResults=200)
        self.newmarketsfilter = "https://jira.starrez.com/issues/?filter=26366"

        self.total = (len(self.portalx) + len(self.web) + len(self.cloud) + len(self.mobile) \
                     + len(self.cloudadoption) + len(self.marketplace) + len(self.devops) \
                     + len(self.enhance) + len(self.explore) + len(self.valueadds) \
                     + len(self.newmarkets))

class TechDebt:
    """ Query JIRA for information on Tech Debt issues """
    def __init__(self):
        print("Querying JIRA for PortalX Tech Debt issues...")
        self.portalx = JQL.search_issues('"Epic Link" = PORTALX-1499 and ' \
        + 'resolution = Unresolved', maxResults=200)

        print("Querying JIRA for Web Tech Debt issues...")
        self.web = JQL.search_issues('"Epic Link" = WEB-7359 and resolution = ' \
        + 'Unresolved', maxResults=200)

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

class Documentation:
    """ Query JIRA for information on Doc jobs """
    def __init__(self):
        print("Querying JIRA for Documentation issues...")
        self.newdocs = JQL.search_issues("project = Documentation AND resolved >= -1w " \
        + "AND resolution = Fixed ORDER BY resolutiondate", maxResults=200)


# Create an email using the assembled information
def createemail(emailbody):
    """ Sent Email Contents to Outlook """
    olmailitem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newmail = obj.CreateItem(olmailitem)
    today = datetime.date.today()
    newmail.Subject = today.strftime("Development Update - %d %b %Y")
    newmail.HTMLBody = emailbody
    newmail.display()


if __name__ == "__main__":
    ENHANCEMENTS = Enhancements()
    BUGS = Bugs()
    TECHHELP = Techhelp()
    TECHDEBT = TechDebt()
    DOCUMENTATION = Documentation()


    # Create Email Contents
    print("Generating Email...")
    with open('emailheader.html', 'r') as emailFormat:
        BODY = emailFormat.read().replace('\n', '')

    BODY += "<p><br><b>Product Health</b><br>"
    BODY += "<br>**Insert Table**</p><br>"

    # TODO: create a function that takes bugs, filter, techdebt, etc as parameters, and handles
    # pluralization

    BODY += "<p><b>Links:</b><ul>"
    BODY += "<li><b>Total Bugs:</b> <a href=\"https://jira.starrez.com/issues/?filter=26367\">%s</a></li>" % BUGS.total
    if BUGS.web or TECHDEBT.web:
        BODY += "<li>Web - <a href=\"https://jira.starrez.com/issues/?filter=19937\">%s</a> open bugs, " % len(BUGS.web)
        BODY += "<a href=\"https://jira.starrez.com/issues/?filter=24217\">%s</a> open Tech Debt issues</li>" % len(TECHDEBT.web)
    if BUGS.portalx or TECHDEBT.portalx:
        BODY += "<li>PortalX - <a href=\"https://jira.starrez.com/issues/?filter=20511\">%s</a> open bugs, " % len(BUGS.portalx)
        BODY += "<a href=\"https://jira.starrez.com/issues/?filter=24218\">%s</a> open Tech Debt issues</li>" % len(TECHDEBT.portalx)
    if BUGS.cloud:
        BODY += "<li>Deployment - <a href=\"%s\">%s</a> open bugs</li>" % \
        (BUGS.cloudfilter, len(BUGS.cloud))
    if BUGS.mobile:
        BODY += "<li>StarRez X - <a href=\"%s\">%s</a> open bugs</li>" % \
        (BUGS.mobilefilter, len(BUGS.mobile))
    if BUGS.cloudadoption:
        BODY += "<li>Cloud Adoption - <a href=\"%s\">%s</a> open bugs</li>" % \
        (BUGS.cloudadoptionfilter, len(BUGS.cloudadoption))
    if BUGS.marketplace:
        BODY += "<li>Marketplace - <a href=\"%s\">%s</a> open bugs</li>" % \
        (BUGS.marketplacefilter, len(BUGS.marketplace))
    if BUGS.devops:
        BODY += "<li>DevOps - <a href=\"%s\">%s</a> open bugs</li>" % \
        (BUGS.devopsfilter, len(BUGS.devops))
    if BUGS.enhance:
        BODY += "<li>Enhance - <a href=\"%s\">%s</a> open bugs</li>" % \
        (BUGS.enhancefilter, len(BUGS.enhance))
    if BUGS.explore:
        BODY += "<li>Explore - <a href=\"%s\">%s</a> open bugs</li>" % \
        (BUGS.explorefilter, len(BUGS.explore))
    if BUGS.valueadds:
        BODY += "<li>Value Adds - <a href=\"%s\">%s</a> open bugs</li>" % \
        (BUGS.valueaddsfilter, len(BUGS.valueadds))
    if BUGS.newmarkets:
        BODY += "<li>New Markets - <a href=\"%s\">%s</a> open bugs</li>" % \
        (BUGS.newmarketsfilter, len(BUGS.newmarkets))
    BODY += "</p></ul>"

    BODY += "<p>**Insert Bug Graph**</p>"

    BODY += "<br><p><b>Techhelps</b> - %s jobs in the last two weeks, %s from %s at the last check</p>" % (len(TECHHELP.in2weeks), TECHHELP.trend, len(TECHHELP.in3weeks))

    BODY += "<br><p>Done in the last week:<ul>"

    # Show bugs closed in the last week
    BODY += "<li>%s Bugs (" % len(BUGS.portalxclosedlastweek + BUGS.webclosedlastweek \
    + BUGS.cloudclosedlastweek + BUGS.cloudadoptionclosedlastweek + BUGS.marketplaceclosedlastweek \
    + BUGS.devopsclosedlastweek + BUGS.enhanceclosedlastweek + BUGS.exploreclosedlastweek \
    + BUGS.valueaddsclosedlastweek + BUGS.newmarketsclosedlastweek)

    if BUGS.portalxclosedlastweek:
        BODY += "<a href=\"https://jira.starrez.com/issues/?filter=22711\">%s PortalX</a>" % len(BUGS.portalxclosedlastweek)
    if BUGS.webclosedlastweek:
        BODY += " / <a href=\"https://jira.starrez.com/issues/?filter=22712\">%s Web</a>" % len(BUGS.webclosedlastweek)
    if BUGS.cloudclosedlastweek:
        BODY += " / <a href=\"https://jira.starrez.com/issues/?filter=24332\">%s Cloud</a>" % len(BUGS.cloudclosedlastweek)
    if BUGS.mobileclosedlastweek:
        BODY += " / <a href=\"https://jira.starrez.com/issues/?filter=24823\">%s Mobile</a>" % len(BUGS.mobileclosedlastweek)
    if BUGS.cloudadoptionclosedlastweek:
        BODY += " / <a href=\"https://jira.starrez.com/issues/?filter=26352\">%s Cloud Adoption</a>" % len(BUGS.cloudadoptionclosedlastweek)
    if BUGS.marketplaceclosedlastweek:
        BODY += " / <a href=\"https://jira.starrez.com/issues/?filter=26354\">%s Marketplace</a>" % len(BUGS.marketplaceclosedlastweek)
    if BUGS.devopsclosedlastweek:
        BODY += " / <a href=\"https://jira.starrez.com/issues/?filter=26357\">%s DevOps</a>" % len(BUGS.devopsclosedlastweek)
    if BUGS.enhanceclosedlastweek:
        BODY += " / <a href=\"https://jira.starrez.com/issues/?filter=26360\">%s Enhance</a>" % len(BUGS.enhanceclosedlastweek)
    if BUGS.exploreclosedlastweek:
        BODY += " / <a href=\"https://jira.starrez.com/issues/?filter=26361\">%s Explore</a>" % len(BUGS.exploreclosedlastweek)
    if BUGS.valueaddsclosedlastweek:
        BODY += " / <a href=\"https://jira.starrez.com/issues/?filter=26363\">%s Value Adds</a>" % len(BUGS.valueaddsclosedlastweek)
    if BUGS.newmarketsclosedlastweek:
        BODY += " / <a href=\"https://jira.starrez.com/issues/?filter=26365\">%s New Markets</a>" % len(BUGS.newmarketsclosedlastweek)
    BODY += ")</li></p>"


    # Show Enhancements for each project
    for issue in ENHANCEMENTS.ux:
        BODY += "<li><a href=\"https://jira.starrez.com/browse/%s\">%s</a> - %s</li>" \
        % (issue, issue, issue.fields.summary)
    for issue in ENHANCEMENTS.mobile:
        BODY += "<li><a href=\"https://jira.starrez.com/browse/%s\">%s</a> - %s</li>" \
        % (issue, issue, issue.fields.summary)
    for issue in ENHANCEMENTS.portalx:
        BODY += "<li><a href=\"https://jira.starrez.com/browse/%s\">%s</a> - %s</li>" \
        % (issue, issue, issue.fields.summary)
    for issue in ENHANCEMENTS.web:
        BODY += "<li><a href=\"https://jira.starrez.com/browse/%s\">%s</a> - %s</li>" \
        % (issue, issue, issue.fields.summary)
    for issue in ENHANCEMENTS.cloud:
        BODY += "<li><a href=\"https://jira.starrez.com/browse/%s\">%s</a> - %s</li>" \
        % (issue, issue, issue.fields.summary)
    for issue in ENHANCEMENTS.cd:
        BODY += "<li><a href=\"https://jira.starrez.com/browse/%s\">%s</a> - %s</li>" \
        % (issue, issue, issue.fields.summary)
    for issue in ENHANCEMENTS.cloudadoption:
        BODY += "<li><a href=\"https://jira.starrez.com/browse/%s\">%s</a> - %s</li>" \
        % (issue, issue, issue.fields.summary)
    for issue in ENHANCEMENTS.marketplace:
        BODY += "<li><a href=\"https://jira.starrez.com/browse/%s\">%s</a> - %s</li>" \
        % (issue, issue, issue.fields.summary)
    for issue in ENHANCEMENTS.devops:
        BODY += "<li><a href=\"https://jira.starrez.com/browse/%s\">%s</a> - %s</li>" \
        % (issue, issue, issue.fields.summary)
    for issue in ENHANCEMENTS.enhance:
        BODY += "<li><a href=\"https://jira.starrez.com/browse/%s\">%s</a> - %s</li>" \
        % (issue, issue, issue.fields.summary)
    for issue in ENHANCEMENTS.explore:
        BODY += "<li><a href=\"https://jira.starrez.com/browse/%s\">%s</a> - %s</li>" \
        % (issue, issue, issue.fields.summary)
    for issue in ENHANCEMENTS.valueadds:
        BODY += "<li><a href=\"https://jira.starrez.com/browse/%s\">%s</a> - %s</li>" \
        % (issue, issue, issue.fields.summary)
    for issue in ENHANCEMENTS.newmarkets:
        BODY += "<li><a href=\"https://jira.starrez.com/browse/%s\">%s</a> - %s</li>" \
        % (issue, issue, issue.fields.summary)
    BODY += "</ul>"


    # Show any Documentation jobs that have been completed in the last week
    if DOCUMENTATION.newdocs:
        BODY += "<p>New Documents:<ul>"
        for issue in DOCUMENTATION.newdocs:
            BODY += "<li><a href=\"https://jira.starrez.com/browse/%s\">%s</a> - %s</li>" \
            % (issue, issue, issue.fields.summary)
    BODY += "</ul></p>"

    BODY += "<p>Thanks,<br><br>Rafe<br></p></body></html>"

    createemail(BODY)
