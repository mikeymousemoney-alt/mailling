import xmltodict
from datetime import datetime
from Vector_Issue.utils import test_log

class VectorIssuesXml:
    """ Vector Issue XML client class """

    def __init__(self, filepath, file, fileIsAPath, categoryFilter=[]):
        test_log("--> vectorIssuesXml")
        self.filepath = filepath
        self.file = file
        self.categoryFilter = categoryFilter
        self.numIssuesIgnored = 0
        self.totalIssues = 0  # Total number of issues
        self.currentIssueIndex = 0  # Index of the current issue being processed
        self.issues = []  # List to store all issues
        self.escanDict = {}

        #check if an xml file is send or "just" a path to it
        if fileIsAPath:        
            with open(self.filepath, encoding='utf-8', errors='ignore') as fd:
                self.doc = xmltodict.parse(fd.read())
        else:
            print("Parsing:",file)
            #convert from Bytes to String
            self.doc = xmltodict.parse(file.decode("utf-8"))
        
        self.loadIssues()
        test_log("vectorIssuesXml <--")

    def loadIssues(self):
        """ Load all issues into a list and count them. """
        raw_issues = self.doc['issueReport']['issues']['issue']
        if not isinstance(raw_issues, list):
            raw_issues = [raw_issues]  # Wrap single issue in a list for consistent processing
        
        self.issues = raw_issues
        self.totalIssues = len(self.issues)  # Total issues count

    def processNextIssue(self):
        """ Process the next issue in the list. """
        if self.currentIssueIndex >= self.totalIssues:
            print("All issues have been processed.")
            return None

        issue = self.issues[self.currentIssueIndex]
        self.currentIssueIndex += 1  # Move to the next issue

        if issue['@category'] in self.categoryFilter:
            self.numIssuesIgnored += 1
            return None

        self.escanDict = {}

        # Prepare issue data
        report_identifier = self.doc['issueReport']['reportData']['reportIdentifier']
        report_time = report_identifier.split('-')[-1]
        creation_date = self.doc['issueReport']['reportData']['reportCreationDate']
        combined_datetime_str = f"{creation_date} {report_time}"
        try:
            iso_combined_datetime = datetime.strptime(combined_datetime_str, "%Y-%m-%d %H:%M:%S").isoformat() + "Z"
        except Exception as e:
            print(f"Error parsing date/time: {e}")
            iso_combined_datetime = f"{creation_date}T{report_time}Z"

        resolution_description = issue.get('resolutionDescription', '')

        identifier = issue['identifier']

        issue_data = (
            issue['componentShortName'],
            issue['firstAffectedVersion'],
            issue['versionsFixed'],
            issue['headline'],
            issue['problemDescription'],
            resolution_description,
            iso_combined_datetime,
            iso_combined_datetime,
            issue['@category']
        )

        self.escanDict[identifier] = issue_data

        return self.escanDict

    def getEscanDict(self):
        """ Get the current issue dictionary. """
        return self.escanDict

    def getTotalIssues(self):
        """ Return the total number of issues. """
        return self.totalIssues

    def getCurrentIssueIndex(self):
        """ Return the current index of the issue being processed. """
        return self.currentIssueIndex
