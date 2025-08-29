import Vector_Issue.genKnownBugsList as genKnownBugsList

c_pathIssueReports = 'C:\\temp\\Vector_issue_mails\\Issue_Reports_Test\\%s\\%s\\IssueReports'
subjectCBDNumber = 'CBD1900137'
version = 'D00'
xmlFile = 'IssueReport_CBD1900137_D00_2020-02-21.xml'

xmlFileDir = c_pathIssueReports % (subjectCBDNumber, version) + '\\' + xmlFile
print("XML file directory: ", xmlFileDir)

genKnownBugsList.main(xmlFileDir)