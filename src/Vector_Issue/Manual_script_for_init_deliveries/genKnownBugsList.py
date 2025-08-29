#!/usr/bin/python
# -*- coding: utf-8 -*-
""" KnownBugs generation tool
Generating the KnownBugs.xlsx file from Vectors IssueReport.xml
"""

import Vector_Issue.MQ as MQ
import datetime
import argparse
import sys
import os
import shutil
import logging


DEFAULT_PATH_KNOWNBUGSLIST_RELATIVE_TO_ISSUE_LIST_XML = "\\..\\.."
version = "0.8"

def main():
    # # Parse the arguments
    # parser = argparse.ArgumentParser()
    # parser.add_argument("xmlfile", help="new Issue Report XML file from Vector")
    # parser.add_argument("-o", "--outfile", help="Output KnownIssues.xlsx file, if not specified default location inside package folder will be used")
    # parser.add_argument("-v", "--version", action='version', version="Known Bug List Generator Version {}".format(version))
    # args = parser.parse_args()
    #
    # # define XML file
    # issueXml = args.xmlfile

    # # define excel file
    # excelFile = args.outfile

    # for debugging
    #issueXml = 'IssueReport_CBD1500052_D01_2017-07-07.xml'
    #issueXml = 'IssueReport_CBD1500057_D01_2017-03-24.xml'
    #issueXml = 'IssueReport_CBD1700556.xml'
    #excelFile = 'KnownBugsList.xlsx'
    
    #issueXml = 'IssueReport_CBD1900137.xml'
    #issueXml = 'IssueReport_CBD1900137_D01_2020-05-20.xml'
    #issueXml = 'IssueReport_CBD1900137_D01_2020-05-29.xml'
    #issueXml = 'IssueReport_CBD1900137_D01_2020-06-16.xml'
    #issueXml = 'IssueReport_CBD1900137_D01_2020-07-14.xml'
    #issueXml = 'IssueReport_CBD1900137_D01_2020-07-20.xml'
    #issueXml = 'IssueReport_CBD1900137_D01_2020-07-23.xml'
    #issueXml = 'IssueReport_CBD1900137_D01_2020-07-30.xml'
    #issueXml = 'IssueReport_CBD1900137_D01_2020-09-28.xml'
    #issueXml = 'IssueReport_CBD1900137_D01_2020-10-07.xml'
    #issueXml = 'IssueReport_CBD1900137_D01_2020-10-21.xml'
    
    #issueXml = 'IssueReport_CBD1900138.xml'
    #issueXml = 'IssueReport_CBD1900138_D01_2020-05-29.xml'
    #issueXml = 'IssueReport_CBD1900138_D01_2020-07-14.xml'
    #issueXml = 'IssueReport_CBD1900138_D01_2020-07-20.xml'
    #issueXml = 'IssueReport_CBD1900138_D01_2020-07-30.xml'
    #issueXml = 'IssueReport_CBD1900138_D01_2020-10-07.xml'
    
    #issueXml = 'IssueReport_CBD2000056.xml'
    #issueXml = 'IssueReport_CBD2000056_D00_2020-05-13.xml'
    #issueXml = 'IssueReport_CBD2000056_D00_2020-05-29.xml'
    #issueXml = 'IssueReport_CBD2000056_D00_2020-06-26.xml'
    #issueXml = 'IssueReport_CBD2000056_D00_2020-07-30.xml'
    #issueXml = 'IssueReport_CBD2000056_D00_2020-10-07.xml'
    
    #issueXml = 'IssueReport_CBD2000374_D00_2020-12-14.xml'
    
    #issueXml = 'FixedIssueReport_CBD1900137.xml'
    #issueXml = 'FixedIssueReport_CBD2000056.xml'
    #issueXml = 'FixedIssueReport_CBD1900138.xml'
    
    #issueXml = 'IssueReport_CBD1900137.xml'
    #issueXml = 'IssueReport_CBD1900137_D01_2020-11-27.xml'
    #issueXml = 'IssueReport_CBD1900137_D01_2020-12-14.xml'
    #issueXml = 'IssueReport_CBD1900137_D01_2021-01-11.xml'
    #issueXml = 'IssueReport_CBD1900137_D01_2021-01-18.xml'
    #issueXml = 'IssueReport_CBD1900137_D01_2021-01-29.xml'
    #issueXml = 'IssueReport_CBD1900137_D01_2021-02-04.xml'
    #issueXml = 'IssueReport_CBD1900137_D01_2021-02-09.xml'
    #issueXml = 'IssueReport_CBD1900137_D01_2020-07-30.xml'
    
    #issueXml = 'IssueReport_CBD2000056.xml'
    #issueXml = 'IssueReport_CBD2000056_D00_2020-05-13.xml'
    #issueXml = 'IssueReport_CBD2000056_D00_2020-05-29.xml'
    #issueXml = 'IssueReport_CBD2000056_D00_2020-06-26.xml'
    #issueXml = 'IssueReport_CBD2000056_D00_2020-07-30.xml'
    #issueXml = 'IssueReport_CBD2000056_D00_2020-10-07.xml'
    #issueXml = 'IssueReport_CBD2000056_D00_2020-11-27.xml'
    #issueXml = 'IssueReport_CBD2000056_D00_2021-01-18.xml'
    #issueXml = 'IssueReport_CBD2000056_D00_2021-01-29.xml'
    
    #issueXml = 'IssueReport_CBD1900138.xml'
    #issueXml = 'IssueReport_CBD1900138_D01_2020-05-29.xml'
    #issueXml = 'IssueReport_CBD1900138_D01_2020-07-14.xml'
    #issueXml = 'IssueReport_CBD1900138_D01_2020-07-20.xml'
    #issueXml = 'IssueReport_CBD1900138_D01_2020-07-30.xml'
    #issueXml = 'IssueReport_CBD1900138_D01_2020-10-07.xml'
    #issueXml = 'IssueReport_CBD1900138_D01_2020-11-27.xml'
    #issueXml = 'IssueReport_CBD1900138_D01_2020-12-14.xml'
    #issueXml = 'IssueReport_CBD1900138_D01_2021-01-29.xml'
    
    #issueXml = 'IssueReport_CBD1900137.xml'
    #issueXml = 'IssueReport_CBD1900138.xml'
    #issueXml = 'IssueReport_CBD2000056.xml'
    
    #issueXml = 'IssueReport_CBD1900137.xml'
    issueXml = 'IssueReport_CBD2200497_NonSecurity.xml'

    print("genKnownBugsList.py --- XML file transferred: %s" % issueXml)

    # Check if file exists
    if not os.path.isfile(issueXml):
        print("%s is not a valid file" % issueXml)
        sys.exit()


    # If option is not set use default location in the package folder
    # if not excelFile:
    #excelFile = os.path.dirname(issueXml) + DEFAULT_PATH_KNOWNBUGSLIST_RELATIVE_TO_ISSUE_LIST_XML + "\\KnownBugsList.xlsx"
    #excelFile = "X:\\ASR_Team\\Vector\\Microsar-Packages\\CBD1900138\\KnownBugsList.xlsx"
    #excelFile = "C:\\temp\\_Vector_Issue_Reports\\Vector_Issue_Reports\\Test\\CBD1900137\\KnownBugsList.xlsx"
    excelFile = "X:\\ASR_Team\\Vector\\Microsar-Packages\\CBD2200497\\KnownBugsList.xlsx"

    print("genKnownBugsList.py --- Issue Xml: ", issueXml)
    print("genKnownBugsList.py --- os.path.dirname(issueXml) :", os.path.dirname(issueXml))
    print("genKnownBugsList.py --- Excel file: ", excelFile)

    # check if file exists
    if os.path.isfile(excelFile):
        # use excel file as it is
        pass
    else:
        shutil.copy2("KnownBugsList_Template.xlsx", excelFile)

    issueObject = MQ.VectorIssuesXml(issueXml, categoryFilter=["apparentIssue", "compilerWarning"])
    newVectorEscanDict = issueObject.getEscanDict()

    excel = MQ.KnownIssuesExcel(excelFile, 0)
    currentEscanExcelDict = excel.getEscanAndDescriptionAsDict()
    # Remove headline in excelsheet
    # escanList.remove("Reference")

    # search Last index in excel sheet
    highestIndex = 0
    for escan in currentEscanExcelDict:
        if currentEscanExcelDict[escan][0] > highestIndex:
            highestIndex = currentEscanExcelDict[escan][0]
    # Set highest Index to the next free line
    highestIndex += 1

    # get report name and release number from issueXml
    #issueXmlFile = issueXml.split('\\')[7]
    #reportName = issueXmlFile.split('.')[0]
    #releaseNumber = issueXmlFile.split('_')[2]
    #reportName = 'IssueReport_CBD1900138'
    #releaseNumber = 'Initial'
    #reportName = issueXml.split('.')[0]
    #releaseNumber = issueXml.split('_')[2]
    reportName = 'IssueReport_CBD2200497_NonSecurity'
    releaseNumber = 'D01'

    # add headings to excel file
    excel.writeHeadingValueToCell(15, 7, 'Report Name')
    excel.writeHeadingValueToCell(16, 7, 'Release Number')
    excel.writeHeadingValueToCell(17, 7, 'Category')
    excel.set_border(15, 7, 16, 7)
    excel.set_border(16, 7, 17, 7)

    # go through new Vector XML entries
    for escan in newVectorEscanDict.keys():
        # check if ESCAN already exists in current excel sheet
        # if not: write it at the end of the excel sheet
        # if yes: write the new content to the line where the old ESCAN was stored
        if escan not in currentEscanExcelDict.keys():
            print("%s not in Excelsheet" % escan)
            logging.info("%s not in Excelsheet" % escan)
            excel.set_border(2, highestIndex, 17, highestIndex)
            if isinstance(excel.sheet.cell(column=2, row=highestIndex-1).value, int) or \
                isinstance(excel.sheet.cell(column=2, row=highestIndex-1).value, float):
                excel.writeValueToCell(2, highestIndex, excel.sheet.cell(column=2, row=highestIndex-1).value + 1)
            else:
                excel.writeValueToCell(2, highestIndex, 1)
            excel.writeValueToCell(3, highestIndex, datetime.date.today())
            excel.writeValueToCell(4, highestIndex, newVectorEscanDict[escan][0])
            excel.writeValueToCell(5, highestIndex, newVectorEscanDict[escan][1])
            excel.writeValueToCell(6, highestIndex, newVectorEscanDict[escan][2])
            excel.writeValueToCell(7, highestIndex, newVectorEscanDict[escan][3])
            excel.writeValueToCell(8, highestIndex, newVectorEscanDict[escan][4])
            excel.writeValueToCell(9, highestIndex, escan)
            excel.writeValueToCell(10, highestIndex, newVectorEscanDict[escan][5])
            excel.writeValueToCell(13, highestIndex, "Not analyzed")
            excel.writeValueToCell(15, highestIndex, reportName)
            excel.writeValueToCell(16, highestIndex, releaseNumber)
            excel.writeValueToCell(17, highestIndex, newVectorEscanDict[escan][6])
            highestIndex += 1
        else:
            print("%s is already in Excelsheet" % escan)
            logging.info("%s is already in Excelsheet" % escan)
            # Number should be already in the sheet --> not updated
            excel.writeValueToCell(3, currentEscanExcelDict[escan][0], datetime.date.today())
            excel.writeValueToCell(4, currentEscanExcelDict[escan][0], newVectorEscanDict[escan][0])
            excel.writeValueToCell(5, currentEscanExcelDict[escan][0], newVectorEscanDict[escan][1])
            excel.writeValueToCell(6, currentEscanExcelDict[escan][0], newVectorEscanDict[escan][2])
            excel.writeValueToCell(7, currentEscanExcelDict[escan][0], newVectorEscanDict[escan][3])
            excel.writeValueToCell(8, currentEscanExcelDict[escan][0], newVectorEscanDict[escan][4])
            excel.writeValueToCell(9, currentEscanExcelDict[escan][0], escan)
            excel.writeValueToCell(10, currentEscanExcelDict[escan][0], newVectorEscanDict[escan][5])
            excel.writeValueToCell(15, currentEscanExcelDict[escan][0], reportName)
            excel.writeValueToCell(16, currentEscanExcelDict[escan][0], releaseNumber)
            excel.writeValueToCell(17, currentEscanExcelDict[escan][0], newVectorEscanDict[escan][6])


if __name__ == "__main__":
    main()
