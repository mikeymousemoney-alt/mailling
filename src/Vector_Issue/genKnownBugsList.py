#!/usr/bin/python
# -*- coding: utf-8 -*-
""" KnownBugs generation tool
Generating the KnownBugs.xlsx file from Vectors IssueReport.xml
And sending the information to the VVM
"""
import re
import uuid
import Vector_Issue.MQ as MQ
import datetime
import argparse
import sys
import os
import shutil
import logging
import requests
import json
import pandas as pd
from Vector_Issue.utils import test_log, get_test_mode # import test_log function

# Path to the KnownBugsList.xlsx template relative to the issueXml file
# Path to the KnownBugsList.xlsx template relative to the issueXml file
PATH_KNOWNBUGSLIST_FROM_ISSUE_LIST_XML = "\\..\\.."
version = "0.8"
VVM_API_URL = None
escan_already_processed = False
exelfile_BugsList = ""

def validate_file(issueXml):
    # Check if file exists
    if not os.path.isfile(issueXml):
        print("%s is not a valid file" % issueXml)
        logging.info("%s is not a valid file" % issueXml)
        return False
    else:
        print("%s is a valid file" % issueXml)
        logging.info("%s is a valid file" % issueXml)
        return True

def copy_template_excel(issueXml):
    # Copy the Excel template if the target file does not exist.
    global exelfile_BugsList
    excelFile = os.path.dirname(issueXml) + PATH_KNOWNBUGSLIST_FROM_ISSUE_LIST_XML + "\\KnownBugsList.xlsx"
    exelfile_BugsList = excelFile

    if not os.path.isfile(excelFile):
        try:
            shutil.copy2("src/Vector_Issue/KnownBugsList_Template.xlsx", excelFile)
        except:
            shutil.copy2("Vector_Issue/KnownBugsList_Template.xlsx", excelFile)

    return excelFile

def parse_issue_xml(issueXml, xmlFile, fileIsAPath):
    #Parse the XML file and return the extracted data.
    global issueObject
    try:
        issueObject = MQ.VectorIssuesXml(issueXml, xmlFile, fileIsAPath, categoryFilter=["apparentIssue", "compilerWarning"])
    except Exception as e:
        print("XML file is not valid!")
        logging.error(f"XML file is not valid! Error: {e}")
        raise
    return issueObject.getTotalIssues()

def get_report_info(issueXml):
    #Extract report name, release number, and CBD number from the issueXml filename.
    issueXmlFile = os.path.basename(issueXml)
    reportName = issueXmlFile.split('.')[0]
    
    try:
        if "SecurityRelated" in reportName:
            releaseNumber = issueXmlFile.split('_')[3]
            cbdNumber = issueXmlFile.split('_')[2]
        else:
            releaseNumber = issueXmlFile.split('_')[2]
            cbdNumber = issueXmlFile.split('_')[1]
    except:
        print("Error: Could not extract report name, release number, and CBD number from the issueXml filename.")
        logging.error("Error: Could not extract report name, release number, and CBD number from the issueXml filename.")
        sys.exit()

    return reportName, releaseNumber, cbdNumber

def is_old_cbd(cbdNumber, cbds):
    #Check if the given CBD number is in the list of old CBDs.
    return cbdNumber in cbds




        #if escan not in currentEscanExcelDict:
        #else:
        #    print(f"{escan} already in Excelsheet KnownBugsList.xlsx")
        #    logging.info(f"{escan} already in Excelsheet KnownBugsList.xlsx")
        #    escan_already_processed = True

# Function to write data to the Excel file
def write_to_excel(excel, reportName, releaseNumber, escanDict, currentEscanExcelDict, highestIndex):
    global escan_already_processed
    #excel = MQ.KnownIssuesExcel(excelFile, cbdOld)
    #Write the data to the Excel file.
    for escan, details in escanDict.items():
        if escan_already_processed:
            print(f"{escan} already in Excelsheet KnownBugsList.xlsx")
            logging.info(f"{escan} already in Excelsheet KnownBugsList.xlsx")
        
        else:
            print(f"{escan} not in Excelsheet")
            logging.info(f"{escan} not in Excelsheet")

            excel.set_border(2, highestIndex, 17, highestIndex)

            if isinstance(excel.sheet.cell(column=2, row=highestIndex-1).value, (int, float)):
                excel.writeValueToCell(2, highestIndex, excel.sheet.cell(column=2, row=highestIndex-1).value + 1)
            else:
                excel.writeValueToCell(2, highestIndex, 1)

            excel.writeValueToCell(3, highestIndex, datetime.date.today())
            excel.writeValueToCell(4, highestIndex, details[0])
            excel.writeValueToCell(5, highestIndex, details[1])
            excel.writeValueToCell(6, highestIndex, details[2])
            excel.writeValueToCell(7, highestIndex, details[3])
            excel.writeValueToCell(8, highestIndex, details[4])
            excel.writeValueToCell(9, highestIndex, escan)
            excel.writeValueToCell(10, highestIndex, details[5])
            excel.writeValueToCell(15, highestIndex, reportName)
            excel.writeValueToCell(16, highestIndex, releaseNumber)

            highestIndex += 1
    
    return highestIndex


def prepare_api_payload(escanDict, cbdNumber, autosar, subject):
    #Prepare the payload for the VVM API.
    
    test_log("--> prepare_api_payload")
    processed_status = {}
    for escan in escanDict.keys():
        processed_status[escan] = check_escan_in_excel(escan, exelfile_BugsList)
        
    logging.debug("<-- genKnownBugsList.check_escan_in_excel")
    test_log("check escan in excel complete")
    payload = []
    for escan, details in escanDict.items():
        test_log("start for loop " + escan)
        escan_already_processed = processed_status.get(escan, False)


        if escan_already_processed:
            test_log(f"{escan} already processed")
            logging.debug(f"{escan} already processed")
            project_number = get_project_number(cbdNumber, autosar)
            
            if project_number is None:
                
                details = list(details)
                details[4] = f"{escan} \nAnalysis Done in {exelfile_BugsList}, But no project number could be found \n{details[4]} \n"
                details = tuple(details)
                logging.debug("no project number found")
                
            else:
                test_log("project number found " + project_number)
                logging.debug(f"project number found {project_number}")
                if isinstance(details, tuple):
                    details = list(details)
                details[4] = escan + " \nAnalysis Done in " + exelfile_BugsList + " for project " + project_number + " \n " + details[4] + "\n"
            details = tuple(details)

        try:                #  Try to convert the escan string to a UUID
            escan1 =str(uuid.UUID(escan))
        except ValueError:
            escan1 = str(uuid.uuid5(uuid.NAMESPACE_DNS, escan))
        test_log("escan1 prepared " + escan1)
        logging.debug("escan uuid (escan1) prepared " + escan1)
        test_log(subject)
        details = list(details)
        details[4] = details[4] + "\n\n E-Mail subject: " + subject
        details = tuple(details)
        test_log("details[4] prepared ")
        issue_data = {                          #  Create a dictionary to store the issue data
            "external_vulnerability_id": escan,
            "vulnerability_type": "SW",
            "source": "PSIRT-Mailbox",
            "publication_created": details[6],
            "publication_updated": details[6],
            "publication_updated": details[6],
            "description":  re.sub(r'([a-z])\n([a-zA-Z])', r'\1 \2',details[4]),
            "time_created": details[7],
            "affected_libraries": [
                {
                    "id": escan1,
                    "name": details[0],
                    "vendor": "Vector",
                    "versions": [
                        {
                            "start_version": details[1],
                            "start_include": True,
                            "end_version": "99.99.99",
                            "end_include": False
                        }
                    ],
                    "vulnerabilityId": escan1
                }
            ],
            "is_confidential": False,
            "analysis": "",
            "mitigation_plan": ""
        }
        payload.append(issue_data)
    test_log(f"Payload prepared for VVM API: {json.dumps(payload, indent=2)}")
    logging.debug(f"Payload prepared for VVM API: {json.dumps(payload, indent=2)}")
    
    return payload
    
def send_data_to_api(payload, token):
    # Send the prepared payload to the VVM API with authorization
    if not token:
        print("No access token. Cannot send data to API.")
        logging.error("No access token given. Cannot send data to API.")
        return
    
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {token}'
    }
    
    response = requests.post(VVM_API_URL, headers=headers, data=json.dumps(payload), verify=False)
    test_log("Headers sent to VVM API:")
    test_log(json.dumps(headers, indent=4))
    logging.debug(f"Headers sent to VVM API: {json.dumps(headers, indent=4)}")
    #test_log("Payload being sent to VVM API:")
    #test_log(json.dumps(payload, indent=4))

    if response.status_code == 200 or response.status_code == 201:
        print("Data successfully sent to VVM API.")
        print(response.status_code)
        logging.debug("Data successfully sent to VVM API. " + str(response.status_code))

    else:
        print(f"Failed to send data to VVM API. Status code: {response.status_code}, Response: {response.text}")
        logging.error(f"Failed to send data to VVM API. Status code: {response.status_code}, Response: {response.text}")

    
def get_access_token():
    test_log("get_access_token")
    logging.debug("-->genKnownBugsList.get_access_token")
    # Get the OAuth2 token
    if get_test_mode() == 0:
        token_url = 'https://auth.prod.asoc.marquardt.de/auth/realms/cloud/protocol/openid-connect/token'
        data = {
            'username': 'vvm.api.user@marquardt.de',
            'password': 'iZdwJwCKXvFKZGDYssEX',
            'client_id': 'argus_api',
            'grant_type': 'password'
        } 
    else:
        token_url = 'https://auth.staging.asoc.marquardt.de/auth/realms/cloud/protocol/openid-connect/token'
        data = {
            'username': 'api-user@argus-sec.com',
            'password': 'PvumHZPwvmqdVZnPSGIg',
            'client_id': 'argus_api',
            'grant_type': 'password'
        }
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    test_log("data prepared")
    response = requests.post(token_url, data=data, headers=headers, verify=False)
    
    if response.status_code == 200:
        token = response.json().get('access_token')
        return token
    else:
        print(f"Failed to obtain access token. Status code: {response.status_code}, Response: {response.text}")
        logging.error(f"Failed to obtain access token. Status code: {response.status_code}, Response: {response.text}")
        return None
    

def get_project_number(cbd_number, Autosar):
    excel_path = Autosar
    try:
        df = pd.read_excel(excel_path, engine='openpyxl')

        # Normalize the column and the CBD number for consistent matching
        column_data = df.iloc[:, 10].astype(str).str.strip().str.upper()
        cbd_number = cbd_number.strip().upper()

        # Perform the search
        matches = column_data.str.contains(re.escape(cbd_number), case=False, na=False)

        if matches.any():
            # Retrieve the project number from the corresponding row
            project_number = df.loc[matches.idxmax(), df.columns[0]]  # Assuming the project number is in column 1
            return project_number
        else:
            print(f"CBD number '{cbd_number}' not found. ")
            logging.info(f"CBD number '{cbd_number}' not found.")
            return None
    except Exception as e:
        print(f"Failed to read Autosarprojects.xlsx: {e}")
        logging.error(f"Failed to read Autosarprojects.xlsx: {e}")
        return None


def check_escan_in_excel(escan, excelFile):
    test_log("-->check_escan_in_excel")
    logging.debug("-->genKnownBugsList.check_escan_in_excel")
    global escan_already_processed
    escan_already_processed = False

    try:
        # Use pandas to read the Excel file (streamlined with KnownIssuesExcel updates)
        test_log("Checking if Escan is already processed using pandas")
        df = pd.read_excel(excelFile, header=6)  # Read Excel file

        # Clean column names (strip spaces and handle unexpected characters)
        df.columns = df.columns.str.strip()

        # Check if the 'Escan_Ref' column exists
        if 'Reference' in df.columns:
            if df['Reference'].isnull().all():
                test_log("Reference column is empty.")
                return False
            df['Reference'] = df['Reference'].astype(str)
            if escan in df['Reference'].values:
                test_log(f"Reference {escan} found in {excelFile}")
                escan_already_processed = True
                return True
            else:
                test_log(f"Reference {escan} not found in {excelFile}")
                return False
        else:
            print("Column 'Escan_Ref' not found in the Excel file")
            return False

    except FileNotFoundError:
        print(f"File {excelFile} not found.")
        logging.error(f"File {excelFile} not found.")
        return False
    except Exception as e:
        print(f"An error occurred while reading {excelFile}: {e}")
        logging.error(f"An error occurred while reading {excelFile}: {e}")
        return False


def check_escan_in_VVM(token, escanDict):
    logging.debug("-->genKnownBugsList.check_escan_in_VVM")
    if token is None:
        logging.error("No access token provided")
        return
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {token}'
    }
    test_log("header built")
    
    for escan, details in escanDict.items():
        payload = {
            "filter": {
                "operator": "and",
                "operands": [
                {
                    "field": "external_vulnerability_id",
                    "simpleCondition": {
                        "operator": "eq",
                        "value": escan
                    },
                }
                ]
            },
            "page": {
                "offset": 0,
                "limit": 100
            },
            "sort": [
                {
                    "field": "ecu_count",
                    "order": "asc",
                    "nulls": "NULLS FIRST"
                }
            ]
        }
    API_URL = str(VVM_API_URL) + "/search"
    response = requests.post(API_URL, headers=headers, data=json.dumps(payload), verify=False)
    response.raise_for_status()
    data = response.json()
    if data["totalFilteredCount"] > 0 and data["items"]:
        issue_id = data["items"][0]["id"]  # Extract the ID of the first matched issue
        description = data["items"][0]["description"]
        test_log(f"Issue found: ID = {issue_id}")
        logging.info(f"Issue found: ID = {issue_id}")

        return 1, issue_id, description
    else:
        return 0, None, None
        return 0, None, None

def patch_data_in_api(escanDict, cbdNumber, issue_id, description, token, autosar, subject):
    logging.debug("-->genKnownBugsList.patch_data_in_api")
    for escan, details in escanDict.items():
        test_log(f"{escan} already in VVM")
        project_number = get_project_number(cbdNumber, autosar)
        
        details = list(details)
        BugsList = re.escape(exelfile_BugsList)
        if re.search(BugsList, description):
            test_log("escan already in VVM already written")

        else:
            
            if project_number is None:

                details = list(details) #  Convert details to a list so it can be modified
                details = list(details) #  Convert details to a list so it can be modified
                details[4] = f"{escan} \nAnalysis Done in {exelfile_BugsList}, But no project number could be found \n{details[4]}\n"
                details = tuple(details) 

            else:
                test_log("project number found " + project_number)
                if isinstance(details, tuple):
                    details = list(details) #  Convert details to a list so it can be modified
                    details = list(details) #  Convert details to a list so it can be modified
                details[4] = escan + " \nAnalysis Done in " + exelfile_BugsList + " for project " + project_number + " \n " + description + "\n"
            details = tuple(details) #  Convert the details list back to a tuple for use in the api call

            details = list(details)
            details[4] = details[4] + "\n\n E-Mail subject: " + subject
            details = tuple(details)

            issue_data = {
                "description":  re.sub(r'([a-z])\n([a-zA-Z])', r'\1 \2',details[4])
            }
            payload=(issue_data)

            if not token:
                print("No access token. Cannot send data to API.")
                logging.error("No access token. Cannot send data to API.")
                return
            
            headers = {
                'Content-Type': 'application/json',
                'Authorization': f'Bearer {token}'
            }
            API_URL = str(VVM_API_URL) + "/" +str(issue_id) #  Create the API URL with the issue ID
            response = requests.patch(API_URL, headers=headers, data=json.dumps(payload), verify=False)
            if response.status_code == 200:
                print("Vulnerability updated successfully.")
                logging.info("Vulnerability updated successfully.")
            else:
                print(f"Failed to update vulnerability. Status code: {response.status_code}, Response: {response.text}")
                logging.error(f"Failed to update vulnerability. Status code: {response.status_code}, Response: {response.text}")

        

def main(issueXml, xmlFile, api_url, autosar, subject, ASR_Functionality_Deactivated):
    test_log("-->genKnownBugsList.main")
    logging.debug("-->genKnownBugsList.main")
    global VVM_API_URL
    VVM_API_URL = api_url
    #test_log(api_url)
    #test_log(VVM_API_URL)


    #validate file returns true if the file is stored on X Drive, false if not
    #filepathValid is used later to ditinguish if Excel functions should be used
    filepathValid = validate_file(issueXml)
    test_log(" file Validated")
    
    # Get Excel file path and ensure it's available
    if filepathValid:
        excelFile = copy_template_excel(issueXml)
        test_log(" Excel file copied")

    # Parse issue XML in folder to get escan data
    num_issues = parse_issue_xml(issueXml, xmlFile, filepathValid)
    test_log(" Escan dict created")
    


    # Get report and CBD information
    reportName, releaseNumber, cbdNumber = get_report_info(issueXml)
    test_log(" Report and CBD info retrieved")
    
    # List of old CBDs
    cbds = ['CBD0800064', 'CBD0800280', 'CBD0900105', 'CBD0900135', 'CBD0900354', 'CBD0900376', 'CBD1000187',
            'CBD1100071', 'CBD1100085', 'CBD1100096', 'CBD1100102', 'CBD1200285', 'CBD1200413', 'CBD1300128',
            'CBD1300404', 'CBD1300405', 'CBD1300581', 'CBD1300669', 'CBD1400105', 'CBD1400620', 'CBD1400794',
            'CBD1500052', 'CBD1500056', 'CBD1500057', 'CBD1500431', 'CBD1500432', 'CBD1500433', 'CBD1500760',
            'CBD1500761', 'CBD1500884', 'CBD1600095', 'CBD1600100', 'CBD1600268', 'CBD1600392', 'CBD1600394',
            'CBD1600489', 'CBD1600671', 'CBD1600734', 'CBD1600781', 'CBD1600788', 'CBD1700205', 'CBD1700227',
            'CBD1700242', 'CBD1700341', 'CBD1700342', 'CBD1700343', 'CBD1700344', 'CBD1700346', 'CBD1700414',
            'CBD1700533', 'CBD1700556', 'CBD1700732', 'CBD1700863', 'CBD1700866', 'CBD1800141', 'CBD1800352',
            'CBD1800379', 'CBD1800728', 'CBD1800813', 'CBD1800883', 'CBD1800899', 'CBD1801020', 'CBD1900222',
            'CBD1900224', 'CBD1900230', 'CBD1900614', 'CBD1900950', 'CBD1901095', 'CBD2000062', 'CBD2000373',
            'CBD2000374', 'CBD2000660', 'CBD2000776', 'CBD2000777', 'CBD2000779', 'CBD2000865']

    # Determine if the CBD number is old
    cbdOld = is_old_cbd(cbdNumber, cbds)
    
    # Initialize Excel operations
    while issueObject.getCurrentIssueIndex() < num_issues:
        test_log("while loop " + str(issueObject.getCurrentIssueIndex()) + " " + str(num_issues))
        #test_log("while loop " )
        escanDict = issueObject.processNextIssue()
        #test_log(escanDict)
        print("filepathValid=",filepathValid)
        #only if the file exists and the ASR functionality is activated
        if filepathValid and not ASR_Functionality_Deactivated:
            try:
            # Log values before passing to KnownIssuesExcel
                print(f"Excel file: {excelFile}")
                print(f"CBD Old status: {cbdOld}")

                # Initialize Excel operations
                excel = MQ.KnownIssuesExcel(excelFile, cbdOld)
                print("Excel initialized.")
                logging.info("Excel initialized.")
            except Exception as e:
                print(f"An error occurred during KnownIssuesExcel initialization: {str(e)}")
                logging.error("Error during KnownIssuesExcel initialization", exc_info=True)
            currentEscanExcelDict = excel.getEscanAndDescriptionAsDict()
            #print(f"Escan dictionary: {currentEscanExcelDict}")

            # Get the highest index in the Excel sheet
            if currentEscanExcelDict:
               highestIndex = max(currentEscanExcelDict[escan][0] for escan in currentEscanExcelDict) + 1
            else:
                highestIndex = 1
            print("highest index done")
            # Write headings to Excel
            try:
                excel.writeHeadingValueToCell(15, 7, 'Report Name')
                test_log(" Heading written  in try block")
            except:
                print("Can't write to KnownBugsList.xlsx")
                logging.info("Can't write to KnownBugsList.xlsx")
            excel.writeHeadingValueToCell(16, 7, 'Release Number')
            excel.writeHeadingValueToCell(17, 7, 'Category')
            excel.set_border(15, 7, 16, 7)
            excel.set_border(16, 7, 17, 7)
        test_log("imported subject  " + str(subject))
        # Prepare and send data to VVM API
        
        if "Non-Cybersecurity-related" in subject:
            test_log("sending to API skipped due to Non-Cybersecurity-related")
        else:
            token = get_access_token()
            logging.debug("<-- genKnownBugsList.get_access_token")
            test_log("get access token exited")
            escan_in_VVM, issue_id, description = check_escan_in_VVM(token, escanDict)
            logging.debug("<-- genKnownBugsList.check_escan_in_VVM")
            test_log("Escan in VVM " + str(escan_in_VVM))
            payload = prepare_api_payload(escanDict, cbdNumber, autosar, subject)
            if escan_in_VVM:
                patch_data_in_api(escanDict, cbdNumber, issue_id, description, token, autosar, subject)
                logging.debug("<-- genKnownBugsList.patch_data_in_api")
                test_log(" Data in VVM API Patched")
                logging.debug("Data in VVM updated")
            else:
                send_data_to_api(payload, token)
                test_log(" Data sent to VVM API")
                logging.debug("Data sent to VVM API")

        if filepathValid and not ASR_Functionality_Deactivated:
            # Write new data to the Excel sheet
            test_log("Attempting to write to Excel")
            highestIndex = write_to_excel(excel, reportName, releaseNumber, escanDict, currentEscanExcelDict, highestIndex)
            test_log(" Excel written")
    
            # Save and close the Excel file
            excel.saveAndClose()
    test_log("<--genKnownBugsList.main")
    logging.debug("<-- genKnownBugsList.main")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Issue List Xml')
    parser.add_argument("issueXml", help="Issue List Xml")
    args = parser.parse_args()
    main(args.issueXml)
