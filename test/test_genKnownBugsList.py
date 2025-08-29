# mailboxissueautomation/test/test_genKnownBugsList.py

import pytest
import os
from unittest.mock import MagicMock, patch
import sys
import datetime
import json
import requests
import pandas as pd
import uuid


# Mock external functions/modules to isolate tests from dependencies
sys.modules['Vector_Issue.MQ'] = MagicMock()
sys.modules['requests'] = MagicMock()

# Assuming all the functions from the provided code are imported directly
from src.Vector_Issue.genKnownBugsList import (
    validate_file,
    copy_template_excel,
    parse_issue_xml,
    get_report_info,
    is_old_cbd,
    write_to_excel,
    prepare_api_payload,
    send_data_to_api,
    get_access_token,
    get_project_number,
    check_escan_in_excel,
    check_escan_in_VVM,
    patch_data_in_api,
    main,
)

# Test validate_file function
def test_validate_file_valid():
    print("Testing validate_file with a valid file path")
    # Create a dummy file
    valid_file = "test_issue.xml"
    with open(valid_file, 'w') as f:
        f.write("<xml>valid</xml>")

    try:
        validate_file(valid_file)
        print("validate_file passed for valid file.")
    finally:
        os.remove(valid_file)

def test_validate_file_invalid():
    print("Testing validate_file with an invalid file path")
    invalid_file = "non_existent_file.xml"
    with pytest.raises(SystemExit):
        validate_file(invalid_file)
    print("validate_file correctly raised an exception for invalid file.")

# Test copy_template_excel function
def test_copy_template_excel():
    print("Testing copy_template_excel function")
    mock_issue_xml = "MockFiles/mock_issue.xml"
    mock_template_path = "src/Vector_Issue/KnownBugsList_Template.xlsx"
    
    # Mock file creation
    os.makedirs(os.path.dirname(mock_issue_xml), exist_ok=True)
    with open(mock_issue_xml, 'w') as f:
        f.write("<xml>mock</xml>")
    
    excel_file = copy_template_excel(mock_issue_xml)
    print(f"Excel file copied to: {excel_file}")

    # Clean up
    os.remove(mock_issue_xml)

# Test parse_issue_xml function
def test_parse_issue_xml():
    print("Testing parse_issue_xml function with valid XML content")
    # Use mock here as XML parsing needs to be tested with real data
    mock_issue_xml = "mock_issue.xml"
    with open(mock_issue_xml, 'w') as f:
        f.write("<issueReport><issues><issue>Test</issue></issues></issueReport>")

    try:
        result = parse_issue_xml(mock_issue_xml)
        print(f"Parsed XML result: {result}")
    finally:
        print()
        #os.remove(mock_issue_xml)

# Test get_report_info function
def test_get_report_info():
    print("Testing get_report_info function")
    mock_issue_xml = "mock_SecurityRelated_1234_5678_91011.xml"
    reportName, releaseNumber, cbdNumber = get_report_info(mock_issue_xml)
    print(f"Report Name: {reportName}, Release Number: {releaseNumber}, CBD Number: {cbdNumber}")
    assert reportName == "mock_SecurityRelated_1234_5678_91011"
    assert releaseNumber == "5678"
    assert cbdNumber == "1234"

# Test is_old_cbd function
def test_is_old_cbd():
    print("Testing is_old_cbd function")
    cbds = ['CBD0800064', 'CBD0800280', 'CBD0900105']
    result = is_old_cbd('CBD0800064', cbds)
    assert result is True
    print("is_old_cbd returned True as expected for known CBD.")

# Test write_to_excel function (Mocked Excel interactions)
def test_write_to_excel():
    print("Testing write_to_excel function")
    mock_excel = MagicMock()
    mock_excel.sheet.cell.return_value.value = None  # Simulate an empty cell
    excel = mock_excel  # This should simulate the Excel object being passed
    reportName = "Report1"
    releaseNumber = "1.0"
    escanDict = {"ES1234": ("libA", "v1.0", "desc", "2022-01-01", "Additional info", "high")}
    currentEscanExcelDict = {}
    highestIndex = 1

    new_highestIndex = write_to_excel(excel, reportName, releaseNumber, escanDict, currentEscanExcelDict, highestIndex)
    print(f"New highest index after writing: {new_highestIndex}")

# Test prepare_api_payload function
def test_prepare_api_payload():
    from src.Vector_Issue.utils import set_test_mode
    set_test_mode(1)
    print("Testing prepare_api_payload function")
    escanDict = {
        "ES1234": ("libA", "1.0", "2.0", "header", "desc", "Workaround", "2022-01-01", "2022-01-01", "high")
    }
    cbdNumber = "CBD1234"
    escan="ES1234"
    try:
        escan1 =str(uuid.UUID(escan))
    except ValueError:
        escan1 = str(uuid.uuid5(uuid.NAMESPACE_DNS, escan))
    payload = prepare_api_payload(escanDict, cbdNumber)
    expected_payload = [{
        "external_vulnerability_id": "ES1234",
        "vulnerability_type": "SW",
        "source": "PSIRT-Mailbox",
        "publication_created": "2022-01-01",
        "description": "desc",
        "time_created": "2022-01-01",
        "affected_libraries": [{
            "id": escan1,
            "name": "libA",
            "vendor": "Vector",
            "versions": [{
                "start_version": "1.0",
                "start_include": True,
                "end_version": "99.99.99",
                "end_include": False
            }],
            "vulnerabilityId": escan1
        }],
        "is_confidential": False,
        "analysis": "",
        "mitigation_plan": ""
    }]
    
    assert payload == expected_payload

def test_handles_empty_input_dict():
    escanDict = {}
    expected_payload = []
    cbdNumber="CBD1234"
    print ("Call the prepare_api_payload function with the empty dictionary")
    result = prepare_api_payload(escanDict, cbdNumber)
    print("Assert that the result is equal to the empty list")
    assert result == expected_payload

# Test send_data_to_api function
def test_send_data_to_api():
    from src.Vector_Issue.utils import set_test_mode
    set_test_mode(1)
    print("Testing send_data_to_api function")
    payload = {"mock_data": "value"}
    token = "mock_token"
    VVM_API_URL = "https://api.mock.com"
    send_data_to_api(payload, token)
    print("Data sent to VVM API (mocked).")

def test_get_access_token():
    from src.Vector_Issue.utils import set_test_mode
    set_test_mode(1)
    print("Test: get_access_token - Mocking token retrieval.")

    with patch("requests.post") as mock_post:
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {"access_token": "mock_token"}
        mock_post.return_value = mock_response

        token = get_access_token()
        assert token == "mock_token"
        print("get_access_token: Successfully retrieved mock token.")


# Test get_project_number function
def test_get_project_number():
    print("Testing get_project_number function")
    cbdNumber = "CBD2100737"
    project_number = get_project_number(cbdNumber)
    print(f"Project number: {project_number}")
    assert isinstance(project_number, str)

# Test check_escan_in_excel function
def test_check_escan_in_excel():
    from src.Vector_Issue.utils import set_test_mode
    set_test_mode(1)
    print("Testing check_escan_in_excel function")
    escan = "ESCAN00118425"
    excelFile = "KnownBugsList.xlsx"  # Use a mock path for testing
    result = check_escan_in_excel(escan, excelFile)
    print(f"Check if Escan found in Excel: {result}")
    assert isinstance(result, bool)

# Test check_escan_in_VVM function
def test_check_escan_in_VVM():
    with patch('src.Vector_Issue.genKnownBugsList.VVM_API_URL',
               'https://vvm.a.staging.asoc.marquardt.de/public-api/vulnerability-management/v1/vulnerabilities'):
        from src.Vector_Issue.utils import set_test_mode
        set_test_mode(1)
        print("Testing check_escan_in_VVM function")

        # Mock the get_access_token function to return a mock token
        with patch('src.Vector_Issue.genKnownBugsList.get_access_token') as mock_get_access_token:
            mock_get_access_token.return_value = "mock_token"

            # Mock the requests.post method to return a mock response
            with patch('requests.post') as mock_post:
                mock_response = MagicMock()
                mock_response.status_code = 200
                mock_response.json.return_value = {
                    "totalFilteredCount": 1,
                    "items": [
                        {
                            "id": "mock_id",
                            "description": "mock_description"
                        }
                    ]
                }
                mock_post.return_value = mock_response

                escanDict = {"ESCAN00118425": ("libA", "v1.0", "v2.0", "header", "desc", "Workaround", "2022-01-01", "2022-01-01", "high")}
                result = check_escan_in_VVM("mock_token", escanDict)
                assert result == (1, "mock_id", "mock_description")

                print("check_escan_in_VVM: Successfully checked Escan in VVM.")

# Test patch_data_in_api function
def test_patch_data_in_api():
    from src.Vector_Issue.utils import set_test_mode
    set_test_mode(1)
    print("Testing patch_data_in_api function")
    escanDict = {"ES1234": ("libA", "v1.0", "desc", "2022-01-01", "Additional info", "high")}
    cbdNumber = "CBD1234"
    issue_id = "mock_issue_id"
    description = "mock description"
    token = "mock_token"
    patch_data_in_api(escanDict, cbdNumber, issue_id, description, token)
    print("Data patched successfully in VVM (mocked).")

@patch('requests.post')
def test_send_data_to_api_success(mock_post):
    """Test successful API response handling."""
    mock_post.return_value.status_code = 200
    mock_post.return_value.text = 'Success'

    payload = [{"test": "data"}]
    token = "mock_token"
    send_data_to_api(payload, token)
    
    mock_post.assert_called_once()
    assert mock_post.return_value.status_code == 200

