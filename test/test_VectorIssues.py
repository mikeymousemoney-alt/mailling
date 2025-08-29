import pytest
import datetime as dt
import re
from unittest.mock import patch, mock_open, MagicMock
from Vector_Issue.Vector_Issue import Vector_IssueApp
from Vector_Issue import Vector_Issue as VI



# Test for the instantiation of Vector_IssueApp with a valid start date
def test_Vector_IssueApp_initialization():
    """Test the initialization of the Vector_IssueApp class with a valid start date."""
    with patch("Vector_Issue.Vector_Issue.Vector_IssueApp.main") as mock_main:
        app = Vector_IssueApp("01012025")
        mock_main.assert_called_once()
        assert app is not None


# Test for updating constants with valid configuration
def test_update_constants_with_valid_config():
    """Test updating constants with a valid configuration."""
    mock_config = {
        "test_mode": "1",
        "set_read": "0",
        "test_mail": "",
        "test_name": "Test User",
        "move_mails": "0",
        "change_to_C_partition": "1",
        "send_mails": "1",
        "mailbox_address": "test@example.com",
        "mailbox_folder": "Inbox",
        "folder_processed_mail_vector": "Processed",
        "folder_processed_mail_others": "Processed Other",
        "email_subject": "Vector Issue Report",
        "vector_mail_address": "vector@example.com",
        "MIA_FILTER_NOSECURITY": ["Test1", "Test2"],
        "folder_no_security": "NoSecurity",
        "VVM_API_URL": "https://api.test.com",
        "cybersecurity_manager_email": "security@example.com",
        "config_mia_processing": "All_Emails",
        "jira_project_key": "JIRA123",
        "path_to_issue_report": "D:\\IssueReports\\path\\to\\xml\\more\\path",
        "autosar_projects_list": "D:\\AutosarProjects.xlsx",
        "path_to_log_file": "D:\\Logs\\VectorIssueLogs.log",
        "column_projectnumber": "0",
        "column_microsar_package_type": "1",
        "column_fbl_package_type": "2",
        "column_hsm_package_type": "3",
        "column_project_status": "4",
        "column_bsw_integrator": "5",
        "column_fbl_integrator": "6",
        "column_extern_integrator": "7",
        "column_microsar_package": "8",
        "column_fbl_package": "9",
        "column_hsm_package": "10",   
    }

    with patch("builtins.open", mock_open(read_data=str(mock_config))):
        with patch("json.load", return_value=mock_config):
            Vector_IssueApp.UpdateConstants(mock_config)
            


# Test for date validation
def test_get_start_date():
    """Test that the correct date is returned for a valid input."""
    with patch("builtins.input", return_value="01012025"):
        date = Vector_IssueApp.get_start_date()
        assert date.strftime("%d%m%Y") == "01012025"


# Test for partition change functionality
def test_change_partition():
    """Test that the partition in a path is correctly changed."""
    new_path = Vector_IssueApp.change_partition("D:\\Folder\\SubFolder", new_partition="C:")
    assert new_path == "C:\\Folder\\SubFolder"

    with pytest.raises(ValueError):
        Vector_IssueApp.change_partition("\\Folder\\SubFolder")


def test_CheckNewMails():
    """Test the CheckNewMails function for email filtering and processing."""
    with patch("win32com.client.Dispatch") as mock_dispatch:
        # Mock objects for inbox and messages
        mock_outlook = MagicMock()
        mock_namespace = MagicMock()
        mock_inbox = MagicMock()
        mock_messages = MagicMock()
        mock_message = MagicMock()

        # Set up mock behavior
        mock_dispatch.return_value = mock_outlook
        mock_outlook.GetNamespace.return_value = mock_namespace
        mock_namespace.Folders.return_value = mock_inbox
        mock_inbox.Folders.return_value = mock_inbox
        mock_inbox.Items.return_value = mock_messages
        mock_messages.Restrict.return_value = [mock_message]

        # Mock message properties
        mock_message.Subject = "Vector Issue Report"
        mock_message.Attachments = []
        mock_message.Sender = "vector@example.com"

        # Call the method
        Vector_IssueApp.CheckNewMails()

        # Assertions
        mock_dispatch.assert_called_once_with("Outlook.Application")
        mock_namespace.Folders.assert_called()

# Test for ReadAutosarprojects functionality
def test_ReadAutosarprojects():
    """Test the ReadAutosarprojects function for parsing project data."""
    with patch("openpyxl.load_workbook") as mock_load_workbook:
        # Mock Excel workbook and sheet
        mock_workbook = MagicMock()
        mock_sheet = MagicMock()
        mock_load_workbook.return_value = mock_workbook
        mock_workbook.__getitem__.return_value = mock_sheet

        # Mock rows in the sheet
        mock_sheet.iter_rows.return_value = [
            ("Project1", "Microsar", "FBL", "HSM", "Active", "Integrator1", "Integrator2", "Integrator3", "Microsar", "FBL", "HSM", "Active", "Integrator1", "Integrator2", "Integrator3"),
            ("Project2", None, None, None, "Stopped", None, None, None, None, None, None, "Stopped", None, None, None),
        ]

        # Call the method
        Vector_IssueApp.ReadAutosarprojects()

        # Assertions
        mock_load_workbook.assert_called_once_with(VI.c_autosarprojectsListDir)
        mock_workbook.__getitem__.assert_called_once_with(VI.c_autosarprojectsSheet)
        mock_sheet.iter_rows.assert_called()

# Test for GetIntegrators functionality
def test_GetIntegrators():
    """Test the GetIntegrators function for extracting integrator information."""
    with patch("Vector_Issue.Vector_Issue.Vector_IssueApp.GetIntegrators") as mock_get_integrators:
        # Mock data
        mock_sheet = MagicMock()
        mock_row = ["Project1", "Microsar", "Integrator1"]
        row_value = "MicrosarPackage"
        package_type = "Microsar"

        # Call the method
        Vector_IssueApp.GetIntegrators(mock_sheet, 1, row_value, package_type)

        # Assertions
        mock_get_integrators.assert_called_with(mock_sheet, 1, row_value, package_type)

# Subject matches c_vectorEmailSubject pattern exactly
def test_subject_matches_pattern_exactly():
    # Arrange
    subject = "Vector Issue Report for CBD123: 01.01.2024"
    c_vectorEmailSubject = "Vector.+Issue Report for CBD.+"

    # Act
    result = re.fullmatch(c_vectorEmailSubject, subject)

    # Assert
    assert result is not None

def test_FilterEmails_VectorIssuesOnly():
    """Test that non-matching emails are skipped if CONFIG_MIA_PROCESSING is Vector_Issues_Only."""
    with patch("Vector_Issue.Vector_Issue.c_mia_processing", "Vector_Issues_Only"):
        with patch("win32com.client.Dispatch") as mock_dispatch:
            mock_message = MagicMock()
            mock_message.Subject = "Non-Vector Email"
            mock_message.Attachments = []
            mock_message.Sender = "random@example.com"

            mock_messages = MagicMock()
            mock_messages.Restrict.return_value = [mock_message]

            mock_dispatch.return_value.GetNamespace.return_value.Folders.return_value.Items.return_value = mock_messages

            # Call the method
            Vector_IssueApp.CheckNewMails()

            # Assert that emails not matching the filter are skipped
            mock_message.mark_as_read.assert_not_called()

def test_MarkEmailsAsRead():
    """Test that parsed emails are marked as read for applicable configurations."""
    with patch("Vector_Issue.Vector_Issue.c_mia_processing", "All_Emails"):
        with patch("Vector_Issue.Vector_Issue.set_read", 1):
            with patch("win32com.client.Dispatch") as mock_dispatch:
                mock_message = MagicMock()
                mock_message.Subject = "Vector Issue Report"
                mock_message.Attachments = []
                mock_message.Sender = "psirt@marquardt.com"
                mock_message.Unread = 1  # Set the initial value of Unread

                mock_messages = MagicMock()
                mock_messages.Restrict.return_value = [mock_message]

                mock_dispatch.return_value.GetNamespace.return_value.Folders.return_value.Items.return_value = mock_messages

                # Call the method
                Vector_IssueApp.CheckNewMails()

                # Assert that matching emails are marked as read
                assert mock_message.Unread == 1

""" def test_ParseMailboxByDate():
"""    """Test parsing mailbox by a given start date.""" """
    start_date = "01012025"
    with patch("Vector_Issue.Vector_Issue.c_start_date", start_date):
        with patch("win32com.client.Dispatch") as mock_dispatch:
            mock_message = MagicMock()
            mock_message.ReceivedTime = "02012025"
            mock_message.Subject = "Vector Issue Report"
            mock_message.Attachments = []
            mock_message.Sender = "vector@example.com"

            mock_messages = MagicMock()
            mock_messages.Restrict.return_value = [mock_message]

            mock_dispatch.return_value.GetNamespace.return_value.Folders.return_value.Items.return_value = mock_messages

            # Call the method
            Vector_IssueApp.CheckNewMails()

            # Assert that emails are filtered by the start date
            assert Vector_IssueApp.filteredMessages == "[ReceivedTime] >= '01012025' AND [Unread] = True" """

def test_ConfigMIAProcessing_AllEmails():
    """Test CONFIG_MIA_PROCESSING configuration for 'All_Emails'."""
    with patch("Vector_Issue.Vector_Issue.c_mia_processing", "All_Emails"):
        assert VI.c_mia_processing == "All_Emails"



if __name__ == "__main__":
    pytest.main()

import unittest
from unittest.mock import patch, MagicMock
import os
import pandas as pd

class TestAdditionalFunctions(unittest.TestCase):

    def test_BuildIntegratorEmail(self):
        integrator = [
            "John, Doe",
            "Jane Doe",
            "john.doe@example.com",
            "Invalid Name"
        ]
        global c_marquardtEmail
        c_marquardtEmail = "example.com"
        
        emails, none_correct = Vector_IssueApp.BuildIntegratorEmail(integrator)
        
        expected_emails = [
            "Doe.John@marquardt.com",
            "Jane.Doe@marquardt.com",
            "john.doe@example.com",
            None
        ]
        expected_none_correct = [0, 0, 0, None]
        
        self.assertEqual(emails, expected_emails)
        self.assertEqual(none_correct, expected_none_correct)

    @patch("os.path.exists", return_value=False)
    @patch("pandas.DataFrame.to_excel")
    def test_create_processed_issues_excel_new_file(self, mock_to_excel, mock_exists):
        processed_issues = [
            ["P1", "CBD1", "1.0", "2025-01-01", "John Doe", "Jane Doe", "Ext1", "Yes"]
        ]
        global c_pathIssuesXlsx, xmlFileDir
        c_pathIssuesXlsx = "\\..\\.."
        xmlFileDir = "\\path\\to\\xml\\more\\path"
        
        with patch("Vector_Issue.Vector_Issue.xmlFileDir", "\\path\\to\\xml\\more\\path"):
            Vector_IssueApp.create_processed_issues_excel(processed_issues)
            mock_to_excel.assert_called_once()

    @patch("pandas.read_excel")
    @patch("os.path.exists", return_value=True)
    @patch("pandas.DataFrame.to_excel")
    def test_create_processed_issues_excel_existing_file(self, mock_to_excel, mock_exists, mock_read_excel):
        processed_issues = [
            ["P1", "CBD1", "1.0", "2025-01-01", "John Doe", "Jane Doe", "Ext1", "Yes"]
        ]
        existing_issues = pd.DataFrame([
            ["P2", "CBD2", "1.0", "2025-01-02", "John Smith", "Jane Smith", "Ext2", "Yes"]
        ], columns=[
            "Project Number", "Subject CBD Number", "Version", "Subject Date", "BSW Integrator", "FBL Integrator", "Ext Integrator", "Email Sent"
        ])
        mock_read_excel.return_value = existing_issues
        
        global c_pathIssuesXlsx, xmlFileDir
        c_pathIssuesXlsx = "\\..\\.."
        xmlFileDir = "\\path\\to\\xml\\more\\path"
        with patch("Vector_Issue.Vector_Issue.xmlFileDir", "\\path\\to\\xml\\more\\path"):
            Vector_IssueApp.create_processed_issues_excel(processed_issues)
            mock_to_excel.assert_called_once()

    @patch("win32com.client.Dispatch")
    def test_send_summary_email(self, mock_dispatch):
        unprocessed_issues = [
            ["P1", "CBD1", "1.0", "2025-01-01", "John Doe", "Jane Doe", "Ext1", "Reason1"]
        ]
        unprocessed_mail = [
            ["testmail", "test@test.com"]
        ]
        global c_cyberSecManager, send_Mails
        c_cyberSecManager = "cybersec@company.com"
        send_Mails = 1

        Vector_IssueApp.send_summary_email(unprocessed_issues, unprocessed_mail)
        mock_dispatch.assert_called_once()

    def test_format_unprocessed_issues(self):
        self.maxDiff = None
        issues = [
            ["P1", "CBD1", "1.0", "2025-01-01", "John Doe", "Jane Doe", "Ext1", "Reason1"],
            ["P2", "Unknown", "Unknown", "Unknown", "John Doe", "Jane Doe", "Ext1", "Unknown"]
        ]
        unprocessed_mail = [
            ["testmail", "test@test.com"],
            ["testmail2", "test2@test.com"]
        ]
        expected_output = (
            "Summary of Unprocessed Emails:\n\n"
            "Sender                             || Subject                                                \n"
            "CBD Number               | Version   | Date           | Reason                                  \n"
            "=============================================================================================\n"
            "test@test.com                      || testmail                                               \n"
            "CBD1                     | 1.0       | 2025-01-01     | Reason1                                 \n"
            "---------------------------------------------------------------------------------------------\n"
            "test2@test.com                     || testmail2                                              \n"
            "N/A                      | N/A       | N/A            | Unknown Reason                          \n"
            "---------------------------------------------------------------------------------------------\n"
        )
        self.assertEqual(Vector_IssueApp.format_unprocessed_issues(issues, unprocessed_mail), expected_output)

if __name__ == "__main__":
    unittest.main()


class TestCheckNewMails(unittest.TestCase):
    
    @patch('win32com.client.Dispatch')
    def test_filter_messages_by_date(self, mock_dispatch):
        # Setup
        mock_inbox = MagicMock()
        mock_messages = MagicMock()
        mock_filtered_messages = MagicMock()

        # Mock the Outlook setup
        mock_dispatch.return_value.GetNamespace.return_value.Folders.return_value.Folders.return_value = mock_inbox
        mock_inbox.Items = mock_messages
        mock_inbox.Items.Restrict.return_value = mock_filtered_messages

        # Create a date for filtering
        global c_start_date
        c_start_date = dt.datetime(2023, 10, 1, 0, 0)

        # Call the function
        Vector_IssueApp.CheckNewMails()

        # Check if the filter was applied correctly
        expected_filter = "[ReceivedTime] >= '01012025' AND [Unread] = True"
        mock_messages.Restrict.assert_called_once_with(expected_filter)

    """ @patch('win32com.client.Dispatch')
    def test_send_reply(self, mock_dispatch):
        # Setup
        mock_message = MagicMock()
        mock_message.Subject = "Test Subject"
        mock_message.SenderEmailAddress = "sender@example.com"
        mock_message.Unread = True
        mock_message.ReceivedTime = dt.datetime(2025, 1, 2, 12, 0)
        
        # Mock the Outlook setup
        mock_outlook = mock_dispatch.return_value
        mock_namespace = mock_outlook.GetNamespace.return_value
        mock_inbox = mock_namespace.Folders.return_value
        mock_messages = mock_inbox.Items
        mock_messages.Restrict.return_value = [mock_message]

        global c_vectorEmailSubject
        c_vectorEmailSubject = "Test Subject"
        global c_vector_mail_address
        c_vector_mail_address = "different@example.com"  # Ensure this is different from the sender
        global send_Mails
        send_Mails = 1
        
        # Call the function
        Vector_IssueApp.CheckNewMails()
        
        # Check if a reply was created and sent
        mock_reply = mock_outlook.CreateItem.return_value
        mock_reply.Send.assert_called_once()
        self.assertEqual(mock_reply.Subject, "Re: Test Subject")
        self.assertEqual(mock_reply.Body, "Thank you for reaching out to Marquardt Product Security Incident Response Team. Your request reached us, we're on it!")
        self.assertEqual(mock_reply.To, "sender@example.com")

    @patch('win32com.client.Dispatch')
    def test_forward_message(self, mock_dispatch):
        # Setup
        mock_message = MagicMock()
        mock_message.Subject = "Test Subject"
        mock_message.Body = "This is the body of the email."
        mock_message.SenderEmailAddress = "sender@example.com"
        
        # Mock the Outlook setup
        mock_outlook = mock_dispatch.return_value
        mock_inbox = mock_outlook.GetNamespace.return_value.Folders.return_value
        mock_inbox.Items.Restrict.return_value = [mock_message]

        global c_cyberSecManager
        c_cyberSecManager = "cybersec@example.com"
        global send_Mails
        send_Mails = 1
        
        # Call the function
        Vector_IssueApp.CheckNewMails()
        
        # Check if the message was forwarded
        mock_forward = mock_outlook.CreateItem.return_value
        mock_forward.Send.assert_called_once()
        self.assertEqual(mock_forward.Subject, "Fwd: Test Subject")
        self.assertIn("Forwarding this message for further review.", mock_forward.Body)
        self.assertEqual(mock_forward.To, "cybersec@example.com") """

if __name__ == '__main__':
    unittest.main()