import unittest
from unittest.mock import patch, MagicMock
from Vector_Issue.jira1 import create_jira_epic_and_task, main  # Passe den Modulnamen an

class TestJiraIntegration(unittest.TestCase):
    
    @patch("Vector_Issue.jira1.JIRA")  # Mock die JIRA-Klasse
    def test_create_jira_epic_and_task(self, mock_jira):
        # Mock die JIRA-Instanz
        mock_instance = mock_jira.return_value
        mock_issue = MagicMock()
        mock_issue.key = "SWTST003-123"
        mock_instance.create_issue.return_value = mock_issue
        
        subject = "Test Epic"
        email_body = "This is a test email body."
        project_key = "SWTST003"
        
        epic_issue = create_jira_epic_and_task(subject, email_body, project_key)
        
        # Überprüfe, ob die JIRA-API korrekt aufgerufen wurde
        mock_instance.create_issue.assert_called_once_with(fields={
            'project': {'key': project_key},
            'summary': subject,
            'description': email_body,
            'issuetype': {'name': 'Epic'},
            'customfield_10102': subject
        })
        
        self.assertEqual(epic_issue.key, "SWTST003-123")
    
    @patch("Vector_Issue.jira1.create_jira_epic_and_task")
    def test_main(self, mock_create_epic):
        mock_epic = MagicMock()
        mock_epic.key = "SWTST003-123"
        mock_create_epic.return_value = mock_epic
        
        message = MagicMock()
        message.Subject = "Test Subject"
        message.Body = "Test Body"
        
        epic_issue = main(message, "SWTST003")
        
        mock_create_epic.assert_called_once_with("Test Subject", "Test Body", "SWTST003")
        self.assertEqual(epic_issue.key, "SWTST003-123")

if __name__ == "__main__":
    unittest.main()
