from jira import JIRA
import urllib3
import logging
import html2text

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
jira_server = "https://jira.marquardt.de"
jira_token = "NDk5MjEwNTgzMTI2On+vRmf3IIl9D+dSp0MdYlAUkxcv"


def create_jira_epic_and_task(subject, email_body, project_key):
    # Authenticate with Jira
    jira_options = {'server': jira_server,
                    'verify': False}
    jira = JIRA(options=jira_options, token_auth=(jira_token))
    
    # Create the Epic
    epic_issue = jira.create_issue(fields={
        'project': {'key': project_key},
        'summary': subject,
        'description': email_body,  
        'issuetype': {'name': 'Epic'},
        'customfield_10102': subject,
        'customfield_11905': 'PSIRT-Mailbox_no_VVM'
    })

    # Create a Task linked to the Epic
#    task_issue = jira.create_issue(fields={
#        'project': {'key': project_key},
#        'summary': subject,
#        'description': email_body,
#        'issuetype': {'name': 'Task'},
#        'customfield_10100': epic_issue.key  # This is the "Epic Link" field
#    })

    return epic_issue

def main(message, jira_project_key):
    # Example usage
    subject = message['subject']
    logging.debug(f"Creating Jira epic for Subject: {subject}")
    email_body = html2text.html2text(message['body']['content'])
    epic_issue = create_jira_epic_and_task(subject, email_body, jira_project_key)
    print(f"Created Epic {epic_issue.key} for subject: {subject}")
    logging.debug(f"Created Epic {epic_issue.key} for subject: {subject}")
    return epic_issue