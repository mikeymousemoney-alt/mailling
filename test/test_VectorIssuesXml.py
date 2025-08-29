# test_vector_issues.py
""" 
import pytest
import xml
from Vector_Issue.MQ.vectorIssuesXml import VectorIssuesXml
from Vector_Issue.utils import set_test_mode

class TestVectorIssuesXml:

    set_test_mode(1)

    def test_initialization_with_valid_xml_and_category_filter(self):
        # Create a string containing XML content
        print(" Creating mockup XML content")
        xml_content = '''
        <issueReport>
            <reportData>
                <reportIdentifier>CBD1234567-D04-2024-07-24-12:30:49</reportIdentifier>
                <reportCreationDate>2024-07-24</reportCreationDate>
            </reportData>
            <issues>
                <issue category="safetyRelevant">
                    <identifier>ISSUE-1</identifier>
                    <package>package1</package>
                    <firstAffectedVersion>1.0</firstAffectedVersion>
                    <versionsFixed>1.1</versionsFixed>
                    <headline>Issue 1 headline</headline>
                    <problemDescription>Issue 1 description</problemDescription>
                    <resolutionDescription>Issue 1 resolution</resolutionDescription>
                </issue>
            </issues>
        </issueReport>
        '''
        # Write the XML content to a file
        with open('test.xml', 'w', encoding='utf-8') as f:
            f.write(xml_content)

        print(" Testing VectorIssuesXml class with correct XML file and category filter")
        # Initialize the VectorIssuesXml class with the XML file and a category filter
        vector_issues = VectorIssuesXml('test.xml', ['@safetyRelevant'])
        # Get the escan dictionary from the VectorIssuesXml class
        escan_dict = vector_issues.getEscanDict()

        print(" Checking the Output on correct behavior")
        # Assert that the issue identifier is in the escan dictionary
        # assert 'ISSUE-1' in escan_dict
        # Assert that the issue details are correct
        assert escan_dict['ISSUE-1'] == ('package1', '1.0', '1.1', 'Issue 1 headline', 'Issue 1 description', 'Issue 1 resolution', '2024-07-24T12:30:49Z', '2024-07-24T12:30:49Z', 'Cat1')

    def test_initialization_with_empty_category_filter(self):
        # Create an XML content string
        print(" Creating mockup XML content")
        xml_content = """ """
        <issueReport>
            <reportData>
                <reportIdentifier>1234-5678-91011</reportIdentifier>
                <reportCreationDate>2023-10-01</reportCreationDate>
            </reportData>
            <issues>
                <issue category="Cat1">
                    <identifier>ISSUE-1</identifier>
                    <package>package1</package>
                    <firstAffectedVersion>1.0</firstAffectedVersion>
                    <versionsFixed>1.1</versionsFixed>
                    <headline>Issue 1 headline</headline>
                    <problemDescription>Issue 1 description</problemDescription>
                    <resolutionDescription>Issue 1 resolution</resolutionDescription>
                </issue>
            </issues>
        </issueReport>
        """
"""         # Write the XML content to a file
        with open('test.xml', 'w', encoding='utf-8') as f:
            f.write(xml_content)

        print(" Testing VectorIssuesXml class with empty category filter")
        # Initialize the VectorIssuesXml object with the file and an empty category filter
        vector_issues = VectorIssuesXml('test.xml', [])
        print(" Checking the Output on correct behavior")
        # Assert that no issues are ignored
        assert vector_issues.numIssuesIgnored == 0
        # Assert that the Escan dictionary contains one issue
        assert len(vector_issues.getEscanDict()) == 1
        # Assert that the issue with identifier 'ISSUE-1' is in the Escan dictionary
        assert 'ISSUE-1' in vector_issues.getEscanDict()

 """