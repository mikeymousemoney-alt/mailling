import xmltodict


class VectorIssuesXml:
    """ Vector Issue XML client class """

    def __init__(self, file, categoryFilter=[]):
        # initialize internal variables
        self.file = file
        with open(self.file, encoding='utf-8', errors='ignore') as fd:
            self.doc = xmltodict.parse(fd.read())
        self.escanDict = {}
        # prepend @ for each category string
        self.categoryFilter = categoryFilter
        self.numIssuesIgnored = 0
        self.createEscanDict()

    def createEscanDict(self):
        if self.file:
            # if issueReport->issues->issue is a list loop through the list else only one entry is there
            if isinstance(self.doc['issueReport']['issues']['issue'], list):
                for issue in self.doc['issueReport']['issues']['issue']:
                    if issue['@category'] in self.categoryFilter:
                        self.numIssuesIgnored += 1
                        continue
                    self.escanDict[issue['identifier']] = (issue['package'],
                                                           issue['firstAffectedVersion'],
                                                           issue['versionsFixed'],
                                                           issue['headline'],
                                                           issue['problemDescription'],
                                                           issue['resolutionDescription'] if 'resolutionDescription' in issue else '')
            else:
                if self.doc['issueReport']['issues']['issue']['@category'] in self.categoryFilter:
                    self.numIssuesIgnored += 1
                self.escanDict[self.doc['issueReport']['issues']['issue']['identifier']] = (self.doc['issueReport']['issues']['issue']['package'],
                                                       self.doc['issueReport']['issues']['issue']['firstAffectedVersion'],
                                                       self.doc['issueReport']['issues']['issue']['versionsFixed'],
                                                       self.doc['issueReport']['issues']['issue']['headline'],
                                                       self.doc['issueReport']['issues']['issue']['problemDescription'],
                                                       self.doc['issueReport']['issues']['issue']['resolutionDescription'] if 'resolutionDescription' in self.doc['issueReport']['issues']['issue'] else '')

        if self.numIssuesIgnored != 0:
            print(str(self.numIssuesIgnored) + " Issues Ignored!!")

    def getEscanDict(self):
        return self.escanDict
