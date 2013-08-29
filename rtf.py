# Module for manipulating RTF files

import datetime

import common

 

class Rtf:
    def __init__(self):
        self.text = '{\\rtf {\\fonttbl {\\f0 Consolas;}}\n'
        self.text += '\pard\plain\n' # single-spaced paragraphs
        #self.text = '{\\rtf {\\fonttbl {\\f0 Courier New;}}'
           
    def addTitle(self, title):
        self.setFontSize(30)
        self.add(title)
        self.setFontSize(18)
        
    def page(self):
        'Add a page break'
        self.text += '\n\n\\page\n\n'

    def para(self, n = 1):
        'Insert N paragraph lines'
        text = ' \\par' * n
        self.text += '\n\n%s\n\n' %(text)

    def add(self, text, newLines = 1):
        if not text: return
        self.text += common.AsAscii(text)
        self.para(newLines)
      
    def setFontSize(self, size):
        self.add('\\f0\\fs%d' % (size), newLines = 0)
        
    def annotation(self, job):
        'Return some RTF text for non-vanilla jobs'
        
        self.para(3)
        #self.add('Approved: ______________________________________')        
        #self.add(approver)
        self.setFontSize(12)        
        self.add(common.get_timestamp())
        
        annotations = []
        if job['Weird']: annotations.append('Unorthodox')
        if job['WIP']: annotations.append('WIP')
        if len(annotations) > 0:
            self.para(2)            
            text = 'Ann: ' + ' '.join(annotations)
            self.add(text)
        self.setFontSize(18)


    def save(self, dir, fname):
        'Save Rtf file'
        common.makedirs(dir)
        fullName = "%s\\%s" % (dir, fname)
        common.spit(fullName, self.text + '}') 