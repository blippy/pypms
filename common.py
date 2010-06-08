import datetime, json, logging, pprint, os, unicodedata

import db, expenses, period


###########################################################################
    
class DataIntegrityError(Exception):
    def __init__(self, value):
        self.value = value

    def __str__(self):
        return repr(self.value)

###########################################################################
# file and directory handling utilties    
outroot = 'M:\\Finance\\pypms\\out'

def makedirs(path):
    path = os.path.abspath(path)
    if not os.path.exists(path): os.makedirs(path)



###########################################################################
# logging

logging.basicConfig(level=logging.ERROR,
                    format='%(asctime)s %(levelname)s %(message)s',
                    filename= outroot + '\\pypms.log',
                    filemode='a')
def logdebug(txt): logging.debug(txt)
def loginfo(txt): logging.info(txt)
def logerror(txt): logging.error(txt)
def logwarn(txt): logging.warn(txt)
    

###########################################################################
# aggregation routines

def mkKeyFunc(fieldName):
    def func(rec): return rec[fieldName]
    return func

###########################################################################

def annotation(job):
    'Return some RTF text for non-vanilla jobs'
    annotations = []
    text = ''
    if job['Weird']: annotations.append('Unorthodox')
    if job['WIP']: annotations.append('WIP')
    if len(annotations) > 0:
        text = '\\par\\par\\f0\\fs12Ann: %s' % (' '.join(annotations))
    else: text = ''
    return text

###########################################################################

def AsAscii(text):
    'convert text into ASCII string'
    return unicodedata.normalize('NFKD', text).encode('ascii','ignore')

def AsFloat(text):
    'Convert text into a float'
    text = text.replace(',', '')
    return float(text)