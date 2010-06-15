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


# FIXME - ensure all file saving goes though this function
def spit(fname, text):
    'Write TEXT to file FNAME'
    with open(fname, "w") as f: f.write(text)

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

def AsAscii(text):
    'convert text into ASCII string'
    # http://www.peterbe.com/plog/unicode-to-ascii
    return unicodedata.normalize('NFKD', text).encode('ascii','ignore')

def AsFloat(text):
    'Convert text into a float'
    text = text.replace(',', '')
    return float(text)