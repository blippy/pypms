import codecs
import datetime, json, logging, pprint, os, unicodedata

import data
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
    
def reportdir(p):
    'Return the report directory'
    dir = outroot + "\\" + p.yyyymm() + "\\reports"
    makedirs(dir)
    return dir


# FIXME - ensure all file saving goes though this function
# FIXME - should create intermediate directories if necessary
def spit(fname, text):
    'Write TEXT to file FNAME'
    #with open(fname, "w") as f: f.write(text)
    #with codecs.open(fname, "w", "utf-8") as f: f.write(text)
    with codecs.open(fname, "w", "Latin-1") as f: f.write(text)

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
    '''Convert text into ASCII string. Useful resource:
    http://www.peterbe.com/plog/unicode-to-ascii'''
    return unicodedata.normalize('NFKD', text).encode('ascii','ignore')

def AsFloat(text):
    'Convert text into a float'
    text = text.replace(',', '')
    return float(text)

###########################################################################

def run_current(main_func):
    'A standard run routine calling using the current yaml'
    # FIXME - use this function in many places, e.g. maininv.py
    d = data.Data()
    d.restore()
    main_func(d)
    d.store()