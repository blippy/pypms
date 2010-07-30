import codecs
import datetime, logging, pprint, os
import unicodedata


###########################################################################
    
class DataIntegrityError(Exception):
    def __init__(self, value):
        self.value = value

    def __str__(self):
        return repr(self.value)

###########################################################################
# file and directory handling utilties    

def makedirs(path):
    path = os.path.abspath(path)
    if not os.path.exists(path): os.makedirs(path)

outroot = 'M:\\Finance\\pypms\\out'
makedirs(outroot)


def camelxls(p):
    'Return the filename for the Camel Excel input file'
    return 'M:\\Finance\\camel\\%s\\camel-%s.xls' % (p.y, p.yyyymm())

def perioddir(p):
    return outroot + "\\" + p.yyyymm()
    
def reportdir(p):
    'Return the report directory'
    dir =  perioddir(p) + "\\reports"
    makedirs(dir)
    return dir

# FIXME - use this more extensively
def reportfile(p, fname):
    'Return a full filename for a report'
    return reportdir(p) + '\\' + fname

# FIXME - ensure all file saving goes though this function
# FIXME - should create intermediate directories if necessary
def spit(fname, text):
    'Write TEXT to file FNAME'
    #with open(fname, "w") as f: f.write(text)
    #with codecs.open(fname, "w", "utf-8") as f: f.write(text)
    with codecs.open(fname, "w", "Latin-1") as f: f.write(text)


# FIXME - use this more
def save_report(p, filename, text):
    fullname = reportfile(p, filename)
    spit(fullname, text)

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

def aggregate(seq, keyfunc):
    #FIXME prolly use this extensively
    dic = {}
    for x in seq:
        key = keyfunc(x)
        if not dic.has_key(key): dic[key] = []
        dic[key].append(x)
    
    lis=[]
    for k in dic.keys():
        lis.append((k, dic[k]))
    lis.sort(key = lambda x: x[0]) # sort on the key
    
    return lis
    
def summate(seq, keyfunc):
    # FIXME - make more use of this function
    total = 0.0
    for el in seq:
        v = keyfunc(el)
        total +=v
    return total
    
    
###########################################################################
# dictionay functions

# FIXME - use this extensively
def dget(dictionary_name, key, value = 0.0):
    '''Obtain a value from a dictionay, using default VALUE if key not found'''    
    try: 
        result = dictionary_name[key]
    except KeyError:
        result = value
    return result

# FIXME - use this extensively
def dplus(dictionary_name, key, value):
    '''Accumulate a value for a dictionary, using 0.00 if the key is not already defined'''
    if not dictionary_name.has_key(key): dictionary_name[key] = 0.0
    dictionary_name[key] += value
    

# FIXME - use this extensively
def combine_dict_keys(list_of_dicts):
    result = set()
    for dic in list_of_dicts:
        for key in dic.keys():
            result.add(key)
    result = list(result)
    result.sort()
    return result
    
###########################################################################

def AsAscii(text):
    '''Convert text into ASCII string. Useful resource:
    http://www.peterbe.com/plog/unicode-to-ascii'''
    if text is None:
        return ''
    elif type(text) == str:
        return text
    else:
        return unicodedata.normalize('NFKD', text).encode('ascii','ignore')

def AsFloat(text):
    'Convert text into a float'
    text = text.replace(',', '')
    if text == '': text = '0.0'
    return float(text)

###########################################################################
# functions difficult to classify


    
def tri(truth, true_result, false_result):
    'Simulate the C ?: operator'
    if truth:
        return true_result
    else:
        return false_result