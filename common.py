from __future__ import print_function

import codecs
import datetime
import decimal
import logging
import pprint
import os
import unicodedata

import _winreg





###########################################################################
    
class DataIntegrityError(Exception):
    def __init__(self, value):
        self.value = value

    def __str__(self):
        return repr(self.value)
    
    
# FIXME prolly need to use this more
def assert_job(d, job_code, source_info):
    "Require that a job exists in the database"
    if not d['jobs'].has_key(job_code):
            fmt = """
            ERR: Missing entry in 'jobs' table cache.
            Either source is wrong, you haven't entered the job in the jobs table, 
            or you need to refresh pydra's cache
            Job   : '{0}'
            Source: '{1}'
            """
            msg = fmt.format(job_code, source_info)
            logerror(msg)
            raise DataIntegrityError("ERR: missing job")
  
def assert_task(d, job_code, task_code, source_info):
    "Require that a task exist in the database"
    if not d['tasks'].has_key((job_code, task_code)):
        fmt = """
        ERR: Missing task entry in 'tasks' table cache.
        Either source is wrong, you haven't entered the task in the tblTasks table, 
        or you need to refresh pydra's cache
        Job   : '{0}'
        Task  : '{1}'
        Source: '{2}'
        """
        msg = fmt.format(job_code, task_code, source_info)
        logerror(msg)
        raise DataIntegrityError("ERR: missing task")
        

###########################################################################
# file and directory handling utilties    

def makedirs(path):
    path = os.path.abspath(path)
    if not os.path.exists(path): os.makedirs(path)

outroot = 'M:\\Finance\\pypms'
makedirs(outroot)


# FIXME - ensure all file saving goes though this function
# FIXME - should create intermediate directories if necessary
def spit(fname, text):
    'Write TEXT to file FNAME'
    #with open(fname, "w") as f: f.write(text)
    #with codecs.open(fname, "w", "utf-8") as f: f.write(text)
    with codecs.open(fname, "w", "Latin-1") as f: f.write(text)




###########################################################################
# logging

#_princ_func = print

# TODO eliminate this function
def princ(text):
    print(text)
    
logging.basicConfig(level=logging.ERROR,
                    format='%(asctime)s %(levelname)s %(message)s',
                    filename= outroot + '\\pypms.log',
                    filemode='a')
def logdebug(txt): logging.debug(txt)
def loginfo(txt): logging.info(txt)

def logerror(txt):
    princ(txt)
    logging.error(txt)
    
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
    
def always(*kargs): return True

def summate(seq, keyfunc, test = always):
    # FIXME - make more use of this function
    total = 0.0
    for el in seq:
        v = keyfunc(el)
        if test(el): total +=v
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

def AsInt(text):
    'Convert text into an integer'
    if text is None: return 0
    return int(text)

    
def AsFloat(text):
    'Convert text into a float'
    if text is None: return 0.0
    if text == '': return 0.0
    if type(text) == decimal.Decimal:
        return float(text)
    elif type(text) == float:
        return text
    else:
        text = text.replace(',', '')    
        return float(text)

###########################################################################
# registry functions

# binary values are kooky on Windows, as there is no default binary 
# registry type. They therefore have to be hand-crafted


__reg_root = 'Software\\Pydra'

def get_reg_key(key_name):
    global __reg_root
    key = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER, __reg_root, 0, _winreg.KEY_ALL_ACCESS)
    v = _winreg.QueryValueEx(key, key_name)
    _winreg.CloseKey(key)
    return v

def get_defaulted_reg_key(key_name, default):    
    try:
        v = get_reg_key(key_name)
    except WindowsError:
        v = default
    return v
    
def set_reg_value(key_name, new_value, reg_type = _winreg.REG_SZ):
    global __reg_root
    try:
        key = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER, __reg_root, 0, _winreg.KEY_ALL_ACCESS)
    except:
        key = _winreg.CreateKey(_winreg.HKEY_CURRENT_USER, __reg_root)
    _winreg.SetValueEx(key, key_name, 0, reg_type, new_value)
    _winreg.CloseKey(key)
    
def get_defaulted_binary_reg_key(key_name, default):
    try:
        raw = get_reg_key(key_name)
        result = raw[0] == 1
    except WindowsError:
        result = default
    return result

def set_binary_reg_value(key_name, value):
    set_reg_value(key_name, value, reg_type = _winreg.REG_DWORD)


###########################################################################
# functions difficult to classify


    
def tri(truth, true_result, false_result):
    'Simulate the C ?: operator'
    if truth:
        return true_result
    else:
        return false_result