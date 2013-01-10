from __future__ import print_function

import codecs
import datetime
import decimal
import logging
import os
import pprint
import time
import unicodedata






###########################################################################

    
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
            raise KeyError("ERR: missing job")
 



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
        raise KeyError("ERR: missing task")
        

###########################################################################
# file and directory handling utilties    

def makedirs(path):
    path = os.path.abspath(path)
    if not os.path.exists(path): os.makedirs(path)

outroot = 'M:\\Finance\\pypms'
makedirs(outroot)


# TODO - should create intermediate directories if necessary
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
    print(str(text))
    
log_level = logging.ERROR
log_level = logging.INFO

logging.basicConfig(level=log_level,
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
# functional routines

def find(item, sequence, key):
    for el in sequence:
        if item == key(el):
            return el
    return None

    
###########################################################################
# list functions

def unique(alist):
    result = []
    for el in alist:
        if el not in result: result.append(el)
    return result

###########################################################################
# aggregation routines

def mkKeyFunc(fieldName):
    def func(rec): return rec[fieldName]
    return func

def aggregate(seq, keyfunc):
    #TODO prolly use this extensively
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
    # TODO - make more use of this function
    total = 0.0
    for el in seq:
        v = keyfunc(el)
        if test(el): total +=v
    return total

def summate_lod(lod, key):
    "Give a list of dictionaries, sum the values for a given KEY"
    return sum(map(lambda x: x[key], lod))

def summate_to_dict(seq, keyfunc, valuefunc):
    "Return a dictionary with key values determined by KEYFUNC, and values determined by VALUEFUNC, over SEQ"
    result = {}
    for el in seq:
        dplus(result, keyfunc(el), valuefunc(el))
    return result
    
def summate_cols(matrix):
    "Sum a list of list of values. E.g. summate_cols([[1,2],[3,4]]) # [4,6]"
    totals = []
    for c in xrange(len(matrix[0])):
        total = 0.0
        for r in matrix:
            try: v = AsFloat(r[c])
            except ValueError: v = 0.0
            total += v
        totals.append(total)
    return totals




###########################################################################
# dictionay functions

# TODO - use these functions extensively


# TODO ought to be able to remove this function entirely and use
# either get() or setdefault()
# http://wiki.python.org/moin/KeyError
def dget(dictionary_name, key, value = 0.0):
    '''Obtain a value from a dictionay, using default VALUE if key not found
    or value is None'''
    try: 
        result = dictionary_name[key]
    except KeyError:
        result = value
    if result is None: result = value
    return result

def dplus(dictionary_name, key, value):
    '''Accumulate a value for a dictionary, using 0.00 if the key is not already defined'''
    if not dictionary_name.has_key(key): dictionary_name[key] = 0.0
    dictionary_name[key] += value
    

   
def mapdict(d, fields):
    return map(lambda x: d[x], fields)

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
    if text is None: return 0.0
    if text == '': return 0.0
    if type(text) == decimal.Decimal:
        return float(text)
    elif type(text) == float:
        return text
    elif type(text) == int:
        return float(text)
    else:
        text = text.replace(',', '')    
        return float(text)

###########################################################################
# functions difficult to classify


def tri(truth, true_result, false_result):
    'Simulate the C ?: operator'
    if truth:
        return true_result
    else:
        return false_result
 


###########################################################################
# common operations on GUIs

def empty_wxgrid(grid):
    num_rows = grid.GetNumberRows()
    if num_rows >0:
        grid.DeleteRows(0, num_rows)

def rectify_num_grid_columns(grid, n):
    num_cols = grid.GetNumberCols()
    num_cols_required = n
    if num_cols_required > num_cols:
        grid.AppendCols(num_cols_required - num_cols)
    elif num_cols_required < num_cols:
        grid.DeleteCols(n - 1, num_cols - num_cols_required)

###########################################################################
# timestamp

__timestamp = None

def reset_timestamp():
    global __timestamp
    __timestamp = None
    
def get_timestamp():
    global __timestamp
    if __timestamp is None:
        __timestamp = datetime.datetime.today().strftime("%d %B %Y %H:%M:%S")
    return __timestamp
