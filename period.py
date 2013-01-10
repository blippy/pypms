# registry example from 
# http://www.blog.pythonlibrary.org/2010/03/20/pythons-_winreg-editing-the-windows-registry/


import calendar
import datetime
import math
import os
import pdb
import sys, traceback

import common
from common import princ
import registry

    
###########################################################################

def create_text_report(filename, lines):    
    
    
    def at(lst, idx):
        try:
            v = lst[i]
        except IndexError:
            v = ""
        return v
    
    # work out how wide the columns are supposed to be
    num_cols = len(lines[0])
    col_widths = [0] * num_cols # assume width 0 to start with
    for row in lines:
        #princ(row)
        for i in range(0, num_cols):
            #princ(i)
            v = at(row, i)
            if isinstance(v, str):
                col_width = len(v)
            else:
                col_width = 9                
            col_widths[i] = max(col_widths[i], col_width)
    #princ(col_widths)

    # create output
    output = []
    for row in lines:
        fields = []
        for i in range(0, num_cols):
            v = at(row, i)
            fmt = '{0:>' + str(col_widths[i])
            if type(v) is float: fmt += '.2f'
            fmt += '}'
            text = fmt.format(v)
            fields.append(text)
        #cols = 
                
        #cols = map(fmt_text, row)
        line = ' '.join(fields)
        output.append(line)
        
    save_report(filename, output)

###########################################################################
# basic initialisation, getting, and setting

g_period = None #init_global_period()

def tuple():
    'Retrieve the global period as the tuple: y, '
    global g_period
    return g_period[0], g_period[1]

def pformat(fmt)    :
    'Format the current period according to a specified format'
    y, m = tuple()
    return fmt.format(y, m)

def describe():
    princ(pformat('YEAR: {0}, MONTH: {1}'))

def yyyymm():
    'Return the period in the form 2010-01'
    #result = "%04d-%02d" % (self.y, self.m)
    return pformat('{0:04d}-{1:02d}')
    #return result
 
def mmmmyyyy():
    y, m = tuple()
    return '{0} {1}'.format(calendar.month_name[m], y)

def XXXfirst(self):
    'Return the first day of the invoice billing period in the form 01/01/2010'
    return pformat('01/{1:02d}/{0:4d}')

    
def dim():
    'Return the number of days in the year/month'
    y, m = tuple()
    return calendar.monthrange(y, m)[1]
    
def stuple(y, m):
    'Set the global period from y,m'
    global g_period
    g_period = (y, m)
    registry.set_reg_value("Period", yyyymm())
    
def init_global_period():
    'Initialise global period'
    try:
        v = registry.get_reg_key("Period")
        ym = str(v[0])
        y = int(ym[:4])
        m = int(ym[5:])
        #p.setym(y, m)
    except WindowsError:
        p = datetime.datetime.now()
        y = p.year
        m = p.month
    stuple(y, m)

init_global_period()




###########################################################################

def inc(num_months):
    'Increase the global period by NUM_MONTHS'
    y, m = tuple()
    month_num = y*12 + m-1
    month_num += num_months
    y = int(math.floor(month_num/12))
    m = month_num - y*12 + 1
    stuple(y, m)
    
def is_weekend(d):
    'Determine if a day is a weekday or weekend'
    y, m = tuple()
    v = calendar.weekday(y, m, d) # mon is 0
    v = v > 4
    return v

###########################################################################
# directories dependent on period



def perioddir():
    return common.outroot + "\\" + yyyymm()
    
def reportdir():
    'Return the report directory'
    dir =  perioddir() + "\\reports"
    common.makedirs(dir)
    return dir

# TODO - use this more extensively
def reportfile(fname):
    'Return a full filename for a report'
    return reportdir() + '\\' + fname

# TODO - use this more
def save_report(filename, text):
    fullname = reportfile(filename)
    if type(text) is list:
        text = '\r\n'.join(text)
    common.spit(fullname, text)
    

 

    
###########################################################################
if __name__ == "__main__":
    #princ("Global period is:" + g_period.yyyymm())
    #p = Period()
    describe()
    inc(-12)
    describe()
    inc(12*13)
    describe()
    stuple(2013,1)
    print yyyymm()
    print mmmmyyyy()
    princ('Finished')