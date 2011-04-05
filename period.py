# registry example from 
# http://www.blog.pythonlibrary.org/2010/03/20/pythons-_winreg-editing-the-windows-registry/


import calendar
import datetime
import math
import os
import pdb
import sys, traceback

#from _winreg import *

import common
from common import princ
import registry

###########################################################################

def now(): return datetime.datetime.now()


def is_weekend(y, m, d):
    'Determine if a day is a weekday or weekend'
    v = calendar.weekday(y, m, d) # mon is 0
    v = v > 4
    return v
    
###########################################################################


class Period:
    def __init__(self, usePrev = False):        
        self.setToCurrent()        
        if usePrev: self.decMonth()
        
    def __eq__(self, other):
        if not other: return False # maybe None was passed in as other
        return self.yyyymm() == other.yyyymm()
    
    def __ne__(self, other):
        return not self.__eq__(other)

    def setym(self, y, m):
        self.y = int(y)
        self.m = int(m)
        
    def setToCurrent(self):
        t = now()
        self.y = t.year
        self.m = t.month
        
    def describe(self):
        princ('YEAR: {0}, MONTH: {1}'.format(self.y, self.m))
        

    def ask(self, text, default):
        inp = raw_input(text)
        if len(inp) == 0: inp = default
        return inp

    def askNumber(self, fieldDescription, default, low, high):
        text = '%s (%d-%d) [%d]:' % (fieldDescription, low, high, default)
        while 1:
            try:
                value = self.ask(text, default)
                value = int(value)
                if value < low or value > high: raise ValueError
                break
            except ValueError:
                princ("Invalid input. Try again")
        return value

    def askYear(self):
        return self.askNumber('Year', now().year, 2000, 2020)
            
    def askMonth(self):
        return self.askNumber('Month', now().month, 1, 12)

    def decMonth(self):
        self.inc(-1)
            
    def inc(self, num_months = 1):
        month_num = self.y*12 + self.m-1
        month_num += num_months
        self.y = int(math.floor(month_num/12))
        self.m = month_num - self.y*12 + 1
        
            
    def inputPeriod(self):
        self.setToCurrent()
        while 1:
            inp = raw_input("Use (r)ecent (c)urrent or (o)ther billing period ([r]/c/o)?")
            if inp == "" or inp =="r":
                self.decMonth()                
                break
            elif inp == "c":
                break
            elif inp == "o":
                self.y = self.askYear()
                self.m = self.askMonth()
                break
            else: 
                princ("Invalid input. Try again" )   
        

    def within(self, d):
        return d.year == self.y and d.month == self.m

    def yyyymm(self):
        'Return the period in the form 2010-01'
        result = "%04d-%02d" % (self.y, self.m)
        return result
    
    def mmmmyyyy(self):
        ' Return the period in the form March 2010'
        d = datetime.date(self.y, self.m, 1)
        result = d.strftime('%B %Y')
        return result
    
    def first(self):
        'Return the first day of the invoice billing period in the form 01/01/2010'
        result = '01/%02d/%04d' % (self.m, self.y)
        return result
    
    def dim(self):
        'Return the number of days in the year/month'
        return calendar.monthrange(self.y, self.m)[1]
    
    def last(self):
        'Return the last day of the invoice billing period in the form 31/01/2010'
        result = '%02d/%02d/%04d' % (self.dim(), self.m, self.y)
        return result
        

###########################################################################
# directories dependent on period

def camelxls(p = None):
    'Return the filename for the Camel Excel input file'
    
    assert(p) # let's demand a period TODO
    if p is None: p = Period(usePrev = True)
    
    return 'M:\\Finance\\camel\\%s\\camel-%s.xls' % (p.y, p.yyyymm())

def perioddir(p = None):
    assert(p) # let's demain a period TODO
    if p is None: p = Period(usePrev = True)
    
    return common.outroot + "\\" + p.yyyymm()
    
def reportdir(p = None):
    'Return the report directory'
    dir =  perioddir(p) + "\\reports"
    common.makedirs(dir)
    return dir

# FIXME - use this more extensively
def reportfile(p, fname):
    'Return a full filename for a report'
    return reportdir(p) + '\\' + fname

# FIXME - use this more
def save_report(p, filename, text):
    fullname = reportfile(p, filename)
    common.spit(fullname, text)
    
###########################################################################

def init_global_period():
    p = Period()
    try:
        v = registry.get_reg_key("Period")
        ym = str(v[0])
        y = ym[:4]
        m = ym[5:]
        p.setym(y, m)
    except WindowsError:
        pass
    return p

g_period = init_global_period()

def global_inc(num_months):
    'Increase the global period by NUM_MONTHS'
    global g_period
    g_period.inc(num_months)
    common.set_reg_value("Period", g_period.yyyymm())

    



if __name__ == "__main__":
    princ("Global period is:" + g_period.yyyymm())
    p = Period()
    p.describe()
    p.inc(-12)
    p.describe()
    princ('Finished')