# registry example from 
# http://www.blog.pythonlibrary.org/2010/03/20/pythons-_winreg-editing-the-windows-registry/


import calendar
import datetime
import math
import os
import pdb
import sys, traceback

from _winreg import *

#import common

def now(): return datetime.datetime.now()


def is_weekend(y, m, d):
    'Determine if a day is a weekday or weekend'
    v = calendar.weekday(y, m, d) # mon is 0
    v = v > 4
    return v
    



class Period:
    def __init__(self, usePrev = False):        
        self.setToCurrent()        
        if usePrev: self.decMonth()
        
    def __eq__(self, other):
        #pdb.set_trace()
        #print "entering equality test"
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
        print 'YEAR: ', self.y
        print 'MONTH:', self.m
        

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
                print "Invalid input. Try again"
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
                print "Invalid input. Try again"    
        


#    def outDir(self):
#        full = "%s\\%s\\" % (common.outroot, self.yyyymm())
#        full = os.path.abspath(full)
#        common.makedirs(full)
#        return full

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
        

def init_global_period():
    p = Period()
    try:
        key = OpenKey(HKEY_CURRENT_USER, 'Software\\Pydra', 0, KEY_ALL_ACCESS)
        v = QueryValueEx(key, "Period")
        CloseKey(key)
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
    keyVal = 'Software\\Pydra'
    try:
        key = OpenKey(HKEY_CURRENT_USER, keyVal, 0, KEY_ALL_ACCESS)
    except:
        key = CreateKey(HKEY_CURRENT_USER, keyVal)
    SetValueEx(key, "Period", 0, REG_SZ, g_period.yyyymm())
    CloseKey(key)

    



if __name__ == "__main__":
    print "Global period is:", g_period.yyyymm()
    p = Period()
    p.describe()
    p.inc(-12)
    p.describe()
    print 'Finished'