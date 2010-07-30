'''
PMS automation facility written in Python.
'''

import copy
import pdb

from common import dget
import db
import data
import expenses
import period
import push
import rtfsprint
import statements
import timesheets


cache = None


###########################################################################


def backup():
    'Make a backup of of the cache, storing it in the relevant period.'
    global cache
    per = copy.copy(cache['period'])
    #pdb.set_trace()
    cache.close()
    data.backup(per)
    cache = data.open()
    
def bye():
    'Save cache and exit'
    global cache
    print "Quitting"
    quit()
    
def dbg():
    'Fetch data for the action period into the cache.'
    global cache
    db.fetch(cache)
    

def expg():
    'Fetch expenses from spreadsheet for period.'
    global cache
    expenses.read_expenses(cache)
    
def exps():
    'Create expenses spreadsheet.'
    global cache
    expenses.create_report(cache)

def perp(): 
    'Set the action period to previous invoicing period.'
    global cache
    
    #print cache
    prev_period = dget(cache, 'period', None)
    p = period.Period(usePrev = True)
    if p != prev_period:
        print 'New period detected. Trashing current cache.'
        for k in cache.keys(): 
            del cache[k]
    cache['period'] = p
    print 'Action period set to', p.mmmmyyyy()

def ptimes():
    'Print the timesheets'
    rtfsprint.main()
    
def stage1():
    dbg()  
    times()
    
def stage2():
    expg()
    exps()
    works()
    
def stage3():
    global cache
    push.main(cache)

def times():
    'Create timesheets.'
    global cache
    timesheets.main(cache)
    
def works():
    'Create the invoice statements.'
    global cache
    statements.main(cache)

    
###########################################################################

def init():
    global cache
    print 'Pydra Arthritic Aardvark 2010-07-31'
    cache = data.open()
    perp()


###########################################################################

if  __name__ == "__main__":
    init()
