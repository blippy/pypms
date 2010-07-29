'''
Write something useful here
'''

import common
from common import dget
import db
import data
import expenses
import period
import push
import statements
import timesheets


cache = None


###########################################################################


def bye():
    'Save cache and exit'
    global cache
    #data.save(cache)
    #cache.close()
    print "Quitting"
    quit()
    
def dbg():
    'Fetch data for the action period into the cache.'
    global cache
    db.fetch(cache)
    
#def load(): cache = data.open()

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
    if p!= prev_period:
        print 'New period detected. Trashing current cache.'
        for k in cache.keys(): del cache[k]
    cache['period'] = p
    print 'Action period set to', p.mmmmyyyy()

def stage1():
    dbg()  
    expg()
    exps()
    works()
    times()
    
def stage2():
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

# def save(): data.save(cache)
    
###########################################################################

def init():
    global cache
    #cache = data.load()
    cache = data.open()
    perp()


###########################################################################

if  __name__ == "__main__":
    init()
