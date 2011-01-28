'''
PMS automation facility written in Python.
'''

from Tkinter import *

import atexit
import copy
import msvcrt
import os
import pdb
import pydoc

import budget
from common import dget
import db
import expenses
import health
import html
import mobil
import period
import push
import rtfsprint
import statements
import timesheets


###########################################################################
# important module instantiation routines

cache = {}

def quit_pydra():
    print "Parting is such sweet sorrow"

atexit.register(quit_pydra)    

###########################################################################


    
def bye():
    'Save cache and exit'
    global cache
    print "Quitting"
    quit()
    
def dbg():
    'Fetch data for the action period into the cache.'
    global cache
    cache = db.fetch()
    

    
def exps():
    'Create expenses spreadsheet.'
    global cache
    expenses.create_report(cache)
    
def expv():
    'View Expenses'
    records = []
    for el in cache['expenses']:
        line = [str(el[key]) for key in ['JobCode', 'Name', 'Amount']]
        line = ' '.join(line)
        records.append(line)
        
    page = 0
    while True:
        os.system('CLS')
        for line in records[page*20:page*20+20]:
            print line
        print
        print 'q-Quit  SPC,n-Next  p-Prev'
        ch = msvcrt.getch().lower()
        if ch == 'q': break
        if ch == 'n' or ch == ' ': page = min( int(len(line)/20), page + 1)
        if ch == 'p': page = max(0, page - 1)
        
    #pydoc.ttypager(txt)
        

def mobils():
    'Create the Mobil work statement'
    global cache
    mobil.main(cache)
    

def ptimes():
    'Print the timesheets'
    rtfsprint.main()
    
def stage1():
    global cache
    dbg()  
    times()

    
def stage2():
    #expg()
    exps()
    works()
    
def stage3():
    global cache
    push.main(cache)
    
def stage4():
    'Create sundry reports'
    global cache
    budget.main(cache)
    health.main(cache)
    html.main()
    mobils()
    
def allstages():
    'Run stages 1-4'
    stage1()
    stage2()
    stage3()
    stage4()
    #cache.sync() # write the cache back out to disk

def times():
    'Create timesheets.'
    global cache
    timesheets.main(cache)
    
def works():
    'Create the invoice statements.'
    global cache
    statements.main(cache)

    
###########################################################################



###########################################################################

if  __name__ == "__main__":
    print "Didn't do anything"
