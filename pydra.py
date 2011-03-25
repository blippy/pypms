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
from common import dget, princ
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



###########################################################################




    
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
            princ(line)
        princ('')
        princ('q-Quit  SPC,n-Next  p-Prev')
        ch = msvcrt.getch().lower()
        if ch == 'q': break
        if ch == 'n' or ch == ' ': page = min( int(len(line)/20), page + 1)
        if ch == 'p': page = max(0, page - 1)
        
    #pydoc.ttypager(txt)
        



def ptimes():
    'Print the timesheets'
    rtfsprint.main()
    

    
def allstages(p):
    'Run stages 1-4'
    global cache
    cache = db.fetch(p)
    
    # stage 1
    timesheets.main(cache)
    
    # stage 2
    exps()
    statements.main(cache)
    
    # stage 3
    push.main(cache)

    # stage 4 - sundry reports
    budget.main(cache)
    health.main(cache)
    html.main()
    mobil.main(cache)

    

    

    
###########################################################################



###########################################################################

if  __name__ == "__main__":
    princ("Didn't do anything")
