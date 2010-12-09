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
#import data
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
    #cache.close()
    print "Parting is such sweet sorrow"

atexit.register(quit_pydra)    

###########################################################################


def XXXbackup():
    'Make a backup of of the cache, storing it in the relevant period.'
    # FIXME remove this module
    #global cache
    #per = copy.copy(cache['period'])
    #pdb.set_trace()
    #cache.close()
    #data.backup(per)
    #cache = data.open()
    pass
    
def bye():
    'Save cache and exit'
    global cache
    print "Quitting"
    quit()
    
def dbg():
    'Fetch data for the action period into the cache.'
    global cache
    cache = db.fetch()
    

def XXXexpg():
    'Fetch expenses from spreadsheet for period.'
    global cache
    expenses.read_expenses(cache)
    
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
        
def gui():
    root = Tk()
    #root.withdraw()
    def finito(): print 'Exiting GUI' ; root.quit() ; root.withdraw()
    root.protocol("WM_DELETE_WINDOW", finito)

    frameMain = Frame(root,borderwidth=2,relief=SUNKEN)
    frameMain.pack(side=LEFT,padx=5,pady=5)
    #Label(root,text='I am a frame').pack(side=LEFT)
    #m1 = PanedWindow(root, orient=VERTICAL)
    
    #m1.pack()
    year = StringVar()
    year.set("2010")
    def print_it(x): print x
    OptionMenu(root, year, "2010","2011", command=print_it).pack(side=LEFT)
    month = StringVar()
    
    monthTuple = ('Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec')
    optMonth = OptionMenu(root, month, *monthTuple, command=print_it)
    currMonth = monthTuple[cache['period'].m -1]
    month.set(currMonth)
    optMonth.pack(side=LEFT)
    
    
    def hello(): print "hello"
    # create a toplevel menu
    menubar = Menu(root)
    data_menu = Menu(menubar, tearoff = 0)
    menubar.add_cascade(label='Data', menu = data_menu)
    data_menu.add_command(label="Expenses",  command=hello)
    data_menu.add_command(label="Hello",  command=hello)
    
    misc_menu = Menu(menubar, tearoff = 0)
    menubar.add_cascade(label='Misc', menu = misc_menu)
    misc_menu.add_command(label="Quit", command=finito)
    

    # display the menu
    root.config(menu=menubar)

    root.title('Hydra')
    root.mainloop()
    #root.withdraw()
   

def mobils():
    'Create the Mobil work statement'
    global cache
    mobil.main(cache)
    
def perp(): 
    'Set the action period to previous invoicing period.'
    return True
    # TODO rest of function unneeded
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

def init():
    global cache
    #cache = data.open()
    perp()


###########################################################################

if  __name__ == "__main__":
    init()
