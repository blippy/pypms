import datetime
import glob
import os

import win32api, win32com.client

import common
import db
import period
from common import princ

def print_file(filename):
    cmd = '"C:\Program Files\Windows NT\Accessories\wordpad.exe" /p ' + filename
    os.system(cmd)
    #win32api.ShellExecute(0, "print", fname, None, ".", 0)

def work_statements():
    'Print workstatements'
    pattern = '{0}\\statements\\*.rtf'.format(period.perioddir())
    files = glob.glob(pattern)
    files.sort()
    for fname in files:
        f0 = os.path.basename(fname)[0]
        if f0 in ['0','5']: continue # special files that don't need to be printed
        print_file(fname)
    
def timesheets():
    #p = period.Period(usePrev = True)
    #p.decMonth()
    jobs = db.records(['job'], 'SELECT job FROM jobs WHERE Autoprint=Yes ORDER BY job')
    jobs = [x[0] for x in jobs] # flatten the jobs list
    for job in jobs:
        pattern = '{0}\\timesheets\\{1}*.rtf'.format(period.perioddir(), job)
        files = glob.glob(pattern)
        files.sort()
        for fname in files: print_file(fname)

    
if  __name__ == "__main__":
    #timesheets()
    work_statements()
    princ("Finished xls")
