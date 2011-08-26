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
    
def timesheets(debug = False):
    jobs = [x['job'] for x in db.GetJobs().values() if x['Autoprint']]
    jobs.sort()
    for job in jobs:
        fname = '{0}\\timesheets\\{1}.rtf'.format(period.perioddir(), job)
        if not os.path.isfile(fname): continue
        if debug: print fname
        else: print_file(fname)

    
if  __name__ == "__main__":
    timesheets(debug = True)
    #work_statements()
    princ("Finished xls")
