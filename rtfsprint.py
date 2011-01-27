import datetime, glob

import win32api, win32com.client

import common
import db


def main():
    p = period.Period()
    p.decMonth()
    jobs = db.records(['job'], 'SELECT job FROM jobs WHERE Autoprint=Yes ORDER BY job')
    jobs = [x[0] for x in jobs] # flatten the jobs list
    for job in jobs:
        # pattern = '%s\\timesheets\\%s*.rtf' % (p.outDir(), job )
        pattern = '{0}\\timesheets\\{1}*.rtf'.format(common.preioddir(p), job)
        files = glob.glob(pattern)
        files.sort()
        for fname in files:
            win32api.ShellExecute(0, "print", fname, None, ".", 0)

    
if  __name__ == "__main__":
    main()
    print "Finished xls"
