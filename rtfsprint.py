"""Print the RTF files"""

import glob
import os

import common
import db
import period

def print_file(filename):
    "Print a file using wordpad"
    cmd = '"C:\Program Files\Windows NT\Accessories\wordpad.exe" /p ' + filename
    os.system(cmd)

def work_statements():
    'Print workstatements'
    pattern = '{0}\\statements\\*.rtf'.format(period.perioddir())
    files = glob.glob(pattern)
    files.sort()
    for fname in files:
        
        # bail out on special files
        fname0 = os.path.basename(fname)[0]
        if fname0 in ['0','5']: 
            continue
        
        print_file(fname)
    
def timesheets(debug = False):
    "Print timesheets"
    jobs = [x['job'] for x in db.GetJobs().values() if x['Autoprint']]
    jobs.sort()
    for job in jobs:
        fname = '{0}\\timesheets\\{1}.rtf'.format(period.perioddir(), job)
        if not os.path.isfile(fname):
            continue
        if debug:
            print fname
        else: print_file(fname)

    
if  __name__ == "__main__":
    work_statements()
    timesheets(debug = True)
    common.princ("Finished xls")
