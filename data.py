import anydbm
import os
import pdb
import pprint
import shelve
import shutil

import common

###########################################################################

# FIXME deprecate this module

__storage_filename = 'M:\\Finance\\pypms\\out\\current.shl'

def XXXopen():
    print 'Loading cache'
    try:
        dbase = shelve.open(__storage_filename, writeback = True)
    except anydbm.error:
        print "Something is wrong with the shelve. Creating a new one"
        os.remove(__storage_filename)
        dbase = shelve.open(__storage_filename, flag = 'n', writeback = True)
    return dbase

def XXXbackup(per):
    shutil.copy(__storage_filename, common.perioddir(per) + "\\current.shl")

def XXXrun_current(main_func):
    'A standard run routine calling using the current yaml'
    # FIXME - use this function in many places, e.g. maininv.py
    d = open()
    main_func(d)
   
###########################################################################

def main(d):
    'Just used for testing purposes'
    print "Didn't do anything"        
    
if  __name__ == "__main__":
    run_current(main)
    print 'Finished'
