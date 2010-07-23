import pdb
#import pickle
import pprint
import shelve

#import yaml

import db, expenses, period





###########################################################################


    
def XXX__init__(self, p = None):
    if not p: p = period.Period(usePrev = True)
    self.p = p # period
    self.restore()

    



###########################################################################


__storage_filename = 'M:\\Finance\\pypms\\out\\current.shl'

def open():
    print 'Loading cache'
    dbase = shelve.open(__storage_filename, writeback = True)
    #with open(__storage_filename, 'rb') as f:
    #    pkl = pickle.load(f)
    return dbase
    
def XXXsave(pkl):
    print 'Saving cache'
    with open(__storage_filename, 'wb') as f:
        pickle.dump(pkl, f, pickle.HIGHEST_PROTOCOL)
        


def run_current(main_func):
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
