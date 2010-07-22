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
    d = Data()
    d.restore()
    main_func(d)
    d.store()
   
###########################################################################

def main():
    'Just used for testing purposes'
    
    usePrevMonth = False # set or unset according to taste
    p = period.Period()
    if usePrevMonth: p.decMonth()
    else: p.inputPeriod()
    
    d = Data()
    d.load()
    d.store()
    
if  __name__ == "__main__":
    main()
    print 'Finished'
