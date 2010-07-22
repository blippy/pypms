'''
Write something useful here
'''

import db
import data
import period


cache = None
#per = None


###########################################################################


def bye():
    'Save cache and exit'
    #data.save(cache)
    #cache.close()
    print "Quitting"
    quit()
    
def dbg():
    'Fetch data for the action period into the cache.'
    db.fetch(cache)
    
#def load(): cache = data.open()

def perp(): 
    'Set the action period to previous invoicing period.'
    p = period.Period(usePrev = True)
    cache['period'] = p
    print 'Action period set to', p.mmmmyyyy()

#def save(): data.save(cache)
    
###########################################################################

def init():
    global cache
    #cache = data.load()
    cache = data.open()
    perp()


###########################################################################

if  __name__ == "__main__":
    init()
