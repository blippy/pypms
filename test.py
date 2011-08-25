# TODO now
# http://www.ironpython.info/index.php/ADODB_API for example usage

import itertools
import timeit
import pdb
#from sqlalchemy import *
import pyodbc
import db, dbnew


#import win32com.client
import adodbapi

from common import princ

class Struct:
    pass

def way1():
    db.records(['JobCode', 'Person'], 'tblTimeItems')
    
def way2():
    dbnew.TableCache('tblTimeItems')
def test():

        
    #print timeit.Timer('way1()', 'from __main__ import way1').timeit(number = 10) # takes about 30.2 s
    #print timeit.Timer('way2()', 'from __main__ import way2').timeit(number = 10) #takes about 4.1 s
    #way2()
    t = dbnew.TableCache('tblTimeItems')
    print t.inmemory_rows



if  __name__ == "__main__": 
    test()
    princ('Finished')
