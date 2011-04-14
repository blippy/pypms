# TODO now
# http://www.ironpython.info/index.php/ADODB_API for example usage

import itertools
import timeit
import pdb

#import win32com.client
import adodbapi

from common import princ

class Struct:
    pass


def test():
    def foo(*args):
        print args
    foo(1,2,3)
    
    for a,b in zip([1,2,3], [4,5,6]):
        print a,b

def test1():
    c = TableCache()
    c.populate()
    print c.rows

if  __name__ == "__main__": 
    test()
    princ('Finished')
