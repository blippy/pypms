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
    c = TableCache()
    c.populate()
    print c.rows

if  __name__ == "__main__": 
    test()
    princ('Finished')
