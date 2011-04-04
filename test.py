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

def testmain():
    'Return an open connection to the database'
    conStr = r'PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=M:/Finance/camel/camel.mdb;'    
    con = adodbapi.connect(conStr)
    cursor = con.cursor()
    sql = "SELECT Person, PersonNAME, IsStaff FROM tblEmployeeDetails"
    cursor.execute(sql)
    ds = cursor.fetchall()
    for row in ds:
        print type(row), row
    con.close()



testmain()
princ('Finished')
