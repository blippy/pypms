# FIXME - moving over to GetRows

import itertools
import timeit
import pdb

import win32com.client

from common import princ

class Struct:
    pass

def testmain():
    'Return an open connection to the database'
    conn = win32com.client.Dispatch(r'ADODB.Connection')
    conn.Open('PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=M:/Finance/camel/camel.mdb;')
    rs = win32com.client.Dispatch(r'ADODB.Recordset')
    sql = "SELECT Person, PersonNAME, IsStaff FROM tblEmployeeDetails"
    rs.Open(sql, conn, 1, 3)
    
    # FIXME - also consider using GetAll and GetArray
    recs = rs.GetRows()
    #Person,PersonNAME, PersonCompany, Active, Complete, IsStaff = recs
    fieldspec = [('Person', str), ('PersonNAME', str), ('IsStaff', bool)]
    for rec in itertools.izip( recs[0], recs[1], recs[2]):
        princ("***" + rec)
    t = Struct()
    
    princ(recs)
    princ(recs[1][2])
    rs.Close()
    conn.Close()

testmain()
princ('Finished')
