# information stored in the database

import datetime, pdb
from itertools import izip

import win32com.client

import common
from common import AsAscii


def DbOpen():
    'Return an open connection to the database'
    conn = win32com.client.Dispatch(r'ADODB.Connection')
    conn.Open('PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=M:/Finance/camel/camel.mdb;')
    return conn


def IsEmpty(rs): return rs.EOF or rs.BOF # is a recordset empty?

# FIXME - consider using the ADODB functions GetAll() or GetRows()

def ForEachRecord(sql, func):
    'Do the FUNC for each record in a recordset returned by a SQL statment'
    try:
        conn = DbOpen()
        rs = win32com.client.Dispatch(r'ADODB.Recordset')
        rs.Open(sql, conn, 1, 3)
        if not IsEmpty(rs): rs.MoveFirst()  
        while not IsEmpty(rs):
            func(rs)            
            rs.MoveNext()    
        rs.Close()
    finally:
        conn.Close()

def records(fieldnames, sql):
    values = []
    def func(rs): 
        rec = [rs.Fields(fieldname).Value for fieldname in fieldnames]
        values.append(rec)
    ForEachRecord(sql, func)
    return values
 

def XXXSingleColumn(sql):
    'Assume that the SQL statement just returns a single field. Obtain the values for that field'
    values = []
    def func(rs): 
        rec = [rs.Fields(0).Value for fieldname in fieldnames]
        values.append(rec)
    ForEachRecord(sql, func)
    return values
    
    
    
def ExecuteSql(sql):
    try:
        conn = DbOpen()
        conn.execute(sql)
    finally:
        conn.Close()
   

def RecordsList(sql, fieldspec):
    result = []
    fieldnames = [x[0] for x in fieldspec]
    for r in records(fieldnames, sql):
        recs = {}
        for (fieldname, fieldtype), fieldvalue in izip(fieldspec, r):
            recs[fieldname] = fieldtype(fieldvalue)
        result.append(recs)
    return result
  
def GetEmployees(p):
    sql = 'SELECT * FROM tblEmployeeDetails'
    fieldspec = [('Person', str), ('PersonNAME', str), ('IsStaff', bool)]
    recs = RecordsList(sql, fieldspec)
    employees = {}    
    for r in recs: employees[r['Person']] = r
    return employees

def GetInvoices(d, field_list):
    sql = "SELECT * FROM tblInvoice WHERE InvBillingPeriod='" +  d['period'].mmmmyyyy() + "'"
    return records(field_list, sql)

    
def GetJobs():
    sql = 'SELECT * FROM jobs'
    fieldspec = [('ID', int), ('job', str), ('title', str), ('address', AsAscii), ('references', AsAscii), ('briefclient', str), 
        ('active', bool), ('vatable', bool), ('exp_factor', float), ('WIP', bool), ('Weird', bool), 
        ('Autoprint', bool), ('TsApprover', AsAscii)]      
    recs = RecordsList(sql, fieldspec)
    jobs = {}
    for r in recs: jobs[r['job']] = r
    return jobs
 



def GetTasks(p):
    sql = 'SELECT * FROM tblTasks ORDER BY JobCode, TaskNo'
    fieldspec = [('JobCode', str), ('JCDescription', str), ('TaskNo', str), ('TaskDes', str)]
    recs = RecordsList(sql, fieldspec)
    tasks = {}
    for r in recs: tasks[ (r['JobCode'], r['TaskNo']) ] = r
    return tasks
   
def GetTimeitems(p):
    fmt =  'SELECT * FROM tblTimeItems '
    fmt += 'WHERE TimeVal<>0 AND LEN(JobCode) > 0 '
    fmt += 'AND Month([DateVal]) = %d and Year([DateVal]) = %d '
    fmt += 'ORDER BY JobCode, Task, Person, DateVal'
    sql = fmt % (p.m , p.y)
    def StdDate(d):
        d1 = datetime.date(d.year, d.month, d.day)
        return d1.strftime('%Y-%m-%d')
    
    fieldspec = [('JobCode', str), ('Person', str), ('DateVal', StdDate), ('TimeVal', float), ('Task', str), ('WorkDone', unicode)]
    recs = RecordsList(sql, fieldspec)
    return recs



def GetCharges(p):
    sql = 'SELECT JobCode, TaskNo, Person, PersonCharge FROM tblCharges;'
    fieldspec = [('JobCode', str), ('TaskNo', str), ('Person', str), ('PersonCharge', float)]
    recs = RecordsList(sql, fieldspec)
    charges = {}
    for r in recs: charges[ (r['JobCode'], r['TaskNo'], r['Person'] )] = r['PersonCharge']
    return charges

def fetch(d):
    p = d['period']
    d['employees'] = GetEmployees(p)
    d['jobs'] = GetJobs()
    d['tasks'] = GetTasks(p)
    d['timeItems'] = GetTimeitems(p)
    d['charges'] = GetCharges(p)
    d['auto_invoices'] = None
    d['manual_invoices'] = None
    d['invoice_tweaks'] = None
    d['expenses'] = None
    
if  __name__ == "__main__": 
    test()
    print 'Finished'