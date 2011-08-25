# information stored in the database
# TODO hopefully everything below will be obsoleted eventually

import datetime, pdb
from itertools import izip
import os
import pickle

import adodbapi
import win32com.client

import common
from common import AsAscii, AsFloat, AsInt, princ, print_timing
import expenses
import period



 
###########################################################################

tbl_billing = None

###########################################################################
        
def DbOpen():
    'Return an open connection to the database'
    conn = win32com.client.Dispatch(r'ADODB.Connection')
    conn.Open('PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=M:/Finance/camel/camel.mdb;')
    return conn


def IsEmpty(rs): return rs.EOF or rs.BOF # is a recordset empty?

# FIXME - consider using the ADODB functions GetAll() or GetRows()

def ForEachRecord(sql, func):
    # TODO deprecate
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
    # TODO deprecate
    values = []
    def func(rs): 
        rec = [rs.Fields(fieldname).Value for fieldname in fieldnames]
        values.append(rec)
    ForEachRecord(sql, func)
    return values
 

def fetch_all(sql):
    # preferred method from 25-Aug-2011 - it's much faster
    try:
        conStr = r'PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=M:/Finance/camel/camel.mdb;'    
        con = adodbapi.connect(conStr)
        cursor = con.cursor()
        #sql = 'SELECT * FROM tblBilling'
        cursor.execute(sql)
        ds = cursor.fetchall()
        rows = [row for row in ds]
        cursor.close()
    finally:
        con.close()
    return rows
        
    
    
def ExecuteSql(sql):
    try:
        conn = DbOpen()
        conn.execute(sql)
    finally:
        conn.Close()
   

def RecordsList(sql, fieldspec):
    # TODO deprecate
    result = []
    fieldnames = [x[0] for x in fieldspec]
    for r in records(fieldnames, sql):
        recs = {}
        for (fieldname, fieldtype), fieldvalue in izip(fieldspec, r):
            recs[fieldname] = fieldtype(fieldvalue)
        result.append(recs)
    return result
 
def StdDate(d):
    if d is None: return None
    d1 = datetime.date(d.year, d.month, d.day)
    return d1.strftime('%Y-%m-%d')

def GetEmployees(p):
    sql = 'SELECT * FROM tblEmployeeDetails'
    fieldspec = [('Person', AsAscii), ('PersonNAME', AsAscii), ('IsStaff', bool), ('MobilSmn', AsAscii)]
    recs = RecordsList(sql, fieldspec)
    employees = {}    
    for r in recs: employees[r['Person']] = r
    return employees

def GetInvoices(field_list):
    sql = "SELECT * FROM tblInvoice WHERE InvBillingPeriod='" +  period.mmmmyyyy() + "'"
    return records(field_list, sql)

    
def GetJobs():
    sql = 'SELECT * FROM jobs ORDER BY job'
    fieldspec = [('ID', int), ('job', str), ('title', str), ('address', AsAscii), 
        ('references', AsAscii), ('briefclient', AsInt),
        ('active', bool), ('vatable', bool), ('exp_factor', AsFloat), ('WIP', bool), ('Weird', bool), 
        ('Autoprint', bool), ('Comments', AsAscii), ('TsApprover', AsAscii), 
        ('UtilisedPOs', AsFloat), ('PoBudget', AsFloat), ('PoStartDate', StdDate), 
        ('PoEndDate', StdDate), ('ProjectManager', AsAscii)]
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
 
def GetTblBilling():
    global tbl_billing
    if tbl_billing is not None: return
    sql = 'SELECT * FROM tblBilling'
    rows = fetch_all(sql)
    tbl_billing = rows
        
    
    
def GetTimeitems():
    GetTblBilling()
    
    #fmt =  'SELECT * FROM tblTimeItems '
    #fmt += 'WHERE TimeVal<>0 AND LEN(JobCode) > 0 '
    #fmt += 'AND Month([DateVal]) = %d and Year([DateVal]) = %d '
    #fmt += 'ORDER BY JobCode, Task, Person, DateVal'
    #p = period.g_period
    #sql = fmt % (p.m , p.y)

    sql = 'SELECT * FROM tblTimeItems WHERE TimeVal<>0 AND LEN(JobCode) > 0 ORDER BY JobCode, Task, Person, DateVal'
    rows = fetch_all(sql)
    fields = 'DateVal,JobCode,Person,Task,TimeVal,WorkDone'
    fields = fields.split(',')
    recs = []
    for row in rows:
        rec = {}
        for field, value in zip(fields, row):
            rec[field] = value
        recs.append(rec)
    
    # filter by period
    global tbl_billing
    def find(item, sequence, key):
        for el in sequence:
            if item == key(el):
                return el
        return None
    per = find(period.billing_key(), tbl_billing, key = lambda x: x[0])
    start = per[1]
    end = per[2]
    print per
    def within(x): return x['DateVal'] >= start and x['DateVal'] <= end
    recs = filter(within, recs)
    #fieldspec = [('JobCode', str), ('Person', str), ('DateVal', StdDate), ('TimeVal', AsFloat), ('Task', str), ('WorkDone', unicode)]
    #recs = RecordsList(sql, fieldspec)
    return recs



def GetCharges(p):
    sql = 'SELECT JobCode, TaskNo, Person, PersonCharge FROM tblCharges;'
    fieldspec = [('JobCode', str), ('TaskNo', str), ('Person', str), ('PersonCharge', AsFloat)]
    recs = RecordsList(sql, fieldspec)
    charges = {}
    for r in recs: charges[ (r['JobCode'], r['TaskNo'], r['Person'] )] = r['PersonCharge']
    return charges

def GetClients():
    sql = 'SELECT * FROM tblClients;'
    fieldspec = [('ID', int), ('brief', str)]
    recs = RecordsList(sql, fieldspec)
    clients = {}
    for r in recs: clients[r['ID']] = r['brief']
    return clients

@print_timing
def fetch():
    d = {}
    p = period.g_period
    
    # table data
    #d['period'] = p # TODO HIGH: eliminate this, as we should now be seitching over to g_period
    d['employees'] = GetEmployees(p)
    d['jobs'] = GetJobs()
    d['tasks'] = GetTasks(p)
    d['timeItems'] = GetTimeitems()
    d['charges'] = GetCharges(p)
    d['clients'] = GetClients()
    d['auto_invoices'] = None
    d['manual_invoices'] = None
    #d['invoice_tweaks'] = None    
    #d['expenses'] = expenses.read_expenses(p)
    return d

###########################################################################

def initials_to_name(data, initials):
    'Convert a persons initials to their full name'
    try: name = data['employees'][initials]['PersonNAME']
    except: name = initials + " name???" # TODO add a warning log
    return name

def task_desc(data, jobcode, taskcode):
    'Obtain a task description from a job and task code'
    try: desc = data['tasks'][(jobcode, taskcode)]['TaskDes']
    except: desc = '{0}-{1} desc??'.format(jobcode, taskcode) # TODO add a warning log
    return desc

###########################################################################

__pickle_dir = os.path.expanduser("~/.config/pypms")
__pickle_file = __pickle_dir + "/pypms.pck"

def load_state():
    fp = open(__pickle_file, "rb")
    d = pickle.load(fp)
    fp.close()
    return d

def save_state(d):
    common.makedirs(__pickle_dir)
    fp = open(__pickle_file, "wb")
    pickle.dump(d, fp)
    fp.close()
    
    
    
###########################################################################

def test():
    #conn = Database()
    #conn.execute('select * from yuk')
    pass
        
###########################################################################
if  __name__ == "__main__": 
    test()
    princ('Finished')