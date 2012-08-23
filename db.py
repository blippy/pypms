# information stored in the database

# TODO have an ability to key on fields

import datetime
import decimal
from itertools import izip
import os
import pdb
import pickle

import adodbapi
import win32com.client

import common
from common import AsAscii, AsFloat, AsInt, princ, print_timing
import expenses
import period



 
###########################################################################

tbl_billing = None
jobs = None

###########################################################################
        
def DbOpen():
    'Return an open connection to the database'
    conn = win32com.client.Dispatch(r'ADODB.Connection')
    conn.Open('PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=M:/Finance/camel/camel.mdb;')
    return conn



def fetch_all(sql):
    # preferred method from 25-Aug-2011 - it's much faster
    try:
        conStr = r'PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=M:/Finance/camel/camel.mdb;'    
        con = adodbapi.connect(conStr)
        cursor = con.cursor()
        cursor.execute(sql)
        ds = cursor.fetchall()
        rows = [row for row in ds]
        cursor.close()
    finally:
        con.close()
    return rows
        
def fetch_and_dictify(sql, fields):
    # preferred method from 25-Aug-2011
    rows = fetch_all(sql)
    fields = fields.split(',')
    recs = []
    for row in rows:
        rec = {}
        for field, value in zip(fields, row):
            if type(value) is decimal.Decimal: value = float(value)
            rec[field] = value
        recs.append(rec)
    return recs
    
    
def ExecuteSql(sql):
    try:
        conn = DbOpen()
        conn.execute(sql)
    finally:
        conn.Close()
   


def GetEmployees():
    sql = 'SELECT * FROM tblEmployeeDetails'
    #fieldspec = [('Person', AsAscii), ('PersonNAME', AsAscii), ('IsStaff', bool), ('MobilSmn', AsAscii)]
    fields = 'Person,PersonNAME,PersonCompany,Active,Complete,IsStaff,MobilSmn'
    recs = fetch_and_dictify(sql, fields)
    employees = {}    
    for r in recs: employees[r['Person']] = r
    return employees

def GetInvoices():
    sql = "SELECT * FROM tblInvoice WHERE InvBillingPeriod='" +  period.mmmmyyyy() + "'"
    fields = 'InvDate,InvBillingPeriod,InvBIA,InvUBI,InvWIP,InvAccrual,InvInvoice,'
    fields += 'Inv3rdParty,InvTime,InvJobCode,InvComments,InvPODatabaseCosts,'
    fields += 'InvCapital,InvStock'
    recs = fetch_and_dictify(sql, fields)
    return recs

    
def GetJobs():
    global jobs
    if jobs is not None: return jobs
    sql = 'SELECT * FROM jobs ORDER BY job'
    fields = 'ID,job,title,address,references,briefclient,active,vatable,exp_factor,'
    fields += 'WIP,Weird,Autoprint,Comments,TsApprover,UtilisedPOs,'
    fields += 'PoBudget,PoStartDate,PoEndDate,ProjectManager'
    recs = fetch_and_dictify(sql, fields)
    result = {}
    for r in recs: result[r['job']] = r
    jobs = result
    return jobs
 



def GetTasks():
    sql = 'SELECT * FROM tblTasks ORDER BY JobCode, TaskNo'
    fields = 'JobCode,JCDescription,TaskNo,TaskDes,JobActive,TotalBudget,CSLBudget,PManager,BManager,IssuingOfficer,CapStock,AgencyFee'
    recs = fetch_and_dictify(sql, fields)
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
    
    sql = 'SELECT * FROM tblTimeItems WHERE TimeVal<>0 AND LEN(JobCode) > 0 ORDER BY JobCode, Task, Person, DateVal'
    fields = 'DateVal,JobCode,Person,Task,TimeVal,WorkDone'    
    recs = fetch_and_dictify(sql, fields)
    
    # filter by period
    global tbl_billing 
    per = common.find(period.mmmmyyyy(), tbl_billing, key = lambda x: x[0])
    start = per[1]
    end = per[2]
    def within(x): return x['DateVal'] >= start and x['DateVal'] <= end
    recs = filter(within, recs)
    
    return recs



def GetCharges():
    sql = 'SELECT * FROM tblCharges'
    fields = 'JobCode,TaskNo,Person,PersonCharge'
    recs = fetch_and_dictify(sql, fields)
    charges = {}
    for r in recs: charges[ (r['JobCode'], r['TaskNo'], r['Person'] )] = r['PersonCharge']
    return charges

def GetClients():
    sql = 'SELECT * FROM tblClients;'
    fields = 'ID,brief'
    recs = fetch_and_dictify(sql, fields)
    clients = {}
    for r in recs: clients[r['ID']] = r['brief']
    return clients

@print_timing
def fetch():
    d = {}
    p = period.g_period
    
    # table data
    d['employees'] = GetEmployees()
    d['jobs'] = GetJobs()
    d['tasks'] = GetTasks()
    d['timeItems'] = GetTimeitems()
    d['charges'] = GetCharges()
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
    data = fetch()
    save_state(data)
    princ('Finished')