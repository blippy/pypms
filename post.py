# post invoices

"""
When the user clicks on Post Invoices in PMS, it activates macro 
    mcrInvoice,
which in turn runs query 
    qappSuumaryReportActiveJobsOnly+HoursInvoiced
This query takes 3 paramemters, in the following order:
    * text8.Value - StartDate - e.g. 01/05/2010
    * text10,Value - EndDate - e.g. 31/05/2010
    * Combo33 - Billing Period - e.g. May 2010

"""

###########################################################################

import csv
import datetime
import itertools 
import pdb
import pprint
import traceback

import ordereddict
import win32com.client

import common
from common import AsAscii, AsFloat, princ
import db
import excel
import period

   


############################################################################

def zap_entries():
    'Remove all recorded invoices for the period PER'
    fmt = "DELETE FROM tblInvoice WHERE InvBillingPeriod='{0}'"
    #print "zap_entries"
    sql = fmt.format(period.mmmmyyyy())
    #print sql
    db.ExecuteSql(sql)
    #db.commit()
    

############################################################################
def get_keys():
    return ['InvDate', 'InvBillingPeriod', 'InvBIA', 'InvUBI', 'InvWIP', 
        'InvAccrual', 'InvInvoice', 'Inv3rdParty', 'InvTime', 
        'InvJobCode', 'InvComments', 'InvPODatabaseCosts', 'InvCapital',
        'InvStock']
  
def create_insertion(jcode):
        d = ordereddict.OrderedDict(map(lambda x: [x, 0.0], get_keys()))
        d['InvDate'] = datetime.date.today().strftime('%d/%m/%Y')
        d['InvBillingPeriod'] = period.mmmmyyyy()
        d['InvJobCode'] = str(jcode)
        d['InvComments'] = "PMS " + common.get_timestamp()
        return d
    
def setup_insertions(tasks):
    invBillingPeriod = period.mmmmyyyy()
    #recs = db.GetTasks()
    activeJobs = set([key[0] for key in tasks.keys()])
    ignoreJobs = set(['010500', '010400', '010300', '010200', '3. Sundry', '404550'])
    jobsToCreate = activeJobs - ignoreJobs
    insertions = {}
    for jcode in jobsToCreate:
        insertions[jcode] = create_insertion(jcode)
    return insertions
        

def process_autos(inserts, autos, d):
    for invoice in autos.values():
        jcode = invoice['JobCode']
        insert = inserts[jcode]
        invoice_time = invoice['work']
        insert['InvTime'] += invoice_time
        party3 = invoice['expenses']
        insert['Inv3rdParty'] += party3
        if d['jobs'][jcode]['WIP']:
            insert['InvUBI'] += invoice_time
            insert['InvWIP'] += party3
        else:        
            insert['InvInvoice'] += invoice['net']
        inserts[jcode] = insert
        
def process_manuals(inserts, manuals):
    "Process manual invoices"
    for m in manuals:
        jcode = m['JobCode']
        inserts[jcode]['InvInvoice'] += m['net']

def process_tweaks(inserts, data):
    tweaks = data['InvTweaks']
    for tweak in tweaks:
        jcode = tweak['JobCode']
        for k in ['Inv3rdParty', 'InvAccrual', 'InvBIA', 'InvInvoice', 'InvTime', 'InvUBI', 'InvWIP']:
            inserts[jcode][k] += tweak[k]
        #print tweak

def process_pos(inserts):
    # Retrieve the relevant info from the database
    #p = period.g_period
    #from_date = p.first()
    #to_date = p.last()
    fmt = "SELECT Code, Cost FROM qryPO WHERE BillingPeriod='%s'"
    sql = fmt % (period.yyyymm())
    costs = {}
    for rec in db.fetch_all(sql):
        costs[str(rec[0])] = rec[1]
    
    # Add PO costs to database
    for jcode in costs:
        cost = costs[jcode]
        if not inserts.has_key(jcode): continue # can happen with '3. Sundry' for example
        #    inserts[jcode] = create_insertion(jcode)
        inserts[jcode]['InvPODatabaseCosts'] = cost
        #fmt = "UPDATE tblInvoice SET InvPODatabaseCosts=%.2f WHERE InvJobCode='%s' AND InvBillingPeriod='%s'"
        #sql = fmt % (cost, job_code, p.mmmmyyyy())
        #conn.Execute(sql)
        
def insert_data(inserts):
    #print inserts
    #conn = db.DbOpen()
    with db.Cursor() as c:
        for ins in inserts.values():
            keys = ' , '.join(get_keys())
            #print keys
            #print ins
            values = []
            for k in get_keys():
                v = ins[k]
                if isinstance(v, str): 
                    v = "'{0}'".format(v)
                else:
                    v = str(v)
                values.append(v)
            values = ' , '.join(values)
        
            sql = "INSERT INTO tblInvoice ({0}) VALUES ({1})".format(keys, values)
            #print sql
            #conn.execute(sql)
            c.execute(sql)
            #print sql
###########################################################################


 
def post_main(data, accumulated_tweaks):
    zap_entries()
    inserts = setup_insertions(data['tasks'])
    #print data.keys()
    autos = data['auto_invoices']
    process_autos(inserts, autos, data)
    manuals = data['ManualInvoices']
    process_manuals(inserts, manuals)
    process_tweaks(inserts, data)
    process_pos(inserts)
    insert_data(inserts)
    #pprint.pprint(inserts)
    
   
if  __name__ == "__main__":
    data = db.load_state()
    #print data
    post_main(data, None)
    princ("Finished")