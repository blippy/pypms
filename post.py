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

import win32com.client

import common
from common import print_timing
import db, excel, unpost
import period
from common import AsAscii, AsFloat, princ
   

###########################################################################


    
def accumulate(d):
    # FIXME - this can probably be used in many other places
    result = {}
    for el in d['manual_invoices']:
        common.dplus(result, el['job'], el['net'])
    return result

############################################################################

def InsertFreshMonth(conn):
    'Put a selection of zeros in tblInvoice'
    
    # Determine which jobs need to be recorded in tblInvoice
    # Do we need to do nearly all of them, or just the active ones?
    invBillingPeriod = period.mmmmyyyy()
    recs = db.GetTasks()
    activeJobs = set([key[0] for key in recs.keys()])
    ignoreJobs = set(['010500', '010400', '010300', '010200', '3. Sundry', '404550'])
    jobsToCreate = activeJobs - ignoreJobs
    
    invDate = datetime.date.today().strftime('%d/%m/%Y')
    for jobcode in jobsToCreate:
        fmt =  "INSERT INTO tblInvoice "
        fmt += "(InvDate, InvBillingPeriod, InvBIA, InvUBI, InvWIP, InvAccrual, InvInvoice, Inv3rdParty, InvTime, InvJobCode, InvComments, InvPODatabaseCosts, InvCapital,InvStock)"
        fmt += " VALUES "
        fmt += "('%s', '%s', 0, 0, 0, 0, 0, 0, 0, '%s', 'PyPms', 0, 0, 0) " #WHERE InvJobCode='%s' AND InvBillingPeriod='%s'"
        sql = fmt % (invDate, invBillingPeriod, jobcode)
        conn.execute(sql)


###########################################################################

def accumulate_tweaks(d, xl):
    '''Accumulate invoice tweaks to job level'''

    #xl = excel.ImportCamelWorksheet('InvTweaks')
    field_names = ['Job', 'InvBIA', 'InvUBI', 'InvWIP', 'InvAccrual', 'InvInvoice', 'Inv3rdParty', 'InvTime', 'Recovery', 'Comment']
    f = common.AsFloat
    field_types = [str, f, f, f, f, f, f, f, f, str]
    tweaks = []

    for line in xl[1:]:
        tweak = {}
        for field, value, converter in itertools.izip(field_names, line, field_types):
            tweak[field] = converter(value)
        tweaks.append(tweak)

    
    jobs = {}
    for entry in tweaks:
        code = entry['Job']
        if not jobs.has_key(code): jobs[code] = {}
        job = jobs[code]
        for key in ['Inv3rdParty', 'InvAccrual', 'InvBIA', 'InvInvoice', 'InvTime', 'InvUBI', 'InvWIP']:
            common.dplus(job, key, common.dget(entry, key))
    return jobs

###########################################################################

def update_pms(conn, d, xl_invtweaks):
    '''Update the invoice table in PMS. Note that AugmentPms() has already inserted any necessary 
    records, so we only have to update, and not insert'''
    
    # Jobs requiring entries
    invBillingPeriod = period.mmmmyyyy()
    code_records = db.GetInvoices()
    #pdb.set_trace()
    codes = [str(rec['InvJobCode']) for rec in code_records]
    invoices = d['auto_invoices']
    
    manual_invoices = accumulate(d)
    tweaks = accumulate_tweaks(d, xl_invtweaks)
    
    for code in codes:
        
        bia = 0.0
        ubi = 0.0
        wip = 0.0
        accrual = 0.0
        invoice_total = 0.0
        party3 = 0.0
        invoice_time = 0.0
        comments = ''
        
        # adjust for automated invoices
        if invoices.has_key(code):
            invoice = invoices[code]
            invoice_time += invoice['work']
            party3 += invoice['expenses']
            if d['jobs'][code]['WIP']:
                ubi += invoice['work']
                wip += invoice['expenses']
            else:        
                invoice_total += invoice['net']
            
            
        # adjust for manual invoices
        invoice_total += common.dget(manual_invoices, code, 0.0)
        
        # adjust for invoice tweaks
        tweak = common.dget(tweaks, code, None)
        if tweak:
            party3 += tweak['Inv3rdParty']
            accrual += tweak['InvAccrual']
            bia += tweak['InvBIA']
            invoice_total += tweak['InvInvoice']
            invoice_time += tweak['InvTime']
            ubi += tweak['InvUBI']
            wip += tweak['InvWIP']
            

        
        # now post those entries
        comment = 'PyPms ' + str(period.now())
        fmt = "UPDATE tblInvoice SET InvBIA=%.2f, InvUBI=%.2f, InvWIP=%.2f, InvAccrual=0,InvInvoice=%.2f, "
        fmt += "Inv3rdParty=%.2f, InvTime=%.2f , InvComments='%s', InvPODatabaseCosts=0, "
        fmt += "InvCapital=0,InvStock=0 WHERE InvJobCode='%s' AND InvBillingPeriod='%s'"
        sql = fmt % (bia, ubi, wip, invoice_total, party3, invoice_time, comment, code, invBillingPeriod)
        conn.Execute(sql)

###########################################################################

def add_purchase_orders(conn):

    # Retrieve the relevant info from the database
    p = period.g_period
    from_date = p.first()
    to_date = p.last()
    # lifted from PMS query qselPOCosts
    fmt = """SELECT tblPurchaseItems.POJobCode AS [code], Sum([POValue]*[POQty]) AS [cost]
        FROM tblPurchaseOrders INNER JOIN tblPurchaseItems ON 
        tblPurchaseOrders.PONumber = tblPurchaseItems.PONumber
        WHERE (((tblPurchaseOrders.PODate) Between #%s# And #%s#))
        GROUP BY tblPurchaseItems.POJobCode"""
    fmt = "SELECT Code, Cost FROM qryPO WHERE BillingPeriod='%s'"
    sql = fmt % (p.yyyymm())
    costs = {}
    for rec in db.fetch_all(sql):
        costs[str(rec[0])] = rec[1]
    
    # Add PO costs to database
    for job_code in costs:
        cost = costs[job_code]
        fmt = "UPDATE tblInvoice SET InvPODatabaseCosts=%.2f WHERE InvJobCode='%s' AND InvBillingPeriod='%s'"
        sql = fmt % (cost, job_code, p.mmmmyyyy())
        conn.Execute(sql)

    

###########################################################################

@print_timing
def post_main(d, xl_invtweaks):
    conn = db.DbOpen()
    #import_manual_invoices(d)
    unpost.zap_entries()
    InsertFreshMonth(conn)
    update_pms(conn, d, xl_invtweaks)
    add_purchase_orders(conn)
    conn.Close()
    
if  __name__ == "__main__":
    data = db.load_state()
    post_main(data)
    princ("Finished")