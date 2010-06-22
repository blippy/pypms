# post invoices

# works on lappy, but not iMac
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

import win32com.client

import common
import data, db, excel, unpost
import invsummary
from common import AsAscii, AsFloat
        

############################################################################

def InsertFreshMonth(conn, d):
    'Put a selection of zeros in tblInvoice'
    
    # Determine which jobs need to be recorded in tblInvoice
    invBillingPeriod = d.p.mmmmyyyy()
    sql = "SELECT JobCode FROM tblTasks WHERE JobActive=Yes"
    sql = "SELECT JobCode FROM tblTasks"
    activeJobs = set([str(rec[0]) for rec in db.records(['JobCode'], sql)])
    ignoreJobs = set(['010500', '010400', '010300', '010200', '3. Sundry', '404550'])
    jobsToCreate = activeJobs - ignoreJobs
    
    invDate = datetime.date.today().strftime('%d/%m/%Y')
    for jobcode in jobsToCreate:
        fmt =  "INSERT INTO tblInvoice "
        fmt += "(InvDate, InvBillingPeriod, InvBIA, InvUBI, InvWIP, InvAccrual, InvInvoice, Inv3rdParty, InvTime, InvJobCode, InvComments, InvPODatabaseCosts, InvCapital,InvStock)"
        fmt += " VALUES "
        fmt += "('%s', '%s', 0, 0, 0, 0, 0, 0, 0, '%s', 'PyPms', 0, 0, 0) " #WHERE InvJobCode='%s' AND InvBillingPeriod='%s'"
        sql = fmt % (invDate, invBillingPeriod, jobcode)
        #print sql
        conn.execute(sql)

    #print jobsToCreate

###########################################################################

def UpdatePms(conn, d):
    '''Update the invoice table in PMS. Note that AugmentPms() has already inserted any necessary 
    records, so we only have to update, and not insert'''
    
    # Jobs requiring entries
    invBillingPeriod = d.p.mmmmyyyy()
    code_records = db.GetInvoices(d, ['InvJobCode'])
    codes = [str(rec[0]) for rec in code_records]
    invoices = d.auto_invoices
    
    for code in codes:
        
        if not invoices.has_key(code): continue # might be an active job for which we have no time or expenses
        # determine what entries should be made to the table of invoices
        invoice = invoices[code]
        ubi = 0
        wip = 0
        invoicedOut = 0
        party3 = 0
        work = invoice['work']
        party3 = invoice['expenses']
        if d.jobs[code]['WIP']:
            ubi = invoice['work']
            wip = invoice['expenses']
        else:        
            invoicedOut = invoice['net']
            
            
        # sum the invoices entered manually
        manual_invoice_total = 0.0
        for inv in d.manual_invoices:
            if inv['job'] == code:
                manual_invoice_total += float(inv['net'])

        
        # now post those entries
        invoice_total = invoicedOut + manual_invoice_total
        fmt = "UPDATE tblInvoice SET InvBIA=0, InvUBI=%.2f, InvWIP=%.2f, InvAccrual=0,InvInvoice=%.2f, Inv3rdParty=%.2f, InvTime=%.2f , InvComments='PyPms autofilled', InvPODatabaseCosts=0, InvCapital=0,InvStock=0 WHERE InvJobCode='%s' AND InvBillingPeriod='%s'"
        sql = fmt % (ubi, wip, invoice_total, party3, work, code, invBillingPeriod)
        conn.Execute(sql)

###########################################################################

def add_purchase_orders(conn, d):

    # Retrieve the relevant info from the database
    from_date = d.p.first()
    to_date = d.p.last()
    # lifted from PMS query qselPOCosts
    fmt = """SELECT tblPurchaseItems.POJobCode AS [code], Sum([POValue]*[POQty]) AS [cost]
        FROM tblPurchaseOrders INNER JOIN tblPurchaseItems ON 
        tblPurchaseOrders.PONumber = tblPurchaseItems.PONumber
        WHERE (((tblPurchaseOrders.PODate) Between #%s# And #%s#))
        GROUP BY tblPurchaseItems.POJobCode"""
    fmt = "SELECT * FROM qryPO WHERE BillingPeriod='%s'"
    sql = fmt % (d.p.yyyymm())
    #sql = fmt % (from_date, to_date)
    #print sql
    costs = {}
    for rec in db.records(['Code', 'Cost'], sql):
        costs[str(rec[0])] = rec[1]
    
    # Add PO costs to database
    #print costs
    for job_code in costs:
        cost = costs[job_code]
        fmt = "UPDATE tblInvoice SET InvPODatabaseCosts=%.2f WHERE InvJobCode='%s' AND InvBillingPeriod='%s'"
        sql = fmt % (cost, job_code, d.p.mmmmyyyy())
        conn.Execute(sql)

    
###########################################################################

def usingconn(conn):
    d = data.Data()
    d.restore()
    invsummary.import_manual_invoices(d)
    unpost.zap_entries(d.p)
    InsertFreshMonth(conn, d)
    UpdatePms(conn, d)
    add_purchase_orders(conn, d)
    #create_invoice_summary(d)

###########################################################################

def main():
    conn = db.DbOpen()
    try: usingconn(conn)
    finally: conn.Close()
    
if  __name__ == "__main__":
    main()
    print 'Finished'