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
    * Combo33 - Billing Perdio - e.g. May 2010

"""
import datetime

import win32com.client

import data, db, excel, unpost
#from db import database
from common import AsAscii, AsFloat

def ImportManualInvoices(d):
    'Import invoices entered manually in spreadsheet'
    fileName = 'M:\\Finance\\Invoices\\Inv summaries %s\\Inv Summary %s.xls' % (d.p.y, d.p.yyyymm())
    wsName = 'Invoices'
    invoiceLines = excel.ImportWorksheet(fileName, wsName)
    for row in invoiceLines:
        job = row[2] #, net, vat, total = row[2:5]
        if not d.jobs.has_key(job): continue
        id = row[0]
        net, vat , total = map(AsFloat, row[3:6])
        try: desc = row[6]
        except IndexError: desc = ''
        #print job, net, vat, total
        # FIXME
        
def XXXAugmentPms(d):
    "Add tblInvoice records for jobs that it doesn't already have"
    
    # Currently in the database
    invBillingPeriod = d.p.mmmmyyyy()
    sql = "SELECT * FROM tblInvoice WHERE InvBillingPeriod='" +  invBillingPeriod + "'"
    alreadyCreated = set()
    for rec in db.records(['InvJobCode'], sql): alreadyCreated.add(str(rec[0]))
    
    # Jobs we need to have
    jobsRequired = set()
    for el in d.expenses + d.timeItems: jobsRequired.add(el['JobCode'])
 
    # Now create the jobs that we need
    jobsToCreate = (jobsRequired - alreadyCreated)  - set(['010500', '010400', '010300', '010200'])
    for job in jobsToCreate:
        invDate = datetime.date.today().strftime('%d/%m/%Y')
        sql = "INSERT INTO tblInvoice (InvDate, InvBillingPeriod, InvJobCode) VALUES ('%s', '%s', '%s')" % (invDate, invBillingPeriod, job)
        try:
            conn = db.DbOpen()
            conn.execute(sql)
        finally:
            conn.Close()

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
    
def UpdatePms(conn, d):
    '''Update the invoice table in PMS. Note that AugmentPms() has already inserted any necessary 
    records, so we only have to update, and not insert'''
    
    # Jobs requiring entries
    invBillingPeriod = d.p.mmmmyyyy()
    sql = "SELECT * FROM tblInvoice WHERE InvBillingPeriod='" +  invBillingPeriod + "'"
    codes = [str(rec[0]) for rec in db.records(['InvJobCode'], sql)]
    invoices = d.invoices
    
    for code in codes:
        
        if not invoices.has_key(code): continue # might be an active job for which we have no time or expenses
        # determine what entries should be made to the table of invoices
        invoice = invoices[code]
        ubi = 0
        wip = 0
        invoicedOut = 0
        party3 = 0
        work = invoice['work']
        if d.jobs[code]['WIP']:
            ubi = invoice['work']
            wip = invoice['expenses']
        else:
            party3 = invoice['expenses']
            invoicedOut = invoice['net']
            
            
        # now post those entries
        fmt = "UPDATE tblInvoice SET InvBIA=0, InvUBI=%.2f, InvWIP=%.2f, InvAccrual=0,InvInvoice=%.2f, Inv3rdParty=%.2f, InvTime=%.2f , InvComments='PyPms autofilled', InvPODatabaseCosts=0, InvCapital=0,InvStock=0 WHERE InvJobCode='%s' AND InvBillingPeriod='%s'"
        sql = fmt % (ubi, wip, invoicedOut, party3, work, code, invBillingPeriod)
        conn.Execute(sql)

        

def usingconn(conn):
    # FIXME some of this stuff will need reinstating - I don't know what, yet
    d = data.Data()
    d.restore()
    # ImportManualInvoices(d) FIXME reinstate
    #AugmentPms(d)
    unpost.ZapEntries(d.p)
    # Run the PMS query to create the monthly invoices
    #print 'So far so good'
    
    #query = 'qappSummaryReportActiveJobsOnly+HoursInvoice'
    #query = 'mcrInvoice'
    #print "query is: ", query
    #conn.DoCmd.RunSQL('qappSummaryReportActiveJobsOnly+HoursInvoiced', '01/05/2010', '31/05/2010', 'May 2010')
    #conn.Run('mcrInvoice', '01/05/2010', '31/05/2010', 'May 2010')
    #conn.Run(query, ['May 2010'])
    InsertFreshMonth(conn, d)
    UpdatePms(conn, d)
 
def main():
    conn = db.DbOpen()
    try: usingconn(conn)
    finally: conn.Close()
    
if  __name__ == "__main__":
    main()
    print 'Finished'