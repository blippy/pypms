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

import csv
import datetime

import win32com.client

import common
import data, db, excel, unpost
#from db import database
from common import AsAscii, AsFloat
import maninv
        

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
        if d.manual_invoices.has_key(code):
            for inv in d.manual_invoices[code]:
                manual_invoice_total += float(inv['net'])

        
        # now post those entries
        invoice_total = invoicedOut + manual_invoice_total
        fmt = "UPDATE tblInvoice SET InvBIA=0, InvUBI=%.2f, InvWIP=%.2f, InvAccrual=0,InvInvoice=%.2f, Inv3rdParty=%.2f, InvTime=%.2f , InvComments='PyPms autofilled', InvPODatabaseCosts=0, InvCapital=0,InvStock=0 WHERE InvJobCode='%s' AND InvBillingPeriod='%s'"
        sql = fmt % (ubi, wip, invoice_total, party3, work, code, invBillingPeriod)
        conn.Execute(sql)

        
def create_invoice_summary(d):
    path = d.p.outDir() + '\\craig'
    common.makedirs(path)
    file_name = path + "\\invoices.csv"
    out = csv.writer(open(file_name, 'wb'))
    out.writerow(['Ref', 'Client', 'Job', 'Net', 'Desc'])
    total = 0.0
    
    #def txt(input): return "'" + input
    def txt(input): return input
    def number(input): return '%.2f' % (input)

    # spit out the manual invoices
    for job_code in d.manual_invoices.keys():
        for inv in d.manual_invoices[job_code]:
            #FIXME - ought to be possible to work out who the client is
            net = inv['net']
            total += net
            out.writerow([txt(inv['id']), inv['client'], txt(job_code), number(net), inv['desc']])
        
    # write out the computed invoices
    for job_code in d.auto_invoices:
        inv = d.auto_invoices[job_code]
        #FIXME - ought to be possible to work out who the client is
        net = inv['net']
        if net <> 0.0:
            total += net
            out.writerow(["", "", txt(job_code), number(inv['net'])])
        
    out.writerow([])
    out.writerow(['Total', '', '', number(total)])
        
def usingconn(conn):
    d = data.Data()
    d.restore()
    maninv.import_manual_invoices(d)
    unpost.zap_entries(d.p)
    InsertFreshMonth(conn, d)
    UpdatePms(conn, d)
    create_invoice_summary(d)
 
def main():
    conn = db.DbOpen()
    try: usingconn(conn)
    finally: conn.Close()
    
if  __name__ == "__main__":
    main()
    print 'Finished'