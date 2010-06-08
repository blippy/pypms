# post invoices

# works on lappy, but not iMac

import datetime

import win32com.client

import data, db, excel
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
        
def AugmentPms(d):
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

    
def UpdatePms(d):
    '''Update the invoice table in PMS. Note that AugmentPms() has already inserted any necessary 
    records, so we only have to update, and not insert'''
    
    # Jobs requiring entries
    invBillingPeriod = d.p.mmmmyyyy()
    sql = "SELECT * FROM tblInvoice WHERE InvBillingPeriod='" +  invBillingPeriod + "'"
    codes = [rec[0] for rec in db.records(['InvJobCode'], sql)]
    
    try:
        conn = db.DbOpen()
        for code in codes:
            
            # determine what entries should be made to the table of invoices
            invoice = d.invoices[code]
            ubi = 0
            wip = 0
            invoicedOut = 0
            party3 = 0
            work = 0
            if d.jobs[code]['WIP']:
                ubi = invoice['work']
                wip = invoice['expenses']
            else:
                party3 = invoice['expenses']
                work = invoice['work']
                invoicedOut = invoice['net']
                
                
            # now post those entries
            fmt = "UPDATE tblInvoice SET InvBIA=0, InvUBI=%d, InvWIP=%d, InvAccrual=0,InvInvoice=%d, Inv3rdParty=%d, InvTime=%d , InvComments='PyPms autofilled', InvPODatabaseCosts=0, InvCapital=0,InvStock=0 WHERE InvJobCode='%s' AND InvBillingPeriod='%s'"
            sql = fmt % (ubi, wip, invoicedOut, party3, work, code, invBillingPeriod)
            conn.Execute(sql)
    finally:
        conn.Close()
        
    
    
    
if  __name__ == "__main__":
    d = data.Data()
    d.restore()
    ImportManualInvoices(d)
    AugmentPms(d)
    UpdatePms(d)
    print 'Finished'