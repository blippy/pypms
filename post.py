# post invoices

# works on lappy, but not iMac

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
        print job, net, vat, total
        

# rough idea how to adjust PMS
#Dim total, expenses, time
#total = ws.Cells(r, 6)
#expenses = ws.Cells(r, 10)
#time = total - expenses
#sql = "UPDATE tblInvoice SET "
#If ws.Cells(r, 4) Then ' WIP
#    sql = sql & "InvUBI = " & time & " , InvWIP =" & expenses
#Else
#    sql = sql & "InvInvoice = " & total & " , Inv3rdParty =" & expenses
#End If
#sql = sql & " WHERE InvJobCode = '" & jobcode & "' AND InvBillingPeriod='" & [nInvBillingPeriod] & "'"
#db.IssueSqlCommand sql

#UPDATE Table1 SET (...) WHERE Column1='SomeValue'
#IF @@ROWCOUNT=0
#    INSERT INTO Table1 VALUES (...)
    
def AdjustPms(d):

    # work out which invoices have already been created
    invoicesAlreadyCreated = set()
    sql = "SELECT * FROM tblInvoice WHERE InvBillingPeriod='" +  d.p.mmmmyyyy() + "'"
    codes = db.records(['InvJobCode'], sql)
    invoicesAlreadyCreated = [ AsAscii(rec[0]) for rec in codes]
    
    try:
        conn = db.DbOpen()
        if foo in invoicesAlreadyCreated:
            sql = "UPDATE something"
            sql = "UPDATE tblInvoice SET InvComments = 'Python' WHERE InvBillingPeriod = 'April 2010' AND InvJobCode = '2736'"
        else:
            sql = "INSERT something"
        conn.Execute(sql)
        # FIXME NOW
    finally:
        conn.Close()
        
    
    
    
if  __name__ == "__main__":
    # perform a simple test
    d = data.Data()
    d.restore()
    ImportManualInvoices(d)
    # FixPms()
    AdjustPms(d)
    print 'Finished'