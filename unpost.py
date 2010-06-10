# zap a month of invoices (tblInvoice)

import db, period

def ZapEntries(per):
    sql = "DELETE FROM tblInvoice WHERE InvBillingPeriod='%s'" % (per.mmmmyyyy())
    db.ExecuteSql(sql)
    
def main():
    print "DANGER WILL ROBINSON! DANGER! DANGER!"
    print "You are about to delete invoices from the database"
    per = period.Period()
    per.inputPeriod()
    ZapEntries(per)
    print 'Finished'
    
    
if  __name__ == "__main__":
    main()