# zap a month of invoices (tblInvoice)

import db, period

def main():
    print "DANGER WILL ROBINSON! DANGER! DANGER!"
    print "You are about to delete invoices from the database"
    p = period.Period()
    p.inputPeriod()
    sql = "DELETE FROM tblInvoice WHERE InvBillingPeriod='%s'" % (p.mmmmyyyy())
    db.ExecuteSql(sql)
    print 'Finished'
    
if  __name__ == "__main__":
    main()