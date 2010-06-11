'Zap a month of invoices (tblInvoice)'

import db
import period

def zap_entries(per):
    'Remove all recorded invoices for the period PER'
    fmt = "DELETE FROM tblInvoice WHERE InvBillingPeriod='%s'"
    sql = fmt % (per.mmmmyyyy())
    db.ExecuteSql(sql)
    
def main():
    'Main point of entry - only for use on the command line, though.'
    print "DANGER WILL ROBINSON! DANGER! DANGER!"
    print "You are about to delete invoices from the database"
    per = period.Period()
    per.inputPeriod()
    zap_entries(per)
    print 'Finished'
    
    
if  __name__ == "__main__":
    main()