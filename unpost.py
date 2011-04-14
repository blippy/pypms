'Zap a month of invoices (tblInvoice)'

import db
import period
from common import princ

def zap_entries():
    'Remove all recorded invoices for the period PER'
    fmt = "DELETE FROM tblInvoice WHERE InvBillingPeriod='%s'"
    sql = fmt % (period.mmmmyyyy())
    db.ExecuteSql(sql)
    
def main():
    'Main point of entry - only for use on the command line, though.'
    princ("DANGER WILL ROBINSON! DANGER! DANGER!")
    princ("You are about to delete invoices from the database")
    per = period.Period()
    per.inputPeriod()
    zap_entries(per)
    princ('Finished')
    
    
if  __name__ == "__main__":
    main()