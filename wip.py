'''Create a WIP summary
'''

import pdb

import common
from common import aggregate, AsAscii, summate
import db
import excel


###########################################################################

def sums(invoices):
    #pdb.set_trace()
    ufunc = lambda x: float(x['InvUBI'])
    wfunc = lambda x: float(x['InvWIP'])
    ubi_ytd = summate(invoices, ufunc)
    wip_ytd = summate(invoices, wfunc)    
    return ubi_ytd, wip_ytd, ubi_ytd + wip_ytd

###########################################################################
def main(d):
    
    fieldspec = [('InvJobCode', AsAscii), ('InvBillingPeriod', AsAscii), 
        ('InvUBI', float), ('InvWIP', float)]
    #wips_for_this_month = db.GetInvoices(d, fieldspec)
    wips = db.RecordsList("SELECT * FROM tblInvoice", fieldspec)
    billing_period = d['period'].mmmmyyyy()
    monthlies = filter(lambda x: x['InvBillingPeriod'] == billing_period, wips)
    
    
    output = []
    output.append(['Job', 'UBI', 'WIP', 'TOTAL', 'UBI', 'WIP', 'TOTAL'])
    output.append(['', 'YTD', 'YTD', 'YTD', 'CUR', 'CUR', 'CUR'])
    
    # output the totals for each job code
    for job_code, invoices in aggregate(wips, common.mkKeyFunc("InvJobCode")):
        ubi_ytd, wip_ytd, sum_ytd = sums(invoices)
        cur = filter(lambda x: x['InvJobCode'] == job_code, monthlies)
        ubi_cur, wip_cur, sum_cur = sums(cur)
        line = [job_code, ubi_ytd, wip_ytd, sum_ytd, ubi_cur, wip_cur, sum_cur]
        
        # maybe add the line to output
        total = reduce( lambda x,y: abs(x)+abs(y), line[1:])
        if total > 0.01: output.append(line)
            

    # output the totals
    output.append([])
    ubi_ytd, wip_ytd, sum_ytd = sums(wips)
    ubi_cur, wip_cur, sum_cur = sums(monthlies)
    line = ['TOTAL', ubi_ytd, wip_ytd, sum_ytd, ubi_cur, wip_cur, sum_cur]
    output.append(line)
    
    excel.create_report(d['period'], "Wip", output, [2, 3 ,4, 5, 6, 7])

        
###########################################################################

if  __name__ == "__main__":
    print "Didn't do anything"