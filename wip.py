'''Create a WIP summary
'''

import pdb

import common
from common import aggregate, AsAscii, summate, princ, print_timing
import db
import excel
import period


###########################################################################

def sums(invoices):
    #pdb.set_trace()
    ufunc = lambda x: float(x['InvUBI'])
    wfunc = lambda x: float(x['InvWIP'])
    ubi_ytd = summate(invoices, ufunc)
    wip_ytd = summate(invoices, wfunc)    
    return ubi_ytd, wip_ytd, ubi_ytd + wip_ytd

###########################################################################

@print_timing
def create_wip_lines():
    fieldspec = [('InvJobCode', AsAscii), ('InvBillingPeriod', AsAscii), 
        ('InvUBI', float), ('InvWIP', float)]
    wips = db.RecordsList("SELECT * FROM tblInvoice", fieldspec)
    billing_period = period.mmmmyyyy()
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
    
    return output
    

###########################################################################
@print_timing
def create_wip_report(output_text = True):
    output = create_wip_lines()
    if output_text:
        period.create_text_report("wip.txt", output)
    else:
        # TODO Consider zapping Excel output option if considered unecessary
        excel.create_report("Wip", output, [2, 3 ,4, 5, 6, 7])

        
###########################################################################

if  __name__ == "__main__":
    princ("Didn't do anything")