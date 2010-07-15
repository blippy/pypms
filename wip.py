'''Create a WIP summary
'''

import common
from common import aggregate, AsAscii
import data
import db
import excel


###########################################################################

def process_job(output, job_code, invoices):
    
    # FIXME NOW
    
    #ubi_total = 0.0
    #wip_total = 0.0
    #total = 0.0
    
    def summate(invoices, fieldname):
        # FIXME HIGH - move to common
        total = 0.0
        for inv in invoices:
            v = float(x[fieldname])
            total +=v
        return total
        
    for invoice in invoices:
        ubi_ytd = summate(invoice, 'InvUBI')
        wip_ytd = summate(invoice, 'InvWIP')
        #job_code, ubi, wip, job_total = wip_entry
        #ubi = float(ubi)
        #wip = float(wip)
        #job_total = float(job_total)
        #if abs(ubi) <0.01 and abs(wip) < 0.01: continue
        #ubi_total += ubi
        #wip_total += wip
        #total += job_total
    return ["something"] # or maybe nothing
        

###########################################################################
def main(d):
    
    fieldspec = [('InvJobCode', AsAscii), ('InvBillingPeriod', AsAscii), 
        ('InvUBI', float), ('InvWIP', float)]
    #wips_for_this_month = db.GetInvoices(d, fieldspec)
    wips = db.RecordsList("SELECT * FROM tblInvoice", fieldspec)
    
    output = []
    output.append(['Job', 'UBI', 'WIP', 'TOTAL'])
    for job_code, invoices in aggregate(wips, common.mkKeyFunc("InvJobCode")):
        job = process_job(output, job_code, invoices)
        if job: output.append(job)
        #output.append([job_code, ubi, wip, job_total])

    output.append([])
    output.append(['TOTAL', ubi_total, wip_total, total])
    excel.create_report(d.p, "Wip", output, [2,3,4])

        
if  __name__ == "__main__":
    data.run_current(main)
    print 'Finished'    