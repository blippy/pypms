'''Spot recoveries in PMS
'''

import common
import db
import excel

# FIXME make use of the recoveries in invoice_tweaks

###########################################################################
def main(d):
    
    # summate manual invoices - they will give clues as to likely recoveries
    manuals = {}
    # FIXME - this kind of incrementing should be refactored and used extensively
    for manual_invoice in d.manual_invoices:
        job_code = manual_invoice['job']
        if not manuals.has_key(job_code): manuals[job_code] = 0.0
        manuals[job_code]  += manual_invoice['net']


    # create a report of the recoveries
    output = [['Job', 'Recovery', 'Manual', 'WIP?', 'Weird?', 'Comment']]
    field_list = ['InvJobCode', 'InvComments', 'InvBIA', 'InvUBI', 'InvWIP', 'InvAccrual', 'InvInvoice', 'Inv3rdParty', 'InvTime']
    invoices = db.GetInvoices(d, field_list)    
    total_recovery = 0.0
    total_manual = 0.0
    for invoice in invoices:
        job_code = invoice[0]
        comment = invoice[1]
        as_floats = map(float, invoice[2:])
        InvBIA, InvUBI, InvWIP, InvAccrual, InvInvoice, Inv3rdParty, InvTime = as_floats
        recovery = InvBIA + InvUBI + InvWIP + InvAccrual + InvInvoice - Inv3rdParty - InvTime
        if abs(recovery)< 20.0: continue # ignore immaterial recoveries
        total_recovery += recovery
        recovery
    
        # FIXME - candidate for refactoring and using extensively
        if manuals.has_key(job_code): manual_amount = manuals[job_code]
        else: manual_amount = 0.0
        total_manual += manual_amount
        
        job = d.jobs[job_code]
        if job['WIP']: wip = 'Y'
        else: wip = ''
        if job['Weird']: weird = 'Y'
        else: weird = ''
        output.append([job_code, recovery, manual_amount, wip, weird, comment])
        
    output.append([])
    output.append(['TOTAL', total_recovery, total_manual])
    excel.create_report(d.p, "Recoveries", output, [2])
        
    #=[InvBIA]+[InvUBI]+[InvWIP]+[InvAccrual]+[InvInvoice]-[Inv3rdParty]-[InvTime]
        
if  __name__ == "__main__":
    common.run_current(main)
    print 'Finished'