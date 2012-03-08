'''Spot recoveries in PMS
'''

import operator
import pdb

import common
from common import dget, princ, print_timing
import db
import excel
import period


###########################################################################

def get_dbase_recoveries(d):
    invoices = db.GetInvoices()
    recoveries = {}
    for invoice in invoices:
        def sum_values(fields): return sum(map(lambda x: float(invoice[x]), fields))
        job_code = common.AsAscii(invoice['InvJobCode'])
        comment = invoice['InvComments']
        recovery = sum_values(['InvBIA', 'InvUBI', 'InvWIP', 'InvAccrual', 'InvInvoice'])
        recovery -= sum_values(['Inv3rdParty', 'InvTime'])
        recoveries[job_code] = recovery
    return recoveries

def get_camel_recoveries(d, xl_invtweaks):
    recoveries = {}
    for line in xl_invtweaks[1:]:
        job_code = line[0]
        amount = common.AsFloat(line[8])
        if amount == 0.0: continue
        comment = line[9]
        if not recoveries.has_key(job_code): recoveries[job_code] = []
        recoveries[job_code].append((amount, comment))
    return recoveries


###########################################################################


def line(amount, comment): return  '    %9.2f %s' % (amount, comment)

###########################################################################

def create_recovery_text(job, camel_job_recoveries, db_job_recovery):
    job_code = job['job']
    wip = common.tri(job['WIP'], 'wip' , '') 
    weird = common.tri(job['Weird'], 'weird ', '')
    output = ['JOB %s %s %s' % (job_code, wip, weird)]    

    tweak_total = 0.0    
    for tweak in camel_job_recoveries:
        amount, comment = tweak
        tweak_total += amount
        output.append(line(amount, comment))
    output.append(line(tweak_total , 'TWEAK TOTAL'))

    output.append(line(db_job_recovery , 'PMS RECOVERY'))
    diff = tweak_total - db_job_recovery
    flag = common.tri(abs(diff) > 20.0, ' **** CHECK THIS', '')
    output.append(line(diff , 'DIFF' + flag))
    output += ['']
    return tweak_total, output

###########################################################################
@print_timing
def create_recovery_report(d, xl_invtweaks):
    
    camel_recoveries = get_camel_recoveries(d, xl_invtweaks)
    jobcodes_to_process = set(camel_recoveries.keys())
    db_recoveries = get_dbase_recoveries(d)
    for k, v in db_recoveries.items():
        if abs(v) >= 0.01: jobcodes_to_process.add(k)    
    jobcodes_to_process = sorted(list(jobcodes_to_process))
    output = ['RECOVERY RECONILIATION', '']

    tweaks_grand_total = 0.0
    pms_grand_total = 0.0
    for job_code in jobcodes_to_process:
        job = d['jobs'][job_code] # TODO - d['jobs'][job_code] is a common idiom which ought to be abstracted        
        camel_job_recoveries = dget(camel_recoveries, job_code, [])
        db_job_recovery = dget(db_recoveries, job_code)
        tweak_amount, recovery_text = create_recovery_text(job, camel_job_recoveries, db_job_recovery)
        output += recovery_text
        tweaks_grand_total += tweak_amount
        pms_grand_total += db_job_recovery
       
    output.append('SUMMARY:')
    output.append(line(tweaks_grand_total, 'TWEAKS GRAND TOTAL'))
    output.append(line(pms_grand_total, 'PMS GRAND TOTAL'))
    output.append(line(tweaks_grand_total - pms_grand_total, 'OVERALL DIFF'))
    period.save_report('recoveries.txt', output)
    
        
if  __name__ == "__main__":
    d = db.load_state()
    create_recovery_report(d)
    princ("Finished")