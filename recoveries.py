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
    #=[InvBIA]+[InvUBI]+[InvWIP]+[InvAccrual]+[InvInvoice]-[Inv3rdParty]-[InvTime]
    #field_list = ['InvJobCode', 'InvComments', 'InvBIA', 'InvUBI', 'InvWIP', 'InvAccrual', 'InvInvoice', 'Inv3rdParty', 'InvTime']
    invoices = db.GetInvoices()
    #pdb.set_trace()

    recoveries = {}
    for invoice in invoices:
        #def value(x): return float(invoice[x])
        def sum_values(fields): return sum(map(lambda x: float(invoice[x]), fields))
        job_code = common.AsAscii(invoice['InvJobCode'])
        comment = invoice['InvComments']
        #as_floats = map(float, invoice[2:])
        recovery = sum_values(['InvBIA', 'InvUBI', 'InvWIP', 'InvAccrual', 'InvInvoice'])
        #InvBIA, InvUBI, InvWIP, InvAccrual, InvInvoice, Inv3rdParty, InvTime = as_floats
        #recovery = InvBIA + InvUBI + InvWIP + InvAccrual + InvInvoice - Inv3rdParty - InvTime
        recovery -= sum_values(['Inv3rdParty', 'InvTime'])
        #if abs(recovery)< 20.0: continue # ignore immaterial recoveries
        recoveries[job_code] = recovery
        #if job_code == '2799': pdb.set_trace()
    return recoveries

def get_camel_recoveries(d):
    xl = excel.ImportCamelWorksheet('InvTweaks')
    recoveries = {}

    for line in xl[1:]:
        job_code = line[0]
        amount = common.AsFloat(line[8])
        if amount == 0.0: continue
        comment = line[9]
        if not recoveries.has_key(job_code): recoveries[job_code] = []
        recoveries[job_code].append((amount, comment))
    return recoveries


###########################################################################


def line(amount, comment): return  '    %9.2f %s\n' % (amount, comment)

###########################################################################

def create_recovery_text(job, camel_job_recoveries, db_job_recovery):
    #pdb.set_trace()
    job_code = job['job']
    wip = common.tri(job['WIP'], 'wip' , '') 
    weird = common.tri(job['Weird'], 'weird ', '')
    output = 'JOB %s %s %s\n' % (job_code, wip, weird)
    
    #output += '  Per InvTweaks\n'
    tweak_total = 0.0
    
    for tweak in camel_job_recoveries:
        amount, comment = tweak
        tweak_total += amount
        output += line(amount, comment)
    output += line(tweak_total , 'TWEAK TOTAL')

        
    
    output += line(db_job_recovery , 'PMS RECOVERY')
    diff = tweak_total - db_job_recovery
    flag = common.tri(abs(diff) > 20.0, ' **** CHECK THIS', '')
    output += line(diff , 'DIFF' + flag + '\n\n')
    return tweak_total, output

###########################################################################
@print_timing
def create_recovery_report(d):
    
    camel_recoveries = get_camel_recoveries(d)
    jobcodes_to_process = set(camel_recoveries.keys())
    db_recoveries = get_dbase_recoveries(d)
    for k, v in db_recoveries.items():
        if abs(v) >= 0.01: jobcodes_to_process.add(k)    
    jobcodes_to_process = sorted(list(jobcodes_to_process))
    output = 'RECOVERY RECONILIATION\n\n'

    #pdb.set_trace()

    tweaks_grand_total = 0.0
    pms_grand_total = 0.0
    for job_code in jobcodes_to_process:
        job = d['jobs'][job_code] # FIXME - d['jobs'][job_code] is a common idiom which ought to be abstracted        
        camel_job_recoveries = dget(camel_recoveries, job_code, [])
        db_job_recovery = dget(db_recoveries, job_code)
        tweak_amount, recovery_text = create_recovery_text(job, camel_job_recoveries, db_job_recovery)
        output += recovery_text
        tweaks_grand_total += tweak_amount
        pms_grand_total += db_job_recovery
       
    output += 'SUMMARY:\n'
    output += line(tweaks_grand_total, 'TWEAKS GRAND TOTAL')
    output += line(pms_grand_total, 'PMS GRAND TOTAL')
    output += line(tweaks_grand_total - pms_grand_total, 'OVERALL DIFF')        
    period.save_report('recoveries.txt', output)
    
        
if  __name__ == "__main__":
    d = db.load_state()
    create_recovery_report(d)
    princ("Finished")