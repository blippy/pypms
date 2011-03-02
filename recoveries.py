'''Spot recoveries in PMS
'''

import operator

import common
from common import dget, princ
import db
import excel


###########################################################################

def get_dbase_recoveries(d):
    #=[InvBIA]+[InvUBI]+[InvWIP]+[InvAccrual]+[InvInvoice]-[Inv3rdParty]-[InvTime]
    field_list = ['InvJobCode', 'InvComments', 'InvBIA', 'InvUBI', 'InvWIP', 'InvAccrual', 'InvInvoice', 'Inv3rdParty', 'InvTime']
    invoices = db.GetInvoices(d, field_list)    

    recoveries = {}
    for invoice in invoices:
        job_code = common.AsAscii(invoice[0])
        comment = invoice[1]
        as_floats = map(float, invoice[2:])
        InvBIA, InvUBI, InvWIP, InvAccrual, InvInvoice, Inv3rdParty, InvTime = as_floats
        recovery = InvBIA + InvUBI + InvWIP + InvAccrual + InvInvoice - Inv3rdParty - InvTime
        if abs(recovery)< 20.0: continue # ignore immaterial recoveries
        recoveries[job_code] = recovery
    return recoveries

def get_camel_recoveries(d):
    xl = excel.ImportWorksheet(common.camelxls(d['period']), 'InvTweaks')
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
def main(d):
    db_recoveries = get_dbase_recoveries(d)
    camel_recoveries = get_camel_recoveries(d)    
    output = 'RECOVERY RECONILIATION\n\n'
    
    def line(amount, comment): return  '    %9.2f %s\n' % (amount, comment)

    tweaks_grand_total = 0.0
    pms_grand_total = 0.0
    for job_code in common.combine_dict_keys([camel_recoveries, db_recoveries]):
        job = d['jobs'][job_code] # FIXME - d['jobs'][job_code] is a common idiom which ought to be abstracted
        wip = common.tri(job['WIP'], 'wip' , '') 
        weird = common.tri(job['Weird'], 'weird ', '')
        output += 'JOB %s %s %s\n' % (job_code, wip, weird)
        
        #output += '  Per InvTweaks\n'
        tweak_total = 0.0
        for tweak in dget(camel_recoveries, job_code, []):
            amount, comment = tweak
            tweak_total += amount
            output += line(amount, comment)
        output += line(tweak_total , 'TWEAK TOTAL')
        tweaks_grand_total += tweak_total
            
        recovery = dget(db_recoveries, job_code)
        output += line(recovery , 'PMS RECOVERY')
        diff = tweak_total - recovery
        flag = common.tri(abs(diff) > 20.0, ' **** CHECK THIS', '')
        output += line(diff , 'DIFF' + flag + '\n\n')
        pms_grand_total += recovery
       
    output += 'SUMMARY:\n'
    output += line(tweaks_grand_total, 'TWEAKS GRAND TOTAL')
    output += line(pms_grand_total, 'PMS GRAND TOTAL')
    output += line(tweaks_grand_total - pms_grand_total, 'OVERALL DIFF')        
    common.save_report(d['period'], 'recoveries.txt', output)
    
        
if  __name__ == "__main__":
    princ("Didn't do anything")