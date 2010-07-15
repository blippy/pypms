'''Process the invoice tweaks'''

import itertools 

#=[InvBIA]+[InvUBI]+[InvWIP]+[InvAccrual]+[InvInvoice]-[Inv3rdParty]-[InvTime]
import common
import data
import db
import excel

###########################################################################

def read_excel_file(d):
    xl = excel.ImportWorksheet(common.camelxls(d.p), 'InvTweaks')
    field_names = ['Job', 'InvBIA', 'InvUBI', 'InvWIP', 'InvAccrual', 'InvInvoice', 'Inv3rdParty', 'InvTime', 'Recovery', 'Comment']
    f = common.AsFloat
    field_types = [str, f, f, f, f, f, f, f, f, str]
    tweaks = []

    for line in xl[1:]:
        tweak = {}
        for field, value, converter in itertools.izip(field_names, line, field_types):
            tweak[field] = converter(value)
        tweaks.append(tweak)
    d.invoice_tweaks = tweaks

def accumulate(d):
    '''Accumulate invoice tweaks to job level'''
    jobs = {}
    for entry in d.invoice_tweaks:
        code = entry['Job']
        if not jobs.has_key(code): jobs[code] = {}
        job = jobs[code]
        for key in ['Inv3rdParty', 'InvAccrual', 'InvBIA', 'InvInvoice', 'InvTime', 'InvUBI', 'InvWIP']:
            common.dplus(job, key, common.dget(entry, key))
    return jobs

###########################################################################

def main(d):
    read_excel_file(d)
    #print accumulate(d)
    
if  __name__ == "__main__":
    data.run_current(main)
    print 'Finished'