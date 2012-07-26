import itertools
import traceback

from common import dget, princ, print_timing
import common

#def accumulate_tweaks(d, xl):
def decipher_tweaks(xl):    

    #xl = excel.ImportCamelWorksheet('InvTweaks')
    field_names = ['Job', 'InvBIA', 'InvUBI', 'InvWIP', 'InvAccrual', 'InvInvoice', 'Inv3rdParty', 'InvTime', 'Recovery', 'Comment']
    f = common.AsFloat
    field_types = [str, f, f, f, f, f, f, f, f, str]
    the_tweaks = []

    all_ok = True
    for line_num in range(2, len(xl)):
        line = xl[line_num]
        #for line in xl[1:]:
            
        tweak = {}
        tweak_ok = True
        for field, value, converter in itertools.izip(field_names, line, field_types):
            try:
                tweak[field] = converter(value)
            except ValueError:
                all_ok = False
                tweak_ok = False
                msg = "Problem with tweaks: line: {0}, field name: {1}, Contents: '{2}'".format(line_num+1, field, value)
                princ(msg)
                princ(traceback.format_exc())
        if tweak_ok and len(tweak['Job']) > 0: 
            the_tweaks.append(tweak) # TODO handle case Job isn't found
        
    if not all_ok:
        raise common.DataIntegrityError("Error encountered in tweaks worksheet. Further processing abandoned")

    #princ(the_tweaks)

    return the_tweaks
        
def tweaked_jobs(the_tweaks):
    jobs = []
    for a_tweak in the_tweaks:
        job = a_tweak['Job']
        if job not in jobs: jobs.append(job)
    jobs.sort()
    return jobs

def tweak_recoveries(the_tweaks):
    #princ(the_tweaks)
    recoveries = {}
    for tweak in the_tweaks:
        #print line
        job_code = tweak['Job']
        amount = tweak['Recovery']
        if amount == 0.0: continue
        comment = tweak['Comment']
        if not recoveries.has_key(job_code): recoveries[job_code] = []
        recoveries[job_code].append((amount, comment))
    return recoveries

def accum_tweaks_to_job_level(the_tweaks):
    
    jobs = {}
    for entry in the_tweaks:
        code = entry['Job']
        if not jobs.has_key(code): jobs[code] = {}
        job = jobs[code]
        for key in ['Inv3rdParty', 'InvAccrual', 'InvBIA', 'InvInvoice', 'InvTime', 'InvUBI', 'InvWIP']:
            common.dplus(job, key, common.dget(entry, key))
    return jobs