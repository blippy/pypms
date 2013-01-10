import itertools
import traceback

from common import dget, princ
import common
import excel

        
def tweaked_jobs(the_tweaks):
    jobs = []
    for a_tweak in the_tweaks:
        job = a_tweak['JobCode']
        if job not in jobs: jobs.append(job)
    jobs.sort()
    return jobs

def tweak_recoveries(the_tweaks):
    #princ(the_tweaks)
    recoveries = {}
    for tweak in the_tweaks:
        #print line
        job_code = tweak['JobCode']
        amount = tweak['Recovery']
        if amount == 0.0: continue
        comment = tweak['Comment']
        if not recoveries.has_key(job_code): recoveries[job_code] = []
        recoveries[job_code].append((amount, comment))
    return recoveries

def accum_tweaks_to_job_level(the_tweaks):
    
    jobs = {}
    for entry in the_tweaks:
        code = entry['JobCode']
        if not jobs.has_key(code): jobs[code] = {}
        job = jobs[code]
        for key in ['Inv3rdParty', 'InvAccrual', 'InvBIA', 'InvInvoice', 'InvTime', 'InvUBI', 'InvWIP']:
            common.dplus(job, key, common.dget(entry, key))
    return jobs