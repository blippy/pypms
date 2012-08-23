import pdb

import common
from common import princ
from common import dget, print_timing
import db
import period

def get_budget_line(d, job_code, cum_utilised):
    fmt = '{0:<7s} {1:<10.10s} {2:<10.10s} {3:<20.20s} {4:10.2f} {5:<10.10s} {6:<10.10s} {7:11.2f}'
    job = d['jobs'][job_code]
    title = job['title']
    client = dget(d['clients'], job['briefclient'], '')
    po = dget(job, 'references', '')
    po = po.split('\n')
    po = po[0]
    po = po.replace('\r', '')
    budget = dget(job, 'PoBudget', 0.0)
    def as_str(v): 
        if v is None: return '' 
        else: return str(v)
    start = as_str(job['PoStartDate'])
    #if start is not None: pdb.set_trace()
    end = as_str(job['PoEndDate'])
    utilised_pos = dget(job, 'UtilisedPOs', 0.0)
    remaining_budget = budget + utilised_pos - cum_utilised
    text = fmt.format(job_code, title, client, po, budget, start, end, remaining_budget)
    return text

###########################################################################

@print_timing
def create_budget(d):
    
    if True: return # this module should be considered semi-deprecated because it's not really serving much of a purpose

    sql = 'SELECT InvJobCode, Sum(InvUBI+InvWIP+InvInvoice) AS SumUsed FROM tblInvoice GROUP BY InvJobCode'
    fields = 'InvJobCode,SumUsed'
    recs = db.fetch_and_dictify(sql, fields)
    utilisations = {}    
    for r in recs: utilisations[r['InvJobCode']] = r['SumUsed']
    
    fmt = '{0:<7s} {1:<10s} {2:10s} {3:20.20s} {4:>10.10s} {5:<10.10s} {6:<10.10s} {7:>11.11s}\n'
    text = [fmt.format('JobNo', 'Title', "Company", "PO", 'Budget', 'Start Date', 'End Date', 'Budget Left')]
    for job_code in sorted(d['jobs'].keys()):
        if not d['jobs'][job_code]['active']: continue
        cum_utilised = dget(utilisations, job_code, 0.0)
        text.append(get_budget_line(d, job_code, cum_utilised))
    period.save_report("budget.txt", text)

    
if  __name__ == "__main__":
    data = db.load_state()
    create_budget(data)
    princ("Finished")    