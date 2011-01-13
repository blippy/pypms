import pdb

import common
from common import dget
import db

def get_budget_line(d, job_code, cum_utilised):
    fmt = '{0:<7s} {1:<10.10s} {2:<10.10s} {3:<20.20s} {4:10.2f} {5:10.10s} {6:10.10s} {7:11.2f}\n'
    job = d['jobs'][job_code]
    title = job['title']
    client = dget(d['clients'], job['briefclient'], '')
    po = job['references']
    po = po.split('\n')
    po = po[0]
    po = po.replace('\r', '')
    budget = job['PoBudget']
    start = job['PoStartDate']
    end = job['PoEndDate']
    utilised_pos = job['UtilisedPOs']
    remaining_budget = budget + utilised_pos - cum_utilised
    text = fmt.format(job_code, title, client, po, budget, start, end, remaining_budget)
    return text

###########################################################################

def main(d):
    sql = 'SELECT InvJobCode, Sum(InvUBI+InvWIP+InvInvoice) AS SumUsed FROM tblInvoice GROUP BY InvJobCode'
    fieldspec = [('InvJobCode', str), ('SumUsed', float)]
    recs = db.RecordsList(sql, fieldspec)
    utilisations = {}    
    for r in recs: utilisations[r['InvJobCode']] = r['SumUsed']
    
    p = d['period']
    fmt = '{0:<7s} {1:<10s} {2:10s} {3:20.20s} {4:>10.10s} {5:<10.10s} {6:<10.10s} {7:>11.11s}\n'
    text = fmt.format('JobNo', 'Title', "Company", "PO", 'Budget', 'Start Date', 'End Date', 'Budget Left')
    for job_code in sorted(d['jobs'].keys()):
        if not d['jobs'][job_code]['active']: continue
        cum_utilised = dget(utilisations, job_code, 0.0)
        text += get_budget_line(d, job_code, cum_utilised)
    common.save_report(p, "budget.txt", text)

if  __name__ == "__main__":
    print "Didn't do anything"