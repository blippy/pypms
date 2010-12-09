# Prepare Mobil work statement

import pdb

import db
#import data
import common
from common import aggregate



###########################################################################

def total_text(amount):
    fmt = '\n{0:10} {1:12} {2:>9} {3:>9} {4:>9.2f}\n\n'
    return fmt.format('', 'TOTAL', '', '', amount)

def create_job_details(job_code, cache, rates):
    'Create a report section for each mobil job'
    output = ''    
    fmt1 = '{0:10} {1:12} {2:9.2f} {3:9.2f} {4:9.2f}\n'
    fmt2 = '\n{0:10} {1:12} {2:>9} {3:>9} {4:>9} {5}\n'
    
    
    # print out the header information
    output += 'SREL Job no: {0}\n'.format(job_code)
    output += cache['jobs'][job_code]['references'] + '\n'
    output += fmt2.format('Serv Mast', 'Personnel', 'Rate', 'Qty', ' Total', 'Timesheet Numbers')
    
    num_items = 0
    total = 0.0
    
    # aggregate to the person level
    invItems = [x for x in cache['timeItems'] if x['JobCode'] == job_code]
    for initials, values in aggregate(invItems, common.mkKeyFunc('Person')):
        time_spent= common.summate(values, lambda x: x['TimeVal'])
        e = cache['employees'][initials]
        smn = e['MobilSmn']
        person = e['PersonNAME']
        #print rates
        rate = rates[(job_code, initials)]
        amount = round(time_spent * rate, 2)
        output += fmt1.format(smn, person, time_spent, rate, amount)
        total += amount
        num_items += 1
        
    # expenses
    expense_total = 0.0
    for expense in cache['expenses']:
        if expense['JobCode'] != job_code: continue
        expense_total += expense['Amount']
    exp_factor = cache['jobs'][job_code]['exp_factor']
    expenses_out = expense_total * exp_factor
    total += expenses_out
    if expense_total != 0:
        output += fmt1.format('3480447', 'Expenses', expense_total, exp_factor, expenses_out)
        num_items += 1
        
    # total
    output += total_text(total)
    output += '\n' + '-' * 53 + '\n\n'
        
    if num_items == 0: output = '' # just ignore everything when there's nothing interesting

    return total, output


###########################################################################

def main(d):
    #conf = config()
    #conf.load()                        
    #print conf.jobs
    output_text = 'SMITH REA ENERGY LIMITED - STATEMENT OF WORK - {0}\n\n'.format(d['period'].mmmmyyyy())

    
    # obtain rates for jobs
    rates = {}
    charges = d['charges']
    for key in charges.keys():
        job_code, task, initials = key
        rate = common.dget(charges, key)
        rates[(job_code, initials)] = rate
    
    
    total = 0.0
    jobs = d['jobs']
    for job_code in sorted(jobs.keys()):
        if jobs[job_code]['briefclient'] != 1: continue # 1 == mobil
        amount, text = create_job_details(job_code, d, rates)
        output_text += text
        total += amount
    
    output_text += total_text(total)
    output_text += '\n\nName: ' + '_' * 30 + ' Signature: ' + '_' * 30 + ' Date: ' + '_' * 15
    
    common.save_report(d['period'], "mobil.txt", output_text)

if  __name__ == "__main__":
    d = db.fetch()
    main(d)
    print 'Finished'