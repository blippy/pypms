# Prepare Mobil work statement

import pdb

import data
import common
from common import aggregate

###########################################################################


class config:
    def __init__(self):
        fileName = 'M:\Finance\camel\mobil.txt'
        self.fin = file(fileName)
        self.lines = self.fin.readlines()
        self.fin.close()
    
    def __readLine(self):
        line = self.lines[0].strip()
        self.lines = self.lines[1:]
        return line        
    
    def __more(self): return len(self.lines) > 0

    def load(self):
        'Load in the configuration information'
        self.jobs = {}
        self.smns = {}
        while self.__more():
            line = self.__readLine()
            #print '*' , line , '*'
            if line == 'job':
                job = {}
                for field in ['job', 'contact', 'title', 'desc', 'po' , 'units']:
                    job[field] = self.__readLine()
                self.jobs[job['job']] = job
            if line == 'smn':
                smn = {}
                for field in ['initials', 'smn']:
                    smn[field] = self.__readLine()
                self.smns[smn['initials']] = smn
                
    def get(self, job_code):
        #pdb.set_trace()
        return self.jobs[job_code]
        
 


###########################################################################

def total_text(amount):
    fmt = '\n{0:10} {1:12} {2:>9} {3:>9} {4:>9.2f}\n\n'
    return fmt.format('', 'TOTAL', '', '', amount)

def create_job_details(job_code, conf, cache, rates):
    'Create a report section for each mobil job'
    output = ''    
    fmt1 = '{0:10} {1:12} {2:9.2f} {3:9.2f} {4:9.2f}\n'
    fmt2 = '\n{0:10} {1:12} {2:>9} {3:>9} {4:>9} {5}\n'
    
    
    # print out the header information
    headers = [ ('PO', 'po') , ('SREL Job no', 'job'), ('Mobil contact', 'contact'), ('Project Title', 'title'), ('Description of work', 'desc'), ('Units', 'units')]
    #fmt1 = '%15s: %s\n'
    job = conf.jobs[job_code]
    for desc, field in headers:        
        output += '{0:<20}: {1}\n'.format(desc, job[field])
    output += fmt2.format('Serv Mast', 'Personnel', 'Rate', 'Qty', ' Total', 'Timesheet Numbers')
    
    #print conf_details
    
    
    num_items = 0
    total = 0.0
    
    # aggregate to the person level
    invItems = [x for x in cache['timeItems'] if x['JobCode'] == job_code]
    for initials, values in aggregate(invItems, common.mkKeyFunc('Person')):
        time_spent= common.summate(values, lambda x: x['TimeVal'])
        smn = conf.smns[initials]['smn']
        person = cache['employees'][initials]['PersonNAME']
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

def main(cache):
    conf = config()
    conf.load()                        
    #print conf.jobs
    output_text = 'SMITH REA ENERGY LIMITED - STATEMENT OF WORK - {0}\n\n'.format(cache['period'].mmmmyyyy())

    
    # obtain rates for jobs
    rates = {}
    charges = cache['charges']
    for key in charges.keys():
        job_code, task, initials = key
        rate = common.dget(charges, key)
        rates[(job_code, initials)] = rate
    
    
    total = 0.0
    job_codes = conf.jobs.keys()
    job_codes.sort()
    for job_code in job_codes:
        amount, text = create_job_details(job_code, conf, cache, rates)
        output_text += text
        total += amount
    
    output_text += total_text(total)
    
    common.save_report(cache['period'], "mobil.txt", output_text)

if  __name__ == "__main__":
    data.run_current(main)
    print 'Finished'