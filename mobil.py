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

def create_job_details(job_code, conf, cache, rates):
    'Create a report section for each mobil job'
    output = ''
    
    # print out the header information
    headers = [ ('PO', 'po') , ('SREL Job no', 'job'), ('Mobil contact', 'contact'), ('Project Title', 'title'), ('Description of work', 'desc'), ('Units', 'units')]
    #fmt1 = '%15s: %s\n'
    job = conf.jobs[job_code]
    for desc, field in headers:        
        output += '{0:<20}: {1}\n'.format(desc, job[field])
    #print conf_details
    
    # aggregate to the person level
    invItems = [x for x in cache['timeItems'] if x['JobCode'] == job_code]
    for initials, values in aggregate(invItems, common.mkKeyFunc('Person')):
        time_spent= common.summate(values, lambda x: x['TimeVal'])
        smn = conf.smns[initials]['smn']
        person = cache['employees'][initials]['PersonNAME']
        rate = rates[(job_code, initials)]
        amount = round(time_spent * rate, 2)
        output += '{0} {1} {2} {3} {4}\n'.format(smn, person, time_spent, rate, amount)
    # FIXME NOW
    return output


###########################################################################

def main(cache):
    conf = config()
    conf.load()                        
    #print conf.jobs
    output_text = 'SMITH REA ENERGY LIMITED - STATEMENT OF WORK\n\n'
    output_text += 'Service Master No	Personnel	Rate	No of Hrs	No of Days	Total	Timesheet Numbers\n'

    
    # obtain rates for jobs
    rates = {}
    charges = cache['charges']
    for key in charges.keys():
        job_code, task, initials = key
        rate = common.dget(charges, key)
        rates[(job_code, initials)] = rate
    
    
    job_codes = conf.jobs.keys()
    job_codes.sort()
    for job_code in job_codes:
        output_text += create_job_details(job_code, conf, cache, rates)
    
    common.save_report(cache['period'], "mobil.txt", output_text)
    # FIXME Outputs: reports\mobil.txt

if  __name__ == "__main__":
    data.run_current(main)
    print 'Finished'