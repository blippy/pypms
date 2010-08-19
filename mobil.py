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
        while self.__more():
            line = self.__readLine()
            #print '*' , line , '*'
            if line == 'job':
                job = {}
                for field in ['job', 'contact', 'title', 'desc', 'po' , 'units']:
                    job[field] = self.__readLine()
                self.jobs[job['job']] = job
                
    def get(self, job_code):
        #pdb.set_trace()
        return self.jobs[job_code]
        
 
###########################################################################

def create_job_details(job_code, conf_details, cache):
    'Create a report section for each mobil job'
    output = ''
    
    # print out the header information
    headers = [ ('PO', 'po') , ('SREL Job no', 'job'), ('Mobil contact', 'contact'), ('Project Title', 'title'), ('Description of work', 'desc')]
    #fmt1 = '%15s: %s\n'
    for desc, field in headers:        
        output += '{0:<20}: {1}\n'.format(desc, conf_details[field])
    #print conf_details
    
    # aggregate to the person level
    invItems = [x for x in cache['timeItems'] if x['JobCode'] == job_code]
    for person, values in aggregate(invItems, common.mkKeyFunc('Person')):
        time_spent= common.summate(values, lambda x: x['TimeVal'])
        output += '{0} {1}\n'.format(person, time_spent)
    # FIXME NOW
    print output


###########################################################################

def main(cache):
    conf = config()
    conf.load()                        
    #print conf.jobs
    output_text = 'SMITH REA ENERGY LIMITED - STATEMENT OF WORK\n\n'
    output_text += 'Service Master No	Personnel	Rate	No of Hrs	No of Days	Total	Timesheet Numbers\n'

    
    job_codes = conf.jobs.keys()
    job_codes.sort()
    for job_code in job_codes:
        create_job_details(job_code, conf.get(job_code), cache)
    
    print output_text
    # FIXME Outputs: reports\mobil.txt

if  __name__ == "__main__":
    data.run_current(main)
    print 'Finished'