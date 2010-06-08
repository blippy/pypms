import json
#import pickle

import yaml

import db, expenses, period





###########################################################################

class Data:
    
    yamlFile = 'M:\\Finance\\pypms\\out\\current.yml'
    
    def __init__(self, p = None):
        if not p: p = period.Period(usePrev = True)
        self.p = p # period

    def load(self):
        p = self.p
        self.employees = db.GetEmployees(p)
        self.jobs = db.GetJobs()
        self.tasks = db.GetTasks(p)
        self.timeItems = db.GetTimeitems(p)
        self.charges = db.GetCharges(p)
        self.expenses = expenses.main(p)
        self.invoices = {}
        

    def store(self): 
        with open(self.__class__.yamlFile, 'w') as f: yaml.dump(self.__dict__, f)
        
        # create a shortened dictionary for developer convenience
        return # maybe reinstate this later
        short = {}
        for key in self.__dict__.keys():
            short[key] = self.__dict__[key][:5]
        print short
        
    def restore(self): 
        with open(self.__class__.yamlFile, 'r') as f: self.__dict__ = yaml.load(f)




   
###########################################################################


    

if  __name__ == "__main__":
    # perform a simple test
    p = period.Period(usePrev = True)
    d = Data()
    d.load()
    d.store()
    print "Finished"
