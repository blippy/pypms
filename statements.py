# Create the monthly workstatements for each job

#import itertools
import pdb
#from itertools import groupby, ifilter
from operator import itemgetter, attrgetter

import common, rtf
from common import aggregate, princ, print_timing
import db
import expenses
import period

debug = False


###########################################################################

# FIXME - probable bad idea
typeWork, typeExpense = range(2) # create enumerated invoice type





###########################################################################

def AddTopInfo(out, job):
    address = job['address'] # whole address
    if address is None: address = ''
    address = address.split('\r')
    address1 = address[0] # first line of address
    out.add(address1)    
    out.add(job['title'])
    job_code = job['job']
    out.add("SREL job code: " + job_code, 2)
    
    #if job_code  == '2638' : pdb.set_trace()
    references = job['references']
    if references is None: references = ''
    references = references.replace('\n', '\\par \n')
    #references = common.AsAscii(references)
    out.add(references, 2)
    
    heading =  '{0:44s}{1:>9s}{2:>9s}{3:>9s}'.format('', 'Quantity', 'Price', 'Value')
    out.add(heading,  2)


 
###########################################################################

# TODO remove this eventually
class Section:
    def __init__(self):
        self.heading = 'Fill with something sensible'
        self.workEntries = {}
        self.expenseEntries = []
        
    def AddWork(self, item, price, personName):
        if not self.workEntries.has_key(personName): self.workEntries[personName] = { 'desc' : personName , 'qty' : 0, 'price' : price }
        self.workEntries[personName]['qty'] += item['TimeVal']
        
    def Work(self):
        items = self.workEntries.values()
        items.sort(key=lambda x: x['desc'])
        return items
    
    def AddExpense(self, item, exp_factor):
        desc = ' - '.join([item['Period'], item['Name'], item['Desc']])
        self.expenseEntries.append( { 'desc' : desc , 'qty' : exp_factor, 'price' : item['Amount'] })
        

 
###########################################################################
 
class LineItem: pass

def subtotal(out, desc, value):
    'Output a subtotal'
    #FIXME - devise a better way so that the format depends on whether something is a string or a number
    text = '   {0:41.41s}{1:9.2s}{2:9.2s}{3:9.2f}'.format(desc , '', '', value)
    out.add(text)
        
def ProcessSubsection(out, items, subtotalTitle):
    if len(items) == 0: return 0, 0 # don't do anything if there are no items
    total = 0
    for el in items:
        value = round(el.qty * el.price, 2)
        total += value
        text = '   {0:41.41s}{1:9.2f}{2:9.2f}{3:9.2f}'.format(el.desc , el.qty, el.price, value)
        out.add(text)
       
    numItems = len(items)
    if numItems >1: subtotal(out, subtotalTitle, total)
    out.para()
    return total, numItems
    
###########################################################################

def create_job_statement(job, all_tasks, exps, times):


    title = "Work Statement: %s" % (period.mmmmyyyy())
    out = rtf.Rtf()
    out.addTitle(title)
    AddTopInfo(out, job)          


    tasks = map(lambda o: getattr(o, 'task'), times + exps)
    tasks = common.unique(tasks)
    tasks.sort()
    if tasks[0] == '' and len(tasks)>1: tasks = tasks[1:] + [tasks[0]] # rotate unassigned to end

    # distribute invoice items into sections
    job_code = job['job']
    sections = [] # a list of Section classes
    for task_key in tasks:
        
        # work out heading
        if len(task_key) == 0: heading = 'Expenses not categorised to a specific task'
        else: 
            desc = all_tasks[(job_code, task_key)]['JCDescription']
            heading = 'Task {0}: {1}'.format(task_key,desc)
        out.add(heading)
        
 
        ProcessSubsection(out, times, 'Work subtotal')        
        ProcessSubsection(out, exps, 'Expenses subtotal')
        
        
        # TODO NOW
        print job_code, " ", heading
        #s = create_section(task)
        #sections.append(s)
    return

    for taskKey, taskGroup in aggregate(invItems, common.mkKeyFunc('Task')):
        s = Section()
        if taskKey == '' : 
            s.heading = 'Expenses not categorised to a specific task'
            s.ordering = 'ZZZ'
        else:
            s.heading = 'Task {0}: {1}'.format(taskKey, db.task_desc(d, jobKey, taskKey))
            s.ordering = s.heading
            
        for item in taskGroup: 
            # FIXME - I think we aggregated work and expenses together - now we're splitting them out - which doesn't make sense - we should have kept them spearate in the first place
            if item['iType'] == typeWork: 
                person = item['Person']
                price = d['charges'][(jobKey, taskKey, person)]
                personName = db.initials_to_name(d, person)
                s.AddWork(item, price, personName)
            else: s.AddExpense(item, exp_factor)
        sections.append(s)
    

        
    # output the sections
    sections.sort(key= lambda x: x.ordering)    
    totalWork = 0
    totalExpenses = 0
    for s in sections:
        out.add(s.heading)
        work, numPeople = ProcessSubsection(out, s.Work(), 'Work subtotal')
        expenses, numExpenses = ProcessSubsection(out, s.expenseEntries, 'Expenses subtotal')
        totalWork += work
        totalExpenses += expenses
        if numPeople > 0 and numExpenses >0: 
            subtotal(out, 'Task subtotal', work+expenses)
            out.para()
        
    # output grand summary
    out.add('Overall summary')
    subtotal(out, 'Work total', totalWork)
    subtotal(out, 'Expenses total', totalExpenses)
    net = totalWork+totalExpenses
    subtotal(out, 'Net total', net)

     
    out.annotation(job, '')
    #out.save(d['period'].outDir() + '\\statements', jobKey + '.rtf')
    outdir = period.perioddir() + '\\statements'
    outfile = jobKey + '.rtf'
    out.save(outdir, outfile)
    
    # remember what we have produced for the invoice summaries
    if job['Weird'] or job['WIP']: 
        billed = 0.0
    else:
        billed = net
    invoice = { 'work': totalWork , 'expenses': totalExpenses, 'net': billed}
    d['auto_invoices'][jobKey] = invoice


###########################################################################

def create_statements(d):
    
    # TODO - print a warning if exp_factor > 1.05
    work_codes = set([common.AsAscii(t['JobCode']) for t in d['timeItems']])
    expense_codes = set([e['JobCode'] for e in expenses.cache.expenses])
    job_codes = list(work_codes.union(expense_codes))
    job_codes.sort()

    for job_code in job_codes:
        #pdb.set_trace()
        if job_code[0:2] == "01": continue
        job = d['jobs'][job_code]        
        
        exps = []
        for exp in expenses.cache.expenses:
            if exp['JobCode'] <> job_code: continue
            item = LineItem()
            item.task = exp['Task']
            item.desc = '{0} - {1} - {2}'.format(exp['Period'], exp['Name'], exp['Desc'])
            item.qty  = job['exp_factor']
            item.price = exp['Amount']
            exps.append(item)

        times = []
        for tim in d['timeItems']:
            if tim['JobCode'] <> job_code: continue
            item = LineItem()
            #if job_code == '2825': pdb.set_trace()
            item.task = tim['Task']
            item.desc = tim['Person']
            item.qty = tim['TimeVal']
            item.price = job['exp_factor']
            times.append(item)
            
        if len(exps) > 0 or len(times)> 0:
            create_job_statement(job, d['tasks'], exps, times)
        


###########################################################################

if  __name__ == "__main__":
    debug = True
    data = db.load_state()
    expenses.cache.import_expenses()
    create_statements(data)
    princ("Finished")
