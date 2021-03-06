# Create the monthly workstatements for each job

import itertools
import pdb
#from itertools import groupby, ifilter
from operator import itemgetter, attrgetter
import pprint

import common, rtf
from common import aggregate, princ
import db
import expenses
import period

debug = False

###########################################################################

def AddTopInfo(out, job):
    address = job['address'] # whole address
    if address is None: address = ''
    address = address.split('\r')
    address1 = address[0] # first line of address
    out.add(address1)    
    out.add(job['title'])
    job_code = job['job']
    out.add("IMA job code: " + job_code, 2)
    
    #if job_code  == '2638' : pdb.set_trace()
    references = job['references']
    if references is None: references = ''
    references = references.replace('\n', '\\par \n')
    #references = common.AsAscii(references)
    out.add(references, 2)
    
    heading =  '{0:44s}{1:>9s}{2:>9s}{3:>9s}'.format('', 'Quantity', 'Price', 'Value')
    out.add(heading,  2)
 
###########################################################################
 
class LineItem: pass

def subtotal(out, desc, value):
    'Output a subtotal'
    #TODO - devise a better way so that the format depends on whether something is a string or a number
    text = '   {0:41.41s}{1:9.2s}{2:9.2s}{3:9.2f}'.format(desc , '', '', value)
    out.add(text)
        
def ProcessSubsection(out, items, task_key, subtotalTitle):
    items = filter(lambda x: x.task == task_key, items)
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


    title = "Work Statement: {0}".format(period.mmmmyyyy())
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
    totalWork = 0.0
    totalExpenses = 0.0
    for task_key in tasks:
        
        # work out heading
        if len(task_key) == 0: heading = 'Expenses not categorised to a specific task'
        else: 
            desc = all_tasks[(job_code, task_key)]['TaskDes']
            heading = 'Task {0}: {1}'.format(task_key,desc)
        out.add(heading)
        
        amount_work, num_times = ProcessSubsection(out, times, task_key, 'Work subtotal')
        totalWork += amount_work
        
        amount_expenses, num_expenses = ProcessSubsection(out, exps, task_key, 'Expenses subtotal')
        totalExpenses += amount_expenses
        
        if num_times > 0 and num_expenses > 0:
            subtotal(out, 'Task subtotal', amount_work + amount_expenses)
            out.para()
        
    # output grand summary
    out.add('Overall summary')
    subtotal(out, 'Work total', totalWork)
    subtotal(out, 'Expenses total', totalExpenses)
    net = totalWork+totalExpenses
    subtotal(out, 'Net total', net)

     
    out.annotation(job)
    outdir = period.perioddir() + '\\statements'
    outfile = job_code + '.rtf'
    out.save(outdir, outfile)
    
    # remember what we have produced for the invoice summaries
    if job['Weird'] or job['WIP']: 
        billed = 0.0
    else:
        billed = net
                        
    invoice = { 'work': totalWork , 'expenses': totalExpenses, 'net': billed}
    return invoice


###########################################################################

def create_statements(data):
    
    # TODO - print a warning if exp_factor > 1.05 
    work_codes = set([common.AsAscii(t['JobCode']) for t in data['timeItems']])
    the_expenses = data['Expenses']
    expense_codes = set([e['JobCode'] for e in the_expenses])
    job_codes = list(work_codes.union(expense_codes))
    job_codes.sort()
    if data['auto_invoices'] is None: data['auto_invoices'] = {}

    for job_code in job_codes:
        if job_code[0:2] == "01": continue
        job = data['jobs'][job_code]        
        
        exps = []
        for exp in the_expenses:
            if exp['JobCode'] <> job_code: continue
            item = LineItem()
            item.task = exp['Task']
            item.desc = '{0} - {1} - {2}'.format(exp['Period'], exp['Name'], exp['Desc'])
            item.qty  = job['exp_factor']
            item.price = exp['Amount']
            exps.append(item)

        times = []
        times1 = filter(lambda x: x['JobCode'] == job_code, data['timeItems'])
        times2 = common.summate_to_dict(times1, lambda x: (x['Task'], x['Person']), common.mkKeyFunc('TimeVal'))
        keys = times2.keys()
        keys = sorted(keys, key = lambda x: x[0] + ' ' + x[1])
        for k in keys:
            item = LineItem()
            item.task = k[0]
            initials = k[1]
            item.desc = db.initials_to_name(data, initials)
            item.qty = times2[k]
            item.price = price = data['charges'][(job_code, item.task, initials)]
            times.append(item)
            
        if len(exps) + len(times)> 0:
            invoice = create_job_statement(job, data['tasks'], exps, times)            
            data['auto_invoices'][job_code] = invoice

        


###########################################################################

class Source:
    def __init__(self):
        self.foo = 1
        self.bar = 2
        self.baz = 'zoing'
        
    def __str__(self):
        return str(self.__dict__)
        
if  __name__ == "__main__":
    debug = True
    data = db.load_state()
    #pprint.pprint(data['Expenses'])
    #s = Source()
    #print s.foo
    #pprint.pprint(props(s))
    #expenses.cache.import_expenses()
    create_statements(data)
    princ("Finished")
