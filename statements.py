# Create the monthly workstatements for each job

#import itertools
import pdb
#from itertools import groupby, ifilter
from operator import itemgetter, attrgetter

import common, rtf
from common import aggregate
import period


###########################################################################

# FIXME - probable bad idea
typeWork, typeExpense = range(2) # create enumerated invoice type





###########################################################################

def AddTopInfo(out, job):
    address = job['address'] # whole address
    address = address.split('\r')
    address1 = address[0] # first line of address
    out.add(address1)    
    out.add(job['title'])
    job_code = job['job']
    out.add("SREL job code: " + job_code, 2)
    
    #if job_code  == '2638' : pdb.set_trace()
    references = job['references']
    references = references.replace('\n', '\\par \n')
    #references = common.AsAscii(references)
    out.add(references, 2)
    
    heading =  '{0:44s}{1:>9s}{2:>9s}{3:>9s}'.format('', 'Quantity', 'Price', 'Value')
    out.add(heading,  2)


 
###########################################################################

class Section:
    def __init__(self):
        self.heading = 'Fill with something sensible'
        self.workEntries = {}
        self.expenseEntries = []
        
    def AddWork(self, item, price, personName):
        #print personName
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
 
def subtotal(out, desc, value):
    'Output a subtotal'
    #FIXME - devise a better way so that the format depends on whether something is a string or a number
    text = '   {0:41.41s}{1:9.2s}{2:9.2s}{3:9.2f}'.format(desc , '', '', value)
    out.add(text)
        
def ProcessSubsection(out, items, subtotalTitle):
    if len(items) == 0: return 0, 0 # don't do anything if there are no items
    total = 0
    for el in items:
        desc, qty, price = el['desc'], el['qty'], el['price']
        value = round(qty * price, 2)
        total += value
        text = '   {0:41.41s}{1:9.2f}{2:9.2f}{3:9.2f}'.format(desc , qty, price, value)
        out.add(text)
       
    numItems = len(items)
    if numItems >1: subtotal(out, subtotalTitle, total)
    out.para()
    return total, numItems
    
def CreateJobStatment(jobKey, invItems, d):
    if jobKey[0:2] == "01": return

    job = d['jobs'][jobKey]
    #if job['Weird']: return # we should even create weird invoices - we their time and expenses anyway
    
    # out = Output(d.p) # deprecated method
    title = "Work Statement: %s" % (d['period'].mmmmyyyy())
    out = rtf.Rtf()
    out.addTitle(title)
    AddTopInfo(out, job)
    
    invItems = list(invItems)
      
    exp_factor = job['exp_factor']
    # FIXME - print a warning if exp_factor > 1.05
    
    # distribute invoice items into sections
    sections = [] # a list of Section classes
    for taskKey, taskGroup in aggregate(invItems, common.mkKeyFunc('Task')):
        s = Section()
        if taskKey == '' : 
            s.heading = 'Expenses not categorised to a specific task'
            s.ordering = 'ZZZ'
        else:
            s.heading = 'Task %s: %s' % (taskKey, d['tasks'][(jobKey, taskKey)]['TaskDes'])
            s.ordering = s.heading
            
        for item in taskGroup: 
            # FIXME - I think we aggregated work and expenses together - now we're splitting them out - which doesn't make sense - we should have kept them spearate in the first place
            if item['iType'] == typeWork: 
                person = item['Person']
                price = d['charges'][(jobKey, taskKey, person)]
                personName = d['employees'][person]['PersonNAME']
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
    outdir = common.perioddir(d['period']) + '\\statements'
    #outfile = '%s\\statements\\%s.rtf'.format(common.perioddir(d['period']), jobKey)
    outfile = jobKey + '.rtf'
    out.save(outdir, outfile)
    
    # remember what we have produced for the invoice summaries
    if job['Weird'] or job['WIP']: 
        billed = 0.0
    else:
        billed = net
    invoice = { 'work': totalWork , 'expenses': totalExpenses, 'net': billed}
    #print jobKey
    #print d['auto_invoices']
    d['auto_invoices'][jobKey] = invoice


###########################################################################

def main(d):
    
    # create invoice items
    invItems = []        
    def AddItem(item, typeText, typeValue):
        x['iType'] = typeValue
        
        # FIXME - use common.assert_job instead of this:
        if not d['jobs'].has_key(x['JobCode']):
            fmt = "Found {0} with Jobcode '{1}', but no entry in jobs table"
            msg = fmt.format(typeText, x['JobCode'])
            raise common.DataIntegrityError(msg)
        
        invItems.append(x)
    for x in d['timeItems']: AddItem(x, "time item", typeWork)        
    for x in d['expenses']: AddItem(x, "expense", typeExpense)
    def invSorter(a, b): return cmp(a['JobCode'], b['JobCode']) or cmp(a['Task'], b['Task']) or (a['iType'] < b['iType'])        
    invItems.sort(invSorter)
        
    d['auto_invoices'] = {}
    for jobKey, jobGroup in aggregate(invItems, common.mkKeyFunc('JobCode')): 
        CreateJobStatment(jobKey, jobGroup, d)
    
if  __name__ == "__main__":
    print "Didn't do anything"
