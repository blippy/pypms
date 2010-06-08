# Create the monthly workstatements for each job

#import itertools
import pdb
from itertools import groupby, ifilter
from operator import itemgetter, attrgetter

import common, data


###########################################################################

typeWork, typeExpense = range(2) # create enumerated invoice type



###########################################################################

class Output:
    
    def __init__(self, p):
        self.outdir = p.outDir() + '\\statements'
        periodText = p.mmmmyyyy()
        self.text = '{\\rtf {\\fonttbl {\\f0 Consolas;}}\\f0\\fs30 Work Statement: %s \\par \\f0\\fs18 ' % (periodText)
        
    def add(self, text, newLines = 1):
        if not text: return
        self.text += text
        self.lines(newLines)
        
    def lines(self, numLines = 1):
        'Add new lines to output'
        self.text += '\\par \n' * numLines

    def heading(self):
        self.add( '\\par \n{0:44s}{1:>9s}{2:>9s}{3:>9s}'.format('', 'Quantity', 'Price', 'Value'),  2)

    def InvLine(self, desc , qty, price, value):
        text = '   {0:41.41s}{1:9.2f}{2:9.2f}{3:9.2f}'.format(desc , qty, price, value)
        self.add(text)
        
    def Subtotal(self, desc, value):
        'Output a subtotal'
        text = '   {0:41.41s}{1:9.2s}{2:9.2s}{3:9.2f}'.format(desc , '', '', value)
        self.add(text)

        
    def save(self, fname):
        common.makedirs(self.outdir)
        fullName = "%s\\%s" % (self.outdir, fname)
        with open(fullName, "w") as f: f.write(self.text + '}')

###########################################################################

def OutputTitle(out, job):
    address = job['address'] # whole address
    address = address.split('\r')
    address1 = address[0] # first line of address
    out.add(address1)    
    out.add(job['title'])
    out.add("SREL job code: " + job['job'], 2)
    
    references = job['references']
    references = references.replace('\n', '\\par \n')
    out.add(references)
    
    out.heading()


 


class Section:
    def __init__(self):
        self.heading = 'Fill with something sensible'
        self.workEntries = {}
        self.expenseEntries = []
        
    def AddWork(self, item, price, personName):
        if not self.workEntries.has_key(personName): self.workEntries[personName] = { 'desc' : personName , 'qty' : 0, 'price' : price }
        #print item
        self.workEntries[personName]['qty'] += item['TimeVal']
        
    def Work(self):
        items = self.workEntries.values()
        items.sort(key=lambda x: x['desc'])
        return items
    
    def AddExpense(self, item, exp_factor):
        desc = ' - '.join([item['Period'], item['Name'], item['Desc']])
        self.expenseEntries.append( { 'desc' : desc , 'qty' : exp_factor, 'price' : item['Amount'] })
        

    
    
def ProcessSubsection(out, items, subtotalTitle):
    if len(items) == 0: return 0, 0 # don't do anything if there are no items
    total = 0
    for el in items:
        desc, qty, price = el['desc'], el['qty'], el['price']
        value = round(qty * price, 2)
        total += value
        out.InvLine(desc, qty, price, value)
       
    numItems = len(items)
    if numItems >1: out.Subtotal(subtotalTitle, total)
    out.lines()
    return total, numItems
    
def CreateJobStatment(jobKey, invItems, d):
    if jobKey[0:2] == "01": return
    job = d.jobs[jobKey]
    #if job['Weird']: return # we should even create weird invoices - we their time and expenses anyway
    out = Output(d.p)
    OutputTitle(out, job)
    
    invItems = list(invItems)
      
    exp_factor = job['exp_factor']
    
    # distribute invoice items into sections
    sections = [] # a list of Section classes
    for taskKey, taskGroup in groupby(invItems, common.mkKeyFunc('Task')):
        s = Section()
        if taskKey == '' : 
            s.heading = 'Expenses not categorised to a specific task'
            s.ordering = 'ZZZ'
        else: 
            s.heading = 'Task %s: %s' % (taskKey, d.tasks[(jobKey, taskKey)]['TaskDes'])
            s.ordering = s.heading
            
        for item in taskGroup: 
            # FIXME - I think we aggregated work and expenses together - now we're splitting them out - which doesn't make sense - we should have kept them spearate in the first place
            if item['iType'] == typeWork: 
                person = item['Person']
                price = d.charges[(jobKey, taskKey, person)]
                personName = d.employees[person]
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
            out.Subtotal('Task subtotal', work+expenses)
            out.lines()
        
    # output grand summary
    out.add('Overall summary')
    out.Subtotal('Work total', totalWork)
    out.Subtotal('Expenses total', totalExpenses)
    net = totalWork+totalExpenses
    out.Subtotal('Net total', net)
    #FIXME - I don't think the following 3 lines do anything
    job['work'] = work
    job['expenses'] = expenses
    job['net'] = net
        
    out.add(common.annotation(job))
    out.save(jobKey + '.rtf')
    
    # remember what we have produced for the invoice summaries
    if job['Weird']: # all bets are off regards what the values should be
        invoice = { 'work': 0 , 'expenses': 0, 'net': 0}
    else:
        invoice = { 'work': totalWork , 'expenses': totalExpenses, 'net': net}
    d.invoices[jobKey] = invoice


###########################################################################

def main(d):
    
    # create invoice items
    invItems = []        
    def AddItem(item, typeText, typeValue):
        x['iType'] = typeValue
        if not d.jobs.has_key(x['JobCode']):
            msg = 'Found %s with Jobcode %s, but no entry in jobs table' % (typeText, x['JobCode'])
            raise common.DataIntegrityError(msg)
        invItems.append(x)
    for x in d.timeItems: AddItem(x, "time item", typeWork)        
    for x in d.expenses: AddItem(x, "expense", typeExpense)
    def invSorter(a, b): return cmp(a['JobCode'], b['JobCode']) or cmp(a['Task'], b['Task']) or (a['iType'] < b['iType'])        
    invItems.sort(invSorter)
        
    for jobKey, jobGroup in groupby(invItems, common.mkKeyFunc('JobCode')): CreateJobStatment(jobKey, jobGroup, d)
    
if  __name__ == "__main__":
    d = data.Data()
    d.restore()
    main(d)
    d.store()
    print 'Finished'
