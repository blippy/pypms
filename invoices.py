'''Create a manual invoice summary
'''

import functools
import itertools
import pdb

import common
import db
import excel
from common import dget, princ, tri
import period


  
VAT = 0.20


def number(input): return '%.2f' % (input)    

###########################################################################    
    
def enumerate_invoices(data):
    mans = data['ManualInvoices']

    job_codes = data['auto_invoices'].keys()
    job_codes.sort()
    autos = []
    for job in job_codes:
        inv = data['auto_invoices'][job]
        inv['JobCode'] = job
        inv['client'] = ''
        inv['irn'] = ''
        inv['desc'] = ''
        autos.append(inv)
    autos = filter(lambda x: x['net'] <> 0.0,  autos)

    the_invoices = mans + autos
    
    # augment vat rates
    jobs = data['jobs']
    for inv in the_invoices:
        job = jobs[inv['JobCode']]
        vatable = job['vatable']
        inv['vat_rate'] = tri(vatable, VAT, 0.0)
    
    net = common.summate(the_invoices, lambda x: x['net'])
    return the_invoices

###########################################################################

def create_text_invoice_summary(the_invoices):
    output = [['Ref', 'Client', 'Job', 'Net', 'VAT', 'Gross', 'Desc']]
    
    net_total = 0.0
    vat_total = 0.0
    for inv in the_invoices:
        irn = inv['irn']
        client = inv['client']
        job = inv['JobCode']
        net = inv['net']
        net_total += net
        vat = inv['vat_rate']* net
        vat_total += vat
        gross = net + vat
        desc = inv['desc']
        desc.strip()
        if len(desc) > 12: desc = desc[0:12]
        line = [irn, client, job, number(net), number(vat), number(gross), desc]        
        output.append(line)
        
    output.append(['total', '', '', number(net_total), number(vat_total), number(net_total + vat_total)])
    output.append([])
    period.create_text_report("invoices.txt", output)
    

    
###########################################################################

        
if  __name__ == "__main__":
    import_manual_invoices({})
    princ("Finished")
