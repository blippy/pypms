'''Create a manual invoice summary
'''

import functools
import itertools
import pdb

import common
import db
import excel
from common import dget, princ, print_timing, tri
import period


  
VAT = 0.20


def number(input): return '%.2f' % (input)    

###########################################################################

def load(data):
    'Import invoices entered manually in spreadsheet'
    fieldspec = [
        (1, 'irn', excel.fix_str, ''),
        (2, 'client', str, ''),
        (3, 'JobCode', excel.fix_str, ''),
        (4, 'net', float, None),
        (7, 'desc', str, '')]
    return excel.import_summary_sheet(data, 'ManualInvoices', fieldspec, 1)

@print_timing
def XXX_import_manual_invoices(data):
    'Import invoices entered manually in spreadsheet'

    invoiceLines = excel.read_worksheet('ManualInvoices')
    manual_invoices = []
    # TODO abstract similarities to expenses.import_expenses()
    fieldspec = [(1, 'irn', excel.fix_str), (2, 'client', str), (3, 'JobCode', excel.fix_str),(4, 'net', float), (7, 'desc', str)]
    row_num = 0
    for row in invoiceLines[1:]:
        row_num += 1
        
        record = {}
        for colNum, fieldName, fieldType in fieldspec:
            text = row[colNum-1]
            try: record[fieldName] = fieldType(text)
            except ValueError: pass
        job_code = record['JobCode']
        if job_code == '' or not record.has_key('net'): continue
        
        #irn, client, job, net = [nth(row, idx) for idx in range(0,4)]
        #desc =nth(row, 6)
        #if job =='Job' or job == '' or job == "TOTAL:": continue
        if not data['jobs'].has_key(job_code):
            fmt = "ERR: No database job '{0}' camel workbook '{1}' sheet manualInvoices row {2}"
            msg = fmt.format(job_code, excel.camelxls(), row_num)
            raise common.DataIntegrityError(msg)
        #net = common.AsFloat(net)        
        #manual_invoices.append({ 'irn' : irn, 'client' : client, 'job' : job, 'net' : net, 'desc' : desc})
        manual_invoices.append(record)
    data['manual_invoices'] = manual_invoices
    
    
def enumerate_invoices(data):
    mans = data['manual_invoices']

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
        #print inv.keys()
        job = jobs[inv['JobCode']]
        vatable = job['vatable']
        inv['vat_rate'] = tri(vatable, VAT, 0.0)
    
    net = common.summate(the_invoices, lambda x: x['net'])
    return the_invoices

###########################################################################
@print_timing
def create_text_invoice_summary(the_invoices):
    output = [['Ref', 'Client', 'Job', 'Net', 'VAT', 'Gross', 'Desc']]
    
    net_total = 0.0
    vat_total = 0.0
    #gross_tot
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
        #desc += ' '*12
        if len(desc) > 12: desc = desc[0:12]
        #print inv.keys()
        line = [irn, client, job, number(net), number(vat), number(gross), desc]        
        output.append(line)
        
    output.append(['total', '', '', number(net_total), number(vat_total), number(net_total + vat_total)])
    output.append([])
    period.create_text_report("invoices.txt", output)
    
###########################################################################
@print_timing
def create_excel_invoice_summary(invoices):
    fname = period.reportfile("invoices.xls")
    xl = excel.Book(fname)
    ws = xl.wb.Worksheets('Sheet1')
    ws.Name = 'Invoices'    
    
    def numbered(alist): return itertools.izip(range(0, len(alist)), alist)        
    def txt(text): return common.AsAscii(text)
    
    


    for coln, coltitle in numbered(['Ref', 'Client', 'Job', 'Net', 'VAT', 'Gross', 'Desc']):
        ws.Cells(1, coln+1).Value = coltitle
        
    rown = 2

            
    for invoice in invoices:
        rown += 1
        for c in [4, 5, 6]: ws.Cells(rown, c).NumberFormat = "0.00"
        for c, f in zip([1, 2, 3, 7], ['irn', 'client', 'job', 'desc']):
            ws.Cells(rown, c).Value = txt(dget(invoice, f, ''))            
        ws.Cells(rown, 4).Value = number(dget(invoice, 'net'))
        formula = '=round(RC[-1]*%f, 2)' % (invoice['vat_rate'])
        ws.Cells(rown, 5).Formula = formula        
        ws.Cells(rown, 6).Formula = '=RC[-2]+RC[-1]' # Gross
    
    # write the totals
    rown += 2
    ws.Cells(rown, 1).Value = 'Total'
    formula = '=SUM(R[-%d]C:R[-1]C)' % (rown -2)
    for col in [4,5,6]:
        ws.Cells(rown, col).Formula = formula
        ws.Cells(rown, col).NumberFormat = "0.00"





    
    
###########################################################################

        
if  __name__ == "__main__":
    import_manual_invoices({})
    princ("Finished")
