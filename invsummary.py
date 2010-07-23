'''Create a manual invoice summary
'''

import functools
import itertools
import pdb

import common
import data
import db
import excel


    
###########################################################################

def import_manual_invoices(d):
    'Import invoices entered manually in spreadsheet'

    #pdb.set_trace()
    invoiceLines = excel.ImportWorksheet(common.camelxls(d['period']), 'ManualInvoices')
    
    def nth(row, index):
        try: result = row[index]
        except IndexError: result = ''
        return result
    
    manual_invoices = []
    for row in invoiceLines:
        irn, client, job, net, desc = [nth(row, idx) for idx in range(0,5)]
        if not d['jobs'].has_key(job): continue
        net = common.AsFloat(net)        
        #if not manual_invoices.has_key(job): manual_invoices[job] = []
        manual_invoices.append({ 'irn' : irn, 'client' : client, 'job' : job, 'net' : net, 'desc' : desc})
    d['manual_invoices'] = manual_invoices
    
def accumulate(d):
    # FIXME - this can probably be used in many other places
    result = {}
    for el in d['manual_invoices']:
        common.dplus(result, el['job'], el['net'])
    return result
    
 
###########################################################################

def create_invoice_summary_func(d, wb):
    def numbered(alist): return itertools.izip(range(0, len(alist)), alist)        
    def txt(text): return common.AsAscii(text)
    def number(input): return '%.2f' % (input)
    
    ws = wb.Worksheets('Sheet1')
    ws.Name = 'Invoices'    

    for coln, coltitle in numbered(['Ref', 'Client', 'Job', 'Net', 'VAT', 'Gross', 'Desc']):
        ws.Cells(1, coln+1).Value = coltitle
        
    total = 0.0
    rown = 2

    def addline(rown, irn, client, job_code, net, desc):
        ws.Cells(rown, 1).Value = txt(irn)
        ws.Cells(rown, 2).Value = client
        ws.Cells(rown, 3).Value = txt(job_code)
        ws.Cells(rown, 4).Value = number(net)
        ws.Cells(rown, 4).NumberFormat = "0.00"
        
        vatable = d['jobs'][job_code]['vatable']
        if vatable: rate = 0.175
        else: rate = 0.0
        formula = '=round(RC[-1]*%f, 2)' % (rate)
        ws.Cells(rown, 5).Formula = formula
        
        ws.Cells(rown, 5).NumberFormat = "0.00"
        ws.Cells(rown, 6).Formula = '=RC[-2]+RC[-1]' # Gross
        ws.Cells(rown, 6).NumberFormat = "0.00"
        ws.Cells(rown, 7).Value = desc
        
        
        
    # spit out the manual invoices
    for invoice in d['manual_invoices']:
        rown += 1
        #FIXME - ought to be possible to work out who the client is            
        net = invoice['net']
        total += net
        addline(rown, invoice['irn'], invoice['client'], invoice['job'], net, invoice['desc'])
    
    # write out the computed invoices
    #pdb.set_trace()
    job_codes = d['auto_invoices'].keys()
    job_codes.sort()
    for job_code in job_codes:
        inv = d['auto_invoices'][job_code]
        #FIXME - ought to be possible to work out who the client is
        net = inv['net']
        if net <> 0.0:
            rown += 1
            total += net
            addline(rown, "", "", job_code, inv['net'], '')
    
    # write the totals
    rown += 2
    ws.Cells(rown, 1).Value = 'Total'
    formula = '=SUM(R[-%d]C:R[-1]C)' % (rown -2)
    for col in [4,5,6]:
        ws.Cells(rown, col).Formula = formula
        ws.Cells(rown, col).NumberFormat = "0.00"


def create_invoice_summary(d): 
    fname = common.reportfile(d['period'], "invoices.xls")
    func = functools.partial(create_invoice_summary_func, d)
    excel.create_workbook(fname, func)

###########################################################################

def create_reconciliation(d):
    
    invoices = {}
    
    # Obtain the invoice amounts from the database
    recs = db.GetInvoices(d, ['InvJobCode', 'InvInvoice'])
    for rec in recs:
        job_code = str(rec[0])
        invoices[job_code] = { 'db' : rec[1] , 'excel' : 0.0}
        
    # How much are we billing out?
    for invoice in d['manual_invoices']:
        net = invoice['net']
        job_code = invoice['job']
        try:
            invoices[job_code]['excel'] += net
        except KeyError:
            msg = "ERR101: No job code " + job_code
            raise common.DataIntegrityError(msg)
    job_codes = d['auto_invoices'].keys()
    for job_code in job_codes:
        inv = d['auto_invoices'][job_code]
        net = inv['net']
        invoices[job_code]['excel'] += net
            
    # now compare the invoices side-by-side
    output_text = 'RECONCILIATION BETWEEN INVOICE SUMMARY AND PMS\n'
    output_text += '%7s %10s %10s %10s\n' % ('JOB', 'DBASE', 'EXCEL', 'DIFF')
    
    def write_line(code, v1, v2):
        diff = float(v1) - float(v2)
        if abs(diff) > 2:
            warning_flag = '***'
        else:
            warning_flag = ''
        return '%7s %10.2f %10.2f %10.2f %s\n' % (code, v1, v2, diff, warning_flag)
        
    job_codes = invoices.keys()
    job_codes.sort()
    db_total = 0.0
    excel_total = 0.0
    for job_code in job_codes:
        db_value = float(invoices[job_code]['db'])
        db_total += db_value
        excel_value = float(invoices[job_code]['excel'])
        excel_total += excel_value
        output_text += write_line(job_code, db_value, excel_value)
        #pdb.set_trace()
        
    output_text += '\n\n'
    output_text += write_line('TOTAL', db_total, excel_total)
    common.spit(common.reportdir(d['period']) + "\\monthrec.txt", output_text) # FIXME maybe it would be better if common.reportdir() just took in d

###########################################################################
def main(d):
    'Extract the manual invoices from the spreadsheets'
    import_manual_invoices(d)
    create_invoice_summary(d)
    create_reconciliation(d)
        
if  __name__ == "__main__":
    data.run_current(main)
    print 'Finished'
