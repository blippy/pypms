'''Create a manual invoice summary
'''

#import csv

import common
import data
import excel

###########################################################################

def import_manual_invoices(d):
    'Import invoices entered manually in spreadsheet'
    #fileName = 'M:\\Finance\\Invoices\\Inv summaries %s\\Inv Summary %s.xls' % (d.p.y, d.p.yyyymm())
    fileName = 'M:\\Finance\\camel\\%s\\camel-%s.xls' % (d.p.y, d.p.yyyymm())
    wsName = 'ManualInvoices'
    invoiceLines = excel.ImportWorksheet(fileName, wsName)
    
    def nth(row, index):
        try: result = row[index]
        except IndexError: result = ''
        return result
    
    manual_invoices = []
    for row in invoiceLines:
        irn, client, job, net, desc = [nth(row, idx) for idx in range(0,5)]
        if not d.jobs.has_key(job): continue
        net = common.AsFloat(net)        
        #if not manual_invoices.has_key(job): manual_invoices[job] = []
        manual_invoices.append({ 'irn' : irn, 'client' : client, 'job' : job, 'net' : net, 'desc' : desc})
    d.manual_invoices = manual_invoices
 
###########################################################################

def create_invoice_summary(d):    
    total = 0.0

    output = []
    output.append(['Ref', 'Client', 'Job', 'Net', 'Desc'])
    
    #def txt(input): return "'" + input
    def txt(input): return input
    def number(input): return '%.2f' % (input)

    # spit out the manual invoices
    for invoice in d.manual_invoices:
        #FIXME - ought to be possible to work out who the client is
        net = invoice['net']
        total += net
        output.append([txt(invoice['irn']), invoice['client'], txt(invoice['job']), 
            number(net), invoice['desc']])
        
    # write out the computed invoices
    job_codes = d.auto_invoices.keys()
    job_codes.sort()
    for job_code in job_codes:
        inv = d.auto_invoices[job_code]
        #FIXME - ought to be possible to work out who the client is
        net = inv['net']
        if net <> 0.0:
            total += net
            output.append(["", "", txt(job_code), number(inv['net'])])
        
    output.append([])
    output.append(['Total', '', '', number(total)])
    
    def excel_func(wb):        
        ws = wb.Worksheets('Sheet1')
        ws.Name = 'Invoices'
        row_num = 0
        for row in output:
            row_num += 1
            col_num = 0
            for col_value in row:
                col_num +=1
                ws.Cells(row_num, col_num).Value = col_value

    path = d.p.outDir() + '\\craig'
    common.makedirs(path)
    file_name = path + "\\invoices.xls"
    excel.create_workbook(file_name, excel_func)

###########################################################################

def create_reconciliation(d):
    print "FIXME NOW"

###########################################################################
def main(d):
    'Extract the manual invoices from the spreadsheets'
    import_manual_invoices(d)
    create_invoice_summary(d)
    create_reconciliation(d)
        
if  __name__ == "__main__":
    common.run_current(main)
    print 'Finished'
