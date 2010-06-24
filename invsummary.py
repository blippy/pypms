'''Create a manual invoice summary
'''

import pdb

import common
import data
import db
import excel


    
###########################################################################

def import_manual_invoices(d):
    'Import invoices entered manually in spreadsheet'
    #fileName = 'M:\\Finance\\Invoices\\Inv summaries %s\\Inv Summary %s.xls' % (d.p.y, d.p.yyyymm())
    fileName = 'M:\\Finance\\camel\\%s\\camel-%s.xls' % (d.p.y, d.p.yyyymm())
    wsName = 'ManualInvoices'
    #import_invoices_from_excel(d, fileName, wsName)
    
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
    
    excel.create_report(d.p, "Invoices", output, 2)

###########################################################################

def create_reconciliation(d):
    
    invoices = {}
    
    # Obtain the invoice amounts from the database
    recs = db.GetInvoices(d, ['InvJobCode', 'InvInvoice'])
    for rec in recs:
        job_code = str(rec[0])
        invoices[job_code] = { 'db' : rec[1] , 'excel' : 0.0}
        
    # How much are we billing out?
    for invoice in d.manual_invoices:
        net = invoice['net']
        job_code = invoice['job']
        try:
            invoices[job_code]['excel'] += net
        except KeyError:
            msg = "ERR101: No job code " + job_code
            raise common.DataIntegrityError(msg)
    job_codes = d.auto_invoices.keys()
    for job_code in job_codes:
        inv = d.auto_invoices[job_code]
        net = inv['net']
        invoices[job_code]['excel'] += net
            
    # now compare the invoices side-by-side
    output_text = 'RECONCILIATION BETWEEN INVOICE SUMMARY AND PMS\n'
    output_text += '%7s %10s %10s %10s\n' % ('JOB', 'DBASE', 'EXCEL', 'DIFF')
    
    def write_line(code, v1, v2):
        diff = float(db_value) - float(excel_value)
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
    common.spit(common.reportdir(d.p) + "\\monthrec.txt", output_text)

###########################################################################
def main(d):
    'Extract the manual invoices from the spreadsheets'
    import_manual_invoices(d)
    create_invoice_summary(d)
    create_reconciliation(d)
        
if  __name__ == "__main__":
    common.run_current(main)
    print 'Finished'
