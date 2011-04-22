from common import print_timing
import db
import period

###########################################################################

@print_timing
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
    period.save_report("monthrec.txt", output_text)
    
if __name__ == '__main__':
    print("Didn't do anything")