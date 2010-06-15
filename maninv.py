'''The processing of manual invoices
'''

import common
import data
import excel

###########################################################################

def import_manual_invoices(d):
    'Import invoices entered manually in spreadsheet'
    fileName = 'M:\\Finance\\Invoices\\Inv summaries %s\\Inv Summary %s.xls' % (d.p.y, d.p.yyyymm())
    wsName = 'Invoices'
    invoiceLines = excel.ImportWorksheet(fileName, wsName)
    manual_invoices = {}
    for row in invoiceLines:
        job = row[2] #, net, vat, total = row[2:5]
        if not d.jobs.has_key(job): continue
        id = row[0]
        client = row[1]
        net, vat , total = map(common.AsFloat, row[3:6])
        try: desc = row[6]
        except IndexError: desc = ''
        if not manual_invoices.has_key(job): manual_invoices[job] = []
        manual_invoices[job].append({ 'id' : id, 'client' : client, 'job' : job, 'net' : net, 'desc' : desc})
    d.manual_invoices = manual_invoices
    

###########################################################################
def main(d = None):
    'Extract the manual invoices from the spreadsheets'
    if not d: d = data.Data(restore = True)
    import_manual_invoices(d)
    
if  __name__ == "__main__":
    main()
    print 'Finished'
