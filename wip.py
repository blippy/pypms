'''Create a WIP summary
'''

import pdb

import common
from common import aggregate, AsAscii, AsFloat, summate, princ, print_timing
import db
import excel
import period



###########################################################################

def create_wip_line(wip_item, invoices):
    job_code, ubi_ytd, wip_ytd = common.mapdict(wip_item, ['InvJobCode', 'SumOfInvUBI', 'SumOfInvWIP'])
    if invoices.has_key(job_code): expenses, work = common.mapdict(invoices[job_code], ['expenses', 'work'])
    else: expenses, work = 0.0, 0.0
    line = [job_code, ubi_ytd + wip_ytd, wip_ytd, ubi_ytd, '', expenses + work, expenses, work]
    return line


###########################################################################

@print_timing
def create_wip_lines(data):
    sql =  "SELECT InvJobCode, Sum(tblInvoice.InvUBI) AS SumOfInvUBI, "
    sql += "Sum(tblInvoice.InvWIP) AS SumOfInvWIP FROM tblInvoice GROUP BY tblInvoice.InvJobCode "
    sql += "ORDER BY tblInvoice.InvJobCode;"
    fields = 'InvJobCode,SumOfInvUBI,SumOfInvWIP'
    wips = db.fetch_and_dictify(sql, fields)
    wips = filter(lambda x: abs(x['SumOfInvWIP']) + abs(x['SumOfInvUBI']) >= 0.01, wips)

    output = []
    output.append(['Job', 'TOTAL', 'WIP', 'UBI',  '      ',  'TOTAL', 'EXPS', 'WORK'])
    output.append(['',    'YTD',   'YTD', 'YTD',  '',  'CUR',   'CUR',  'CUR'])
    
    # output the totals for each job code
    for wip in wips:
        line = create_wip_line(wip, data['auto_invoices'])
        output.append(line)
    total = common.summate_cols(output)
    total[0] = "TOTAL"
    total[4] = ''
    output.append(total)

    return output
    

###########################################################################
@print_timing
def create_wip_report(data, output_text = True):
    output = create_wip_lines(data)
    if output_text:
        period.create_text_report("wip.txt", output)
    else:
        # TODO Consider zapping Excel output option if considered unecessary
        excel.create_report("Wip", output, [2, 3 ,4, 5, 6, 7])

        
###########################################################################

if  __name__ == "__main__":
    data = db.load_state()
    create_wip_report(data)
    princ("Finish")