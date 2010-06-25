'''Create a WIP summary
'''

import common
import db
import excel

###########################################################################
def main(d):
    
    fieldspec = ['InvJobCode', 'SumOfInvUBI', 'SumOfInvWIP', 'TotalWIP']
    wips = db.GetQryWip(fieldspec)
    
    output = [['Job', 'UBI', 'WIP', 'TOTAL']]
    ubi_total = 0.0
    wip_total = 0.0
    total = 0.0
    for wip_entry in wips:
        job_code, ubi, wip, job_total = wip_entry
        ubi = float(ubi)
        wip = float(wip)
        job_total = float(job_total)
        if abs(ubi) <0.01 and abs(wip) < 0.01: continue
        ubi_total += ubi
        wip_total += wip
        total += job_total
        output.append([job_code, ubi, wip, job_total])

    output.append([])
    output.append(['TOTAL', ubi_total, wip_total, total])
    excel.create_report(d.p, "Wip", output, [2,3,4])

        
if  __name__ == "__main__":
    common.run_current(main)
    print 'Finished'    