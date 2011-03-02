# The reports are usually stored in
# \\Cbegbrabssrv01\quality\Intertek\HSE Records\Monthly Returns

import common
from common import summate, mkKeyFunc, princ
import period
    
###########################################################################

class S:
    def __init__(self):
        pass

def main(d):
    stats = {}
    
    total_all_hours = 0.0
    total_staff_hours = 0.0
    total_leave_and_sickness = 0.0
    time_items = d['timeItems']
    las_test = lambda(x): employee['IsStaff'] and (x['JobCode'] == '010400' or x['JobCode'] == '010500')
    fmt = '{0:4} {1:15} {2:>10} {3:>10} {4:>10}'
    
    dir = common.reportdir(d['period'])
    filename = dir +'\\health.txt' 
    output = file(filename, "w")

    print >>output, fmt.format('INI', 'NAME', 'LEAV/SICK', 'STAFF', 'ALL')
    for employee_initials, worker_times in common.aggregate(time_items, common.mkKeyFunc('Person')):
        try: employee = d['employees'][employee_initials]
        except KeyError: continue
    
        full_name = employee['PersonNAME']
        time_val = mkKeyFunc('TimeVal')

        staff_hours = summate(worker_times, time_val, lambda(x): employee['IsStaff'])
        total_staff_hours += staff_hours
                
        leave_and_sickness = summate(worker_times, time_val, las_test)
        total_leave_and_sickness += leave_and_sickness
        
        all_hours = summate(worker_times, time_val)
        total_all_hours += all_hours
        
        print >>output, fmt.format(employee_initials, full_name, leave_and_sickness, staff_hours, all_hours)

    print >>output, fmt.format('', 'TOTALS' , total_leave_and_sickness, total_staff_hours, total_all_hours)

    print >>output, '\nHealth & Safety Stats:'
    print >>output, 'Staff hours less leave & sickness:', total_staff_hours - total_leave_and_sickness
    print >>output, 'Total all hours:', total_all_hours

    output.close()
 

    

###########################################################################

if  __name__ == "__main__":
    princ("Didn't do anything")