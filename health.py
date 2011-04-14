# The reports are usually stored in
# \\Cbegbrabssrv01\quality\Intertek\HSE Records\Monthly Returns

import pdb

import common
from common import summate, mkKeyFunc, princ, print_timing
import period
    
###########################################################################


@print_timing
def create_health_report(d):
    stats = {}
    
    total_all_hours = 0.0
    total_staff_hours = 0.0
    total_leave_and_sickness = 0.0
    time_items = d['timeItems']
    las_test = lambda(x): employee['IsStaff'] and (x['JobCode'] == '010400' or x['JobCode'] == '010500')
    #fmt = '{0:4} {1:15} {2:>10} {3:>10} {4:>10}'
    
    output = []
    def add_line(text):
        output.append(text)
    def add_formatted_line(*fields):
        text = []
        for fmt, value in zip('{0:4} {0:15} {0:>10} {0:>10} {0:>10}'.split(' '), fields):
            #pdb.set_trace()
            #print fmt, "---", value
            text.append(fmt.format(value))            
        text = ' '.join(text)
        add_line(text)
    #dir = period.reportdir(d['period'])
    #filename = dir +'\\health.txt' 
    #output = file(filename, "w")

    add_formatted_line('INI', 'NAME', 'LEAV/SICK', 'STAFF', 'ALL')
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
        
        add_formatted_line(employee_initials, full_name, leave_and_sickness, staff_hours, all_hours)

    add_formatted_line('', 'TOTALS' , total_leave_and_sickness, total_staff_hours, total_all_hours)

    add_line('\nHealth & Safety Stats:')
    add_line('Staff hours less leave & sickness: {0}'.format(total_staff_hours - total_leave_and_sickness))
    add_line('Total all hours: {0}'.format(total_all_hours))

    #output.close()
    
    period.save_report('health.txt', output)
 

    

###########################################################################

if  __name__ == "__main__":
    princ("Didn't do anything")