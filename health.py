# The reports are usually stored in
# \\Cbegbrabssrv01\quality\Intertek\HSE Records\Monthly Returns

import common
import data
import period
    

def main():
    # FIXME - determine if selecting the period really has any effect, or just uses current.yml
    p = period.Period()
    p.inputPeriod()
    d = data.Data(p)
    d.load()
    
    total_all_hours = 0.0
    total_staff_hours = 0.0
    leave_and_sickness = 0.0
    for time_item in d.timeItems:
        time_val = time_item['TimeVal']
        total_all_hours += time_val
        
        employee_initials = time_item['Person']
        employee = d.employees[employee_initials]
        if employee['IsStaff']:
            total_staff_hours += time_val
            jobcode = time_item['JobCode']
            if jobcode == '010400' or jobcode == '010500': 
                leave_and_sickness += time_val
        
    output = 'Statistics\n'
    def line(text, value): return '%40s %8.f\n' % (text, value)
    output += line('Staff hours less leave and sickness', total_staff_hours - leave_and_sickness)
    output += line('Total hours', total_all_hours)
    print output
    dir = common.reportdir(d.p)
    filename = dir +'\\health.txt'    
    common.spit(filename, output)
    print 'A copy of the results are in ' + filename
    
    

if  __name__ == "__main__":
    main()
    print 'Finished'