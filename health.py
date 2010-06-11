import data
import period
    

def main():
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
        
    print 'Statistics'
    def line(text, value): return '%40s %8.f' % (text, value)
    print line('Staff hours less leave and sickness', total_staff_hours - leave_and_sickness)
    print line('Total hours', total_all_hours)
    

if  __name__ == "__main__":
    main()
    print 'Finished'