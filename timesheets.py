# Create timesheet files

import datetime, pdb
#from itertools import groupby

import common, rtf
from common import aggregate

###########################################################################


def AddHeader(title, job_code, job_title, task_title, person_name, out):
    def descline(name, value): out.add('{0:14s}{1:s}'.format(name, value))
    out.addTitle(title)
    descline('Job number:', job_code)
    descline('Description:', job_title)
    descline('Task:', task_title)
    descline('Person:', person_name)
    out.para()

def AddPersonToJobsheet(personKey, personGroup, out):    
    out.add('{0:10s} {1:>9s}   {2:s}'.format('Date', 'Qty', 'Description') , 2)
    totalTime = 0
    for timeItem in personGroup:
        timeVal = timeItem['TimeVal']
        totalTime += timeVal
        dateVal = timeItem['DateVal']
        asDate = datetime.date(int(dateVal[0:4]), int(dateVal[5:7]), int(dateVal[8:]))
        day = asDate.strftime('%a %d %b')
        workDone = timeItem['WorkDone']
        if workDone == '..': workDone = ''
        out.add('%s %9.3f   %s' % (day, timeVal, workDone))
    out.para()
    out.add('{0:10s} {1:9.3f}'.format('Total', totalTime))



###########################################################################

def CreateJobsheet(jobcode, job_times, d, title, outdir):
    out = rtf.Rtf()
    common.assert_job(d, jobcode, "In tblTimeItems, but not jobs")
    job = d['jobs'][jobcode]         
    sheets = aggregate(job_times, lambda x: (x['Task'], x['Person']))
    last_key = sheets[-1][0]
    for key, values in sheets:
        task, initials = key
        taskTitle = task + ' - ' + d['tasks'][(jobcode, task)]['TaskDes']
        person_name = d['employees'][initials]['PersonNAME']
        AddHeader(title, jobcode, job['title'], taskTitle, person_name, out)
        #pdb.set_trace()
        AddPersonToJobsheet(initials, values, out)
        out.annotation(job, job['TsApprover'])
        if key != last_key: out.page()
    out.save(outdir, jobcode + ".rtf" )

###########################################################################

def main(d):
    p = d['period']
    title = 'Timesheet: ' + p.mmmmyyyy()
    outdir = common.perioddir(p) + '\\timesheets'
    for jobKey, job_items in aggregate(d['timeItems'], common.mkKeyFunc('JobCode')):
        #pdb.set_trace()
        CreateJobsheet(jobKey, job_items, d, title, outdir)

if  __name__ == "__main__":
    print "Didn't do anything"
