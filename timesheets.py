# Create timesheet files

import datetime, pdb
from itertools import groupby

import common, data, rtf

###########################################################################

def CreateJobsheet(jobcode, jobTimes, d, title, outdir):

    out = rtf.Rtf()

    job = d.jobs[jobcode]
    jobTitle = job['title']
    approver = job['TsApprover']
    if approver == 'None': approver = ''
    
    def descline(name, value): out.add('{0:14s}{1:s}'.format(name, value))
         
    for taskKey, taskGroup in groupby(jobTimes, common.mkKeyFunc('Task')):
        taskTitle = taskKey + ' - ' + d.tasks[(jobcode, taskKey)]['TaskDes']
        for personKey, personGroup in groupby( taskGroup, common.mkKeyFunc('Person')):
            out.addTitle(title)
            descline('Job number:', jobcode)
            descline('Description:',	jobTitle)
            descline('Task:', taskTitle)
            personName = d.employees[personKey]['PersonNAME']
            descline('Person:', personName)
            out.para()
            
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
            out.annotation(job, approver)
            out.page()

    out.save(outdir, jobcode + ".rtf" )


def main(d):
    title = 'Timesheet: ' + d.p.mmmmyyyy()
    outdir = d.p.outDir() + '\\timesheets'
    for jobKey, jobGroup in groupby(d.timeItems, common.mkKeyFunc('JobCode')):
        CreateJobsheet(jobKey, jobGroup, d, title, outdir)

if  __name__ == "__main__":
    d = data.Data()
    d.restore()
    main(d)
    print 'Finished'
