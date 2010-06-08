import datetime, pdb

from itertools import groupby #, ifilter

import common, data

###########################################################################


###########################################################################

def CreateJobsheet(jobcode, jobTimes, periodText, d):

    body = []
    job = d.jobs[jobcode]
    title = job['title']
    approver = job['TsApprover']
    #if job['job'] == '2669': pdb.set_trace()
    if approver == 'None': approver = ''
    
    def descline(name, value): return '{0:14s}{1:s} \\par \n'.format(name, value)
     
    for taskKey, taskGroup in groupby(jobTimes, common.mkKeyFunc('Task')):
        taskTitle = taskKey + ' - ' + d.tasks[(jobcode, taskKey)]['TaskDes']
        for personKey, personGroup in groupby( taskGroup, common.mkKeyFunc('Person')):
            #pdb.set_trace()
            text = '\\f0\\fs30 Timesheet: %s \\par \\f0\\fs18 ' % (periodText)
            text+= descline('Job number:', jobcode)
            text+= descline('Description:',	title)
            text+= descline('Task:', taskTitle)
            personName = d.employees[personKey]
            text+= descline('Person:', personName)
            text+= '\\par \\par \n'
            
            text += '{0:10s} {1:>9s}   {2:s} \\par \\par \n'.format('Date', 'Qty', 'Description')
            totalTime = 0
            for timeItem in personGroup:
                timeVal = timeItem['TimeVal']
                totalTime += timeVal
                dateVal = timeItem['DateVal']
                asDate = datetime.date(int(dateVal[0:4]), int(dateVal[5:7]), int(dateVal[8:]))
                day = asDate.strftime('%a %d %b')
                workDone = timeItem['WorkDone']
                if workDone == '..': workDone = ''
                text += '%s %9.3f   %s \\par \n' % (day, timeVal, workDone)
                
            text += '\n \\par \n{0:10s} {1:9.3f} \\par \n'.format('Total', totalTime)
            text += '\n \\par \\par \\par \nApproved: ______________________________________ \\par {0:s}'.format(approver)
            text += common.annotation(job)
            body.append(text)
                
    return '\n\\page\n'.join(body)


def main(d):
    root = d.p.outDir() + '\\timesheets'
    common.makedirs(root)
    periodText = d.p.mmmmyyyy()

    for jobKey, jobGroup in groupby(d.timeItems, common.mkKeyFunc('JobCode')):
        text = '{\\rtf {\\fonttbl {\\f0 Consolas;}}'
        text += CreateJobsheet(jobKey, jobGroup, periodText, d)
        text += '}'
        fullName = root + "\\" + jobKey + ".rtf"
        with open(fullName, "w") as f: f.write(text)

if  __name__ == "__main__":
    d = data.Data()
    d.restore()
    main(d)
    print 'Finished'
