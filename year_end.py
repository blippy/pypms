# year-end sort-out

# NB You cannot close jobs with WIP, UBI, BIA. This is not picked up currently.
# It would also be useful to have a last invoicing date

from pprint import pprint
from sets import Set

import db

actives = Set()
jobs = db.GetJobs()
#for j in jobs[0:3]: pass # pprint(type(j))
#print jobs.keys()

tasks = db.GetTasks()
for idx, t in enumerate(tasks.items()):
    _, d = t
    jcode = d['JobCode']

    #if idx <3: 
    #    pprint(d)
    if d['JobActive'] == 1:
        if jcode[0] == '0': continue
        if jcode == u'3. Sundry': continue
        if jcode[0:2] == "40": continue
        if jcode[0:2] == "75": continue
        actives.add(jcode)


alist = list(actives)
alist.sort()

print "Number of active jobs:", len(alist)
print

for jcode in alist:
    print jcode
    job = jobs[jcode]
    print job['title']
    print job['ProjectManager']
    print "Last billed:"
    print
    
#pprint(alist)
