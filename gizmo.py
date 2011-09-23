import pdb
import pickle
import pprint

import common
import db

def pic():
    data = db.fetch()
    db.save_state(data)

#pic()

#data = db.load_state()
#pic()
data = db.fetch()    
pp = pprint.PrettyPrinter(indent=4)
for e in sorted(data['employees']):
    print e
    times = filter(lambda x: x['Person'] == e, data['timeItems'])
    aggregates = common.aggregate(times, lambda x: (x['JobCode'], x['Task']))
    for key, vals in aggregates:
        total_times = common.summate(vals, common.mkKeyFunc('TimeVal'))
        print key, " ", total_times
    #pdb.set_trace()
    #pp.pprint(aggregates)
    print
#print data['timeItems']

print "Finished"
