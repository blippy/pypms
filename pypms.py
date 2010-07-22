import code, cProfile, datetime, msvcrt, os, pdb, pprint, pstats, sys, traceback

import common
import data
import db
import period
import statements
import timesheets

    
def core(usePrevMonth = False):
    p = period.Period()
    if usePrevMonth: p.decMonth()
    else: p.inputPeriod()
    print "Results will be stored in ", p.outDir()
    common.loginfo("begin")
    d = data.Data(p)
    d.load()
    statements.main(d)
    timesheets.main(d)
    d.store()
    common.loginfo("end")


def XXXprotectedProcess(usePrevMonth = False):        
    try:        
        core(usePrevMonth)
        print "########     ###     ######   ######  "
        print "##     ##   ## ##   ##    ## ##    ## "
        print "##     ##  ##   ##  ##       ##       "
        print "########  ##     ##  ######   ######  "
        print "##        #########       ##       ## "
        print "##        ##     ## ##    ## ##    ## "
        print "##        ##     ##  ######   ######  "
    except:
        print "########    ###    #### ##       "
        print "##         ## ##    ##  ##       "
        print "##        ##   ##   ##  ##       "
        print "######   ##     ##  ##  ##       "
        print "##       #########  ##  ##       "
        print "##       ##     ##  ##  ##       "
        print "##       ##     ## #### ######## "
        exception, msg, tb = sys.exc_info()
        txt = traceback.format_exc(tb)
        print txt
        common.logerror(txt)

def XXXmain(): 
    protectedProcess()
    print "Finished"
    #msvcrt.getch()

def profileit():
    "Use this if you want to perform a profile analysis of the code"
    fname = '../profresults.dat'
    cProfile.run('pypms.protectedProcess(True)', fname)
    p = pstats.Stats(fname)
    p.sort_stats('cumulative')
    p.print_stats()
    #p.dump_stats('../profile.txt')
    


if  __name__ == "__main__":
    print "You're on your own"
    main()
    print 'Finished'
