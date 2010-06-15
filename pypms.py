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
    #d.store()
    common.loginfo("end")


def protectedProcess(usePrevMonth = False):        
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

def main():
    protectedProcess()
    print "Finished. Press any key to exit"
    msvcrt.getch()

def profileit():
    "Use this if you want to perform a profile analysis of the code"
    cProfile.run('pypms.protectedProcess(True)', 'profresults')
    p = pstats.Stats('profresults')
    p.sort_stats('cumulative')
    p.print_stats()
    


if  __name__ == "__main__": main()
