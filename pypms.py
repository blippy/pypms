# TODO deprecate this module
import code,  datetime, msvcrt, os, pdb, pprint, sys, traceback

import common
import db
import period
import statements
import timesheets

    



def XXXprotectedProcess(usePrevMonth = False):        
    try:        
        #core(usePrevMonth)
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

    


if  __name__ == "__main__":
    print "You're on your own"
