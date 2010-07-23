import data
import invsummary
import invtweaks
import post
import recoveries
import wip

###########################################################################

def main(d):
    invsummary.main(d)
    invtweaks.main(d)
    recoveries.main(d)
    wip.main(d)
    post.main(d)

if  __name__ == "__main__":
    data.run_current(main)
    print 'Finished'
