import data
import invsummary
import post
import recoveries
import wip

###########################################################################

def main(d):
    post.main()
    invsummary.main(d)
    recoveries.main(d)
    wip.main(d)

if  __name__ == "__main__":
    data.run_current(main)
    print 'Finished'
