import invsummary
import invtweaks
import post
import recoveries
import wip

###########################################################################

def main(d):
    invsummary.import_manual_invoices(d)
    invsummary.create_invoice_summary(d)
    invtweaks.main(d)
    post.main(d)
    recoveries.main(d)
    wip.main(d)
    invsummary.create_reconciliation(d)

if  __name__ == "__main__":
    print "Didn't do anything"
