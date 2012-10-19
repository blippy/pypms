import budget
import common
from common import princ, print_timing
import db
import excel
import expenses
import health
import html
import invoices
import period
import post
import recoveries
import statements
import timesheets
import tweaks
import wip

@print_timing
def main():
    p = period.g_period
    p.describe()
    cache = db.fetch()
    timesheets.create_timesheets(cache)
    
    excel.encache(cache)
    db.save_state(cache)
    excel.verify(cache)
    expenses.create_expense_report(cache)
    statements.create_statements(cache)
    
    all_invoices = invoices.enumerate_invoices(cache)
    invoices.create_text_invoice_summary(all_invoices)
        
    the_tweaks = cache['InvTweaks'] # tweaks.load(cache)
    accumulated_tweaks = tweaks.accum_tweaks_to_job_level(the_tweaks)
    post.post_main(cache, accumulated_tweaks)
    recoveries.create_recovery_report(cache, the_tweaks)
    wip.create_wip_report(cache, True) # self.cbox_text_wip.IsChecked())
    budget.create_budget(cache)
    health.create_health_report(cache)
    html.create_html()
    return cache # useful if we want to pickle it
    
if __name__ == "__main__":
    main()
    print 'Finished'