import budget
import common
from common import princ, print_timing
import db
import excel
import expenses
import health
import html
import invsummary
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
    
    # TODO the excel data should be cached instead of passed around
    excel_data = excel.import_excel_data()
    # spreadsheet importation
    #if self.cbox_expenses.IsChecked():
    expenses.cache.import_expenses(excel_data['Expenses'])
    invsummary.import_manual_invoices(cache, excel_data['ManualInvoices'])            
    
    
    #if self.cbox_expenses.IsChecked():
    expenses.cache.create_expense_report()
    statements.create_statements(cache)
    
    invoices = invsummary.enumerate_invoices(cache)
    #if self.cbox_text_invoices.IsChecked():
    invsummary.create_text_invoice_summary(invoices)
    #else:
    #    invsummary.create_excel_invoice_summary(invoices)
        
    the_tweaks = tweaks.decipher_tweaks(excel_data['InvTweaks'])
    accumulated_tweaks = tweaks.accum_tweaks_to_job_level(the_tweaks)
    post.post_main(cache, accumulated_tweaks)
    recoveries.create_recovery_report(cache, the_tweaks)
    wip.create_wip_report(cache, True) # self.cbox_text_wip.IsChecked())
    budget.create_budget(cache)
    health.create_health_report(cache)
    html.create_html()
    #princ('Finished')
    return cache # useful if we want to pickle it
    
if __name__ == "__main__":
    main()