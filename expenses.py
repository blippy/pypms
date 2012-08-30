# Extract expenses from the spreadsheet



import pdb
import time

import win32com.client.dynamic

import common
import db
import excel
import period
from common import dget, princ, print_timing


###########################################################################

def XXXload(data):
    fieldspec = [
        (1, 'JobCode', excel.fix_str, ''),
        (2, 'Task', str, ''), 
        (4, 'Period', excel.fix_date, ''),
        (6, 'Name', str, ''),
        (8, 'Desc', str, ''), 
        (10, 'Amount', float, None)]
    return excel.import_summary_sheet(data, 'Expenses', fieldspec, 1)

def create_expense_report(data, output_text = True):
    the_expenses =  data['Expenses']
    sorted_expenses= sorted(the_expenses, key= lambda x: x['JobCode'])

    output = [['Job', 'Amount', 'Period', '', 'Name', '', 'Desc']]
    
    xs = filter(lambda x: x['Amount'] != 0, sorted_expenses)
    for x in xs:
        desc = x['Desc']
        desc_upper = desc.upper()
        if 'ACCOM' in desc_upper: desc += ' - Munros'
        if 'FLIGHT' in desc_upper: desc += ' - Munros'
        if 'TAXI' in desc_upper: desc += ' - Rainbow'        
        output.append([ x['JobCode'], x['Amount'], x['Period'], '', x['Name'], '', desc])

    output.append([])
    total = common.summate_lod(sorted_expenses, 'Amount')
    output.append(['TOTAL', total])
    
    if output_text:
        period.create_text_report("expenses.txt", output)
    else:
        # TODO Consider zapping Excel output option if considered unecessary
        excel.create_report("expenses", output, [2, 3 ,4, 5, 6, 7])

def XXX_create_expense_report(sorted_expenses, output_text = True):
    total = 0.0

    output = [['Job', 'Amount', 'Period', '', 'Name', '', 'Desc']]
    
    for expense_item in sorted_expenses:
        job_code = dget(expense_item,'JobCode', "")
        per = dget(expense_item, 'Period', "")
        name = dget(expense_item, 'Name', "")
        desc = dget(expense_item, 'Desc', "")
        desc_upper = desc.upper()
        if 'ACCOM' in desc_upper: desc += ' - Munros'
        if 'FLIGHT' in desc_upper: desc += ' - Munros'
        if 'TAXI' in desc_upper: desc += ' - Rainbow'        
        amount = dget(expense_item, 'Amount', 0.0)
        if amount == 0.0: continue
        total += amount
        output.append([job_code, amount, per, '', name, '', desc])

    output.append([])
    output.append(['TOTAL', total])
    
    if output_text:
        period.create_text_report("expenses.txt", output)
    else:
        # TODO Consider zapping Excel output option if considered unecessary
        excel.create_report("expenses", output, [2, 3 ,4, 5, 6, 7])
            

    
###########################################################################
if  __name__ == "__main__":
    debug = True
    process({})
    princ("Finished")

