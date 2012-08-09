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

def import_expenses():
    excel_rows = excel.read_worksheet('Expenses')
    result = []
    fieldspec = [(1, 'JobCode', excel.fix_str), (2, 'Task', str), (4, 'Period', excel.fix_date),(6, 'Name', str), (8, 'Desc', str), (10, 'Amount', float)]
    row_num = 0
    for row in excel_rows[1:]:
        row_num += 1
        expense = {}
        for colNum, fieldName, fieldType in fieldspec:
            text = row[colNum-1]
            try: expense[fieldName] = fieldType(text)
            except ValueError: pass
        job_code = expense['JobCode']
        if job_code == '' or not expense.has_key('Amount'): continue
    
        # TODO HIGH error checking
        
        result.append(expense)
    return result
        
def create_expense_report(sorted_expenses, output_text = True):
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
            
            
def process(data):
    the_expenses =  import_expenses()
    the_expenses.sort(key= lambda x: x['JobCode'])
    create_expense_report(the_expenses)
    data['expenses'] = the_expenses
    
###########################################################################
if  __name__ == "__main__":
    debug = True
    process({})
    princ("Finished")

