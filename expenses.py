# Extract expenses from the spreadsheet

import win32com.client.dynamic

import common
import data
import excel
import period
#from common import AsAscii



def read_expenses(p):    
    rows = excel.ImportWorksheet(common.camelxls(p), 'Expenses')
    expenses = []
    fieldspec = [(1, 'JobCode', str), (2, 'Task', str), (4, 'Period', str),(6, 'Name', str), (8, 'Desc', str), (10, 'Amount', float)]
    for row in rows[1:]:
        expense = {}
        for colNum, fieldName, fieldType in fieldspec:
            text = row[colNum-1]
            try: expense[fieldName] = fieldType(text)
            except ValueError: pass
        if expense['JobCode']: expenses.append(expense) # only add an expense if it's "genuine" - i.e. it has a job number
    return expenses

def create_report(p, expenses):
    total = 0.0

    output = [['Job', 'Amount', 'Period', '', 'Name', '', 'Desc']]
    sorted_expenses = sorted(expenses, key= lambda x: x['JobCode'])
    for expense_item in sorted_expenses:
        job_code = expense_item['JobCode']
        per = expense_item['Period']
        name = expense_item['Name']
        desc = expense_item['Desc']
        desc_upper = desc.upper()
        if 'ACCOM' in desc_upper: desc += ' - Munros'
        if 'FLIGHT' in desc_upper: desc += ' - Munros'
        if 'TAXI' in desc_upper: desc += ' - Rainbow'        
        amount = expense_item['Amount']
        total += amount
        output.append([job_code, amount, per, '', name, '', desc])

    output.append([])
    output.append(['TOTAL', total])
    excel.create_report(p, "expenses", output, [2])
  
###########################################################################

def main(d):
    expenses = read_expenses(d.p)
    create_report(d.p, expenses)
    d.expenses = expenses
    
###########################################################################

if  __name__ == "__main__":
    data.run_current(main)
    print 'Finished'
