# Extract expenses from the spreadsheet

import win32com.client.dynamic
#import yaml

import data
import excel
import period
#from common import AsAscii



def read_expenses(p):
    fileName = 'M:\\Finance\\camel\\%s\\camel-%s.xls' % (p.y, p.yyyymm())
    rows = excel.ImportWorksheet(fileName, 'Expenses')
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
    excel.create_report(p, "expenses", output, 2)
    
def main(p):
    expenses = read_expenses(p)
    create_report(p, expenses)
    return expenses
    


if  __name__ == "__main__":
    d = data.Data()
    d.restore()
    #p = period.Period(usePrev = True)
    expenses = main(d.p)
    #print expenses
    print 'Finished'
