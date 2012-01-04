# Extract expenses from the spreadsheet



import pdb
import time

import win32com.client.dynamic

import common
import db
import dbnew
from dbnew import TableCache
import excel
import period
from common import dget, princ, print_timing

debug = False



class ExpenseCache(TableCache):
    def __init__(self):
        TableCache.__init__(self, 'tblCamelExpenses')
        
    @print_timing
    def import_expenses(self):
        #princ("Importing expenses")
        #input_filename = period.camelxls()
        rows = excel.ImportCamelWorksheet('Expenses', 200, 10) # set likely maximum rows an cols in sheet
        expenses = []
        fieldspec = [(1, 'JobCode', str), (2, 'Task', str), (4, 'Period', str),(6, 'Name', str), (8, 'Desc', str), (10, 'Amount', float)]
        row_num = 0
        for row in rows[1:]:
            row_num += 1
            expense = {}
            for colNum, fieldName, fieldType in fieldspec:
                text = row[colNum-1]
                try: expense[fieldName] = fieldType(text)
                except ValueError: pass
            job_code = expense['JobCode']
            if job_code == '' or not expense.has_key('Amount'): continue
        
            # peform error checking
            # TODO should possibly reinstate this somehow
            #err_msg = "Workbook '{0}', sheet '{1}', row {2}".format(input_filename, 'Expenses', row_num+1)
            #common.assert_job(d, job_code, err_msg)
            #if expense['Task']:
            #    common.assert_task(d, job_code, expense['Task'], err_msg)
            
            # everything seems OK, so add it
            expenses.append(expense)
        self.expenses = expenses

    @print_timing
    def create_expense_report(self, output_text = True):
        total = 0.0

        output = [['Job', 'Amount', 'Period', '', 'Name', '', 'Desc']]
        sorted_expenses = sorted(self.expenses, key= lambda x: x['JobCode'])
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
    
cache = ExpenseCache()


###########################################################################

if  __name__ == "__main__":
    debug = True
    cache.import_expenses()
    cache.create_expense_report(True)
    pdb.set_trace()
    princ("Finished")

