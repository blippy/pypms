# Extract expenses from the spreadsheet

import time

import win32com.client.dynamic

import common
import dbnew
from dbnew import TableCache
import excel
import period
from common import princ, print_timing




class ExpenseCache(TableCache):
    def __init__(self):
        TableCache.__init__(self, 'tblCamelExpenses')
        
    @print_timing
    def import_expenses(self):
        princ("Importing expenses")
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
            if job_code == '': continue
        
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
    def create_expense_report(self):
        total = 0.0

        output = [['Job', 'Amount', 'Period', '', 'Name', '', 'Desc']]
        sorted_expenses = sorted(self.expenses, key= lambda x: x['JobCode'])
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
        excel.create_report("expenses", output, [2])
    
#cache = db.TableCache('tblCamelExpenses')
cache = ExpenseCache()
#print cache.table_name


    


  
###########################################################################

def main(d):
    read_expenses(d)
    create_report(d)
    
###########################################################################

if  __name__ == "__main__":
    princ("Didn't do anything")
