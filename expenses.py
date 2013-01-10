# Extract expenses from the spreadsheet



import pdb
import time

import win32com.client.dynamic

import common
import db
import excel
import period
from common import dget, princ


###########################################################################


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

            

    
###########################################################################
if  __name__ == "__main__":
    debug = True
    process({})
    princ("Finished")

