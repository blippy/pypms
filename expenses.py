# Extract expenses from the spreadsheet

import win32com.client.dynamic
#import yaml

import excel, period
#from common import AsAscii


def main(p):
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

def XXXmain(p):
    xlapp = win32com.client.dynamic.Dispatch("Excel.Application")
    fileName = 'M:\\Finance\\camel\\%s\\camel-%s.xls' % (p.y, p.yyyymm())
    wb = xlapp.Workbooks.Open(fileName)
    ws = wb.Worksheets('Expenses')
    numRows = ws.UsedRange.Rows.Count # not necessarily truly accurate, but good enough
    expenses = []
    fieldspec = [(1, 'JobCode', str), (2, 'Task', str), (4, 'Period', str),(6, 'Name', str), (8, 'Desc', str), (10, 'Amount', float)]
    for rowNum in xrange(2, numRows+1):
        expense = {}
        for colNum, fieldName, fieldType in fieldspec:
            text = ws.Cells(rowNum, colNum).Text
            try: expense[fieldName] = fieldType(text)
            except ValueError: pass
        if expense['JobCode']: expenses.append(expense) # only add an expense if it's "genuine" - i.e. it has a job number
    wb.Close(SaveChanges = 0)
    xlapp.Quit()
    del xlapp
    return expenses





if  __name__ == "__main__":
    p = period.Period(usePrev = True)
    expenses = main(p)
    print expenses
