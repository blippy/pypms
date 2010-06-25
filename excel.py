# Excel manipulation

import os

import win32com.client.dynamic
#import numpy

#import yaml

import common
#from common import AsAscii
#import period




    
def ImportWorksheet(fileName, wsName):
    xlapp = win32com.client.dynamic.Dispatch("Excel.Application")
    wb = xlapp.Workbooks.Open(fileName)
    ws = wb.Worksheets(wsName)
    numApproxRows = ws.UsedRange.Rows.Count # not necessarily truly accurate, but good enough
    numApproxCols = ws.UsedRange.Columns.Count # ditto
    numActualRows = 0
    numActualCols = 0
    sloppy = []
    for rowNum in xrange(0, numApproxRows):
        row = []
        for colNum in xrange(0,numApproxCols):
            value = common.AsAscii(ws.Cells(rowNum+1, colNum+1).Text)
            row.append(value)
            if len(value) > 0:
                numActualRows = rowNum +1
                if colNum >= numActualCols: numActualCols = colNum+1                                
        sloppy.append(row)
    wb.Close(SaveChanges = 0)
    xlapp.Quit()
    del xlapp
    
    # now reduce the number of rows and columns read in
    sloppy = sloppy[0:numActualRows]
    result = []
    for sloppyRow in sloppy:
        row = sloppyRow[0:numActualCols]
        result.append(row)
        
    return result


def create_workbook(file_name, func):
    if os.path.isfile(file_name): os.remove(file_name)
    xlapp = win32com.client.dynamic.Dispatch("Excel.Application")
    wb = xlapp.Workbooks.Add()
    func(wb)
    wb.SaveAs(file_name)
    wb.Close(SaveChanges = 0)
    xlapp.Quit()
    del xlapp



def create_report(p, desc, list_of_rows, numeric_fields):
    
    def excel_func(wb):        
        ws = wb.Worksheets('Sheet1')
        ws.Name = desc
        row_num = 0
        for row in list_of_rows:
            row_num += 1
            col_num = 0
            for col_value in row:
                col_num +=1
                ws.Cells(row_num, col_num).Value = col_value
            for col_number in numeric_fields:
                ws.Cells(row_num, col_number).NumberFormat = "0.00"

    file_name = common.reportdir(p) + '\\' + desc.lower() + '.xls'
    create_workbook(file_name, excel_func)
    
    
if  __name__ == "__main__":
    create_workbook('M:\\Finance\\pypms\\out\\2010-06\\craig\\test.xls')
    print 'Finished'
