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



if  __name__ == "__main__":
    create_workbook('M:\\Finance\\pypms\\out\\2010-06\\craig\\test.xls')
    print 'Finished'
