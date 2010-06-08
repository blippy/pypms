# Excel manipulation


import win32com.client.dynamic
#import numpy

#import yaml

import common
#from common import AsAscii
#import period




    
def ImportWorksheet(fileName, wsName):
    xlapp = win32com.client.dynamic.Dispatch("Excel.Application")
    #fileName = 'M:\\Finance\\camel\\%s\\camel-%s.xls' % (p.y, p.yyyymm())
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





if  __name__ == "__main__":
    print ImportWorksheet('M:\\Finance\\Invoices\\Inv summaries 2010\\Inv Summary 2010-04.xls', 'Invoices')
