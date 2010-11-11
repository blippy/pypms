# Excel manipulation

import os
import pdb

import win32com.client.dynamic
#import xlrd
import common

 
def XXXxlopen():
    return win32com.client.dynamic.Dispatch("Excel.Application")

def XXXxlclose(xlapp):
    xlapp.Quit()
    del xlapp    
 
###########################################################################

#xlFormulas, , xlByRows, xlPrevious
def XXXnum_rows(ws):
    #global xlFormulas, , xlByRows, xlPrevious
    #rng = ws.Range("A1")
    #const = win32com.client.constants
    #find = ws.Cells.Find("*", "A1",   xlFormulas, , SearchOrder = const.xlByRows, SearchDirectory = const.xlPrevious)
    xlByRows = 1
    xlPrevious = 2
    cells = ws.Cells
    find = cells.Find("*", "A1", SearchOrder = xlByRows, SearchDirection = xlPrevious)
    num = find.Row
    return num
        
###########################################################################

def gcell(ws, r, c):
    'Return the contents of a cell as text'
    return ws.Cells(r, c).Text

def read_worksheet(fileName, wsName):
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
            value = common.AsAscii(gcell(ws, rowNum+1, colNum+1))
            row.append(value)
            if len(value) > 0:
                numActualRows = rowNum +1
                if colNum >= numActualCols: numActualCols = colNum+1                                
        sloppy.append(row)
    wb.Close(SaveChanges = 0)
    xlapp.Quit()
    del xlapp
    return sloppy, numActualRows, numActualCols

def prune_results(sloppy, numActualRows, numActualCols):
    # now reduce the number of rows and columns read in
    sloppy = sloppy[0:numActualRows]
    pruned = []
    for sloppyRow in sloppy:
        row = sloppyRow[0:numActualCols]
        pruned.append(row)        
    return pruned


def ImportWorksheet(fileName, wsName):
    sloppy, numActualRows, numActualCols = read_worksheet(fileName, wsName)
    pruned = prune_results(sloppy, numActualRows, numActualCols)
    return pruned


###########################################################################

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
    def xlfunc(wb):
        ws = wb.Worksheets('Sheet1')
        ws.Cells(1,1).Value = 'Basic test works'
    fname = 'c:\\test.xls'
    create_workbook(fname, xlfunc)
    print 'Simple spreadsheet created at', fname
    print 'Finished'
