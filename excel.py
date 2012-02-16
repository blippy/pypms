# Excel manipulation

import os
import pdb

import win32com.client.dynamic
import common
from common import princ
import period

 
###########################################################################

MAX_ROWS = 65535
MAX_COLS = 255
        
###########################################################################

def gcell(ws, r, c):
    'Return the contents of a cell as text'
    #print r, " ", c
    return ws.Cells(r, c).Text

def read_worksheetXXX(fileName, wsName, max_rows = MAX_ROWS, max_cols = MAX_COLS):
    '''max_rows - the maximum number of rows you are prepared to import (default = 70000)
    max_cols - ditto for columns (default = 300)
    '''
    xlapp = win32com.client.dynamic.Dispatch("Excel.Application")
    wb = xlapp.Workbooks.Open(fileName)
    ws = wb.Worksheets(wsName)
    numApproxRows = min(max_rows, ws.UsedRange.Rows.Count) # not necessarily truly accurate, but good enough
    numApproxCols = min(max_cols, ws.UsedRange.Columns.Count) # ditto
    
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

def prune_resultsXXX(sloppy, numActualRows, numActualCols):
    # now reduce the number of rows and columns read in
    sloppy = sloppy[0:numActualRows]
    pruned = []
    for sloppyRow in sloppy:
        row = sloppyRow[0:numActualCols]
        pruned.append(row)        
    return pruned


def ImportWorksheetXXX(fileName, wsName, max_rows = MAX_ROWS, max_cols = MAX_COLS):
    
    sloppy, numActualRows, numActualCols = read_worksheet(fileName, wsName, max_rows, max_cols)
    pruned = prune_results(sloppy, numActualRows, numActualCols)
    return pruned


def camelxlsXXX():
    #global g_period
    p = period.g_period
    file_name = "M:\\Finance\\pypms\\{0}\\pypms.xls".format(p.yyyymm())  # M:\Finance\camel\2011\camel-2011-04.xls
    #princ(file_name)
    return file_name

def ImportCamelWorksheetXXX(wsName, max_rows = MAX_ROWS, max_cols = MAX_COLS):
    #def camelxls():
    #'Return the filename for the Camel Excel input file'
    #fname = 'M:\\Finance\\camel\\%s\\camel-%s.xls' % (period.g_period.y, period.g_period.yyyymm())
    return ImportWorksheet(camelxls(), wsName, max_rows, max_cols)
    
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


###########################################################################


class Book:
    def __init__(self, filename):
        self.filename = filename
        if os.path.isfile(filename): os.remove(filename)
        self.xlapp = win32com.client.dynamic.Dispatch("Excel.Application")
        self.wb = self.xlapp.Workbooks.Add()
        
    def __del__(self):
        princ('Book:__del__()')
        self.wb.SaveAs(self.filename)
        self.wb.Close(SaveChanges = 0)
        self.xlapp.Quit()
        del self.xlapp 
        
###########################################################################

def create_report(desc, list_of_rows, numeric_fields):
    
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

    file_name = period.reportfile(desc.lower() + '.xls')
    create_workbook(file_name, excel_func)
    
###########################################################################
def read_worksheet(wb, wsName, max_rows = MAX_ROWS, max_cols = MAX_COLS):
    ws = wb.Worksheets(wsName)    
    numApproxRows = min(max_rows, ws.UsedRange.Rows.Count) # not necessarily truly accurate, but good enough
    numApproxCols = min(max_cols, ws.UsedRange.Columns.Count) # ditto
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
    sloppy = sloppy[0:numActualRows]
    
    pruned = []
    for sloppyRow in sloppy:
        row = sloppyRow[0:numActualCols]
        pruned.append(row)        
    return pruned


def import_excel_data():
    xlapp = None
    wb = None    
    try:
        p = period.g_period
        file_name = "M:\\Finance\\pypms\\{0}\\pypms.xls".format(p.yyyymm()) 
        excel_data = {}
        xlapp = win32com.client.dynamic.Dispatch("Excel.Application")
        wb = xlapp.Workbooks.Open(file_name)
        for sheet in [('Expenses', 200, 10), ('InvTweaks', 200, 10) , ('ManualInvoices', 200, 7)]:
            wsname, mrows, mcols = sheet
            excel_data[wsname] = read_worksheet(wb, wsname, mrows, mcols)

    finally:
        if wb is not None: wb.Close(SaveChanges = 0)
        if xlapp is not None:
            xlapp.Quit()
            del xlapp
    return excel_data # sloppy, numActualRows, numActualCols


###########################################################################
if  __name__ == "__main__":
    data = import_excel_data()
    princ(data)
    #def xlfunc(wb):
    #    ws = wb.Worksheets('Sheet1')
    #    ws.Cells(1,1).Value = 'Basic test works'
    #fname = 'c:\\test.xls'
    #create_workbook(fname, xlfunc)
    #princ('Simple spreadsheet created at ' + fname)
    princ('Finished')
