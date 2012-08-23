# Excel manipulation

import datetime
import math
import os
import pdb

import win32com.client.dynamic

import xlrd

import common
from common import princ
import period

 
###########################################################################

MAX_ROWS = 65535
MAX_COLS = 255
        
def camelxls():
    return r"M:\Finance\pypms\{0}\summary-{0}.xls".format(period.yyyymm())

    
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

def gcell(ws, r, c):
    'Return the contents of a cell as text'
    #print r, " ", c
    return ws.Cells(r, c).Text


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
        file_name = camelxls()
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

def fix_date(v):
    if isinstance(v, basestring):
        result = v
    else:
        d1 = int(math.floor(v)) + datetime.datetime(1900,1,1).toordinal() -2
        d2 = datetime.datetime.fromordinal(d1)
        result = d2.strftime('%d-%b-%Y')
    return result

def fix_str(s):
    "Fix quirk in xlrd where what we think should be a string became converted to a float"
    s1 = str(s)
    if len(s1)>1:
        if s1[-2:] == '.0': s1 = s1[:-2]
    return s1

def workbook_name():
    return r"M:\Finance\pypms\{0}\summary-{0}.xls".format(period.yyyymm())

def read_worksheet(wsname):
    wb = None
    try:
        wb = xlrd.open_workbook(workbook_name())
        ws = wb.sheet_by_name(wsname)
        result = [ws.row_values(i) for i in xrange(ws.nrows)] #    nal.extend(sh.row_values(rowx))
        return result
    finally:
        if wb is not None: wb.release_resources()
        
def assert_worksheet_job(data, job_code, worksheet_name, row0):
    "Require that a job in an Excel spreadsheet exists in the database"
    if not data['jobs'].has_key(job_code):
        #msg = fmt.format(job_code, source_info)
        # logerror(msg) # TODO reinstate
        raise KeyError("E100", workbook_name(), worksheet_name, row0 + 1, job_code) 
    #fmt = "Excel workbook: {0}, sheet: {1}, row: {2}"
    #msg = fmt.format(workbook_name(), worksheet_name, row0 + 1)
    #common.assert_job(data, job_code, msg)
    
###########################################################################

def import_summary_sheet(data, wsname, fieldspec, start_row = 1):    
    rows = read_worksheet(wsname)    
    result = []
    for row_num in range(start_row, len(rows)):
        row = rows[row_num]
        record = {}
        for colNum, fieldName, fieldType, default in fieldspec:
            text = row[colNum-1]
            try: record[fieldName] = fieldType(text)
            except ValueError: record[fieldName] = default
        job_code = record['JobCode']
        if len(job_code) == 0: continue
        record['ExcelRow'] = row_num + 1
        
        #print record
        assert_worksheet_job(data, job_code, wsname, row_num)        
        
        # check that no records have None in them - they should all be set to something
        for k in record.keys():
            if record[k] is None:
                msg = "Can't have a key value of None for {0} in summary worksheet {1} with record {2}".format(k, wsname, record)
                raise common.DataIntegrityError(msg)
        
        result.append(record)
    return result

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
