# Excel manipulation

import datetime
import math
import os
import pdb
import pprint

import win32com.client.dynamic

import xlrd

import common
from common import princ
import db
import period

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
        #if s1[-2:] == '.0': s1 = s1[:-2]
        if s1[0] == '~': s1 = s1[1:]
    return s1

f = common.AsFloat
s = fix_str

schema = [  ('Expenses', 1, [       (1, 'JobCode', s, ''),
                                    (2, 'Task', s, ''), 
                                    (4, 'Period', fix_date, ''),
                                    (6, 'Name', str, ''),
                                    (8, 'Desc', str, ''), 
                                    (10, 'Amount', float, None)]),
            ('InvTweaks', 2, [      (1, 'JobCode', s, ''),
                                    (2, 'InvBIA', f, 0.0),
                                    (3, 'InvUBI', f, 0.0),
                                    (4, 'InvWIP', f, 0.0),
                                    (5, 'InvAccrual', f, 0.0),
                                    (6, 'InvInvoice', f, 0.0),
                                    (7, 'Inv3rdParty', f, 0.0),
                                    (8, 'InvTime', f, 0.0),
                                    (9, 'Recovery', f, 0.0),
                                    (10, 'Comment', s, '')]),
            ('ManualInvoices', 1, [ (1, 'irn', s, ''),
                                    (2, 'client', str, ''),
                                    (3, 'JobCode', s, ''),
                                    (4, 'net', float, None),
                                    (7, 'desc', str, '')])]

###########################################################################


def camelxls():
    return r"M:\Finance\pypms\{0}\summary-{0}.xls".format(period.yyyymm())

    
###########################################################################


class Source:
    def __init__(self, wsname, row):
        self.wsname = wsname
        self.row = row
        
    def __repr__(self):
        return self.__str__()
        
    def __str__(self):
        text = 'Source: Excel: ' + str(self.__dict__)
        return text
        
  
def import_summary_sheet(wb, wsname, start_row, fieldspec):    
    rows = read_worksheet(wb, wsname)
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
        record['source'] = Source(wsname, row_num + 1)
        #record['type'] = wsname
        
        # check that no records have None in them - they should all be set to something
        for k in record.keys():
            if record[k] is None:
                msg = "Can't have a key value of None for {0} in summary worksheet {1} with record {2}".format(k, wsname, record)
                raise KeyError(msg)
        
        result.append(record)
    return result


#m_xlmodtime = 0
#m_xldata = None

def load():
    global schema                                    
    xldata = {}
    wb = None
    try:
        wb = xlrd.open_workbook(camelxls())
        for layout in schema:
            wsname, start_row, fieldspec = layout
            xldata[wsname] = import_summary_sheet(wb, wsname, start_row, fieldspec)
    finally:
        if wb is not None: wb.release_resources()
    return xldata
    
###########################################################################

def verify(dbase):
    #global m_xldata
    #pprint.pprint(dbase)
    #pprint.pprint(m_xldata)
    #return
    global schema
    
    # collect errors
    errors = []
    
    
    for rec in dbase['Expenses']:
        if rec['Task'] == '':
            if not dbase['jobs'].has_key(rec['JobCode']): errors.append(rec)
        else:
            if not dbase['tasks'].has_key((rec['JobCode'], rec['Task'])): errors.append(rec)
            
    for rec in dbase['InvTweaks'] + dbase['ManualInvoices']:
        if not dbase['jobs'].has_key(rec['JobCode']): errors.append(rec)
        
    if len(errors) == 0: return

    # we have some errors, so print them
    def pr(): print "-----------------------------------"
    pr()
    print "Errors in Excel data:"
    for e in errors:
        print e['source']
    pr()
    raise KeyError("Spreadsheet job(s)")

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




def read_worksheet(wb, wsname):
    ws = wb.sheet_by_name(wsname)
    result = [ws.row_values(i) for i in xrange(ws.nrows)] #    nal.extend(sh.row_values(rowx))
    return result
        
def assert_worksheet_job(data, job_code, worksheet_name, row0):
    "Require that a job in an Excel spreadsheet exists in the database"
    if not data['jobs'].has_key(job_code):
        #msg = fmt.format(job_code, source_info)
        # logerror(msg) # TODO reinstate
        raise KeyError("E100", workbook_name(), worksheet_name, row0 + 1, job_code) 
    
###########################################################################

def encache(cache):
    xldata = load()
    cache.update(xldata)
    


###########################################################################
if  __name__ == "__main__":
    xldata = load()
    pprint.pprint(xldata)
    #cache = db.load_state()    
    #pprint.pprint(cache.keys())
    #encache(cache)
    #verify(cache)
    princ('Finished')
