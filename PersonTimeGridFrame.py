# -*- coding: iso-8859-15 -*-
# generated by wxGlade 0.6.3 on Thu Jul 26 14:34:55 2012

import wx

# begin wxGlade: dependencies
import wx.grid
# end wxGlade

# begin wxGlade: extracode

# end wxGlade

import common
import db
import period

def fetch_people():
    result = []
    emps = db.GetEmployees()
    for k  in sorted(emps.keys()):
        e = emps[k]
        if e['Active'] == 0: continue
        entry = [k, '{0:4.4s} | {1:s}'.format( k, e['PersonNAME'])]
        result.append(entry)
    return result

def aggregate_time_items(initials):
    titems = db.GetTimeitems()
    titems = filter(lambda(x) : x['Person'] == initials, titems)
    #print titems
    
    codings = [ (x['JobCode'], x['Task']) for x in titems]
    codings.sort()
    codings = common.unique(codings)
    
    result = []
    for coding in codings:
        time_vals = []
        for d in range(1, 32):
            total = 0
            for titem in titems:
                if titem['JobCode'] == coding[0] and titem['Task'] == coding[1] and titem['DateVal'].day == d:
                    total += titem['TimeVal']
            time_vals.append(total)
        result.append([coding[0], coding[1], time_vals])
    return result

class PersonTimeGridFrame(wx.Frame):
    def __init__(self, *args, **kwds):
        # begin wxGlade: PersonTimeGridFrame.__init__
        kwds["style"] = wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(self, *args, **kwds)
        self.button_fetch_peope = wx.Button(self, -1, "Fetch People")
        self.choice_person = wx.Choice(self, -1, choices=[])
        self.button_calculate = wx.Button(self, -1, "Calculate")
        self.grid_time = wx.grid.Grid(self, -1, size=(1, 1))

        self.__set_properties()
        self.__do_layout()

        self.Bind(wx.EVT_BUTTON, self.clicked_fetch_people, self.button_fetch_peope)
        self.Bind(wx.EVT_BUTTON, self.clicked_calculate, self.button_calculate)
        # end wxGlade

        # added by mcarter
        self.Bind(wx.EVT_CLOSE, self.OnClose)

    def OnClose(self, event):
        self.Hide()
        
        
    def __set_properties(self):
        # begin wxGlade: PersonTimeGridFrame.__set_properties
        self.SetTitle("Person time Grid")
        self.grid_time.CreateGrid(1, 3)
        self.grid_time.SetColLabelValue(0, "JOB")
        self.grid_time.SetColLabelValue(1, "TASK")
        self.grid_time.SetColLabelValue(2, "SUM")
        self.grid_time.SetMinSize((1250,600))
        # end wxGlade

    def __do_layout(self):
        # begin wxGlade: PersonTimeGridFrame.__do_layout
        sizer_16 = wx.BoxSizer(wx.VERTICAL)
        sizer_17 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_17.Add(self.button_fetch_peope, 0, 0, 0)
        sizer_17.Add(self.choice_person, 0, 0, 0)
        sizer_17.Add(self.button_calculate, 0, 0, 0)
        sizer_16.Add(sizer_17, 0, wx.EXPAND, 0)
        sizer_16.Add(self.grid_time, 1, wx.EXPAND, 0)
        self.SetSizer(sizer_16)
        sizer_16.Fit(self)
        self.Layout()
        # end wxGlade

    def clicked_fetch_people(self, event): # wxGlade: PersonTimeGridFrame.<event_handler>
        #print "Event handler `clicked_fetch_people' not implemented"
        #event.Skip()
        self.choice_person.Clear()
        font1 = wx.Font(10, wx.MODERN, wx.NORMAL, wx.NORMAL, False, u'Consolas')
        self.choice_person.SetFont(font1)
        self.choice_entries = fetch_people()
        for entry in self.choice_entries:
            #print entry
            self.choice_person.Append(entry[1])
        self.choice_person.Select(0)
        

    def clicked_calculate(self, event): # wxGlade: PersonTimeGridFrame.<event_handler>
        #print "Event handler `clicked_calculate' not implemented"
        #event.Skip()
        if self.choice_person.GetCount() == 0: return
        DATE_COL0 = 3
        initials = self.choice_entries[self.choice_person.GetSelection()] #  'CJO' # fixme
        initials = initials[0]
        #print "initials = ", initials
        rows = aggregate_time_items(initials)
        grid = self.grid_time
        common.empty_wxgrid(grid)
        #p = period.g_period
        #self.label_period.SetLabel(p.yyyymm())
        dim = period.dim()
        common.rectify_num_grid_columns(grid, dim + DATE_COL0)
        for c in range(0, dim):
            grid.SetColLabelValue(c + DATE_COL0 , str(c + 1))
            grid.SetColSize(c+DATE_COL0, 30)
        grid.ForceRefresh()
        
        # totals
        total_row = []
        for c in range(0, dim):
            total = 0
            for row in rows:
                total +=  row[2][c]
                #print total
            #total = [row[c+2]  for row in rows]            
            total_row.append(total)
        total_row = [ 'TOTAL', '', total_row]
        rows.append(total_row)
            
        
        row_num = -1
        for row in rows:
            row_num += 1
            grid.AppendRows(1)
            job_code, task, time_spent = row
            grid.SetCellValue(row_num, 0, job_code)
            grid.SetCellValue(row_num, 1, task)
            grid.SetCellValue(row_num, 2, str(sum(time_spent)))
            for d in range(0, dim):
                v = time_spent[d]
                if v == 0: v = ''
                grid.SetCellValue(row_num,d + DATE_COL0, str(v))

                # TODO a lot of this code is in common with TimegridFrame
                # colour the cell according to whether it is a week day or end
                if period.is_weekend(d+1):
                    bcolour = wx.LIGHT_GREY
                else:
                    bcolour = wx.WHITE
                grid.SetCellBackgroundColour(row_num,d+DATE_COL0, bcolour)
            #print row
            #grid.SetCellValue(row_num,c+DATE_COL0,v)

# end of class PersonTimeGridFrame


if __name__ == "__main__":
    #fetch_people()
    #aggregate_time_items('CJO')
    #exit(0)
    
    app = wx.PySimpleApp(0)
    wx.InitAllImageHandlers()
    frame = PersonTimeGridFrame(None, -1, "")
    app.SetTopWindow(frame)
    frame.Show()
    app.MainLoop()
    frame.Destroy()

