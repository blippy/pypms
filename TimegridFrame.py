# -*- coding: iso-8859-15 -*-
# generated by wxGlade HG on Thu Feb 24 16:17:32 2011

import array

import wx

# begin wxGlade: dependencies
import wx.grid
# end wxGlade

# begin wxGlade: extracode

# end wxGlade

import db
import period

class DataRow:
    def __init__(self, initials, name, dim):
        self.initials = initials
        self.name = name
        self.times = array.array('f', [0.0] * dim)
        
class TimegridFrame(wx.Frame):
    def __init__(self, *args, **kwds):
        # begin wxGlade: TimegridFrame.__init__
        kwds["style"] = wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(self, *args, **kwds)
        self.btnCalculate = wx.Button(self, -1, "Calculate")
        self.label_period = wx.StaticText(self, -1, "YYYY-MM", style=wx.ALIGN_CENTRE)
        self.grid_time = wx.grid.Grid(self, -1, size=(1, 1))

        self.__set_properties()
        self.__do_layout()

        self.Bind(wx.EVT_BUTTON, self.btn_calculate_clicked, self.btnCalculate)
        # end wxGlade
        
        # added by mcarter
        self.Bind(wx.EVT_CLOSE, self.OnClose)

    def OnClose(self, event):
        self.Hide()
        
    def __set_properties(self):
        # begin wxGlade: TimegridFrame.__set_properties
        self.SetTitle("Time grid")
        self.SetSize((1334, 1041))
        self.label_period.SetBackgroundColour(wx.Colour(255, 255, 255))
        self.grid_time.CreateGrid(0, 3)
        self.grid_time.SetColLabelValue(0, "IN")
        self.grid_time.SetColLabelValue(1, "NAME")
        self.grid_time.SetColLabelValue(2, "SUM")
        # end wxGlade

    def __do_layout(self):
        # begin wxGlade: TimegridFrame.__do_layout
        sizer_2 = wx.BoxSizer(wx.VERTICAL)
        sizer_4 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_4.Add(self.btnCalculate, 0, 0, 0)
        sizer_4.Add(self.label_period, 0, wx.EXPAND|wx.ALIGN_CENTER_VERTICAL, 0)
        sizer_2.Add(sizer_4, 0, wx.EXPAND|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 0)
        sizer_2.Add(self.grid_time, 1, wx.EXPAND, 0)
        self.SetSizer(sizer_2)
        self.Layout()
        # end wxGlade

    def btn_calculate_clicked(self, event): # wxGlade: TimegridFrame.<event_handler>
    
        grid = self.grid_time
        num_rows = grid.GetNumberRows()
        if num_rows >0:
            grid.DeleteRows(0, num_rows)
        

        
        p = period.g_period
        self.label_period.SetLabel(p.yyyymm())
        dim = p.dim()
        
        #maybe delete some columns
        DATE_COL0 = 3
        num_cols = grid.GetNumberCols()
        num_cols_required = dim + DATE_COL0
        if num_cols_required > num_cols:
            grid.AppendCols(num_cols_required - num_cols)
        elif num_cols_required < num_cols:
            grid.DeleteCols(dim + DATE_COL0 - 1, num_cols - num_cols_required)
        #number the columns
        for c in range(0, dim):
            grid.SetColLabelValue(c + DATE_COL0 , str(c + 1))
            grid.SetColSize(c+DATE_COL0, 30)
        grid.ForceRefresh()
        
        
        employees = db.GetEmployees(p)
        data_rows = {}
        for initials in employees.keys():
            data_rows[initials] = DataRow(initials, employees[initials]['PersonNAME'], dim)
        #print employees

        # accumulate time items by person and date
        timeItems = db.GetTimeitems(p)
        for timeItem in timeItems:
            initials = timeItem['Person']
            dstamp = timeItem['DateVal']
            day = int(dstamp[-2:])
            #print day
            qty = timeItem['TimeVal']
            if not data_rows.has_key(initials):
                data_rows[initials] = DataRow(initials, '???', dim)
            data_rows[initials].times[day-1] += qty
        #print timeItems

        
        # output the results in a grid
        row_num = -1
        for initials in sorted(data_rows):
            row_num += 1
            data_row = data_rows[initials]
            grid.AppendRows(1)
            grid.SetCellValue(row_num,0, initials)
            grid.SetCellValue(row_num,1, data_row.name)
            grid.SetCellValue(row_num, 2, str(sum(data_row.times)))
            for c in range(0, dim):
                v = data_row.times[c]
                if v == 0: v = ''
                v = str(v)
                #v = 'foo'
                grid.SetCellValue(row_num,c+DATE_COL0,v)
                
                # colour the cell according to whether it is a week day or end
                if period.is_weekend(p.y, p.m, c+1):
                    bcolour = wx.LIGHT_GREY
                else:
                    bcolour = wx.WHITE
                grid.SetCellBackgroundColour(row_num,c+DATE_COL0, bcolour)
                
        grid.SetFocus()

# end of class TimegridFrame



if __name__ == "__main__":
    app = wx.PySimpleApp(0)
    wx.InitAllImageHandlers()
    frame = TimegridFrame(None, -1, "")
    app.SetTopWindow(frame)
    frame.Show()
    app.MainLoop()
    frame.Destroy()
