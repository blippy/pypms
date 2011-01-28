#!/usr/bin/env python
# -*- coding: iso-8859-15 -*-
# generated by wxGlade HG on Thu Jan 13 15:09:26 2011


# begin wxGlade: extracode
# end wxGlade

import os

import win32api
import wx
import wx.richtext

import common
import period
import pydra
import rtfsprint


def open_file(filename):
    win32api.ShellExecute(0, "open", filename, None, ".", 0)

    

class MainFrame(wx.Frame):
    def __init__(self, *args, **kwds):
        # begin wxGlade: MainFrame.__init__
        kwds["style"] = wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(self, *args, **kwds)
        self.notebook_1 = wx.Notebook(self, -1, style=0)
        self.notebook_1_pane_1 = wx.Panel(self.notebook_1, -1)
        self.btnAllStages = wx.Button(self.notebook_1_pane_1, -1, "All Stages")
        self.notebook_1_pane_2 = wx.Panel(self.notebook_1, -1)
        self.btnExpenses = wx.Button(self.notebook_1_pane_2, -1, "Expenses")
        self.btnGizmo = wx.Button(self.notebook_1_pane_2, -1, "Gizmo")
        self.btnInvoiceSummary = wx.Button(self.notebook_1_pane_2, -1, "Invoice Summary")
        self.btnOpenReportsFolder = wx.Button(self.notebook_1_pane_2, -1, "Open Reports Folder")
        self.notebook_1_pane_3 = wx.Panel(self.notebook_1, -1)
        self.btnPrintWorkstatements = wx.Button(self.notebook_1_pane_3, -1, "Workstatements")
        self.btnPrintTimesheets = wx.Button(self.notebook_1_pane_3, -1, "Timesheets")

        self.__set_properties()
        self.__do_layout()

        self.Bind(wx.EVT_BUTTON, self.click_all_stages, self.btnAllStages)
        self.Bind(wx.EVT_BUTTON, self.click_expenses, self.btnExpenses)
        self.Bind(wx.EVT_BUTTON, self.click_gizmo, self.btnGizmo)
        self.Bind(wx.EVT_BUTTON, self.click_invoice_summary, self.btnInvoiceSummary)
        self.Bind(wx.EVT_BUTTON, self.click_open_reports_folder, self.btnOpenReportsFolder)
        self.Bind(wx.EVT_BUTTON, self.click_print_workstatements, self.btnPrintWorkstatements)
        self.Bind(wx.EVT_BUTTON, self.click_print_timesheets, self.btnPrintTimesheets)
        # end wxGlade

    def __set_properties(self):
        # begin wxGlade: MainFrame.__set_properties
        self.SetTitle("Pydra GUI")
        self.SetSize((294, 218))
        # end wxGlade

    def __do_layout(self):
        # begin wxGlade: MainFrame.__do_layout
        sizer_1 = wx.BoxSizer(wx.VERTICAL)
        sizer_2 = wx.BoxSizer(wx.VERTICAL)
        sizer_4 = wx.BoxSizer(wx.VERTICAL)
        sizer_3 = wx.BoxSizer(wx.VERTICAL)
        sizer_3.Add(self.btnAllStages, 0, 0, 0)
        self.notebook_1_pane_1.SetSizer(sizer_3)
        sizer_4.Add(self.btnExpenses, 0, 0, 0)
        sizer_4.Add(self.btnGizmo, 0, 0, 0)
        sizer_4.Add(self.btnInvoiceSummary, 0, 0, 0)
        sizer_4.Add(self.btnOpenReportsFolder, 0, 0, 0)
        self.notebook_1_pane_2.SetSizer(sizer_4)
        sizer_2.Add(self.btnPrintWorkstatements, 0, 0, 0)
        sizer_2.Add(self.btnPrintTimesheets, 0, 0, 0)
        self.notebook_1_pane_3.SetSizer(sizer_2)
        self.notebook_1.AddPage(self.notebook_1_pane_1, "Process")
        self.notebook_1.AddPage(self.notebook_1_pane_2, "Externals")
        self.notebook_1.AddPage(self.notebook_1_pane_3, "Printing")
        sizer_1.Add(self.notebook_1, 1, wx.EXPAND, 0)
        self.SetSizer(sizer_1)
        self.Layout()
        # end wxGlade

    def click_print_timesheets(self, event): # wxGlade: MainFrame.<event_handler>
        rtfsprint.timesheets()
        wx.MessageBox('Finished', 'Info')

    def click_all_stages(self, event): # wxGlade: MainFrame.<event_handler>
        pydra.allstages()
        wx.MessageBox('Finished', 'Info')
        #print "Event handler `click_all_stages' not implemented"
        #event.Skip()

    def click_open_reports_folder(self, event): # wxGlade: MainFrame.<event_handler>        
        cmd = 'explorer ' + common.reportdir()
        os.system(cmd)
        
    def click_print_workstatements(self, event): # wxGlade: MainFrame.<event_handler>
        rtfsprint.work_statements()

    def click_expenses(self, event): # wxGlade: MainFrame.<event_handler>
        open_file(common.camelxls())
        #print "Event handler `click_expenses' not implemented"
        #event.Skip()

    def click_gizmo(self, event): # wxGlade: MainFrame.<event_handler>
        open_file('M:\Finance\gizmo\gizmo04.xls')

    def click_invoice_summary(self, event): # wxGlade: MainFrame.<event_handler>
        p = period.Period(usePrev = True)
        fname = 'M:\Finance\Invoices\Inv summaries {0}\Inv Summary {0}-{1}.xls'.format(p.y, p.m)
        open_file(fname)

# end of class MainFrame


if __name__ == "__main__":
    app = wx.PySimpleApp(0)
    wx.InitAllImageHandlers()
    frmMain = MainFrame(None, -1, "")
    app.SetTopWindow(frmMain)
    frmMain.Show()
    app.MainLoop()
