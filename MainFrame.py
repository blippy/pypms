# -*- coding: iso-8859-15 -*-
# generated by wxGlade HG on Fri Feb 11 15:13:32 2011

import wx

# begin wxGlade: dependencies
# end wxGlade

# begin wxGlade: extracode

# end wxGlade

import os

import win32api
#import wx
#import wx.richtext

import common
#import db
import period
import pydra
import rtfsprint

import JobsFrame
import TimegridFrame

def open_file(filename):
    win32api.ShellExecute(0, "open", filename, None, ".", 0)
    
def long_calc(parent):
    'Allow bailout of action'
    caption = 'Danger'
    question = 'Long or dangerous action. Continue?'
    dlg = wx.MessageDialog(parent, question, caption, wx.YES_NO | wx.ICON_QUESTION)
    do_it = dlg.ShowModal() == wx.ID_YES
    dlg.Destroy()
    return do_it

class MainFrame(wx.Frame):
    def __init__(self, *args, **kwds):
        # begin wxGlade: MainFrame.__init__
        kwds["style"] = wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(self, *args, **kwds)
        
        # Menu Bar
        self.frmMain_menubar = wx.MenuBar()
        wxglade_tmp_menu = wx.Menu()
        self.menu_data_jobs = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Jobs", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.menu_data_jobs)
        self.menu_data_timegrid = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Time grid", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.menu_data_timegrid)
        self.frmMain_menubar.Append(wxglade_tmp_menu, "Data")
        wxglade_tmp_menu = wx.Menu()
        self.menu_externals_expenses = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Expenses", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.menu_externals_expenses)
        self.menu_externals_gizmo = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Gizmo", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.menu_externals_gizmo)
        self.menu_externals_html = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Html", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.menu_externals_html)
        self.menu_externals_invoice_summary = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Invoice Summary", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.menu_externals_invoice_summary)
        self.menu_externals_open_reports_folder = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Open Reports Folder", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.menu_externals_open_reports_folder)
        self.frmMain_menubar.Append(wxglade_tmp_menu, "Externals")
        wxglade_tmp_menu = wx.Menu()
        self.menu_print_timesheets = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Timesheets", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.menu_print_timesheets)
        self.menu_print_workstatements = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Workstatements", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.menu_print_workstatements)
        self.frmMain_menubar.Append(wxglade_tmp_menu, "Print")
        self.SetMenuBar(self.frmMain_menubar)
        # Menu Bar end
        self.notebook_1 = wx.Notebook(self, -1, style=0)
        self.notebook_1_pane_1 = wx.Panel(self.notebook_1, -1)
        self.btnAllStages = wx.Button(self.notebook_1_pane_1, -1, "All Stages")
        self.notebook_1_pane_2 = wx.Panel(self.notebook_1, -1)
        self.label_period = wx.StaticText(self.notebook_1_pane_2, -1, "label_1")
        self.btn_dec_period = wx.Button(self.notebook_1_pane_2, -1, "-")
        self.btn_inc_period = wx.Button(self.notebook_1_pane_2, -1, "+")
        self.text_output = wx.TextCtrl(self, -1, "", style=wx.TE_MULTILINE)

        self.__set_properties()
        self.__do_layout()

        self.Bind(wx.EVT_MENU, self.menu_data_jobs_selected, self.menu_data_jobs)
        self.Bind(wx.EVT_MENU, self.menu_data_timegrid_selected, self.menu_data_timegrid)
        self.Bind(wx.EVT_MENU, self.menu_externals_expenses_selected, self.menu_externals_expenses)
        self.Bind(wx.EVT_MENU, self.menu_externals_gizmo_selected, self.menu_externals_gizmo)
        self.Bind(wx.EVT_MENU, self.menu_externals_html_selected, self.menu_externals_html)
        self.Bind(wx.EVT_MENU, self.menu_externals_invoice_summary_selected, self.menu_externals_invoice_summary)
        self.Bind(wx.EVT_MENU, self.menu_externals_open_reports_folder_selected, self.menu_externals_open_reports_folder)
        self.Bind(wx.EVT_MENU, self.menu_print_timesheets_selected, self.menu_print_timesheets)
        self.Bind(wx.EVT_MENU, self.menu_print_workstatements_selected, self.menu_print_workstatements)
        self.Bind(wx.EVT_BUTTON, self.click_all_stages, self.btnAllStages)
        self.Bind(wx.EVT_BUTTON, self.btn_dec_period_clicked, self.btn_dec_period)
        self.Bind(wx.EVT_BUTTON, self.btn_inc_period_clicked, self.btn_inc_period)
        # end wxGlade
 

        # added by mcarter
        self.Bind(wx.EVT_CLOSE, self.OnClose)
        self.jobs_frame = JobsFrame.JobsFrame(self)
        self.timegrid_frame = TimegridFrame.TimegridFrame(self)
        
        def princ_func(text):
            self.text_output.AppendText(str(text) + '\n')
        
        common._princ_func = princ_func

    def __set_properties(self):
        # begin wxGlade: MainFrame.__set_properties
        self.SetTitle("Pydra GUI")
        self.SetSize((568, 507))
        self.text_output.SetFont(wx.Font(10, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, "Consolas"))
        # end wxGlade
        self.label_period.SetLabel(period.g_period.yyyymm())

    def __do_layout(self):
        # begin wxGlade: MainFrame.__do_layout
        sizer_1 = wx.BoxSizer(wx.VERTICAL)
        sizer_5 = wx.BoxSizer(wx.VERTICAL)
        sizer_10 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_3 = wx.BoxSizer(wx.VERTICAL)
        sizer_3.Add(self.btnAllStages, 0, 0, 0)
        self.notebook_1_pane_1.SetSizer(sizer_3)
        sizer_10.Add(self.label_period, 0, 0, 0)
        sizer_10.Add(self.btn_dec_period, 0, 0, 0)
        sizer_10.Add(self.btn_inc_period, 0, 0, 0)
        sizer_5.Add(sizer_10, 1, wx.EXPAND, 0)
        self.notebook_1_pane_2.SetSizer(sizer_5)
        self.notebook_1.AddPage(self.notebook_1_pane_1, "Process")
        self.notebook_1.AddPage(self.notebook_1_pane_2, "Settings")
        sizer_1.Add(self.notebook_1, 1, wx.EXPAND, 0)
        sizer_1.Add(self.text_output, 1, wx.EXPAND, 0)
        self.SetSizer(sizer_1)
        self.Layout()
        # end wxGlade

    def OnClose(self, event):        
        if self.jobs_frame.IsShown():
            wx.MessageBox("Close jobs form first", "Error")
        else:
            princ("Going down")
            self.Destroy()
        
    def menu_data_jobs_selected(self, event): # wxGlade: MainFrame.<event_handler>
        if not self.jobs_frame.IsShown():
            self.jobs_frame.Show()
    
    def menu_data_timegrid_selected(self, event): # wxGlade: MainFrame.<event_handler>
        if not self.timegrid_frame.IsShown():
            self.timegrid_frame.Show()
        
    def click_all_stages(self, event): # wxGlade: MainFrame.<event_handler>
        if not long_calc(self): return
        pydra.allstages()
        princ('Finished')
        wx.MessageBox('Finished', 'Info')



    def menu_externals_expenses_selected(self, event): # wxGlade: MainFrame.<event_handler>
        open_file(common.camelxls())

    def menu_externals_gizmo_selected(self, event): # wxGlade: MainFrame.<event_handler>
        open_file('M:\\Finance\\gizmo\\gizmo04.xls')

    def menu_externals_html_selected(self, event): # wxGlade: MainFrame.<event_handler>
        open_file('M:\\Finance\\pypms\\texts.htm')

    def menu_externals_invoice_summary_selected(self, event): # wxGlade: MainFrame.<event_handler>
        p = period.Period(usePrev = True)
        fname = '"M:\\Finance\\Invoices\\Inv summaries {0}\\Inv Summary {1}.xls"'.format(p.y, p.yyyymm())
        open_file(fname)

    def menu_externals_open_reports_folder_selected(self, event): # wxGlade: MainFrame.<event_handler>
        cmd = 'explorer ' + common.reportdir()
        os.system(cmd)



    def menu_print_timesheets_selected(self, event): # wxGlade: MainFrame.<event_handler>
        if not long_calc(self): return
        rtfsprint.timesheets()
        wx.MessageBox('Finished', 'Info')


    def menu_print_workstatements_selected(self, event): # wxGlade: MainFrame.<event_handler>
        if not long_calc(self): return
        rtfsprint.work_statements()


        

    def change_period(self, num_months):
        period.global_inc(num_months)
        self.label_period.SetLabel(period.g_period.yyyymm())
        
    def btn_dec_period_clicked(self, event): # wxGlade: MainFrame.<event_handler>
        self.change_period(-1)

    def btn_inc_period_clicked(self, event): # wxGlade: MainFrame.<event_handler>
        self.change_period(1)

# end of class MainFrame


