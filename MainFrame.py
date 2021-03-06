# -*- coding: iso-8859-15 -*-
# generated by wxGlade HG on Fri Feb 11 15:13:32 2011

import wx

# begin wxGlade: dependencies
# end wxGlade

# begin wxGlade: extracode

# end wxGlade

# notes on windows redirects:
# http://www.velocityreviews.com/forums/t515815-wxpython-redirect-the-stdout-to-a-textctrl.html    

import datetime
import os
import sys

import win32api

import common
from common import princ
import db
import excel
import period
import pydra
import registry
import rtfsprint


import AboutDialog
import ExpensesFrame
import JobsFrame
import TimegridFrame
import PersonTimeGridFrame

def open_file(filename):
    win32api.ShellExecute(0, "open", filename, None, ".", 0)
    
def long_calc(parent, question = 'Long or dangerous action. Continue?'):
    'Allow bailout of action'
    caption = 'Danger'
    dlg = wx.MessageDialog(parent, question, caption, wx.YES_NO | wx.ICON_QUESTION)
    do_it = dlg.ShowModal() == wx.ID_YES
    dlg.Destroy()
    return do_it

class RedirectText:
    def __init__(self,aWxTextCtrl):
        self.out=aWxTextCtrl
        
    def write(self,string):
        self.out.WriteText(string)

#-------------------------- Create Icon --------------------------
def GetMondrianData():
    return \
'\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00 \x00\x00\x00 \x08\x06\x00\
\x00\x00szz\xf4\x00\x00\x00\x04sBIT\x08\x08\x08\x08|\x08d\x88\x00\x00\x00qID\
ATX\x85\xed\xd6;\n\x800\x10E\xd1{\xc5\x8d\xb9r\x97\x16\x0b\xad$\x8a\x82:\x16\
o\xda\x84pB2\x1f\x81Fa\x8c\x9c\x08\x04Z{\xcf\xa72\xbcv\xfa\xc5\x08 \x80r\x80\
\xfc\xa2\x0e\x1c\xe4\xba\xfaX\x1d\xd0\xde]S\x07\x02\xd8>\xe1wa-`\x9fQ\xe9\
\x86\x01\x04\x10\x00\\(Dk\x1b-\x04\xdc\x1d\x07\x14\x98;\x0bS\x7f\x7f\xf9\x13\
\x04\x10@\xf9X\xbe\x00\xc9 \x14K\xc1<={\x00\x00\x00\x00IEND\xaeB`\x82' 
def GetMondrianBitmap():
    return wx.BitmapFromImage(GetMondrianImage())
def GetMondrianImage():
    import cStringIO
    stream = cStringIO.StringIO(GetMondrianData())
    return wx.ImageFromStream(stream)
def GetMondrianIcon():
    icon = wx.EmptyIcon()
    icon.CopyFromBitmap(GetMondrianBitmap())
    return icon  

class MainFrame(wx.Frame):
    def __init__(self, *args, **kwds):
        # begin wxGlade: MainFrame.__init__
        kwds["style"] = wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(self, *args, **kwds)
        
        # Menu Bar
        self.frmMain_menubar = wx.MenuBar()
        wxglade_tmp_menu = wx.Menu()
        self.menu_data_expenses = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Expenses", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.menu_data_expenses)
        self.menu_data_jobs = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Jobs", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.menu_data_jobs)
        self.menu_data_timegrid = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Time grid", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.menu_data_timegrid)
        self.menu_data_persontimegrid = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Person Time grid", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.menu_data_persontimegrid)
        wxglade_tmp_menu.AppendSeparator()
        self.menu_data_pickle = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Pickle", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.menu_data_pickle)
        self.frmMain_menubar.Append(wxglade_tmp_menu, "Data")
        wxglade_tmp_menu = wx.Menu()
        self.menu_externals_gizmo = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Gizmo", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.menu_externals_gizmo)
        self.menu_externals_html = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Html", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.menu_externals_html)
        self.menu_externals_open_reports_folder = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Open Reports Folder", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.menu_externals_open_reports_folder)
        self.menu_externals_spreadsheet = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Spreadsheet", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.menu_externals_spreadsheet)
        self.frmMain_menubar.Append(wxglade_tmp_menu, "Externals")
        wxglade_tmp_menu = wx.Menu()
        self.menu_print_workstatements = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Workstatements", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.menu_print_workstatements)
        self.menu_print_timesheets = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Timesheets", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.menu_print_timesheets)
        self.frmMain_menubar.Append(wxglade_tmp_menu, "Print")
        wxglade_tmp_menu = wx.Menu()
        self.menu_help_about = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "About", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.menu_help_about)
        self.frmMain_menubar.Append(wxglade_tmp_menu, "Help")
        self.SetMenuBar(self.frmMain_menubar)
        # Menu Bar end
        self.notebook_1 = wx.Notebook(self, -1, style=0)
        self.notebook_1_pane_2 = wx.Panel(self.notebook_1, -1)
        self.label_period = wx.StaticText(self.notebook_1_pane_2, -1, "label_1")
        self.btn_dec_period = wx.Button(self.notebook_1_pane_2, -1, "-")
        self.btn_inc_period = wx.Button(self.notebook_1_pane_2, -1, "+")
        self.notebook_1_pane_1 = wx.Panel(self.notebook_1, -1)
        self.btnAllStages = wx.Button(self.notebook_1_pane_1, -1, "All Stages")
        self.cbox_expenses = wx.CheckBox(self.notebook_1_pane_1, -1, "Import expenses")
        self.cbox_text_expenses = wx.CheckBox(self.notebook_1_pane_1, -1, "Expenses as text, not XL")
        self.cbox_text_invoices = wx.CheckBox(self.notebook_1_pane_1, -1, "Invoice summary as text, not XL")
        self.cbox_text_wip = wx.CheckBox(self.notebook_1_pane_1, -1, "WIP as text, not XL")
        self.notebook_options = wx.Panel(self.notebook_1, -1)
        self.cbox_autopickle = wx.CheckBox(self.notebook_options, -1, "Auto-pickle data")
        self.text_output = wx.TextCtrl(self, -1, "", style=wx.TE_MULTILINE)

        self.__set_properties()
        self.__do_layout()

        self.Bind(wx.EVT_MENU, self.menu_data_expenses_selected, self.menu_data_expenses)
        self.Bind(wx.EVT_MENU, self.menu_data_jobs_selected, self.menu_data_jobs)
        self.Bind(wx.EVT_MENU, self.menu_data_timegrid_selected, self.menu_data_timegrid)
        self.Bind(wx.EVT_MENU, self.menu_data_persontimegrid_selected, self.menu_data_persontimegrid)
        self.Bind(wx.EVT_MENU, self.menu_data_pickle_selected, self.menu_data_pickle)
        self.Bind(wx.EVT_MENU, self.menu_externals_gizmo_selected, self.menu_externals_gizmo)
        self.Bind(wx.EVT_MENU, self.menu_externals_html_selected, self.menu_externals_html)
        self.Bind(wx.EVT_MENU, self.menu_externals_open_reports_folder_selected, self.menu_externals_open_reports_folder)
        self.Bind(wx.EVT_MENU, self.menu_externals_spreadsheet_selected, self.menu_externals_spreadsheet)
        self.Bind(wx.EVT_MENU, self.menu_print_workstatements_selected, self.menu_print_workstatements)
        self.Bind(wx.EVT_MENU, self.menu_print_timesheets_selected, self.menu_print_timesheets)
        self.Bind(wx.EVT_MENU, self.menu_help_about_selected, self.menu_help_about)
        self.Bind(wx.EVT_BUTTON, self.btn_dec_period_clicked, self.btn_dec_period)
        self.Bind(wx.EVT_BUTTON, self.btn_inc_period_clicked, self.btn_inc_period)
        self.Bind(wx.EVT_BUTTON, self.click_all_stages, self.btnAllStages)
        # end wxGlade
 

        # added by mcarter
        self.Bind(wx.EVT_CLOSE, self.OnClose)
        self.expenses_frame = ExpensesFrame.ExpensesFrame(self)
        self.jobs_frame = JobsFrame.JobsFrame(self)
        self.timegrid_frame = TimegridFrame.TimegridFrame(self)
        self.persontimegrid_frame = PersonTimeGridFrame.PersonTimeGridFrame(self)
        registry.RegistryBoundCheckbox(self, self.cbox_expenses, 'import_expenses', True)
        registry.RegistryBoundCheckbox(self, self.cbox_text_wip, 'wip_as_text', True)
        registry.RegistryBoundCheckbox(self, self.cbox_text_expenses, 'expenses_as_text', True)
        registry.RegistryBoundCheckbox(self, self.cbox_text_invoices, 'invoices_as_text', True)
        registry.RegistryBoundCheckbox(self, self.cbox_autopickle, 'autopickle', False)
        self.cache = None
        
        # redirect stdout and sterr to window
        redir = RedirectText(self.text_output)
        sys.stdout = redir
        sys.stderr = redir
        
        #def princ_func(text):
        #    self.text_output.AppendText(str(text) + '\n')        
        #common._princ_func = princ_func
 
        # iconification and taskbar
        # http://wxpython-users.1045709.n5.nabble.com/minimize-to-try-question-td2359957.html
        self.icon = GetMondrianIcon()
        self.SetIcon(self.icon) 
        self.tbicon = wx.TaskBarIcon()                
        self.Bind(wx.EVT_ICONIZE, self.OnIconify)
        self.tbicon.Bind(wx.EVT_TASKBAR_LEFT_DCLICK, self.OnTaskBarActivate) 
        self.Show() 
 
    def OnTaskBarActivate(self, evt):
        if self.IsIconized():
            self.Iconize(False)
            self.Show()
            self.Raise()
            self.tbicon.RemoveIcon()
 
    def OnIconify(self, evt):  
        if evt.Iconized():
            self.Iconize(True)
            self.Hide()
            self.tbicon.SetIcon(self.icon) 

    def __set_properties(self):
        # begin wxGlade: MainFrame.__set_properties
        self.SetTitle("Pydra GUI")
        self.SetSize((568, 507))
        self.text_output.SetFont(wx.Font(10, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, "Consolas"))
        # end wxGlade
        self.label_period.SetLabel(period.yyyymm())
        
        #use_prev_month = common.get_defaulted_binary_reg_key('UsePrevMonth', True)
        #self.cbox_use_prev_month.SetValue(use_prev_month)

    def __do_layout(self):
        # begin wxGlade: MainFrame.__do_layout
        sizer_1 = wx.BoxSizer(wx.VERTICAL)
        sizer_15 = wx.BoxSizer(wx.VERTICAL)
        sizer_3 = wx.BoxSizer(wx.VERTICAL)
        sizer_5 = wx.BoxSizer(wx.VERTICAL)
        sizer_10 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_10.Add(self.label_period, 0, 0, 0)
        sizer_10.Add(self.btn_dec_period, 0, 0, 0)
        sizer_10.Add(self.btn_inc_period, 0, 0, 0)
        sizer_5.Add(sizer_10, 1, wx.EXPAND, 0)
        self.notebook_1_pane_2.SetSizer(sizer_5)
        sizer_3.Add(self.btnAllStages, 0, 0, 0)
        sizer_3.Add(self.cbox_expenses, 0, 0, 0)
        sizer_3.Add(self.cbox_text_expenses, 0, 0, 0)
        sizer_3.Add(self.cbox_text_invoices, 0, 0, 0)
        sizer_3.Add(self.cbox_text_wip, 0, 0, 0)
        self.notebook_1_pane_1.SetSizer(sizer_3)
        sizer_15.Add(self.cbox_autopickle, 0, 0, 0)
        self.notebook_options.SetSizer(sizer_15)
        self.notebook_1.AddPage(self.notebook_1_pane_2, "Settings")
        self.notebook_1.AddPage(self.notebook_1_pane_1, "Process")
        self.notebook_1.AddPage(self.notebook_options, "Options")
        sizer_1.Add(self.notebook_1, 0, wx.EXPAND, 0)
        sizer_1.Add(self.text_output, 1, wx.EXPAND, 0)
        self.SetSizer(sizer_1)
        self.Layout()
        # end wxGlade

    def OnClose(self, event):        
        if self.jobs_frame.IsShown() or self.expenses_frame.IsShown():
            wx.MessageBox("Close other open forms first", "Error")
        else:
            princ("Going down")
            self.Destroy()
        
    def menu_data_jobs_selected(self, event): # wxGlade: MainFrame.<event_handler>
        if not self.jobs_frame.IsShown():
            self.jobs_frame.Show()
    
    def menu_data_timegrid_selected(self, event): # wxGlade: MainFrame.<event_handler>
        if not self.timegrid_frame.IsShown():
            self.timegrid_frame.Show()

    def menu_data_persontimegrid_selected(self, event): # wxGlade: MainFrame.<event_handler>
        if not self.persontimegrid_frame.IsShown():
            self.persontimegrid_frame.Show()
            
    def click_all_stages(self, event): # wxGlade: MainFrame.<event_handler>
        # main process loop
        question = 'Process period {0}?'.format(period.yyyymm())
        if not long_calc(self, question): return
    
        self.text_output.SetValue("")
        self.text_output.Refresh()
        princ('Started: ' + str(datetime.datetime.now()))
        wx.SetCursor ( wx.StockCursor ( wx.CURSOR_WAIT ) ) 
        try:
            self.cache = pydra.main()
            princ('Finished: ' + str(datetime.datetime.now()))
        finally:
            #princ('Entered finally')
            wx.SetCursor ( wx.StockCursor ( wx.CURSOR_ARROW ) )
            
        wx.MessageBox('Finished', 'Info')




    def menu_externals_gizmo_selected(self, event): # wxGlade: MainFrame.<event_handler>
        open_file('M:\\Finance\\gizmo\\gizmo04.xls')

    def menu_externals_html_selected(self, event): # wxGlade: MainFrame.<event_handler>
        open_file('M:\\Finance\\pypms\\texts.htm')


    def menu_externals_open_reports_folder_selected(self, event): # wxGlade: MainFrame.<event_handler>
        cmd = 'explorer ' + period.reportdir()
        os.system(cmd)



    def menu_print_timesheets_selected(self, event): # wxGlade: MainFrame.<event_handler>
        if not long_calc(self): return
        rtfsprint.timesheets()
        wx.MessageBox('Finished', 'Info')


    def menu_print_workstatements_selected(self, event): # wxGlade: MainFrame.<event_handler>
        if not long_calc(self): return
        rtfsprint.work_statements()


        

    def change_period(self, num_months):
        period.inc(num_months)
        self.label_period.SetLabel(period.yyyymm())
        
    def btn_dec_period_clicked(self, event): # wxGlade: MainFrame.<event_handler>
        self.change_period(-1)

    def btn_inc_period_clicked(self, event): # wxGlade: MainFrame.<event_handler>
        self.change_period(1)

    def cbox_use_prev_month_changed(self, event): # wxGlade: MainFrame.<event_handler>
        checked = self.cbox_use_prev_month.IsChecked()
        common.set_binary_reg_value('UsePrevMonth', checked)

    def menu_data_expenses_selected(self, event): # wxGlade: MainFrame.<event_handler>
        if not self.expenses_frame.IsShown():
            self.expenses_frame.Show()

    def menu_data_pickle_selected(self, event): # wxGlade: MainFrame.<event_handler>
        db.save_state(self.cache)


    def menu_externals_spreadsheet_selected(self, event): # wxGlade: MainFrame.<event_handler>
        open_file(excel.camelxls())


    def menu_help_about_selected(self, event):  # wxGlade: MainFrame.<event_handler>
        #print "Event handler `menu_help_about_selected' not implemented"
        #event.Skip()
        dlg = AboutDialog.AboutDialog(self)
        dlg.ShowModal()
        dlg.Destroy()

# end of class MainFrame


