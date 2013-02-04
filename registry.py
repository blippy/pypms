###########################################################################
# registry functions

# binary values are kooky on Windows, as there is no default binary 
# registry type. They therefore have to be hand-crafted

import _winreg

import wx

from common import princ


__reg_root = 'Software\\Pydra'

def get_reg_key(key_name):
    global __reg_root
    key = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER, __reg_root, 0, _winreg.KEY_ALL_ACCESS)
    v = _winreg.QueryValueEx(key, key_name)
    _winreg.CloseKey(key)
    return v

def get_defaulted_reg_key(key_name, default):    
    try:
        v = get_reg_key(key_name)
    except WindowsError:
        v = default
    return v
    
def set_reg_value(key_name, new_value, reg_type = _winreg.REG_SZ):
    global __reg_root
    try:
        key = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER, __reg_root, 0, _winreg.KEY_ALL_ACCESS)
    except:
        key = _winreg.CreateKey(_winreg.HKEY_CURRENT_USER, __reg_root)
    _winreg.SetValueEx(key, key_name, 0, reg_type, new_value)
    _winreg.CloseKey(key)
    
def get_defaulted_binary_reg_key(key_name, default):
    try:
        raw = get_reg_key(key_name)
        result = raw[0] == 1
    except WindowsError:
        result = default
    return result

def set_binary_reg_value(key_name, value):
    set_reg_value(key_name, value, reg_type = _winreg.REG_DWORD)

###########################################################################
class RegistryBoundCheckbox:
    '''A checkbox that is bound to the registry'''
    
    def __init__(self, frame, checkbox, registry_key, default_value):
        self.registry_key = registry_key
        frame.Bind(wx.EVT_CHECKBOX, self.cbox_change_event, checkbox)
        value = get_defaulted_binary_reg_key(registry_key, default_value)
        checkbox.SetValue(value)
        self.checkbox = checkbox
        
    def cbox_change_event(self, event):
        princ("cbox_change_event")
        set_binary_reg_value(self.registry_key, self.checkbox.GetValue())
        


###########################################################################
