# to construct the executable, type
# python setup.py py2exe

import glob
import sys

from distutils.core import setup
import py2exe

#ms_dll_path=r'C:\WINDOWS\WinSxS\x86_Microsoft.VC90.CRT_1fc8b3b9a1e18e3b_9.0.30729.1_x-ww_6f74963e'
#sys.path.append(ms_dll_path)
#ms_dlls = glob.glob(ms_dll_path + '\\*.*')
#setup(data_files = [("Microsoft.VC90.CRT", ms_dlls)],
#    windows = ['app.py'])

settings = [{
    'script' : 'app.py',
    'dest_base' : 'pydra'
    }]
 
setup(windows = settings)
#setup(dest_base = 'pydrag-app', windows = ['app.py'])

#distutils.core.run_setup("setup", "py2exe")