"""
Runtime hook for PyInstaller: make tkinter able to find Tcl/Tk
when the app is frozen into a one-file executable.

PyInstaller extracts data to sys._MEIPASS. Tcl/Tk data lives in
'tcl8.6' and 'tk8.6' subdirectories under the extraction root.
Set TCL_LIBRARY / TK_LIBRARY accordingly so _tkinter can locate
its initialization files.
"""
import os
import sys

if getattr(sys, 'frozen', False):
    base = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    tcl_dir = os.path.join(base, 'tcl8.6')
    tk_dir  = os.path.join(base, 'tk8.6')
    if os.path.isdir(tcl_dir):
        os.environ['TCL_LIBRARY'] = tcl_dir
    if os.path.isdir(tk_dir):
        os.environ['TK_LIBRARY'] = tk_dir
