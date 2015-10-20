# coding: utf-8
import win32com.client
import win32gui
import win32process
from ctypes import *

user32 = windll.user32
#user32.MessageBoxA(
#        0,
#        "Hello",
#        "Python to Windows API",
#        0x00000040)
user32.AllowSetForegroundWindow("ASFW_ANY")

WMI = win32com.client.GetObject("winmgmts:") #WMIの取得_windows-managements
processes = WMI.InstancesOf("Win32_Process")
print len(processes)

#name_list = []
#for process in processes:
#    if process.Properties_('Name').Value == "gvim.exe":
#        print process.Properties_('Name').Value
#        name_list.append(process.Properties_('Name').Value)

id_list = []
for process in processes:
    if process.Properties_('Name').Value == "gvim.exe":
        id_list.append(process.Properties_('ProcessId').Value)
#print id_list

shell = win32com.client.Dispatch("Wscript.Shell")
if id_list != []:
    a = shell.AppActivate(id_list[0])
    print a
    #うまく開けなかった場合は、新しく開く
