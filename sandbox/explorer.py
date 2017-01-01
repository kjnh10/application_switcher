# coding: utf-8
import win32com.client
import win32gui
import win32process
import subprocess

WMI = win32com.client.GetObject("winmgmts:") #WMIの取得_windows-managements
processes = WMI.InstancesOf("Win32_Process")
print(len(processes))

name_list = []
for process in processes:
    if process.Properties_('Name').Value == "explorer.exe":
        print(process.Properties_('Name').Value)
        name_list.append(process.Properties_('Name').Value)

id_list = []
for process in processes:
    if process.Properties_('Name').Value == "explorer.exe":
        id_list.append(process.Properties_('ProcessId').Value)
#print(id_list)

shell = win32com.client.Dispatch("Wscript.Shell")
if id_list != []:
    b = False
    print(id_list[0])
    a = shell.AppActivate(id_list[0])
    if len(id_list) >= 2:
        b = shell.AppActivate(id_list[1])
    print(a)
    #うまく開けなかった場合は、新しく開く
    if a == False or b == False:
        #shell.run("C:\\Program\ Files\ (x86)\\Google\\Chrome\\Application\\chrome.exe")
        subprocess.call("C:\Windows\explorer.exe")
else:
    subprocess.call("C:\Windows\explorer.exe")
