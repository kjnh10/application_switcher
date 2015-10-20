# coding: utf-8
import win32com.client
import win32gui
import win32process

WMI = win32com.client.GetObject("winmgmts:") #WMIの取得_windows-managements
processes = WMI.InstancesOf("Win32_Process")
print len(processes)

chrome_name_list = []
for process in processes:
    if process.Properties_('Name').Value == "ConEmu.exe":
        print process.Properties_('Name').Value
        chrome_name_list.append(process.Properties_('Name').Value)

id_list = []
for process in processes:
    if process.Properties_('Name').Value == "ConEmu.exe":
        id_list.append(process.Properties_('ProcessId').Value)
#print id_list

shell = win32com.client.Dispatch("Wscript.Shell")
if id_list != []:
    a = shell.AppActivate(id_list[0])
    #うまく開けなかった場合は、新しく開く
    if a == False:
        shell.run("C:\Users\koji\program\Cmder\Cmder.exe")

else:
    shell.run("C:\Users\koji\program\Cmder\Cmder.exe")
