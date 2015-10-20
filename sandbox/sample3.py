# coding: utf-8
import win32com.client
import win32gui
import win32process

WMI = win32com.client.GetObject("winmgmts:")
processes = WMI.InstancesOf("Win32_Process")
print len(processes)

chrome_name_list = []
for process in processes:
    if process.Properties_('Name').Value == "chrome.exe":
        chrome_name_list.append(process.Properties_('Name').Value)
print chrome_name_list

chrome_id_list = []
for process in processes:
    if process.Properties_('Name').Value == "chrome.exe":
        chrome_id_list.append(process.Properties_('ProcessId').Value)
print chrome_id_list
