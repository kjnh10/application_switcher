# coding: utf-8
import win32com.client
import win32gui
import win32process

shell = win32com.client.Dispatch("Wscript.Shell")
hwnd = win32gui.GetForegroundWindow()
pid = win32process.GetWindowThreadProcessId(hwnd)
#a = shell.AppActivate(pid[1])
a = shell.AppActivate(4520)
print pid
print a
