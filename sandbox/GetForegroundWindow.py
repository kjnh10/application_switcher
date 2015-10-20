import win32com
import win32gui
import win32process
hwnd = win32gui.GetForegroundWindow()
pid = win32process.GetWindowThreadProcessId(hwnd)
print (pid)
