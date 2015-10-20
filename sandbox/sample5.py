import win32gui




win32gui.ShowWindow(HWND, win32con.SW_RESTORE)

win32gui.SetWindowPos(HWND,win32con.HWND_NOTOPMOST, 0, 0, 0,0, win32con.SWP_NOMOVE + win32con.SWP_NOSIZE )

win32gui.SetWindowPos(HWND,win32con.HWND_TOPMOST, 0, 0, 0,0,win32con.SWP_NOMOVE + win32con.SWP_NOSIZE )

win32gui.SetWindowPos(HWND,win32con.HWND_NOTOPMOST, 0, 0, 0,0,win32con.SWP_SHOWWINDOW + win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)
