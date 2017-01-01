# coding: utf-8
from win32gui import *
from win32con import * #SW_SHOWNORMALを呼ぶため
import subprocess

window_list = []
def getAllwindowAslist(hwnd, lParam):
    window_name = GetWindowText(hwnd).decode("shift-jis").encode("utf8")
    window_class = GetClassName(hwnd).decode("shift-jis").encode("utf8")
    window_list.append({"name":window_name,"hwnd":hwnd, "class":window_class})

c_window_list = []
def getAllChildWindow(hwnd, param):
    c_window_name = GetWindowText(hwnd).decode("shift-jis").encode("utf8")
    c_window_list.append({"name":c_window_name,"hwnd":hwnd})

def activate_app(hwnd):
    #print(window["name"])
    if IsIconic(hwnd):
        print("最小化されていたので,リストアしました")
        ShowWindow(hwnd,SW_RESTORE)
    else:
        print("最小化されていなかったので、SetForegroundWindowを使いました。")
        result = SetForegroundWindow(hwnd)

EnumWindows(getAllwindowAslist, None)
if window_list != []:
    for window in window_list:
        #print( window["class"] +":::::"+ window["name"])
        if window["class"] == "CabinetWClass":
            EnumChildWindows(window["hwnd"], getAllChildWindow,None)
            for c_window in c_window_list:
                if c_window["name"] == "ShellView":
                    activate_app(c_window["hwnd"])
            break
    else:
        subprocess.call("explorer.exe")
