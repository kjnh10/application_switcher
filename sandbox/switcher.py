# coding: utf-8
from win32gui import *
from win32con import * #SW_SHOWNORMALを呼ぶため
import subprocess

#TOPLEVEL_WINDOWの取得
def getAllwindowAslist(hwnd, lParam):
    window_name = GetWindowText(hwnd).decode("shift-jis").encode("utf8")
    window_class = GetClassName(hwnd).decode("shift-jis").encode("utf8")
    window_list.append({"name":window_name,"hwnd":hwnd, "class":window_class})

def getAllChildWindow(hwnd, param):
    c_window_name = GetWindowText(hwnd).decode("shift-jis").encode("utf8")
    c_window_list.append({"name":c_window_name,"hwnd":hwnd})

def activate_app(hwnd):
    if IsIconic(hwnd):
        print "最小化されていたので,リストアしました"
        ShowWindow(hwnd,SW_RESTORE)
    else:
        print "最小化されていなかったので、SetForegroundWindowを使いました。"
        result = SetForegroundWindow(hwnd)

window_list = []
c_window_list = []
class Window():
    """
    this is Window class. A window object has its name,its hwnd,its class name.
    """
    def __init__(self, class_name, window_title=None, application=None):
        self.pro = {}
        self.application = application
        #propertyの取得
        self.class_hit_count = 0
        self.title_hit_count = 0

        EnumWindows(getAllwindowAslist, None)
        for window in window_list:
            if window["class"] == class_name:
                self.class_hit_count += 1
                print( "name:" + window["name"])

                #child window が指定されていれば、取得する
                if window_title != None:
                    EnumChildWindows(window["hwnd"], getAllChildWindow,None)
                    for c_window in c_window_list:
                        if c_window["name"] == window_title:
                            self.title_hit_count += 1
                            self.pro["hwnd"] = c_window["hwnd"]
                            self.pro["name"] = c_window["name"]
                #そうでなければ、そのままclassでヒットしたトップレベルウインドウを取得する。
                else:
                        self.pro["hwnd"] = window["hwnd"]
                        self.pro["name"] = window["name"]

    def activate_app(self):
        if self.class_hit_count == 0:
            print "specified class_name found no windows hits…"
            subprocess.call(self.application)
        elif IsIconic(self.pro["hwnd"]):
            print "最小化されていたので,リストアしました"
            ShowWindow(self.pro["hwnd"],SW_RESTORE)
        else:
            print "最小化されていなかったので、SetForegroundWindowを使いました。"
            result = SetForegroundWindow(self.pro["hwnd"])

    def debug(self):
        print(self.class_hit_count)
        print(self.title_hit_count)
        print(self.pro["hwnd"])
        EnumWindows(getAllwindowAslist, None)
        for window in window_list:
            print( window["class"] +":::::"+ window["name"])

