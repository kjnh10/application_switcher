# coding: utf-8
from win32gui import *
from win32con import * #SW_SHOWNORMALを呼ぶため

import win32com.client
import subprocess
import os

shell = win32com.client.Dispatch("Wscript.Shell")

window_list = []
class Window():
    """
    this is Window class. A window object has its name,its hwnd,its class name.
    """
    def __init__(self, class_title=None, window_title=None, application=None):
        self.searching_class_title = class_title
        self.searching_window_title = window_title
        self.searching_application = application
        self.prop = {}
        self.application = application
        self.hit_count = 0

        global window_list
        window_list = [] #初期化
        EnumChildWindows(None,getWindows, None) #window_listの生成
        for window in window_list:
            if (class_title == None or window["class"] == class_title) and (window_title == None or window_title in window["title"]):
                self.hit_count += 1
                self.prop["hwnd"] = window["hwnd"]
                self.prop["title"] = window["title"]
                self.prop["class"] = window["class"]

    def become_child(self, class_title, window_title):
        if self.hit_count != 0:
            global window_list
            window_list = [] #初期化
            EnumChildWindows(self.prop["hwnd"], getWindows, None) #window_listの生成
            for window in window_list:
                if (class_title == None or window["class"] == class_title) and (window_title == None or window_title in window["title"]):
                    self.hit_count += 1
                    self.prop["hwnd"] = window["hwnd"]
                    self.prop["title"] = window["title"]
                    self.prop["class"] = window["class"]
        else:
            return

    def activate_app(self):
        if self.hit_count == 0:
            print "The specified condition doesn't match any windows…"
            #subprocess.call(self.application)
            #最大化がの方法がsubprocess.callではわからなかったため。
            shell.run(self.application, 3)
            self.__init__(self.searching_class_title, self.searching_window_title, self.searching_application)
            return
        elif IsIconic(self.prop["hwnd"]):
            print "最小化されていたので,リストアしました"
            ShowWindow(self.prop["hwnd"],SW_RESTORE)
        else:
            print "最小化されていなかったので、SetForegroundWindowを使いました。"
            result = SetForegroundWindow(self.prop["hwnd"])

    def debug(self):
        print("The number of hits is " + str(self.hit_count))
        if self.hit_count == 0:
            print "The specified condition doesn't match any windows…"
        else:
            print("class is " + self.prop["class"])
            print("title is " + self.prop["title"])
            print("window handle is " + str(self.prop["hwnd"]))
        self.show_scope()

    #うまく作動しない。
    def maximize(self):
        ShowWindow(self.prop["hwnd"], SW_MAXIMIZE)

    def show_scope(self):
        print("\n検索対象は以下のものです。")
        for window in window_list:
            print( window["class"] +":::::"+ window["title"])
        return

#---------------------sub procedure----------------------------------------------------

def getWindows(hwnd, lParam):
    window_name = GetWindowText(hwnd).decode("shift-jis").encode("utf8")
    window_class = GetClassName(hwnd).decode("shift-jis").encode("utf8")
    window_list.append({"title":window_name,"hwnd":hwnd, "class":window_class})

def activate_app(hwnd):
    if IsIconic(hwnd):
        print "最小化されていたので,リストアしました"
        ShowWindow(hwnd,SW_RESTORE)
    else:
        print "最小化されていなかったので、SetForegroundWindowを使いました。"
        result = SetForegroundWindow(hwnd)
