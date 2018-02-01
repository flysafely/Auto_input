from tkinter import *
from tkinter import ttk
from tkinter.filedialog import *
from tkinter.messagebox import *
from PIL import Image, ImageTk

import tkinter
import urllib.request
import re
import os
import http.cookiejar 
import win32com.client
import threading
import pythoncom  # 多线程调用COM
import time
import win32api,win32con 
import winshell
import subprocess
import MainFunction as MF

global download_ProgressValue
global cookies,GET_MAIL_HEADER

Default_Params = {
    "UserName":"xafei", 
    "UserPassword":"111", 
    "Method":"ZIP", 
    "TargetHOST":"130.130.200.30", 
    "LoginURL":"http://130.130.200.49/kmext/ext/sso.jsp", 
    "MailreceiveURL":"http://130.130.200.30/applications/email/mailreceive.aspx", 
    "TimeOut":5, 
    "KeywordArray":[" ","."]
}

class myThread (threading.Thread):

    def __init__(self, functions):
        threading.Thread.__init__(self)
        self.functions = functions
        self.result = object

    def run(self):
        self.functions()


def Add_Thread(function):
    thread = myThread(function)
    thread.setDaemon(True)
    thread.start()
    return thread

def DownLoad(dbnum, dbsize, size):
    global download_ProgressValue

    '''''回调函数 
    dbnum: 已经下载的数据块 
    dbsize: 数据块的大小 
    size: 远程文件的大小 
    '''
    percent = 100.0 * dbnum * dbsize / size
    if percent > 100:
        percent = 100

    download_ProgressValue.set(percent)

def loadview():
    global download_ProgressValue
    download_windows = tkinter.Tk()
    download_windows.title("下载进度...")
    download_ProgressValue = DoubleVar()
    download_ProgressValue.set(0.0)
    ttk.Progressbar(download_windows, 
                                        orient="horizontal",
                                        length=300,
                                        mode="determinate",
                                        variable=download_ProgressValue).grid(column=1,
                                                                                                                              row=1,
                                                                                                                              sticky=W,
                                                                                                                              columnspan=1)
    download_windows.withdraw()
    info = GetInfoFromFile()
    if info != None:
        Add_Thread(lambda:Check_Update(info[1],info[2], download_windows, DownLoad))
    else:
        tkinter.messagebox.showinfo('提示！','获取版本信息错误,请重试！')
        os._exit(0)
    download_windows.mainloop()

def GetInfoFromFile():
    file_open =open(r"C:\UpdateInfo.ini","r")
    info = file_open.readlines()
    file_open.close()
    if len(info) > 0:
        info_tuple = info[0].split('-')
    else:
        return False

    if len(info_tuple) > 2:
        return info_tuple
    else:
        return False

def Check_Update(appname, version, progress, callbackfunc):
    pythoncom.CoInitialize()
    if subprocess.call('ping 130.130.200.30 -w 100', shell=True) == 0:    
        check_ini_Md5 = MF.Start(appname, version, Default_Params["UserName"], Default_Params["UserPassword"], Default_Params["Method"], Default_Params["TargetHOST"], Default_Params["LoginURL"], Default_Params["MailreceiveURL"], progress, callbackfunc, Default_Params["TimeOut"], Default_Params["KeywordArray"])
    else:
        tkinter.messagebox.showinfo("提示！","未能连接更新服务器！")
        os._exit(0)

if __name__ == '__main__':
    '''
    print(os.popen('tasklist'))
    os.system('taskkill /IM auto_input.exe /F')

    '''
    loadview()
