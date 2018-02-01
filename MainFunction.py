from tkinter import *
from tkinter import ttk
from tkinter.filedialog import *
from tkinter.messagebox import *
from PIL import Image, ImageTk
from oscrypto._win import symmetric

import tkinter
import urllib.request
import re
import os
import http.cookiejar 
import win32com.client
import win32api,win32con 
import threading
import pythoncom  # 多线程调用COM
import time
import winshell
import zipfile
import subprocess
import psutil
import webbrowser
import hashlib
import platform
import datetime
import uuid
global download_ProgressValue
global keyword,cookies,Request_Header
Name_Mapping = {
    "sm":"空间管理",
    "ae":"费用录入"
}

Path_Mapping = {
    "sm":"Space_Managementer.exe",
    "ae":"auto_input.exe"
}

Code_Mapping = {
    "sm":"registrationcode",
    "ae":"input-registrationcode"
}

Keyword_Mapping = {
    "INI":"key",
    "ZIP":"version"
}

REG_EXP2 = "/PublicFunction.*?</a></td></tr>"

Email_URL="http://130.130.200.30/applications/email/"

default_cookie = 'EIIS=Login=xafei&ThemePath=/DesktopTheme/Blue1'

Request_Header = {
    'Host': None,
    'Connection': 'keep-alive',
    'Upgrade-Insecure-Requests': '1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36',
    'Referer': None,
    'Accept-Encoding':'',
    'Accept-Language': 'zh-CN,zh;q=0.8',
    'Cookie': None    
}
'''
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

    Add_thread(lambda:Start("sm","3.0","xafei","111",
                                                            "130.130.200.30",
                                                            "http://130.130.200.49/kmext/ext/sso.jsp", 
                                                            "http://130.130.200.30/applications/email/mailreceive.aspx", 
                                                            download_windows, 
                                                            DownLoad,
                                                            5, 
                                                            [" ","."]))
    download_windows.withdraw()
    download_windows.mainloop()
'''
def Set_Header(host, referer, cookie):
    global Request_Header
    Request_Header['Host'] = host
    Request_Header['Referer'] = referer
    Request_Header['Cookie'] = cookie
    return Request_Header

class AppBody(object):

    def __init__(self, *args, **kwargs):
        object.__init__(self)
        self.AppNAME = args[0]
        self.AppINFO = args[1]
        self.AppePATH = args[2]
        self.UserName = args[3]
        self.UserPassWord = args[4]
        self.targetHOST = kwargs['targetHOST']
        self.loginURL = kwargs['loginURL']
        self.mailreceiveURL = kwargs['mailreceiveURL']
        self.targetURL = "http://" + self.targetHOST + "/loginA.aspx?login=" + self.UserName + "&pass=" + self.UserPassWord


class myThread (threading.Thread):

    def __init__(self, functions):
        threading.Thread.__init__(self)
        self.functions = functions
        self.result = object

    def run(self):
        self.functions()

def Add_thread(function):
    thread = myThread(function)
    thread.setDaemon(True)
    thread.start()
    return thread

def Get_Cookie(URL, host, mailreceiveURL, refererURL, cookie, timeout = 5):
    global Request_Header
    req = urllib.request.Request(URL,headers=Set_Header(host, refererURL, cookie))
    cj = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
    Add_thread(lambda:opener.open(req,timeout = timeout))
    result = Extract_Cookie(host, mailreceiveURL, cj, timeout)
    return result

def Extract_Cookie(host, mailreceiveURL, cookiejar, timeout):
    isSuccessed = False
    save_time=time.time()
    while time.time()-save_time < timeout:    
        if len(list(enumerate(cookiejar)))!= 0:
            cookie_str = str(list(cookiejar)[0]).split(" ")[1]
            start_index=cookie_str.index("=") + 1
            SessionId =cookie_str[start_index:]
            Set_Header(host, 
                                     mailreceiveURL, 
                                     'ASP.NET_SessionId=' + SessionId + '; EIIS=Login=xafei&ThemePath=/DesktopTheme/Blue1')
            isSuccessed = True
            break        
    return isSuccessed

def Open_MailreceiveURL(URL, name, info, Method, IngoreArray):
    global Request_Header
    try:
        req = urllib.request.Request(URL,headers = Request_Header)
        html = urllib.request.urlopen(req)
    except urllib.error.HTTPError as err:
        print(err.msg)
        time.sleep(1)
        tkinter.messagebox.showinfo("警告！","连接更新服务器错误！请确认‘强制登录’后重试！")
        Refrush_Login()
    else:
        download_IDs=[]
        web_app_names=[]
        web_app_infos=[]
        content=html.read().decode('utf8')
        REG_EXP_ZIP = "<a href=.*?%s</a>" % Keyword_Mapping[Method]
        REG_EXP_INI_1 = "%s.*%s</font>" % (Keyword_Mapping[Method], Keyword_Mapping[Method])
        REG_EXP_INI_2 = "<a href=.*>%s" % Keyword_Mapping[Method]

        if Method == "ZIP" and len(re.findall(REG_EXP_ZIP, content)) > 0:
            reg_content = re.findall(REG_EXP_ZIP, content)
            for i in reg_content:
                download_IDs.append(i.split("\"")[1])
                web_app_names.append(i.split("\"")[2].split(":")[1])
                web_app_infos.append(i.split("\"")[2].split(":")[2])
        elif Method == "INI" and len(re.findall(REG_EXP_INI_1, content)) > 0:
            reg_content_1 = re.findall(REG_EXP_INI_1, content)
            reg_content_2 = re.findall(REG_EXP_INI_2, content)
            for i in reg_content_1:
                web_app_names.append(i.split(":")[1])
                web_app_infos.append(i.split(":")[2])
            for i in reg_content_2:
                download_IDs.append(i.split("\"")[1])
        else:
            return False

        if name in web_app_names:
            index_num=web_app_names.index(name)
            if Compare_Info(info, web_app_infos[index_num], Method, IngoreArray):
                print([download_IDs[index_num], web_app_names[index_num], web_app_infos[index_num]])
                return [download_IDs[index_num], web_app_names[index_num], web_app_infos[index_num]]
            else:
                print("相同")
                return False
        else:
            print("未找到",name)
            return False
            
def Get_Download_URL(host, id_info):
    global Request_Header
    URL = "http://"+ host + "/applications/email/" + id_info
    try:
        req = urllib.request.Request(URL,headers=Request_Header)
        html = urllib.request.urlopen(req)
    except urllib.error.HTTPError as err:
        print(err.msg)
        time.sleep(1)
        Get_Download_URL(URL)
    else:    
        content=html.read().decode('utf8')

        download_info=str(re.findall(REG_EXP2,content)).split("\"")[0]

        download_url_without_host = download_info.replace("['","")
        if len(download_url_without_host) > 0:
            return  "http://"+ host + download_url_without_host
        else:
            return 0

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

def Download_File(URL, name, info, Method, progress, callbackfunc):
    global Request_Header
    if progress != None:
        progress.deiconify()
    try:
        opener=urllib.request.build_opener()
        opener.addheaders=[('Host', Request_Header["Host"]),
                            ('Connection', Request_Header["Connection"]),
                            ('Upgrade-Insecure-Requests', Request_Header["Upgrade-Insecure-Requests"]),
                            ('User-Agent', Request_Header["User-Agent"]),
                            ('Referer', Request_Header["Referer"]),
                            ('Accept-Encoding',Request_Header["Accept-Encoding"]),
                            ('Accept-Language', Request_Header["Accept-Language"]),
                            ('Cookie', Request_Header["Cookie"])]
        urllib.request.install_opener(opener)
        if Method == "ZIP":
            filename = get_desktop() + "\\" + name + info +".zip"
            urllib.request.urlretrieve(URL, filename, callbackfunc)
            progress.withdraw()
            if tkinter.messagebox.askyesno("提示！","下载完成！是否立刻开始安装？"):
                KillProcessByName(Path_Mapping[name])
                setup_path = GetInfoFromFile()
                if setup_path:
                    UnZip_File(filename, setup_path[0], name, info, progress, callbackfunc)
                else:
                    UnZip_File(filename, setup_path, name, info, progress, callbackfunc)
        else:
            filename = os.getcwd() + "\\" + Code_Mapping[name] + ".ini"
            urllib.request.urlretrieve(URL, filename, None)
            tkinter.messagebox.showinfo("提示！", "激活信息已更新！")
    except urllib.error.HTTPError as err:
        tkinter.messagebox.showinfo("警告！", err.msg)
    else:    
        if progress != None:
            progress.deiconify()

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

def UnZip_File(file, path, name, version, progress, callbackfunc):
    if path:
        setup_path = path
    else:
        tkinter.messagebox.showinfo("警告！", "未获取到程序安装位置！请手动指定安装路径！")
        setup_path = tkinter.filedialog.askdirectory().replace("/", "\\")
    f = zipfile.ZipFile(file, 'r')
    size = len(f.namelist())
    n=0
    if progress != None:
        progress.deiconify()

    for file_obj in f.namelist():
        n+=1
        callbackfunc(n, 1, size)
        unzipfile = f.extract(file_obj, setup_path)
    f.close()
    if progress != None:
        progress.withdraw()
    Create_Shortcut(name, version, setup_path)
    os.remove(file)
    os._exit(0)

def Refrush_Login():
    webbrowser.open_new("http://130.130.200.30/loginA.aspx?login=xafei&pass=111")

def GetFileMd5(filename):
    if not os.path.isfile(filename):
        return
    myhash = hashlib.md5()
    f = open(filename,'rb')
    while True:
        b = f.read(8096)
        if not b :
            break
        myhash.update(b)
    f.close()
    return myhash.hexdigest()

def Get_download_Info():
    pass

def CleanStr(Str, StrArray):
    if len(StrArray) > 0:
        for s in StrArray:
            newString = Str.replace(s,"")
        return newString
    else:
        return Str

def Compare_Info(AppInfo, WebInfo, Method, IngoreArray):
    if Method == "ZIP":
        if int(CleanStr(AppInfo,IngoreArray)) < int(CleanStr(WebInfo,IngoreArray)):
            return 1
        else:
            return 0
    else:
        if AppInfo != WebInfo:
            return 1
        else:
            return 0
 
def get_desktop(): #获取桌面路径 最后没有/
    key = win32api.RegOpenKey(win32con.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders', 0, win32con.KEY_READ)  
    return win32api.RegQueryValueEx(key, 'Desktop')[0]

def Download_ZIP(App, Method, keywordArray, progress, callbackfunc, timeout):
    if Get_Cookie(App.targetURL, App.targetHOST, App.loginURL, default_cookie, timeout):

        Download_Info_Array = Open_MailreceiveURL(App.mailreceiveURL, App.AppNAME, App.AppINFO, Method, keywordArray)

        if Download_Info_Array:

            Download_URL = Get_Download_URL(App.targetHOST, Download_Info_Array[0])

            if Download_URL:

                if tkinter.messagebox.askyesno("提示！","检查到有更新版(%s) 是否立刻开始下载？" % Download_Info_Array[2]):

                    Download_File(Download_URL, App.AppNAME, App.AppINFO, Method, progress, callbackfunc)
                else:
                    os._exit(0)
            else:

                tkinter.messagebox.showinfo("提示！","%s版本下载链接错误！请联系作者！" % App.AppINFO)

                os._exit(0)

        else:
            tkinter.messagebox.showinfo("提示！","%s版本已经是最新版！" % App.AppINFO)

            os._exit(0)

    else:

        Refrush_Login()

def Download_INI(App, Method, keywordArray, registration_check, callbackfunc,timeout):

    if Get_Cookie(App.targetURL, App.targetHOST, App.loginURL, default_cookie, timeout):

        Download_Info_Array = Open_MailreceiveURL(App.mailreceiveURL, App.AppNAME, App.AppINFO, Method, keywordArray)

        if Download_Info_Array:

            print("Download_Info_Array")
            Download_URL = Get_Download_URL(App.targetHOST, Download_Info_Array[0])

            if Download_URL[1] != App.AppINFO:
                print("Download_File")
                Download_File(Download_URL, App.AppNAME, App.AppINFO, Method, None, None)

                return True

            else:

                return True

        else:
            print("相同返回")
            return True

    else:

        return False

def KillProcessByName(Pname):
    TargetPID = []
    pids = psutil.pids()
    for pid in pids:
        p = psutil.Process(pid)
        if  p.name() ==  Pname:
            TargetPID.append(pid)
    for pid in TargetPID:
        cmd = "TSKILL " + str(pid)
        subprocess.call(cmd, shell=True)

def Start(appname, appinfo, username, userpassword, Method, targetHOST, loginURL, mailreceiveURL, progress, callbackfunc,timeout, keywordArray):

    App = AppBody(appname, appinfo, None, username, userpassword, None, targetHOST = targetHOST, loginURL = loginURL, mailreceiveURL= mailreceiveURL)

    Switch_Method = {"ZIP":Download_ZIP, "INI":Download_INI}

    result = Switch_Method[Method](App, Method,  keywordArray, progress, callbackfunc, timeout)

    return result

