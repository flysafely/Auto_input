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
import zipfile
import subprocess
import webbrowser

global download_ProgressValue
global cookies,GET_MAIL_HEADER
Name_Mapping = {
    "sm":"空间管理",
    "ae":"费用录入"
}

Path_Mapping = {
    "sm":"Space_Managementer.exe",
    "ae":"auto_input.exe"
}

cookies=""
REG_EXP1 = "<a href=.*?version</a>"
REG_EXP2 = "/PublicFunction.*?</a></td></tr>"
#下载地址路径为http://130.130.200.30/applications/email/MailReceiveContent.aspx?ID=********&PageNum=1
main_url="http://130.130.200.30/applications/email/"
#re.findall(ADD_REG_EXP, html_content)
class myThread (threading.Thread):

    def __init__(self, functions):
        threading.Thread.__init__(self)
        self.functions = functions
        self.result = object

    def run(self):
        self.functions()


GET_COOKIE_HEADER = {
    #获取登录cookies
    'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36',
    'Referer': 'http://www.oa.com/kmext/ext/sso.jsp',
    'Host':'130.130.200.30',
    'Connection':'keep-alive',
    'Upgrade-Insecure-Requests':'Upgrade-Insecure-Requests',
    'Accept-Encoding':'',
    'Accept-Language':'zh-CN,zh;q=0.8',
    'Cookie':'EIIS=Login=xafei&ThemePath=/DesktopTheme/Blue1'
}

def loadview():
    global download_ProgressValue
    download_windows = tkinter.Tk()
    #ico = os.getcwd() + r'\sm.ico'
    download_windows.title("下载进度...")
    #download_windows.iconbitmap(ico)
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
        Add_thread(lambda:check_update("130.130.200.30", info[0],info[1], info[2], download_windows, DownLoad))
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
        return None

    if len(info_tuple) > 2:
        return info_tuple
    else:
        return None

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

def Add_thread(function):
    thread = myThread(function)
    thread.setDaemon(True)
    thread.start()
    return thread

def get_login_cookie(website):
    global cookies
    req = urllib.request.Request(website,headers=GET_COOKIE_HEADER)
    cj = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))

    Add_thread(lambda:opener.open(req,timeout=1))
    Add_thread(lambda:get_login_info(cj))
    save_time=time.time()
    while time.time()-save_time<5:
        if cookies!="":
            set_header_with_login_cookie(cookies)
            break

def set_header_with_login_cookie(cookies):
    global GET_MAIL_HEADER

    GET_MAIL_HEADER = {
    #获取登录cookies
    'Host': '130.130.200.30',
    'Connection': 'keep-alive',
    'Upgrade-Insecure-Requests': '1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
    'Referer': 'http://130.130.200.30/applications/email/mailreceive.aspx',
    'Accept-Encoding':'',
    'Accept-Language': 'zh-CN,zh;q=0.8',
    'Cookie': 'ASP.NET_SessionId='+cookies+'; EIIS=Login=xafei&ThemePath=/DesktopTheme/Blue1'
    }

def open_url_without_cookie(website,name,version):
    global GET_MAIL_HEADER
    try:
        req = urllib.request.Request(website,headers=GET_MAIL_HEADER)
        html = urllib.request.urlopen(req)
    except urllib.error.HTTPError as err:
        print(err.msg)
        time.sleep(1)
        tkinter.messagebox.showinfo("警告！","连接更新服务器错误！请确认‘强制登录’后重试！")
        refrushlogin()
    else:
        download_page_id=[]
        software_name=[]
        software_version=[]
        content=html.read().decode('utf8')
        for i in re.findall(REG_EXP1,content):
            download_page_id.append(i.split("\"")[1])
            software_name.append(i.split("\"")[2].split(":")[1])
            software_version.append(i.split("\"")[2].split(":")[2])
        if name in software_name:
            index_num=software_name.index(name)
            if float(version) < float(software_version[index_num]):
                return [download_page_id[index_num],software_name[index_num],software_version[index_num]]
            else:
                return 0
        else:
            return 0

def get_download_file_with_login(website):
    global GET_MAIL_HEADER
    try:
        req = urllib.request.Request(website,headers=GET_MAIL_HEADER)
        html = urllib.request.urlopen(req)
        print(html)
    except urllib.error.HTTPError as err:
        print(err.msg)
        time.sleep(1)
        get_download_file_with_login(website)
    else:    
        content=html.read().decode('utf8')

        download_info=str(re.findall(REG_EXP2,content)).split("\"")[0]

        download_url_without_host=download_info.replace("['","")
        if len(download_url_without_host)>0:
            return download_url_without_host
        else:
            return 0

def get_login_info(cj):
    global cookies
    save_time=time.time()
    while time.time()-save_time<5:
        if len(list(enumerate(cj)))!=0:
            start_index=str(list(cj)[0]).split(" ")[1].index("=")+1
            cookies=str(list(cj)[0]).split(" ")[1][start_index:]
            break

def download_file(website, path, name, version, progress, callbackfunc):
    global GET_MAIL_HEADER
    session='ASP.NET_SessionId='+ cookies +'; EIIS=Login=xafei&ThemePath=/DesktopTheme/Blue1'
    try:
        opener=urllib.request.build_opener()
        opener.addheaders=[('Host', '130.130.200.30'),
                                                    ('Connection', 'keep-alive'),
                                                    ('Upgrade-Insecure-Requests', '1'),
                                                    ('User-Agent', 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36'),
                                                    ('Accept', 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8'),
                                                    ('Referer', 'http://130.130.200.30/applications/email/mailreceive.aspx'),
                                                    ('Accept-Encoding',''),
                                                    ('Accept-Language', 'zh-CN,zh;q=0.8'),
                                                    ('Cookie', session)]
        urllib.request.install_opener(opener)
        filename = get_desktop() + "\\" + name + version +"-setup.zip"
        urllib.request.urlretrieve(website, filename, callbackfunc)
    except urllib.error.HTTPError as err:
        tkinter.messagebox.showinfo("警告！",website)
    else:    
        progress.withdraw()
        if tkinter.messagebox.askyesno("提示！","下载完成！是否立刻开始安装？"):
            UnZip_File(filename, path, name, version, progress, callbackfunc)
            os._exit(0)
        else:
            os.remove(filename)
            os._exit(0)

def UnZip_File(file, path,name, version, progress, callbackfunc):
    #Choosed_dir = tkinter.filedialog.askdirectory().replace("/", "\\")
    f = zipfile.ZipFile(file, 'r')
    size = len(f.namelist())
    n=0
    progress.deiconify()
    for file_obj in f.namelist():
        n+=1
        callbackfunc(n, 1, size)
        unzipfile = f.extract(file_obj, path)
    f.close()
    progress.withdraw()
    Create_Shortcut(name, version, path)
    os.remove(file)
    os._exit(0)

def Create_Shortcut(name, version, path):
    destDir = winshell.desktop()  
    filename = Name_Mapping[name] + "-" + version 
    target = path + '\\' + Path_Mapping[name]
    winshell.CreateShortcut(  
                            Path = os.path.join(destDir, os.path.basename(filename)+".lnk"),  
                            Target = target,  
                            StartIn = str(os.path.dirname(target)),
                            Icon = (target, 0),  
                            Description = "")  

def check_update(host, path, name, version, progress, callbackfunc):
    pythoncom.CoInitialize()
    if subprocess.call('ping 130.130.200.30 -w 100', shell=True) == 0:    
        get_login_cookie("http://"+host+"/loginA.aspx?login=xafei&pass=111")
        version_info = open_url_without_cookie("http://"+host+"/applications/email/mailreceive.aspx",name,version)
        if version_info != 0:
            download_id = get_download_file_with_login(main_url+version_info[0])
            if download_id != 0:
                if tkinter.messagebox.askyesno("提示！","检查到有更新版(%s) 是否立刻开始下载？" % (version_info[2])):
                    subprocess.call('TASKKILL /F /IM '+ Path_Mapping[name], shell=True)
                    progress.deiconify()
                    download_file("http://" + host + download_id, path, version_info[1], version_info[2], progress, callbackfunc)
                else:
                    os._exit(0)
        else:
            tkinter.messagebox.showinfo("提示！","%s版本已经是最新版！" % version)
            os._exit(0)
    else:
        tkinter.messagebox.showinfo("提示！","未能连接更新服务器！")
        os._exit(0)

def get_desktop(): #获取桌面路径 最后没有/
    key =win32api.RegOpenKey(win32con.HKEY_CURRENT_USER,r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders',0,win32con.KEY_READ)  
    return win32api.RegQueryValueEx(key,'Desktop')[0]  

def refrushlogin():
    webbrowser.open_new("http://130.130.200.30/loginA.aspx?login=xafei&pass=111")


if __name__ == '__main__':

    loadview()