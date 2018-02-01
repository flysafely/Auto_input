from oscrypto._win import symmetric
import tkinter
from tkinter.filedialog import *
import hashlib
import threading
import pythoncom  # 多线程调用COM
import platform
import os
import time
import datetime
import uuid

import urllib.request
import http.cookiejar 
import re
import hashlib
import MainFunction as MF

ADD_REG_EXP = "now : '(.*?) "

Default_Params = {
    "UserName":"xafei", 
    "UserPassword":"111", 
    "Method":"INI", 
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

    def get_result(self):
        return self.result


def Add_Thread(function):
    thread = myThread(function)
    thread.setDaemon(True)
    thread.start()
    return thread

def get_mac_address():
    mac = uuid.UUID(int=uuid.getnode()).hex[-12:]
    return "-".join([mac[e:e + 2] for e in range(0, 11, 2)])

def get_Server_time(ip):
    global Server_time
    try:
        req = urllib.request.Request(ip)
        html = urllib.request.urlopen(req, timeout=1)
        html_content = html.read().decode('utf8')
        object_element_list = re.findall(ADD_REG_EXP, html_content)
        time = datetime.date(int(object_element_list[0].split("-")[0]), 
                                                     int(object_element_list[0].split("-")[1]), 
                                                     int(object_element_list[0].split("-")[2]))
        return time
    except urllib.error.URLError as err:
        return datetime.date.today()

def GetFileMd5(filepath):
    if not os.path.isfile(filepath):
        return
    myhash = hashlib.md5()
    f = open(filepath,'rb')
    while True:
        b = f.read(8096)
        if not b :
            break
        myhash.update(b)
    f.close()
    return myhash.hexdigest()

def Check_INI_info(timeHost, filepath, keyvalue):

    expiry_date=get_Server_time(timeHost)
    if os.path.exists(filepath):
        registration_dict = {}
        openfile = open(filepath, "r+")
        content_list2 = openfile.readlines()
        openfile.close()
        for j in content_list2:
            text = symmetric.aes_cbc_pkcs7_decrypt(
                keyvalue, eval(j.replace("\n", "")), keyvalue)
            plaintext = text.decode("utf8").replace("\n", "")
            plaintext_list = plaintext.split("&")
            content_dict = {}
            content_dict["MacIp"] = plaintext_list[0]
            content_dict["ExpData"] = plaintext_list[1]
            content_dict["UserName"] = plaintext_list[2]
            content_dict["Company"] = plaintext_list[3]
            content_dict["Department"] = plaintext_list[4]

            registration_dict[plaintext_list[0]] = content_dict

        Local_MacIP = get_mac_address().upper()
        if Local_MacIP in list(registration_dict.keys()):
            date_array = registration_dict[Local_MacIP]["ExpData"].split("-")
            year = int(date_array[0])
            month = int(date_array[1])
            day = int(date_array[2])
            if datetime.date(year, month, day) > expiry_date:
                return [True,registration_dict[Local_MacIP]]
            else:
                return [False,None]
        else:
            return [False,None]
    else:
        return [False,None]


def registration_check(timeHost, appname, Md5, inputview, filename, keyvalue):
    ini_path = os.getcwd() + "\\"+filename
    loacl_ini_Md5 = GetFileMd5(ini_path)
    if Md5 == None:

        check_ini_Md5 = MF.Start(appname, loacl_ini_Md5, Default_Params["UserName"], Default_Params["UserPassword"], Default_Params["Method"], Default_Params["TargetHOST"], Default_Params["LoginURL"], Default_Params["MailreceiveURL"], inputview, None, Default_Params["TimeOut"], Default_Params["KeywordArray"])

        if check_ini_Md5:
            
            result = Check_INI_info(timeHost, ini_path, keyvalue)
            return result

        else:
            print("show windos")
            inputview.deiconify()
    else:

        if Md5 == loacl_ini_Md5:
            result = Check_INI_info(timeHost, ini_path, keyvalue)
            return result

        else:

            tkinter.messagebox.showinfo("提示！","激活码验证错误！重启程序或询问作者获取更新激活信息！")

            os._exit(0)
'''
def Init_Register(timeHost, appname, filename, keyvalue):
    input_window = Toplevel()
    input_window.title("输入验证码")
    input_window.geometry("282x24+%s+%s" % (input_window.winfo_screenwidth() // 2 - 140, input_window.winfo_screenheight() // 2 - 200))
    input_StrVar = StringVar()
    Entry(input_window, font='微软雅黑 -10',
                  width=38,
                  textvariable=input_StrVar,
                  justify=LEFT).grid(column=1,
                                     row=1,
                                     sticky=N + S + E + W)
    Button(input_window, text="验证",
       width=7,
       font='微软雅黑 -9 bold',
       command = lambda:registration_check(timeHost, appname, input_StrVar.get(), input_window, filename, keyvalue)).grid(column=2,
                            row=1,
                            sticky=W)
    input_window.withdraw()
    Add_Thread(lambda:registration_check(timeHost, appname, None, input_window, filename, keyvalue))



    input_window.mainloop()
'''
if __name__ == '__main__':
    pass
    #print(Init_Register("http://130.130.200.49", "sm", "registrationcode.ini",b"1234567890123456"))
    #registration_check("http://130.130.200.49", "sm", None, None, "registrationcode.ini",b"1234567890123456")

    