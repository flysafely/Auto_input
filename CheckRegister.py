from oscrypto._win import symmetric
import threading
import pythoncom  # 多线程调用COM
import platform
import os
import time
import datetime
import uuid
import urllib.request
import re
ADD_REG_EXP = "now : '(.*?) "
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
        time = datetime.date(int(object_element_list[0].split("-")[0]), int(
            object_element_list[0].split("-")[1]), int(object_element_list[0].split("-")[2]))

        return time

    except urllib.error.URLError as err:
        return datetime.date.today()


def registration_check(ip,filename,keyvalue):
    #global Server_time, isRegistered
    #registrationcode.ini
    mac_check_path = os.getcwd() + "\\"+filename
    expiry_date=get_Server_time(ip)
    if os.path.exists(mac_check_path):
        registration_dict = {}
        openfile = open(os.getcwd() + "\\"+filename, "r+")
        content_list2 = openfile.readlines()
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

        openfile.close()
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

if __name__ == '__main__':
    #print(registration_check("http://130.130.200.49","registrationcode.ini",b"1234567890123456"))
    pass
    