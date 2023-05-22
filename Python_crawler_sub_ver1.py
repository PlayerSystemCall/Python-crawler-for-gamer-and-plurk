# -*- coding: utf-8 -*-
import re
import json
import pytz
import time
import ntplib
import psutil
import socket
import requests
import ipaddress
from fake_useragent import UserAgent
from datetime import datetime, timedelta

def TWtime(): #獲取台灣時間
    #import pytz
    #import ntplib
    #from datetime import datetime, timedelta
    timeserver = ["ntp.ntu.edu.tw", "tw.pool.ntp.org", "time.stdtime.gov.tw", "tock.stdtime.gov.tw", "watch.stdtime.gov.tw", "clock.stdtime.gov.tw", "tick.stdtime.gov.tw", "118.163.81.61", "time.windows.com", "time.google.com"] #時間伺服器網址串列
    server_number = 0 #時間伺服器網址編號(0~9)
    status = 0 #設定狀態碼為0
    try:
        while server_number < 10 and status != 1: #在透過NTP獲取時間超過10次或狀態碼等於1時，結束迴圈
            NTPClient = ntplib.NTPClient() #啟用NTP客戶端
            try:
                NTPServer = NTPClient.request(timeserver[server_number]) #連結NTP伺服器
                NTPtime = NTPServer.tx_time #取得今天現在時間戳
                Taipeitime = datetime.fromtimestamp(NTPtime, pytz.timezone("Asia/Taipei")) #取得今天現在的時間，時區設定亞洲台北
                ad_year_today = str(Taipeitime.year)+"年" #今天現在台北時區的西元紀年
                mg_year_today = str(int(Taipeitime.year)-1911)+"年" #今天現在台北時區的民國紀年
                date_today = str(Taipeitime.month)+"月"+str(Taipeitime.day)+"日" #今天現在台北時區的西元日期
                ad_year_yesterday = str((Taipeitime+timedelta(days=-1)).year)+"年" #昨天台北時區的西元紀年
                mg_year_yesterday = str((Taipeitime+timedelta(days=-1)).year-1911)+"年" #昨天台北時區的民國紀年
                date_yesterday = str((Taipeitime+timedelta(days=-1)).month)+"月"+str((Taipeitime+timedelta(days=-1)).day)+"日" #昨天台北時區的西元日期
                status = 1 #設定狀態碼為1
            except BaseException: #如果上面執行失敗，執行此區
                server_number = server_number+1 #時間伺服器網址編號加1
        timeserver_link = timeserver[server_number] #時間伺服器網址
    except BaseException: #如果用NTP網址取得時間失敗，執行此區，原理未明
        Taipeitime = datetime.now(pytz.timezone("Asia/Taipei")) #取得今天現在的時間，時區設定亞洲台北
        ad_year_today = str(Taipeitime.year)+"年" #今天現在台北時區的西元紀年
        mg_year_today = str(int(Taipeitime.year)-1911)+"年" #今天現在台北時區的民國紀年
        date_today = str(Taipeitime.month)+"月"+str(Taipeitime.day)+"日" #今天現在台北時區的西元日期
        ad_year_yesterday = str((Taipeitime+timedelta(days=-1)).year)+"年" #昨天台北時區的西元紀年
        mg_year_yesterday = str((Taipeitime+timedelta(days=-1)).year-1911)+"年" #昨天台北時區的民國紀年
        date_yesterday = str((Taipeitime+timedelta(days=-1)).month)+"月"+str((Taipeitime+timedelta(days=-1)).day)+"日" #昨天台北時區的西元日期
        timeserver_link = "無" #設定時間伺服器網址為無
        server_number = "無" #時間伺服器網址編號為無
    return Taipeitime, ad_year_today, mg_year_today, date_today, ad_year_yesterday, mg_year_yesterday, date_yesterday , timeserver_link, server_number #回傳一堆時間

def go_to_web(web_URL): #測試網路連結的狀態
    #import time
    #import requests
    #from fake_useragent import UserAgent
    web_status = 0 #程式調整的網路狀態碼
    web_testtime = 0 #測試網路連線的次數
    while web_status != 200 and web_testtime != 3: #當網路狀態碼等於200或測試次數達到3次時，結束迴圈
        headers = {"User-Agent" : UserAgent().random} #設置http頭欄位，裡面夾帶瀏覽器識別標籤
        Go_to_web = requests.get(web_URL, headers = headers, timeout = 60, allow_redirects = False, stream = True, verify = False) #對web_URL夾帶headers發出GET請求，timeout為最長反應時間，allow_redirects為禁止重新定向，stream為強制解壓縮，verify為SSL憑證檢查功能       
        requests.packages.urllib3.disable_warnings() #關閉InsecureRequestWarning的顯示
        if Go_to_web.status_code != 200: #如果網路狀態碼不等於200
            time.sleep(30) #停頓30秒
            web_testtime = web_testtime + 1 #測試網路連線的次數加1
        else: #等於200
            web_status = Go_to_web.status_code 
            web_testtime = web_testtime + 1 #測試網路連線的次數加1
        Go_to_web.close() #關閉對web_URL夾帶headers發出GET請求
    return web_status #回傳程式調整的網路狀態碼

def get_followday(number, start_time, i, status, startfollowdate, followday, string, endfollowdate, worksheet): #取得追蹤天數
    #from datetime import datetime, timedelta
    if status == "old": #如果狀態為old(老追蹤者)
        try:
            startfollowdate = datetime.strptime(startfollowdate, "%Y-%m-%d %H:%M:%S") #將開始追蹤時間轉成datetime格式
            today_date = datetime.strptime(str(start_time).split(".")[0], "%Y-%m-%d %H:%M:%S") #將程式開始執行的時間start_time取出年月日
            if "days," in str(today_date-startfollowdate): #時間相減等於追蹤時間
                followday_out = (str(today_date-startfollowdate).split(" days,")[0]+"日")+(str(today_date-startfollowdate).split("days, ")[1].split(":")[0])+"時"+(str(today_date-startfollowdate).split("days, ")[1].split(":")[1])+"分"+(str(today_date-startfollowdate).split("days, ")[1].split(":")[2])+"秒"
            elif "day," in str(today_date-startfollowdate):
                followday_out = (str(today_date-startfollowdate).split(" day,")[0]+"日")+(str(today_date-startfollowdate).split("day, ")[1].split(":")[0])+"時"+(str(today_date-startfollowdate).split("day, ")[1].split(":")[1])+"分"+(str(today_date-startfollowdate).split("day, ")[1].split(":")[2])+"秒"
        except:
            startfollowdate = "{}{}".format(string, i) #來源字串(gamer/plurk;friend/followerfan)加編號
            if str(start_time).split(" ")[0] == str(worksheet.get_value("E{}".format(number+2))): #start_time的日期等於傳入工作表（追蹤名單）("D{}".format(number+1))的字串
                followday_out = followday 
            else:
                followday_out = timedelta(days=int(followday.split("日")[0]), seconds=(((int(followday.split("日")[1].split("時")[0])*60)+(int(followday.split("日")[1].split("時")[1].split("分")[0])))*60)+(int(followday.split("日")[1].split("時")[1].split("分")[1].split("秒")[0])))+timedelta(days=1) #追蹤日加1
                if "days," in str(followday_out): #時間相減等於追蹤時間
                    followday_out = (str(followday_out).split(" days,")[0]+"日")+(str(followday_out).split("days, ")[1].split(":")[0])+"時"+(str(followday_out).split("days, ")[1].split(":")[1])+"分"+(str(followday_out).split("days, ")[1].split(":")[2])+"秒"
                elif "day," in str(followday_out):
                    followday_out = (str(followday_out).split(" day,")[0]+"日")+(str(followday_out).split("day, ")[1].split(":")[0])+"時"+(str(followday_out).split("day, ")[1].split(":")[1])+"分"+(str(followday_out).split("day, ")[1].split(":")[2])+"秒"
    else: #如果狀態為leave(退追蹤者)
        startfollowdate = datetime.strptime(startfollowdate, "%Y-%m-%d %H:%M:%S") #將開始追蹤時間轉成datetime格式
        end_date = endfollowdate #結束時間
        followday_out = (str(end_date-startfollowdate).split(" days,")[0]+"日")+(str(end_date-startfollowdate).split("days, ")[1].split(":")[0])+"時"+(str(end_date-startfollowdate).split("days, ")[1].split(":")[1])+"分"+(str(end_date-startfollowdate).split("days, ")[1].split(":")[2])+"秒" #時間相減等於追蹤時間
    return followday_out #回傳追蹤天數

def lastday_of_month(year, month): #取得該年該月最後一天
    #from datetime import datetime
    try:
        datetime(year, month, 31)
        return 31
    except:
        try:
            datetime(year, month, 30)
            return 30
        except:
            try:
                datetime(year, month, 29)
                return 29
            except:
                datetime(year, month, 28)
                return 28

def get_nic_data(): #取得正在使用的網路卡資料
    #import re
    #import psutil
    #import socket
    ni_list = psutil.net_if_addrs() #取得網路介面清單
    ni_list_status = psutil.net_if_stats() #取得網路介面連線狀態清單
    device_name = socket.getfqdn(socket.gethostname()).split(".")[0] #裝置名稱，從DNS連線網址中擷取第1段
    local_ip = socket.gethostbyname(device_name)
    for ni in ni_list: #取出其中一個網路介面
        if (ni_list_status[ni].isup == True) & (len(ni_list[ni]) == 3): #如果該網路介面正在連線和不是Loopback Pseudo-Interface 1
            ni_data = [ni, {"mac" : ni_list[ni][0].address,"ipv4" : ni_list[ni][1].address, "ipv6" : ni_list[ni][2].address}, {ni_list[ni][0].address : "mac", ni_list[ni][1].address : "ipv4", ni_list[ni][2].address : "ipv6"}] #獲取該網路介面資訊
    if re.findall("\d+.\d+.\d+.\d+", local_ip) != []:
        local_ip_ver = "IPv4"
    else:
        local_ip_ver = "None"
    return ni_data[0], ni_data[1], ni_data[2], local_ip_ver #回傳網路類型(非SSID)、其mac、ipv4和ipv6、正在使用的內網板本

def get_user(): #取得本機裝置使用者名稱
    #import psutil
    user_list = psutil.users() #取得本機使用者
    users = [] #放置本機使用者名稱格式轉換後的串列
    for i in range(0, len(user_list)): #利用本機使用者數量限定範圍
        users.append(user_list[i].name) #抽出本機使用者名稱後加入串列 
    return users, len(users) #回傳本機使用者和其人數

def get_ip_and_version(): #取得網際網路(外網)IP
    #import json
    #import requests
    #import ipaddress
    #from fake_useragent import UserAgent
    get_ip_link = ["http://httpbin.org/ip", "https://ifconfig.me/ip"] #查詢IP網址串列
    i = 0 #IP網址串列順序
    status = 0 #取得IP的狀態碼
    while i < 2 and status != 1: #當順序小於2和狀態碼等於0時，持續運作
        try:
            headers = {"User-Agent" : UserAgent().random} #設置http頭欄位，裡面夾帶瀏覽器識別標籤
            go_to_url = requests.get(get_ip_link[i], headers = headers, timeout = 60, allow_redirects = False, stream = True, verify = False) #對get_ip_link[i]的網址夾帶headers發出GET請求，timeout為最長反應時間，allow_redirects為禁止重新定向，stream為強制解壓縮，verify為SSL憑證檢查功能
            if i == 0: #依據網址原始碼做相對應資訊之取出
                ip = json.loads(go_to_url.text)["origin"] #將字串變成python的字典再取出相對的值
                status = 1
            elif i == 1: #依據網址原始碼做相對應資訊之取出
                ip = go_to_url.text #直接取出值
                status = 1
            go_to_url.close() #關閉與get_ip_link[i]之網址的連結
        except:
            i = i+1 #順序加1
    if ip == "": #如果上面執行失敗，將""置換成None
        ip = None
    try: #判定網際網路通訊協議版本
        try:
            if ipaddress.IPv4Address(str(ip)).version == 4:
                version = "IPv4"
        except:
            if ipaddress.IPv6Address(str(ip)).version == 6:
                version = "IPv6"
    except:
        version = None
    return ip, version #回傳外網IP

def get_ip_data(ip): #取得網際網路IP的所在地區資料
    #import json
    #import requests
    #from fake_useragent import UserAgent
    get_ip_data_link = "http://ip-api.com/json/{}?fields=status,message,country,countrycode,region,regionname,city,zip,lat,lon,timezone,isp,org,as,query&lang=zh-CN".format(ip) #查詢IP網址的資訊
    try:
        headers = {"User-Agent" : UserAgent().random} #設置http頭欄位，裡面夾帶瀏覽器識別標籤
        go_to_url = requests.get(get_ip_data_link, headers = headers, timeout = 60, allow_redirects = False, stream = True, verify = False) #對get_ip_link[i]的網址夾帶headers發出GET請求，timeout為最長反應時間，allow_redirects為禁止重新定向，stream為強制解壓縮，verify為SSL憑證檢查功能
        ip_data = json.loads(go_to_url.text) #將字串變成python的字典
        go_to_url.close() #關閉
    except: #如果上面執行失敗，執行此區
        ip_data = None
    return ip_data #回傳IP資訊