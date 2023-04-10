# -*- coding: utf-8 -*-
import socket
import requests
from datetime import datetime, timedelta

def TWtime(): #獲取台灣時間
    import pytz
    import ntplib
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
    import time    
    web_status = 0 #程式調整的網路狀態碼
    web_testtime = 0 #測試網路連線的次數
    while web_status != 200 or web_testtime != 3: #當網路狀態碼等於200或測試次數達到3次時，結束迴圈
        headers = {"User-Agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:99.0) Gecko/20100101 Firefox/99.0"} #設置http頭欄位，裡面夾帶瀏覽器識別標籤
        Go_to_web = requests.get(web_URL, headers = headers, timeout = 60, allow_redirects = False, stream = True, verify = False) #對web_URL夾帶headers發出GET請求，timeout為最長反應時間，allow_redirects為禁止重新定向，stream為強制解壓縮，verify為SSL憑證檢查功能       
        requests.packages.urllib3.disable_warnings() #關閉InsecureRequestWarning的顯示
        if Go_to_web.status_code != 200: #如果網路狀態碼不等於200
            time.sleep(30) #停頓30秒
            web_testtime = web_testtime + 1 #測試網路連線的次數加1
        else: #等於200
            web_status = Go_to_web.status_code 
            web_testtime = web_testtime + 1 #測試網路連線的次數加1
        Go_to_web.close() #關閉對web_URL夾帶headers發出GET請求
        if web_status == 200 or web_testtime == 3: #如果網路正常或是測試網路連線3次
            return web_status #回傳程式調整的網路狀態碼
        else:
            return web_status #回傳程式調整的網路狀態碼

def get_followday(number, start_time, i, status, startfollowdate, followday, string, endfollowdate, worksheet): #取得追蹤天數
    if status == "old": #如果狀態為old(老追蹤者)
        try:
            startfollowdate = datetime.strptime(startfollowdate, "%Y-%m-%d %H:%M:%S") #將開始追蹤時間轉成datetime格式
            today_date = datetime.strptime(str(start_time).split(".")[0], "%Y-%m-%d %H:%M:%S") #將程式開始執行的時間start_time取出年月日
            followday_out = (str(today_date-startfollowdate).split(" days,")[0]+"日 ")+(str(today_date-startfollowdate).split("days, ")[1].split(":")[0])+"時"+(str(today_date-startfollowdate).split("days, ")[1].split(":")[1])+"分"+(str(today_date-startfollowdate).split("days, ")[1].split(":")[2])+"秒" #時間相減等於追蹤時間
        except:
            startfollowdate = "{}{}".format(string, i) #來源字串(gamer/plurk;friend/followerfan)加編號
            if str(start_time).split(" ")[0] == str(worksheet.get_value("E{}".format(number+2))): #start_time的日期等於傳入工作表（追蹤名單）("D{}".format(number+1))的字串
                followday_out = followday 
            else:
                followday_plus = timedelta(days=int(followday.split("日 ")[0]), seconds=(((int(followday.split("日")[1].split("時")[0])*60)+(int(followday.split("日")[1].split("時")[1].split("分")[0])))*60)+(int(followday.split("日")[1].split("時")[1].split("分")[1].split("秒")[0])))+timedelta(days=1) #追蹤日加1
                followday_out = (str(followday_plus).split(" days,")[0]+"日 ")+(str(followday_plus).split("days, ")[1].split(":")[0])+"時"+(str(followday_plus).split("days, ")[1].split(":")[1])+"分"+(str(followday_plus).split("days, ")[1].split(":")[2])+"秒"
    else: #如果狀態為leave(退追蹤者)
        startfollowdate = datetime.strptime(startfollowdate, "%Y-%m-%d %H:%M:%S") #將開始追蹤時間轉成datetime格式
        end_date = endfollowdate #結束時間
        followday_out = (str(end_date-startfollowdate).split(" days,")[0]+"日 ")+(str(end_date-startfollowdate).split("days, ")[1].split(":")[0])+"時"+(str(end_date-startfollowdate).split("days, ")[1].split(":")[1])+"分"+(str(end_date-startfollowdate).split("days, ")[1].split(":")[2])+"秒" #時間相減等於追蹤時間
    return followday_out #回傳追蹤天數

def lastday_of_month(year, month): #取得該年該月最後一天
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
    import psutil
    nic_list = psutil.net_if_addrs() #取得網路連線清單
    nic_list_status = psutil.net_if_stats() #取得網路連線狀態清單
    for nic in nic_list: #取出其中一個網路
        if (nic_list_status[nic].isup == True) & (len(nic_list[nic]) == 3): #如果該網路正在連線和不是Loopback Pseudo-Interface 1
            nic_data = [nic, {"mac" : nic_list[nic][0].address,"ipv4" : nic_list[nic][1].address, "ipv6" : nic_list[nic][2].address}] #獲取該網路資訊
    return nic_data[0], nic_data[1] #回傳網路名稱(非SSID)和其mac、ipv4和ipv6

def get_user(): #取得本機裝置使用者名稱
    import psutil
    user_list = psutil.users() #取得本機使用者
    users = [] #放置本機使用者名稱格式轉換後的串列
    for i in list(range(0, len(user_list))): #利用本機使用者數量限定範圍
        users.append(user_list[i].name) #抽出本機使用者名稱後加入串列 
    return users, len(users) #回傳本機使用者和其人數

def get_ip(): #取得網際網路(外網)IP
    import json
    get_ip_link = ["http://httpbin.org/ip", "https://ifconfig.me/ip"] #查詢IP網址串列
    i = 0 #IP網址串列順序
    status = 0 #取得IP的狀態碼
    while i < 2 and status != 1: #當順序小於2和狀態碼等於0時，持續運作
        try:
            headers = {"User-Agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36 Edge/18.19577"} #設置http頭欄位，裡面夾帶瀏覽器識別標籤
            go_to_url = requests.get(get_ip_link[i], headers = headers, timeout = 60, allow_redirects = False, stream = True, verify = False) #對get_ip_link[i]的網址夾帶headers發出GET請求，timeout為最長反應時間，allow_redirects為禁止重新定向，stream為強制解壓縮，verify為SSL憑證檢查功能
            if i == 0: #依據網址原始碼做相對應資訊之取出
                ip = json.loads(go_to_url.text)["origin"] #將字串變成python的字典再取出相對的值
                status = 1
            if i == 1: #依據網址原始碼做相對應資訊之取出
                ip = go_to_url.text #直接取出值
                status = 1
            go_to_url.close() #關閉與get_ip_link[i]之網址的連結
        except:
            i = i+1 #順序加1
    if ip == "": #如果上面執行失敗，將""置換成None
        ip = None 
    return ip #回傳外網IP

def type_of_ip(ip): #網際網路通訊協定版本判定(IPv4/IPv6)
    import ipaddress
    try:
        try:
            if ipaddress.IPv4Address(ip).version == 4:
                return "IPv4"
        except:
            if ipaddress.IPv6Address(ip).version == 6:
                return  "IPv6"
    except:
        return "None"

def usingnow_of_ip(ip): #判定IP使用狀態
    try:
        netlink = socket.socket(socket.AF_INET, socket.SOCK_DGRAM) #UDP宣告
        netlink.connect(ip, 80) #指定客戶端串接的ip跟Port
        netlink.getsockname()[0] #指定客戶端的ip跟Port並和伺服器連接
        netlink.close #關閉連線
        return "正在使用"
    except:
        return "非正在使用"

def get_ip_data(ip): #取得網際網路IP的所在地區資料
    import json
    get_ip_data_link = "http://ip-api.com/json/{}?fields=status,message,country,countrycode,region,regionname,city,zip,lat,lon,timezone,isp,org,as,query&lang=zh-CN".format(ip) #查詢IP網址的資訊
    try:
        headers = {"User-Agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36 Edge/18.19577"} #設置http頭欄位，裡面夾帶瀏覽器識別標籤
        go_to_url = requests.get(get_ip_data_link, headers = headers, timeout = 60, allow_redirects = False, stream = True, verify = False) #對get_ip_link[i]的網址夾帶headers發出GET請求，timeout為最長反應時間，allow_redirects為禁止重新定向，stream為強制解壓縮，verify為SSL憑證檢查功能
        ip_data = json.loads(go_to_url.text) #將字串變成python的字典
        go_to_url.close() #關閉
    except: #如果上面執行失敗，執行此區
        ip_data = None
    return ip_data #回傳IP資訊


