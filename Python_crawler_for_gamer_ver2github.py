# -*- coding: utf-8 -*-
import os
import re
import socket
import requests
import pygsheets
from lxml import etree
from dotenv import load_dotenv
from datetime import datetime, timedelta

Player_SystemCall_gamer_id = os.getenv('PLAYER_SYSTEMCALL_GAMER_ID')

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

#取得起始時間
start_timecall = TWtime() #呼叫副程式TWtime()獲取時間
start_time = start_timecall[0] #取得程式開始的時間
ad_year_today = start_timecall[1] #台北時區的西元紀年
mg_year_today = start_timecall[2] #台北時區的民國紀年
date_today = start_timecall[3] #今天現在台北時區的西元日期
ad_year_yesterday = start_timecall[4] #台北時區的西元紀年
mg_year_yesterday = start_timecall[5] #台北時區的民國紀年
date_yesterday = start_timecall[6] #昨天台北時區的西元日期

try:    
    #取得裝置資訊
    device_name = socket.getfqdn(socket.gethostname()).split(".")[0] #裝置名稱，從DNS連線網址中擷取第1段
    device_user = get_user() #裝置使用者名稱和數量
    mac_adderss = get_nic_data()[1]["mac"] #裝置網路卡號碼
    net_name = get_nic_data()[0] #裝置網路卡名稱
    device_addr_IPv4 = get_nic_data()[1]["ipv4"] #IPv4位址(內網IP)
    device_addr_IPv6 = get_nic_data()[1]["ipv6"] #IPv6位址(內網IP)
    ip = get_ip() #取得裝置對外的IP
    if ip != None:
        ip_data = get_ip_data(ip) #查詢IP所在地資料
    code_section_1_status = "〇"
except:
    code_section_1_status = "✕"

try:
    #取得巴哈小屋網頁原始碼，找到資料的的html區塊，取出資料
    gamer_url = "https://home.gamer.com.tw/homeindex.php?owner={}".format(Player_SystemCall_gamer_id) #巴哈小屋網址
    gamer_statuscode = go_to_web(gamer_url) #回傳網路狀態碼
    headers = {"User-Agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36 Edge/18.19577"} #設置http頭欄位，裡面夾帶瀏覽器識別標籤
    Go_to_gamer = requests.get(gamer_url, headers = headers, timeout = 60, allow_redirects = False, stream = True, verify = False) #對gamer_url夾帶headers發出GET請求，timeout為最長反應時間，allow_redirects為禁止重新定向，stream為強制解壓縮，verify為SSL憑證檢查功能       
    requests.packages.urllib3.disable_warnings() #關閉InsecureRequestWarning的顯示
    if gamer_statuscode == 200: #如果網路正常
        Go_to_gamer.encoding = "utf-8" #指定網頁的編碼格式
        gamer_sourcecode = etree.HTML(Go_to_gamer.text) #取得網頁原始碼
        #小屋/小屋統計/圖表資料
        xpath_on_web = "//div[@id='BH-background']/div[@id='BH-wrapper']/div[@id='BH-slave']/div[@class='BH-rbox MSG-list1']/script/text()" #text()在網頁程式碼的位置(Xpath表達式)
        jscode_on_web = gamer_sourcecode.xpath(xpath_on_web)[0] #使用Xpath表達式提出，為results_on_web
        gamer_viewnumbers = re.findall('"\d\d"', jscode_on_web) #使用正則表達式提出，為gamer_viewnumbers串列
        for i in range(len(gamer_viewnumbers)):
            gamer_viewnumbers[i] = re.findall("\d\d", gamer_viewnumbers[i])[0]
        gamer_dates = re.findall("\d\d\d\d..\d\d..\d\d", jscode_on_web)
        for i in range(len(gamer_dates)): #將7個日期全部由字串轉成datetime格式
            gamer_dates[i] = datetime.strptime(gamer_dates[i], "%Y\/%m\/%d")
        #小屋/小屋統計/好友圈人數
        xpath_on_web = "//div[@id='BH-background']/div[@id='BH-wrapper']/div[@id='BH-slave']/div[@class='BH-rbox MSG-list1']/ul/li[5]/text()" #text()在網頁程式碼的位置(Xpath表達式)
        gamer_friend_number = int(gamer_sourcecode.xpath(xpath_on_web)[0][4:]) #指定第1個元素的第5個元素之後的字串，並數字化 #取得好友人數
        #小屋/小屋統計/追蹤者人數
        xpath_on_web = "//div[@id='BH-background']/div[@id='BH-wrapper']/div[@id='BH-slave']/div[@class='BH-rbox MSG-list1']/ul/li[6]/text()" #text()在網頁程式碼的位置(Xpath表達式)
        gamer_follower_number = int(gamer_sourcecode.xpath(xpath_on_web)[0][4:]) #指定第1個元素的第5個元素之後的字串，並數字化 #取得粉絲人數
        Go_to_gamer.close() #關閉對gamer_url夾帶headers發出GET請求
    code_section_2_status = "〇"
except:
    code_section_2_status = "✕"

gamer_data_get = {"follower" : {}, "friend" : {}, "other" : {}} #建立字典，並在裡面建立follower、friend和other鍵，功能為放置從巴哈姆特取得(昨天)的資料
    #account為巴哈帳戶，nickname為名字，regdate巴哈帳號建立日，lastondate巴哈最後登入日 #"account" : [], "nickname" :[], "regdate" : [], "lastondate" : []

try:
    #取得巴哈追蹤者名單網頁原始碼，找到資料的的html區塊，取出資料
    gamer_followerlist_url = "https://home.gamer.com.tw/friendList.php?user={}&t=4".format(Player_SystemCall_gamer_id) #巴哈追蹤者列表網址
    gamer_followerlist_statuscode = go_to_web(gamer_followerlist_url) #回傳網路狀態碼
    headers = {"User-Agent" : "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36"} #設置http頭欄位，裡面夾帶瀏覽器識別標籤
    Go_to_gamer_followerlist = requests.get(gamer_followerlist_url, headers = headers, timeout = 60, allow_redirects = False, stream = True, verify = False) #對gamer_follower_url夾帶headers發出GET請求，timeout為最長反應時間，allow_redirects為禁止重新定向，stream為強制解壓縮，verify為SSL憑證檢查功能       
    requests.packages.urllib3.disable_warnings() #關閉InsecureRequestWarning的顯示
    if gamer_followerlist_statuscode == 200: #如果網路正常
        Go_to_gamer_followerlist.encoding = "utf-8" #指定網頁的編碼格式
        gamer_followerlist_sourcecode = etree.HTML(Go_to_gamer_followerlist.text) #取得網頁原始碼
        xpath_on_web = "/html/body/div[@class='friendlist_container']/div[@class='friendlist_wrapper']/div[@class='sort_content for-ios']/div[@class='user_list']/div[@class='user_friend']/div[@class='user_column']/div[@class='user_name']/div[@class='user_id']/text()" #text()在網頁程式碼的位置(Xpath表達式) #帳號
        gamer_data_get["other"]["gamer_follower_accountlist"] = list(map(str, gamer_followerlist_sourcecode.xpath(xpath_on_web))) #使用Xpath表達式提出，為巴哈追蹤者帳號account清單
        xpath_on_web = "/html/body/div[@class='friendlist_container']/div[@class='friendlist_wrapper']/div[@class='sort_content for-ios']/div[@class='user_list']/div[@class='user_friend']/div[@class='user_column']/div[@class='user_name']/div[@class='nickname']/text()" #text()在網頁程式碼的位置(Xpath表達式) #暱稱
        gamer_data_get["other"]["gamer_follower_nicknamelist"] = list(map(str, gamer_followerlist_sourcecode.xpath(xpath_on_web))) #使用Xpath表達式提出，為巴哈追蹤者暱稱nickname清單
        Go_to_gamer_followerlist.close() #關閉對gamer_followerlist_url夾帶headers發出GET請求
    
    for i in range(len(gamer_data_get["other"]["gamer_follower_accountlist"])): #依照順序數字建立字典，再將資訊分別塞進去
        gamer_data_get["follower"][i] = {}
        gamer_data_get["follower"][i]["account"] = gamer_data_get["other"]["gamer_follower_accountlist"][i]
        gamer_data_get["follower"][i]["nickname"] = gamer_data_get["other"]["gamer_follower_nicknamelist"][i]
        try:
            try:
                gamer_follower_url = "https://home.gamer.com.tw/homeindex.php?owner={}".format(gamer_data_get["follower"][i]["account"]) #組合成追蹤者的舊版巴哈小屋連結
                gamer_follower_statuscode = go_to_web(gamer_follower_url) #回傳網路狀態碼
                version = "Old" #標示為舊版小屋
            except:
                gamer_follower_url = "https://home.gamer.com.tw/profile/index.php?owner={}".format(gamer_data_get["follower"][i]["account"]) #組合成追蹤者的新版巴哈小屋連結
                gamer_follower_statuscode = go_to_web(gamer_follower_url) #回傳網路狀態碼
                version = "New" #標示為新版小屋
        except:
            version = "Delete"      
        headers = {"User-Agent" : "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_9_3) AppleWebKit/537.75.14 (KHTML, like Gecko) Version/7.0.3 Safari/7046A194A"} #設置http頭欄位，裡面夾帶瀏覽器識別標籤
        Go_to_gamer_follower = requests.get(gamer_follower_url, headers = headers, timeout = 60, allow_redirects = False, stream = True, verify = False) #對gamer_follower_url夾帶headers發出GET請求，timeout為最長反應時間，allow_redirects為禁止重新定向，stream為強制解壓縮，verify為SSL憑證檢查功能       
        requests.packages.urllib3.disable_warnings() #關閉InsecureRequestWarning的顯示
        Go_to_gamer_follower.encoding = "utf-8" #指定網頁的編碼格式
        gamer_follower_sourcecode = etree.HTML(Go_to_gamer_follower.text) #取得網頁原始碼
        gamer_title_number = 0 #巴哈小屋大標的序號
        status = False #表示狀態
        if gamer_follower_statuscode == 200 and version == "Old": #如果可以用舊版小屋連線
            gamer_title = list(map(str, gamer_follower_sourcecode.xpath("//div[@id='BH-background']/div[@id='BH-wrapper']/div[@id='BH-slave']/h5/text()"))) #從巴哈小屋的原始碼篩出大標
            while status == False and gamer_title_number < len(gamer_title): #當status等於False和gamer_title_number小於巴哈大標數量時，持續運作
                if "個人紀錄" == gamer_title[gamer_title_number]:
                    xpath_on_web = "//div[@id='BH-background']/div[@id='BH-wrapper']/div[@id='BH-slave']/div[@class='BH-rbox BH-list1']/ul/li[4]/text()" #取得註冊日期
                    gamer_data_get["follower"][i]["regdate"] = str(gamer_follower_sourcecode.xpath(xpath_on_web)[0]).split("：")[1] #使用Xpath表達式提出，將帳號建立日放進regdate鍵的值
                    xpath_on_web = "//div[@id='BH-background']/div[@id='BH-wrapper']/div[@id='BH-slave']/div[@class='BH-rbox BH-list1']/ul/li[5]/text()" #取得最後上站日期
                    gamer_data_get["follower"][i]["lastondate"] = str(gamer_follower_sourcecode.xpath(xpath_on_web)[0]).split("：")[1] #使用Xpath表達式提出，將帳號最後上線日放進lastondate鍵的值
                    status = True
                else:
                    gamer_title_number = gamer_title_number+1 #巴哈小屋大標的序號加1
            if status == False:
                gamer_data_get["follower"][i]["regdate"] = "舊版無區塊" #巴哈姆特模塊化特色所致
                gamer_data_get["follower"][i]["lastondate"] = "舊版無區塊"
        elif gamer_follower_statuscode == 200 and version == "New": #如果可以用新版小屋連線，目前程式還不支援去新版尋找
            gamer_data_get["follower"][i]["regdate"] = "新版不支援"
            gamer_data_get["follower"][i]["lastondate"] = "新版不支援"
        else: #如果都不行
            gamer_data_get["follower"][i]["regdate"] = None
            gamer_data_get["follower"][i]["lastondate"] = None
        Go_to_gamer_follower.close() #關閉對gfol_url夾帶headers發出GET請求
    code_section_3_status = "〇"
except:
    code_section_3_status = "✕"

try:
    #取得巴哈朋友名單網頁原始碼，找到資料的的html區塊，取出資料
    gamer_friendlist_url = "https://home.gamer.com.tw/friendList.php?user={}&t=1".format(Player_SystemCall_gamer_id) #巴哈好友列表網址
    gamer_friendlist_statuscode = go_to_web(gamer_friendlist_url) #回傳網路狀態碼
    headers = {"User-Agent" : "Opera/9.80 (X11; Linux i686; Ubuntu/14.10) Presto/2.12.388 Version/12.16"} #設置http頭欄位，裡面夾帶瀏覽器識別標籤
    Go_to_gamer_friendlist = requests.get(gamer_friendlist_url, headers = headers, timeout = 60, allow_redirects = False, stream = True, verify = False) #對gamerfriend_url夾帶headers發出GET請求，timeout為最長反應時間，allow_redirects為禁止重新定向，stream為強制解壓縮，verify為SSL憑證檢查功能       
    requests.packages.urllib3.disable_warnings() #關閉InsecureRequestWarning的顯示
    if gamer_friendlist_statuscode == 200:
        Go_to_gamer_friendlist.encoding = "utf-8" #指定網頁的編碼格式
        gamer_friend_sourcecode = etree.HTML(Go_to_gamer_friendlist.text) #取得網頁原始碼
        xpath_on_web = "/html/body/div[@class='friendlist_container']/div[@class='friendlist_wrapper']/div[@class='sort_content for-ios']/div[@class='user_list']/div[@class='user_friend']/div[@class='user_column']/div[@class='user_name']/div[@class='user_id']/text()" #text()在網頁程式碼的位置(Xpath表達式) #帳號
        gamer_data_get["other"]["gamer_friend_accountlist"] = list(map(str, gamer_friend_sourcecode.xpath(xpath_on_web))) #使用Xpath表達式提出，為巴哈好友帳號account清單
        xpath_on_web = "/html/body/div[@class='friendlist_container']/div[@class='friendlist_wrapper']/div[@class='sort_content for-ios']/div[@class='user_list']/div[@class='user_friend']/div[@class='user_column']/div[@class='user_name']/div[@class='nickname']/text()" #text()在網頁程式碼的位置(Xpath表達式) #暱稱
        gamer_data_get["other"]["gamer_friend_nicknamelist"] = list(map(str, gamer_friend_sourcecode.xpath(xpath_on_web))) #使用Xpath表達式提出，為巴哈好友暱稱nickname清單
        Go_to_gamer_friendlist.close() #關閉對gamer_friendlist_url夾帶headers發出GET請求
    
    for i in range(len(gamer_data_get["other"]["gamer_friend_accountlist"])): #依照順序數字建立字典，再將資訊分別塞進去
        gamer_data_get["friend"][i] = {}
        gamer_data_get["friend"][i]["account"] = gamer_data_get["other"]["gamer_friend_accountlist"][i]
        gamer_data_get["friend"][i]["nickname"] = gamer_data_get["other"]["gamer_friend_nicknamelist"][i]
        gamer_friend_url = "https://home.gamer.com.tw/homeindex.php?owner={}".format(gamer_data_get["friend"][i]["account"]) #組合成追蹤者的巴哈小屋(舊版)連結
        gamer_friend_statuscode = go_to_web(gamer_friend_url) #回傳網路狀態碼
        try:
            try:
                gamer_friend_url = "https://home.gamer.com.tw/homeindex.php?owner={}".format(gamer_data_get["friend"][i]["account"]) #組合成好友的舊版巴哈小屋連結
                gamer_friend_statuscode = go_to_web(gamer_friend_url) #回傳網路狀態碼
                version = "Old" #標示為舊版小屋
            except:
                gamer_friend_url = "https://home.gamer.com.tw/profile/index.php?owner={}".format(gamer_data_get["friend"][i]["account"]) #組合成好友的新版巴哈小屋連結
                gamer_friend_statuscode = go_to_web(gamer_friend_url) #回傳網路狀態碼
                version = "New" #標示為新版小屋
        except:
            version = "Delete"
        headers = {"User-Agent" : "Mozilla/5.0 (Windows; U; Windows NT 6.1; rv:2.2) Gecko/20110201"} #設置http頭欄位，裡面夾帶瀏覽器識別標籤
        Go_to_gamer_friend = requests.get(gamer_friend_url, headers = headers, timeout = 60, allow_redirects = False, stream = True, verify = False) #對gfri_url夾帶headers發出GET請求，timeout為最長反應時間，allow_redirects為禁止重新定向，stream為強制解壓縮，verify為SSL憑證檢查功能       
        requests.packages.urllib3.disable_warnings() #關閉InsecureRequestWarning的顯示
        Go_to_gamer_friend.encoding = "utf-8" #指定網頁的編碼格式
        gamer_friend_sourcecode = etree.HTML(Go_to_gamer_friend.text) #取得網頁原始碼
        gamer_title_number = 0 #巴哈小屋大標的序號
        status = False #表示狀態
        if gamer_friend_statuscode == 200 and version == "Old": #如果可以用舊版小屋連線
            gamer_title = list(map(str, gamer_friend_sourcecode.xpath("//div[@id='BH-background']/div[@id='BH-wrapper']/div[@id='BH-slave']/h5/text()"))) #從巴哈小屋的原始碼篩出大標
            while status == False and gamer_title_number < len(gamer_title): #當status等於False和gamer_title_number小於巴哈大標數量時，持續運作
                if "個人紀錄" == gamer_title[gamer_title_number]:
                    xpath_on_web = "//div[@id='BH-background']/div[@id='BH-wrapper']/div[@id='BH-slave']/div[@class='BH-rbox BH-list1']/ul/li[4]/text()" #取得註冊日期
                    gamer_data_get["friend"][i]["regdate"] = str(gamer_friend_sourcecode.xpath(xpath_on_web)[0]).split("：")[1] #使用Xpath表達式提出，將帳號建立日放進regdate鍵的值
                    xpath_on_web = "//div[@id='BH-background']/div[@id='BH-wrapper']/div[@id='BH-slave']/div[@class='BH-rbox BH-list1']/ul/li[5]/text()" #取得最後上站日期
                    gamer_data_get["friend"][i]["lastondate"] = str(gamer_friend_sourcecode.xpath(xpath_on_web)[0]).split("：")[1] #使用Xpath表達式提出，將帳號最後上線日放進lastondate鍵的值
                    status = True
                else:
                    gamer_title_number = gamer_title_number+1 #巴哈小屋大標的序號加1
            if status == False: #巴哈姆特模塊化特色所致
                gamer_data_get["friend"][i]["regdate"] = "舊版無區塊"
                gamer_data_get["friend"][i]["lastondate"] = "舊版無區塊"
        elif gamer_friend_statuscode == 200 and version == "New": #如果可以用新版小屋連線，目前程式還不支援去新版尋找
            gamer_data_get["friend"][i]["regdate"] = "新版不支援"
            gamer_data_get["friend"][i]["lastondate"] = "新版不支援"
        else: #如果都不行
            gamer_data_get["friend"][i]["regdate"] = None
            gamer_data_get["friend"][i]["lastondate"] = None
        Go_to_gamer_friend.close() #關閉對gamer_friend_url夾帶headers發出GET請求
    code_section_4_status = "〇"
except:
    code_section_4_status = "✕"

try:
    #開啟試算表
    certificate = pygsheets.authorize(service_file='.\google_sheets_API_key.json') #取得位置在同層級目錄的Google sheets API憑證
    googlesheets_url = "https://docs.google.com/spreadsheets/d/1vLopfsKHRNaS02bI5AmKHsBbbqtL4EbY4k47SRivMSY" #有spreadsheetId的google sheets網址
    open_googlesheets = certificate.open_by_url(googlesheets_url) #開啟Google sheets
    open_googlesheets_status = True
    print("open_googlesheets_status", open_googlesheets_status)
except:
    open_googlesheets_status = False
    print("open_googlesheets_status", open_googlesheets_status)

if open_googlesheets_status == True:
    try:
        #獲取運作天數
        worksheet = open_googlesheets.worksheet('id', 1888051586) #以sheetId定位試算表位置為倒數第2張的"系統訊息"
        number = 0 #天數寫了總共幾格的計數
        days = [] #總共幾天
        while worksheet.get_value("B{}".format(number+3)) != "": #一直運作直到得到的字串為「完全空白」
            days.append(worksheet.get_value("B{}".format(number+3))) #將得到的字串加入days串列
            number = number+1
        days = list(set(days)) #將days串列中重複的元素變成1個再組成一個新的days串列
        date_number = len(days) #天數為days元素的數量
        
        runtime = int(worksheet.get_value("A{}".format(number+2))) #獲取運作次數，number+2為「完全空白」前一格的位置
        basic_status = True
        print("basic_status", basic_status)
    except:
        basic_status = False
        print("basic_status", basic_status)
    
    if basic_status == True:
        try: #寫入試算表1："人氣紀錄"
            worksheet = open_googlesheets.worksheet('id', 0) #以sheetId定位試算表位置為第1張的"人氣紀錄表"
            
            writesit = 3 #日期被寫下的格數
            while worksheet.get_value("A{}".format(writesit)) != "": #一直運作直到得到的字串為「完全空白」
                writesit = writesit+1
            
            if str(datetime.strptime(worksheet.get_value("A{}".format(writesit-1))+worksheet.get_value("C{}".format(writesit-1)), "%Y年%m月%d日")) == str(datetime.strptime(ad_year_yesterday+date_yesterday, "%Y年%m月%d日")): #如果最後一天的日期等於「昨天的台北時間」
                gamer_writesit = writesit-1 #巴哈資料被寫下的格數
            elif str(datetime.strptime(worksheet.get_value("A{}".format(writesit-1))+worksheet.get_value("C{}".format(writesit-1)), "%Y年%m月%d日")) == str(datetime.strptime(ad_year_today+date_today, "%Y年%m月%d日")): #如果最後一天的日期等於「今天的台北時間」
                gamer_writesit = writesit-2 #巴哈資料被寫下的格數
            
            month_lastday = lastday_of_month(int(worksheet.get_value("A{}".format(writesit-2)).split("年")[0]), int(worksheet.get_value("C{}".format(writesit-2)).split("月")[0])) #取得該月最後一天的日期
            
            gamer_yesterday_allview = "=394+SUM($F$3:F{})".format(gamer_writesit) #計算到昨天為止的巴哈總人氣數
            if worksheet.get_value("A{}".format(writesit-1))+worksheet.get_value("C{}".format(writesit-1)) != ad_year_today+date_today: #如果最後一天的日期不等於今天的日期
                worksheet.update_value("A{}".format(writesit-1), ad_year_yesterday) #寫入昨日的西元紀年
                worksheet.update_value("B{}".format(writesit-1), mg_year_yesterday) #寫入昨日的民國紀年
                worksheet.update_value("C{}".format(writesit-1), date_yesterday) #寫入昨日的日期
                worksheet.update_value("A{}".format(writesit), ad_year_today) #寫入今日的西元紀年
                worksheet.update_value("B{}".format(writesit), mg_year_today) #寫入今日的民國紀年
                worksheet.update_value("C{}".format(writesit), date_today) #寫入今日的日期
            if str(datetime.strptime(worksheet.get_value("A{}".format(gamer_writesit))+worksheet.get_value("C{}".format(gamer_writesit)), "%Y年%m月%d日")) == str(gamer_dates[6]): #如果試算表上所寫的「昨天的日期」等於「ad_year+date_yesterday」
                worksheet.update_value("D{}".format(gamer_writesit), gamer_friend_number) #寫入巴哈好友圈人數
                worksheet.update_value("E{}".format(gamer_writesit), gamer_follower_number) #寫入巴哈追蹤者人數
                worksheet.update_value("F{}".format(gamer_writesit), gamer_viewnumbers[6]) #寫入巴哈當日人氣數
                worksheet.update_value("G{}".format(gamer_writesit), gamer_yesterday_allview) #寫入巴哈總人氣數
                if (gamer_writesit > (lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0]))-1)) and (int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[1].split("日")[0]) == month_lastday): #如果試算表日期等於月的最後一天
                    worksheet.update_values("H{}".format(gamer_writesit-3), [["本月總和"], ["=SUM(F{}:F{})".format(gamer_writesit+1-(lastday_of_month(int(worksheet.get_value("A{}".format(gamer_writesit)).split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0]))), gamer_writesit)], ["本月日平均"], ["=AVERAGE(F{}:F{})".format(gamer_writesit+1-(lastday_of_month(int(worksheet.get_value("A{}".format(gamer_writesit)).split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0]))), gamer_writesit)]]) #寫入本月總和和日平均
                if int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0]) in [3, 6, 9, 12]:
                    if (gamer_writesit > (lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0])-2)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0])-1)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0]))-1)) and (int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[1].split("日")[0]) == month_lastday):
                        worksheet.update_values("H{}".format(gamer_writesit-7), [["本季總和"], ["=SUM(F{}:F{})".format(gamer_writesit+1-(lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0])-2)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0])-1)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0]))), gamer_writesit)], ["本季日平均"], ["=AVERAGE(F{}:F{})".format(gamer_writesit+1-(lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0])-2)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0])-1)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0]))), gamer_writesit)]]) #寫入本季總和和日平均
                if int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0]) in [6, 12]:
                    if (gamer_writesit > (lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0])-5)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0])-4)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0])-3)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0])-2)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0])-1)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0]))-1)) and (int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[1].split("日")[0]) == month_lastday):
                        worksheet.update_values("H{}".format(gamer_writesit-11), [["本半年總和"], ["=SUM(F{}:F{})".format(gamer_writesit+1-(lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0])-5)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0])-4)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0])-3)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0])-2)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0])-1)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0]))), gamer_writesit)], ["本半年日平均"], ["=AVERAGE(F{}:F{})".format(gamer_writesit+1-(lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0])-5)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0])-4)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0])-3)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0])-2)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0])-1)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0]))), gamer_writesit)]]) #寫入本季總和和日平均
                if int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[0]) == 12:
                    if (gamer_writesit > 364) and (int(worksheet.get_value("C{}".format(gamer_writesit)).split("月")[1].split("日")[0]) == month_lastday):
                        worksheet.update_values("H{}".format(gamer_writesit-15), [["本年總和"], ["=SUM(F{}:F{})".format(gamer_writesit+1-(lastday_of_month(int(ad_year_yesterday.split("年")[0]), 1)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 2)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 3)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 4)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 5)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 6)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 7)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 8)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 9)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 10)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 11)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 12)))], ["本年日平均"], ["=AVERAGE(F{}:F{})".format(gamer_writesit+1-(lastday_of_month(int(ad_year_yesterday.split("年")[0]), 1)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 2)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 3)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 4)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 5)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 6)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 7)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 8)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 9)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 10)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 11)+lastday_of_month(int(ad_year_yesterday.split("年")[0]), 12)))]]) #寫入本年總和和日平均
            worksheet.update_value("H4", "=394+SUM($F$3:F{})".format(gamer_writesit)) #寫入巴哈總人氣數
            worksheet.update_value("H56", "=394+SUM($F$3:F{})".format(gamer_writesit)) #寫入巴哈總人氣數
            code_section_5_status = "〇"
        except:
            code_section_5_status = "✕"
        
        try:
            #寫入試算表2："巴哈追蹤名單"
            worksheet = open_googlesheets.worksheet('id', 1960934836) #以sheetId定位試算表位置為"巴哈姆特追蹤名單"
            
            gamer_data_last = {"follower" : {}, "friend" : {}, "other" : {}} #建立字典，並在裡面建立fan、friena和other3鍵，功能為放置從噗浪取得(昨天)的資料
                #account為巴哈帳戶，nickname為名字，regdate巴哈帳號建立日，lastondate巴哈最後登入日 #"account" : [], "nickname" :[], "regdate" : [], "lastondate" : []
            
            gamer_follower_number_last = 0 #下方迴圈的計數
            gamer_data_last["other"]["gamer_follower_accountlist"] = [] #建立與鍵對應的值，此值為串列，功能為放置account
            gamer_data_last["other"]["gamer_follower_nicknamelist"] = [] #建立與鍵對應的值，此值為串列，功能為放置nickname
            while worksheet.get_value("B{}".format(gamer_follower_number_last+3)) != "": #一直運作直到得到的字串為「沒有東西」
                gamer_data_last["follower"][gamer_follower_number_last] = {} #以gamer_follower_number_last為鍵建立對應的值，此值為字典
                gamer_data_last["follower"][gamer_follower_number_last]["account"] = worksheet.get_value("B{}".format(gamer_follower_number_last+3)) #建立鍵，自試算表的指定位置讀取值
                gamer_data_last["follower"][gamer_follower_number_last]["nickname"] = worksheet.get_value("C{}".format(gamer_follower_number_last+3))
                gamer_data_last["follower"][gamer_follower_number_last]["regdate"] = worksheet.get_value("D{}".format(gamer_follower_number_last+3))
                gamer_data_last["follower"][gamer_follower_number_last]["lastondate"] = worksheet.get_value("E{}".format(gamer_follower_number_last+3))
                gamer_data_last["follower"][gamer_follower_number_last]["startfollowdate"] = worksheet.get_value("F{}".format(gamer_follower_number_last+3))
                gamer_data_last["follower"][gamer_follower_number_last]["followday"] = worksheet.get_value("G{}".format(gamer_follower_number_last+3))
                gamer_data_last["follower"][gamer_follower_number_last]["endfollowdate"] = worksheet.get_value("H{}".format(gamer_follower_number_last+3))
                gamer_data_last["other"]["gamer_follower_accountlist"].append(gamer_data_last["follower"][gamer_follower_number_last]["account"]) #將account這個鍵的值加入串列
                gamer_data_last["other"]["gamer_follower_nicknamelist"].append(gamer_data_last["follower"][gamer_follower_number_last]["nickname"]) #將nickname這個鍵的值加入串列
                gamer_follower_number_last = gamer_follower_number_last+1
            
            gamer_friend_number_last = 0 #下方迴圈的計數
            gamer_data_last["other"]["gamer_friend_accountlist"] = [] #建立與鍵對應的值，此值為串列，功能為放置account
            gamer_data_last["other"]["gamer_friend_nicknamelist"] = [] #建立與鍵對應的值，此值為串列，功能為放置nickname
            while worksheet.get_value("I{}".format(gamer_friend_number_last+3)) != "": #一直運作直到得到的字串為END
                gamer_data_last["friend"][gamer_friend_number_last] = {} #以gamer_friend_number_last為鍵建立對應的值，此值為字典
                gamer_data_last["friend"][gamer_friend_number_last]["account"] = worksheet.get_value("I{}".format(gamer_friend_number_last+3)) #建立鍵，自試算表的指定位置讀取值
                gamer_data_last["friend"][gamer_friend_number_last]["nickname"] = worksheet.get_value("J{}".format(gamer_friend_number_last+3))
                gamer_data_last["friend"][gamer_friend_number_last]["regdate"] = worksheet.get_value("K{}".format(gamer_friend_number_last+3))
                gamer_data_last["friend"][gamer_friend_number_last]["lastondate"] = worksheet.get_value("L{}".format(gamer_friend_number_last+3))
                gamer_data_last["friend"][gamer_friend_number_last]["startfollowdate"] = worksheet.get_value("M{}".format(gamer_friend_number_last+3))
                gamer_data_last["friend"][gamer_friend_number_last]["followday"] = worksheet.get_value("N{}".format(gamer_friend_number_last+3))
                gamer_data_last["friend"][gamer_friend_number_last]["endfollowdate"] = worksheet.get_value("O{}".format(gamer_friend_number_last+3))
                gamer_data_last["other"]["gamer_friend_accountlist"].append(gamer_data_last["friend"][gamer_friend_number_last]["account"]) #將account這個鍵的值加入串列
                gamer_data_last["other"]["gamer_friend_nicknamelist"].append(gamer_data_last["friend"][gamer_friend_number_last]["nickname"]) #將nickname這個鍵的值加入串列
                gamer_friend_number_last = gamer_friend_number_last+1
            
            gamer_data_new = {"follower" : {}, "friend" : {}} #建立字典，並在裡面建立fan、friena和other3鍵，功能為放置從噗浪取得(昨天)的資料
                #account為巴哈帳戶，nickname為名字，regdate巴哈帳號建立日，lastondate巴哈最後登入日 #"account" : [], "nickname" :[], "regdate" : [], "lastondate" : []
            
            gamer_follower_number_new = 0 #下方迴圈的計數
            gamer_follower_string = "gamer_follower-" #沒有起始日則在表格填上這個字串
            for i in range(len(gamer_data_get["follower"])): #範圍為follower鍵的值(字典)內所含有的數量
                gamer_data_new["follower"][gamer_follower_number_new] = {} #以gamer_follower_number_new為鍵建立對應的值，此值為字典
                try:
                    if gamer_data_get["follower"][i]["account"] in gamer_data_last["other"]["gamer_follower_accountlist"]: #如果從巴哈取得的追蹤者account的值有在串列裡
                        gamer_data_new["follower"][gamer_follower_number_new]["status"] = "old" #標記狀態為old
                        gamer_data_new["follower"][gamer_follower_number_new]["account"] = gamer_data_get["follower"][i]["account"] #將從巴哈取得的追蹤者account、nickname、建立日和最後上線日記錄進gamer_data_new
                        gamer_data_new["follower"][gamer_follower_number_new]["nickname"] = gamer_data_get["follower"][i]["nickname"]
                        gamer_data_new["follower"][gamer_follower_number_new]["regdate"] = gamer_data_get["follower"][i]["regdate"]
                        gamer_data_new["follower"][gamer_follower_number_new]["lastondate"] = gamer_data_get["follower"][i]["lastondate"]
                        gamer_data_new["follower"][gamer_follower_number_new]["startfollowdate"] = gamer_data_last["follower"][gamer_data_last["other"]["gamer_follower_accountlist"].index(gamer_data_new["follower"][gamer_follower_number_new]["account"])]["startfollowdate"] #尋找指定序號後，從前一筆(前天)紀錄取得開始追蹤的日期
                        gamer_data_new["follower"][gamer_follower_number_new]["followday"] = get_followday(number, start_time, i, gamer_data_new["follower"][gamer_follower_number_new]["status"], gamer_data_new["follower"][gamer_follower_number_new]["startfollowdate"], gamer_data_last["follower"][gamer_data_last["other"]["gamer_follower_accountlist"].index(gamer_data_new["follower"][gamer_follower_number_new]["account"])]["followday"], gamer_follower_string, 0, worksheet) #傳入副程式number, start_time, i, status, startfollowdate, followday, string, endfollowdate這些參數，取得追蹤天數
                        gamer_data_new["follower"][gamer_follower_number_new]["endfollowdate"] = "" #結束時間為空白
                        gamer_follower_number_new = gamer_follower_number_new+1
                    else:
                        del gamer_data_new["follower"][gamer_follower_number_new] #如果從巴哈取得的追蹤者account的值沒有在串列裡，刪除以gamer_follower_number_new為名的鍵
                except:
                    del gamer_data_new["follower"][gamer_follower_number_new] #如果從巴哈取得的追蹤者account的值沒有在串列裡，刪除以gamer_follower_number_new為名的鍵
            
            for i in range(len(gamer_data_get["follower"])): #範圍為follower鍵的值(字典)內所含有的數量
                gamer_data_new["follower"][gamer_follower_number_new] = {} #以gamer_follower_number_new為鍵建立對應的值，此值為字典
                try:
                    if gamer_data_get["follower"][i]["account"] not in gamer_data_last["other"]["gamer_follower_accountlist"]: #如果從巴哈取得的追蹤者account的值沒有在串列裡
                        gamer_data_new["follower"][gamer_follower_number_new]["status"] = "new" #標記狀態為new
                        gamer_data_new["follower"][gamer_follower_number_new]["account"] = gamer_data_get["follower"][i]["account"] #將從巴哈取得的粉絲account、nickname、建立日和最後上線日記錄進gamer_data_new
                        gamer_data_new["follower"][gamer_follower_number_new]["nickname"] = gamer_data_get["follower"][i]["nickname"]
                        gamer_data_new["follower"][gamer_follower_number_new]["regdate"] = gamer_data_get["follower"][i]["regdate"]
                        gamer_data_new["follower"][gamer_follower_number_new]["lastondate"] = gamer_data_get["follower"][i]["lastondate"]
                        gamer_data_new["follower"][gamer_follower_number_new]["startfollowdate"] = str(start_time.year)+"-"+str(start_time.month)+"-"+str(int(start_time.day)-1) #以昨天為開始追蹤日
                        gamer_data_new["follower"][gamer_follower_number_new]["followday"] = 0 #第0天
                        gamer_data_new["follower"][gamer_follower_number_new]["endfollowdate"] = "" #結束時間為空白
                        gamer_follower_number_new = gamer_follower_number_new+1
                    else:
                        del gamer_data_new["follower"][gamer_follower_number_new] #如果從巴哈取得的追蹤者account的值沒有在串列裡，刪除以gamer_follower_number_new為名的鍵
                except:
                    del gamer_data_new["follower"][gamer_follower_number_new] #如果從巴哈取得的追蹤者account的值沒有在串列裡，刪除以gamer_follower_number_new為名的鍵
            
            for i in range(len(gamer_data_get["follower"])): #範圍為follower鍵的值(字典)內所含有的數量
                gamer_data_new["follower"][gamer_follower_number_new] = {} #以gamer_follower_number_new為鍵建立對應的值，此值為字典
                try:
                    if gamer_data_last["follower"][i]["account"] not in gamer_data_get["other"]["gamer_follower_accountlist"]: #如果從試算表取得的追蹤者account的值沒有在串列裡
                        gamer_data_new["follower"][gamer_follower_number_new]["status"] = "leave" #標記狀態為leave
                        gamer_data_new["follower"][gamer_follower_number_new]["account"] = gamer_data_get["follower"]["account"][i] #將從巴哈取得的粉絲account、nickname、建立日和最後上線日記錄進gamer_data_new
                        gamer_data_new["follower"][gamer_follower_number_new]["nickname"] = gamer_data_get["follower"]["nickname"][i]
                        gamer_data_new["follower"][gamer_follower_number_new]["regdate"] = gamer_data_get["follower"]["regdate"][i]
                        gamer_data_new["follower"][gamer_follower_number_new]["lastondate"] = gamer_data_get["follower"]["lastondate"][i]
                        gamer_data_new["follower"][gamer_follower_number_new]["startfollowdate"] = gamer_data_last["follower"][gamer_data_last["other"]["gamer_follower_accountlist"].index(gamer_data_new["follower"][gamer_follower_number_new]["account"])]["startfollowdate"] #尋找指定序號後，從前一筆(前天)紀錄取得開始追蹤的日期
                        if gamer_data_last["follower"][gamer_data_last["other"]["gamer_follower_accountlist"].index(gamer_data_new["follower"][gamer_follower_number_new]["account"])]["endfollowdate"] == "": #如果日期等於空白
                            gamer_data_new["follower"][gamer_follower_number_new]["endfollowdate"] = str(start_time.year)+"-"+str(start_time.month)+"-"+str(int(start_time.day)-1) #登記昨天為結束追蹤的日期
                        else: #如果日期不等於空白
                            gamer_data_new["follower"][gamer_follower_number_new]["endfollowdate"] = gamer_data_last["follower"][gamer_data_last["other"]["gamer_follower_accountlist"].index(gamer_data_new["fan"][gamer_follower_number_new]["account"])]["endfollowdate"] #尋找指定序號後，從前一筆(前天)紀錄取得結束追蹤的日期
                        gamer_data_new["follower"][gamer_follower_number_new]["followday"] = get_followday(number, start_time, i, gamer_data_new["follower"][gamer_follower_number_new]["status"], gamer_data_new["follower"][gamer_follower_number_new]["startfollowdate"], gamer_data_last["follower"][gamer_data_last["other"]["gamer_follower_accountlist"].index(gamer_data_new["follower"][gamer_follower_number_new]["account"])]["followday"], gamer_follower_string, gamer_data_new["follower"][gamer_follower_number_new]["endfollowdate"], worksheet) #傳入副程式number, start_time, i, status, startfollowdate, followday, string, endfollowdate這些參數，取得追蹤天數
                        gamer_follower_number_new = gamer_follower_number_new+1
                    else:
                        del gamer_data_new["follower"][gamer_follower_number_new] #如果從巴哈取得的追蹤者account的值沒有在串列裡，刪除以gamer_follower_number_new為名的鍵
                except:
                    del gamer_data_new["follower"][gamer_follower_number_new] #如果從巴哈取得的追蹤者account的值沒有在串列裡，刪除以gamer_follower_number_new為名的鍵
            
            gamer_friend_number_new = 0 #下方迴圈的計數
            gamer_friend_string = "gamer_friend-" #沒有起始日則在表格填上這個字串
            for i in range(len(gamer_data_get["friend"])): #範圍為friend鍵的值(字典)內所含有的數量
                gamer_data_new["friend"][gamer_friend_number_new] = {} #以gamer_friend_number_new為鍵建立對應的值，此值為字典
                try:
                    if gamer_data_get["friend"][i]["account"] in gamer_data_last["other"]["gamer_friend_accountlist"]: #如果從巴哈取得的朋友account的值有在串列裡
                        gamer_data_new["friend"][gamer_friend_number_new]["status"] = "old" #標記狀態為old
                        gamer_data_new["friend"][gamer_friend_number_new]["account"] = gamer_data_get["friend"][i]["account"] #將從巴哈取得的朋友account、nickname、建立日和最後上線日記錄進gamer_data_new
                        gamer_data_new["friend"][gamer_friend_number_new]["nickname"] = gamer_data_get["friend"][i]["nickname"]
                        gamer_data_new["friend"][gamer_friend_number_new]["regdate"] = gamer_data_get["friend"][i]["regdate"]
                        gamer_data_new["friend"][gamer_friend_number_new]["lastondate"] = gamer_data_get["friend"][i]["lastondate"]
                        gamer_data_new["friend"][gamer_friend_number_new]["startfollowdate"] = gamer_data_last["friend"][gamer_data_last["other"]["gamer_friend_accountlist"].index(gamer_data_new["friend"][gamer_friend_number_new]["account"])]["startfollowdate"] #尋找指定序號後，從前一筆(前天)紀錄取得開始追蹤的日期
                        gamer_data_new["friend"][gamer_friend_number_new]["followday"] = get_followday(number, start_time, i, gamer_data_new["friend"][gamer_friend_number_new]["status"], gamer_data_new["friend"][gamer_friend_number_new]["startfollowdate"], gamer_data_last["friend"][gamer_data_last["other"]["gamer_friend_accountlist"].index(gamer_data_new["friend"][gamer_friend_number_new]["account"])]["followday"], gamer_friend_string, 0, worksheet) #傳入副程式number, start_time, i, status, startfollowdate, followday, string, endfollowdate這些參數，取得追蹤天數
                        gamer_data_new["friend"][gamer_friend_number_new]["endfollowdate"] = "" #結束時間為空白
                        gamer_friend_number_new = gamer_friend_number_new+1
                    else:
                        del gamer_data_new["friend"][gamer_friend_number_new] #如果從巴哈取得的朋友account的值沒有在串列裡，刪除以gamer_friend_number_new為名的鍵
                except:
                    del gamer_data_new["friend"][gamer_friend_number_new] #如果從巴哈取得的朋友account的值沒有在串列裡，刪除以gamer_friend_number_new為名的鍵
            
            for i in range(len(gamer_data_get["friend"])): #範圍為friend鍵的值(字典)內所含有的數量
                gamer_data_new["friend"][gamer_friend_number_new] = {} #以gamer_friend_number_new為鍵建立對應的值，此值為字典
                try:
                    if gamer_data_get["friend"][i]["account"] not in gamer_data_last["other"]["gamer_friend_accountlist"]: #如果從巴哈取得的朋友account的值有不在串列裡
                        gamer_data_new["friend"][gamer_friend_number_new]["status"] = "new" #標記狀態為new
                        gamer_data_new["friend"][gamer_friend_number_new]["account"] = gamer_data_get["friend"][i]["account"] #將從巴哈取得的朋友account、nickname、建立日和最後上線日記錄進gamer_data_new
                        gamer_data_new["friend"][gamer_friend_number_new]["nickname"] = gamer_data_get["friend"][i]["nickname"]
                        gamer_data_new["friend"][gamer_friend_number_new]["regdate"] = gamer_data_get["friend"][i]["regdate"]
                        gamer_data_new["friend"][gamer_friend_number_new]["lastondate"] = gamer_data_get["friend"][i]["lastondate"]
                        gamer_data_new["friend"][gamer_friend_number_new]["startfollowdate"] = str(start_time.year)+"-"+str(start_time.month)+"-"+str(int(start_time.day)-1) #以昨天為開始追蹤日
                        gamer_data_new["friend"][gamer_friend_number_new]["followday"] = 0 #第0天
                        gamer_data_new["friend"][gamer_friend_number_new]["endfollowdate"] =  "" #結束時間為空白
                        gamer_friend_number_new = gamer_friend_number_new+1
                    else:
                        del gamer_data_new["friend"][gamer_friend_number_new] #如果從巴哈取得的朋友account的值沒有在串列裡，刪除以gamer_friend_number_new為名的鍵
                except:
                    del gamer_data_new["friend"][gamer_friend_number_new] #如果從巴哈取得的朋友account的值沒有在串列裡，刪除以gamer_friend_number_new為名的鍵
            
            for i in range(len(gamer_data_get["friend"])): #範圍為friend鍵的值(字典)內所含有的數量
                gamer_data_new["friend"][gamer_friend_number_new] = {} #以gamer_friend_number_new為鍵建立對應的值，此值為字典
                try:
                    if gamer_data_last["friend"][i]["account"] not in gamer_data_get["other"]["gamer_friend_accountlist"]: #如果從試算表取得的朋友account的值沒有在串列裡
                        gamer_data_new["friend"][gamer_friend_number_new]["status"] = "leave" #標記狀態為leave
                        gamer_data_new["friend"][gamer_friend_number_new]["account"] = gamer_data_get["friend"][i]["account"] #將從巴哈取得的朋友account、nickname、建立日和最後上線日記錄進gamer_data_new
                        gamer_data_new["friend"][gamer_friend_number_new]["nickname"] = gamer_data_get["friend"][i]["nickname"]
                        gamer_data_new["friend"][gamer_friend_number_new]["regdate"] = gamer_data_get["friend"][i]["regdate"]
                        gamer_data_new["friend"][gamer_friend_number_new]["lastondate"] = gamer_data_get["friend"][i]["lastondate"]
                        gamer_data_new["friend"][gamer_friend_number_new]["startfollowdate"] = gamer_data_last["friend"][gamer_data_last["other"]["gamer_friend_accountlist"].index(gamer_data_new["friend"][gamer_friend_number_new]["account"])]["startfollowdate"] #尋找指定序號後，從前一筆(前天)紀錄取得開始追蹤的日期
                        
                        if gamer_data_last["friend"][gamer_data_last["other"]["gamer_friend_accountlist"].index(gamer_data_new["friend"][gamer_friend_number_new]["account"])]["endfollowdate"] == "": #如果日期等於空白
                            gamer_data_new["friend"][gamer_friend_number_new]["endfollowdate"] = str(start_time.year)+"-"+str(start_time.month)+"-"+str(int(start_time.day)-1) #登記昨天為結束追蹤的日期
                        else: #如果日期不等於空白
                            gamer_data_new["friend"][gamer_friend_number_new]["endfollowdate"] = gamer_data_last["friend"][gamer_data_last["other"]["gamer_friend_accountlist"].index(gamer_data_new["friend"][gamer_friend_number_new]["account"])]["endfollowdate"] #尋找指定序號後，從前一筆(前天)紀錄取得結束追蹤的日期
                        gamer_data_new["friend"][gamer_friend_number_new]["followday"] = get_followday(number, start_time, i, gamer_data_new["friend"][gamer_friend_number_new]["status"], gamer_data_new["friend"][gamer_friend_number_new]["startfollowdate"], gamer_data_last["friend"][gamer_data_last["other"]["gamer_friend_accountlist"].index(gamer_data_new["friend"][gamer_friend_number_new]["account"])]["followday"], gamer_friend_string, gamer_data_new["friend"][gamer_friend_number_new]["endfollowdate"], worksheet) #傳入副程式number, start_time, i, status, startfollowdate, followday, string, endfollowdate這些參數，取得追蹤天數
                        gamer_friend_number_new = gamer_friend_number_new+1
                    else:
                        del gamer_data_new["friend"][gamer_friend_number_new] #如果從巴哈取得的朋友account的值沒有在串列裡，刪除以gamer_friend_number_new為名的鍵
                except:
                    del gamer_data_new["friend"][gamer_friend_number_new] #如果從巴哈取得的朋友account的值沒有在串列裡，刪除以gamer_friend_number_new為名的鍵
            
            for i in range(3, (max(gamer_follower_number_last, gamer_friend_number_last)+4)):
                worksheet.update_values("B{}".format(i), [["", "", "", "", "", "", "", "", "", "", "", "", "", ""]]) #寫入「完全空白」來清除版面
            
            for gamer_follower_number in range(len(gamer_data_new["follower"])): #範圍為fan鍵的值(字典)內所含有的數量
                worksheet.update_value("B{}".format(gamer_follower_number+3), gamer_data_new["follower"][gamer_follower_number]["account"]) #寫入追蹤者帳號
                worksheet.update_value("C{}".format(gamer_follower_number+3), gamer_data_new["follower"][gamer_follower_number]["nickname"]) #寫入追蹤者暱稱
                worksheet.update_value("D{}".format(gamer_follower_number+3), gamer_data_new["follower"][gamer_follower_number]["regdate"]) #寫入註冊日期
                worksheet.update_value("E{}".format(gamer_follower_number+3), gamer_data_new["follower"][gamer_follower_number]["lastondate"]) #寫入最後上站日期
                worksheet.update_value("F{}".format(gamer_follower_number+3), gamer_data_new["follower"][gamer_follower_number]["startfollowdate"]) #寫入開始追蹤日期
                worksheet.update_value("G{}".format(gamer_follower_number+3), gamer_data_new["follower"][gamer_follower_number]["followday"]) #寫入追蹤天數
                worksheet.update_value("H{}".format(gamer_follower_number+3), gamer_data_new["follower"][gamer_follower_number]["endfollowdate"]) #寫入結束追蹤日期
            
            for gamer_friend_number in range(len(gamer_data_new["friend"])): #範圍為friend鍵的值(字典)內所含有的數量
                worksheet.update_value("I{}".format(gamer_friend_number+3), gamer_data_new["friend"][gamer_friend_number]["account"]) #寫入朋友帳號
                worksheet.update_value("J{}".format(gamer_friend_number+3), gamer_data_new["friend"][gamer_friend_number]["nickname"]) #寫入朋友帳號
                worksheet.update_value("K{}".format(gamer_friend_number+3), gamer_data_new["friend"][gamer_friend_number]["regdate"]) #寫入註冊日期
                worksheet.update_value("L{}".format(gamer_friend_number+3), gamer_data_new["friend"][gamer_friend_number]["lastondate"]) #寫入上次上站日期
                worksheet.update_value("M{}".format(gamer_friend_number+3), gamer_data_new["friend"][gamer_friend_number]["startfollowdate"]) #寫入開始追蹤日期
                worksheet.update_value("N{}".format(gamer_friend_number+3), gamer_data_new["friend"][gamer_friend_number]["followday"]) #寫入追蹤天數
                worksheet.update_value("O{}".format(gamer_friend_number+3), gamer_data_new["friend"][gamer_friend_number]["endfollowdate"]) #寫入結束追蹤日期
            code_section_6_status = "〇"
        except:
            code_section_6_status = "✕"
        
        try:
            #寫入試算表4："系統訊息"
            worksheet= open_googlesheets.worksheet('id', 1888051586) #以sheetId定位試算表位置為倒數第2張的"系統訊息"
            
            worksheet.update_value("A{}".format(number+3), runtime+1) #寫入本次的運作次數
            if str(start_time).split(" ")[0] == str(worksheet.get_value("E{}".format(number+2))): #如果start_time的日期等於("D{}".format(number+1))即今天的字串
                worksheet.update_value("B{}".format(number+3), date_number) #寫入同樣天數
                worksheet.update_value("C{}".format(number+3), int(worksheet.get_value("C{}".format(number+2)))+1) #取得前一格的天次數後加1再寫入
            else:
                worksheet.update_value("B{}".format(number+3), date_number+1) #寫入天數+1
                worksheet.update_value("C{}".format(number+3), "1") #寫入天次數為1
            worksheet.update_value("D{}".format(number+3), "study_gamer_code(Python)") #寫入現在執行的檔案
            worksheet.update_value("E{}".format(number+3), str(start_time).split(" ")[0].split("+")[0]) #寫入現在的時間
            worksheet.update_value("F{}".format(number+3), str(start_time).split(" ")[1].split("+")[0]) #寫入程式開始的時間
            worksheet.update_value("I{}".format(number+3), code_section_1_status) #寫入程式運作狀態 #裝置基本資訊取得
            worksheet.update_value("J{}".format(number+3), code_section_2_status) #寫入程式運作狀態 #追蹤者、好友和流量統計數取得
            worksheet.update_value("K{}".format(number+3), code_section_3_status) #寫入程式運作狀態 #追蹤者名單取得
            worksheet.update_value("L{}".format(number+3), code_section_4_status) #寫入程式運作狀態 #好友名單取得
            worksheet.update_value("M{}".format(number+3), code_section_5_status) #寫入程式運作狀態 #人氣記錄表寫入
            worksheet.update_value("N{}".format(number+3), code_section_6_status) #寫入程式運作狀態 #名單寫入
            worksheet.update_value("P{}".format(number+3), start_timecall[7]) #寫入時間伺服器連結
            worksheet.update_value("Q{}".format(number+3), start_timecall[8]) #寫入時間伺服器連結在時間伺服器串列中的編號
            if code_section_1_status == "〇":
                worksheet.update_value("R{}".format(number+3), device_name) #寫入裝置名稱
                worksheet.update_value("S{}".format(number+3), str(device_user[0])+","+str(device_user[1])) #寫入裝置使用者
                worksheet.update_value("T{}".format(number+3), mac_adderss) #寫入裝置網路卡號碼
                worksheet.update_value("U{}".format(number+3), net_name) #寫入裝置網路卡名稱
                worksheet.update_value("V{}".format(number+3), device_addr_IPv4) #寫入局域網IP(IPv4)
                worksheet.update_value("W{}".format(number+3), device_addr_IPv6) #寫入局域網IP(IPv6)
                if usingnow_of_ip(device_addr_IPv4) == "正在使用": #對局域網IP的不同版本進行測試
                    worksheet.update_value("X{}".format(number+3), "IPv4")
                elif usingnow_of_ip(device_addr_IPv6) == "正在使用":
                    worksheet.update_value("X{}".format(number+3), "IPv6")
                else:
                    worksheet.update_value("X{}".format(number+3), "None")
                if ip != "None": #如果沒有抓到IP
                    worksheet.update_value("Y{}".format(number+3), ip) #寫入網際網路(外網)IP
                    worksheet.update_value("Z{}".format(number+3), type_of_ip(ip)) #寫入外網IP的種類(IPv4/IPv6)
                    worksheet.update_value("AA{}".format(number+3), str(ip_data["lat"])+"/"+str(ip_data["lon"])) #寫入外網IP的座標
                    worksheet.update_value("AB{}".format(number+3), ip_data["city"]) #寫入外網IP所在的城市
                    worksheet.update_value("AC{}".format(number+3), ip_data["region"]) #寫入外網IP所在的省級區域
                    worksheet.update_value("AD{}".format(number+3), ip_data["country"]) #寫入外網IP所在的國家/地區
                    worksheet.update_value("AE{}".format(number+3), ip_data["timezone"]) #寫入外網IP所在的時區
                    worksheet.update_value("AF{}".format(number+3), ip_data["org"]) #寫入外網IP所在的網路服務提供公司
                    worksheet.update_value("AG{}".format(number+3), ip_data["as"]) #寫入外網IP所在的自治系統
            code_section_7_status = "〇"
        except:
            code_section_7_status = "✕" 
            worksheet.update_value("A{}".format(number+3), runtime+1) #寫入本次的運作次數
            if str(start_time).split(" ")[0] == str(worksheet.get_value("E{}".format(number+2))): #start_time的日期等於("D{}".format(number+1))的字串
                worksheet.update_value("B{}".format(number+3), date_number) #寫入同樣天數
                worksheet.update_value("C{}".format(number+3), int(worksheet.get_value("C{}".format(number+2)))+1) #取得前一格的天次數後加1再寫入
            else:
                worksheet.update_value("B{}".format(number+3), date_number+1) #寫入天數+1
                worksheet.update_value("C{}".format(number+3), "1") #寫入天次數為1
            worksheet.update_value("D{}".format(number+3), "study_gamer_code(Python)") #寫入現在執行的檔案
            worksheet.update_value("E{}".format(number+3), str(start_time).split(" ")[0].split("+")[0]) #寫入現在的時間
            worksheet.update_value("F{}".format(number+3), str(start_time).split(" ")[1].split("+")[0]) #寫入程式開始的時間
    worksheet.update_value("O{}".format(number+3), code_section_7_status) #寫入程式運作狀態 #系統訊息寫入
    worksheet.update_value("H{}".format(number+3), str(TWtime()[0]).split(" ")[1].split("+")[0]) #寫入程式執行的時間
    worksheet.update_value("G{}".format(number+3), str(TWtime()[0]-start_time)) #寫入程式結束的時間
