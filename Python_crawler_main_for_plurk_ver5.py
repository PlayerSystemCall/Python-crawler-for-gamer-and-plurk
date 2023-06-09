# -*- coding: utf-8 -*-
import os
import re
import json
import socket
import requests
import pygsheets
import Python_crawler_sub_ver1 as sub_program
from lxml import etree
from dotenv import load_dotenv
from fake_useragent import UserAgent
from datetime import datetime

load_dotenv()
Player_SystemCall_plurk_id = os.getenv("PLAYER_SYSTEMCALL_PLURK_ID")

#取得起始時間
start_timecall = sub_program.TWtime() #呼叫副程式TWtime()獲取時間
start_time = start_timecall[0] #取得程式開始的時間
ad_year_today = start_timecall[1] #台北時區的西元紀年
mg_year_today = start_timecall[2] #台北時區的民國紀年
date_today = start_timecall[3] #今天現在台北時區的西元日期
ad_year_yesterday = start_timecall[4] #台北時區的西元紀年
mg_year_yesterday = start_timecall[5] #台北時區的民國紀年
date_yesterday = start_timecall[6] #昨天台北時區的西元日期

code_section_1_status, times_1 = "✕", 0
while code_section_1_status == "✕" and times_1 < 3:
    try:
        #取得裝置資訊
        device_name = socket.getfqdn(socket.gethostname()).split(".")[0] #裝置名稱，從DNS連線網址中擷取第1段
        device_user = sub_program.get_user() #裝置使用者名稱和數量
        mac_adderss = sub_program.get_nic_data()[1]["mac"] #裝置網路卡號碼
        net_name = sub_program.get_nic_data()[0] #裝置網路卡名稱
        device_addr_IPv4 = sub_program.get_nic_data()[1]["ipv4"] #IPv4位址(內網IP)
        device_addr_IPv6 = sub_program.get_nic_data()[1]["ipv6"] #IPv6位址(內網IP)
        ip = sub_program.get_ip_and_version()[0] #取得裝置對外的IP
        if ip != None:
            ip_data = sub_program.get_ip_data(ip) #查詢IP所在地資料
        code_section_1_status = "〇"
    except:
        code_section_1_status = "✕"
        times_1 = times_1+1

code_section_2_status, times_2 = "✕", 0
while code_section_2_status == "✕" and times_2 < 3:
    try:
        #取得噗浪網頁原始碼，找到資料的的html區塊，取出資料
        plurk_url = "https://www.plurk.com/Player_SystemCall" #噗浪個人頁面
        plurk_statuscode = sub_program.go_to_web(plurk_url) #回傳網路狀態碼
        headers = {"User-Agent" : UserAgent().random} #設置http頭欄位，裡面夾帶瀏覽器識別標籤
        Go_to_plurk = requests.get(plurk_url, headers = headers, timeout = 60, allow_redirects = False, stream = True, verify = False) #對plurk_url夾帶headers發出GET請求，timeout為最長反應時間，allow_redirects為禁止重新定向，stream為強制解壓縮，verify為SSL憑證檢查功能       
        requests.packages.urllib3.disable_warnings() #關閉InsecureRequestWarning的顯示
        if plurk_statuscode == 200: #如果連線狀態正常
            Go_to_plurk.encoding = "utf-8" #指定網頁的編碼格式
            plurk_sourcecode = etree.HTML(Go_to_plurk.text) #取得網頁原始碼
            #個人頁面/朋友數
            xpath_on_web = "//span[@id='num_of_friends']/text()" #指定text()在網頁程式碼的位置(Xpath表達式)
            plurk_friend_number = int(plurk_sourcecode.xpath(xpath_on_web)[0]) #使用Xpath表達式提出，為plurk_friend_number串列的第一個元素
            #個人頁面/粉絲數
            xpath_on_web = "//span[@id='num_of_fans']/text()" #指定text()在網頁程式碼的位置(Xpath表達式)
            plurk_fan_number = int(plurk_sourcecode.xpath(xpath_on_web)[0])-plurk_friend_number #使用Xpath表達式提出，為plurk_fan_number串列的第一個元素，然後再與朋友數量相減為實際追蹤人數
            #個人頁面/人氣指數
            xpath_on_web = "//td[@id='profile_views']/text()" #指定text()在網頁程式碼的位置(Xpath表達式)
            plurk_allview_number_yesterday = int(plurk_sourcecode.xpath(xpath_on_web)[0]) #使用Xpath表達式提出，為plurk_allview_number_yesterday串列的第一個元素
            Go_to_plurk.close() #關閉對web_URL夾帶headers發出GET請求
        code_section_2_status = "〇"
    except:
        code_section_2_status = "✕"
        times_2 = times_2+1

plurk_data_get = {"fan" : {}, "friend" : {}, "other" : {}} #建立字典，並在裡面建立fan、friena和other3鍵，功能為放置從噗浪取得(昨天)的資料
    #id為噗浪ID，account為噗浪nick_name噗浪帳號，displayname為名字，regdate巴哈帳號建立日，lastondate巴哈最後登入日 #"id" : [], "account" : [], "displayname" :[], "regdate" : [], "lastondate" : []

code_section_3_status, times_3 = "✕", 0
while code_section_3_status == "✕" and times_3 < 3:
    try:
        #取得噗浪粉絲名單網頁原始碼，找到資料的的html區塊，取出資料
        plurk_fanlist_url = "https://www.plurk.com/Friends/getFansByOffset" #噗浪的粉絲名單請求網址
        headers = {"User-Agent" : UserAgent().random} #設置http頭欄位，裡面夾帶瀏覽器識別標籤
        data = {"user_id": Player_SystemCall_plurk_id, "offset":"0", "limit": "10000000"} #夾帶資料，個人ID、offset和單次送出來的人數
        Go_to_plurk_fanlist = requests.post(plurk_fanlist_url, headers = headers, data = data, timeout = 60, allow_redirects = False, stream = True, verify = False) #對plurk_fanlist_url夾帶headers和data發出POST請求，timeout為最長反應時間，allow_redirects為禁止重新定向，stream為強制解壓縮，verify為SSL憑證檢查功能       
        requests.packages.urllib3.disable_warnings() #關閉InsecureRequestWarning的顯示
        plurk_data_get["other"]["plurk_fan_idlist"] = [] #建立串列放粉絲ID
        plurk_data_get["other"]["plurk_fan_accountlist"] = [] #建立串列放粉絲帳號
        plurk_data_get["other"]["plurk_fan_displaynamelist"] = [] #建立串列放粉絲名字
        if Go_to_plurk_fanlist.status_code == 200: #如果網路正常
            plurk_fan_Datalist = json.loads(Go_to_plurk_fanlist.text) #將字串變成python的字典
            for i in range(len(plurk_fan_Datalist)):
                plurk_data_get["other"]["plurk_fan_idlist"].append(plurk_fan_Datalist[i]["id"]) #將粉絲ID加進和鍵plurk_fan_idlist相對應的串列
                plurk_data_get["other"]["plurk_fan_accountlist"].append(plurk_fan_Datalist[i]["nick_name"]) #將粉絲ID加進和鍵plurk_fan_accountlist相對應的串列
                plurk_data_get["other"]["plurk_fan_displaynamelist"].append(plurk_fan_Datalist[i]["display_name"]) #將unicode(代碼)加進和鍵plurk_fan_displaynamelist相對應的串列 
            Go_to_plurk_fanlist.close() #關閉對plurk_fanlist_url夾帶headers和data發出POST請求
        
        for i in range(len(plurk_data_get["other"]["plurk_fan_accountlist"])): #依照順序數字建立字典，再將資訊分別塞進去
            plurk_data_get["fan"][i] = {} 
            plurk_data_get["fan"][i]["id"] = plurk_data_get["other"]["plurk_fan_idlist"][i]
            plurk_data_get["fan"][i]["account"] = plurk_data_get["other"]["plurk_fan_accountlist"][i]
            plurk_data_get["fan"][i]["displayname"] = plurk_data_get["other"]["plurk_fan_displaynamelist"][i]
            plurk_fan_url = "https://www.plurk.com/{}".format(plurk_data_get["fan"][i]["account"]) #組合成噗浪粉絲的個人網址
            plurk_fan_statuscode = sub_program.go_to_web(plurk_fan_url) #回傳狀態碼
            headers = {"User-Agent" : UserAgent().random} #設置http頭欄位，裡面夾帶瀏覽器識別標籤
            Go_to_plurk_fan = requests.get(plurk_fan_url, headers = headers, timeout = 60, allow_redirects = False, stream = True, verify = False) #對plurk_fan_url夾帶headers發出GET請求，timeout為最長反應時間，allow_redirects為禁止重新定向，stream為強制解壓縮，verify為SSL憑證檢查功能       
            requests.packages.urllib3.disable_warnings() #關閉InsecureRequestWarning的顯示
            Go_to_plurk_fan.encoding = "utf-8" #指定網頁的編碼格式
            plurk_fan_sourcecode = etree.HTML(Go_to_plurk_fan.text) #取得網頁原始碼
            number = 0 #編號
            if plurk_fan_statuscode == 200: #如果網路狀態碼正常
                #個人頁面/帳號註冊日和最後登入日
                try: #測試第一個script的起始順序
                    number = number+1
                    jscode_1st = plurk_fan_sourcecode.xpath("/html/body/script[{}]/text()".format(number))[0] #取得text()在網頁程式碼的位置(Xpath表達式)
                except:
                    number = number+1
                    jscode_1st = plurk_fan_sourcecode.xpath("/html/body/script[{}]/text()".format(number))[0] #取得text()在網頁程式碼的位置(Xpath表達式)
                number = number+2
                xpath_on_web = "/html/body/script[{}]/text()".format(number) #指定text()在網頁程式碼的位置(Xpath表達式)
                jscode_2nd = plurk_fan_sourcecode.xpath(xpath_on_web)[0] #使用Xpath表達式提出，為jscode_2nd串列的第一個元素
                plurk_data_get["fan"][i]["lastondate"], plurk_data_get["fan"][i]["regdate"] = re.findall("\d\d\d\d-\d\d-\d\d \d\d:\d\d:\d\d", jscode_2nd) #從jscode_2nd抽出日期
            else: #應該不會有這情況
                plurk_data_get["fan"][i]["regdate"] = "版面無區塊"
                plurk_data_get["fan"][i]["lastondate"] = "版面無區塊" 
            Go_to_plurk_fan.close() #關閉對plurk_fan_url夾帶headers發出GET請求
        code_section_3_status = "〇"
    except:
        code_section_3_status = "✕"
        times_3 = times_3+1

code_section_4_status, times_4 = "✕", 0
while code_section_4_status == "✕" and times_4 < 3:
    try:
        #取得噗浪朋友名單網頁原始碼，找到資料的的html區塊，取出資料
        plurk_friendlist_url = "https://www.plurk.com/Friends/getFriendsByOffset" #噗浪的朋友名單請求網址
        headers = {"User-Agent" : UserAgent().random} #設置http頭欄位，裡面夾帶瀏覽器識別標籤
        data = {"user_id": Player_SystemCall_plurk_id, "offset":"0", "limit": "10000000"} #夾帶資料，個人ID、offset和單次送出來的人數
        Go_to_plurk_friendlist = requests.post(plurk_friendlist_url, headers = headers, data = data,timeout = 60, allow_redirects = False, stream = True, verify = False) #對plurk_friendlist_url夾帶headers發出GET請求，timeout為最長反應時間，allow_redirects為禁止重新定向，stream為強制解壓縮，verify為SSL憑證檢查功能       
        requests.packages.urllib3.disable_warnings() #關閉InsecureRequestWarning的顯示
        plurk_data_get["other"]["plurk_friend_idlist"] = [] #建立串列放好友ID
        plurk_data_get["other"]["plurk_friend_accountlist"] = [] #建立串列放好友帳號
        plurk_data_get["other"]["plurk_friend_displaynamelist"] = [] #建立串列放好友名字
        if Go_to_plurk_friendlist.status_code == 200: #如果網路正常
            plurk_friend_Datalist = json.loads(Go_to_plurk_friendlist.text) #取得網頁原始碼
            for i in range(len(plurk_friend_Datalist)):
                plurk_data_get["other"]["plurk_friend_idlist"].append(plurk_friend_Datalist[i]["id"]) #將朋友ID加進和鍵plurk_friend_idlist相對應的串列
                plurk_data_get["other"]["plurk_friend_accountlist"].append(plurk_friend_Datalist[i]["nick_name"]) #將朋友ID加進和鍵plurk_friend_accountlist相對應的串列
                plurk_data_get["other"]["plurk_friend_displaynamelist"].append(plurk_friend_Datalist[i]["display_name"]) #將unicode(代碼)加進和鍵plurk_friend_displaynamelist相對應的串列 
            Go_to_plurk_friendlist.close() #關閉對plurk_friendlist_url夾帶headers和data發出POST請求
        
        for i in range(len(plurk_data_get["other"]["plurk_friend_accountlist"])): #依照順序數字建立字典，再將資訊分別塞進去
            plurk_data_get["friend"][i] = {}
            plurk_data_get["friend"][i]["id"] = plurk_data_get["other"]["plurk_friend_idlist"][i]
            plurk_data_get["friend"][i]["account"] = plurk_data_get["other"]["plurk_friend_accountlist"][i]
            plurk_data_get["friend"][i]["displayname"] = plurk_data_get["other"]["plurk_friend_displaynamelist"][i]
            plurk_friend_url = "https://www.plurk.com/{}".format(plurk_data_get["friend"][i]["account"]) #組合成噗浪朋友的個人網址
            plurk_friend_statuscode = sub_program.go_to_web(plurk_friend_url) #回傳狀態碼
            headers = {"User-Agent" : UserAgent().random} #設置http頭欄位，裡面夾帶瀏覽器識別標籤
            Go_to_plurk_friend = requests.get(plurk_friend_url, headers = headers, timeout = 60, allow_redirects = False, stream = True, verify = False) #對plurk_friend_url夾帶headers發出GET請求，timeout為最長反應時間，allow_redirects為禁止重新定向，stream為強制解壓縮，verify為SSL憑證檢查功能       
            requests.packages.urllib3.disable_warnings() #關閉InsecureRequestWarning的顯示
            Go_to_plurk_friend.encoding = "utf-8" #指定網頁的編碼格式
            plurk_friend_sourcecode = etree.HTML(Go_to_plurk_friend.text) #取得網頁原始碼
            number = 0 #編號
            if plurk_friend_statuscode == 200:
                #個人頁面/人氣指數
                try: #測試第一個script的起始順序
                    number = number+1
                    jscode_1st = plurk_friend_sourcecode.xpath("/html/body/script[{}]/text()".format(number))[0] #取得text()在網頁程式碼的位置(Xpath表達式)
                except:
                    number = number+1
                    jscode_1st = plurk_friend_sourcecode.xpath("/html/body/script[{}]/text()".format(number))[0] #取得text()在網頁程式碼的位置(Xpath表達式)
                number = number+2
                xpath_on_web = "/html/body/script[{}]/text()".format(number) #指定text()在網頁程式碼的位置(Xpath表達式)
                jscode_2nd = plurk_friend_sourcecode.xpath(xpath_on_web)[0] #使用Xpath表達式提出，為jscode_2nd串列的第一個元素
                jscode_to_xml = re.findall("\d\d\d\d-\d\d-\d\d \d\d:\d\d:\d\d", jscode_2nd) #將results_on_web轉成字串再轉換成xml標記的文本
                plurk_data_get["friend"][i]["lastondate"], plurk_data_get["friend"][i]["regdate"] = re.findall("\d\d\d\d-\d\d-\d\d \d\d:\d\d:\d\d", jscode_2nd) #從jscode_2nd抽出日期
            else: #應該不會有這情況
                plurk_data_get["friend"][i]["regdate"] = "版面無區塊"
                plurk_data_get["friend"][i]["lastondate"] = "版面無區塊"
            Go_to_plurk_friend.close() #關閉對plurk_friend_url夾帶headers發出GET請求
        code_section_4_status = "〇"
    except:
        code_section_4_status = "✕"
        times_4 = times_4+1

open_googlesheets_status, times_open_googlesheets = False, 0
while open_googlesheets_status == False and times_open_googlesheets < 3:
    try:
        #開啟試算表
        certificate = pygsheets.authorize(service_account_env_var = "GOOGLE_SHEETS_API_KEY") #取得位置在同層級目錄的Google sheets API憑證
        googlesheets_url = os.getenv("GOOGLESHEETS_URL") #有spreadsheetId的google sheets網址
        open_googlesheets = certificate.open_by_url(googlesheets_url) #開啟Google sheets
        open_googlesheets_status = True
    except:
        open_googlesheets_status = False
        times_open_googlesheets = times_open_googlesheets+1

if open_googlesheets_status == True:
    basic_status, times_basic = False, 0
    while basic_status == False and times_basic < 3:
        try:
            #獲取運作天數
            worksheet = open_googlesheets.worksheet('id', 1888051586) #以sheetId定位試算表位置為倒數第2張的"系統訊息"
            number = 0 #天數寫了總共幾格的計數
            days = worksheet.get_col(2)[2:] #將得到的字串加入days串列
            while days[number] != "": #一直運作直到得到的字串為「完全空白」
                number = number+1
            days = list(set(days))[1:] #將days串列中重複的元素變成1個再組成一個新的days串列
            date_number = len(days) #天數為days元素的數量
            
            try:
                runtime = int(worksheet.get_value("A{}".format(number+2))) #獲取運作次數，number+2為「完全空白」前一格的位置
            except:
                runtime = 0
            basic_status = True
        except:
            basic_status = False
            times_basic = times_basic+1
    
    if basic_status == True:
        code_section_5_status, time_5 = "✕", 0
        while code_section_5_status == "✕" and time_5 < 3:
            try: #寫入試算表1："人氣紀錄"
                worksheet = open_googlesheets.worksheet('id', 0) #以sheetId定位試算表位置為第1張的"人氣紀錄表"
                
                writesit = 3 #日期被寫下的格數
                A_col = worksheet.get_col(1)[2:]
                while A_col[writesit-3] != "": #一直運作直到得到的字串為「完全空白」
                    writesit = writesit+1
                
                if writesit > 3 and str(datetime.strptime(worksheet.get_value("A{}".format(writesit-1))+worksheet.get_value("C{}".format(writesit-1)), "%Y年%m月%d日")) == str(datetime.strptime(ad_year_yesterday+date_yesterday, "%Y年%m月%d日")): #噗浪資料被寫下的格數 #如果最後一天的日期等於「昨天的台北時間」
                    plurk_writesit = writesit-1
                elif writesit > 3 and str(datetime.strptime(worksheet.get_value("A{}".format(writesit-1))+worksheet.get_value("C{}".format(writesit-1)), "%Y年%m月%d日")) == str(datetime.strptime(ad_year_today+date_today, "%Y年%m月%d日")): #如果最後一天的日期等於「今天的台北時間」
                    plurk_writesit = writesit-2
                elif writesit == 3 and worksheet.get_value("A{}".format(writesit-1))+worksheet.get_value("C{}".format(writesit-1)) == "西元紀年日期":
                    plurk_writesit = writesit
                
                if plurk_writesit != writesit: #取得該月最後一天的日期
                    month_lastday = sub_program.lastday_of_month(int(worksheet.get_value("A{}".format(writesit-2)).split("年")[0]), int(worksheet.get_value("C{}".format(writesit-2)).split("月")[0]))
                else:
                    month_lastday = sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(date_yesterday.split("月")[0]))
                
                try:
                    plurk_yesterday_viewnum = plurk_allview_number_yesterday-int(worksheet.get_value("M{}".format(plurk_writesit-1))) #用昨天全部的噗浪瀏覽數減前天全部的噗浪瀏覽數為昨天單日的瀏覽數
                except:
                    plurk_yesterday_viewnum = plurk_allview_number_yesterday
                if plurk_writesit != writesit:
                    if str(datetime.strptime(worksheet.get_value("A{}".format(writesit-1))+worksheet.get_value("C{}".format(writesit-1)), "%Y年%m月%d日")) != str(datetime.strptime(ad_year_today+date_today, "%Y年%m月%d日")): #如果最後一天的日期不等於今天的日期
                        worksheet.update_value("A{}".format(writesit-1), ad_year_yesterday) #寫入昨日的西元紀年
                        worksheet.update_value("B{}".format(writesit-1), mg_year_yesterday) #寫入昨日的民國紀年
                        worksheet.update_value("C{}".format(writesit-1), date_yesterday) #寫入昨日的日期
                        worksheet.update_value("A{}".format(writesit), ad_year_today) #寫入今日的西元紀年
                        worksheet.update_value("B{}".format(writesit), mg_year_today) #寫入今日的民國紀年
                        worksheet.update_value("C{}".format(writesit), date_today) #寫入今日的日期
                elif plurk_writesit == writesit:
                    worksheet.update_value("A{}".format(writesit), ad_year_yesterday) #寫入昨日的西元紀年
                    worksheet.update_value("B{}".format(writesit), mg_year_yesterday) #寫入昨日的民國紀年
                    worksheet.update_value("C{}".format(writesit), date_yesterday) #寫入昨日的日期
                    worksheet.update_value("A{}".format(writesit+1), ad_year_today) #寫入今日的西元紀年
                    worksheet.update_value("B{}".format(writesit+1), mg_year_today) #寫入今日的民國紀年
                    worksheet.update_value("C{}".format(writesit+1), date_today) #寫入今日的日期
                if str(datetime.strptime(worksheet.get_value("A{}".format(plurk_writesit))+worksheet.get_value("C{}".format(plurk_writesit)), "%Y年%m月%d日")) == str(datetime.strptime(ad_year_yesterday+date_yesterday, "%Y年%m月%d日")): #如果試算表上所寫的「昨天的日期」等於「ad_year+date_yesterday」
                    worksheet.update_value("J{}".format(plurk_writesit), plurk_friend_number) #寫入噗浪好友人數
                    worksheet.update_value("K{}".format(plurk_writesit), plurk_fan_number) #寫入噗浪粉絲人數
                    worksheet.update_value("L{}".format(plurk_writesit), plurk_yesterday_viewnum) #寫入噗浪今天的訪客數
                    worksheet.update_value("M{}".format(plurk_writesit), "=1015+SUM($L$66:$L{})".format(plurk_writesit)) #寫入噗浪訪客總數(不直接輸入數字，而是函數加總)
                    if int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0]) == 12:
                        if (plurk_writesit > (sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 1)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 2)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 3)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 4)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 5)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 6)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 7)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 8)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 9)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 10)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 11)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 12)-1)) and (int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[1].split("日")[0]) == month_lastday):
                            worksheet.update_values("N{}".format(plurk_writesit-15), [["本年總和"], ["=SUM($L{}:$L{})".format(plurk_writesit+1-(sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 1)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 2)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 3)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 4)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 5)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 6)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 7)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 8)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 9)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 10)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 11)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 12)))], ["本年日平均"], ["=AVERAGE(L{}:L{})".format(plurk_writesit+1-(sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 1)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 2)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 3)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 4)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 5)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 6)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 7)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 8)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 9)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 10)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 11)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 12)))]]) #寫入本年總和和日平均
                        elif (plurk_writesit < (sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 1)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 2)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 3)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 4)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 5)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 6)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 7)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 8)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 9)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 10)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 11)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), 12)-1)) and (int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[1].split("日")[0]) == month_lastday):
                            start_point = 3
                            while worksheet.get_value("J{}".format(start_point)) == "✕": #一直運作直到得到的字串為「完全空白」
                                start_point = start_point+1
                            if start_point < 18:
                                worksheet.update_values("N3", [["本年總和"], ["=SUM($L{}:$L{})".format(start_point, plurk_writesit)], ["本年日平均"], ["=AVERAGE(F{}:F{})".format(start_point, plurk_writesit)]]) #寫入本年總和和日平均
                            else:
                                worksheet.update_values("N{}".format(plurk_writesit-15), [["本年總和"], ["=SUM($L{}:$L{})".format(start_point, plurk_writesit)], ["本年日平均"], ["=AVERAGE(F{}:F{})".format(start_point, plurk_writesit)]]) #寫入本年總和和日平均
                    if int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0]) in [6, 12]:
                        if (plurk_writesit > (sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-5)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-4)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-3)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-2)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-1)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0]))-1)) and (int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[1].split("日")[0]) == month_lastday):
                            worksheet.update_values("N{}".format(plurk_writesit-11), [["本半年總和"], ["=SUM($L{}:$L{})".format(plurk_writesit+1-(sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-5)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-4)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-3)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-2)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-1)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0]))), plurk_writesit)], ["本半年日平均"], ["=AVERAGE(L{}:L{})".format(plurk_writesit+1-(sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-5)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-4)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-3)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-2)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-1)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0]))), plurk_writesit)]]) #寫入本季總和和日平均
                        elif (plurk_writesit < (sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-5)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-4)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-3)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-2)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-1)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0]))-1)) and (int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[1].split("日")[0]) == month_lastday):
                            start_point = 3
                            while worksheet.get_value("J{}".format(start_point)) == "✕": #一直運作直到得到的字串為「完全空白」
                                start_point = start_point+1
                            if start_point < 14 and int(date_yesterday.split("月")[0]) == 12:
                                worksheet.update_values("N7", [["本半年總和"], ["=SUM($L{}:$L{})".format(start_point, plurk_writesit)], ["本半年日平均"], ["=AVERAGE(F{}:F{})".format(start_point, plurk_writesit)]])
                            elif start_point < 14 and int(date_yesterday.split("月")[0]) == 6:
                                worksheet.update_values("N3", [["本半年總和"], ["=SUM($L{}:$L{})".format(start_point, plurk_writesit)], ["本半年日平均"], ["=AVERAGE(F{}:F{})".format(start_point, plurk_writesit)]])
                            else:
                                worksheet.update_values("N{}".format(plurk_writesit-11), [["本半年總和"], ["=SUM($L{}:$L{})".format(start_point, plurk_writesit)], ["本半年日平均"], ["=AVERAGE(F{}:F{})".format(start_point, plurk_writesit)]])
                    if int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0]) in [3, 6, 9, 12]:
                        if (plurk_writesit > (sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-2)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-1)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0]))-1)) and (int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[1].split("日")[0]) == month_lastday):
                            worksheet.update_values("N{}".format(plurk_writesit-7), [["本季總和"], ["=SUM($L{}:$L{})".format(plurk_writesit+1-(sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-2)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-1)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0]))), plurk_writesit)], ["本季日平均"], ["=AVERAGE(L{}:L{})".format(plurk_writesit+1-(sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-2)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-1)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0]))), plurk_writesit)]]) #寫入本季總和和日平均
                        elif (plurk_writesit < (sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-2)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0])-1)+sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0]))-1)) and (int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[1].split("日")[0]) == month_lastday):
                            start_point = 3
                            while worksheet.get_value("J{}".format(start_point)) == "✕": #一直運作直到得到的字串為「完全空白」
                                start_point = start_point+1
                            if start_point < 11 and int(date_yesterday.split("月")[0]) == 12:
                                worksheet.update_values("N11", [["本季總和"], ["=SUM($L{}:$L{})".format(start_point, plurk_writesit)], ["本季日平均"], ["=AVERAGE(F{}:F{})".format(start_point, plurk_writesit)]])
                            elif start_point < 11 and int(date_yesterday.split("月")[0]) == 6:
                                worksheet.update_values("N7", [["本季總和"], ["=SUM($L{}:$L{})".format(start_point, plurk_writesit)], ["本季日平均"], ["=AVERAGE(F{}:F{})".format(start_point, plurk_writesit)]])
                            elif start_point < 11 and (int(date_yesterday.split("月")[0]) in [3, 9]):
                                worksheet.update_values("N3", [["本季總和"], ["=SUM($L{}:$L{})".format(start_point, plurk_writesit)], ["本季日平均"], ["=AVERAGE(F{}:F{})".format(start_point, plurk_writesit)]])
                            else:
                                worksheet.update_values("N{}".format(plurk_writesit-7), [["本季總和"], ["=SUM($L{}:$L{})".format(start_point, plurk_writesit)], ["本季日平均"], ["=AVERAGE(F{}:F{})".format(start_point, plurk_writesit)]])
                    if (plurk_writesit > (sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0]))-1)) and (int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[1].split("日")[0]) == month_lastday): #如果試算表日期等於月的最後一天
                        worksheet.update_values("N{}".format(plurk_writesit-3), [["本月總和"], ["=SUM($L{}:$L{})".format(plurk_writesit+1-(sub_program.lastday_of_month(int(worksheet.get_value("A{}".format(plurk_writesit)).split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0]))), plurk_writesit)], ["本月日平均"], ["=AVERAGE(L{}:L{})".format(plurk_writesit+1-(sub_program.lastday_of_month(int(worksheet.get_value("A{}".format(plurk_writesit)).split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0]))), plurk_writesit)]]) #寫入本月總和和日平均
                    elif (plurk_writesit < (sub_program.lastday_of_month(int(ad_year_yesterday.split("年")[0]), int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[0]))-1)) and (int(worksheet.get_value("C{}".format(plurk_writesit)).split("月")[1].split("日")[0]) == month_lastday):
                        start_point = 3
                        while worksheet.get_value("J{}".format(start_point)) == "✕": #一直運作直到得到的字串為「完全空白」
                            start_point = start_point+1
                        if start_point < 7 and int(date_yesterday.split("月")[0]) == 12:
                            worksheet.update_values("N15", [["本月總和"], ["=SUM($L{}:$L{})".format(start_point, plurk_writesit)], ["本月日平均"], ["=AVERAGE(L{}:L{})".format(start_point, plurk_writesit)]]) #寫入本月總和和日平均
                        elif start_point < 7 and int(date_yesterday.split("月")[0]) == 6:
                            worksheet.update_values("N11", [["本月總和"], ["=SUM($L{}:$L{})".format(start_point, plurk_writesit)], ["本月日平均"], ["=AVERAGE(L{}:L{})".format(start_point, plurk_writesit)]])
                        elif start_point < 7 and (int(date_yesterday.split("月")[0]) in [3, 9]):
                            worksheet.update_values("N7", [["本月總和"], ["=SUM($L{}:$L{})".format(start_point, plurk_writesit)], ["本月日平均"], ["=AVERAGE(L{}:L{})".format(start_point, plurk_writesit)]])
                        elif start_point < 7 and (int(date_yesterday.split("月")[0]) in [1, 2, 4, 5, 7, 8, 10, 11]):
                            worksheet.update_values("N3", [["本月總和"], ["=SUM($L{}:$L{})".format(start_point, plurk_writesit)], ["本月日平均"], ["=AVERAGE(L{}:L{})".format(start_point, plurk_writesit)]])
                        else:
                            worksheet.update_values("N{}".format(plurk_writesit-3), [["本月總和"], ["=SUM($L{}:$L{})".format(start_point, plurk_writesit)], ["本月日平均"], ["=AVERAGE(L{}:L{})".format(start_point, plurk_writesit)]])
                worksheet.update_values("N3", [["至今總人氣數"], ["=1015+SUM($L$66:$L{})".format(plurk_writesit)]])
                worksheet.update_values("N55", [["至今總人氣數"], ["=1015+SUM($L$66:$L{})".format(plurk_writesit)]])
                code_section_5_status = "〇"
            except:
                code_section_5_status = "✕"
                time_5 = time_5+1
        
        code_section_6A_status, time_6A = "✕", 0
        while code_section_6A_status == "✕" and time_6A < 3:
            try:
                #寫入試算表3："噗浪追蹤名單"
                worksheet = open_googlesheets.worksheet('id', 1068388458) #以sheetId定位試算表位置為"噗浪追蹤名單"
                
                plurk_data_last= {"fan" : {}, "friend" : {}, "other" : {}} #建立字典，並在裡面建立fan、friena和other3鍵，功能為放置從試算表取得的噗浪前一筆(前天)資料
                    #id為噗浪ID，account為噗浪nick_name噗浪帳號，displayname為名字，regdate巴哈帳號建立日，lastondate巴哈最後登入日 #"id" : [], "account" : [], "displayname" :[], "regdate" : [], "lastondate" : []
                
                if code_section_3_status == "〇":
                    plurk_fan_number_last = 0 #下方迴圈的計數
                    plurk_data_last["other"]["plurk_fan_accountlist"] = [] #建立與鍵對應的值，此值為串列，功能為放置account
                    plurk_data_last["other"]["plurk_fan_displaynamelist"] = [] #建立與鍵對應的值，此值為串列，功能為放置displayname
                    while worksheet.get_value("B{}".format(plurk_fan_number_last+3)) != "": #一直運作直到得到的字串為「完全空白」
                        plurk_data_last["fan"][plurk_fan_number_last] = {} #以plurk_fan_number_last為鍵建立對應的值，此值為字典
                        plurk_data_last["fan"][plurk_fan_number_last]["account"] = worksheet.get_value("B{}".format(plurk_fan_number_last+3)) #建立鍵，自試算表的指定位置讀取值
                        plurk_data_last["fan"][plurk_fan_number_last]["displayname"] = worksheet.get_value("C{}".format(plurk_fan_number_last+3))
                        plurk_data_last["fan"][plurk_fan_number_last]["regdate"] = worksheet.get_value("D{}".format(plurk_fan_number_last+3))
                        plurk_data_last["fan"][plurk_fan_number_last]["lastondate"] = worksheet.get_value("E{}".format(plurk_fan_number_last+3))
                        plurk_data_last["fan"][plurk_fan_number_last]["startfollowdate"] = worksheet.get_value("F{}".format(plurk_fan_number_last+3))
                        plurk_data_last["fan"][plurk_fan_number_last]["followday"] = worksheet.get_value("G{}".format(plurk_fan_number_last+3))
                        plurk_data_last["fan"][plurk_fan_number_last]["endfollowdate"] = worksheet.get_value("H{}".format(plurk_fan_number_last+3))
                        plurk_data_last["other"]["plurk_fan_accountlist"].append(plurk_data_last["fan"][plurk_fan_number_last]["account"]) #將account這個鍵的值加入串列
                        plurk_data_last["other"]["plurk_fan_displaynamelist"].append(plurk_data_last["fan"][plurk_fan_number_last]["displayname"]) #將displayname這個鍵的值加入串列
                        plurk_fan_number_last = plurk_fan_number_last+1
                
                if code_section_4_status == "〇":
                    plurk_friend_number_last = 0 #下方迴圈的計數
                    plurk_data_last["other"]["plurk_friend_accountlist"] = [] #建立與鍵對應的值，此值為串列，功能為放置account
                    plurk_data_last["other"]["plurk_friend_displaynamelist"] = [] #建立與鍵對應的值，此值為串列，功能為放置displayname
                    while worksheet.get_value("I{}".format(plurk_friend_number_last+3)) != "": #一直運作直到得到的字串為「完全空白」
                        plurk_data_last["friend"][plurk_friend_number_last] = {} #以plurk_friend_number_last為鍵建立對應的值，此值為字典
                        plurk_data_last["friend"][plurk_friend_number_last]["account"] = worksheet.get_value("I{}".format(plurk_friend_number_last+3)) #建立鍵，自試算表的指定位置讀取值
                        plurk_data_last["friend"][plurk_friend_number_last]["displayname"] = worksheet.get_value("J{}".format(plurk_friend_number_last+3))
                        plurk_data_last["friend"][plurk_friend_number_last]["regdate"] = worksheet.get_value("K{}".format(plurk_friend_number_last+3))
                        plurk_data_last["friend"][plurk_friend_number_last]["lastondate"] = worksheet.get_value("L{}".format(plurk_friend_number_last+3))
                        plurk_data_last["friend"][plurk_friend_number_last]["startfollowdate"] = worksheet.get_value("M{}".format(plurk_friend_number_last+3))
                        plurk_data_last["friend"][plurk_friend_number_last]["followday"] = worksheet.get_value("N{}".format(plurk_friend_number_last+3))
                        plurk_data_last["friend"][plurk_friend_number_last]["endfollowdate"] = worksheet.get_value("O{}".format(plurk_friend_number_last+3))
                        plurk_data_last["other"]["plurk_friend_accountlist"].append(plurk_data_last["friend"][plurk_friend_number_last]["account"]) #將account這個鍵的值加入串列
                        plurk_data_last["other"]["plurk_friend_displaynamelist"].append(plurk_data_last["friend"][plurk_friend_number_last]["displayname"]) #將displayname這個鍵的值加入串列
                        plurk_friend_number_last = plurk_friend_number_last+1
                code_section_6A_status = "〇"
            except:
                code_section_6A_status = "✕"
                time_6A = time_6A+1
                
        code_section_6B_status, time_6B = "✕", 0
        while code_section_6B_status == "✕" and time_6B < 3:
            try:
                plurk_data_new = {"fan" : {}, "friend" : {}} #建立字典，並在裡面建立fan、friena和other3鍵，功能為放置經過比對後，今天要寫入試算表的(昨天)資料
                    #id為噗浪ID，account為噗浪nick_name噗浪帳號，displayname為名字，regdate巴哈帳號建立日，lastondate巴哈最後登入日 #"id" : [], "account" : [], "displayname" :[], "regdate" : [], "lastondate" : []
                
                if code_section_3_status == "〇":
                    plurk_fan_number_new = 0 #下方迴圈的計數
                    plurk_fan_string = "plurk_fan-" #沒有起始日則在表格填上這個字串
                    for i in range(len(plurk_data_get["fan"])): #範圍為fan鍵的值(字典)內所含有的數量
                        plurk_data_new["fan"][plurk_fan_number_new] = {} #以plurk_fan_number_new為鍵建立對應的值，此值為字典
                        try:
                            if plurk_data_get["fan"][i]["account"] in plurk_data_last["other"]["plurk_fan_accountlist"]: #如果從噗浪取得的粉絲account的值有在串列裡
                                plurk_data_new["fan"][plurk_fan_number_new]["status"] = "old" #標記狀態為old
                                plurk_data_new["fan"][plurk_fan_number_new]["account"] = plurk_data_get["fan"][i]["account"] #將從噗浪取得的粉絲account、displayname、建立日和最後上線日記錄進plurk_data_new
                                plurk_data_new["fan"][plurk_fan_number_new]["displayname"] = plurk_data_get["fan"][i]["displayname"]
                                plurk_data_new["fan"][plurk_fan_number_new]["regdate"] = plurk_data_get["fan"][i]["regdate"]
                                plurk_data_new["fan"][plurk_fan_number_new]["lastondate"] = plurk_data_get["fan"][i]["lastondate"]
                                plurk_data_new["fan"][plurk_fan_number_new]["startfollowdate"] = plurk_data_last["fan"][plurk_data_last["other"]["plurk_fan_accountlist"].index(plurk_data_new["fan"][plurk_fan_number_new]["account"])]["startfollowdate"] #尋找指定序號後，從前一筆(前天)紀錄取得開始追蹤的日期
                                plurk_data_new["fan"][plurk_fan_number_new]["followday"] = sub_program.get_followday(number, start_time, i, plurk_data_new["fan"][plurk_fan_number_new]["status"], plurk_data_new["fan"][plurk_fan_number_new]["startfollowdate"], plurk_data_last["fan"][plurk_data_last["other"]["plurk_fan_accountlist"].index(plurk_data_new["fan"][plurk_fan_number_new]["account"])]["followday"], plurk_fan_string, 0, worksheet) #傳入副程式number, start_time, i, status, startfollowdate, followday, string, endfollowdate這些參數，取得追蹤天數
                                plurk_data_new["fan"][plurk_fan_number_new]["endfollowdate"] = "" #結束時間為空白
                                plurk_fan_number_new = plurk_fan_number_new+1
                            else:
                                del plurk_data_new["fan"][plurk_fan_number_new] #如果從噗浪取得的粉絲account的值沒有在串列裡，刪除以plurk_fan_number_new為名的鍵
                        except:
                            del plurk_data_new["fan"][plurk_fan_number_new] #如果從噗浪取得的粉絲account的值有在串列裡，刪除以plurk_fan_number_new為名的鍵
                    
                    for j in range(len(plurk_data_get["fan"])): #範圍為fan鍵的值(字典)內所含有的數量
                        plurk_data_new["fan"][plurk_fan_number_new] = {} #以plurk_fan_number_new為鍵建立對應的值，此值為字典
                        try:
                            if plurk_data_get["fan"][j]["account"] not in plurk_data_last["other"]["plurk_fan_accountlist"]: #如果從噗浪取得的粉絲account的值沒有在串列裡
                                plurk_data_new["fan"][plurk_fan_number_new]["status"] = "new" #標記狀態為new
                                plurk_data_new["fan"][plurk_fan_number_new]["account"] = plurk_data_get["fan"][j]["account"] #將從噗浪取得的粉絲account、displayname、建立日和最後上線日記錄進plurk_data_new
                                plurk_data_new["fan"][plurk_fan_number_new]["displayname"] = plurk_data_get["fan"][j]["displayname"]
                                plurk_data_new["fan"][plurk_fan_number_new]["regdate"] = plurk_data_get["fan"][j]["regdate"]
                                plurk_data_new["fan"][plurk_fan_number_new]["lastondate"] = plurk_data_get["fan"][j]["lastondate"]
                                plurk_data_new["fan"][plurk_fan_number_new]["startfollowdate"] = str(start_time.year)+"-"+str(start_time.month)+"-"+str(int(start_time.day)-1)+" 0:00:00" #以昨天為開始追蹤日
                                plurk_data_new["fan"][plurk_fan_number_new]["followday"] = "0日0時0分0秒" #第0天
                                plurk_data_new["fan"][plurk_fan_number_new]["endfollowdate"] = "" #結束時間為空白
                                plurk_fan_number_new = plurk_fan_number_new+1
                            else:
                                del plurk_data_new["fan"][plurk_fan_number_new] #如果從噗浪取得的粉絲account的值有在串列裡，刪除以plurk_fan_number_new為名的鍵
                        except:
                            del plurk_data_new["fan"][plurk_fan_number_new] #如果從噗浪取得的粉絲account的值有在串列裡，刪除以plurk_fan_number_new為名的鍵
                    
                    for k in range(len(plurk_data_last["fan"])): #範圍為fan鍵的值(字典)內所含有的數量
                        plurk_data_new["fan"][plurk_fan_number_new] = {} #以plurk_fan_number_new為鍵建立對應的值，此值為字典
                        try:
                            if plurk_data_last["fan"][k]["account"] not in plurk_data_get["other"]["plurk_fan_accountlist"]: #如果從試算表取得的粉絲account的值沒有在串列裡
                                plurk_data_new["fan"][plurk_fan_number_new]["status"] = "leave" #標記狀態為leave
                                plurk_data_new["fan"][plurk_fan_number_new]["account"] = plurk_data_last["fan"][k]["account"] #將從噗浪取得的粉絲account、displayname、建立日和最後上線日記錄進plurk_data_new
                                plurk_data_new["fan"][plurk_fan_number_new]["displayname"] = plurk_data_last["fan"][k]["displayname"]
                                plurk_data_new["fan"][plurk_fan_number_new]["regdate"] = plurk_data_last["fan"][k]["regdate"]
                                plurk_data_new["fan"][plurk_fan_number_new]["lastondate"] = plurk_data_last["fan"][k]["lastondate"]
                                plurk_data_new["fan"][plurk_fan_number_new]["startfollowdate"] = plurk_data_last["fan"][plurk_data_last["other"]["plurk_fan_accountlist"].index(plurk_data_new["fan"][plurk_fan_number_new]["account"])]["startfollowdate"] #尋找指定序號後，從前一筆(前天)紀錄取得開始追蹤的日期
                                if plurk_data_last["fan"][plurk_data_last["other"]["plurk_fan_accountlist"].index(plurk_data_new["fan"][plurk_fan_number_new]["account"])]["endfollowdate"] == "": #如果日期等於空白
                                    plurk_data_new["fan"][plurk_fan_number_new]["endfollowdate"] = str(start_time.year)+"-"+str(start_time.month)+"-"+str(int(start_time.day)-1) #登記昨天為結束追蹤的日期
                                else: #如果日期不等於空白
                                    plurk_data_new["fan"][plurk_fan_number_new]["endfollowdate"] = plurk_data_last["fan"][plurk_data_last["other"]["plurk_fan_accountlist"].index(plurk_data_new["fan"][plurk_fan_number_new]["account"])]["endfollowdate"] #尋找指定序號後，從前一筆(前天)紀錄取得結束追蹤的日期
                                plurk_data_new["fan"][plurk_fan_number_new]["followday"] = sub_program.get_followday(number, start_time, k, plurk_data_new["fan"][plurk_fan_number_new]["status"], plurk_data_new["fan"][plurk_fan_number_new]["startfollowdate"], plurk_data_last["fan"][plurk_data_last["other"]["plurk_fan_accountlist"].index(plurk_data_new["fan"][plurk_fan_number_new]["account"])]["followday"], plurk_fan_string, plurk_data_new["fan"][plurk_fan_number_new]["endfollowdate"], worksheet) #傳入副程式number, start_time, i, status, startfollowdate, followday, string, endfollowdate這些參數，取得追蹤天數
                                plurk_fan_number_new = plurk_fan_number_new+1
                            else:
                                del plurk_data_new["fan"][plurk_fan_number_new] #如果從試算表取得的粉絲account的值有在串列裡，刪除以plurk_fan_number_new為名的鍵
                        except:
                            del plurk_data_new["fan"][plurk_fan_number_new]
                
                if code_section_4_status == "〇":
                    plurk_friend_number_new = 0 #下方迴圈的計數
                    plurk_friend_string = "plurk_friend-" #沒有起始日則在表格填上這個字串
                    for i in range(len(plurk_data_get["friend"])): #範圍為friend鍵的值(字典)內所含有的數量
                        plurk_data_new["friend"][plurk_friend_number_new] = {} #以plurk_friend_number_new為鍵建立對應的值，此值為字典
                        try:
                            if plurk_data_get["friend"][i]["account"] in plurk_data_last["other"]["plurk_friend_accountlist"]: #如果從噗浪取得的朋友account的值有在串列裡
                                plurk_data_new["friend"][plurk_friend_number_new]["status"] = "old" #標記狀態為old
                                plurk_data_new["friend"][plurk_friend_number_new]["account"] = plurk_data_get["friend"][i]["account"] #將從噗浪取得的朋友account、displayname、建立日和最後上線日記錄進plurk_data_new
                                plurk_data_new["friend"][plurk_friend_number_new]["displayname"] = plurk_data_get["friend"][i]["displayname"]
                                plurk_data_new["friend"][plurk_friend_number_new]["regdate"] = plurk_data_get["friend"][i]["regdate"]
                                plurk_data_new["friend"][plurk_friend_number_new]["lastondate"] = plurk_data_get["friend"][i]["lastondate"]
                                plurk_data_new["friend"][plurk_friend_number_new]["startfollowdate"] = plurk_data_last["friend"][plurk_data_last["other"]["plurk_friend_accountlist"].index(plurk_data_new["friend"][plurk_friend_number_new]["account"])]["startfollowdate"] #尋找指定序號後，從前一筆(前天)紀錄取得開始追蹤的日期
                                plurk_data_new["friend"][plurk_friend_number_new]["followday"] = sub_program.get_followday(number, start_time, i, plurk_data_new["friend"][plurk_friend_number_new]["status"], plurk_data_new["friend"][plurk_friend_number_new]["startfollowdate"], plurk_data_last["friend"][plurk_data_last["other"]["plurk_friend_accountlist"].index(plurk_data_new["friend"][plurk_friend_number_new]["account"])]["followday"], plurk_friend_string, 0, worksheet) #傳入副程式number, start_time, i, status, startfollowdate, followday, string, endfollowdate這些參數，取得追蹤天數
                                plurk_data_new["friend"][plurk_friend_number_new]["endfollowdate"] = "" #結束時間為空白
                                plurk_friend_number_new = plurk_friend_number_new+1
                            else:
                                del plurk_data_new["friend"][plurk_friend_number_new] #如果從噗浪取得的朋友account的值沒有在串列裡，刪除以plurk_friend_number_new為名的鍵
                        except:
                            del plurk_data_new["friend"][plurk_friend_number_new] #如果從噗浪取得的朋友account的值沒有在串列裡，刪除以plurk_friend_number_new為名的鍵
                    
                    for j in range(len(plurk_data_get["friend"])): #範圍為friend鍵的值(字典)內所含有的數量
                        plurk_data_new["friend"][plurk_friend_number_new] = {} #以plurk_friend_number_new為鍵建立對應的值，此值為字典
                        try:
                            if plurk_data_get["friend"][j]["account"] not in plurk_data_last["other"]["plurk_friend_accountlist"]: #如果從噗浪取得的朋友account的值沒有在串列裡
                                plurk_data_new["friend"][plurk_friend_number_new]["status"] = "new" #標記狀態為new
                                plurk_data_new["friend"][plurk_friend_number_new]["account"] = plurk_data_get["friend"][j]["account"] #將從噗浪取得的朋友account、displayname、建立日和最後上線日記錄進plurk_data_new
                                plurk_data_new["friend"][plurk_friend_number_new]["displayname"] = plurk_data_get["friend"][j]["displayname"]
                                plurk_data_new["friend"][plurk_friend_number_new]["regdate"] = plurk_data_get["friend"][j]["regdate"]
                                plurk_data_new["friend"][plurk_friend_number_new]["lastondate"] = plurk_data_get["friend"][j]["lastondate"]
                                plurk_data_new["friend"][plurk_friend_number_new]["startfollowdate"] = str(start_time.year)+"-"+str(start_time.month)+"-"+str(int(start_time.day)-1)+" 0:00:00" #以昨天為開始追蹤日
                                plurk_data_new["friend"][plurk_friend_number_new]["followday"] = "0日0時0分0秒" #第0天
                                plurk_data_new["friend"][plurk_friend_number_new]["endfollowdate"] =  "" #結束時間為空白
                                plurk_friend_number_new = plurk_friend_number_new+1
                            else:
                                del plurk_data_new["friend"][plurk_friend_number_new] #如果從噗浪取得的粉絲account的值沒有在串列裡，刪除以plurk_friend_number_new為名的鍵
                        except:
                            del plurk_data_new["friend"][plurk_friend_number_new] #如果從噗浪取得的朋友account的值沒有在串列裡，刪除以plurk_friend_number_new為名的鍵
                    
                    for k in range(len(plurk_data_get["friend"])): #範圍為friend鍵的值(字典)內所含有的數量
                        plurk_data_new["friend"][plurk_friend_number_new] = {} #以plurk_friend_number_new為鍵建立對應的值，此值為字典
                        try:
                            if plurk_data_last["friend"][k]["account"] not in plurk_data_get["other"]["plurk_friend_accountlist"]: #如果從試算表取得的朋友account的值沒有在串列裡
                                plurk_data_new["friend"][plurk_friend_number_new]["status"] = "leave" #標記狀態為leave
                                plurk_data_new["friend"][plurk_friend_number_new]["account"] = plurk_data_get["friend"][k]["account"] #將從噗浪取得的朋友account、displayname、建立日和最後上線日記錄進plurk_data_new
                                plurk_data_new["friend"][plurk_friend_number_new]["displayname"] = plurk_data_get["friend"][k]["displayname"]
                                plurk_data_new["friend"][plurk_friend_number_new]["regdate"] = plurk_data_get["friend"][k]["regdate"]
                                plurk_data_new["friend"][plurk_friend_number_new]["lastondate"] = plurk_data_get["friend"][k]["lastondate"]
                                plurk_data_new["friend"][plurk_friend_number_new]["startfollowdate"] = plurk_data_last["friend"][plurk_data_last["other"]["plurk_friend_accountlist"].index(plurk_data_new["friend"][plurk_friend_number_new]["account"])]["startfollowdate"] #尋找指定序號後，從前一筆(前天)紀錄取得開始追蹤的日期
                                
                                if plurk_data_last["friend"][plurk_data_last["other"]["plurk_friend_accountlist"].index(plurk_data_new["friend"][plurk_friend_number_new]["account"])]["endfollowdate"] == "": #如果日期等於空白
                                    plurk_data_new["friend"][plurk_friend_number_new]["endfollowdate"] = str(start_time.year)+"-"+str(start_time.month)+"-"+str(int(start_time.day)-1) #登記昨天為結束追蹤的日期
                                else: #如果日期不等於空白
                                    plurk_data_new["friend"][plurk_friend_number_new]["endfollowdate"] = plurk_data_last["friend"][plurk_data_last["other"]["plurk_friend_accountlist"].index(plurk_data_new["friend"][plurk_friend_number_new]["account"])]["endfollowdate"] #尋找指定序號後，從前一筆(前天)紀錄取得結束追蹤的日期
                                plurk_data_new["friend"][plurk_friend_number_new]["followday"] = sub_program.get_followday(number, start_time, k, plurk_data_new["friend"][plurk_friend_number_new]["status"], plurk_data_new["friend"][plurk_friend_number_new]["startfollowdate"], plurk_data_last["friend"][plurk_data_last["other"]["plurk_friend_accountlist"].index(plurk_data_new["friend"][plurk_friend_number_new]["account"])]["followday"], plurk_friend_string, plurk_data_new["friend"][plurk_friend_number_new]["endfollowdate"], worksheet) #傳入副程式number, start_time, i, status, startfollowdate, followday, string, endfollowdate這些參數，取得追蹤天數
                                plurk_friend_number_new = plurk_friend_number_new+1
                            else:
                                del plurk_data_new["friend"][plurk_friend_number_new] #如果從噗浪取得的朋友account的值沒有在串列裡，刪除以plurk_friend_number_new為名的鍵
                        except:
                            del plurk_data_new["friend"][plurk_friend_number_new] #如果從噗浪取得的朋友account的值沒有在串列裡，刪除以plurk_friend_number_new為名的鍵
                code_section_6B_status = "〇"
            except:
                code_section_6B_status = "✕"
                time_6B = time_6B+1
                
        code_section_6C_status, time_6C = "✕", 0
        while code_section_6C_status == "✕" and time_6C < 3:
            try:
                if code_section_3_status == "〇":
                    for l in range(3, (plurk_fan_number_last+4)):
                        worksheet.update_values("B{}".format(l), [["", "", "", "", "", "", ""]]) #寫入「完全空白」來清除版面
                    
                    for plurk_fan_number in range(len(plurk_data_new["fan"])): #範圍為fan鍵的值(字典)內所含有的數量
                        worksheet.update_value("B{}".format(plurk_fan_number+3), plurk_data_new["fan"][plurk_fan_number]["account"]) #寫入追蹤者帳號
                        worksheet.update_value("C{}".format(plurk_fan_number+3), plurk_data_new["fan"][plurk_fan_number]["displayname"]) #寫入追蹤者暱稱
                        worksheet.update_value("D{}".format(plurk_fan_number+3), plurk_data_new["fan"][plurk_fan_number]["regdate"]) #寫入註冊日期
                        worksheet.update_value("E{}".format(plurk_fan_number+3), plurk_data_new["fan"][plurk_fan_number]["lastondate"]) #寫入上次上站日期
                        worksheet.update_value("F{}".format(plurk_fan_number+3), plurk_data_new["fan"][plurk_fan_number]["startfollowdate"]) #寫入開始追蹤日期
                        worksheet.update_value("G{}".format(plurk_fan_number+3), plurk_data_new["fan"][plurk_fan_number]["followday"]) #寫入追蹤天數
                        worksheet.update_value("H{}".format(plurk_fan_number+3), plurk_data_new["fan"][plurk_fan_number]["endfollowdate"]) #寫入結束追蹤日期
                
                if code_section_4_status == "〇":
                    for m in range(3, (plurk_friend_number_last+4)):
                        worksheet.update_values("I{}".format(m), [["", "", "", "", "", "", ""]]) #寫入「完全空白」來清除版面
                    
                    for plurk_friend_number in range(len(plurk_data_new["friend"])): #範圍為friend鍵的值(字典)內所含有的數量
                        worksheet.update_value("I{}".format(plurk_friend_number+3), plurk_data_new["friend"][plurk_friend_number]["account"]) #寫入朋友帳號
                        worksheet.update_value("J{}".format(plurk_friend_number+3), plurk_data_new["friend"][plurk_friend_number]["displayname"]) #寫入朋友帳號
                        worksheet.update_value("K{}".format(plurk_friend_number+3), plurk_data_new["friend"][plurk_friend_number]["regdate"]) #寫入註冊日期
                        worksheet.update_value("L{}".format(plurk_friend_number+3), plurk_data_new["friend"][plurk_friend_number]["lastondate"]) #寫入上次上站日期
                        worksheet.update_value("M{}".format(plurk_friend_number+3), plurk_data_new["friend"][plurk_friend_number]["startfollowdate"]) #寫入開始追蹤日期
                        worksheet.update_value("N{}".format(plurk_friend_number+3), plurk_data_new["friend"][plurk_friend_number]["followday"]) #寫入追蹤天數
                        worksheet.update_value("O{}".format(plurk_friend_number+3), plurk_data_new["friend"][plurk_friend_number]["endfollowdate"]) #寫入結束追蹤日期
                code_section_6C_status = "〇"
            except:
                code_section_6C_status = "✕"
                time_6C = time_6C+1
        if code_section_6A_status == "〇" and code_section_6B_status == "〇" and code_section_6C_status=="〇":
            code_section_6_status = "〇"
        else:
            code_section_6_status = "✕"
        
        code_section_7_status, time_7 = "✕", 0
        while code_section_7_status == "✕" and time_7 < 3:
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
                worksheet.update_value("D{}".format(number+3), "get_plurk_data(Python)") #寫入現在執行的檔案
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
                    worksheet.update_value("X{}".format(number+3), sub_program.get_nic_data()[3])
                    if ip != None: #如果沒有抓到IP
                        worksheet.update_value("Y{}".format(number+3), ip) #寫入網際網路(外網)IP
                        worksheet.update_value("Z{}".format(number+3), sub_program.get_ip_and_version()[1]) #寫入外網IP的種類(IPv4/IPv6)
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
                time_7 = time_7+1
        worksheet= open_googlesheets.worksheet('id', 1888051586)
        worksheet.update_value("O{}".format(number+3), code_section_7_status) #寫入程式運作狀態 #系統訊息寫入
        worksheet.update_value("H{}".format(number+3), str(sub_program.TWtime()[0]).split(" ")[1].split("+")[0]) #寫入程式執行的時間
        worksheet.update_value("G{}".format(number+3), str(sub_program.TWtime()[0]-start_time)) #寫入程式結束的時間