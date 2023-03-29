# Python-crawler-for-gamer-and-plurk
 從 Gamer(巴哈姆特) 和 Plurk(噗浪) 的帳號首頁取得資料，然後將之填進 Google Sheets 裡。  
 Get the data from account homepages of Gamer(巴哈姆特) and Plurk(噗浪) and then write it into Google Sheets.   
## 專案文件 Project File   
 ├必要文件Necessary Documents   
 │├**Player_SystemCall_account.py**   
 │└**google_sheets_API_key.json**   
 ├副程式Subprogram   
 │└Python_crawler_subprogram_ver1.py   
 ├巴哈姆特Gamer (an operating social networking site)   
 │├study_gamer_data_ver1.py   
 │├study_gamer_data_ver2.py=c.n.=>**Python_crawler_for_gamer_ver2.py**   
 │└Python_crawler_mainprogram_for_gamer_ver3.py   
 ├噗浪Plurk (a working social networking site)   
 │├study_plurk_data_ver1.py   
 │├study_plurk_data_ver2.py   
 │├study_plurk_data_ver3.py=c.n.=>**Python_crawler_for_plurk_ver3.py**   
 │└Python_crawler_mainprogram_for_plurk_ver4.py   
 └探路克Timelog (a closed social networking site)   
 　└No_file   
## 版本分支圖 Version Branch Map   
 study_web_data_ver0.py   
 ├study_plurk_data_ver1.py   
 │├study_plurk_data_ver2.py   
 ││└study_plurk_data_ver3.py=c.n.=>"Python_crawler_for_plurk_ver3.py"   
 ││　└Python_crawler_main_for_plurk_ver4.py   
 │└study_gamer_data_ver1.py   
 │　└study_gamer_data_ver2.py=c.n.=>"Python_crawler_for_gamer_ver2.py"   
 │　　└Python_crawler_main_for_gamer_ver3.py   
 └──"Python_crawler_for_plurk_ver3.py" and Python_crawler_for_gamer_ver2.py   
 　　　├"Player_SystemCall_account.py"   
 　　　└Python_crawler_sub_ver1.py   
## 備註Annotation   
 1. change name==>c.n.   
 2. **file.py** is mean the run-time version of file of python crawler   
 3. If some code can do the repetitive work, create a file to put it. Import it when program needs to use the code of the file.   
 4. The version 2.0.6 of pygsheet sisn't using.   