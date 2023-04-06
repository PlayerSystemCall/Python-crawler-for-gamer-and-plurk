# -*- coding: utf-8 -*-
import os
import json
import pygsheets
from dotenv import load_dotenv

"""
GCP_PROJECT_ID = os.getenv('GCP_PROJECT_ID')
print("GCP_PROJECT_ID=", GCP_PROJECT_ID)
SERVICE_ACCOUNT_FILE = os.getenv('SERVICE_ACCOUNT_FILE')
print("SERVICE_ACCOUNT_FILE=", SERVICE_ACCOUNT_FILE)
STORAGE_BUCKET_NAME = os.getenv('STORAGE_BUCKET_NAME')
print("STORAGE_BUCKET_NAME=", STORAGE_BUCKET_NAME)
"""

try:
    certificate = pygsheets.authorize(service_file='google_sheets_API_key.json') #取得位置在同層級目錄的Google sheets API憑證
    googlesheets_url = "https://docs.google.com/spreadsheets/d/1vLopfsKHRNaS02bI5AmKHsBbbqtL4EbY4k47SRivMSY"
    open_googlesheets = certificate.open_by_url(googlesheets_url) #開啟Google sheets
    print("Yes Link")
except:
    print("No Link")
    