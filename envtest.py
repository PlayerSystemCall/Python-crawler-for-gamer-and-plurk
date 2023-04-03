# -*- coding: utf-8 -*-
import os
from dotenv import load_dotenv

A = load_dotenv()
print("A", A)

GCP_PROJECT_ID = os.getenv('GCP_PROJECT_ID')
print("GCP_PROJECT_ID", GCP_PROJECT_ID)
SERVICE_ACCOUNT_FILE = os.getenv('SERVICE_ACCOUNT_FILE')
print("SERVICE_ACCOUNT_FILE", SERVICE_ACCOUNT_FILE)
STORAGE_BUCKET_NAME = os.getenv('STORAGE_BUCKET_NAME')
print("STORAGE_BUCKET_NAME", STORAGE_BUCKET_NAME)
