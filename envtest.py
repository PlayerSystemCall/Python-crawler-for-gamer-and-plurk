# -*- coding: utf-8 -*-
import os
from dotenv import load_dotenv

GCP_PROJECT_ID = os.getenv('GCP_PROJECT_ID')
print("GCP_PROJECT_ID=", GCP_PROJECT_ID)