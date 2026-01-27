# import json
# import requests
# import datetime
# import pandas as pd
#
# from constants import SYMBOL_DATA_API, HEADERS, DATE
# from utils import get_duration_params
#
# # ================= CONFIG =================
# SHARE = "PAYTM"
# TOKEN = "1716481"
#
#
#
# # ================= FETCH DATA =================
# start_date = f"{(datetime.datetime.strptime(DATE, '%d.%m.%y') - datetime.timedelta(days=5)).strftime('%d.%m.%y')}"
#
# URL = f"{SYMBOL_DATA_API}/12602626/day"
# PARAMS = get_duration_params(f"{start_date} 09:15:00", f"{DATE} 15:30:00")
#
# PARAMS |= {"continuous":"1"}
#
# resp = requests.get(URL, headers=HEADERS, params=PARAMS)
# candles = resp.json()["data"]["candles"]
#
# print(candles)
import shutil

from constants import COPY_TO_CASH, BASE_FOLDER_PATH

for symbol in COPY_TO_CASH:
    file_daily = rf"{BASE_FOLDER_PATH}\DAILY\{symbol}.xlsx"
    file_daily_cash = rf"{BASE_FOLDER_PATH}\CASH\{symbol}.xlsx"
    file_daily_raghav = rf"C:\Users\RITIK\Desktop\STUDY MATERIAL\CASH\DAILY\{symbol}.xlsx"
    file_daily_cash_raghav = rf"C:\Users\RITIK\Desktop\STUDY MATERIAL\CASH\{symbol}.xlsx"

    shutil.copy(file_daily, file_daily_cash)
    shutil.copy(file_daily_raghav, file_daily_cash_raghav)