import json
import requests
import datetime
import pandas as pd

from constants import SYMBOL_DATA_API, HEADERS, DATE, EQ_SYMBOLS, FO_SYMBOLS, FO_SYMBOLS_WITH_EXPIRY, LTP_DATA_API, \
    LTP_PREV_PATH, BASE_FOLDER_PATH, APPEND
from utils import get_duration_params, sanitize_url

# ================= CONFIG =================
SHARE = "PAYTM"
TOKEN = "1716481"



# ================= FETCH DATA =================
start_date = f"{(datetime.datetime.strptime(DATE, '%d.%m.%y') - datetime.timedelta(days=5)).strftime('%d.%m.%y')}"

URL = f"{SYMBOL_DATA_API}/12602626/day"
URL = "https://api.kite.trade/quote/ltp?i="
#
# # for symbol in EQ_SYMBOLS:
# #     URL += f"NSE:{symbol}&i="
#
# for symbol in FO_SYMBOLS:
#     URL += f"NFO:{FO_SYMBOLS_WITH_EXPIRY[symbol]}&i="
#
# URL = URL[:-3]
# print(URL)
# PARAMS = get_duration_params(f"21.01.26 09:15:00", f"21.01.26 15:30:00")
# PARAMS |= {"continuous": "1"}
# resp = requests.get(URL, headers=HEADERS, params=PARAMS)
# candles = resp.json()["data"]
# print(candles)


import openpyxl as xl
wb = xl.load_workbook(LTP_PREV_PATH)
s = wb["Sheet1"]
i = 2

for sh in EQ_SYMBOLS:
    file_daily = rf"{BASE_FOLDER_PATH}\DAILY\{sh}.xlsx"

    wb = xl.load_workbook(file_daily)
    sheet = wb['D']
    ltp = s.cell(i, 2).value
    input_row = EQ_SYMBOLS[sh][1] + APPEND

    sheet.cell(input_row, 4).value = ltp  # close
    sheet.cell(input_row, 5).value = ltp  # LTP

    wb.save(file_daily)
    print(f"{sh} done")
    i += 1