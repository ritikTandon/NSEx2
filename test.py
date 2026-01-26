import json
import requests
import datetime
import pandas as pd

from constants import SYMBOL_DATA_API, HEADERS, DATE
from utils import get_duration_params

# ================= CONFIG =================
SHARE = "PAYTM"
TOKEN = "1716481"



# ================= FETCH DATA =================
start_date = f"{(datetime.datetime.strptime(DATE, '%d.%m.%y') - datetime.timedelta(days=5)).strftime('%d.%m.%y')}"

URL = f"{SYMBOL_DATA_API}/12602626/day"
PARAMS = get_duration_params(f"{start_date} 09:15:00", f"{DATE} 15:30:00")

PARAMS |= {"continuous":"1"}

resp = requests.get(URL, headers=HEADERS, params=PARAMS)
candles = resp.json()["data"]["candles"]

print(candles)
