import logging, os, io, requests, pandas as pd
from kiteconnect import KiteConnect
from constants import EQ_SYMBOLS, EQ_INSTRUMENTS_URL, HEADERS, FO_INSTRUMENTS_URL
from datetime import datetime
from dotenv import load_dotenv
from utils import get_fut_instrument_token

logging.basicConfig(level=logging.DEBUG)

load_dotenv()

api_key = os.getenv("API_KEY")
access_token = os.getenv("ACCESS_TOKEN")

kite = KiteConnect(api_key=api_key)
kite.set_access_token(access_token)

# # EQ SYMBOLS
# response = requests.get(EQ_INSTRUMENTS_URL, headers=HEADERS)
#
# ds = io.StringIO(response.text)
# df = pd.read_csv(ds)
# df.to_excel(f'EQ_instruments{datetime.now().strftime("%d-%m-%Y")}.xlsx', index=False)
#
# # 1. Filter the dataframe for symbols present in your list
# # 2. Create a mapping of {tradingsymbol: instrument_token}
# found_mapping = dict(zip(df['tradingsymbol'], df['instrument_token']))
#
# # 3. Build final dictionary and track missing symbols
# final_dict = {}
# missing_symbols = []
#
# for symbol in EQ_SYMBOLS:
#     if symbol in found_mapping:
#         final_dict[symbol] = found_mapping[symbol]
#     else:
#         final_dict[symbol] = -1
#         missing_symbols.append(symbol)
#
# # Output results
# print("Final Dictionary:", final_dict)
# print("Missing Symbols:", missing_symbols)

# FO SYMBOLS
response = requests.get(FO_INSTRUMENTS_URL, headers=HEADERS)
ds = io.StringIO(response.text)
df = pd.read_csv(ds)
df.to_excel(f'FO_instruments{datetime.now().strftime("%d-%m-%Y")}.xlsx', index=False)
print(get_fut_instrument_token(df, is_expiry_today=False))  # make it true only on day of expiry
