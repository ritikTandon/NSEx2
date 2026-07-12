import os
from dotenv import load_dotenv
from openpyxl.styles import Font, Alignment

load_dotenv()

DATE = r"13.07.26"
MONTH = r"JUL"
YEAR = r"2026"

APPEND = 762  # 762 is 13-JUL-26

# share name to expiry symbol mapping to get LTP price
FO_SYMBOLS_WITH_EXPIRY = {'NIFTY': 'NIFTY26JULFUT', 'BANKNIFTY': 'BANKNIFTY26JULFUT'} # check fut march sysmbols

NO_FORMAT_LIST = ['APOLLOTYRE', 'BANDHANBANK', 'BANKBARODA', 'COAL INDIA', 'DLF CHL', 'TATAMOTOR CHL',
                  'TATASTEEL', 'TATAPOWER', 'M&MFINANCE', 'FEDRAL BANK', 'HINDALCO', 'NTPC']

# Shares to copy to CASH folder
COPY_TO_CASH = ['ABB', 'APOLLOHOSP', 'ASHOKLEY', 'BAJFINANCE', 'BANKBARODA', 'BHEL', 'BSE', 'CANBK', 'COFORGE', 'DIXON', 'DLF',
                'ETERNAL', 'GLENMARK', 'HDFCAMC', 'HEROMOTOCO', 'HINDALCO', 'JINDALSTEL', 'JIOFIN', 'LAURUSLABS', 'LTF',
                'MCX', 'NAUKRI', 'PAYTM', 'RECLTD', 'RELIANCE', 'TMPV', 'VEDL']

# {SYMBOL: (INSTRUMENT_TOKEN, ROW_NUMBER)} dictionary for all equity shares
EQ_SYMBOLS = {'AARTIIND': (1793, 947), 'ABB': (3329, 947), 'ABCAPITAL': (5533185, 947), 'ABFRL': (7707649, 947),
              'ADANIENT': (6401, 1579), 'ADANIPORTS': (3861249, 947), 'ALIVUS': (1347841, 947), 'ALKEM': (2995969, 947),
              'AMBUJACEM': (325121, 947), 'APOLLOHOSP': (40193, 947), 'APOLLOTYRE': (41729, 947),
              'ASHOKLEY': (54273, 947), 'ASTRAL': (3691009, 947), 'ATUL': (67329, 947), 'AUBANK': (5436929, 947),
              'AUROPHARMA': (70401, 947), 'BAJAJ-AUTO': (4267265, 947), 'BAJAJFINSV': (4268801, 1579),
              'BAJFINANCE': (81153, 1579), 'BALKRISIND': (85761, 947), 'BALRAMCHIN': (87297, 947),
              'BANDHANBNK': (579329, 1579), 'BANKBARODA': (1195009, 1579), 'BATAINDIA': (94977, 947),
              'BEL': (98049, 947),
              'BHARATFORG': (108033, 947), 'BHEL': (112129, 947), 'BIOCON': (2911489, 947), 'BRITANNIA': (140033, 947),
              'BSE': (5013761, 683), 'BSOFT': (1790465, 947), 'CANBK': (2763265, 947), 'CANFINHOME': (149249, 947),
              'CHAMBLFERT': (163073, 947), 'CHOLAFIN': (175361, 947), 'CIPLA': (177665, 947),
              'COALINDIA': (5215745, 3232),
              'COFORGE': (2955009, 947), 'CONCOR': (1215745, 947), 'COROMANDEL': (189185, 947),
              'CROMPTON': (4376065, 947), 'CUMMINSIND': (486657, 947), 'DABUR': (197633, 947),
              'DALBHARAT': (2067201, 947), 'DEEPAKFERT': (211713, 947), 'DEEPAKNTR': (5105409, 947),
              'DELTACORP': (3851265, 947), 'DIVISLAB': (2800641, 947), 'DIXON': (5552641, 947), 'DLF': (3771393, 4058),
              'DRREDDY': (225537, 947), 'EICHERMOT': (232961, 2715), 'ESCORTS': (245249, 947),
              'ETERNAL': (1304833, 447),
              'EXIDEIND': (173057, 947), 'FEDERALBNK': (261889, 1579), 'GLENMARK': (1895937, 947),
              'GNFC': (300545, 947),
              'GODREJCP': (2585345, 947), 'GODREJPROP': (4576001, 947), 'GRANULES': (3039233, 947),
              'GRASIM': (315393, 947), 'GUJGASLTD': (2713345, 947), 'HAL': (589569, 947), 'HAVELLS': (2513665, 947),
              'HCLTECH': (1850625, 1579), 'HDFCAMC': (1086465, 947), 'HDFCBANK': (341249, 3936),
              'HDFCLIFE': (119553, 947),
              'HEROMOTOCO': (345089, 683), 'HINDALCO': (348929, 947), 'HINDCOPPER': (4592385, 947),
              'ICICIBANK': (1270529, 1579), 'ICICIGI': (5573121, 947), 'ICICIPRULI': (4774913, 947),
              'IEX': (56321, 947),
              'IGL': (2883073, 947), 'INDHOTEL': (387073, 947), 'INDIACEM': (387841, 947), 'INDIAMART': (2745857, 947),
              'INDIGO': (2865921, 947), 'INDUSINDBK': (1346049, 1579), 'INDUSTOWER': (7458561, 947),
              'INFY': (408065, 2765),
              'INTELLECT': (1517057, 947), 'IPCALAB': (418049, 947), 'JINDALSTEL': (1723649, 5195),
              'JIOFIN': (4644609, -12), 'JKCEMENT': (3397121, 947), 'JSWSTEEL': (3001089, 947),
              'JUBLFOOD': (4632577, 947), 'KOTAKBANK': (492033, 947), 'LALPATHLAB': (2983425, 947),
              'LAURUSLABS': (4923905, 947), 'LICHSGFIN': (511233, 947), 'LTF': (6386689, 683), 'LTM': (4561409, 947),
              'LTTS': (4752385, 947), 'LUPIN': (2672641, 947), 'M&M': (519937, 1579), 'M&MFIN': (3400961, 1579),
              'MANAPPURAM': (4879617, 947), 'MARICO': (1041153, 947), 'MCX': (7982337, 947),
              'METROPOLIS': (2452737, 947), 'MFSL': (548353, 947), 'MGL': (4488705, 947), 'MPHASIS': (1152769, 947),
              'MUTHOOTFIN': (6054401, 947), 'NAM-INDIA': (91393, 947), 'NAUKRI': (3520257, 947),
              'NAVINFLUOR': (3756033, 947), 'NMDC': (3924993, 947), 'NTPC': (2977281, 947),
              'OBEROIRLTY': (5181953, 947), 'ONGC': (633601, 947), 'PAYTM': (1716481, 447),
              'PERSISTENT': (4701441, 947),
              'PETRONET': (2905857, 947), 'PIDILITIND': (681985, 947), 'PIRAMALFIN': (194445057, 947),
              'POLYCAB': (2455041, 947), 'POWERGRID': (3834113, 947), 'RAIN': (3926273, 947), 'RAMCOCEM': (523009, 947),
              'RBLBANK': (4708097, 947), 'RECLTD': (3930881, 947), 'RELIANCE': (738561, 4793),
              'SBICARD': (4600577, 947),
              'SBILIFE': (5582849, 947), 'SBIN': (779521, 4860), 'SIEMENS': (806401, 947), 'SRF': (837889, 947),
              'STAR': (1887745, 947), 'SUNPHARMA': (857857, 947), 'SUNTV': (3431425, 1579), 'SYNGENE': (2622209, 947),
              'TATACHEM': (871681, 1579), 'TATACOMM': (952577, 947), 'TATAPOWER': (877057, 1579),
              'TATASTEEL': (895745, 4570),
              'TCS': (2953217, 947), 'TECHM': (3465729, 947), 'TITAN': (897537, 947), 'TMPV': (884737, 4434),
              'TORNTPHARM': (900609, 947), 'TORNTPOWER': (3529217, 947), 'TRENT': (502785, 947),
              'TVSMOTOR': (2170625, 947), 'UBL': (4278529, 947), 'ULTRACEMCO': (2952193, 2696),
               'UPL': (2889473, 947), 'VEDL': (784129, 947), 'VOLTAS': (951809, 947),
              'ZEEL': (975873, 947), 'ZYDUSLIFE': (2029825, 947)}

HEADERS = {
    "X-Kite-Version": "3",
    "Authorization": f"token {os.getenv("API_KEY")}:{os.getenv("ACCESS_TOKEN")}"
}

# NIFTY and BN symbols but continuous = 1 for appi calls so that we don't need to keep updating symbols
# {SYMBOL: INSTRUMENT_TOKEN} dictionary for all F&O shares
FO_SYMBOLS = {'NIFTY': (15639810, 0), 'BANKNIFTY': (15638530, 0)}

EQ_INSTRUMENTS_URL = "https://api.kite.trade/instruments/NSE"
FO_INSTRUMENTS_URL = "https://api.kite.trade/instruments/NFO"

SYMBOL_DATA_API = "https://api.kite.trade/instruments/historical"

LTP_DATA_API = "https://api.kite.trade/quote/ltp?i="

BASE_FOLDER_PATH = rf"E:\Daily Data work"

FIXED_WIDTH = 11  # COLUMN WIDTH FOR FORMATTING

red = Font("Arial", 11, color='ff0000', bold=True)
blue = Font("Arial", 11, color="0000ff", bold=True)
bold = Font("Arial", 11, bold=True)
alignment = Alignment(horizontal='center')

SHARE_LIST = EQ_SYMBOLS | FO_SYMBOLS

MAX_POINTS = 10   # last 10 LTPs

LTP_PREV_PATH = rf"C:\Users\RITIK\PycharmProjects\NseEx2\ltp_prev.xlsx"
LTP_PREV_BACKUP_PATH = rf"C:\Users\RITIK\PycharmProjects\NseEx2\ltp_prev_backup.xlsx"
