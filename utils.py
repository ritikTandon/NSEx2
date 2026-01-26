from datetime import datetime

from constants import DATE
import pandas as pd

def get_duration_params(from_datetime, to_datetime):
    """
    Convert input datetime strings to required query params format
    Example: 2026-01-23 09:15:00
    """

    from_date, from_time = from_datetime.split(" ")
    to_date, to_time = to_datetime.split(" ")

    from_datetime = f"{datetime.strptime(from_date, '%d.%m.%y').strftime('%Y-%m-%d')} {from_time}"
    to_datetime = f"{datetime.strptime(to_date, '%d.%m.%y').strftime('%Y-%m-%d')} {to_time}"

    return {
        "from": from_datetime,
        "to": to_datetime
    }

def get_fut_instrument_token(df: pd.DataFrame, is_expiry_today=False):
    """
    Method to get instrument token for nifty and bn based with
    additional check if expiry is today then get next expiry
    """
    # Ensure expiry is datetime
    df['expiry'] = pd.to_datetime(df['expiry'])

    result = {}

    for symbol in ["NIFTY", "BANKNIFTY"]:
        subset = df[
            (df["name"] == symbol) &
            (df["instrument_type"] == "FUT")
        ].sort_values("expiry")

        if subset.empty:
            continue

        # index 0 = front month, index 1 = next month
        row_index = 1 if is_expiry_today else 0

        # safety check
        if len(subset) > row_index:
            row = subset.iloc[row_index]
            result[row["tradingsymbol"]] = int(row["instrument_token"])

    return result