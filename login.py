import logging, os
from kiteconnect import KiteConnect
from dotenv import load_dotenv

logging.basicConfig(level=logging.DEBUG)

load_dotenv()

api_key = os.getenv("API_KEY")
api_secret = os.getenv("API_SECRET")

kite = KiteConnect(api_key=api_key)

request_token = input(f"Please follow the URL and paste the request token from URL: {kite.login_url()}: ")
data = kite.generate_session(request_token, api_secret=api_secret)
access_token = data["access_token"]

print(f"Access Token: {access_token}")



