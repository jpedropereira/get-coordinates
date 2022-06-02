import os
from dotenv import load_dotenv
import xlwings as xw

from geopy.geocoders import Bing

load_dotenv()

BING_API_KEY = os.getenv("BING_API_KEY")
WB_PATH = "addresses_template.xlsx" #path to the file containing the addresses

locator = Bing(api_key=BING_API_KEY)

#builds connection to workbook containing the addresses
wb = xw.Book(WB_PATH)
#builds connection to worksheer containing the addresses
ws = wb.sheets[0]

