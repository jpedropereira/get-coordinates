import os
from types import NoneType
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

#Creates dynamic range address containing the list of addresses
addresses_list_range = ws.range("C3").expand("down")

#Loops through the addresses and, when successful, returns the coordinates in columns D and E

for index, value in enumerate(addresses_list_range, start=3):
    address = ws.range(f"C{index}").value
    location = locator.geocode(address)
    ws.range(f"D{index}").value = location.latitude
    ws.range(f"E{index}").value = location.longitude




