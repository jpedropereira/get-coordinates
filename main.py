import os
from dotenv import load_dotenv
import xlwings as xw

from geopy.geocoders import Bing

load_dotenv()

BING_API_KEY = os.getenv("BING_API_KEY")