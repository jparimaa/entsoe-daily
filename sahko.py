import requests
import pandas
import json
import openpyxl
import time
import datetime

def to_gmt_plus_2(hour):
    hour += 2
    if hour >= 24:
        hour -= 24
    return hour

response = requests.get("https://api.porssisahko.net/v1/latest-prices.json")
print(response)

json_data = response.json()

dates = []
times = []
prices = []

for price_data in json_data["prices"]:
    price = price_data['price']
    start_date = price_data['startDate']
    end_date = price_data['endDate']
    start_time = to_gmt_plus_2(int(start_date[11:13]))
    end_time = to_gmt_plus_2(int(end_date[11:13]))
    date = start_date[0:10] if end_time == 0 else ""

    time = str(start_time) + "-" + str(end_time)

    dates.append(date)
    times.append(time)
    prices.append(price)

data = {'pvm': dates, 'klo': times, 'hinta snt / kWh, sis. alv.': prices}

df = pandas.DataFrame(data=data)
print(df)
df.to_excel('hinnat.xlsx')