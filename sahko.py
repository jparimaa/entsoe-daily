import requests
import pandas
import json
import time
import datetime

def to_gmt_plus_2(hour):
    hour += 2
    if hour >= 24:
        hour -= 24
    return hour

def log(msg):
    date = datetime.datetime.now()
    stamped_msg = str(date) + ": " + msg
    print(stamped_msg)
    with open("sahko.log", "a") as myfile:
        myfile.write(stamped_msg+ "\n")

def sleep_until_17():
    t = datetime.datetime.today()
    future = datetime.datetime(t.year,t.month,t.day,17,0)
    if t.hour >= 17:
        future += datetime.timedelta(days=1)
    log("Sleeping until " + str(future) + " and after that query new data")
    sleep_time = (future-t).total_seconds()
    time.sleep(sleep_time)


log("This program will get the latest electricity prices from https://api.porssisahko.net/v1/latest-prices.json and write the data to an excel file called hinnat.xlsl. The file should not be open at the time of writing. Data is queried at program startup and then daily at 17:00.")

while True:
    response = requests.get("https://api.porssisahko.net/v1/latest-prices.json")
    if response.status_code != 200:
        log("ERROR: Unable to get data from server. Return code = " + str(response.status_code) + ". Trying again in 15 minutes.")
        time.sleep(15*60)
        continue

    log("Data obtained from https://api.porssisahko.net/v1/latest-prices.json")
    json_data = response.json()

    dates = []
    times = []
    prices = []

    for price_data in json_data["prices"]:
        price = price_data["price"]
        start_date = price_data["startDate"]
        end_date = price_data["endDate"]
        start_time = to_gmt_plus_2(int(start_date[11:13]))
        end_time = to_gmt_plus_2(int(end_date[11:13]))
        date_num = start_date[0:10] if end_time == 0 else ""

        time_str = str(start_time) + "-" + str(end_time)

        dates.append(date_num)
        times.append(time_str)
        prices.append(price)

    d = {"pvm": dates, "klo": times, "hinta snt / kWh, sis. alv.": prices}
    df = pandas.DataFrame(data=d)

    filename = "hinnat.xlsx"
    log("Writing latest data to " + filename)
    try:
        df.to_excel(filename)
        log("File write completed")
        sleep_until_17()
    except PermissionError:
        log("ERROR: Permission error - Unable to write file. Maybe the file is already open? Trying again in 15 minutes.")
        time.sleep(15*60)
    except FileNotFoundError:
        log("ERROR: File not found error - Check if the directory exists.")
    except ValueError as e:
        log(f"ERROR: Value error - {e}")
    except OSError as e:
        log(f"ERROR: OS error - {e}")
    except Exception as e:
        log(f"ERROR: An unexpected error occurred - {e}")


