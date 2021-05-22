import http.client
import json
import datetime 
from  datetime import date
import openpyxl

client_id = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #enter client id
client_secret = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx0" #enter client secret

all_cities = [['Edinburgh', '-3.201549', '55.950724'], ['Glasgow', '-4.258124', '55.858819'], ['Dundee', '-2.970447', '56.462230'], ['Aberdeen', '-2.099616', '57.144591']]

conn = http.client.HTTPSConnection("api-metoffice.apiconnect.ibmcloud.com")

def get_data(city, location_long, location_lat):
    headers = {
        'x-ibm-client-id': client_id,
        'x-ibm-client-secret': client_secret,
        'accept': "application/json"
        }

    conn.request("GET", "/metoffice/production/v0/forecasts/point/daily?excludeParameterMetadata=true&includeLocationName=true&latitude=" + location_lat + "&longitude=" + location_long, headers=headers)

    res = conn.getresponse()
    data = res.read()
    weather_data = json.loads(data)

    for day in range(8):
        day_code = weather_data["features"][0]["properties"]["timeSeries"][day]["daySignificantWeatherCode"]
        weather_codes.append(day_code)

    reqest_date = weather_data["features"][0]["properties"]["timeSeries"][0]["time"]
    reqest_date = datetime.datetime.strptime(reqest_date, "%Y-%m-%dT%H:%M%SZ") 

    file2write=open(city + "_" + str(date.today()),'w')
    file2write.write(json.dumps(weather_data, indent=4))
    file2write.close()

    return weather_codes, reqest_date

weatherFile = openpyxl.load_workbook(filename = "Weather_Records_v2.xlsx", read_only = False)
currentSheet = weatherFile['Data']

date_values = []
for cells in list(currentSheet.rows)[0]:
    date_values.append(cells.value)


city_counter = 0
for city in all_cities:
    weather_codes = []
    city_counter += 1
    weather_codes, reqest_date = get_data(*city)

    index = date_values.index(reqest_date)

    print(weather_codes, reqest_date, index, date_values[366])

    weather_code_index = 0
    for col in range(index + 1, index + len(weather_codes) + 1):
        _ = currentSheet.cell(column=col, row=city_counter+1, value=weather_codes[weather_code_index])
        weather_code_index += 1

weatherFile.save("Weather_Records_v2.xlsx")
