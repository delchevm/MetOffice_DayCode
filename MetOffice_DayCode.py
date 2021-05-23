import http.client
import json
import datetime 
from  datetime import date
import openpyxl

# Enter API keys
client_id = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #enter client id
client_secret = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx0" #enter client secret

# Define locations in the format: City, Longitude, Latitude
all_cities = [['Edinburgh', '-3.201549', '55.950724'], ['Glasgow', '-4.258124', '55.858819'], ['Dundee', '-2.970447', '56.462230'], ['Aberdeen', '-2.099616', '57.144591']]

# Make HTTP connection
conn = http.client.HTTPSConnection("api-metoffice.apiconnect.ibmcloud.com")


def get_data(city, location_long, location_lat):
    # Returns a list of weather codes starting from (today's date - 1); and today's date - 1
    headers = {
        'x-ibm-client-id': client_id,
        'x-ibm-client-secret': client_secret,
        'accept': "application/json"
        }
    
    # Make request
    conn.request("GET", "/metoffice/production/v0/forecasts/point/daily?excludeParameterMetadata=true&includeLocationName=true&latitude=" + location_lat + "&longitude=" + location_long, headers=headers)
    
    # Store JSON response in weather_data
    res = conn.getresponse()
    data = res.read()
    weather_data = json.loads(data)

    # Fill a list with weather codes for (today's date - 1) to (today + 7)
    for day in range(8):
        day_code = weather_data["features"][0]["properties"]["timeSeries"][day]["daySignificantWeatherCode"]
        weather_codes.append(day_code)

    reqest_date = weather_data["features"][0]["properties"]["timeSeries"][0]["time"]
    reqest_date = datetime.datetime.strptime(reqest_date, "%Y-%m-%dT%H:%M%SZ") 
    
    # Dumps JSON file with the response if in future we are interested in more than the daily weather codes
    file2write=open(city + "_" + str(date.today()),'w')
    file2write.write(json.dumps(weather_data, indent=4))
    file2write.close()

    return weather_codes, reqest_date

# Opens Weather_Records_v2.xlsx Data tab
weatherFile = openpyxl.load_workbook(filename = "Weather_Records_v2.xlsx", read_only = False)
currentSheet = weatherFile['Data']

# Stores all dates in row 1 in a list
date_values = []
for cells in list(currentSheet.rows)[0]:
    date_values.append(cells.value)


city_counter = 0
for city in all_cities:
    weather_codes = []
    city_counter += 1
    # Requests the weather_codes and reqest_date
    weather_codes, reqest_date = get_data(*city)
    
    # Finds the position of the reqest_date in the list of dates
    index = date_values.index(reqest_date)
    
    # Writes the weather codes for each date, overwriting any previous codes as the latest reponse should be the most accurate
    weather_code_index = 0
    for col in range(index + 1, index + len(weather_codes) + 1):
        _ = currentSheet.cell(column=col, row=city_counter+1, value=weather_codes[weather_code_index])
        weather_code_index += 1

weatherFile.save("Weather_Records_v2.xlsx")
