# MetOffice_DayCode

Python script to request and retrieve MetOffice data for given specified locations. Currently requesting data for Edinburgh, Glasgow, Aberdeen and Dundee.

Requests are made with the Met Office API Weather DataHub service using the 'Global spot data bundle'.

# Usage:

Requests data for specified coordinate locations. The response will include weather observation for the current day and the day before, and weather forecasts for the following 6 days.
The current focus of the script is to extract the weather significant daily code corresponding to a daily observation/forecast (e.g. "Sunny day", "Light rain", etc.).
These get recorded in a pre-set excel workbook for easy distribution against the corresponding date. It's worth noting here that, existing data (weather codes) for a particular day will be overwritten if the current API response covers that date as information in the latest response is considered to be more accurate.
The script outputs the raw JSON response from the API for each city in case a future need is identified to analyse data available from the response but not currently covered by the script.


# Depends:

API key for MetOffice

https://metoffice.apiconnect.ibmcloud.com/metoffice/production/

# Libraries used:

http.client - To attempt making connections and making HTTP requests

json - JSON encoder and decoder

datetime - For manipulating dates and times

openpyxl - For interacting with excel files
