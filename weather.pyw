import requests, json, os
import datetime

now = datetime.datetime.now()

api_key = "902d54724fa024a1493cf006abc26fa8"
base_url = "http://api.openweathermap.org/data/2.5/weather?"
city_name = "Sofia"
#city_name = input("Enter city name : ") 
complete_url = base_url + "appid=" + api_key + "&q=" + city_name + "&units=metric"
response = requests.get(complete_url) 
x = response.json() 
if x["cod"] != "404":
    y = x["main"] 
    current_temperature = y["temp"] 
    current_pressure = y["pressure"] 
    current_humidiy = y["humidity"] 
    z = x["weather"] 
    weather_description = z[0]["description"] 
    print(" Temperature (in celsius unit) = " +
                    str(current_temperature) + 
          "\n atmospheric pressure (in hPa unit) = " +
                    str(current_pressure) +
          "\n humidity (in percentage) = " +
                    str(current_humidiy) +
          "\n description = " +
                    str(weather_description)) 
else: 
    print(" City Not Found ") 

file = open("weather.txt", "w")
file.write(str(city_name) + ": " + str(current_temperature)+"° "+"at: "+str(now.hour)+":"+str(now.minute)+"h")
file.close()