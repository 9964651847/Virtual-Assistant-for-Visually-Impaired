import pyttsx3
import json
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import smtplib
import requests
from bs4 import BeautifulSoup
import win32com.client as wincl

from urllib.request import urlopen


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[1].id)
engine.setProperty('voice', voices[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon!")   

    else:
        speak("Good Evening!")  

    speak("I am Jarvis. Please tell me how may I help you")       

def takeCommand():
    #It takes microphone input from the user and returns string output
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        # print(e)    
        print("Say that again please...")  
        return "None"
    return query

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('youremail@gmail.com', 'your-password')
    server.sendmail('youremail@gmail.com', to, content)
    server.close()

if __name__ == "__main__":
    wishMe()
    while True:
    # if 1:
        query = takeCommand().lower()

        # Logic for executing tasks based on query
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            webbrowser.open("google.com")   

        elif 'news' in query:
            try:
                jsonObj = urlopen('''https://newsapi.org//v1//articles?source=the-times-of-india&sortBy=top&apiKey=\\times of India Api key\\''')
                data = json.load(jsonObj)
                i = 1
                 
                speak('here are some top news from the times of india')
                print('''=============== TIMES OF INDIA ============'''+ '\n')
                 
                for item in data['articles']:
                     
                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e:
                 
                print(str(e))
                
        elif "weather" in query:
            # Google Open weather website
            # to get API of Open weather
            api_key = "Api key"
            base_url = "http://api.openweathermap.org/data/2.5/weather?"
            params = {"q": "Bengaluru", "appid": "your_api_key"}
            response = requests.get(base_url, params=params)
            speak(" City name ")
            print("City name : ")
            city_name = takeCommand()
            complete_url = base_url + "appid =" + api_key + "&q =" + city_name
            response = requests.get(complete_url)

            x = response.json()

            if response.status_code == 200:
                data = response.json
                y = x["main"]
                temperature = data["main"]["temp"]        # Access the desired information from the JSON data
                description = data['weather'][0]['description']
                print('The temperature in Bengaluru is', temperature, 'degrees Celsius with', description)
                z = x["weather"]
                weather_description = z[0]["description"]
                #print(" Temperature (in kelvin unit) = " +str(temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description)) 
            else:
                speak(" City Not Found ")

            """if response.status_code == 200:
            # Extract the JSON response data
                 data = response.json()
                 temperature = data["main"]["temp"]        # Access the desired information from the JSON data
                 description = data['weather'][0]['description']
                 print('The temperature in Bengaluru is', temperature, 'degrees Celsius with', description)
            else:
                # Handle the error
                print('Error: HTTP status code', response.status_code)"""
            
        elif 'time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak(f"The time is {strTime}")
        
        elif 'date' in query:
            strDate = datetime.datetime.now().strftime('%d /%m /%Y')
            speak(f"The Date is {strDate}")
            
        elif 'quit' in query or 'bye' in query:
            speak("Quitting , Thanks For Your Time")
            exit()

        elif 'open code' in query:
            codePath = "C:\\Users\\manas\\Downloads\\Creating_Virtual_Assistant\\code"
            os.startfile(codePath)
            
        
        elif 'email to harry' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "harryyourEmail@gmail.com"    
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry my friend harry. I am not able to send this email")

"""
if response.status_code == 200:
    # Process the JSON response
    data = response.json()
    temperature = data['main']['temp']
    description = data['weather'][0]['description']
    print('The temperature in Bengaluru is', temperature, 'degrees Celsius with', description)
else:
    # Handle the error
    print('Error: HTTP status code', response.status_code)

import requests

# Define the API endpoint and parameters
url = "http://api.openweathermap.org/data/2.5/weather"
params = {"q": "Bengaluru", "appid": "your_api_key"}

# Send a GET request to the API and store the response
response = requests.get(url, params=params)

# Check if the request was successful
if response.status_code == 200:
    # Extract the JSON response data
    data = response.json()

    # Access the desired information from the JSON data
    temperature = data["main"]["temp"]


    try:
                jsonObj = urlopen('''https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\times of India Api key\\''')
                #jsonObj = request.get(urlopen("https://www.googleadservices.com/pagead/aclk?sa=L&ai=DChcSEwj5lvy-3L_9AhX8kmYCHSuECOIYABABGgJzbQ&ase=2&ohost=www.google.com&cid=CAESbeD2W2ifNJn_AIedkI9BiHdRNEIekx-iZIV9taNJPJllB8SYclOqv9UBdRq_gIVitpdJmSG_qi9b59Grcr_tzi3OTbceJ6YQixtJkthf1bWhmw0QIFmfNu3tBrF0xPc0-tEPAvm9mAZl3JmN7Is&sig=AOD64_3PqGICchvHDe-_hBQBzwihtgxUPQ&q&nis=4&adurl&ved=2ahUKEwjl0_W-3L_9AhXN4DgGHSjUAIIQ0Qx6BAgKEAE"))
                data = json.load(jsonObj)
                i = 1
                speak('here are some top news from the times of india')
                print('''=============== TIMES OF INDIA ============'''+ '\n')
                for item in data['articles']: 
                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e: 
                print(str(e))




if response.status_code == 200:
            # Extract the JSON response data
                 data = response.json()
                 temperature = data["main"]["temp"]        # Access the desired information from the JSON data
                 description = data['weather'][0]['description']
                 print('The temperature in Bengaluru is', temperature, 'degrees Celsius with', description)
            else:
                # Handle the error
                print('Error: HTTP status code', response.status_code)
"""


#jsonObj = request.get(urlopen("https://www.googleadservices.com/pagead/aclk?sa=L&ai=DChcSEwj5lvy-3L_9AhX8kmYCHSuECOIYABABGgJzbQ&ase=2&ohost=www.google.com&cid=CAESbeD2W2ifNJn_AIedkI9BiHdRNEIekx-iZIV9taNJPJllB8SYclOqv9UBdRq_gIVitpdJmSG_qi9b59Grcr_tzi3OTbceJ6YQixtJkthf1bWhmw0QIFmfNu3tBrF0xPc0-tEPAvm9mAZl3JmN7Is&sig=AOD64_3PqGICchvHDe-_hBQBzwihtgxUPQ&q&nis=4&adurl&ved=2ahUKEwjl0_W-3L_9AhXN4DgGHSjUAIIQ0Qx6BAgKEAE"))

"""
elif 'play music' in query:
            music_dir = 'D:\\Non Critical\\songs\\Favorite Songs2'
            songs = os.listdir(music_dir)
            print(songs)    
            os.startfile(os.path.join(music_dir, songs[0]))

import requests

# Define the API endpoint and parameters
url = "http://api.openweathermap.org/data/2.5/weather"
params = {"q": "Bengaluru", "appid": "your_api_key"}

# Send a GET request to the API and store the response
response = requests.get(url, params=params)

# Check if the request was successful
if response.status_code == 200:
    # Extract the JSON response data
    data = response.json()

    # Access the desired information from the JSON data
    temperature = data["main"]["temp"]
"""

#base_url = api_key.request('GET', 'http://api.openweathermap.org / data / 2.5 / weather?')
            #base_url = "https://openweathermap.org/api"
            #r = http.request('GET', 'http://google.com/')

"""
import requests

# Set up the API endpoint and parameters
url = 'https://api.openweathermap.org/data/2.5/weather'
params = {'q': 'Bengaluru', 'appid': 'your_api_key'}

# Send an HTTP GET request to the API endpoint
response = requests.get(url, params=params)

# Check the status code of the response
if response.status_code == 200:
    # Process the JSON response
    data = response.json()
    temperature = data['main']['temp']
    description = data['weather'][0]['description']
    print('The tem
"""