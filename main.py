import pyttsx3
import speech_recognition as sr
import datetime
import os
import json
from requests import get
import wmi
import psutil
import platform
import wikipedia
import webbrowser
import pywhatkit
import sys
import smtplib
import pyautogui
import time
from time import sleep
import pyjokes
import requests
import PyPDF2

# from PyQt5 import QtWidgets, QtCore, QtGui
# from PyQt5.QtCore import QTimer, QTime, QDate, Qt
# from PyQt5.QtGui import QMovie
# from PyQt5.QtCore import *
# from PyQt5.QtGui import *
# from PyQt5.QtWidgets import *

engine = pyttsx3.init('sapi5')
voice = engine.getProperty('voices')
engine.setProperty('voices', voice[1].id)


# Jarvis speaking code
def speak(audio):
    engine.say(audio)
    print(audio)
    engine.runAndWait()


# Function to send email
def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('youremail@gmail.com', 'your-password')
    server.sendmail('youremail@gmail.com', to, content)
    server.close()


# Fetching news function
def news():
    news_url = "https://newsapi.org/v2/top-headlines?sources=techcrunch&apiKey=3d7d2570874a4de697f837fc650bb57b"
    news_page = requests.get(news_url).json()
    articles = news_page["articles"]
    headlines = []
    days = ["first", "second", "third", "fourth", "fifth", "sixth", "seventh", "eighth", "ninth", "tenth"]
    for a in articles:
        headlines.append(a["title"])
    for i in range(len(days)):
        speak(f"Today's {days[i]}  news is: {headlines[i]}")


# Jarvis greeting function
def greeting():
    hour = int(datetime.datetime.now().hour)
    timee = time.strftime("%I:%M %p")
    if hour > 0 and hour <= 12:
        speak(f"Good morning sir its {timee}")
    elif hour > 12 and hour < 18:
        speak(f"Good afternoon sir its {timee}")
    else:
        speak(f"Good evening sir its {timee}")
    speak("Your assistant Jarvis at your service sir...How can I help you")


# Pdf reading function
def pdf_reader():
    speak("Sir can u please enter the correct file location of the pdf u need me to read out please")
    pdflocation = input()
    book = open(pdflocation, 'rb')
    pdfreader = PyPDF2.PdfFileReader(book)
    pagecount = pdfreader.numPages
    speak(f"Sir there are totally {pagecount} pages in the book u told me to read sir")
    speak("Sir please enter the page number of the book u need me to read ")
    pagenumber = int(input("Enter the page number sir: "))
    page = pdfreader.getPage(pagenumber)
    contents = page.extractText()
    speak(contents)


# Takes input from user
def receivecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening.....")
        r.pause_threshold = 1
        audio = r.listen(source, timeout=2, phrase_time_limit=6)
    try:
        print("Recognizing....")
        query = r.recognize_google(audio, language="en-in")
        print(f"Boss said:{query}")

    except Exception as e:
        speak("Can you come again sir please")
        return "none"
    return query


if __name__ == '__main__':
    greeting()

    while True:
        query = receivecommand().lower()
        if "open notepad" in query:
            npath = "C:\\Windows\\system32\\notepad.exe"
            os.startfile(npath)

        elif "close notepad" in query:
            speak("Ok sir, Closing notepad")
            os.system("taskkill /f /im notepad.exe")

        elif "open d drive" in query:
            dpath = "D:"
            os.startfile(dpath)
        elif "open Code Playground" in query:
            dpath = "E:"
            os.startfile(dpath)

        elif "open chrome" in query:
            cpath = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
            os.startfile(cpath)

        elif "close chrome" in query:
            speak("Ok sir, Closing chrome")
            os.system("taskkill /f /im chrome.exe")

        elif "open command prompt" in query:
            os.system("start cmd")

        elif "close command prompt" in query:
            speak("Ok sir, Closing cmd")
            os.system("taskkill /f /im cmd.exe")

        elif "play movie" in query:
            movie_dir = "D:\\movies"
            movies = os.listdir(movie_dir)
            os.startfile(os.path.join(movie_dir, movies[1]))

        elif "study material" in query:
            academicspath = "D:\\sem 4"
            os.startfile(academicspath)

        elif "data science" in query:
            aiprojectspath = "E:\\Data Science"
            os.startfile(aiprojectspath)

        elif "downloads" in query:
            bipath = "C:\\Users\\Anonymous\\Downloads"
            os.startfile(bipath)

        elif "photos" in query:
            certificatespath = "D:\\Photos"
            os.startfile(certificatespath)


        elif "your source folder" in query:
            jarvispath = "E:\\Jarvis"
            os.startfile(jarvispath)


        elif "my poetries" in query:
            myimagepath = "C:\\Users\\Anonymous\\OneDrive\\Pictures\\yq"
            os.startfile(myimagepath)


        elif 'wikipedia' in query:

            speak('Searching Wikipedia...')

            query = query.replace("wikipedia", "")

            results = wikipedia.summary(query, sentences=2)

            speak("According to Wikipedia")

            print(results)

            speak(results)

        elif "ip address" in query:
            ip = get("https://api.ipify.org").text
            speak(f"Sir your ip address is: {ip}")

        elif "running process" in query:
            f = wmi.WMI()
            speak("Sir your running processes are:")
            for process in f.Win32_Process():
                # Displaying the P_ID and P_Name of the process
                speak(f"{process.Name:}")

        elif "task manager" in query:
            taskmanagerpath = "C:\\Windows\\system32\\Taskmgr.exe"
            os.startfile(taskmanagerpath)

        elif "ram usage" in query:
            psutil.cpu_percent()
            psutil.virtual_memory()
            dict(psutil.virtual_memory()._asdict())
            speak(f"Sir you have used {psutil.virtual_memory().percent} percentage of RAM and ")

        elif "memory available" in query:
            psutil.cpu_percent()
            psutil.virtual_memory()
            dict(psutil.virtual_memory()._asdict())
            speak(
                f"Sir your available memory is: {int(psutil.virtual_memory().available * 100 / psutil.virtual_memory().total)} percentage")

        elif "cpu usage" in query:
            speak(f"Sir you have used: {psutil.cpu_percent()} percentage of your cpu")


        elif "system information" in query:
            uname = platform.uname()
            speak(f"System: {uname.system}")
            speak(f"Node Name: {uname.node}")
            speak(f"Release: {uname.release}")
            speak(f"Version: {uname.version}")
            speak(f"Machine: {uname.machine}")
            speak(f"Processor: {uname.processor}")

        elif "cpu information" in query:
            speak(f"Physical cores:{psutil.cpu_count(logical=False)}")
            speak(f"Total cores:{psutil.cpu_count(logical=True)}")
            cpufreq = psutil.cpu_freq()
            speak(f"Max Frequency: {cpufreq.max:.2f}Mhz")
            speak(f"Min Frequency: {cpufreq.min:.2f}Mhz")
            speak(f"Current Frequency: {cpufreq.current:.2f}Mhz")
            speak("CPU Usage Per Core:")
            for i, percentage in enumerate(psutil.cpu_percent(percpu=True, interval=1)):
                speak(f"Core {i}: {percentage}%")
            speak(f"Total CPU Usage: {psutil.cpu_percent()}%")

        elif "battery information" in query:
            speak(f"Sir you have used {psutil.sensors_battery()} percentage of battery")


        elif "open youtube" in query:
            webbrowser.open("https://www.youtube.com/")

        elif "open hotstar" in query:
            webbrowser.open("https://www.hotstar.com/in")

        elif "open whatsapp" in query:
            webbrowser.open("https://web.whatsapp.com/")

        elif "open linkedin" in query:
            webbrowser.open("https://www.linkedin.com/feed/")

        elif "open my github" in query:
            webbrowser.open("https://github.com/tushar-max")

        elif "django videos" in query:
            webbrowser.open("https://www.youtube.com/playlist?list=PLu0W_9lII9ah7DDtYtflgwMwpT3xmjXY9")

        elif "chess.com" in query:
            webbrowser.open("https://www.chess.com/")
        elif "lichess" in query:
            webbrowser.open("https://lichess.org/")

        elif "gmail" in query:
            webbrowser.open("https://mail.google.com/")

        elif "open hackerank" in query:
            webbrowser.open("https://www.hackerrank.com/dashboard")

        elif "open youtube music" in query:
            webbrowser.open("https://music.youtube.com/")

        elif "open stack overflow" in query:
            webbrowser.open("https://stackoverflow.com/")

        elif "open google" in query:
            speak("Sir, what do you want me to search on google?")
            command = receivecommand().lower()
            webbrowser.open(f"{command}")


        elif "play song in youtube" in query:
            speak("Sir, what song do you want me to play on youtube")
            songcommand = receivecommand().lower()
            pywhatkit.playonyt(f"{songcommand}")

        elif "play video in Youtube" in query:
            speak("Sir, what video do you want me to play on youtube")
            videocommand = receivecommand().lower()
            pywhatkit.playonyt(f"{videocommand}")


        elif "tell me a joke" in query:
            joke = pyjokes.get_joke()
            speak(joke)

        elif "shut down the system" in query:
            os.system("shutdown /s /t 5")

        elif "restart the system" in query:
            os.system("shutdown /r /t 5")

        elif "switch the window" in query:
            pyautogui.keyDown("alt")
            pyautogui.press("tab")
            sleep(1)
            pyautogui.keyUp("alt")

        elif "tell me news" in query:
            speak("Just give me a moment fetching some top headlines for you......")
            news()



        elif "location" in query:
            speak("Wait a moment sir, let me check")
            send_url = "http://api.ipstack.com/check?access_key=09dcd7a4a1569efaef378a9d076224df"
            geo_req = requests.get(send_url)
            geo_json = json.loads(geo_req.text)
            latitude = geo_json['latitude']
            longitude = geo_json['longitude']
            city = geo_json['city']
            region = geo_json['region_name']
            country = geo_json['country_name']
            speak(f"I guess if I am not wrong we are near {city} area in {region} of {country} sir")

        elif "pdf" in query:
            pdf_reader()

        elif "exit yourself" in query:
            speak(
                "Thank you for using me, Have a good day sir, whenever again you need help your assistant will be at service. Bye sir")
            sys.exit()

        speak("Do you need me to do any other help sir")
