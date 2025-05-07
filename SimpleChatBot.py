import subprocess  						# Used for executing system commands like shutdown, sleep, etc.
import wolframalpha  					# Used to compute expert-level algorithms and perform complex calculations.
import pyttsx3  						# Provides functionality for text-to-speech synthesis.
import wikipedia  						# Used for searching and retrieving information from Wikipedia.
import webbrowser  						# Allows opening web pages and performing web searches.
import os  								# Provides functions to interact with the operating system, like file operations.
import smtplib  						# Used for sending emails via Simple Mail Transfer Protocol (SMTP).
import time  							# Used to represent time and to pause the program or introduce time-based functionality.
import requests  						# Allows sending and receiving HTTP requests and simplifies the process of interacting with websites.
import shutil  							# Provides functions for copying and creating files and directories.
from twilio.rest import Client  		# Used for making calls and sending messages via the Twilio API.
import ctypes  							# Provides access to Windows-specific functions and constants.
import datetime  						# Allows working with date and time data.
import winshell  						# Provides functionality for creating Windows shortcuts.
import pyjokes  						# Used to generate jokes.
import feedparser  						# Parses and processes RSS feeds.
import speech_recognition as sr  		# Used for speech recognition and voice input.
import json  							# Allows working with JSON data.
import random  							# Provides functions for generating random numbers.
import operator  						# Used for performing operations like sorting.
import tkinter  						# Used for building graphical user interfaces (GUIs).
from bs4 import BeautifulSoup  			# A library for parsing HTML and XML documents.
from urllib.request import urlopen  	# Allows opening URLs.
from ecapture import ecapture as ec 	# Provides functionality for capturing images or screens.
import win32com.client as wincl  		# Used for interfacing with COM objects on Windows.



# Initialize the text-to-speech engine
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

# Function to speak out text
def speak(audio):
    engine.say(audio)
    engine.runAndWait()

# Function to greet the user
def greet():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning My lord!")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon My lord!")

    else:
        speak("Good Evening My lord!")

    assname = "Lexi Lore"
    speak("I am your Assistant")
    speak(assname)

# Function to get the user's name
def username():
    speak("What should I call you, my lord?")
    user_input = takeCommand()
    print("User:", user_input)  # Print what the user said for debugging purposes

    # Check if the user provided a name with "call me" and extract the name
    if "call me" in user_input:
        uname = user_input.split("call me", 1)[1].strip()
    else:
        uname = user_input

    if uname:
        speak(f"Welcome, Lord {uname}")
    else:
        speak("Welcome, My Lord")
    speak("How can I help you, My Lord")


# Function to recognize user's voice input
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)
    
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User: {query}\n")
    except Exception as e:
        print(e)
        print("I'm sorry my lord but I am unable to recognize your voice.")
        return "None"
    return query

# Function to send an email
def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('your email id', 'your email password')
    server.sendmail('your email id', to, content)
    server.close()

if __name__ == '__main__':
    clear = lambda: os.system('cls')
    clear()
    greet()
    username()
    
    while True:
        query = takeCommand().lower()
        
        if 'open youtube' in query:
            speak("Opening YouTube")
            webbrowser.open("https://www.youtube.com")
        elif 'open google' in query:
            speak("Opening Google")
            webbrowser.open("https://www.google.com")
        elif 'play music' in query:
            music_dir = 'C:\\Path\\To\\Your\\Music\\Directory' 
            songs = os.listdir(music_dir)
            random.shuffle(songs)
            os.startfile(os.path.join(music_dir, songs[0]))
        elif 'tell me a joke' in query:
            joke = pyjokes.get_joke()
            speak(joke)
        elif 'take a screenshot' in query:
            ec.capture(0, "screenshot.png")
            speak("Screenshot taken, my lord!")
        elif 'send email' in query:
            try:
                speak("What should I say, my lord?")
                content = takeCommand()
                speak("To whom should I send the email, my lord?")
                to = input() 
                sendEmail(to, content)
                speak("Email has been sent, my lord!")
            except Exception as e:
                speak("I'm sorry, my lord. I couldn't send the email at the moment.")
        elif 'search in wikipedia' in query:
            speak("What should I search for, my lord?")
            search_query = takeCommand()
            try:
                result = wikipedia.summary(search_query, sentences=2)
                speak("According to Wikipedia, " + result)
            except wikipedia.exceptions.PageError:
                speak("I'm sorry, my lord. I couldn't find any information on that topic.")
            except wikipedia.exceptions.DisambiguationError:
                speak("There are multiple possible matches, my lord. Please be more specific.")
            except wikipedia.exceptions.HTTPTimeoutError:
                speak("I'm facing some connectivity issues, my lord. Please try again later.")
        elif 'time' in query:
            current_time = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"My lord, the current time is {current_time}")
        elif 'exit' in query:
            speak("Goodbye, my lord. Have a great day!")
            exit()
        # Add more commands here...


