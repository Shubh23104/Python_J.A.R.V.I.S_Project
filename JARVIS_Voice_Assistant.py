from win32com.client import Dispatch
import speech_recognition as sr
import webbrowser
import wikipedia
import time as dt
from pygame import mixer as m
import os

r = sr.Recognizer()


def speek(text):
    speak = Dispatch("SAPI.SpVoice")

    speak.Speak(text)


def listen():
    try:
        with sr.Microphone() as source2:
            print("Listening...")
            r.adjust_for_ambient_noise(source2, duration=0.2)
            audio2 = r.listen(source2)

            print("Recognizing...")
            mytext = r.recognize_google(audio2)
            mytext = mytext.lower()

            print("did You Say:", mytext)

            return mytext

    except sr.RequestError as e:
        print("Could not request results: {0}".format(e))
        return "error"

    except sr.UnknownValueError:
        print("Unkonwn value error occured")
        return "error"


if __name__ == '__main__':
    t = dt.strftime("%H")

    if '0' <= t <= '11':
        speek("Good Morning Shubh")

    elif '12' <= t < '17':
        speek("Good Afternoon Shubh")

    elif '17' <= t < '19':
        speek("Good Evening Shubh")

    else:
        speek("Good night Shubh")

    speek("What May I Help You Shubh")

    while True:
        mytext = listen()

        if "open youtube" in mytext:
            webbrowser.get("C:/Program Files/Google/Chrome/Application/chrome.exe %s").open('youtube.com')

        elif "open google" in mytext:
            webbrowser.get("C:/Program Files/Google/Chrome/Application/chrome.exe %s").open('google.com')

        elif "open stackoverflow" in mytext:
            webbrowser.get("C:/Program Files/Google/Chrome/Application/chrome.exe %s").open('stackoverflow.com')

        elif 'play music' in mytext:
            li = os.listdir("E:\\python\\music.mp3")
            print(li)

            m.init()
            flag = False

            for song in li:
                m.music.load(f"E:\\python\\music.mp3\\{song}")
                m.music.play()

                while True:
                    responce = input("Enter p for pause the music, c for continue , s for stop and n for next song: ")
                    responce = responce.lower()

                    if 'p' == responce:
                        m.music.pause()

                    elif 'c' == responce:
                        m.music.unpause()

                    elif 'n' == responce:
                        m.music.stop()
                        break

                    elif 's' == responce:
                        flag = True
                        m.music.stop()
                        break

                if flag:
                    break


        elif "wikipedia" in mytext:
            mytext = str(mytext)
            li = mytext.split(" ")
            suma = wikipedia.summary(title=li[0], sentences=2)
            print("Summary: ", suma)
            speek(suma)

        elif "what's your name" in mytext or 'introduce your self' in mytext:
            speek("Hey My Name Is JARVIS I am Developed By Shubh Patel")

        elif 'close' in mytext:
            speek("Thank You Shubh I Hope You are Enjoy")
            break
