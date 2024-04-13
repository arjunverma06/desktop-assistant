import pyttsx3  # pip install pyttsx3
import speech_recognition as sr  # pip install speechRecognition
import wikipedia
import webbrowser
import os
import datetime
import smtplib
import operator
import ctypes
import wolframalpha
import pyjokes
# from ecapture import ecapture as ec
engine = pyttsx3.init("sapi5")
voices = engine.getProperty("voices")
# print(voices[1].id)
engine.setProperty("voice", voices[1].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning!")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon!")

    else:
        speak("Good Evening!")

    speak("I am   assistant of arjun verma. Please tell me how may I help you")


def takeCommand():
    # It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language='en-in') #Using google for voice recognition.
        print(f"User said: {query}\n")  #User query will be printed.

    except Exception as e:
        # print(e)    
        print("Say that again please...")   #Say that again will be printed in case of improper voice 
        return "None" #None string will be returned
    return query


def sendEmail(to, content):
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.ehlo()
    server.starttls()
    server.login("arjun225553@gmail.com", "csahxdncmmniojpi")
    server.sendmail("arjun225553@gmail.com", to, content)
    server.close()


if __name__ == "__main__":
    wishMe()

    while True:
        # if 1:
        query = takeCommand().lower()  # Converting user query into lower case

        # Logic for executing tasks based on query
        if (
            "wikipedia" in query
        ):  # if wikipedia found in the query then this block will be executed
            speak('Searching wikipedia" ...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif "open youtube" in query:
            webbrowser.open("youtube.com")
        elif "open google" in query:
            webbrowser.open("google.com")

        elif "open stack overflow" in query:
            webbrowser.open("stackoverflow.com")
        elif "play music" in query:
            music_dir = "E:\SONG"
            songs = os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir, songs[0]))
        elif "tell me time" in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            # print("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")
        elif "open code" in query:
            codePath = "C:\\Program Files\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)
        elif "shutdown" in query:
            os_dir = "C:\\Users\\Dell\\OneDrive\\Desktop\\arjun"
            
            shutdown = os.listdir(os_dir)
            print(shutdown)
            os.startfile(os.path.join(os_dir, shutdown[0]))

        elif "search" in query:
            webbrowser.open(f"www.google.com/search?q={query.replace('search','')}")

        # elif 'email to arjun' in query:
        # try:
        # speak("What should I say?")
        # content = takeCommand()
        #   to = "arjun225553@gmail.com"
        # sendEmail(to, content)
        #  speak("Email has been sent!")
        # except Exception as e:
        # print(e)
        # speak("Sorry my friend  arjun bhai. I am not able to send this email")

        elif "how are you" in query:
            speak("I am fine, Thank you sir")
            speak("How are you, Sir")
       

        elif "sketch design" in query:
            speak("ok i am draw the sketch  design sir")
            os_dir = "D:\\Majar project folder\\TASLLEM SIR PROJECT _voice_asistance\\radhakrishna code"
            design = os.listdir(os_dir)
            print(design)
            os.startfile(os.path.join(os_dir, design[0]))
            
        elif 'lock window' in query:
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()
                
        elif 'exit' in query:
            speak("Thanks for giving me your time")
            exit()
                 
                 
        elif 'joke' in query:
            speak(pyjokes.get_joke())
        elif "camera" in query or "take a photo" in query:
            ec.capture(0, "Jarvis Camera ", "img.jpg")
