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
import win32com.client #this for text to speak coverted
import openai
from ecapture import ecapture as ec
# from config import apikey
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

    speak("I am  your virtual assistant . Please tell me how may I help you")
    
# def sendEmail(to, content):
#     server = smtplib.SMTP('smtp.gmail.com', 587)
#     server.ehlo()
#     server.starttls()
     
#     # Enable low security in gmail
#     server.login('arjun225553@gamil.com', '')
#     server.sendmail('arjun225553@gamil.com', to, content)
#     server.close() 
def ai(prompt):
     openai.api_key = apikey
     text = f"OpenAI response for Prompt: {prompt} \n *************************\n\n"

     response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    
    # todo: Wrap this inside of a  try catch block
    # print(response["choices"][0]["text"])
     text += response["choices"][0]["text"]
     if not os.path.exists("Openai"):
        os.mkdir("Openai")

    # with open(f"Openai/prompt- {random.randint(1, 2343434356)}", "w") as f:
     with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip() }.txt", "w") as f:
        f.write(text)

def say(text):
    os.system(f'say "{text}"')

  

    
def takeCommand():
    # It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 0.5
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
    
    speaker=win32com.client.Dispatch("SAPI.SpVoice") 
    while 1:
       print("entert the word") 
       s=input()
       speaker.Speak(s)






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
            music_dir = "E:\\Mp3 song\\Mp3 song"
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
            speak("Hold On a Sec ! Your system is on its way to shut down")
            os_dir = "C:\\Users\\Dell\\OneDrive\\Desktop\\arjun"
            
            shutdown = os.listdir(os_dir)
            print(shutdown)
            os.startfile(os.path.join(os_dir, shutdown[0]))
        elif "Using artificial intelligence".lower() in query.lower():
            ai(prompt=query)

        elif "search" in query:
            webbrowser.open(f"www.google.com/search?q={query.replace('search','')}")

        elif "how are you" in query:
            speak("I am fine, Thank you sir")
            speak("How are you, Sir")
       

        elif "sketch design" in query:
            speak("ok i am draw the sketch  design sir")
            os_dir = "D:\\Majar project folder\\TASLLEM SIR PROJECT _voice_asistance\\radhakrishna code\\"
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
        elif "open" in query:
             webbrowser.open("chat.openai.com")  
        # elif 'send a mail' in query:
        #     try:
        #         speak("What should I say?")
        #         content = takeCommand()
        #         speak("whome should i send")
        #         to = input()    
        #         sendEmail(to, content)
        #         speak("Email has been sent !")
        #     except Exception as e:
        #         print(e)
        #         speak("I am not able to send this email") 
        elif "change my name to" in query:
            query = query.replace("change my name to", "")
            assname = query 
            
        elif "change name" in query:
            speak("What would you like to call me, Sir ")
            assname = takeCommand()
            speak("Thanks for naming me")
        elif "who made you" in query or "who created you" in query: 
            speak("I have been created by Arjun Abhishek Aniket.")
        elif 'powerpoint presentation' in query:
            speak("opening Power Point presentation")
            power = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs"
            os.startfile(power) 
        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20, 
                                                       0, 
                                                       "https://images.pexels.com/photos/33109/fall-autumn-red-season.jpg?auto=compress&cs=tinysrgb&w=600.png",
                                                       0)
            speak("Background changed successfully")


        
