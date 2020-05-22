from win32com.client import Dispatch
import datetime
import speech_recognition as babbu
import os
import wikipedia
import webbrowser
import random


def command():
    r = babbu.Recognizer()
    with babbu.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio=r.listen(source)

    try:
        print("getting...")
        query= r.recognize_google(audio,language='en-in')
        print(f"you said:{query}\n")
    except Exception as e:
        print("please say that again")
        return "none"
    return query

def kanwal(s):

    speak = Dispatch(("SAPI.SpVoice"))
    speak.speak(s)
def wish():
    hour= int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        kanwal("good morning kanwal")

    elif hour>=12 and hour<16:
        kanwal("good afternoon kanwal")
    

    else:
        kanwal("good evening kanwal")

    kanwal("im jarvis your personal assistent,what can i do for you")
    
if __name__ == '__main__':
    wish()   
    query=command().lower()
    if "wikipedia" in query:
        kanwal("searching in wikipedia...")
        #query=query.replace("wikipedia","")
        results= wikipedia.summary(query,sentences=2)
        kanwal("according to wikipedia")
        print(results)
        kanwal(results)
    elif "book" in query:
        with open("text.txt") as file:
            for item in file.readlines():
                kanwal("im gonna read from second lesson")
                kanwal(item)
    elif "youtube" in query:
        webbrowser.open("youtube.com")
    elif "stackoverflow" in query:
        webbrowser.open("stackoverflow.com")
    elif "github" in query:
        webbrowser.open("github.com")

    elif "song" in query:
        kanwal("im gonna play your favourite number")
        dir="C:\\songs"
        songs=os.listdir(dir)
        
        os.startfile(os.path.join(dir,songs[random.randint(1,15)]))

    elif "current time" in query:
        time= datetime.datetime.now().strftime("%H:%M:%S")
        kanwal(f"kanwal the current time is {time}")

    elif "movie" in query:
        kanwal("im gonna play your favourite number")
        dir="C:\\movies"
        movie=os.listdir(dir)
        os.startfile(os.path.join(dir,movie[random.randint(1,15)]))
    

    
