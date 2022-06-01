import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import smtplib
import imaplib, email
from email.header import decode_header
import webbrowser as wb
import os
import time
import psutil
import pyautogui
import pyjokes

engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

global memory
memory = "nothing yet!"

def time_():
    time = datetime.datetime.now().strftime("%I:%M:%S")
    speak("current time is ")
    speak(time)

def year_to_string(year):
    thousands = int(year / 1000)
    hundreds = int((year - (thousands * 1000)) / 100)
    tens = int((year - (thousands * 1000) - (hundreds * 100)) / 10)
    ones = int(year % 10)
    tens_ones = (tens * 10) + ones
    hundreds_tens_ones = tens_ones + (hundreds * 100)
    answer = ""
    if thousands != 0:
        if hundreds_tens_ones != 0:
            answer += str(thousands) + "thousand"
        else:
            answer += str(thousands) + "thousandth"
    if hundreds != 0:
        if tens_ones != 0:
            answer += str(hundreds) + "hundred"
        else:
            answer += str(hundreds) + "hundredth"
    if tens != 0:
        if tens == 2:
            answer += "twenty"
        elif tens == 3:
            answer += "thirty"
        elif tens == 4:
            answer += "forty"
        elif tens == 5:
            answer += "fifty"
        elif tens == 6:
            answer += "sixty"
        elif tens == 7:
            answer+= "seventy"
        elif tens == 8:
            answer += "eighty"
        elif tens == 9:
            answer += "ninety"
        else:
            if tens_ones == 11:
                answer += "eleventh"
            elif tens_ones == 12:
                answer += "twelfth"
            elif tens_ones == 13:
                answer += "thirteenth"
            elif tens_ones == 14:
                answer += "fourteenth"
            elif tens_ones == 15:
                answer += "fifteenth"
            elif tens_ones == 16:
                answer += "sixteenth"
            elif tens_ones == 17:
                answer += "seventeenth"
            elif tens_ones == 18:
                answer += "eighteenth"
            elif tens_ones == 19:
                answer += "nineteenth"
            elif tens_ones == 10:
                answer += "tenth"
    if tens != 1:
        if ones == 1:
            answer += "first"
        elif ones == 2:
            answer += "second"
        elif ones == 3:
            answer += "third"
        elif ones == 4:
            answer += "fourth"
        elif ones == 5:
            answer += "fifth"
        elif ones == 6:
            answer += "sixth"
        elif ones == 7:
            answer += "seventh"
        elif ones == 8:
            answer += "eightieth"
        elif ones == 9:
            answer += "ninth"

    return answer

def day_to_string(day):
    answer = ""
    ones = day % 10
    tens = (day - ones) / 10
    if tens <= 0:
        if ones == 1:
            answer = "first"
        elif ones == 2:
            answer = "second"
        elif ones == 3:
            answer = "third"
        elif ones == 4:
            answer = "fourth"
        elif ones == 5:
            answer = "fifth"
        elif ones == 6:
            answer = "sixth"
        elif ones == 7:
            answer = "seventh"
        elif ones == 8:
            answer = "eightieth"
        elif ones == 9:
            answer = "ninth"
    elif tens == 1:
        if ones == 0:
            answer = "tenth"
        elif ones == 1:
            answer = "eleventh"
        elif ones == 2:
            answer = "twelfth"
        elif ones == 3:
            answer = "thirteenth"
        elif ones == 4:
            answer = "fourteenth"
        elif ones == 5:
            answer = "fifteenth"
        elif ones == 6:
            answer = "sixteenth"
        elif ones == 7:
            answer = "seventeenth"
        elif ones == 8:
            answer = "eighteenth"
        else:
            answer = "nineteenth"
    elif tens == 2:
        if ones == 0:
            answer = "twentieth"
        else:
            answer += "twenty"
            if ones == 1:
                answer += "first"
            elif ones == 2:
                answer += "second"
            elif ones == 3:
                answer += "third"
            elif ones == 4:
                answer += "fourth"
            elif ones == 5:
                answer += "fifth"
            elif ones == 6:
                answer += "sixth"
            elif ones == 7:
                answer += "seventh"
            elif ones == 8:
                answer += "eightieth"
            elif ones == 9:
                answer += "ninth"
    else:
        if ones == 0:
            answer = "thirtieth"
        elif ones == 1:
            answer = "thirty first"
    return answer

def date_():
    year = datetime.datetime.now().year
    month = datetime.datetime.now().month
    day = datetime.datetime.now().day
    speak("Today is ")
    speak(str(year_to_string(year)))
    speak("year")
    speak(str(day_to_string(day)) + "th")
    if month == 1:
        speak("January")
    elif month == 2:
        speak("February")
    elif month == 3:
        speak("March")
    elif month == 4:
        speak("April")
    elif month == 5:
        speak("May")
    elif month == 6:
        speak("June")
    elif month == 7:
        speak("July")
    elif month == 8:
        speak("August")
    elif month == 9:
        speak("September")
    elif month == 10:
        speak("October")
    elif month == 11:
        speak("November")
    else:
        speak("December")

def greeting():
    hour = datetime.datetime.now().hour

    if hour >= 6 and hour < 12:
        speak("Good Morning Sir!")
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon Sir!")
    elif hour >= 18 and hour < 24:
        speak("Good Evening Sir!")
    else:
        speak("Good Night Sir!")

def send_mail(to, content, subject):
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.ehlo()
        server.starttls()
        # message = 'Subject: {}\n\n{}'.format(subject, content)
        server.login("email","password")
        server.sendmail("email",to,content)
        server.close()
        speak("Email sent successfully!")
    except Exception as e:
        speak("Exception occurred!")

def check_mail(user):
    if user == "main" or user == "first":
        username = "email"
        password = "password"
    else:
        username = "email"
        password = "password"

    def clean(text):
        return "".join(c if c.isalnum() else "_" for c in text)
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login(username, password)

    status, messages = imap.select("INBOX")
    N = 3
    messages = int(messages[0])
    for i in range(messages, messages-N, -1):
        res, msg = imap.fetch(str(i),"(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1])
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding)
                From, encoding = decode_header(msg.get("From"))[0]
                if isinstance(From, bytes):
                    From = From.decode(encoding)
                print("Subject: ", subject)
                print("From: ", From)
                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            print(body)
                        elif "attachment" in content_disposition:
                            filename = part.get_filename()
                            if filename:
                                folder_name = clean(subject)
                                if not os.path.isdir(folder_name):
                                    os.mkdir(folder_name)
                                filepath = os.path.join(folder_name, filename)
                                # open(filepath, "wb").write(part.get_payload(decode=True))
                else:
                    content_type = msg.get_content_type()
                    body = msg.get_payload(decode=True).decode()
                    if content_type == "text/plain":
                        # print only text email parts
                        print(body)
                if content_type == "text/html":
                    # if it's HTML, create a new HTML file and open it in browser
                    folder_name = clean(subject)
                    if not os.path.isdir(folder_name):
                        # make a folder for this email (named after the subject)
                        os.mkdir(folder_name)
                    filename = "index.html"
                    filepath = os.path.join(folder_name, filename)
                    # write the file
                    open(filepath, "w").write(body)
                    # open in the default browser
                    wb.open(filepath)
                print("=" * 100)
                # close the connection and logout
                imap.close()
                imap.logout()

def cpu():
    battery = psutil.sensors_battery().percent
    if battery == 100:
        speak("Yor battery is full!")
    elif battery == 50:
        speak("Your battery is half")
    elif battery < 15 and battery > 5:
        speak("Your battery is below fifteen percent!")
    elif battery < 5:
        speak("Your battery is running out, please plug in!")
    else:
        speak("Battery is at " + str(battery))

def joke():
    speak(pyjokes.get_joke())

def find_location(query):
    if ("find" in query):
        query = query.replace("find","")
    if "show" in query:
        query = query.replace("show", "")
    if ("location" in query):
        query = query.replace("location","")
    if ("address" in query):
        query = query.replace("address","")
    speak("showing the location")
    wb.open("https://www.google.com/maps/place/" + query)

def screenshot():
    img = pyautogui.screenshot()
    img.save('C:/Users/hp/Desktop/screenshot.png')
    speak("Screenshot saved on desktop")


def speak(audio):
    global memory
    memory = audio
    engine.say(audio)
    engine.runAndWait()


def take_command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)
    try:
        print("Recognizing.....")
        query = r.recognize_google(audio, language='en-US').lower()
        print(query)
    except Exception as e:
        print(e)
        print("Say that again.....")
        return "None"
    return query

def process(empty, just):
    if empty == True:
        query = take_command()
    else:
        query = just
    if (("time" in query) and ("now" in query)) or (("what" in query) and ("time" in query)):
        time_()
    elif (("date" in query) and ("today" in query)) or (("what" in query) and ("day" in query) and ("today" in query)):
        date_()
    elif "wikipedia" in query:
        speak("Searching...")
        query = query.replace("wikipedia", "")
        failure = False
        try:
            result = wikipedia.summary(query, sentences=3)
        except Exception as e:
            failure = True
            pass
        if failure:
            speak("Failure occurred!!!")
            process(True, "None")
        else:
            speak("Here is what I found from wikipedia about")
            speak(query)
            print(result)
            speak("Want me to read that?")
            answer = "None"
            while answer == "None":
                answer = take_command()
            if ("yes" in answer) or ("yeah" in answer) or ("why not" in answer):
                speak(result)
        process(True, "None")
    elif ("send mail" in query) or ("send email" in query) or ("send letter" in query):
        speak("To whom?")
        receiver = input("Insert email address:")
        to = receiver
        speak("What should I say?")
        content = take_command()
        while content == "None":
            content = take_command()
        send_mail(to, content, "TEST Subject")
    elif ("check email" in query) or ("check mail" in query) or ("check my email" in query) or ("check my mail" in query) or ("open my email" in query) or ("open my mail" in query) or (("news" in query) and ("email" in query) and ("my" in query)):
        try:
            print("Checking....")
            speak("Checking....")
            speak("Here you go!")
            check_mail("lol")
        except Exception as e:
            pass
    elif (("thank" in query) and ("you" in query) or ("thanks" in query)):
        speak("You are always welcome")
    elif ("what is your name" in query) or (("what" in query) and ("your" in query) and ("name" in query)) or (("your" in query) and ("name" in query)):
        speak("My name is Veronica")
    elif ("hello" in query) or ("hi " in query) or ("good" in query and (("morning" in query) or ("afternoon" in query) or ("evening" in query) or ("night" in query))):
        greeting()
        speak("How are you doing?")
    elif (("search" in query) or ("open" in query)) and (("internet" in query) or ("google" in query) or ("chrome" in query)):
        speak("what you want to search?")
        answer = "None"
        while answer == "None":
            answer = take_command()
        chrome_path = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'
        wb.get(chrome_path).open_new_tab("https://www.google.com/search?q=" + answer + "&rlz=1C1GCEA_enUZ938UZ938&sxsrf=ALiCzsbqICJxTzjJW2HF5l-ONOpwhsOXNg%3A1653344647243&ei=hwmMYum_DqKSxc8P_82iwAo&ved=0ahUKEwipn4_j1Pb3AhUiSfEDHf-mCKgQ4dUDCA4&uact=5&oq=apulatjonov&gs_lcp=Cgdnd3Mtd2l6EAMyBggAEA0QCjIICAAQHhANEAoyCAgAEB4QDRAKMgoIABAeEA8QDRAKMgoIABAeEA0QBRAKMgwIABAeEA8QDRAFEAoyCggAEB4QDxANEAoyCggAEB4QDxANEAo6BwgAEEcQsAM6BwgAELADEEM6BAgjECc6BQgAEJECOggILhDUAhCRAjoLCC4QgAQQxwEQ0QM6CwguEIAEEMcBEKMCOgUIABCABDoECAAQQzoKCC4QxwEQowIQQzoICC4QgAQQ1AI6BQguEIAEOgcIABCABBAKOgQIABAKSgQIQRgASgQIRhgAUKMFWLMPYLAQaAJwAXgAgAHXAYgBiA2SAQUwLjguMpgBAKABAcgBCsABAQ&sclient=gws-wiz")
    elif (("search" in query) or ("open" in query)) and (("youtube" in query) or (("you" in query) and ("tube" in query))):
        speak("what you want to search?")
        answer = "None"
        while answer == "None":
            answer = take_command()
        chrome_path = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'
        wb.get(chrome_path).open_new_tab("https://www.youtube.com/search?q=" + answer + "&rlz=1C1GCEA_enUZ938UZ938&sxsrf=ALiCzsbqICJxTzjJW2HF5l-ONOpwhsOXNg%3A1653344647243&ei=hwmMYum_DqKSxc8P_82iwAo&ved=0ahUKEwipn4_j1Pb3AhUiSfEDHf-mCKgQ4dUDCA4&uact=5&oq=apulatjonov&gs_lcp=Cgdnd3Mtd2l6EAMyBggAEA0QCjIICAAQHhANEAoyCAgAEB4QDRAKMgoIABAeEA8QDRAKMgoIABAeEA0QBRAKMgwIABAeEA8QDRAFEAoyCggAEB4QDxANEAoyCggAEB4QDxANEAo6BwgAEEcQsAM6BwgAELADEEM6BAgjECc6BQgAEJECOggILhDUAhCRAjoLCC4QgAQQxwEQ0QM6CwguEIAEEMcBEKMCOgUIABCABDoECAAQQzoKCC4QxwEQowIQQzoICC4QgAQQ1AI6BQguEIAEOgcIABCABBAKOgQIABAKSgQIQRgASgQIRhgAUKMFWLMPYLAQaAJwAXgAgAHXAYgBiA2SAQUwLjguMpgBAKABAcgBCsABAQ&sclient=gws-wiz")
    elif (("check" in query) or (("What" in query) and ("my" in query))) and ("battery" in query):
        cpu()
    elif ((("tell" in query) or ("can") in query) and ("joke" in query)) or (("tell" in query) and ("funny" in query)):
        joke()
        speak("haha")
    elif (("go" in query) and ("offline" in query)) or (("shut" in query) and ("down" in query)):
        speak("Going Offline!")
        quit()
    elif ("telegram" in query) and ("open" in query):
        speak("Opening telegram...")
        telegram_path = r'C:/Users/hp/AppData/Roaming/Telegram Desktop/Telegram.exe'
        os.startfile(telegram_path)
    elif (("take" in query) or ("write" in query)) and ("note" in query):
        speak("what should I write, Sire?")
        notes = "None"
        silent = False
        while notes == "None":
            notes = take_command()
            if notes == "None":
                speak("Couldn't get Sir, can you repeat please?")
                silent = True
        if silent:
            speak("Oh, OK!")
        file = open('C:/Users/hp/Desktop/notes.txt', 'a')
        speak("Sire should I include date and time?")
        ans = "None"
        while ans == "None":
            ans = take_command()
        if ("yes" in ans) or ("yeah" in ans) or ("sure" in ans):
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            file.write(strTime)
            file.write(":-")
        file.write(notes + "\n")
        speak("Done!")
    elif (("show" in query) or ("what" in query)) and ("note" in query):
        speak("Showing notes!")
        file = open('C:/Users/hp/Desktop/notes.txt', "r")
        print(file.read())
    elif ("snapshot" in query) or ("screenshot" in query):
        screenshot()
    elif (("find" in query) or ("show" in query)) and (("location" in query) or ("address" in query)):
        find_location(query)
    elif ("log out" in query):
        speak("Logging out")
        os.system("shutdown -l")
    elif ("switch off" in query or "turn off" in query) and ("pc" in query or "computer" in query):
        speak("Shutting down the system")
        os.system("shutdown /s /t 1")
    elif ("restart" in query and ("pc" in query or "system" in query)) or "restart" in query:
        speak("Restarting...")
        os.system("shutdown /r /t 1")
    elif ("repeat" in query) or ("what" in query and "say" in query and "you" in query):
        engine.say("I said")
        speak(memory)




if __name__ == "__main__":
    greeting()
    while True:
        process(True, "")
