from __future__ import print_function
import pickle
import datetime
import os.path
import pyttsx3
import speech_recognition as sr
import pytz
import subprocess
import yagmail
import xlwt
import xlrd
import webscraping
from transformers import AutoTokenizer, AutoModelForCausalLM

# Google Calendar API dependencies
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# ------------------------------ GLOBAL VARIABLES ------------------------------
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'September', 'October', 'November', 'December']
DAY_EXTENSIONS = ['rd', 'th', 'st', 'nd']
CALENDAR_STRS = ["what do I have", "do I have plans", "am I busy"]
NOTE_STRS = ["remember this", "make a note", "write this down"]
WAKE = "hey man"
QUIT = ["quit", "sleep", "exit", "turn off"]
WEBSCRAPING_STR = ["get information of", "search for "]

# Load Llama model and tokenizer
tokenizer = AutoTokenizer.from_pretrained("allenai/llama")
model = AutoModelForCausalLM.from_pretrained("allenai/llama")

# ------------------------------ Utility Functions ------------------------------
def speak(text):
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()

def get_audio():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        audio = recognizer.listen(source)
        said = ""

    try:
        said = recognizer.recognize_google(audio)
        print(said)
    except Exception as e:
        print(f"Exception: {e}")

    return said.lower()

# ------------------------------ Google Calendar Functions ------------------------------
def authorization_google():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return build('calendar', 'v3', credentials=creds)

def get_event(day, service):
    date = datetime.datetime.combine(day, datetime.datetime.min.time())
    end_date = datetime.datetime.combine(day, datetime.datetime.max.time())
    utc = pytz.UTC
    date = date.astimezone(utc)
    end_date = end_date.astimezone(utc)

    events_result = service.events().list(calendarId='primary', timeMin=date.isoformat(), 
                                          timeMax=end_date.isoformat(), singleEvents=True, orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        speak('No upcoming events found.')
    else:
        speak(f"You have {len(events)} events on this day.")
    
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        print(start, event['summary'])
        start_time = str(start.split("T")[1].split("-")[0])
        start_time = start_time + "am" if int(start_time.split(":")[0]) < 12 else start_time + "pm"
        speak(f"{event['summary']} at {start_time}")

# ------------------------------ Date Functions ------------------------------
def get_date(text):
    text = text.lower()
    today = datetime.date.today()

    if "today" in text:
        return today

    day = -1
    day_of_week = -1
    month = -1
    year = today.year

    for word in text.split(" "):
        if word in MONTHS:
            month = MONTHS.index(word) + 1
        elif word in DAYS:
            day_of_week = DAYS.index(word)
        elif word.isdigit():
            day = int(word)
        else:
            for ext in DAY_EXTENSIONS:
                if word.endswith(ext):
                    try:
                        day = int(word[:-2])
                    except ValueError:
                        pass

    if month < today.month and month != -1:
        year += 1
    if month == -1 and day != -1:
        month = today.month + 1 if day < today.day else today.month

    if day_of_week != -1:
        current_day_of_week = today.weekday()
        dif = day_of_week - current_day_of_week
        dif += 7 if dif < 0 and "next" in text else 0
        return today + datetime.timedelta(dif)
    elif day != -1:
        return datetime.date(month=month, day=day, year=year)

# ------------------------------ Note Functions ------------------------------
def note(text):
    date = datetime.datetime.now()
    file_name = str(date).replace(":", "-") + "-note.txt"
    with open(file_name, "w", encoding="utf-8") as f:
        f.write(text)
    subprocess.Popen(["notepad.exe", file_name])

# ------------------------------ Email Functions ------------------------------
def send_email(address, message):
    yag = yagmail.SMTP("EMAIL", password="PASSWORD")
    yag.send(to=address, subject="Python", contents=message)

# ------------------------------ Excel Functions ------------------------------
def make_excel_first(username, password):
    workbook = xlwt.Workbook()  
    sheet = workbook.add_sheet("Sheet Name") 
    style = xlwt.easyxf('font: bold 1') 
    sheet.write(0, 0, username, style)
    sheet.write(0, 1, password, style)
    workbook.save("ADDRESS")

def excel_checker():
    return os.path.isfile('ADDRESS')

def read_user_passwd():
    loc = "C:/Users/Mahdi/Desktop/env/datas.xls"
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    return sheet.cell_value(0, 0), sheet.cell_value(0, 1)

# ------------------------------ OpenAI Connection ------------------------------
def call_llm(text):
    inputs = tokenizer(text, return_tensors="pt", padding=True, truncation=True, max_length=512)
    with torch.no_grad():
        outputs = model.generate(inputs.input_ids)
    llm_response = tokenizer.decode(outputs[0], skip_special_tokens=True)
    return llm_response

# ------------------------------ Main Function ------------------------------
def main():
    print("Start")
    text = get_audio()
    llm_response = call_llm(text)
    speak(llm_response)
    
    for phrase in WEBSCRAPING_STR:
        if phrase in text:
            speak("What product are you looking for?")
            product = get_audio()
            scrap = webscraping.webscrape(product)
            llm_scrap_response = call_llm(scrap)
            speak(llm_scrap_response)

if __name__ == "__main__":
    SERVICE = authorization_google()
    main()
