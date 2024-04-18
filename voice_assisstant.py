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
    """
    Converts text to speech.
    
    Input:
    - text (str): Text to be spoken.
    
    Output:
    - None
    """
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()

def get_audio():
    """
    Records audio from microphone and converts it to text.
    
    Input:
    - None
    
    Output:
    - str: Recognized text from audio.
    """
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
     """
    Authorizes and builds the Google Calendar service.
    
    Input:
    - None
    
    Output:
    - Google Calendar service object
    """
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
    """
    Fetches and speaks the events for a given day.
    
    Input:
    - day (datetime.date): Date for which events are to be fetched.
    - service: Google Calendar service object
    
    Output:
    - None
    """
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
    """
    Parses and returns the date mentioned in the input text.
    
    Input:
    - text (str): Input text containing date information
    
    Output:
    - datetime.date: Parsed date
    """
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
    """
    Saves a note with the provided text to a text file.
    
    Input:
    - text (str): Text to be saved in the note
    
    Output:
    - None
    """
    date = datetime.datetime.now()
    file_name = str(date).replace(":", "-") + "-note.txt"
    with open(file_name, "w", encoding="utf-8") as f:
        f.write(text)
    subprocess.Popen(["notepad.exe", file_name])

# ------------------------------ Email Functions ------------------------------
def send_email(address, message):
     """
    Sends an email to the specified address with the given message.
    
    Input:
    - address (str): Email address of the recipient
    - message (str): Email message content
    
    Output:
    - None
    """
    yag = yagmail.SMTP("EMAIL", password="PASSWORD")
    yag.send(to=address, subject="Python", contents=message)

# ------------------------------ Excel Functions ------------------------------
def make_excel_first(username, password):
    """
    Creates an Excel file with the provided username and password.
    
    Input:
    - username (str): Username to be saved
    - password (str): Password to be saved
    
    Output:
    - None
    """
    workbook = xlwt.Workbook()  
    sheet = workbook.add_sheet("Sheet Name") 
    style = xlwt.easyxf('font: bold 1') 
    sheet.write(0, 0, username, style)
    sheet.write(0, 1, password, style)
    workbook.save("ADDRESS")

def excel_checker():
     """
    Checks if the Excel file exists.
    
    Input:
    - None
    
    Output:
    - bool: True if file exists, False otherwise
    """
    return os.path.isfile('ADDRESS')

def read_user_passwd():
    """
    Reads and returns the username and password from the Excel file.
    
    Input:
    - None
    
    Output:
    - tuple: (username, password)
    """
    loc = "ADDRESS"
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    return sheet.cell_value(0, 0), sheet.cell_value(0, 1)

# ------------------------------ OpenAI Connection ------------------------------
def call_llm(text):
     """
    Generates text using the LLM model from the input text.
    
    Input:
    - text (str): Input text for text generation
    
    Output:
    - str: Generated text
    """
    inputs = tokenizer(text, return_tensors="pt", padding=True, truncation=True, max_length=512)
    with torch.no_grad():
        outputs = model.generate(inputs.input_ids)
    llm_response = tokenizer.decode(outputs[0], skip_special_tokens=True)
    return llm_response

# ------------------------------ Main Function ------------------------------
def main():
    """
    Main function to execute the program.
    
    Input:
    - None
    
    Output:
    - None
    """
    print("Start")
    # Record audio from microphone and convert it to text
    text = get_audio()
    
    # Generate text response using the LLM model based on the recorded text
    llm_response = call_llm(text)
    
    # Speak the generated LLM response
    speak(llm_response)
    
    # Loop through predefined web scraping phrases
    for phrase in WEBSCRAPING_STR:
        # Check if the recorded text contains a web scraping phrase
        if phrase in text:
            # Ask the user what product they are looking for
            speak("What product are you looking for?")
            
            # Record audio from microphone to capture the product name
            product = get_audio()
            
            # Scrape information related to the product
            scrap = webscraping.webscrape(product)
            
            # Generate text response using the LLM model based on the scraped information
            llm_scrap_response = call_llm(scrap)
            
            # Speak the generated LLM response related to the scraped information
            speak(llm_scrap_response)


if __name__ == "__main__":
    SERVICE = authorization_google()
    main()
