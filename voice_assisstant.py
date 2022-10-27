from __future__ import print_function
import pickle
import datetime
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import pyttsx3
import speech_recognition as sr
import pytz
import subprocess
from subprocess import PIPE
import yagmail
import xlwt 
import xlrd
import webscraping
# If modifying these scopes, delete the file token.pickle.
#------------------------------GLOBAL VARIABLES--------------------------------
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
DAYS = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']
MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'September', 'October', 'November', 'december']
DAY_EXTENTIONS = ['rd', 'th', 'st', 'nd']
CALENDAR_STRS = ["what do i have", "do i have plans", "am i busy"]
NOTE_STRS = ["remmeber this", "make a note", "write this down"]
WAKE = "hey man"
QUIT = ["quit", "sleep", "exit", "turn off"]
GLOBALINT = 0
WEBSCRAPING_STR = ["get information of", "search for "]
#-------------------------------converting text to sound ----------------------
def speak(text):
	engine = pyttsx3.init()
	engine.say(text)
	engine.runAndWait()
#------------------------------getting events-------------------------	
def get_event(day, service):
	date = datetime.datetime.combine(day, datetime.datetime.min.time())
	end_date = datetime.datetime.combine(day, datetime.datetime.max.time())
	utc = pytz.UTC
	date = date.astimezone(utc)
	end_date = end_date.astimezone(utc)

	events_result = service.events().list(calendarId='primary',	timeMin=date.isoformat(), timeMax=end_date.isoformat(), singleEvents=True, orderBy='startTime').execute()
	events = events_result.get('items', [])

	if not events:
		speak('No upcoming events found.')
	else:
		speak(f"you have {len(events)} events on this day.")
	for event in events:
		start = event['start'].get('dateTime', event['start'].get('date'))
		print(start, event['summary'])
		start_time = str(start.split("T")[1].split("-")[0])
		if int(start_time.split(":")[0]) < 12:
			start_time = start_time + "am"
		else: 
			start_time = start_time + "pm"
		speak(event["summary"] + "at" + start_time)
#-----------------------------converting audio to text----------------


def get_audio():
	r = sr.Recognizer()
	with sr.Microphone() as source:
		audio = r.listen(source)
		said = ""
		
	try:
		said = r.recognize_google(audio)
		print(said)
	except Exception as e:
		print("Exception: " + str(e))
	return said
#-----------------------------------using google calndar ------------

def authorization_google():
	"""Shows basic usage of the Google Calendar API.
	Prints the start and name of the next 10 events on the user's calendar.
	"""
	creds = None
	# The file token.pickle stores the user's access and refresh tokens, and is
	# created automatically when the authorization flow completes for the first
	# time.
	if os.path.exists('token.pickle'):
		with open('token.pickle', 'rb') as token:
			creds = pickle.load(token)
	# If there are no (valid) credentials available, let the user log in.
	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			creds.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file(
				'credentials.json', SCOPES)
			creds = flow.run_local_server(port=0)
		# Save the credentials for the next run
		with open('token.pickle', 'wb') as token:
			pickle.dump(creds, token)

	service = build('calendar', 'v3', credentials=creds)
	return service
#---------------------------------planning for dates -----------------

def get_date(text):
	text = text.lower()
	today = datetime.date.today()

	if text.count("today") > 0:
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
			for ext in DAY_EXTENTIONS:
				found = word.find(ext)
				if found > 0 :
					try:
						day = int(word[:found])
					except:
						pass
# THE NEW PART STARTS HERE
	if month < today.month and month != -1: # if the month mentioned is before the current month set the year to the next
		year = year+1
	if month == -1 and day != -1: # if we didn't find a month, but we have a day
		if day < today.day:
			month = today.month + 1
		else:
			month = today.month
# if we only found a dta of the week
	if month == -1 and day == -1 and day_of_week != -1:
		current_day_of_week = today.weekday()
		dif = day_of_week - current_day_of_week
		if dif < 0:
			dif += 7
			if text.count("next") >= 1:
				dif += 7
		return today + datetime.timedelta(dif)
	if day != -1:
		return datetime.date(month=month, day=day, year=year)

#------------------opening note pad -----------------
def note(text):
	date = datetime.datetime.now()
	file_name = str(date).replace(":","-") + "-note.txt"
	with open(file_name,"w", encoding="utf-8") as f:
		f.write(text)

	subprocess.Popen(["notepad.exe", file_name])
	
'''def music():
	music = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Accessories\Windows Media Player.lnk"
	subprocess.run(shlex.split(music), stdout=PIPE, stderr=PIPE)
'''

#----------------sending email ------------------
def send_email(ad, message):
	receiver = "{}".format(ad)
	body = "{}".format(message)
	yag = yagmail.SMTP("a.mahdi1013@gmail.com",password = "m.ansari722")
	yag.send(
	to=receiver,
	subject="Python",
	contents=body,
	)
#------------------making an excel file ----------------
def make_excel_fisrt(username, password):
	workbook = xlwt.Workbook()  
  	
	sheet = workbook.add_sheet("Sheet Name") 

	style = xlwt.easyxf('font: bold 1') 
  
	sheet.write(GLOBALINT, GLOBALINT, username, style)
	sheet.write(GLOBALINT, GLOBALINT+1,password,style) 
	workbook.save("C:/Users/Mahdi/Desktop/env/datas.xls")

def excel_checker():
	return os.path.isfile('C:/Users/Mahdi/Desktop/env/datas.xls')

def read_use_passwd():
	loc = ("C:/Users/Mahdi/Desktop/env/datas.xls") 
	  
	# To open Workbook 
	wb = xlrd.open_workbook(loc) 
	sheet = wb.sheet_by_index(0) 
	  
	# For row 0 and column 0 
	return (sheet.cell_value(0, 0), sheet.cell_value(0,1))
	
#-------------------------main------------------------
SERVICE = authorization_google()
def main():
	
	print("start")
	text = get_audio().lower()
	for phrase in WEBSCRAPING_STR:
		if phrase in text:
			speak("what product are you looking for?")	
			product = get_audio().lower()
			scrap = webscraping.webscrape(product)
			note(scrap)

			
if __name__ == "__main__":
	main()


