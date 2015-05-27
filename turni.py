import gspread
import httplib2
import time
import locale
from collections import OrderedDict
from apiclient.discovery import build
from oauth2client.client import Credentials


## OAuth credentials functions
def __create_cal_service():
    with open('credentials', 'rw') as f:
        credentials = Credentials.new_from_json(f.read())
	http = httplib2.Http()
	http = credentials.authorize(http)
	return build('calendar', 'v3', http=http)

def __create_sheets_service():
	with open('credentials', 'rw') as f:
		credentials = Credentials.new_from_json(f.read())
	return credentials

## Spreadsheet Section
credentials = __create_sheets_service()
gc = gspread.authorize(credentials)

# Open a worksheet from spreadsheet with one shot
wks = gc.open("Turni operation 2014").sheet1
month = wks.title
if any ( [month == 'Aprile', month=='Giugno', month=='Settembre', month=='Novembre'] ):
	turn_list = wks.range('E2:E31')
	days_list = wks.range('A2:A31')
elif ( month == 'Febbraio' ):
	turn_list = wks.range('E2:E29')
	days_list = wks.range('A2:A29')
else:
	turn_list = wks.range('E2:E32')
	days_list = wks.range('A2:A32')
	
days = []
for member in days_list:
     #print member.value
     date_string=member.value+" "+month+" "+"2014"
     locale.setlocale(locale.LC_TIME, 'it_IT')
     date_obj = time.strptime(date_string, '%a %d %B %Y')
     locale.setlocale(locale.LC_TIME, 'en_US')
     date_string = time.strftime('%Y-%m-%d', date_obj)
     days.append(date_string)

turns = []
for member in turn_list:
    #print member.value
    turns.append(member.value)

dictionary = OrderedDict(zip(days, turns))
#print dictionary


## Calendar Section
service = __create_cal_service()
calendar='CALENDAR_ID'
oldEvents = service.events().list(calendarId=calendar).execute()
for event in oldEvents['items']:
    print event['id'] + ' ' + event['summary'] + ' ' + 'deleted'
    service.events().delete(calendarId=calendar, eventId=event['id']).execute()
for key in dictionary:
	print key , 'corresponds to', dictionary[key]
	if dictionary[key] == '1':
		startTime = '07:00:00'
		endTime = '16:00:00'
		event = {'summary' : 'Turno '+dictionary[key],'end' : {'timeZone' : 'Europe/Rome', 'dateTime' : key+'T'+endTime }, 'start' : { 'timeZone' : 'Europe/Rome', 'dateTime' : key+'T'+startTime }}
	elif dictionary[key] == '2':
		startTime = '09:00:00'
		endTime = '18:00:00'
		event = {'summary' : 'Turno '+dictionary[key],'end' : {'timeZone' : 'Europe/Rome', 'dateTime' : key+'T'+endTime }, 'start' : { 'timeZone' : 'Europe/Rome', 'dateTime' : key+'T'+startTime }}
	elif dictionary[key] == '3':
		startTime = '10:30:00'
		endTime = '19:30:00'
		event = {'summary' : 'Turno '+dictionary[key],'end' : {'timeZone' : 'Europe/Rome', 'dateTime' : key+"T"+endTime }, 'start' : { 'timeZone' : 'Europe/Rome', 'dateTime' : key+'T'+startTime }}
	else:
		continue
	request = service.events().insert(calendarId=calendar, body=event)
	response = request.execute()
	print response
