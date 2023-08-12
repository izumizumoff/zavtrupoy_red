import datetime
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import config

# read&write mode
SCOPES = ['https://www.googleapis.com/auth/calendar']
# import token
creds = Credentials.from_authorized_user_info(config.TOKEN)
# main object
service = build('calendar', 'v3', credentials=creds)
calendar = config.CALENDAR

def addEvent(date,duration,title,description,place):
	# add events
    for_end = datetime.timedelta(minutes=duration)
    event = {
  'summary': title,
  'location': place,
  'description': description,
  'start': {
    'dateTime': f'{date.isoformat()}',
    'timeZone': 'Asia/Yekaterinburg',
  },
  'end': {
    'dateTime': f'{(date + for_end).isoformat()}',
    'timeZone': 'Asia/Yekaterinburg',
  },
}
    event = service.events().insert(calendarId='primary', body=event).execute()


def parseEvent(search_event, start, end):
    events_result = service.events().list(calendarId=calendar, timeMin=f'{start.isoformat()}Z', timeMax=f'{end.isoformat()}Z', singleEvents=True, orderBy='startTime', q=search_event).execute()

    events = events_result.get('items', [])

    if not events:
        return None
    else:
        list_result = []
        for event in events:
            result_dict = {}
            result_dict['summary'] = event.get('summary', '')
            result_dict['description'] = event.get('description', '').translate({ord(i): None for i in '/ui<>bnr'}).replace('\n', '')
            result_dict['place'] = event.get('location', '')
            result_dict['date'] = event['start']['dateTime'].split('T')[0]
            list_result.append(result_dict)
        return list_result

def searchActIndex(act, date):
	# search index of act from googlecalendar events
	# act: str --> название спектакля
	# date: datetime object --> дата от которой искать последний спектакль
    search_step = datetime.timedelta(days=365*3) # from period 3 years
    current_date = date
    search = f'{act}. Спектакль'
    events_result = service.events().list(calendarId=calendar, timeMin=f'{(current_date - search_step).isoformat()}Z', timeMax=f'{current_date.isoformat()}Z', singleEvents=True, q=search, orderBy='startTime').execute()
    events = events_result.get('items', [])
    if events:
        return int(events[-1].get('summary', 'Empty Summary').split('Спектакль')[1].split('.')[0].replace(' ', ''))
	# if not found -->
    return 0

def checkTurn(act, role, actor, date):
    # check whose turn is it to play this role
    # act: str --> название спектакля
    # role: str --> название роли
    # actor: str --> имя актера
    # date: datetime --> дата от которой искать последний спектакль
    search_step = datetime.timedelta(days=365*3) # from period 3 years
    current_date = date
    search = f'{act}. Спектакль'
    events_result = service.events().list(calendarId=calendar, timeMin=f'{(current_date - search_step).isoformat()}Z', timeMax=f'{current_date.isoformat()}Z', singleEvents=True, q=search, orderBy='startTime').execute()
    events = events_result.get('items', [])
    if events:
        if f'{actor}({role})' in events[-1].get('description', '').replace('<b>', '').replace('</b>', ''):
            return True
        else:
            return False
    else:
        return None