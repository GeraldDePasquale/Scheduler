from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import datetime

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'C:\ProgramData\MathnasiumScheduler\Working Directory\client_secret_calendar.json'
APPLICATION_NAME = 'Google Calendar API Python Quickstart'
home_dir = os.path.expanduser('~')
credential_dir = os.path.join(home_dir, '.credentials')
if not os.path.exists(credential_dir):
    os.makedirs(credential_dir)
credential_path = os.path.join(credential_dir,
                               'calendar-python-quickstart.json')
store = Storage(credential_path)
credentials = store.get()
if not credentials or credentials.invalid:
    flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
    flow.user_agent = APPLICATION_NAME
    if flags:
        credentials = tools.run_flow(flow, store, flags)
    else: # Needed only for compatibility with Python 2.6
        credentials = tools.run(flow, store)
    print('Storing credentials to ' + credential_path)

class GoogleEventScheduler:

    @staticmethod
    def insert_event(start_time, end_time, email_address):
        summary = 'Mathnasium Scheduler'
        location = 'Mathnasium of Stafford, 263 Garrisonville Road (Ste 104), Stafford, VA 22554'
        description = 'Work Event'
        time_zone = 'America/New_York'
        recurrence = 'RRULE:FREQ=DAILY;COUNT=1'
        center_email = 'stafford@mathnasium.com'
        this_event = {
          'summary': summary,
          'location': location,
          'description': description,
          'start': {'dateTime': start_time, 'timeZone': time_zone},
          'end': {'dateTime': end_time, 'timeZone': time_zone},
          'recurrence': [recurrence],
          'attendees': [{'email': email_address, 'email': center_email}],
          'reminders': {
            'useDefault': False,
            'overrides': [
              {'method': 'email', 'minutes': 24 * 60},
              {'method': 'popup', 'minutes': 60},
            ],
          },
        }
        http = credentials.authorize(httplib2.Http())
        service = discovery.build('calendar', 'v3', http=http)
        this_event = service.events().insert(calendarId='primary', sendNotifications= True, body=this_event).execute()
        return this_event