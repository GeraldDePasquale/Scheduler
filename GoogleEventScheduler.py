from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import datetime
from Center import Center

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

    def create_event_bodies(self, instructor, center):
        #ToDo create event data from instructor_schedule
        summary = instructor.name + '@' + center.name
        location = center.location
        description = 'You have bene scheduled to work today. Please be ready to serve students when your shift begins.'
        time_zone = center.location
        recurrence = 'RRULE:FREQ=DAILY;COUNT=1'
        center_email =
        event_bodies = []
        return event_bodies


    def insert_event(start_time, end_time, email_address):
        summary = 'Instructor_First_Name Work Event'
        location = 'Mathnasium of Stafford, 263 Garrisonville Road (Ste 104), Stafford, VA 22554'
        description = 'You work today. Please be ready to serve students when your shift begins.'
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
        #ToDo add calendar id to configuration file to enable mapping and setting master schedule & employee events
        this_event = service.events().insert(calendarId='primary', sendNotifications= True, body=this_event).execute()
        return this_event

    @staticmethod
    def insert_events(instructors):
        for instructor in instructors:
            self create_events(instructor.schedule)



if __name__ == '__main__':
    main()