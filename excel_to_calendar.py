#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: Corbett Redden
This program allows me to easily create lots of events at once on my 
Google Calendar. I enter all the information in a specially formatted
Excel spreadsheet. Then, I run the Python command 
”xlsx_to_calendar(‘cal_template.xlsx’, ‘SheetName’)" 
where the first argument is the Excel filename, and the second (optional)
argument is the Sheet Name (default is 1st sheet). The program then 
automatically creates all the individual events in my Google Calendar.

Note: must have "credentials.json" saved in same directory. This is obtained
from Google https://developers.google.com/workspace/guides/create-credentials

Some initial code for Google Calendar interface was borrowed from
https://developers.google.com/calendar/api/quickstart/python
https://karenapp.io/articles/how-to-automate-google-calendar-with-python-using-the-calendar-api/
"""
from __future__ import print_function

import datetime as dt
import os.path
import pandas as pd
import warnings

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar']


def xlsx_to_calendar(file, sheet=None):
    """
    Primary function that does everything:
        - connects to Google Calendar
        - converts input .xlsx 'file' to Pandas DataFrame
        - creates the appropriate dictionary (json) for each row
        - uploads each of these events to Google Calendar
        
    Inputs:
        file: str of xlsx filename
        sheet: (optional) str of sheet name, otherwise use 1st sheet
    """
    print("Connecting to Google Calendar...")
    service = get_calendar_service()
    print("Converting spreadsheet...")
    df = xlsx_to_df(file, sheet)
    df = df_cal_format(df, service)
    print("Updating Google Calendar...")
    results = df.apply(create_new_event, axis=1, service=service)
    print("Success!")
    return results

    
def get_calendar_service():
    """ returns the Google Calendar 'service' """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('calendar', 'v3', credentials=creds)
        return service
    except HttpError as error:
        print('An error occurred: %s' % error)


def xlsx_to_df(file, sheet=None):
    """ Inputs xlsx file, outputs df, no reformatting yet """
    # My excel template uses data validation to select items from list, and 
    # read_excel warns that it can't support this. I am ignoring the warning.
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Data Validation extension is not supported", category=UserWarning)
        if sheet:
            df = pd.read_excel(file, sheet_name=sheet)
        else:
            df = pd.read_excel(file)
    return df


def df_cal_format(df, service):
    """ Takes excel df, creates needed columns for google cal """
    # rename some columns
    column_dict = {'Summary' : 'summary', 'Description' : 'description'}
    df = df.rename(columns=column_dict)

    # If timezone isn't specified, will use default timezone
    tz = default_timezone(service)
    df['Start Time Zone'].fillna(tz, inplace=True)
    df['End Time Zone'].fillna(tz, inplace=True)
    df['description'].fillna("", inplace=True)  # So blank descriptions don't say NaN
    df['Calendar Name'].fillna('Primary', inplace=True)  # default to Primar calendar

    # lookup calendar ids from calendar names
    calendars = linked_calendars(service)
    calendars['Primary'] = 'primary'   # can use 'primary' instead of email address/id
    df['id'] = df['Calendar Name'].map(calendars)

    # Create start/end information
    df[["start", "end"]] = df.apply(format_date, axis="columns", 
                                    result_type='expand').astype("object")  
    
    return df


def default_timezone(service):
    """ return the default timezone (str) of user """
    settings = service.settings().list().execute()
    return [s['value'] for s in settings['items'] if s['id']=='timezone'][0]

    
def linked_calendars(service, can_edit=True):
    """ return a dict of linked calendars (that can be edited) """
    calendars_result = service.calendarList().list().execute()
    calendars = calendars_result.get('items', [])  # list of calendars

    if can_edit:
        access_roles = {'owner', 'writer'}
        return { cal['summary'] : cal['id'] for cal in calendars
                 if cal['accessRole'] in access_roles }       
    else:
        return { cal['summary'] : cal['id'] for cal in calendars }


def format_date(row):
    """ Input row in df, create appropriate start/end info for google cal """
    # Check that the "Start Date" is not blank
    if pd.isna(row["Start Date"]):
        raise Exception("Event must contain a Start Date")
    
    # All-day event
    if pd.isna(row["Start Time"]):
        start_date = str(row["Start Date"].date())
        if pd.isna(row["End Date"]):  # Single-day event
            end_date = start_date
        else:
            end_date = str(row["End Date"].date())  # Multi-day event
        return [ {"date" : start_date}, {"date" : end_date} ]
 
    # Include Start/Stop times
    if pd.notna(row["Start Time"]): 
        d = row["Start Date"]  # pd Timestamp format
        t = row["Start Time"]  # datetime.time format        
        start_datetime = dt.datetime(d.year, d.month, d.day, t.hour, t.minute)

        if pd.isna(row["End Time"]):  # Default to 1-hour duration
            end_datetime = start_datetime + dt.timedelta(hours=1)
        elif pd.isna(row["End Date"]):  # One-day event
            t = row["End Time"]
            end_datetime = dt.datetime(d.year, d.month, d.day, t.hour, t.minute)
        else:  # Specific start/stop times, but on different days
            d = row["End Date"]  # pd Timestamp format
            t = row["End Time"]  # datetime.time format        
            end_datetime = dt.datetime(d.year, d.month, d.day, t.hour, t.minute)
            
        return [
            {"dateTime" : start_datetime.isoformat(), 
             "timeZone" : row["Start Time Zone"]},
            {"dateTime" : end_datetime.isoformat(), 
             "timeZone" : row["End Time Zone"]}
            ]


def create_new_event(row, service):
    """ Create new Google Cal event from df row """
    cal_id = row["id"]    
    event_dict = dict(row[["summary", "description", "start", "end"]])
    result = service.events().insert(calendarId=cal_id, body=event_dict).execute()
    return result


#%%
# Print function that aren't needed, but they are useful when first
# checking you are connected to the google calendar.
def print_upcoming_events(service, num_events=5):
    now = dt.datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
    print(f'Getting the upcoming {num_events} events')
    events_result = service.events().list(calendarId='primary', timeMin=now,
                                          maxResults=num_events, singleEvents=True,
                                          orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        print('No upcoming events found.')
        return

    # Prints the start and name of the next events
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        print(start, event['summary'])


def print_calendars(service):
   # Call the Calendar API
   print('Getting list of calendars')
   calendars_result = service.calendarList().list().execute()

   calendars = calendars_result.get('items', [])

   if not calendars:
       print('No calendars found.')
   for calendar in calendars:
       summary = calendar['summary']
       id = calendar['id']
       primary = "Primary" if calendar.get('primary') else ""
       print("%s\t%s\t%s" % (summary, id, primary))
