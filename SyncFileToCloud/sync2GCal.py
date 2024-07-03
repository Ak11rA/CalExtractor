#  Copyright (c) 2024 Loopback.ORG GmbH.

# Sync an ICS calendar file to Office 365 calendar online
# configuration is done in config.py

import os
import config
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

def connect():
    # If modifying these scopes, delete the file Googletoken.json.
    SCOPES = ["https://www.googleapis.com/auth/calendar"]
    creds = None
    # The file Googletoken.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("Googletoken.json"):
        creds = Credentials.from_authorized_user_file("Googletoken.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "googlecredentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("Googletoken.json", "w") as token:
            token.write(creds.to_json())
    else:
        print("Google credentials are valid.")

def delete_events(calendar, period_begin, period_end):
    # Delete old entries in Google calendar
    SCOPES = ["https://www.googleapis.com/auth/calendar"]
    creds = Credentials.from_authorized_user_file("Googletoken.json", SCOPES)
    try:
        service = build('calendar', "v3", credentials=creds)
        print("Deleting previous Google Calendar events.")
        events_result = (
            service.events()
            .list(
                calendarId=config.gcal_calendar_id,
                singleEvents=True,
                orderBy="startTime"
            )
            .execute()
        )
        events = events_result.get("items", [])
        if not events:
            print("No events found.")
        for event in events:
            start = event["start"].get("dateTime", event["start"].get("date"))
            print("Deleting GCal event: " + start, event["summary"])
            (service.events()
             .delete(
                calendarId=config.gcal_calendar_id,
                eventId=event["id"]
            ).execute())
    except HttpError as error:
        print(f"An error occurred: {error}")

def add_event(
        summary,
        location,
        start,
        end
):
    SCOPES = ["https://www.googleapis.com/auth/calendar"]
    creds = Credentials.from_authorized_user_file("Googletoken.json", SCOPES)
    gc_event = {
        'summary': summary,
        'location': location,
        'start': {
            'dateTime': start.isoformat()
        },
        'end': {
            'dateTime': end.isoformat()
        }
    }
    try:
        service = build('calendar', "v3", credentials=creds)
        new_event = (service.events()
                     .insert(
            calendarId=config.gcal_calendar_id,
            body=gc_event
        ).execute())
    except HttpError as error:
        print(f"An error occurred: {error}")
