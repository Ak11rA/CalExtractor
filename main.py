#  Copyright (c) 2024 Loopback.ORG GmbH.

import os.path
import shutil

import recurring_ical_events
import SyncFileToCloud.sync2GCal
import SyncFileToCloud.sync2O365
import config
from icalendar import Calendar
import datetime
import urllib.request

def ical_walk(filename):
    # init the calendar file
    cal = Calendar()
    # Loop through file and create O365 events from each VEvent
    e = open(config.ics_filename, 'rb')
    ecal = Calendar.from_ical(e.read())

    for component in ecal.walk():
        if component.name == "VEVENT":
            for event in recurring_ical_events.of(component).between(period_begin, period_end):

                # TZ healing
                try:
                    if event.decoded("dtstart").tzid() is None:
                        cleaned_tz = config.default_TZ
                    else:
                        cleaned_tz = event.decoded("dtstart").tzid()
                except AttributeError:
                    print("Something was wrong with this events timezone.")
                    cleaned_tz = config.default_TZ

                if config.o365_tenant_id != 'zzz':
                    print('Adding event: '
                            + event.get("summary")
                            + ' at '
                            + event.decoded("dtstart").strftime("%B %d, %Y")
                            + ' (TZ: '
                            + cleaned_tz
                            + ') to Office 365.')
                    SyncFileToCloud.sync2O365.add_event(
                                        event.get("summary"),
                                        event.get("location"),
                                        event.decoded("dtstart"),
                                        event.decoded("dtend"),
                                        cleaned_tz
                    )

                if config.gcal_calendar_id != 'xxx@group.calendar.google.com':
                    print('Adding event: '
                            + event.get("summary")
                            + ' at '
                            + event.decoded("dtstart").strftime("%B %d, %Y")
                            + ' (TZ: '
                            + cleaned_tz
                            + ') to Google Calendar.')
                    SyncFileToCloud.sync2GCal.add_event(
                        event.get("summary"),
                        event.get("location"),
                        event.decoded("dtstart"),
                        event.decoded("dtend"),
                        cleaned_tz
                    )

    e.close()

# check for sane config or create config from conf directory when in Docker
if os.path.isfile('/conf/config.py') is True:
    print("Config file found in directory, running inside Docker?")
    shutil.copy('/conf/config.py', './config.py')
    if os.path.isfile('/conf/googlecredentials.json') is True:
        shutil.copy('/conf/googlecredentials.json', './googlecredentials.json')
    else:
        print("Google credentials not found in config directory.")
    if os.path.isfile('/conf/Googletoken.json') is True:
        shutil.copy('/conf/Googletoken.json', './Googletoken.json')
    else:
        print("Google token not found in config directory.")
    print("Config files copied to working directory. Please start the container again if this was the first run.")

if os.path.isfile('config.py') is False:
    print("My config file is missing. Please create config.py from config.py-EXAMPLE and set the correct values.")
    print("If you are running this inside a Docker container, please mount the config file to /conf/config.py.")
    exit(1)

period_begin = datetime.date.today() - datetime.timedelta(days=config.lookbehind_days)
period_end = datetime.date.today() + datetime.timedelta(days=config.lookahead_days)

if config.o365_tenant_id != 'zzz':
    SyncFileToCloud.sync2O365.connect()
    if config.cleanup_calendar:
        SyncFileToCloud.sync2O365.delete_events(config.o365_calendar_name, period_begin, period_end)
    else:
        print("Skipping cleanup of O365 calendar.")

if config.gcal_calendar_id != 'xxx@group.calendar.google.com':
    SyncFileToCloud.sync2GCal.connect()
    if config.cleanup_calendar:
        SyncFileToCloud.sync2GCal.delete_events(config.gcal_calendar_id, period_begin, period_end)
    else:
        print("Skipping cleanup of GCal calendar.")

# Download the ICS file
print('Downloading ICS file from ' + config.ics_url + ' to ' + config.ics_filename)
urllib.request.urlretrieve(config.ics_url, config.ics_filename)

# Process the ICS file
print("Processing ICS file.")
ical_walk(config.ics_filename)

print("Syncing done.")
