#  Copyright (c) 2024 Loopback.ORG GmbH.

import config
from O365 import Account
from zoneinfo import ZoneInfo

def connect():
    # init the O365 connection
    account = Account(config.o365_credentials, auth_flow_type='credentials', tenant_id=config.o365_tenant_id)
    if account.authenticate():
        print('O365 Authentication successful.')

    # Authenticate and get access to the calendar
    if not account.is_authenticated:  # will check if there is a token and has not expired
        account.authenticate(scopes=['basic', 'Calendars.ReadWrite.Shared'])

def delete_events(calendar, period_begin, period_end):
    # Delete old entries in O365 calendar
    account = Account(config.o365_credentials, auth_flow_type='credentials', tenant_id=config.o365_tenant_id)
    schedule = account.schedule(resource=config.o365_resource)
    calendar = schedule.get_calendar(calendar_name=config.o365_calendar_name)
    q = calendar.new_query('start').greater_equal(period_begin)
    q.chain('and').on_attribute('end').less_equal(period_end)
    print(
        "Deleting O365 events from " + period_begin.strftime("%B %d, %Y") + " to " + period_end.strftime("%B %d, %Y") + ".")
    prevevents = calendar.get_events(query=q, limit=500, include_recurring=True)
    for event in prevevents:
        print("Deleting O365 event: " + event.subject)
        event.delete()

def add_event(
        summary,
        location,
        start,
        end,
        timezone):

    try:
        tz_info = ZoneInfo(timezone)
        starttime_tz = start.replace(tzinfo=tz_info)
        endtime_tz = end.replace(tzinfo=tz_info)
    except TypeError:
        print("Error: Timezone not found: (" + timezone + "). Using default timezone.")
        #print(" Event started at: " + start.strftime("%m.%d.%Y, %H:%M:%S"))
        #print(" Event started at: " + end.strftime("%m.%d.%Y, %H:%M:%S"))

    account = Account(config.o365_credentials, auth_flow_type='credentials', tenant_id=config.o365_tenant_id)
    schedule = account.schedule(resource=config.o365_resource)
    calendar = schedule.get_calendar(calendar_name=config.o365_calendar_name)

    try:
        new_event = calendar.new_event()
        new_event.subject = summary
        new_event.location = location
        new_event.start = starttime_tz
        new_event.end = endtime_tz
        new_event.tz = timezone
        new_event.save()
    except:
        print("Error: Could not add event to O365 calendar.")
