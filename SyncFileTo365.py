#  Copyright (c) 2024 Loopback.ORG GmbH.

# Sync an ICS calendar file to Office 365 calendar online
# configuration is done in config.py

# imports
from O365 import Account
from icalendar import Calendar
import recurring_ical_events
import config
import datetime

# init the calendar file
cal = Calendar()

# init the O365 connection
account = Account(config.o365_credentials, auth_flow_type='credentials', tenant_id=config.o365_tenant_id)
if account.authenticate():
    print('O365 Authentication successful.')

# Authenticate and get access to the calendar
if not account.is_authenticated:  # will check if there is a token and has not expired
    account.authenticate(scopes=['basic', 'Calendars.ReadWrite.Shared'])

schedule = account.schedule(resource=config.o365_resource)
calendar = schedule.get_calendar(calendar_name=config.o365_calendar_name)

period_begin = datetime.date.today() - datetime.timedelta(days=config.lookbehind_days)
period_end = datetime.date.today() + datetime.timedelta(days=config.lookahead_days)

# Delete old entries in O365 calendar
if config.cleanup_calendar:
    q = calendar.new_query('start').greater_equal(period_begin)
    q.chain('and').on_attribute('end').less_equal(period_end)
    print(
        "Deleting events from " + period_begin.strftime("%B %d, %Y") + " to " + period_end.strftime("%B %d, %Y") + ".")
    prevevents = calendar.get_events(query=q, limit=500, include_recurring=True)
    for event in prevevents:
        print("Deleting event: " + event.subject)
        event.delete()
else:
    print("Skipping cleanup of O365 calendar.")

# Loop through file and create O365 events from each VEvent
e = open(config.ics_filename, 'rb')
ecal = Calendar.from_ical(e.read())
for component in ecal.walk():
    if component.name == "VEVENT":
        for event in recurring_ical_events.of(component).between(period_begin, period_end):
            print('Adding event: ' + event.get("summary") + ' at ' + event.decoded("dtstart").strftime("%B %d, %Y")
                  )
            new_event = calendar.new_event()
            new_event.subject = event.get("summary")
            new_event.location = event.get("location")
            new_event.start = event.decoded("dtstart")
            new_event.end = event.decoded("dtend")
            new_event.save()
e.close()

print("Finished.")
# EOF
