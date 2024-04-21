# Sync an ICS calendar file to Office 365 calendar online
# configuration is done in config.py

# imports
from O365 import Account
from icalendar import Calendar
import config

# init the calendar file
cal = Calendar()

# init the O365 connection
account =  Account(config.o365_credentials, auth_flow_type='credentials', tenant_id=config.o365_tenant_id)

# Authenticate and get access to the calendar
if not account.is_authenticated:    # will check if there is a token and has not expired
    account.authenticate(scopes=['basic', 'Calendars.ReadWrite.Shared'])

# A Schedule instance can list and create calendars.
# It can also list or create events on the default user calendar.
# To use other calendars use a Calendar instance.
schedule = account.schedule(resource=config.o365_resource)
#calendar = schedule.new_calendar(calendar_name=config.o365_calendar_name)
calendar = schedule.get_calendar(calendar_name=config.o365_calendar_name)

e = open(config.ics_filename, 'rb')

ecal = Calendar.from_ical(e.read())
for component in ecal.walk():
    if component.name == "VEVENT":
        print('Processing event: ' + component.get("SUMMARY"))
        new_event = calendar.new_event()
        new_event.subject = component.get("SUMMARY")
        new_event.location = component.get("location")
        new_event.start = component.decoded("dtstart")
        new_event.end = component.decoded("dtend")
        new_event.attendees = component.get("attendees")
        new_event.organizer = component.get("organizer")
        new_event.save()

e.close()

print("Finished.")
# EOF
