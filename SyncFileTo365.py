# Sync an ICS calendar file to Office 365 calendar online
# configuration is done in config.py

# imports
from O365 import Account
from icalendar import Calendar
import config
import datetime

# init the calendar file
cal = Calendar()

# init the O365 connection
account =  Account(config.o365_credentials, auth_flow_type='credentials', tenant_id=config.o365_tenant_id)
if account.authenticate():
   print('O365 Authentication successful.')

# Authenticate and get access to the calendar
if not account.is_authenticated:    # will check if there is a token and has not expired
    account.authenticate(scopes=['basic', 'Calendars.ReadWrite.Shared'])

# A Schedule instance can list and create calendars.
# It can also list or create events on the default user calendar.
# To use other calendars use a Calendar instance.
schedule = account.schedule(resource=config.o365_resource)
calendar = schedule.get_calendar(calendar_name=config.o365_calendar_name)

# Delete old entries in O365 calendar
period_begin = datetime.date.today() - datetime.timedelta(days=config.o365_lookbehind_days)
period_end = datetime.date.today() + datetime.timedelta(days=config.o365_lookahead_days)
q = calendar.new_query('start').greater_equal(period_begin)
q.chain('and').on_attribute('end').less_equal(period_end)
print ("Deleting events from "+period_begin.strftime("%B %d, %Y")+" to "+period_end.strftime("%B %d, %Y")+".")
prevevents = calendar.get_events(query=q, limit=500, include_recurring=True)
for event in prevevents:
        print("Deleting event: " + event.subject)
        event.delete()

# Loop through file and create O365 events from each VEvent
e = open(config.ics_filename, 'rb')
ecal = Calendar.from_ical(e.read())
for component in ecal.walk():
    if component.name == "VEVENT":
        print('Adding event: ' + component.get("SUMMARY") + ' at ' + component.decoded("dtstart").strftime("%B %d, %Y"))
        new_event = calendar.new_event()
        new_event.subject = component.get("SUMMARY")
        new_event.location = component.get("location")
        new_event.start = component.decoded("dtstart")
        new_event.end = component.decoded("dtend")
        new_event.save()
e.close()

print("Finished.")
# EOF
