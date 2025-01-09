# CalExtractor

Some scripts to manager calendar synchronization between different calendar platforms.

- config.py         # General configuration for all scripts (access tokens, calendar names etc.)
- main.py         # This script downloads an .ics calendar file and writes all entries to a Office 365 and/or Google calendar. All content of this calendar is deleted first.

There is also a docker image available: https://hub.docker.com/repository/docker/ak1ras/calextractor/general

## Read an ICS file and sync your calendar

#### configure ics file location

`# ical file configuration
cs_filename = 'Cal_File.ics'                # internal name of the ical file to download and sync
ics_url = 'https://drive.google.com/uc?export=download&id=XXX'  # URL to download the ical file from`

### sync to Office 365

#### configure config.py and run SynFileTo365.py 

Office 365 has to be configured to access the API.

`# O365 connection credentials
o365_credentials = ('xxx', 'yyy')
o365_tenant_id = 'zzz'
o365_calendar_name='My Synced Cal'           # O365 calendar to sync to
o365_resource='address@email.org'            # O365 user name`

See: 
https://github.com/O365/python-o365#authentication
https://learn.microsoft.com/en-us/graph/tutorials/python?tabs=aad

### sync to Google Calendar

Basic configuration of the Google Calendar API:
https://developers.google.com/calendar/api/quickstart/python

### configure config.py and run SyncFileToGCal.py

The ics file should be in a Google Drive folder accessible to the python script.

Configure: `gcal_calendar_id` to the calendar ID of the Google Calendar you want to sync to. You can find the ID in the settings of the calendar.

### configure Google Credentials

Copy file googlecredentials.json-EXAMPLE to googlecredentials.json and fill in client_id anf secret from your Google API project (see quickstart guide).

## How to generate an ICS file 

### from Outlook 2019

You can export your calendar from Microsoft Outlook / Microsoft Exchange Server 2019 etc. by 
running a PowerShell command like:

`PS C:\Users\XXX> Invoke-WebRequest -Uri https://webmail.xyz.com/OWA/calendar/<id>/calendar.ics -OutFile OWACalFile.ics`

After that, you could send this file per EMail with a command like:

`PS C:\Users\XXX> send-mailmessage -to <EMailAdresse> -subject <Betreff> -from <XXX-Mail-Addresse> -smtpserver mailrelay.xxx.de`
You can use Windows Task Planner to execute these Powershell command periodically.

![Aufgabenplanung1.png](Images%2FAufgabenplanung1.png)
![Aufgabenplanung2.png](Images%2FAufgabenplanung2.png)

## How to save the .ics file from email to a file automatically

### to O365 OneDrive or Google Drive via Power Automate

You can use Microsoft Power Automate to process and delete a specific incoming email and save the .ics file attached to the email to a folder in OneDrive, from where the Python scripts can process it.

![PowerAutomate.png](Images%2FPowerAutomate.png)

For each attachment in the email, you have to decode the attachment using a decode function:

`decodeBase64(outputs('Anlage_abrufen_(V2)')?['body/contentBytes'])`

### to Google Drive via Google Apps Script

Create an Google Apps Script which processes your InBox and saves the attachment to a Google Drive folder.
Details: https://medium.com/@pablopallocchi/automatically-save-email-attachments-to-google-drive-using-google-apps-script-7a751a5d3ac9

1. Create a new Google Apps Script project within your Google Drive: Create a new Google Apps Script project by clicking on “New” > “More” > “Google Apps Script.”
2. Replace the default code with the script saveNewAttachmentsToDrive.js
3. Make sure to replace "XXX" with the actual ID of the destination folder in your Google Drive. To get the folder ID, open the folder in Google Drive, and the ID can be found in the URL after "folders/"
4. Modify the searchQuery variable to match the criteria for emails you want to process. The example query looks for emails with subject OWACalSync with attachments, but you can adjust it based on your requirements 
5. Save the script and give it a name
6. Run the script manually the first time to grant permissions for accessing Gmail and Drive
7. Set up a trigger to run the script periodically (e.g., every hour)
   8. Open Script Editor: In your Google Apps Script project, click on the menu “Extensions”, “Apps Script.” 
   9. Create a Trigger: In the Script Editor, click on the clock icon (Triggers) located on the left sidebar 
   10. Add New Trigger: Click on the “+ Add Trigger” button in the bottom right corner
   11. Configure Trigger: In the trigger configuration window 
       12. Select event source: Choose “Time-driven.”
       13. Select type of time based trigger: Select your preferred interval (e.g., every hour, every day, every week)
   14. Save Trigger: Once you’ve configured the trigger settings, click the “Save” button. The trigger will now be active and will automatically execute your script at the specified intervals.
