# CalExtractor

Some scripts to manager calendar synchronization between different calendar platforms.

- config.py         # General configuration for all scripts (access tokens, calendar names etc.)
- SyncFileTo365.py  # This script reads an .ics calendar file and writes all entries to a Office 365 calendar defined in config.py. All content of this calendar is deleted first.

## How to generate an ICS file 

### from Outlook 2019

You can export your calendar from Microsoft Outlook / Microsoft Exchange Server 2019 etc. by 
running a PowerShell command like:

`PS C:\Users\XXX> Invoke-WebRequest -Uri https://webmail.xyz.com/OWA/calendar/<id>/calendar.ics -OutFile OWACalFile.ics`

After that, you could send this file per EMail with a command like:

`PS C:\Users\XXX> send-mailmessage -to <EMailAdresse> -subject <Betreff> -from <XXX-Mail-Addresse> -smtpserver mailrelay.xxx.de`
You can use Windows Task Planner to execute these Powershell command periodically.

![](/Users/akira/Downloads/Aufgabenplanung 1.png)
![](/Users/akira/Downloads/Aufgabenplanung 2.png)

## How to save the .ics file from email to a file automatically

### to O365 OneDrive or Google Drive via Power Automate

You can use Microsoft Power Automate to process and delete a specific incoming email and save the .ics file attached to the email to a folder in OneDrive, from where the Python scripts can process it.

![](/Users/akira/Downloads/PowerAutomate.png)

For each attachment in the email, you have to decode the attachment using a decode function:

`decodeBase64(outputs('Anlage_abrufen_(V2)')?['body/contentBytes'])`