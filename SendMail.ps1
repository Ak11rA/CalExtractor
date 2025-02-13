#'       |||      
#'      (o o)     
#'╔═ooO══(_)══Ooo════════════════════════════════════════════════════════════════════════════════════════════╗
#'║                                                                                                          ║
#'║                                            Send-SerialMail.ps1                                           ║
#'║                                                                                                          ║
#'╟───────────────────────┬──────────────────────────────────────────────────────────────────────────────────╢
#'║ AUTHORs               │ Martin Schroedl, Jan Schreiber                                                   ║
#'╟───────────────────────┼──────────────────────────────────────────────────────────────────────────────────╢
#'║ Version               │ 2025.02.12.0001                                                                  ║
#'╚═══════════════════════╧══════════════════════════════════════════════════════════════════════════════════╝

# Hier wird die Betreffzeile festgelegt
  $Subject = "Das ist eine Testnachricht"

# Hier wird die E-Mail Nachricht als HTML abgelegt
  $HTMLBody = "<html><body>%%%ANREDE%%%<br /><br />Das ist ein Test</body></html>"

# Hier wird festgelegt, ob ein Attachment angefügt werden soll
  $Attachment = "S:<...>\icsUpload\OWACalFile.ics"

# Hier wird über die Variable 0 (Save) oder 1 (Send) festgelegt, ob die Nachricht nur gespeichert oder sofort
# gesendet werden soll:
  $SaveorSend = 0

# Hier wird die Verbindung zu Outlook hergestellt
  $OutlookConnection = New-Object -ComObject Outlook.Application

# Hier wird nun für jeden Empänger eine Nachricht erstellt
  $Empfaenger = @{
    'E-Mail' = "xy@mail.adress"
    Anrede = "Herr/Frau"
  }
  
  $EMailMessage.Recipients.Add($Empfaenger['E-Mail'])
  $EMailMessage = $OutlookConnection.CreateItem("olMailItem")
  $EMailMessage.Subject = $Subject
  $EMailMessage.Recipients.Add($Empfaenger.'E-Mail')
  $EMailMessage.HTMLBody = $HTMLBody -replace("%%%ANREDE%%%",$Empfaenger.Anrede)
  if ($Attachment -ne $null) { $EMailMessage.Attachments.Add($Attachment) }
  if ($SaveorSend -eq 1) {$EMailMessage.Send()} else {$EMailMessage.Save()}
  Remove-Variable EMailMessage

  Remove-Variable OutlookConnection