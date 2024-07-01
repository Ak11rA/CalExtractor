# Copyright: © 2024 Jipink
# Original: https://gist.github.com/pallocchi/

function saveNewAttachmentsToDrive() {
  var folderId = "XXX";
  var searchQuery = "subject:OWACalSync has:attachment";
  var lastExecutionTime = getLastExecutionDate();
  var threads = GmailApp.search(searchQuery + " after:" + lastExecutionTime);
  var driveFolder = DriveApp.getFolderById(folderId);
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var attachments = message.getAttachments();
      for (var k = 0; k < attachments.length; k++) {
        var attachment = attachments[k];
        var attachmentBlob = attachment.copyBlob();
        var fileName = attachment.getName();
        driveFolder.createFile(attachmentBlob).setName(fileName);
      }
    }
  }
  updateLastExecutionDate();
}

function getLastExecutionDate() {
  var properties = PropertiesService.getUserProperties();
  return properties.getProperty("lastExecutionDate") || "2023-09-01";
}

function resetLastExecutionDate() {
  PropertiesService.getUserProperties().deleteProperty("lastExecutionDate");
}

function updateLastExecutionDate() {
  var now = new Date();
  var dateString = now.toISOString().split("T")[0];
  var properties = PropertiesService.getUserProperties();
  properties.setProperty("lastExecutionDate", dateString);
}
