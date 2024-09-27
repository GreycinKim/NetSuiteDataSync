function importCSVFromEmail() {
  // Define search criteria for the email subject
  var query = 'subject:"EXAMPLE SUBJECT TITLE" has:attachment';
  var threads = GmailApp.search(query);
  
  if (threads.length === 0) {
    Logger.log('No emails found with the specified subject.');
    return;
  }

  // Get the most recent thread (first thread returned)
  var mostRecentThread = threads[0];
  var messages = mostRecentThread.getMessages();
  
  // Get the most recent message in the thread (first message returned)
  var message = messages[messages.length - 1];
  var attachments = message.getAttachments();

  // Loop through attachments
  for (var k = 0; k < attachments.length; k++) {
    var attachment = attachments[k];
    var attachmentName = attachment.getName().toLowerCase();
    
    // Check if the file is a CSV based on the file extension
    if (attachmentName.endsWith('.csv')) {
      var csvData = attachment.getDataAsString();
      var parsedData = Utilities.parseCsv(csvData);

      // Get "Sheet2" and clear its contents before importing new data
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SHEET1'); // Google Sheet Name
      if (!sheet) {
        Logger.log('(JY) French Bull Inventory does not exist.');
        return;
      }

      // Clear the existing data starting from row 2, but leave row 1 (headers) intact
      sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns()).clearContent();

      // Manually set the labels in the first row
      var headers = ["Item", "Description", "Inv Value", "% of Inv Value", "On Hand", "Last Purchase Price", "Average Cost", "Purchase Price"];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]); // Set the headers in row 1

   

      // Check if any item data was found before proceeding
      if (itemsData.length > 0) {
        // Write the filtered data (only actual items) starting at row 2 in Sheet2
        sheet.getRange(2, 1, itemsData.length, itemsData[0].length).setValues(itemsData); // Start at row 2
        Logger.log('CSV file imported successfully into Sheet2.');
      } else {
        Logger.log('No items found in the CSV file.');
      }
      
    } else {
      Logger.log('Attachment is not a CSV file: ' + attachmentName);
    }
  }
}
