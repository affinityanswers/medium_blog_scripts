function getEmailsAndWriteToSpreadsheet() {
    var targetDomain = 'domainName';
  
    // Log in to Gmail
    var gmailApp = GmailApp;
    var threads = gmailApp.search('from:' + targetDomain);
  
    // Create a new Google Sheet or use an existing one
    var spreadsheet = SpreadsheetApp.openById('sheetID');
    var sheet = spreadsheet.getSheetByName('sheetName'); // Change the sheet name as needed
  
    // Get the list of processed message IDs from Script Properties
    var scriptProperties = PropertiesService.getScriptProperties();
    var processedMessageIds = scriptProperties.getProperty('processedMessageIds');
    processedMessageIds = processedMessageIds ? processedMessageIds.split(',') : [];
  
    // Write headers to the spreadsheet if the sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1).setValue('Subject');
      sheet.getRange(1, 2).setValue('Date');
      sheet.getRange(1, 3).setValue('Name');
      sheet.getRange(1, 4).setValue('Email Address');
      sheet.getRange(1, 5).setValue('Address');
      sheet.getRange(1, 6).setValue('Request Received via');
    }
  
    // Loop through email threads and retrieve messages
    for (var i = 0; i < threads.length; i++) {
      var messages = threads[i].getMessages();
      for (var j = 0; j < messages.length; j++) {
        var message = messages[j];
        var messageId = message.getId();
  
        // Check if the message has already been processed
        if (!processedMessageIds.includes(messageId)) {
          var subject = message.getSubject();
          var date = message.getDate();
          var body = message.getPlainBody();
          var senderEmail = extractSenderEmail(message.getFrom());
  
          // Search for "Name," "E-mail address," "State," and "Country" in the email body
          var nameMatch = /Name: ([^\n]+)/i.exec(body);
          var emailMatch = /E-mail address: ([^\n]+)/i.exec(body);
          var addressMatch = /Address: ([^\n]+)/i.exec(body);
  
          if (nameMatch && emailMatch) {
            var name = nameMatch[1].trim();
            var email = emailMatch[1].trim();
            var address = addressMatch ? addressMatch[1].trim() : '';
  
            // Write email details, name, email-address, state code, and country to the spreadsheet
            sheet.appendRow([subject, date, name, email, address, senderEmail]);
  
            // Add the message ID to the list of processed message IDs
            processedMessageIds.push(messageId);
          }
        }
      }
    }
  
    // Update the list of processed message IDs in Script Properties
    scriptProperties.setProperty('processedMessageIds', processedMessageIds.join(','));
  }
  
  // Function to extract sender's email address from the sender information
  function extractSenderEmail(senderInfo) {
    var emailMatch = /<([^>]+)>/.exec(senderInfo);
    return emailMatch ? emailMatch[1] : senderInfo;
  }
  